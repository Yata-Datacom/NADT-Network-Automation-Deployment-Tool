"""
NADT — Network Automation Deployment Tool / 网络自动化交付工具
===============================================================
基于 DHCP option 66/67 的华为交换机批量自动化开局工具。

原理（Zero-Touch Provisioning）:
  1. 交换机出厂开机 → 从 DHCP 获取 IP
  2. DHCP 服务器通过 option 66 下发 TFTP 服务器地址（本机）
  3. 通过 option 67 下发引导文件名（lswnet.cfg — ESN 映射表）
  4. 交换机下载 lswnet.cfg，按自己的 ESN 找到对应配置文件
  5. 下载自己的 .cfg 并自动应用 → 开局完成

功能:
  - 设备清单管理（CSV/Excel 导入、手动编辑）
  - 模板渲染引擎（{{占位符}}，支持任意华为配置模板）
  - 批量生成交换机 .cfg + lswnet.cfg ESN 映射表
  - 内置 TFTP 服务器（纯 stdlib，RFC 1350）
  - DHCP option 66/67 配置片段生成

依赖: 仅 Python 标准库（可选 openpyxl 用于 Excel 导入）
"""
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import os
import sys
import re
import csv
import io
import json
import socket
import struct
import asyncio
import ipaddress
import time
import shutil
import copy
from datetime import datetime

__version__ = "1.12.2"

# ═══════════════════════════════════════════════════════════════
# 常量 / Constants
# ═══════════════════════════════════════════════════════════════

# 打包成 exe 后 __file__ 在临时目录，必须用 exe 所在目录作为可写基目录
# when frozen (PyInstaller), use the exe's own directory as the writable base
BASE_DIR = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
    else os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
DEFAULT_TEMPLATE = os.path.join(TEMPLATE_DIR, "access-switch.cfg")
OUTPUT_DIR   = os.path.join(BASE_DIR, "output")
BOOTFILE     = "lswnet.cfg"          # ESN 映射表文件名（option 67 指向它）

CSV_HEADER = ["esn", "hostname", "type", "mgmt_ip", "mgmt_vlan", "access_vlan",
              "mask", "gateway", "note"]


# ═══════════════════════════════════════════════════════════════
# 核心逻辑 / Core
# ═══════════════════════════════════════════════════════════════

def derive_defaults(dev: dict) -> dict:
    """补全通用默认值：标准键兜底、掩码、网关（.254）、vlan_batch。

    type 字段是自由文本，只用于「类型→模板映射」选模板，不做 VLAN 推导。
    """
    d = dict(dev)
    for k in CSV_HEADER:
        d.setdefault(k, "")                     # 保证标准键存在 / ensure keys exist
    if not (d.get("mask") or "").strip():
        d["mask"] = "255.255.255.0"
    mgmt_ip = (d.get("mgmt_ip") or "").strip()
    gw = (d.get("gateway") or "").strip()
    if not gw and mgmt_ip:
        try:
            ip = ipaddress.ip_address(mgmt_ip)
            parts = str(ip).split(".")
            parts[-1] = "254"
            gw = ".".join(parts)
        except ValueError:
            pass
    d["gateway"] = gw
    mv = (d.get("mgmt_vlan") or "").strip()
    av = (d.get("access_vlan") or "").strip()
    d["vlan_batch"] = f"{mv} {av}".strip()
    d["hostname"] = (d.get("hostname") or "").strip()
    return d


# 占位符：允许花括号内有空格（{{ hostname }} 与 {{hostname}} 等价）
_PLACEHOLDER_RE = re.compile(r"\{\{\s*(\w+)\s*\}\}")
# 任何残留的 {{...}}（含 {{#if}}/{{#for}}/带空格写法）——用于「模板缺变量」保护
_LEFTOVER_RE = re.compile(r"\{\{[^{}]*\}\}")


def _subst(line: str, v: dict) -> str:
    """
    替换单行里的 {{var}}，未定义的保留原样（最后统一检查）。

    ⚠️ 花括号内允许空格（`{{ hostname }}`）：早期实现的正则只认紧贴写法（两个花括号 + 词 + 两个花括号），
    带空格时既不替换、又不被残留检测发现，会**静默生成含字面占位符的配置**。
    """
    def rep(m):
        k = m.group(1)
        return str(v[k]) if k in v else m.group(0)
    return _PLACEHOLDER_RE.sub(rep, line)


def _eval_truthy(v: dict, expr: str) -> bool:
    """求值 {{#if}} 条件：var / !var / var=值。"""
    expr = expr.strip()
    if expr.startswith("!"):
        return not _eval_truthy(v, expr[1:])
    if "=" in expr:
        k, _, val = expr.partition("=")
        return str(v.get(k.strip(), "")).strip() == val.strip()
    return str(v.get(expr, "")).strip() != ""


_RANGE = r"(\d+|\{\{\w+\}\})"          # for 范围：数字 或 {{变量}}
_INL_FOR_RE = re.compile(rf"^\{{{{#for\s+(\w+)\s+from\s+{_RANGE}\s+to\s+{_RANGE}\s*\}}}}(.*)\{{{{#endfor}}}}\s*$")
_FOR_RE = re.compile(rf"\{{{{#for\s+(\w+)\s+from\s+{_RANGE}\s+to\s+{_RANGE}\s*}}}}")


def _find_block_end(lines: list, start: int):
    """从 start 找匹配的 {{#endif}}/{{#endfor}}，返回 (block, else_block, next_idx)。

    else_block 仅对 if 有效（for 无 else）。找不到抛 ValueError。
    单行完整内联块（{{#if}}…{{#endif}} 同行 / {{#for}}…{{#endfor}} 同行）不参与深度计数。
    """
    depth = 1
    block, els = [], None
    i = start
    while i < len(lines):
        s = lines[i].strip()
        if (s.startswith("{{#if") or s.startswith("{{#for")) and \
                (s.endswith("{{#endif}}") or s.endswith("{{#endfor}}")):
            (block if els is None else els).append(lines[i])
            i += 1
            continue
        if s.startswith("{{#if") or s.startswith("{{#for"):
            depth += 1
        elif s.startswith("{{#endif") or s.startswith("{{#endfor"):
            if s not in ("{{#endif}}", "{{#endfor}}"):
                raise ValueError(f"块结束指令不能带其他内容（会静默丢文本），请换行写：{s[:40]}")
            depth -= 1
            if depth == 0:
                return block, els, i + 1
        if s.startswith("{{#else"):
            # 与 {{#if}}/{{#for}} 带 body 的行为保持一致：宁可报错，也不静默丢掉行内文本
            if s != "{{#else}}":
                raise ValueError(f"{{#else}} 指令行不能带其他内容（会静默丢文本），请换行写：{s[:40]}")
            if depth == 1:
                els = []
                i += 1
                continue
        (block if els is None else els).append(lines[i])
        i += 1
    raise ValueError("模板缺少块结束符 {{#endif}}/{{#endfor}}")


def _render_lines(lines: list, v: dict, out: list):
    """渲染一组行，处理 {{#if}}/{{#for}} 指令与 {{var}} 替换。"""
    i = 0
    n = len(lines)
    while i < n:
        s = lines[i].strip()
        m_if = re.match(r"\{\{#if\s+(.+?)\s*\}\}", s)
        m_for = _FOR_RE.match(s)

        # 单行内联块: {{#if x}}A{{#else}}B{{#endif}} / {{#for n from 1 to N}}...{{#endfor}}
        m_inl_if = re.match(r"^\{\{#if\s+(.+?)\}\}(.*)\{\{#endif\}\}\s*$", s)
        m_inl_for = _INL_FOR_RE.match(s)
        if m_inl_if and "{{#if" not in m_inl_if.group(2) and "{{#for" not in m_inl_if.group(2):
            expr, body = m_inl_if.group(1), m_inl_if.group(2)
            if "{{#else}}" in body:
                then_s, _, else_s = body.partition("{{#else}}")
            else:
                then_s, else_s = body, ""
            out.append(_subst(then_s if _eval_truthy(v, expr) else else_s, v))
            i += 1
            continue
        if m_inl_for and "{{#if" not in m_inl_for.group(4) and "{{#for" not in m_inl_for.group(4):
            name, a_expr, b_expr, body = (m_inl_for.group(1), m_inl_for.group(2),
                                          m_inl_for.group(3), m_inl_for.group(4))
            try:
                a = int(_subst(a_expr, v).strip())
                b = int(_subst(b_expr, v).strip())
            except ValueError:
                raise ValueError(f"{{{{#for}}}} 范围不是数字: {s.strip()}")
            parts = []
            for k in range(a, b + 1):
                parts.append(_subst(body, {**v, name: str(k)}))
            out.append("".join(parts))
            i += 1
            continue
        # 单行块但含嵌套 → 明确报错（不支持）
        if (s.startswith("{{#if") or s.startswith("{{#for")) and (
                "{{#endif}}" in s or "{{#endfor}}" in s):
            raise ValueError(f"单行块不支持嵌套指令: {s.strip()}")

        if m_if:
            tail = s[m_if.end():].strip()
            if tail:
                raise ValueError(
                    f"{{{{#if}}}} 指令行不能带内容: {s.strip()}\n"
                    "请把内容放下一行，或写成单行 {{#if}}…{{#endif}}")
            block, els, ni = _find_block_end(lines, i + 1)
            if _eval_truthy(v, m_if.group(1)):
                _render_lines(block, v, out)
            elif els:
                _render_lines(els, v, out)
            i = ni
            continue
        if m_for:
            name, a_expr, b_expr = m_for.group(1), m_for.group(2), m_for.group(3)
            tail = s[m_for.end():].strip()
            if tail:
                raise ValueError(
                    f"{{{{#for}}}} 指令行不能带内容: {s.strip()}\n"
                    "请把内容放下一行，或写成单行 {{#for}}…{{#endfor}}")
            try:
                a = int(_subst(a_expr, v).strip())
                b = int(_subst(b_expr, v).strip())
            except ValueError:
                raise ValueError(f"{{{{#for}}}} 范围不是数字: {s.strip()}")
            if a > b:
                raise ValueError(f"{{{{#for}}}} 范围无效: {a} > {b}")
            block, _, ni = _find_block_end(lines, i + 1)
            for k in range(a, b + 1):
                _render_lines(block, {**v, name: str(k)}, out)
            i = ni
            continue
        out.append(_subst(lines[i], v))
        i += 1


def render_template(template_text: str, vars_map: dict) -> str:
    """渲染模板：支持 {{var}}、{{#if}}/{{#else}}/{{#endif}}、{{#for}}/{{#endfor}}。

    未定义变量保留 {{var}} 原样（由调用方统一检查遗漏）。
    """
    out: list = []
    _render_lines(template_text.splitlines(), vars_map, out)
    return "\n".join(out) + ("\n" if template_text.endswith("\n") else "")


def build_switch_config(template_text: str, dev: dict) -> str:
    """渲染单台交换机配置。变量 = 标准字段 + 派生值 + 扩展字段(_extras)。"""
    d = derive_defaults(dev)
    vars_map = {
        "hostname":    d["hostname"],
        "vlan_batch":  d["vlan_batch"],
        "mgmt_vlan":   d["mgmt_vlan"],
        "mgmt_ip":     d["mgmt_ip"],
        "mask":        d["mask"],
        "access_vlan": d["access_vlan"],
        "gateway":     d["gateway"],
    }
    extras = dev.get("_extras") or {}
    for k, val in extras.items():
        if val is not None:
            vars_map[k] = str(val)
    if "{{vlan_batch}}" in template_text and not vars_map["vlan_batch"]:
        raise ValueError("模板使用 vlan_batch，但设备缺少 管理VLAN/接入VLAN")
    cfg = render_template(template_text, vars_map)
    # 模板里残留的未替换占位符 = 模板缺变量 → 抛错提示
    # 任何残留的 {{...}} 都算模板问题：{{ hostname }}（带空格）、{{#if}}（指令不在行首时整行
    # 原样输出）都在内 —— 宁可报错，也绝不把字面占位符下发给交换机
    leftovers = _LEFTOVER_RE.findall(cfg)
    if leftovers:
        names = sorted({re.sub(r"\s+", " ", x)[2:-2].strip() for x in leftovers})
        raise ValueError(f"模板缺少变量: {', '.join(names)}")
    return cfg


def build_lswnet(devices: list, vrpfile: str = "", patchfile: str = "") -> str:
    """生成 ESN 映射表（lswnet.cfg）。

    支持 vrpfile（版本文件 .cc）与 patchfile（补丁 .pat）：
    设备 CSV 列 / 扩展字段 vrp_file、patch_file 优先，其次用全局默认参数。
    """
    lines = []
    for dev in devices:
        d = derive_defaults(dev)
        esn = (d.get("esn") or "").strip()
        host = (d.get("hostname") or "").strip()
        if not esn or not host:
            continue
        extra = d.get("_extras") or {}
        vf = (extra.get("vrp_file") or d.get("vrp_file") or vrpfile or "").strip()
        pf = (extra.get("patch_file") or d.get("patch_file") or patchfile or "").strip()
        seg = f"esn={esn};"
        if vf:
            seg += f"vrpfile={vf};"
        if pf:
            seg += f"patchfile={pf};"
        seg += f"cfgfile={host}.cfg;"
        lines.append(seg)
    return "\n".join(lines) + ("\n" if lines else "")


def build_usb_ini(devices: list, fileserver: str = "file:/usb:",
                  system_software: str = "", system_pat: str = "") -> str:
    """生成华为 EasyDeploy USB 开局 ini（usb_config.ini）。

    每台设备一个 [DEVICEn DESCRIPTION] 块：按 ESN（或 MAC）匹配；
    SYSTEM-CONFIG 自动取设备主机名.cfg；SYSTEM-SOFTWARE/SYSTEM-PAT
    可用设备扩展字段（system_software/system_pat）或全局默认；
    STACK-MEMBER-ID 来自扩展字段 stack_member_id。
    """
    out = [";BEGIN DC", "[GLOBAL CONFIG]", f"FILESERVER={fileserver}"]
    idx = 0
    for dev in devices:
        d = derive_defaults(dev)
        ex = d.get("_extras") or {}
        esn = (d.get("esn") or "").strip()
        mac = (d.get("mac") or ex.get("mac") or "").strip()
        host = (d.get("hostname") or "").strip()
        if (not esn and not mac) or not host:
            continue
        extra = d.get("_extras") or {}
        out.append("")
        out.append(f"[DEVICE{idx} DESCRIPTION]")
        idx += 1
        out.append(f"ESN={esn if esn else 'DEFAULT'}")
        # MAC 非空但格式非法时 normalize_mac 返回 ""：此时**整行跳过**，
        # 否则 ini 里会出现空的 `MAC=` 项（设备可能忽略或解析异常）
        mac_norm = normalize_mac(mac, "huawei") if mac else ""
        if mac_norm:
            out.append(f"MAC={mac_norm}")
        out.append("DEVICETYPE=DEFAULT")
        sof = (extra.get("system_software") or system_software or "").strip()
        if sof:
            out.append(f"SYSTEM-SOFTWARE={sof}")
        out.append(f"SYSTEM-CONFIG={host}.cfg")
        pat = (extra.get("system_pat") or system_pat or "").strip()
        if pat:
            out.append(f"SYSTEM-PAT={pat}")
        mid = (extra.get("stack_member_id") or d.get("stack_member_id") or "").strip()
        if mid:
            out.append(f"STACK-MEMBER-ID={mid}")
    out.append(";END DC")
    return "\r\n".join(out) + "\r\n"      # 华为要求 CRLF，函数级保证


def _parse_range_values(expr: str) -> list:
    """解析 {..} 内容：'1-10' → 1..10；'1,10' → [1,10]；'01-10' → 零填充。"""
    vals = []
    for p in (x.strip() for x in expr.split(",") if x.strip()):
        if "-" in p:
            a, _, b = p.partition("-")
            try:
                lo, hi = int(a), int(b)
            except ValueError:
                return []
            if lo > hi:
                lo, hi = hi, lo
            pad = 0
            if (a.startswith("0") and len(a) > 1) or (b.startswith("0") and len(b) > 1):
                pad = max(len(a), len(b))
            if pad:
                vals.extend(str(i).zfill(pad) for i in range(lo, hi + 1))
            else:
                vals.extend(str(i) for i in range(lo, hi + 1))
        else:
            vals.append(p)
    return vals


def expand_ranges(text: str) -> list:
    """展开字符串里的 {1-10} / {1,10} / {01-10}，返回展开后的字符串列表。

    无范围语法时返回 [原字符串]；解析失败也按原文返回。
    """
    if "{" not in text:
        return [text]
    matches = re.findall(r"\{([^{}]*)\}", text)
    if not matches:
        return [text]
    choices = []
    for m in matches:
        vals = _parse_range_values(m)
        if not vals:
            return [text]
        choices.append(vals)
    results = []
    for combo in __import__("itertools").product(*choices):
        out = text
        for m, v in zip(matches, combo):
            out = out.replace("{" + m + "}", v, 1)
        results.append(out)
    return results


def expand_device(dev: dict) -> list:
    """展开设备字段里的 {1-10} 批量编号 → 设备列表（多字段联动、一一对应）。

    标准字段和扩展字段都支持；无范围时返回 [原设备]。
    """
    def _has(v):
        return isinstance(v, str) and "{" in v

    ex = dev.get("_extras") or {}
    range_fields = [k for k, v in dev.items() if k != "_extras" and _has(v)]
    range_extras = [k for k, v in ex.items() if _has(v)]
    if not range_fields and not range_extras:
        return [dev]

    expansions = {}
    for k in range_fields:
        expansions[("f", k)] = expand_ranges(dev[k])
    for k in range_extras:
        expansions[("e", k)] = expand_ranges(ex[k])
    count = min(len(v) for v in expansions.values())

    results = []
    for i in range(count):
        nd = dict(dev)
        nd["_extras"] = dict(ex)
        for (kind, k), vals in expansions.items():
            if kind == "f":
                nd[k] = vals[i]
            else:
                nd["_extras"][k] = vals[i]
        results.append(nd)
    return results


def _count_range_values(expr: str) -> int:
    """
    **只数不展开**地计算 `{..}` 内容能展开多少项（与 `_parse_range_values` 分支规则保持一致）。

    ⚠️ 存在的理由：`count_expansions` 号称"轻量预估"，早期却调用会**完整物化**区间列表的
    `_parse_range_values` —— `{1-1000000}` 峰值约 63MB、`{1-100000000}` 直接吃满内存，
    防组合爆炸的保护形同虚设。这里改成纯算术（O(段数) 时间、零额外内存）。
    """
    total = 0
    for part in (x.strip() for x in expr.split(",") if x.strip()):
        if "-" in part:
            a, _, b = part.partition("-")
            try:
                lo, hi = int(a), int(b)
            except ValueError:
                return 0
            if lo > hi:
                lo, hi = hi, lo
            total += hi - lo + 1
        elif part.isdigit():
            total += 1
        else:
            return 0          # 非法项 → 与 _parse_range_values 返回 [] 的行为对齐
    return total


def count_expansions(dev: dict) -> int:
    """轻量预估展开数量（**纯算术，不展开**，防 {1-1000000} 组合爆炸）。"""
    def _count(text: str) -> int:
        if "{" not in text:
            return 1
        n = 1
        for m in re.findall(r"\{([^{}]*)\}", text):
            vals = _count_range_values(m)
            n *= vals if vals else 1
        return n
    counts = []
    for k, v in dev.items():
        if k != "_extras" and isinstance(v, str) and "{" in v:
            counts.append(_count(v))
    ex = dev.get("_extras") or {}
    for k, v in ex.items():
        if isinstance(v, str) and "{" in v:
            counts.append(_count(v))
    return min(counts) if counts else 1


def check_range_lengths(dev: dict) -> str:
    """检查各字段范围长度是否一致，不一致返回警告文本（一致返回空串）。"""
    def _has(v):
        return isinstance(v, str) and "{" in v
    lengths = {}
    for k, v in dev.items():
        if k != "_extras" and _has(v):
            lengths[k] = len(expand_ranges(v))
    ex = dev.get("_extras") or {}
    for k, v in ex.items():
        if _has(v):
            lengths["扩展." + k] = len(expand_ranges(v))
    distinct = set(lengths.values())
    if len(distinct) > 1:
        return f"范围长度不一致 {lengths}，按最短 {min(distinct)} 台生成"
    return ""


def validate_devices(devices: list) -> list:
    """批量校验设备清单，返回 [(severity, host, message)]。

    severity: error（重复 ESN/IP/主机名、缺必填、IP 格式错）/ warning（缺 ESN）。
    """
    problems = []
    seen_esn, seen_host, seen_ip = {}, {}, {}
    for d in devices:
        dd = derive_defaults(d)
        esn = dd["esn"].strip()
        host = dd["hostname"].strip()
        ip = dd["mgmt_ip"].strip()
        tag = host or esn or "?"
        if not host:
            problems.append(("error", esn or "?", "缺主机名"))
        if not ip:
            problems.append(("error", tag, "缺管理 IP"))
        else:
            try:
                ipaddress.ip_address(ip)
            except ValueError:
                hint = ""
                if "{" in ip or re.search(r"\d+-\d+", ip):
                    hint = "（如需批量编号请用花括号格式，如 192.168.1.{1-100}）"
                problems.append(("error", tag, f"管理 IP 格式错误: {ip}{hint}"))
        if esn:
            if esn in seen_esn:
                problems.append(("error", tag, f"ESN 重复: {esn}（与 {seen_esn[esn]} 冲突）"))
            else:
                seen_esn[esn] = host or esn
        else:
            problems.append(("warning", tag, "缺 ESN（无法写入映射表，设备不能自动匹配配置）"))
        if host:
            if host in seen_host:
                problems.append(("error", host, f"主机名重复: {host}"))
            else:
                seen_host[host] = ip
        if ip:
            if ip in seen_ip:
                problems.append(("error", tag, f"管理 IP 重复: {ip}（与 {seen_ip[ip]} 冲突）"))
            else:
                seen_ip[ip] = host or ip
    return problems


DB_PATH = os.path.join(BASE_DIR, "devices.db")


def save_devices_db(devices: list):
    """设备清单持久化到 SQLite（devices.db），失败静默。"""
    try:
        import sqlite3
        conn = sqlite3.connect(DB_PATH)
        conn.execute("CREATE TABLE IF NOT EXISTS devices ("
                     "esn TEXT, hostname TEXT, type TEXT, mgmt_ip TEXT, mgmt_vlan TEXT, "
                     "access_vlan TEXT, mask TEXT, gateway TEXT, note TEXT, extras TEXT)")
        conn.execute("DELETE FROM devices")
        rows = []
        for d in devices:
            rows.append((d.get("esn", ""), d.get("hostname", ""), d.get("type", ""),
                         d.get("mgmt_ip", ""), d.get("mgmt_vlan", ""), d.get("access_vlan", ""),
                         d.get("mask", ""), d.get("gateway", ""), d.get("note", ""),
                         json.dumps(d.get("_extras") or {}, ensure_ascii=False)))
        conn.executemany("INSERT INTO devices VALUES (?,?,?,?,?,?,?,?,?,?)", rows)
        conn.commit()
        conn.close()
    except Exception:
        pass


def load_devices_db() -> list:
    """从 SQLite 加载设备清单，失败返回 []。"""
    try:
        import sqlite3
        conn = sqlite3.connect(DB_PATH)
        conn.execute("CREATE TABLE IF NOT EXISTS devices ("
                     "esn TEXT, hostname TEXT, type TEXT, mgmt_ip TEXT, mgmt_vlan TEXT, "
                     "access_vlan TEXT, mask TEXT, gateway TEXT, note TEXT, extras TEXT)")
        conn.row_factory = sqlite3.Row
        rows = conn.execute("SELECT * FROM devices").fetchall()
        conn.close()
        devices = []
        for r in rows:
            d = {k: (r[k] or "") for k in ("esn", "hostname", "type", "mgmt_ip", "mgmt_vlan",
                                           "access_vlan", "mask", "gateway", "note")}
            try:
                d["_extras"] = json.loads(r["extras"] or "{}")
            except Exception:
                d["_extras"] = {}
            devices.append(d)
        return devices
    except Exception:
        return []


# 表头中文/常见别名 → 标准字段 / header aliases
CSV_ALIASES = {
    "esn": "esn", "序列号": "esn", "sn": "esn", "serial": "esn", "设备序列号": "esn",
    "hostname": "hostname", "设备名": "hostname", "名称": "hostname", "主机名": "hostname",
    "设备名称": "hostname", "设备hostname": "hostname",
    "type": "type", "型号": "type", "类型": "type", "设备类型": "type",
    "mgmt_ip": "mgmt_ip", "管理ip": "mgmt_ip", "ip": "mgmt_ip", "ip地址": "mgmt_ip",
    "管理地址": "mgmt_ip", "管理ip地址": "mgmt_ip",
    "mgmt_vlan": "mgmt_vlan", "管理vlan": "mgmt_vlan", "vlanif": "mgmt_vlan",
    "access_vlan": "access_vlan", "接入vlan": "access_vlan", "用户vlan": "access_vlan",
    "mask": "mask", "掩码": "mask", "子网掩码": "mask",
    "gateway": "gateway", "网关": "gateway", "默认网关": "gateway",
    "note": "note", "备注": "note", "说明": "note", "位置": "note",
}

# 列映射弹窗可选字段 / column mapping options
MAPPING_OPTIONS = ["esn", "hostname", "type", "mgmt_ip", "mgmt_vlan", "access_vlan",
                   "mask", "gateway", "note", "忽略", "扩展字段"]


def parse_sheet_rows(rows: list, sheet_name: str, warnings: list):
    """从二维表解析 模板+设备（Excel 宏工作流结构）→ (templates, devices)。

    结构：
      A1 = 【模板名】（或 A 列含 {{变量}} 也算模板 sheet）
      第 1 行 C 列起 = 参数名列头（{{sysname}} → 参数名 sysname）
      A 列第 2 行起 = 配置模板文本（{{参数}} 占位符，'#' 行保留）
      第 2 行起每行 = 一台设备（C 列起为该行各参数值）

    非模板 sheet → ({}, [])；映射表 sheet 跳过并写 warnings。
    """
    if not rows:
        return {}, []
    a1 = str(rows[0][0] or "") if rows[0] else ""
    a_col_has_var = any("{{" in str((r[0] if r else "") or "") for r in rows[1:7])
    if "【" not in a1 and not a_col_has_var:
        return {}, []
    # 模板文本（A 列第 2 行起，保留 # 等分隔行）
    tpl_lines = []
    for r in rows[1:]:
        v = r[0] if r else None
        if v is None:
            continue
        s = str(v).rstrip()
        if not s.strip() and not tpl_lines:
            continue
        tpl_lines.append(s)
    tpl_text = "\n".join(tpl_lines)
    # 识别 ESN 映射表 sheet（如 lswnet：每行都是 esn=...;cfgfile=...;）——NADT 会自动生成，跳过
    tpl_effective = [x for x in tpl_text.splitlines() if x.strip()]
    # ⚠️ 必须先判非空：all([]) 恒为 True，空模板会被误判成「映射表」而丢掉「A 列无模板文本」提示
    if tpl_effective and all("esn=" in l and ("cfgfile=" in l or "netfile=" in l)
                             for l in tpl_effective):
        warnings.append(f"{sheet_name}: 检测为 ESN 映射表（{tpl_text.splitlines()[0][:60]}…），"
                        f"NADT 会按设备清单自动生成映射表，已跳过此 sheet")
        return {}, []
    # 参数名（第 1 行 C 列起，去花括号；重复警告并保留第一个）
    hdr = rows[0]
    params, seen = [], set()
    for c in range(2, len(hdr)):
        v = str(hdr[c] or "").strip()
        if not v:
            continue
        name = v.replace("{{", "").replace("}}", "").strip()
        if not name:
            continue
        if name in seen:
            warnings.append(f"{sheet_name}: 参数名 {name} 重复（第 {c + 1} 列），已忽略该列")
            continue
        seen.add(name)
        params.append(name)
    if not params:
        warnings.append(f"{sheet_name}: 无参数列（第 1 行 C 列起应为 {{参数名}}）")
        return {}, []
    if not tpl_text.strip():
        warnings.append(f"{sheet_name}: A 列无模板文本")
        return {}, []
    tname = sheet_name.strip()
    templates = {tname: tpl_text}
    # 设备（第 2 行起每行 = 一台）
    devices = []
    for r in rows[1:]:
        vals = {}
        for i, p in enumerate(params):
            c = 2 + i
            vals[p] = "" if (c >= len(r) or r[c] is None) else str(r[c]).strip()
        if not any(vals.values()):
            continue
        dev = {k: "" for k in CSV_HEADER}
        dev["type"] = tname
        # ESN / MAC 参数提升为标准字段（U盘 ini / lswnet 映射 / DHCP 绑定都读标准字段）
        for pk in ("esn", "sn", "serial", "序列号", "mac"):
            for k, v in vals.items():
                if k.lower() == pk:
                    dev[pk] = v
        host = (vals.get("sysname") or vals.get("hostname") or vals.get("name") or "").strip()
        if not host:
            host = f"{tname}-{len(devices) + 1}"
        dev["hostname"] = host
        dev["_extras"] = vals
        devices.append(dev)
    return templates, devices


def parse_xlsm_templates(path: str):
    """解析 Excel 模板文件（xlsm/xlsx，师傅的宏脚本工作流）→ (模板dict, 设备列表, 警告列表)。

    结构（每个模板一个 sheet）：
      - A1 = 【模板名】标题（或 A 列含 {{变量}} 也算模板 sheet）
      - 第 1 行 C 列起 = 参数名列头（{{sysname}} → 参数名 sysname）
      - A 列第 2 行起 = 配置模板文本（含 {{参数}} 占位符，'#' 行保留）
      - 第 2 行起每行 = 一台设备（C 列起为该行各参数值）
    """
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    skip = {"主页", "模板数据", "帮助说明", "说明", "帮助", "help"}
    templates, devices, warnings = {}, [], []
    for sn in wb.sheetnames:
        if sn.strip().lower() in skip:
            continue
        ws = wb[sn]
        rows = [[ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
                for r in range(1, ws.max_row + 1)]
        t, d = parse_sheet_rows(rows, sn, warnings)
        templates.update(t)
        devices.extend(d)
    return templates, devices, warnings


def parse_device_mapped(text: str, mapping: dict) -> list:
    """按列映射解析 CSV（mapping: {列索引: MAPPING_OPTIONS 之一}）。

    '扩展字段' 保留原表头为 key；'忽略' 跳过该列；标准字段进 dev。
    """
    reader = csv.reader(io.StringIO(text))
    rows = [r for r in reader if r and any(c.strip() for c in r)]
    if not rows:
        return []
    header = rows[0]
    devices = []
    for r in rows[1:]:
        dev = {}
        extras = {}
        for i, key in mapping.items():
            val = r[i].strip() if i < len(r) else ""
            if key == "忽略":
                continue
            if key == "扩展字段":
                h = (header[i] if i < len(header) else f"col{i + 1}").strip().lower()
                extras[h] = val
            else:
                dev[key] = val
        for k in CSV_HEADER:
            dev.setdefault(k, "")
        dev["_extras"] = extras
        # 过滤空设备（全忽略列 / 空数据行）：无主机名且无 ESN 且无 IP 的不要
        if not (dev.get("hostname") or "").strip() and not (dev.get("esn") or "").strip() \
                and not (dev.get("mgmt_ip") or "").strip():
            continue
        devices.append(dev)
    return devices


def parse_device_csv(text: str) -> list:
    """解析设备清单 CSV（首行表头，兼容缺列）。返回 dict 列表。"""
    reader = csv.reader(io.StringIO(text))
    rows = [r for r in reader if r and any(c.strip() for c in r)]
    if not rows:
        return []
    header = [h.strip().lower() for h in rows[0]]
    # 表头规范化：支持 中文/英文 常见别名
    header = [CSV_ALIASES.get(h, h) for h in header]
    devices = []
    for r in rows[1:]:
        dev = {}
        extras = {}
        for i, k in enumerate(header):
            val = r[i] if i < len(r) else ""
            if k in CSV_HEADER:
                dev[k] = val
            else:
                extras[k] = val                 # 任意额外列 → 模板变量
        for k in CSV_HEADER:
            dev.setdefault(k, "")
        dev["_extras"] = extras
        devices.append(dev)
    return devices


def load_devices_csv(path: str) -> list:
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        return parse_device_csv(f.read())


def load_devices_excel(path: str) -> list:
    """Excel 导入（需 openpyxl）。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = [[("" if c is None else str(c)).strip() for c in row] for row in ws.iter_rows(values_only=True)]
    text = "\n".join(",".join(r) for r in rows)
    return parse_device_csv(text)


def devices_to_csv(devices: list) -> str:
    """导出 CSV：标准列 + 所有出现过的扩展列。"""
    buf = io.StringIO()
    w = csv.writer(buf)
    extras_keys: list = []
    for d in devices:
        for k in (d.get("_extras") or {}):
            if k not in extras_keys:
                extras_keys.append(k)
    header = CSV_HEADER + extras_keys
    w.writerow(header)
    for d in devices:
        row = [d.get(k, "") for k in CSV_HEADER]
        ex = d.get("_extras") or {}
        row += [ex.get(k, "") for k in extras_keys]
        w.writerow(row)
    return buf.getvalue()


def dhcp_snippet(tftp_ip: str, bootfile: str = BOOTFILE,
                 mgmt_vlan: str = "11", network: str = "192.168.100.0",
                 gateway: str = "192.168.100.254", mode: str = "67",
                 fileserver_url: str = "") -> str:
    """生成 DHCP 配置片段。

    mode="146" → 华为 EasyDeploy 标准（option 66 文件服务器 + option 146 netfile），
                 需 fileserver_url（如 sftp://user:pass@ip:port，空则用 tftp://ip）；
    mode="67"  → 通用 PXE 标准（option 66 ip + option 67 引导文件名）。
    """
    if mode == "146":
        url = (fileserver_url or "").strip() or f"tftp://{tftp_ip}"
        opt66 = f"option 66 ascii {url}"
        opt67 = f"option 146 ascii opervalue=1;delaytime=0;netfile={bootfile};"
        title = "option 66 + option 146（华为 EasyDeploy）"
        note67 = ("option 146 ascii opervalue=1;delaytime=0;netfile={bootfile}   # EasyDeploy 参数")
        note66 = f"option 66 ascii {url}"
    else:
        opt66 = f"option 66 ip-address {tftp_ip}"
        opt67 = f"option 67 ascii {bootfile}"
        title = "option 66 + option 67（通用 PXE）"
        note66 = f"option 66 ip-address {tftp_ip};"
        note67 = f'option 67 "{bootfile}";'
    return f"""# ============================================================
# DHCP 配置（华为 VRP / S 系列交换机作 DHCP 服务器）— {title}
# 1. 全局开启 DHCP
dhcp enable

# 2. 管理网段 VLANIF 启用 DHCP（按实际 VLAN 改）
interface Vlanif {mgmt_vlan}
 dhcp select global

# 3. 配置地址池（按实际网段改）
ip pool mgmt
 network {network} mask 255.255.255.0
 gateway-list {gateway}
 {opt66}       # ⬅ 文件服务器
 {opt67}       # ⬅ 引导文件 = ESN 映射表

# 4. （可选）按 MAC 绑定固定 IP
#    static-bind ip-address 192.168.100.2 mac-address xxxx-xxxx-xxxx

# ============================================================
# 通用说明（其它 DHCP 服务器）
#   ISC dhcpd:   {note66}  {note67}
#   dnsmasq:     dhcp-option=66,{tftp_ip}   dhcp-option=67,"{bootfile}"
#   Windows DHCP: 服务器选项 → 066 引导服务器 = {tftp_ip}
#                 服务器选项 → 067 引导文件名 = {bootfile}
# ============================================================
"""


def normalize_mac(mac: str, style: str = "colon") -> str:
    """MAC 规范化：输入任意格式（aa:bb:cc:dd:ee:ff / aabb.ccdd.eeff / 连字符），
    style="colon" → aa:bb:cc:dd:ee:ff（ISC dhcpd/思科），
    style="huawei" → aaaa-bbbb-cccc（华为 EasyDeploy ini）。非法输入返回 ""。
    """
    m = re.sub(r"[^0-9a-fA-F]", "", mac or "")
    if len(m) != 12:
        return ""
    m = m.lower()
    if style == "huawei":
        return "-".join(m[i:i + 4] for i in range(0, 12, 4))
    return ":".join(m[i:i + 2] for i in range(0, 12, 2))


# ═══════════════════════════════════════════════════════════════
# 厂商预设 / Vendor presets（核心引擎厂商无关，预设只是参数集）
# 选预设自动填充: 引导文件名 / 是否生成映射表 / DHCP 输出风格
# 所有参数在 GUI 里都可手动覆盖
# ═══════════════════════════════════════════════════════════════
VENDOR_PRESETS = {
    "generic": {
        "name": "通用 Generic（按 MAC 绑定，推荐）",
        "bootfile": "",
        "mapping": False,
        "dhcp": "isc",
    },
    "huawei": {
        "name": "华为 Huawei（EasyDeploy ESN 映射表）",
        "bootfile": "lswnet.cfg",
        "mapping": True,
        "dhcp": "vrp146",
    },
    "cisco": {
        "name": "思科 Cisco（auto-install，按 MAC 绑定）",
        "bootfile": "network-confg",
        "mapping": False,
        "dhcp": "isc",
    },
}


DHCP_STYLES = {
    "isc": "ISC dhcpd（Linux 通用 / 思科 auto-install）",
    "dnsmasq": "dnsmasq（轻量通用）",
    "vrp146": "华为 VRP option 146（EasyDeploy）",
    "vrp67": "华为 VRP option 67（通用）",
}


def build_dhcp_dnsmasq(devices: list, tftp_ip: str, bootfile: str = "",
                       network: str = "", gateway: str = "",
                       range_start: str = "", range_end: str = "") -> str:
    """生成 dnsmasq 配置 — 通用 DHCP option 66/67 开局（思科 auto-install 可用）。"""
    out = [
        "# ============================================================",
        "# dnsmasq 配置 — DHCP option 66/67 开局（通用 / 思科 auto-install）",
        "# 用法: 写入 /etc/dnsmasq.conf 或 dnsmasq 配置文件，重启服务生效",
        "# ============================================================",
        "# 文件服务器（option 66）",
        f"dhcp-option=66,{tftp_ip}",
    ]
    if bootfile:
        out.append('# 引导文件（option 67，华为映射表模式；思科 auto-install 用下面 host 段）')
        out.append(f'dhcp-option=67,"{bootfile}"')
    if network and gateway and range_start and range_end:
        out += [
            "",
            "# 管理网段",
            f"dhcp-range={range_start},{range_end},12h",
            f"dhcp-option=3,{gateway}",
        ]
    bound = 0
    for dev in devices:
        d = derive_defaults(dev)
        ex = dev.get("_extras") or {}
        mac = normalize_mac(d.get("mac") or ex.get("mac") or "", "colon")
        host = (d.get("hostname") or "").strip()
        ip = (d.get("mgmt_ip") or "").strip()
        if not mac or not host:
            continue
        tag = host.replace(".", "_").replace("-", "_")
        out += [
            "",
            f"# {host} — 按 MAC 绑定专属配置（思科 auto-install）",
            f'dhcp-host={mac},set:{tag}' + (f",{ip}" if ip else ""),
            f'dhcp-option=tag:{tag},67,"{host}.cfg"',
        ]
        bound += 1
    out.append("")
    out.append(f"# 已生成 {bound} 条 MAC 绑定；"
               f"{len(devices) - bound} 台设备缺 MAC（清单加 mac 列才会生成绑定）")
    return "\n".join(out) + "\n"


def build_dhcp_isc(devices: list, tftp_ip: str, bootfile: str = "",
                   network: str = "", gateway: str = "",
                   netmask: str = "255.255.255.0", range_start: str = "",
                   range_end: str = "", dns: str = "", domain: str = "") -> str:
    """生成 ISC dhcpd 配置（思科 auto-install / 通用 DHCP 开局）。

    有 mac 的设备生成 host 绑定段（每台独立 bootfile-name=主机名.cfg）；
    无 mac 的设备用全局 bootfile（华为 lswnet.cfg 映射表模式）。
    """
    out = [
        "# ============================================================",
        "# ISC dhcpd 配置 — 通用 DHCP 开局（思科 auto-install / 华为通用模式）",
        "# 用法: 贴入 /etc/dhcp/dhcpd.conf 或 dhcpd 配置，重启服务生效",
        "# ============================================================",
        "# 1. 全局 option（文件服务器 = 本机）",
        f'option tftp-server-name "{tftp_ip}";',            # option 66
    ]
    if bootfile:
        out.append(f'option bootfile-name "{bootfile}";'
                   "    # 华为映射表模式（思科 auto-install 不用全局，看下面 host 段）")
    if network and gateway:
        out += [
            "",
            "# 2. 管理网段",
            f"subnet {network} netmask {netmask} {{",
        ]
        if range_start and range_end:
            out.append(f"  range {range_start} {range_end};")
        out.append(f"  option routers {gateway};")
        if dns:
            out.append(f"  option domain-name-servers {dns};")
        if domain:
            out.append(f'  option domain-name "{domain}";')
        out.append("}")
    # 3. 每台设备按 MAC 绑定（思科 auto-install 核心）
    bound = 0
    for dev in devices:
        d = derive_defaults(dev)
        ex = dev.get("_extras") or {}
        mac = normalize_mac(d.get("mac") or ex.get("mac") or "", "colon")
        host = (d.get("hostname") or "").strip()
        ip = (d.get("mgmt_ip") or "").strip()
        if not mac or not host:
            continue
        out += ["", f"host {host} {{",
                f"  hardware ethernet {mac};"]
        if ip:
            out.append(f"  fixed-address {ip};")
        out.append(f'  option bootfile-name "{host}.cfg";   # ⬅ 该设备专属配置文件')
        out.append("}")
        bound += 1
    out.append("")
    out.append(f"# 已生成 {bound} 条 MAC 绑定；"
               f"{len(devices) - bound} 台设备缺 MAC（清单加 mac 列才会生成绑定）")
    return "\n".join(out) + "\n"


def get_local_ips() -> list:
    """获取本机 IPv4 地址列表。"""
    ips = []
    try:
        host = socket.gethostname()
        for info in socket.getaddrinfo(host, None, socket.AF_INET):
            ip = info[4][0]
            if ip not in ips and not ip.startswith("127."):
                ips.append(ip)
    except Exception:
        pass
    return ips


def parse_extras_text(text: str) -> dict:
    """解析扩展字段文本（每行一个 K:V，回车分隔；支持 = / : / ： 分隔符）。"""
    extras = {}
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        sep = next((s for s in ("=", ":", "：") if s in line), None)
        if sep:
            k, _, v = line.partition(sep)
            extras[k.strip()] = v.strip()
        else:
            extras[line] = ""
    return extras


def format_extras_text(extras: dict) -> str:
    return "\n".join(f"{k}={v}" for k, v in (extras or {}).items())


def list_templates() -> list:
    """列出模板库中的模板文件（*.cfg）。"""
    try:
        names = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith(".cfg")]
        return sorted(names)
    except OSError:
        return []


def load_template_mapping() -> dict:
    """读取 类型→模板 映射（templates/mapping.json）。"""
    path = os.path.join(TEMPLATE_DIR, "mapping.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_template_mapping(mapping: dict):
    path = os.path.join(TEMPLATE_DIR, "mapping.json")
    try:
        os.makedirs(TEMPLATE_DIR, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def find_template_for_type(devtype: str, mapping: dict, fallback: str) -> str:
    """类型 → 模板文件名；未映射用 fallback。"""
    t = (devtype or "").strip().upper()
    if t and t in mapping:
        return mapping[t]
    return fallback


def get_template_path(name: str) -> str:
    """模板文件名 → 完整路径（防目录穿越）。"""
    return os.path.join(TEMPLATE_DIR, os.path.basename(name))


# ═══════════════════════════════════════════════════════════════
# TFTP 服务器 / TFTP Server (asyncio 异步引擎, RFC 1350)
# 千并发场景：单事件循环多路复用 + 内存文件缓存 + 并发上限
# ═══════════════════════════════════════════════════════════════

class TFTPTransfer(asyncio.DatagramProtocol):
    """单次 RRQ 传输会话（独立 TID 端口，异步逐块发送+ACK+超时重传）。"""

    def __init__(self, server, file_bytes: bytes, client, fname: str, log, done):
        self.server = server
        self.data = file_bytes
        self.client = client
        self.fname = fname
        self.log = log
        self.done = done
        self.transport = None
        self.block = 0
        self._ack_future = None

    def connection_made(self, transport):
        self.transport = transport
        task = asyncio.get_event_loop().create_task(self._run())
        self.server._tasks.add(task)         # 登记，停止时可取消
        task.add_done_callback(self.server._tasks.discard)

    def datagram_received(self, data, addr):
        if addr[0] != self.client[0]:
            return
        try:
            op, blk = struct.unpack("!HH", data[:4])
        except Exception:
            return
        if op == 4 and blk == self.block and self._ack_future and not self._ack_future.done():
            self._ack_future.set_result(True)

    async def _run(self):
        loop = asyncio.get_event_loop()
        offset, block, size = 0, 1, len(self.data)
        try:
            while True:
                chunk = self.data[offset:offset + 512]
                pkt = struct.pack("!HH", 3, block) + chunk
                acked = False
                for _ in range(6):                     # 超时重传 6 次
                    self.transport.sendto(pkt, self.client)
                    self.block = block
                    fut = loop.create_future()
                    self._ack_future = fut
                    try:
                        await asyncio.wait_for(fut, timeout=1.0)
                        acked = True
                        break
                    except asyncio.TimeoutError:
                        continue
                if not acked:
                    self.log(f"{self.client[0]}: RRQ {self.fname} → ACK 超时中止")
                    return
                if len(chunk) < 512:
                    break
                offset += 512
                block += 1
            self.log(f"{self.client[0]}: RRQ {self.fname} 完成（{size} B, {block} 块）")
            self.server.record_deployed(self.fname, self.client[0])
        finally:
            try:
                self.transport.close()
            except Exception:
                pass
            try:
                self.done.set()
            except Exception:
                pass

    def error_received(self, exc):
        pass

    def connection_lost(self, exc):
        try:
            self.done.set()
        except Exception:
            pass


class TFTPListener(asyncio.DatagramProtocol):
    """端口 69 监听器：收到 RRQ 后交给服务器异步处理。"""

    def __init__(self, server):
        self.server = server

    def datagram_received(self, data, addr):
        self.server._on_datagram(data, addr)

    def error_received(self, exc):
        pass


class TFTPServer:
    """异步 TFTP 服务器：单事件循环处理数千并发传输，文件内存缓存。"""

    def __init__(self, root_dir: str, log_callback=None, port: int = 69,
                 max_concurrent: int = 500):
        self.root_dir = root_dir
        self.log = log_callback or (lambda msg: None)
        self.port = port
        self.max_concurrent = max_concurrent
        self.running = False
        self._thread: threading.Thread | None = None
        self._loop = None
        self._listener_transport = None
        self._sem = None
        self._cache: dict = {}
        self._tasks: set = set()
        self.deployed: list = []           # 开局记录 [(时间, 文件名, 客户端IP)]
        self._deployed_lock = threading.Lock()

    def record_deployed(self, fname: str, ip: str):
        """记录一次成功下发的文件（开局完成）。"""
        with self._deployed_lock:
            self.deployed.append((datetime.now().strftime("%Y-%m-%d %H:%M:%S"), fname, ip))

    def start(self):
        if self.running:
            return
        if not os.path.isdir(self.root_dir):
            self.log(f"TFTP 根目录不存在: {self.root_dir}")
            return
        self._preload()
        self.running = True
        self._thread = threading.Thread(target=self._run_loop, daemon=True)
        self._thread.start()
        # 等待监听就绪（最多 3 秒）
        for _ in range(30):
            if self._listener_transport:
                break
            time.sleep(0.1)
        if self._listener_transport:
            self.log(f"TFTP server 已启动 0.0.0.0:{self.port}  root={self.root_dir}  "
                     f"(异步引擎, 并发上限 {self.max_concurrent})")
        else:
            self.log(f"TFTP 启动失败（端口 {self.port} 可能被占用）")

    def stop(self):
        self.running = False
        if self._loop:
            try:
                self._loop.call_soon_threadsafe(self._shutdown_loop)
            except Exception:
                pass
        if self._thread:
            self._thread.join(timeout=3)
            self._thread = None
        self.log("TFTP server 已停止")

    def _shutdown_loop(self):
        for t in list(self._tasks):
            t.cancel()
        if self._listener_transport:
            try:
                self._listener_transport.close()
            except Exception:
                pass
            self._listener_transport = None

    def _run_loop(self):
        self._loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self._loop)
        try:
            self._loop.run_until_complete(self._serve())
        except Exception as e:
            self.log(f"TFTP 事件循环异常: {e}")
        finally:
            self._loop.close()
            self._loop = None

    async def _serve(self):
        loop = self._loop
        self._sem = asyncio.Semaphore(self.max_concurrent)
        try:
            transport, _ = await loop.create_datagram_endpoint(
                lambda: TFTPListener(self), local_addr=("0.0.0.0", self.port))
            self._listener_transport = transport
        except OSError as e:
            self.log(f"TFTP 监听失败: {e}")
            self.running = False            # 启动失败必须复位，否则 GUI 误显示"运行中"
            return
        try:
            while self.running:
                await asyncio.sleep(0.2)
        finally:
            try:
                transport.close()
            except Exception:
                pass
            self._listener_transport = None

    def _on_datagram(self, data, addr):
        if not self.running:
            return
        try:
            opcode = struct.unpack("!H", data[:2])[0]
        except Exception:
            return
        if opcode != 1:
            self.log(f"{addr[0]}: 仅支持读请求 RRQ")
            return
        task = self._loop.create_task(self._handle_rrq(data, addr))
        self._tasks.add(task)
        task.add_done_callback(self._tasks.discard)

    async def _handle_rrq(self, data, addr):
        parts = data[2:].split(b"\x00", 2)
        if len(parts) < 2:
            return
        fname = parts[0].decode("utf-8", "replace").strip()
        safe = os.path.basename(fname)               # 防目录穿越
        file_bytes = self._get_file(safe)
        if file_bytes is None:
            self.log(f"{addr[0]}: RRQ {safe} → 文件不存在")
            self._send_error(addr, 1, "File not found")
            return
        if self._sem.locked():
            self.log(f"{addr[0]}: RRQ {safe} → 服务器繁忙（客户端稍后重试）")
            self._send_error(addr, 0, "Server busy")
            return
        async with self._sem:
            done = asyncio.Event()
            try:
                await self._loop.create_datagram_endpoint(
                    lambda: TFTPTransfer(self, file_bytes, addr, safe, self.log, done),
                    local_addr=("0.0.0.0", 0))
            except OSError as e:
                self.log(f"{addr[0]}: 会话创建失败 {e}")
                return
            await done.wait()

    def _send_error(self, addr, code: int, msg: str):
        if not self._listener_transport:
            return
        pkt = struct.pack("!HH", 5, code) + msg.encode("utf-8", "replace") + b"\x00"
        try:
            self._listener_transport.sendto(pkt, addr)
        except Exception:
            pass

    def _get_file(self, name: str) -> bytes | None:
        if name in self._cache:
            return self._cache[name]
        path = os.path.join(self.root_dir, name)
        if not os.path.isfile(path):
            return None
        try:
            with open(path, "rb") as f:
                data = f.read()
            self._cache[name] = data
            return data
        except OSError:
            return None

    def _preload(self):
        """预热内存缓存：小文件全量载入；总量超 200MB 则改按需加载。"""
        total, count = 0, 0
        try:
            names = os.listdir(self.root_dir)
        except OSError:
            return
        for fn in names:
            fp = os.path.join(self.root_dir, fn)
            if not os.path.isfile(fp):
                continue
            try:
                total += os.path.getsize(fp)
            except OSError:
                continue
            if total > 200 * 1024 * 1024:
                self._cache = {}
                self.log("TFTP 文件总量过大，改用按需加载")
                return
            try:
                with open(fp, "rb") as f:
                    self._cache[fn] = f.read()
                count += 1
            except OSError:
                pass
        if count:
            self.log(f"TFTP 缓存预热: {count} 个文件 / {total / 1024:.0f} KB")


# ═══════════════════════════════════════════════════════════════
# 主题 / Theme（蓝色，与 NIS 一致）
# ═══════════════════════════════════════════════════════════════

def setup_app_style(root: tk.Tk) -> None:
    BG = "#f3f5f9"; PANEL = "#ffffff"; FG = "#1f2937"; MUTED = "#64748b"
    ACCENT = "#2563eb"; ACCENT2 = "#3b82f6"; ACCENTD = "#1e40af"
    LINE = "#d7dde8"; HEADER = "#eef2f9"; ROW_SEL = "#dbeafe"; DANGER = "#dc2626"
    GREEN = "#16a34a"
    FONT = ("Microsoft YaHei UI", 9); FONT_B = ("Microsoft YaHei UI", 9, "bold")
    FONT_H = ("Microsoft YaHei UI", 10, "bold")
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    root.configure(bg=BG)
    style.configure(".", background=BG, foreground=FG, bordercolor=LINE,
                    lightcolor=BG, darkcolor=BG, focuscolor=ACCENT)
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG, foreground=FG, font=FONT)
    style.configure("TLabelframe", background=BG, bordercolor=LINE, relief="solid", borderwidth=1)
    style.configure("TLabelframe.Label", background=BG, foreground=ACCENTD, font=FONT_H)
    style.configure("TButton", background="#e8edf5", foreground=FG, borderwidth=0,
                    focusthickness=0, padding=(12, 7), font=FONT)
    style.map("TButton", background=[("pressed", ACCENTD), ("active", "#dbe6f8")],
              foreground=[("pressed", "white")])
    style.configure("Accent.TButton", background=ACCENT, foreground="white",
                    padding=(14, 8), font=FONT_B)
    style.map("Accent.TButton", background=[("pressed", ACCENTD), ("active", ACCENT2)],
              foreground=[("pressed", "white"), ("disabled", "#bfdbfe")])
    style.configure("Danger.TButton", background=DANGER, foreground="white", padding=(12, 7), font=FONT_B)
    style.map("Danger.TButton", background=[("pressed", "#991b1b"), ("active", "#ef4444")],
              foreground=[("disabled", "#fecaca")])
    style.configure("GreenAccent.TButton", background=GREEN, foreground="white",
                    padding=(14, 8), font=FONT_B)
    style.map("GreenAccent.TButton", background=[("pressed", "#15803d"), ("active", "#4ade80")],
              foreground=[("pressed", "white"), ("disabled", "#bbf7d0")])
    style.configure("TEntry", fieldbackground="white", foreground=FG, bordercolor=LINE, padding=4, insertcolor=ACCENT)
    style.configure("TCombobox", fieldbackground="white", foreground=FG, bordercolor=LINE, padding=4, arrowcolor=ACCENT)
    style.configure("TCheckbutton", background=BG, foreground=FG, font=FONT, focuscolor=ACCENT)
    style.configure("TRadiobutton", background=BG, foreground=FG, font=FONT, focuscolor=ACCENT)
    style.configure("TNotebook", background=BG, borderwidth=0)
    style.configure("TNotebook.Tab", background="#e2e8f0", foreground=MUTED,
                    padding=(16, 8), borderwidth=0, font=FONT)
    style.map("TNotebook.Tab", background=[("selected", PANEL)],
              foreground=[("selected", ACCENT)], font=[("selected", FONT_H)])
    style.configure("Treeview", background="white", fieldbackground="white",
                    foreground=FG, rowheight=26, borderwidth=0, font=FONT)
    style.configure("Treeview.Heading", background=HEADER, foreground="#334155",
                    padding=(6, 5), relief="flat", font=FONT_B)
    style.map("Treeview", background=[("selected", ROW_SEL)], foreground=[("selected", FG)])
    style.configure("Vertical.TScrollbar", background="#cbd5e1", troughcolor=BG, borderwidth=0, arrowsize=12)
    style.configure("Horizontal.TScrollbar", background="#cbd5e1", troughcolor=BG, borderwidth=0, arrowsize=12)


# ═══════════════════════════════════════════════════════════════
# 主 GUI / Main GUI
# ═══════════════════════════════════════════════════════════════

class NADTGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("NADT — Network Automation Deployment Tool")
        self.root.geometry("1180x780")
        self.root.minsize(960, 640)
        setup_app_style(root)

        self.devices: list = []            # 设备清单 / device list
        self.output_dir = OUTPUT_DIR
        self.tftp = TFTPServer(OUTPUT_DIR, log_callback=self._tftp_log)
        self.template_mapping: dict = load_template_mapping()
        self.current_template = "access-switch.cfg"
        self._page = 0                      # 设备表分页状态
        self._page_size = 200
        self._filtered_idx: list | None = None
        self._checked: set = set()          # 勾选（checkbox 多选）的设备索引
        self._undo_stack: list = []         # 撤销栈（最多 20 步）
        self._last_gen_dir: str | None = None   # 最近一次生成的实际输出目录（归档子目录）

        # 表格工作台数据（Excel 宏式：A 列模板文本 + 第 1 行参数名 + 每行一台设备）
        self.sheet_rows: list = [
            ["【工作表】", "文本备注", "{{sysname}}", "{{vlan}}", "{{ip}}", "{{wg}}", "{{vlanif}}"],
            ["#", "1楼", "SW-1-1F", "10 20", "192.168.10.11", "192.168.10.254", "10"],
            ["sysname {{sysname}}", "", "SW-2-2F", "20 30", "192.168.20.11", "192.168.20.254", "20"],
            ["#", "", "", "", "", "", ""],
            ["vlan batch {{vlan}}", "", "", "", "", "", ""],
            ["#", "", "", "", "", "", ""],
            ["interface Vlanif{{vlanif}}", "", "", "", "", "", ""],
            [" ip address {{ip}} 255.255.255.0", "", "", "", "", "", ""],
            ["#", "", "", "", "", "", ""],
            ["ip route-static 0.0.0.0 0.0.0.0 {{wg}}", "", "", "", "", "", ""],
            ["#", "", "", "", "", "", ""],
            ["return", "", "", "", "", "", ""],
            ["save", "", "", "", "", "", ""],
            ["y", "", "", "", "", "", ""],
        ]
        self._sheet_cursor: tuple = (0, 0)   # 表格右键菜单定位 (row, col)

        self._build_ui()
        self._load_template(self.current_template)
        self.root.bind("<Control-z>", lambda e: self._on_undo())
        self.root.bind("<Control-Z>", lambda e: self._on_undo())
        loaded = load_devices_db()          # 从本地设备库恢复
        if loaded:
            self.devices = loaded
            self._set_status(f"已从本地设备库加载 {len(loaded)} 台设备 (devices.db)")
        self._update_dev_table()
        self._update_ip_display()

    # ── 顶层布局 / top-level layout ──
    def _build_ui(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.tab_sheet    = ttk.Frame(nb); nb.add(self.tab_sheet,    text="◎ 工作表 Sheet（推荐）")
        self.tab_devices  = ttk.Frame(nb); nb.add(self.tab_devices,  text="① 设备清单 Devices")
        self.tab_template = ttk.Frame(nb); nb.add(self.tab_template, text="② 配置模板 Template")
        self.tab_generate = ttk.Frame(nb); nb.add(self.tab_generate, text="③ 脚本生成 Generate")
        self.tab_tftp     = ttk.Frame(nb); nb.add(self.tab_tftp,     text="④ TFTP 服务器")
        self.tab_guide    = ttk.Frame(nb); nb.add(self.tab_guide,    text="⑤ 部署说明 Guide")

        self._build_sheet_tab()
        self._build_devices_tab()
        self._build_template_tab()
        self._build_generate_tab()
        self._build_tftp_tab()
        self._build_guide_tab()

        # 底部状态栏
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN,
                  anchor=tk.W, padding=(5, 2)).pack(fill=tk.X, side=tk.BOTTOM)

    # ── ◎ 工作表 / sheet workbench tab（Excel 宏式表格，推荐入口）──
    def _build_sheet_tab(self):
        bar = ttk.Frame(self.tab_sheet)
        bar.pack(fill=tk.X, padx=10, pady=(10, 6))
        ttk.Label(bar, text="表格工作台：A 列 = 配置模板文本，第 1 行 C 列起 = 参数名，每行 = 一台设备（双击单元格编辑）",
                  foreground="#1e40af").pack(side=tk.LEFT)
        ttk.Button(bar, text="导入 xlsx", command=self._on_sheet_import).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(bar, text="保存 xlsx", command=self._on_sheet_save).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(bar, text="添加设备行", command=self._on_sheet_add_dev_row).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(bar, text="生成脚本", style="Accent.TButton",
                   command=self._on_sheet_generate).pack(side=tk.RIGHT, padx=(6, 0))

        # 表格
        wrap = ttk.Frame(self.tab_sheet)
        wrap.pack(fill=tk.BOTH, expand=True, padx=10)
        self.sheet_tree = ttk.Treeview(wrap, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(wrap, orient=tk.VERTICAL, command=self.sheet_tree.yview)
        hsb = ttk.Scrollbar(wrap, orient=tk.HORIZONTAL, command=self.sheet_tree.xview)
        self.sheet_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.sheet_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.sheet_tree.bind("<Double-1>", self._on_sheet_dclick)
        self.sheet_tree.bind("<Button-3>", self._on_sheet_rmb)

        # 右键菜单
        self.sheet_menu = tk.Menu(self.root, tearoff=0)
        self.sheet_menu.add_command(label="编辑单元格…", command=lambda: self._on_sheet_edit_at(*self._sheet_cursor))
        self.sheet_menu.add_separator()
        self.sheet_menu.add_command(label="↓ 插入行", command=self._on_sheet_ins_row)
        self.sheet_menu.add_command(label="✕ 删除行", command=self._on_sheet_del_row)
        self.sheet_menu.add_command(label="⇥ 添加设备行（末尾）", command=self._on_sheet_add_dev_row)
        self.sheet_menu.add_separator()
        self.sheet_menu.add_command(label="+ 追加参数列", command=self._on_sheet_add_col)
        self.sheet_menu.add_command(label="✕ 删除最后参数列", command=self._on_sheet_del_col)
        self.sheet_menu.add_separator()
        self.sheet_menu.add_command(label="清空表格", command=self._on_sheet_clear)

        # 日志
        self.sheet_logbox = scrolledtext.ScrolledText(self.tab_sheet, height=6, font=("Consolas", 9),
                                                      state=tk.DISABLED, bg="#fbfdf9")
        self.sheet_logbox.pack(fill=tk.X, padx=10, pady=(6, 10))
        self._refresh_sheet()

    # ── 表格数据刷新 ──
    def _refresh_sheet(self):
        """按 self.sheet_rows 重建列与行。"""
        tree = self.sheet_tree
        cols = ["A", "B"] + [f"P{i}" for i in range(2, max(len(self.sheet_rows[0]) if self.sheet_rows else 0, 2))]
        tree.configure(columns=cols)
        # 列头
        heads = ["A·模板文本", "B·备注"]
        row0 = self.sheet_rows[0] if self.sheet_rows else []
        for i in range(2, len(cols)):
            h = str(row0[i]) if i < len(row0) and str(row0[i]).strip() else f"参数{i - 1}"
            heads.append(h[:20])
        for c, h in zip(cols, heads):
            tree.heading(c, text=h)
            tree.column(c, anchor=tk.W, stretch=(c == "A"))
        self._auto_col_widths()
        tree.delete(*tree.get_children())
        for r_i, row in enumerate(self.sheet_rows):
            values = [str(row[c]) if c < len(row) else "" for c in range(len(cols))]
            tree.insert("", tk.END, iid=str(r_i), values=values)
            if r_i == 0:
                tree.tag_configure("hdr", background="#eef2ff", font=("Microsoft YaHei UI", 9, "bold"))
                tree.item(str(r_i), tags=("hdr",))

    def _auto_col_widths(self):
        """按内容自适应列宽（中文按 2 字符宽算）；A 列限宽不占满，其它列留出可视空间。

        用户仍可手动拖列头边界调宽（Treeview 原生支持）。
        """
        tree = self.sheet_tree
        cols = tree["columns"]

        def disp_len(s):
            return sum(2 if ord(ch) > 127 else 1 for ch in str(s))

        for c in cols:
            idx = cols.index(c)
            maxlen = disp_len(tree.heading(c, "text"))
            for row in self.sheet_rows:
                v = str(row[idx]) if idx < len(row) else ""
                if v:
                    maxlen = max(maxlen, disp_len(v))
            if c == "A":
                width = min(max(maxlen, 18) * 8, 380)     # 模板文本列：上限 380，留出右侧空间
            elif c == "B":
                width = min(max(maxlen, 4) * 8, 120)
            else:
                width = min(max(maxlen, 5) * 8, 160)
            tree.column(c, width=width)

    def _sheet_pos(self, ev):
        """鼠标位置 → (row, col) 索引。"""
        row_iid = self.sheet_tree.identify_row(ev.y)
        col = self.sheet_tree.identify_column(ev.x)
        r = int(row_iid) if row_iid else 0
        c = int(col.replace("#", "")) - 1 if col else 0
        return r, max(c, 0)

    def _on_sheet_dclick(self, ev):
        r, c = self._sheet_pos(ev)
        self._sheet_cursor = (r, c)
        self._on_sheet_edit_at(r, c)

    def _on_sheet_rmb(self, ev):
        r, c = self._sheet_pos(ev)
        self._sheet_cursor = (r, c)
        try:
            self.sheet_menu.tk_popup(ev.x_root, ev.y_root)
        finally:
            self.sheet_menu.grab_release()

    def _on_sheet_edit_at(self, r, c):
        """双击单元格 → 弹编辑框（支持多行模板文本）。"""
        if not self.sheet_rows:
            self.sheet_rows = [[""] * (c + 1)]
        while len(self.sheet_rows) <= r:
            self.sheet_rows.append([""] * (c + 1))
        while len(self.sheet_rows[r]) <= c:
            self.sheet_rows[r].append("")
        win = tk.Toplevel(self.root)
        win.title(f"编辑单元格 行{r + 1} 列{c + 1}")
        win.geometry("520x220")
        win.transient(self.root)
        txt = scrolledtext.ScrolledText(win, font=("Consolas", 10), height=6)
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        txt.insert("1.0", self.sheet_rows[r][c])
        btns = ttk.Frame(win)
        btns.pack(pady=(0, 8))
        def ok():
            self.sheet_rows[r][c] = txt.get("1.0", tk.END).rstrip("\n")
            self._refresh_sheet()
            win.destroy()
        def cancel():
            win.destroy()
        ttk.Button(btns, text="确定", style="Accent.TButton", command=ok).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="取消", command=cancel).pack(side=tk.LEFT, padx=4)
        txt.focus_set()

    def _on_sheet_ins_row(self):
        """当前行下方插入一行（复制上一行结构）。"""
        r = self._sheet_cursor[0]
        ncol = max(len(rw) for rw in self.sheet_rows) if self.sheet_rows else 3
        if not self.sheet_rows:
            self.sheet_rows = [[""] * ncol]
            r = 0
        self.sheet_rows.insert(r + 1, [""] * ncol)
        self._refresh_sheet()

    def _on_sheet_del_row(self):
        r = self._sheet_cursor[0]
        if len(self.sheet_rows) <= 2:
            messagebox.showinfo("提示", "至少保留表头 + 一行")
            return
        self.sheet_rows.pop(r)
        self._refresh_sheet()

    def _on_sheet_add_dev_row(self):
        """末尾添加一台设备（参数列空，等待填写）。"""
        if not self.sheet_rows:
            self.sheet_rows = [[""] * 3]
        ncol = max(len(rw) for rw in self.sheet_rows)
        self.sheet_rows.append([""] * ncol)
        self._refresh_sheet()

    def _on_sheet_add_col(self):
        """末尾追加参数列。"""
        if not self.sheet_rows:
            self.sheet_rows = [[""] * 3]
        ncol = max(len(rw) for rw in self.sheet_rows)
        for rw in self.sheet_rows:
            rw.append("")
        if len(self.sheet_rows[0]) <= 2:
            self.sheet_rows[0] = [self.sheet_rows[0][0], self.sheet_rows[0][1]] + [""]
        self.sheet_rows[0][ncol] = f"{{{{new_param_{ncol - 1}}}}}"
        self._refresh_sheet()

    def _on_sheet_del_col(self):
        ncol = max(len(rw) for rw in self.sheet_rows) if self.sheet_rows else 0
        if ncol <= 3:
            messagebox.showinfo("提示", "至少保留 A 模板 / B 备注 / 1 个参数列")
            return
        for rw in self.sheet_rows:
            if len(rw) > ncol - 1:
                rw.pop(ncol - 1)
        self._refresh_sheet()

    def _on_sheet_clear(self):
        if messagebox.askyesno("清空表格", "清空表格内容？（保留表头）"):
            self.sheet_rows = self.sheet_rows[:1] if self.sheet_rows else [[""] * 3]
            self._refresh_sheet()

    def _sheet_log(self, msg: str):
        self.sheet_logbox.config(state=tk.NORMAL)
        self.sheet_logbox.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.sheet_logbox.see(tk.END)
        self.sheet_logbox.config(state=tk.DISABLED)

    def _sheet_name(self) -> str:
        """模板名：A1 的【】内文本，否则 '工作表'。"""
        a1 = str(self.sheet_rows[0][0] or "") if self.sheet_rows else ""
        m = re.search(r"【(.+?)】", a1)
        return m.group(1).strip() if m else "工作表"

    def _on_sheet_import(self):
        """导入 xlsx → 表格（取第一个可识别模板 sheet）。"""
        try:
            import openpyxl  # noqa
        except ImportError:
            messagebox.showerror("缺少依赖", "需要 openpyxl\n请运行: pip install openpyxl")
            return
        path = filedialog.askopenfilename(title="选择 Excel 模板", filetypes=[("Excel 模板", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            tpls, devs, warns = parse_xlsm_templates(path)
        except Exception as e:
            messagebox.showerror("解析失败", str(e))
            return
        if not tpls:
            messagebox.showwarning("未识别", "没找到模板 sheet（A1 需为【模板名】或 A 列含 {{变量}}）")
            return
        # 重新读第一个模板 sheet 的原始行
        wb = openpyxl.load_workbook(path, data_only=True)
        sn = next(iter(tpls.keys()))
        ws = wb[sn]
        rows = [[ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
                for r in range(1, ws.max_row + 1)]
        self.sheet_rows = [[("" if v is None else str(v)) for v in row] for row in rows]
        self._refresh_sheet()
        self._sheet_log(f"已导入 {sn}（{len(self.sheet_rows)} 行）")
        self._set_status(f"工作表已载入: {sn}")

    def _on_sheet_save(self):
        """表格 → xlsx（sheet 名 = 模板名）。"""
        try:
            import openpyxl
        except ImportError:
            messagebox.showerror("缺少依赖", "需要 openpyxl\n请运行: pip install openpyxl")
            return
        path = filedialog.asksaveasfilename(title="保存表格为 Excel", defaultextension=".xlsx",
                                            initialfile=f"{self._sheet_name()}.xlsx",
                                            filetypes=[("Excel 模板", "*.xlsx")])
        if not path:
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = (self._sheet_name() or "工作表")[:31]
            for r, row in enumerate(self.sheet_rows, start=1):
                for c, v in enumerate(row, start=1):
                    cell = ws.cell(row=r, column=c, value=v)
                    if cell.data_type != "s":
                        cell.data_type = "s"
            ws.column_dimensions["A"].width = 44
            wb.save(path)
        except Exception as e:
            messagebox.showerror("保存失败", str(e))
            return
        self._set_status(f"表格已保存: {path}")
        try:
            os.startfile(path)
        except Exception:
            pass

    def _on_sheet_generate(self):
        """表格直接渲染 → 输出 cfg（Excel 宏工作流，不依赖 ① ② 页）。"""
        warnings = []
        templates, devices = parse_sheet_rows(self.sheet_rows, self._sheet_name(), warnings)
        if not templates:
            messagebox.showwarning("无法生成",
                                   "表格里没有识别到模板：\n· A1 需为【模板名】或 A 列含 {{变量}}\n"
                                   "· 第 1 行 C 列起要有参数名（{{sysname}}）\n" +
                                   ("\n".join(f"  ⚠ {w}" for w in warnings)))
            return
        if not devices:
            messagebox.showwarning("无法生成", "表格里没有设备（第 2 行起 C 列起填参数，全空行会被跳过）")
            return
        # 同步设备进①页清单（U盘包 / lswnet / DHCP 生某园区读 self.devices；可撤销）
        self._push_undo("工作表生成")
        self.devices = devices
        self._checked = set()
        save_devices_db(self.devices)
        self._update_dev_table()
        name, tpl = next(iter(templates.items()))
        out_dir = self._latest_output_dir()
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("目录错误", str(e))
            return
        ok, fail = 0, 0
        self.sheet_logbox.delete("1.0", tk.END)
        for dev in devices:
            host = dev.get("hostname") or dev.get("esn") or "?"
            try:
                cfg = build_switch_config(tpl, dev)
                with open(os.path.join(out_dir, host + ".cfg"), "w", encoding="utf-8", newline="\n") as f:
                    f.write(cfg)
                ok += 1
                self._sheet_log(f"  ✓ {host}.cfg ({len(cfg.splitlines())} 行)")
            except Exception as e:
                fail += 1
                self._sheet_log(f"  ✗ {host}: {e}")
        if self.gen_archive_var.get():
            self.tftp_root_var.set(out_dir)
        self._last_gen_dir = out_dir
        self._sheet_log(f"模板 [{name}] → {out_dir}")
        self._sheet_log(f"完成: 成功 {ok} / 失败 {fail}")
        self._set_status(f"工作表生成完成: 成功 {ok} / 失败 {fail}")
        try:
            os.startfile(out_dir)
        except Exception:
            pass

    # ── ① 设备清单 / devices tab ──
    def _build_devices_tab(self):
        # 表单行 / input form
        form = ttk.LabelFrame(self.tab_devices, text="设备信息 Device", padding=8)
        form.pack(fill=tk.X, padx=10, pady=(10, 6))

        fields = [
            ("esn", "ESN 序列号", 22), ("hostname", "主机名(自动=cfg文件名)", 24),
            ("type", "类型(自由文本，用于模板映射)", 18), ("mgmt_ip", "管理IP", 14),
            ("mgmt_vlan", "管理VLAN", 8), ("access_vlan", "接入VLAN", 8),
            ("mask", "掩码", 15), ("gateway", "网关(空=自动.254)", 15),
        ]
        self.entry_vars = {}
        for i, (key, label, width) in enumerate(fields):
            r, c = divmod(i, 4)
            ttk.Label(form, text=label).grid(row=r*2, column=c, sticky=tk.W, padx=(8 if c else 0, 2), pady=(6, 0))
            var = tk.StringVar()
            ttk.Entry(form, textvariable=var, width=width).grid(
                row=r*2+1, column=c, sticky=tk.W, padx=(8 if c else 0, 2))
            self.entry_vars[key] = var

        btns = ttk.Frame(form)
        btns.grid(row=6, column=0, columnspan=4, sticky=tk.W, pady=(8, 0))
        ttk.Button(btns, text="添加 Add", style="Accent.TButton", command=self._on_add_device).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btns, text="更新选中 Update", command=self._on_update_device).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btns, text="删除选中 Delete", style="Danger.TButton", command=self._on_delete_device).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(btns, text="导入 CSV", command=self._on_import_csv).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btns, text="导入 Excel", command=self._on_import_excel).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btns, text="导入 Excel 模板", command=self._on_import_excel_template).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btns, text="生成 Excel 模板", command=self._on_export_excel_template).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btns, text="导出 CSV", command=self._on_export_csv).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btns, text="清空", command=self._on_clear_devices).pack(side=tk.LEFT)
        ttk.Button(btns, text="校验 Validate", style="GreenAccent.TButton",
                   command=self._on_validate).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Button(btns, text="清除示例数据", style="Danger.TButton",
                   command=self._on_clear_demo).pack(side=tk.LEFT, padx=(12, 0))

        # 扩展字段 / extra variables: 每行 K:V（回车分隔）→ 模板 {{key}}
        ex_frame = ttk.LabelFrame(self.tab_devices,
            text="扩展字段 Extras（每行一个 K:V，回车分隔；模板用 {{key}} 引用，例如 ntp_server:10.0.0.1）",
            padding=5)
        ex_frame.pack(fill=tk.X, padx=10, pady=(0, 6))
        self.extras_text = scrolledtext.ScrolledText(ex_frame, height=7, font=("Consolas", 9))
        self.extras_text.pack(fill=tk.X)
        ttk.Label(self.tab_devices,
                  text="💡 批量编号: 主机名/管理IP/ESN/扩展字段填 {1-254} 按顺序生成，{1,254} 只生成 1 和 254，"
                       "{01-10} 零填充；多字段联动一一对应（如 SW-{1-3} ↔ 10.10.10.{1-3}）",
                  foreground="#16a34a", font=("Microsoft YaHei UI", 8)).pack(anchor=tk.W, padx=12, pady=(0, 4))

        # 搜索 + 分页工具条（千台场景必备）
        toolbar = ttk.Frame(self.tab_devices)
        toolbar.pack(fill=tk.X, padx=10, pady=(0, 4))
        ttk.Label(toolbar, text="🔍 搜索:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_change)
        ttk.Entry(toolbar, textvariable=self.search_var, width=22).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="☑ 全选", width=7, command=self._check_all).pack(side=tk.LEFT, padx=(4, 2))
        ttk.Button(toolbar, text="反选", width=5, command=self._check_invert).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(toolbar, text="取消", width=5, command=self._check_none).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(toolbar, text="批量删除选中", style="Danger.TButton",
                   command=self._on_batch_delete).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="批量修改选中", command=self._on_batch_edit).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(toolbar, text="导出勾选 CSV", command=self._on_export_checked).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(toolbar, text="↶ 撤销", command=self._on_undo).pack(side=tk.LEFT, padx=(6, 0))
        self.page_info_var = tk.StringVar(value="")
        ttk.Label(toolbar, textvariable=self.page_info_var,
                  foreground="#1e40af", font=("Consolas", 9)).pack(side=tk.LEFT, padx=8)
        ttk.Button(toolbar, text="◀ 上一页", width=9, command=self._page_prev).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(toolbar, text="下一页 ▶", width=9, command=self._page_next).pack(side=tk.RIGHT)
        ttk.Label(toolbar, text="每页:").pack(side=tk.RIGHT, padx=(4, 0))
        self.page_size_var = tk.StringVar(value="200")
        self.page_size_var.trace_add("write", self._on_search_change)
        ttk.Combobox(toolbar, textvariable=self.page_size_var, values=("100", "200", "500", "1000"),
                     width=5, state="readonly").pack(side=tk.RIGHT)

        # 设备表 / device table（首列 ☐ 多选，点击切换；行点击填充表单）
        tbl_frame = ttk.LabelFrame(self.tab_devices, text="设备列表（☐ 勾选多选 / 行点击填充表单）", padding=5)
        tbl_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        cols = ("sel", "esn", "hostname", "type", "mgmt_ip", "mgmt_vlan", "access_vlan",
                "mask", "gateway", "note")
        headers = {"sel": "✓", "esn": "ESN", "hostname": "主机名", "type": "类型", "mgmt_ip": "管理IP",
                   "mgmt_vlan": "管理VLAN", "access_vlan": "接入VLAN", "mask": "掩码",
                   "gateway": "网关", "note": "备注"}
        widths = {"sel": 36, "esn": 130, "hostname": 220, "type": 90, "mgmt_ip": 110, "mgmt_vlan": 70,
                  "access_vlan": 70, "mask": 110, "gateway": 110, "note": 140}
        self.dev_tree = ttk.Treeview(tbl_frame, columns=cols, show="headings", selectmode="browse")
        for c in cols:
            self.dev_tree.heading(c, text=headers[c])
            self.dev_tree.column(c, width=widths[c], anchor=tk.W)
        vsb = ttk.Scrollbar(tbl_frame, orient=tk.VERTICAL, command=self.dev_tree.yview)
        self.dev_tree.configure(yscrollcommand=vsb.set)
        self.dev_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.dev_tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.dev_tree.bind("<Button-1>", self._on_tree_click)

    def _on_add_device(self):
        base = {k: v.get().strip() for k, v in self.entry_vars.items()}
        base["note"] = ""
        base["_extras"] = parse_extras_text(self.extras_text.get("1.0", tk.END))
        if not base["esn"] and not base["hostname"]:
            messagebox.showwarning("提示", "ESN 或主机名至少填一个")
            return
        if not base["mgmt_ip"]:
            messagebox.showwarning("提示", "管理 IP 必填")
            return
        est = count_expansions(base)          # 先轻量预估，防爆炸
        if est > 500:
            messagebox.showerror("范围过大",
                                 f"将生成约 {est} 台设备，超过上限 500，已取消。\n"
                                 "请缩小 {范围}（如 {1-254}）或分批添加")
            return
        devs = expand_device(base)
        if len(devs) > 500:
            messagebox.showerror("范围过大", f"将生成 {len(devs)} 台设备，超过上限 500，已取消。\n"
                                           "请缩小 {范围}（如 {1-254}）或分批添加")
            return
        if len(devs) > 20 and not messagebox.askyesno("批量添加",
                f"检测到批量编号，将生成 {len(devs)} 台设备，继续？"):
            return
        self._push_undo("添加设备")
        self.devices.extend(devs)
        self._clear_form()
        save_devices_db(self.devices)
        self._update_dev_table()
        label = base["hostname"] or base["esn"]
        warn = check_range_lengths(base)
        status = f"已添加 {len(devs)} 台（{label}）" if len(devs) > 1 else f"已添加 {label}"
        if warn:
            status += f" ⚠ {warn}"
        self._set_status(status)

    def _on_update_device(self):
        idx = self._selected_index()
        if idx is None:
            return
        dev = {k: v.get().strip() for k, v in self.entry_vars.items()}
        dev["note"] = self.devices[idx].get("note", "")
        dev["_extras"] = parse_extras_text(self.extras_text.get("1.0", tk.END))
        self._push_undo("更新设备")
        self.devices[idx] = dev
        save_devices_db(self.devices)
        self._update_dev_table(keep_selection=idx)
        self._set_status(f"已更新 {dev['hostname'] or dev['esn']}")

    def _on_delete_device(self):
        idx = self._selected_index()
        if idx is None:
            return
        host = self.devices[idx].get("hostname") or self.devices[idx].get("esn")
        if messagebox.askyesno("确认", f"删除设备 {host}？"):
            self._push_undo("删除设备")
            self.devices.pop(idx)
            self._clear_form()
            self._checked = set()         # 索引已移位，勾选失效
            save_devices_db(self.devices)
            self._update_dev_table()

    def _on_tree_select(self, _ev=None):
        idx = self._selected_index()
        if idx is None:
            return
        dev = self.devices[idx]
        for k, var in self.entry_vars.items():
            var.set(dev.get(k, ""))
        self.extras_text.delete("1.0", tk.END)
        self.extras_text.insert("1.0", format_extras_text(dev.get("_extras")))

    def _selected_index(self):
        sel = self.dev_tree.selection()
        if not sel:
            return None
        iid = sel[0]
        tags = self.dev_tree.item(iid, "tags")
        if tags:
            return int(tags[0])
        return int(iid)

    def _on_validate(self):
        if not self.devices:
            messagebox.showwarning("提示", "设备清单为空")
            return
        problems = validate_devices(self.devices)
        errors = [p for p in problems if p[0] == "error"]
        warns = [p for p in problems if p[0] == "warning"]
        win = tk.Toplevel(self.root)
        win.title(f"设备校验结果 — {len(self.devices)} 台")
        win.geometry("780x480")
        tk.Label(win,
                 text=f"共 {len(self.devices)} 台设备：错误 {len(errors)} 项 / 警告 {len(warns)} 项",
                 font=("Microsoft YaHei UI", 11, "bold"),
                 fg="#dc2626" if errors else "#16a34a").pack(anchor=tk.W, padx=12, pady=(10, 4))
        txt = scrolledtext.ScrolledText(win, font=("Consolas", 9))
        txt.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 10))
        txt.tag_configure("err", foreground="#dc2626")
        txt.tag_configure("warn", foreground="#b45309")
        if not problems:
            txt.insert(tk.END, "✓ 全部通过，没有发现问题")
        else:
            for sev, host, msg in problems:
                tag = "err" if sev == "error" else "warn"
                txt.insert(tk.END, f"[{'错误' if sev == 'error' else '警告'}] {host}: {msg}\n", tag)
        txt.config(state=tk.DISABLED)
        self._set_status(f"校验完成: 错误 {len(errors)} / 警告 {len(warns)}")

    def _on_search_change(self, *_):
        self._page = 0
        self._update_dev_table()

    def _page_prev(self):
        if self._page > 0:
            self._page -= 1
            self._update_dev_table()

    def _page_next(self):
        self._page += 1
        self._update_dev_table()          # 越界自动收敛

    # ── 多选 / 批量操作 ──
    def _push_undo(self, label: str):
        """变更操作前调用：压入撤销快照（最多 20 步）。"""
        self._undo_stack.append((label, copy.deepcopy(self.devices)))
        if len(self._undo_stack) > 20:
            self._undo_stack.pop(0)

    def _on_undo(self):
        """撤销最近一次设备清单变更（误删/误改后悔药）。"""
        if not self._undo_stack:
            messagebox.showinfo("提示", "没有可撤销的操作")
            return
        label, snap = self._undo_stack.pop()
        self.devices = snap
        self._checked = set()
        self._page = 0
        save_devices_db(self.devices)
        self._update_dev_table()
        self._set_status(f"↶ 已撤销: {label}（剩余 {len(self._undo_stack)} 步可撤销）")

    def _check_all(self):
        """全选当前过滤结果（没搜索=全部设备）。"""
        self._checked = set(self._filtered_idx) if self._filtered_idx is not None \
            else set(range(len(self.devices)))
        self._update_dev_table()

    def _check_none(self):
        self._checked = set()
        self._update_dev_table()

    def _check_invert(self):
        total = self._filtered_idx if self._filtered_idx is not None \
            else list(range(len(self.devices)))
        self._checked = set(total) - self._checked
        self._update_dev_table()

    def _on_tree_click(self, ev):
        """点击首列 ☐ 切换勾选（其它列走默认行为）。"""
        if self.dev_tree.identify_region(ev.x, ev.y) != "cell":
            return
        if self.dev_tree.identify_column(ev.x) != "#1":
            return
        row = self.dev_tree.identify_row(ev.y)
        if not row:
            return
        tags = self.dev_tree.item(row, "tags")
        idx = int(tags[0]) if tags else int(row)
        if idx in self._checked:
            self._checked.discard(idx)
        else:
            self._checked.add(idx)
        self._update_dev_table()

    def _on_batch_delete(self):
        if not self._checked:
            messagebox.showinfo("提示", "请先勾选要删除的设备（点击行首的 ☐，或点「全选」）")
            return
        n = len(self._checked)
        hosts = [self.devices[i].get("hostname") or self.devices[i].get("esn") or "?"
                 for i in sorted(self._checked)[:5]]
        preview = "、".join(hosts) + (" …" if n > 5 else "")
        if not messagebox.askyesno("批量删除",
                f"删除勾选的 {n} 台设备？\n\n{preview}"):
            return
        self._push_undo("批量删除")
        self.devices = [d for i, d in enumerate(self.devices) if i not in self._checked]
        self._checked = set()
        save_devices_db(self.devices)
        self._update_dev_table()
        self._set_status(f"已删除 {n} 台设备")

    def _on_batch_edit(self):
        if not self._checked:
            messagebox.showinfo("提示", "请先勾选要修改的设备（点击行首的 ☐，或点「全选」）")
            return
        win = tk.Toplevel(self.root)
        win.title(f"批量修改 {len(self._checked)} 台设备")
        win.geometry("560x420")
        win.transient(self.root)
        tk.Label(win,
                 text=f"勾选 {len(self._checked)} 台：填写的字段会覆盖到全部选中设备（留空 = 不改）",
                 font=("Microsoft YaHei UI", 9)).pack(anchor=tk.W, padx=12, pady=(10, 4))
        fields = [("type", "类型"), ("mgmt_vlan", "管理VLAN"), ("access_vlan", "接入VLAN"),
                  ("mask", "掩码"), ("gateway", "网关"), ("note", "备注"),
                  ("esn", "ESN"), ("mgmt_ip", "管理IP")]
        vars_map = {}
        frm = ttk.Frame(win)
        frm.pack(fill=tk.X, padx=12)
        for i, (key, label) in enumerate(fields):
            r, c = divmod(i, 4)
            ttk.Label(frm, text=label).grid(row=r * 2, column=c, sticky=tk.W, padx=(0, 4))
            var = tk.StringVar()
            ttk.Entry(frm, textvariable=var, width=14).grid(row=r * 2 + 1, column=c,
                                                            padx=(0, 10), pady=(0, 6))
            vars_map[key] = var
        tk.Label(win, text="⚠ 批量改 ESN / 管理IP 请自行确保不重复（生成前有校验拦截）",
                 foreground="#b45309", font=("Microsoft YaHei UI", 8)).pack(anchor=tk.W, padx=12)
        # 扩展字段批量添加/删除
        tk.Label(win, text="扩展字段（每行一个 K:V，应用到选中设备的模板变量；K: 空值 = 删除该字段）",
                 font=("Microsoft YaHei UI", 9)).pack(anchor=tk.W, padx=12, pady=(8, 2))
        self._batch_extras_text = scrolledtext.ScrolledText(win, height=4, font=("Consolas", 9))
        self._batch_extras_text.pack(fill=tk.X, padx=12)

        def apply():
            filled = {k: v.get().strip() for k, v in vars_map.items() if v.get().strip()}
            extras_text = self._batch_extras_text.get("1.0", tk.END)
            extras_map = parse_extras_text(extras_text)
            if not filled and not extras_map:
                messagebox.showwarning("提示", "至少填一个要修改的字段或扩展字段")
                return
            self._push_undo("批量修改")
            for i in sorted(self._checked):
                dev = self.devices[i]
                dev.update(filled)
                ex = dev.setdefault("_extras", {})
                for k, v in extras_map.items():
                    if v == "":           # K: 空值 → 删除该扩展字段
                        ex.pop(k, None)
                    else:
                        ex[k] = v
            save_devices_db(self.devices)
            self._update_dev_table()
            parts = list(filled) + [f"ext.{k}" for k in extras_map]
            self._set_status(f"已批量修改 {len(self._checked)} 台设备: {', '.join(parts)}")
            win.destroy()

        ttk.Button(win, text="应用修改", style="Accent.TButton", command=apply).pack(pady=8)

    def _on_export_checked(self):
        """导出勾选的设备到 CSV（交接/筛选子集用）。"""
        if not self._checked:
            messagebox.showinfo("提示", "请先勾选要导出的设备（点击行首的 ☐，或点「全选」）")
            return
        selected = [self.devices[i] for i in sorted(self._checked)]
        path = filedialog.asksaveasfilename(title="导出勾选设备 CSV", defaultextension=".csv",
                                            initialfile=f"selected_{len(selected)}.csv",
                                            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(devices_to_csv(selected))
        self._set_status(f"已导出勾选 {len(selected)} 台设备: {path}")

    def _update_dev_table(self, keep_selection=None):
        # 搜索过滤（ESN/主机名/IP/类型 子串，无视大小写）
        kw = self.search_var.get().strip().lower()
        if kw:
            self._filtered_idx = [i for i, d in enumerate(self.devices)
                                  if kw in (d.get("esn", "") + " " + d.get("hostname", "") + " " +
                                            d.get("mgmt_ip", "") + " " + d.get("type", "")).lower()]
        else:
            self._filtered_idx = None
        total = len(self._filtered_idx) if self._filtered_idx is not None else len(self.devices)
        # 分页（越界收敛）
        pages = max(1, (total + self._page_size - 1) // self._page_size)
        self._page = max(0, min(self._page, pages - 1))
        start = self._page * self._page_size
        end = min(start + self._page_size, total)
        idxs = (self._filtered_idx[start:end] if self._filtered_idx is not None
                else list(range(start, end)))
        # 渲染（分块插入保持响应）
        self.dev_tree.delete(*self.dev_tree.get_children())
        for n, i in enumerate(idxs):
            d = self.devices[i]
            mark = "☑" if i in self._checked else "☐"
            self.dev_tree.insert("", tk.END, iid=str(i), tags=(str(i),),
                values=(mark, d.get("esn", ""), d.get("hostname", ""), d.get("type", ""),
                        d.get("mgmt_ip", ""), d.get("mgmt_vlan", ""), d.get("access_vlan", ""),
                        d.get("mask", ""), d.get("gateway", ""), d.get("note", "")))
            if n % 200 == 199:
                self.root.update_idletasks()
        self.page_info_var.set(f"共 {total} 台 · 第 {self._page + 1}/{pages} 页 · 显示 {start + 1}-{end}")
        if keep_selection is not None:
            try:
                self.dev_tree.selection_set(str(keep_selection))
            except Exception:
                pass

    def _clear_form(self):
        for var in self.entry_vars.values():
            var.set("")
        self.extras_text.delete("1.0", tk.END)

    def _on_import_csv(self):
        path = filedialog.askopenfilename(title="选择设备 CSV", filetypes=[("CSV", "*.csv"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
                text = f.read()
        except Exception as e:
            messagebox.showerror("导入失败", str(e))
            return
        self._finish_import(text, "CSV", path)

    def _on_import_excel(self):
        try:
            import openpyxl  # noqa
        except ImportError:
            messagebox.showerror("缺少依赖", "Excel 导入需要 openpyxl\n请运行: pip install openpyxl")
            return
        path = filedialog.askopenfilename(title="选择设备 Excel", filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb.active
            rows = [["" if c is None else str(c).strip() for c in row]
                    for row in ws.iter_rows(values_only=True)]
            buf = io.StringIO()
            csv.writer(buf).writerows(rows)
            text = buf.getvalue()
        except Exception as e:
            messagebox.showerror("导入失败", str(e))
            return
        self._finish_import(text, "Excel", path)

    def _finish_import(self, text: str, kind: str, path: str):
        """导入前弹列映射窗口（自动识别表头，可手动调整）→ 解析 → 展开 → 载入。"""
        rows = [r for r in csv.reader(io.StringIO(text)) if r and any(c.strip() for c in r)]
        if not rows:
            messagebox.showwarning("导入失败", f"{kind} 无有效数据（首行应为表头）")
            return
        header = rows[0]
        mapping = {}
        for i, h in enumerate(header):
            mapping[i] = CSV_ALIASES.get(h.strip().lower(), "扩展字段")

        win = tk.Toplevel(self.root)
        win.title(f"{kind} 列映射 — 已自动识别，可调整")
        win.geometry("600x440")
        win.transient(self.root)
        tk.Label(win,
                 text="每列含义可下拉调整：标准字段 / 忽略 / 扩展字段（进模板变量 {{列名}}）",
                 font=("Microsoft YaHei UI", 9)).pack(anchor=tk.W, padx=12, pady=(10, 4))
        frm = ttk.Frame(win)
        frm.pack(fill=tk.BOTH, expand=True, padx=12)
        canvas = tk.Canvas(frm)
        vsb = ttk.Scrollbar(frm, orient=tk.VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_canvas_configure(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _on_canvas_configure)

        vars_map = {}
        for i, h in enumerate(header):
            ttk.Label(inner, text=f"列{i + 1}: {h.strip() or '(空表头)'}").grid(
                row=i, column=0, sticky=tk.W, pady=2, padx=(0, 10))
            var = tk.StringVar(value=mapping.get(i, "扩展字段"))
            ttk.Combobox(inner, textvariable=var, values=MAPPING_OPTIONS,
                         state="readonly", width=16).grid(row=i, column=1, sticky=tk.W, pady=2)
            vars_map[i] = var

        def apply():
            final = {i: var.get() for i, var in vars_map.items()}
            devs = parse_device_mapped(text, final)
            if not devs:
                messagebox.showwarning("导入失败", f"{kind} 无有效数据")
                return
            self._push_undo(f"导入 {kind}")
            expanded = []
            for d in devs:
                expanded.extend(expand_device(d))
            self.devices = expanded
            self._checked = set()
            save_devices_db(self.devices)
            self._update_dev_table()
            self._set_status(f"已导入 {len(expanded)} 台设备: {path}")
            win.destroy()

        ttk.Button(win, text="确认导入", style="Accent.TButton", command=apply).pack(pady=8)

    def _on_export_excel_template(self):
        """生成一个 Excel 模板（sheet1 示例 demo + sheet2 操作方式），填好内容后可用「导入 Excel 模板」导入。"""
        try:
            import openpyxl
        except ImportError:
            messagebox.showerror("缺少依赖", "Excel 生成需要 openpyxl\n请运行: pip install openpyxl")
            return
        path = filedialog.asksaveasfilename(
            title="生成 Excel 模板",
            defaultextension=".xlsx",
            initialfile="NADT设备模板.xlsx",
            filetypes=[("Excel 模板", "*.xlsx")])
        if not path:
            return
        try:
            wb = openpyxl.Workbook()
            # ── sheet1 示例 demo ──
            ws = wb.active
            ws.title = "示例模板"
            ws["A1"] = "【示例模板】"
            ws["B1"] = "文本备注（不参与生成，随便写）"
            params = ["{{sysname}}", "{{vlan}}", "{{ip}}", "{{wg}}", "{{vlanif}}"]
            for i, p in enumerate(params):
                ws.cell(row=1, column=3 + i, value=p)
            demo_tpl = [
                "#",
                "sysname {{sysname}}",
                "#",
                "vlan batch {{vlan}}",
                "#",
                "interface Vlanif{{vlanif}}",
                " ip address {{ip}} 255.255.255.0",
                "#",
                "ip route-static 0.0.0.0 0.0.0.0 {{wg}}",
                "#",
                "return",
                "save",
                "y",
            ]
            for r, line in enumerate(demo_tpl, start=2):
                ws.cell(row=r, column=1, value=line)
            # 设备行（第 2 行起每行一台）
            devices_demo = [
                ("SW-1-1F", "10 20", "192.168.10.11", "192.168.10.254", "10", "1楼接入交换机"),
                ("SW-2-2F", "20 30", "192.168.20.11", "192.168.20.254", "20", "2楼接入交换机"),
                ("SW-3-3F", "30 40", "192.168.30.11", "192.168.30.254", "30", "3楼接入交换机"),
            ]
            for row, (sysname, vlan, ip, wg, vlanif, note) in enumerate(devices_demo, start=2):
                ws.cell(row=row, column=3, value=sysname)
                ws.cell(row=row, column=4, value=vlan)
                ws.cell(row=row, column=5, value=ip)
                ws.cell(row=row, column=6, value=wg)
                ws.cell(row=row, column=7, value=vlanif)
                ws.cell(row=row, column=2, value=note)
            # 格式：标题加粗、列宽
            from openpyxl.styles import Font, Alignment
            ws["A1"].font = Font(bold=True, size=12)
            for c in "ABCDEFGH":
                ws.column_dimensions[c].width = 18
            for r in range(1, 6):
                ws.cell(row=r, column=1).alignment = Alignment(vertical="top")
            # ── sheet2 操作方式 ──
            ws2 = wb.create_sheet("操作方式")
            ws2.column_dimensions["A"].width = 105
            guide = [
                "NADT Excel 模板 · 使用说明",
                "────────────────────────────────────────────────",
                "",
                "一、模板结构（每个 sheet = 一个模板，本文件 sheet1 是示例）",
                "  A1          = 【模板名】（sheet 名也会作为模板名）",
                "  第 1 行      = B 列起为备注说明；C 列起为参数名（必须用 {{参数名}} 格式）",
                "  A 列第 2 行起 = 配置模板文本（{{变量}} 会被替换；'#' 分隔行会保留）",
                "  第 2 行起每行 = 一台交换机（C 列起填该设备各参数值，全空行会被跳过）",
                "",
                "二、操作步骤",
                "  1. 复制「示例模板」sheet，重命名为你的模板名（如：核心交换机）",
                "  2. 修改 A 列配置文本、第 1 行参数名、每行设备参数（可以加任意多列参数）",
                "  3. 在 NADT ① 设备清单页点「导入 Excel 模板」，选择本文件",
                "  4. 确认预览 → 模板进②模板库（类型=sheet名自动映射），设备进①清单",
                "  5. 在③ 生成脚本页点「生成脚本」→ 每台交换机生成一个 .cfg",
                "",
                "三、变量语法（写在 A 列模板文本里）",
                "  {{sysname}}     普通变量替换（参数名列定义的任意名字）",
                "  {{#if 变量}}...{{#else}}...{{#endif}}   条件块（变量非空才输出）",
                "  {{#for n from 1 to 8}}...{{#endfor}}     循环块（n 可用）",
                "  {{#for n from 1 to {{port_count}}}}      循环次数用另一个变量",
                "  # 提醒：{{#if}}/{{#for}} 指令要单独一行，或整行写成",
                "    {{#if x}}A{{#endif}} 这种单行形式；指令行带内容会报错提示",
                "",
                "四、U盘 ZTP 开局（无 DHCP / 无网络的现场，插 U 盘就能开）",
                "  1. 参数列加 {{esn}}（交换机序列号，U盘匹配依据；没有就加 {{mac}} 用 MAC 匹配）",
                "  2. 在③ 生成脚本页点「生成脚本」→ 得到全部 .cfg",
                "  3. ③ 页勾选「☰ 高级选项」→ 点「生成 U盘 ZTP 包」",
                "     → 自动生成 usb_config.ini（每台设备一个 [DEVICEn] 块，按 ESN 指向对应 cfg）",
                "  4. 把 输出目录/usb/ 里的 usb_config.ini + 全部 .cfg 拷到 U 盘根目录",
                "  5. U 盘插交换机 → 开机 → 设备按 ESN 自动匹配加载配置，开局完成",
                "  · 堆叠设备：参数列加 {{stack_member_id}}，每个成员一行（同配置不同 ESN）",
                "  · 顺带刷版本/补丁：加 {{system_software}}（.cc 版本）和 {{system_pat}}（.pat 补丁）",
                "",
                "五、注意事项",
                "  · 参数名要唯一，重复列会被忽略并警告",
                "  · 映射表类 sheet（每行 esn=...;cfgfile=...;）会被自动跳过，NADT 会自己生成映射表",
                "  · 同一台设备不要出现在多个 sheet（同名会互相覆盖），导入后会提示重复",
                "  · 导入后建议点「校验 Validate」检查重复/格式问题",
                "  · 本文件不含宏，就是普通 Excel，哪里都能编辑",
            ]
            for r, line in enumerate(guide, start=1):
                cell = ws2.cell(row=r, column=1, value=line)
                if cell.data_type != "s":
                    cell.data_type = "s"      # 防 "=" 开头被存成公式触发 Excel 告警
                if r == 1:
                    cell.font = Font(bold=True, size=14)
            wb.save(path)
        except Exception as e:
            messagebox.showerror("生成失败", str(e))
            return
        self._set_status(f"Excel 模板已生成: {path}")
        try:
            os.startfile(path)
        except Exception:
            pass

    def _on_import_excel_template(self):
        """导入 Excel 模板（xlsm）：模板 sheet → 模板库 + 设备清单。

        支持师傅的宏脚本工作流：A 列模板文本 + 第 1 行参数名 + 每行一台设备。
        """
        try:
            import openpyxl  # noqa
        except ImportError:
            messagebox.showerror("缺少依赖", "Excel 导入需要 openpyxl\n请运行: pip install openpyxl")
            return
        path = filedialog.askopenfilename(title="选择 Excel 模板（xlsm/xlsx）",
                                          filetypes=[("Excel 模板", "*.xlsm *.xlsx")])
        if not path:
            return
        try:
            templates, devices, warnings = parse_xlsm_templates(path)
        except Exception as e:
            messagebox.showerror("解析失败", str(e))
            return
        if not templates:
            messagebox.showwarning("未识别",
                "没找到模板 sheet。\n要求：sheet 的 A1 为【模板名】，或 A 列含 {{变量}}，\n"
                "第 1 行 C 列起为参数名（如 {{sysname}}），A 列第 2 行起为模板文本。")
            return
        win = tk.Toplevel(self.root)
        win.title("Excel 模板导入预览")
        win.geometry("640x460")
        win.transient(self.root)
        tk.Label(win,
                 text=f"识别到 {len(templates)} 个模板 / {len(devices)} 台设备 — 确认导入？",
                 font=("Microsoft YaHei UI", 10, "bold")).pack(anchor=tk.W, padx=12, pady=(10, 4))
        txt = scrolledtext.ScrolledText(win, font=("Consolas", 9))
        txt.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 4))
        for name, t in templates.items():
            lines = t.count("\n") + 1
            n_dev = sum(1 for d in devices if d["type"] == name)
            txt.insert(tk.END, f"■ 模板 [{name}]  ({lines} 行, {n_dev} 台设备)\n")
            txt.insert(tk.END, f"  预览: {t.splitlines()[0][:70] if t.splitlines() else ''}\n\n")
        if warnings:
            txt.insert(tk.END, "⚠ 警告:\n" + "\n".join(f"  - {w}" for w in warnings) + "\n")
        txt.config(state=tk.DISABLED)
        note = ("导入后：模板进②模板库（类型=sheet名自动映射），设备进①清单（每行一台，"
                "参数进扩展字段），直接点③生成脚本即可")
        tk.Label(win, text=note, foreground="#16a34a",
                 font=("Microsoft YaHei UI", 8)).pack(anchor=tk.W, padx=12)

        def do_import():
            self._push_undo("导入 Excel 模板")
            os.makedirs(TEMPLATE_DIR, exist_ok=True)
            mapping = load_template_mapping()
            for name, text in templates.items():
                fname = re.sub(r'[\\/:*?"<>|]+', "_", name) + ".cfg"
                with open(os.path.join(TEMPLATE_DIR, fname), "w", encoding="utf-8", newline="\n") as f:
                    f.write(text)
                mapping[name] = fname
            save_template_mapping(mapping)
            self._refresh_mapping()
            expanded = []
            for d in devices:
                expanded.extend(expand_device(d))
            self.devices = expanded
            self._checked = set()
            save_devices_db(self.devices)
            self._update_dev_table()
            self._set_status(f"已导入 {len(templates)} 个模板 + {len(expanded)} 台设备: {path}")
            # 跨 sheet 同名设备提示（xlsm 里 JR 与命名规范等常含同名设备）
            dup = [p for p in validate_devices(self.devices)
                   if p[0] == "error" and ("主机名重复" in p[2] or "ESN 重复" in p[2])]
            if dup:
                names = sorted({p[1] for p in dup})
                messagebox.showwarning(
                    "发现重复",
                    f"导入的设备中有 {len(dup)} 台与其它 sheet 同名/同 ESN：\n"
                    f"{', '.join(names[:8])}{' …' if len(names) > 8 else ''}\n\n"
                    "同名设备生成配置时会互相覆盖。建议在①页用「校验 Validate」查看，"
                    "或用「导出勾选 CSV」筛选后处理。")
            win.destroy()

        ttk.Button(win, text="确认导入", style="Accent.TButton", command=do_import).pack(pady=8)

    def _on_clear_demo(self):
        """清除试用/示例数据：设备清单 + 导入的模板（保留默认 access-switch.cfg）。"""
        extra_tpls = [t for t in list_templates() if t.lower() != "access-switch.cfg"]
        n_dev = len(self.devices)
        if n_dev == 0 and not extra_tpls:
            messagebox.showinfo("提示", "当前没有示例数据\n（设备清单为空，模板库只有默认模板）")
            return
        if not messagebox.askyesno(
                "清除示例数据",
                f"将清除：\n· 设备清单 {n_dev} 台\n· 导入的模板 {len(extra_tpls)} 个\n\n"
                "（保留默认模板 access-switch.cfg）\n此操作不可撤销，确认清除？"):
            return
        # 清设备清单（保存空库）
        self._push_undo("清除示例数据")
        self.devices = []
        self._checked = set()
        self._page = 0
        save_devices_db(self.devices)
        # 删导入的模板文件 + 清映射
        for t in extra_tpls:
            try:
                os.remove(get_template_path(t))
            except OSError:
                pass
        mapping = load_template_mapping()
        changed = False
        for t in extra_tpls:                       # t 是文件名（示例模板.cfg）
            for k in [k for k, v in mapping.items() if v == t]:   # 按 value=文件名 删
                del mapping[k]
                changed = True
        if changed:
            save_template_mapping(mapping)
        self._refresh_mapping()
        self._refresh_tpl_combo()
        # 当前模板被删则回退默认
        if self.current_template.lower() != "access-switch.cfg" and \
                not os.path.isfile(get_template_path(self.current_template)):
            self.current_template = "access-switch.cfg"
            self._load_template(self.current_template)
        self._update_dev_table()
        self._clear_form()
        self._set_status(f"已清除示例数据: {n_dev} 台设备 + {len(extra_tpls)} 个模板")

    def _on_export_csv(self):
        if not self.devices:
            messagebox.showwarning("提示", "设备清单为空")
            return
        path = filedialog.asksaveasfilename(title="导出设备 CSV", defaultextension=".csv",
                                            initialfile="devices.csv", filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(devices_to_csv(self.devices))
        self._set_status(f"已导出 {len(self.devices)} 台设备: {path}")

    def _on_clear_devices(self):
        if self.devices and messagebox.askyesno("确认", "清空全部设备？"):
            self._push_undo("清空设备")
            self.devices = []
            self._clear_form()
            self._checked = set()
            save_devices_db(self.devices)
            self._update_dev_table()
    # ── ② 模板 / template tab ──
    def _build_template_tab(self):
        bar = ttk.LabelFrame(self.tab_template, text="模板库 Template Library", padding=6)
        bar.pack(fill=tk.X, padx=10, pady=(10, 5))
        ttk.Label(bar, text="当前模板:").pack(side=tk.LEFT)
        self.tpl_combo = ttk.Combobox(bar, state="readonly", width=32)
        self.tpl_combo.pack(side=tk.LEFT, padx=6)
        self.tpl_combo.bind("<<ComboboxSelected>>", lambda e: self._on_tpl_load())
        ttk.Button(bar, text="载入", command=self._on_tpl_load).pack(side=tk.LEFT)
        ttk.Button(bar, text="保存", style="Accent.TButton", command=self._on_save_template).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(bar, text="另存为新模板", command=self._on_save_template_as).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(bar, text="删除模板", style="Danger.TButton", command=self._on_delete_template).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(bar, text="重置默认", command=self._on_reset_template).pack(side=tk.LEFT, padx=(6, 0))

        # 类型→模板 映射 / type→template mapping
        map_frame = ttk.LabelFrame(self.tab_template, text="类型→模板映射（生成时按设备类型选模板；空类型用当前模板）", padding=6)
        map_frame.pack(fill=tk.X, padx=10, pady=(0, 5))
        self.map_tree = ttk.Treeview(map_frame, columns=("type", "template"), show="headings", height=3)
        self.map_tree.heading("type", text="设备类型")
        self.map_tree.heading("template", text="模板文件")
        self.map_tree.column("type", width=180)
        self.map_tree.column("template", width=320)
        self.map_tree.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        self.map_type_var = tk.StringVar()
        self.map_tpl_var = tk.StringVar()
        ttk.Entry(map_frame, textvariable=self.map_type_var, width=16).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Entry(map_frame, textvariable=self.map_tpl_var, width=24).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(map_frame, text="添加/更新", command=self._on_map_add).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(map_frame, text="删除", command=self._on_map_del).pack(side=tk.LEFT)
        self.map_tree.bind("<<TreeviewSelect>>", self._on_map_select)

        self.placeholder_var = tk.StringVar(value="")
        ttk.Label(self.tab_template, textvariable=self.placeholder_var,
                  foreground="#64748b").pack(anchor=tk.W, padx=12)

        self.template_text = scrolledtext.ScrolledText(self.tab_template, font=("Consolas", 9),
                                                       undo=True, wrap=tk.NONE)
        self.template_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(6, 10))
        self._refresh_tpl_combo()
        self._refresh_mapping()

    def _refresh_tpl_combo(self):
        names = list_templates()
        self.tpl_combo["values"] = names
        if self.current_template in names:
            self.tpl_combo.set(self.current_template)
        elif names:
            self.tpl_combo.set(names[0])

    def _refresh_mapping(self):
        self.template_mapping = load_template_mapping()
        self.map_tree.delete(*self.map_tree.get_children())
        for t, f in sorted(self.template_mapping.items()):
            self.map_tree.insert("", tk.END, values=(t, f))

    def _load_template(self, name: str = None):
        """加载模板文件到编辑器；首次运行从打包内嵌复制默认模板。"""
        name = name or self.current_template
        path = get_template_path(name)
        try:
            bundled = ""
            meipass = getattr(sys, "_MEIPASS", "")
            if meipass:
                bundled = os.path.join(meipass, "templates", "access-switch.cfg")
            if not os.path.isfile(path) and name == "access-switch.cfg":
                os.makedirs(TEMPLATE_DIR, exist_ok=True)
                if bundled and os.path.isfile(bundled):
                    with open(bundled, "r", encoding="utf-8") as f:
                        data = f.read()
                    with open(path, "w", encoding="utf-8", newline="\n") as f:
                        f.write(data)
            with open(path, "r", encoding="utf-8") as f:
                self.template_text.delete("1.0", tk.END)
                self.template_text.insert("1.0", f.read())
            self.current_template = name
            self._refresh_tpl_combo()
            self._scan_placeholders()
        except Exception as e:
            self._set_status(f"模板加载失败: {e}")

    def _scan_placeholders(self):
        text = self.template_text.get("1.0", tk.END)
        vars_ = sorted(set(re.findall(r"\{\{(\w+)\}\}", text)))
        loop_vars = set(re.findall(r"\{\{#for\s+(\w+)", text))
        vars_ = [v for v in vars_ if v not in ("if", "else", "endif", "for", "endfor")
                 and v not in loop_vars]
        extra_note = ""
        if vars_:
            missing = [v for v in vars_ if v not in CSV_HEADER
                       and v not in ("hostname", "vlan_batch", "mgmt_vlan", "mgmt_ip",
                                     "mask", "access_vlan", "gateway")]
            if missing:
                extra_note = "  （需在设备扩展字段或 CSV 额外列提供: " + ", ".join(missing) + "）"
        self.placeholder_var.set("模板占位符: " + (" ".join("{{" + v + "}}" for v in vars_) if vars_ else "（无）") + extra_note)

    def _on_tpl_load(self):
        name = self.tpl_combo.get().strip()
        if name:
            self._load_template(name)

    def _on_save_template(self):
        try:
            os.makedirs(TEMPLATE_DIR, exist_ok=True)
            with open(get_template_path(self.current_template), "w", encoding="utf-8", newline="\n") as f:
                f.write(self.template_text.get("1.0", tk.END))
            self._scan_placeholders()
            self._set_status(f"模板已保存: {self.current_template}")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    def _on_save_template_as(self):
        """另存为新模板：固定写入模板库目录（避免路径混乱）。"""
        import tkinter.simpledialog as sd
        name = sd.askstring("另存为新模板", "模板文件名（自动加 .cfg）:", initialvalue="new-template")
        if not name:
            return
        if not name.endswith(".cfg"):
            name += ".cfg"
        name = os.path.basename(name)
        try:
            with open(get_template_path(name), "w", encoding="utf-8", newline="\n") as f:
                f.write(self.template_text.get("1.0", tk.END))
            self.current_template = name
            self._refresh_tpl_combo()
            self._set_status(f"模板已另存: {name}")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    def _on_delete_template(self):
        name = self.current_template
        if name == "access-switch.cfg":
            messagebox.showwarning("提示", "默认模板 access-switch.cfg 不可删除")
            return
        if messagebox.askyesno("确认", f"删除模板 {name}？"):
            try:
                os.remove(get_template_path(name))
                self._load_template("access-switch.cfg")
                self._set_status(f"模板已删除: {name}")
            except Exception as e:
                messagebox.showerror("删除失败", str(e))

    def _on_reset_template(self):
        if messagebox.askyesno("确认", "载入默认 access-switch.cfg？（当前未保存修改会丢失）"):
            self._load_template("access-switch.cfg")

    def _on_map_add(self):
        t = self.map_type_var.get().strip().upper()
        f = self.map_tpl_var.get().strip()
        if not t or not f:
            messagebox.showwarning("提示", "类型和模板文件都要填")
            return
        if not os.path.isfile(get_template_path(f)):
            messagebox.showwarning("提示", f"模板文件不存在: {f}")
            return
        self.template_mapping[t] = f
        save_template_mapping(self.template_mapping)
        self._refresh_mapping()
        self._set_status(f"映射已保存: {t} → {f}")

    def _on_map_del(self):
        sel = self.map_tree.selection()
        if not sel:
            return
        t = self.map_tree.item(sel[0], "values")[0]
        if messagebox.askyesno("确认", f"删除类型映射 {t}？"):
            self.template_mapping.pop(t, None)
            save_template_mapping(self.template_mapping)
            self._refresh_mapping()

    def _on_map_select(self, _ev=None):
        sel = self.map_tree.selection()
        if not sel:
            return
        t, f = self.map_tree.item(sel[0], "values")
        self.map_type_var.set(t)
        self.map_tpl_var.set(f)

    # ── ③ 生成 / generate tab ──
    def _build_generate_tab(self):
        bar = ttk.LabelFrame(self.tab_generate, text="生成设置", padding=8)
        bar.pack(fill=tk.X, padx=10, pady=(10, 6))
        ttk.Label(bar, text="输出目录:").pack(side=tk.LEFT)
        self.output_var = tk.StringVar(value=self.output_dir)
        ttk.Entry(bar, textvariable=self.output_var, width=52).pack(side=tk.LEFT, padx=6)
        ttk.Button(bar, text="浏览", command=self._on_browse_output).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(bar, text="生成脚本", style="Accent.TButton", command=self._on_generate).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(bar, text="打开目录", command=self._on_open_output).pack(side=tk.LEFT)
        self.gen_archive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bar, text="自动归档", variable=self.gen_archive_var,
                        command=self._on_archive_toggle).pack(side=tk.LEFT, padx=(8, 0))
        self.gen_archive_hint = tk.StringVar(value="每次生成存到 时间戳子目录，保留最近 10 次")
        ttk.Label(bar, textvariable=self.gen_archive_hint,
                  foreground="#64748b").pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(bar, text="  厂商预设:").pack(side=tk.LEFT, padx=(12, 2))
        self.gen_vendor_var = tk.StringVar(value=VENDOR_PRESETS["generic"]["name"])
        self.gen_vendor_var.trace_add("write", self._on_vendor_change)
        vendor_names = [VENDOR_PRESETS[k]["name"] for k in VENDOR_PRESETS]
        ttk.Combobox(bar, textvariable=self.gen_vendor_var, state="readonly", width=34,
                     values=vendor_names).pack(side=tk.LEFT)
        ttk.Label(bar, text="  引导文件名:").pack(side=tk.LEFT, padx=(8, 2))
        self.gen_bootfile_var = tk.StringVar(value="")
        ttk.Entry(bar, textvariable=self.gen_bootfile_var, width=16).pack(side=tk.LEFT, padx=4)
        self.gen_mapping_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(bar, text="生成 ESN 映射表", variable=self.gen_mapping_var).pack(side=tk.LEFT)

        ttk.Label(self.tab_generate,
                  text="生成内容: 每台设备一个 .cfg（文件名=主机名.cfg）+ lswnet.cfg（ESN 映射表，DHCP option 66/146 引导）",
                  foreground="#64748b").pack(anchor=tk.W, padx=12)
        # 高级选项折叠（U盘 ZTP / 版本 / 补丁 / 堆叠 —— 复杂场景才用）
        self.ztp_adv_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.tab_generate, text="☰ 高级选项（U盘 ZTP 包 / 版本补丁 / 堆叠）",
                        variable=self.ztp_adv_var, command=self._on_ztp_adv_toggle,
                        style="TCheckbutton").pack(anchor=tk.W, padx=12, pady=(2, 2))
        # U盘 ZTP（EasyDeploy USB 开局）+ 版本文件（复杂场景）
        self.ztp_frame = ttk.LabelFrame(self.tab_generate, text="U盘 ZTP 包（EasyDeploy USB 开局，无需 DHCP）", padding=6)
        ztp = self.ztp_frame
        ttk.Label(ztp, text="FILESERVER:").pack(side=tk.LEFT)
        self.ztp_server_var = tk.StringVar(value="file:/usb:")
        ttk.Entry(ztp, textvariable=self.ztp_server_var, width=24).pack(side=tk.LEFT, padx=4)
        ttk.Label(ztp, text="SYSTEM-SOFTWARE(.cc):").pack(side=tk.LEFT, padx=(8, 2))
        self.ztp_sw_var = tk.StringVar(value="")
        ttk.Entry(ztp, textvariable=self.ztp_sw_var, width=24).pack(side=tk.LEFT, padx=4)
        ttk.Label(ztp, text="SYSTEM-PAT(.pat):").pack(side=tk.LEFT, padx=(8, 2))
        self.ztp_pat_var = tk.StringVar(value="")
        ttk.Entry(ztp, textvariable=self.ztp_pat_var, width=20).pack(side=tk.LEFT, padx=4)
        ttk.Button(ztp, text="生成 U盘 ZTP 包", style="Accent.TButton",
                   command=self._on_gen_usb).pack(side=tk.LEFT, padx=(10, 0))
        ztp2 = ttk.Frame(self.tab_generate)
        ztp2.pack(fill=tk.X, padx=12, pady=(0, 4))
        ttk.Label(ztp2, text="lswnet 版本文件 vrpfile(.cc):").pack(side=tk.LEFT)
        self.ztp_vrp_var = tk.StringVar(value="")
        ttk.Entry(ztp2, textvariable=self.ztp_vrp_var, width=26).pack(side=tk.LEFT, padx=4)
        ttk.Label(ztp2, text="补丁 patchfile(.pat):").pack(side=tk.LEFT, padx=(8, 2))
        self.ztp_patch_var = tk.StringVar(value="")
        ttk.Entry(ztp2, textvariable=self.ztp_patch_var, width=24).pack(side=tk.LEFT, padx=4)
        ttk.Label(ztp2,
                  text="留空则不含该字段；设备扩展字段 vrp_file/patch_file/system_software/system_pat/stack_member_id 可单台覆盖",
                  foreground="#64748b").pack(side=tk.LEFT, padx=(8, 0))
        prog = ttk.Frame(self.tab_generate)
        prog.pack(fill=tk.X, padx=10, pady=(4, 2))
        ttk.Label(prog, text="进度:").pack(side=tk.LEFT)
        self.gen_progress = ttk.Progressbar(prog, maximum=100)
        self.gen_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        self.gen_progress_label = ttk.Label(prog, text="", width=10, anchor=tk.E)
        self.gen_progress_label.pack(side=tk.LEFT)
        self.gen_log = scrolledtext.ScrolledText(self.tab_generate, height=10, font=("Consolas", 9),
                                                 state=tk.DISABLED, bg="#fbfdf9")
        self.gen_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(6, 10))

    def _gen_log(self, msg: str, tag=None):
        self.gen_log.config(state=tk.NORMAL)
        self.gen_log.insert(tk.END, msg + "\n", tag)
        self.gen_log.see(tk.END)
        self.gen_log.config(state=tk.DISABLED)

    def _on_ztp_adv_toggle(self):
        """高级选项折叠：勾选展开 U盘 ZTP 面板，取消收起。"""
        if self.ztp_adv_var.get():
            self.ztp_frame.pack(fill=tk.X, padx=10, pady=(0, 4))
        else:
            self.ztp_frame.pack_forget()

    def _on_browse_output(self):
        path = filedialog.askdirectory(title="选择输出目录", initialdir=self.output_dir)
        if path:
            self.output_var.set(path)

    def _on_open_output(self):
        path = self.output_var.get().strip()
        if not os.path.isdir(path):
            os.makedirs(path, exist_ok=True)
        os.startfile(path)

    def _on_generate(self):
        if not self.devices:
            messagebox.showwarning("提示", "设备清单为空，请先在 ① 添加/导入设备")
            return
        # 批量校验：错误需确认，警告仅提示
        problems = validate_devices(self.devices)
        errors = [p for p in problems if p[0] == "error"]
        warns = [p for p in problems if p[0] == "warning"]
        if errors:
            first = "；".join(f"{p[1]}: {p[2]}" for p in errors[:3])
            if not messagebox.askyesno("发现错误",
                    f"校验发现 {len(errors)} 个错误（重复 ESN/IP/主机名、IP 格式等）。\n"
                    f"仍要生成吗？\n\n示例:\n{first}\n\n"
                    f"（建议先在 ① 点「校验 Validate」查看全部问题）"):
                return
        elif warns:
            messagebox.showinfo("提示",
                f"{len(warns)} 台设备缺 ESN，无法写入映射表（设备不能自动匹配配置）。\n继续生成。")
        out_dir = self._latest_output_dir()
        try:
            os.makedirs(out_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("目录错误", str(e))
            return
        # 自动保存当前模板（编辑器内容可能未落盘）
        try:
            os.makedirs(TEMPLATE_DIR, exist_ok=True)
            with open(get_template_path(self.current_template), "w", encoding="utf-8", newline="\n") as f:
                f.write(self.template_text.get("1.0", tk.END))
        except Exception:
            pass
        # 预加载所有用到的模板（避免并行中重复读盘）
        template_cache = {}
        for tname in set(find_template_for_type(d.get("type"), self.template_mapping,
                                                self.current_template) for d in self.devices):
            try:
                with open(get_template_path(tname), "r", encoding="utf-8") as f:
                    template_cache[tname] = f.read()
            except OSError as e:
                messagebox.showerror("模板错误", f"模板 {tname} 读取失败: {e}")
                return
        self.gen_log.delete("1.0", tk.END)
        self._gen_log(f"[{datetime.now().strftime('%H:%M:%S')}] 开始生成 {len(self.devices)} 台设备 → {out_dir}")
        self.gen_progress.config(value=0, maximum=len(self.devices))
        self.gen_progress_label.config(text=f"0/{len(self.devices)}")
        self.root.update_idletasks()

        # 并行渲染（build_switch_config 是纯函数，线程安全）
        import concurrent.futures

        def _build(i):
            dev = self.devices[i]
            try:
                tname = find_template_for_type(dev.get("type"), self.template_mapping,
                                               self.current_template)
                d = derive_defaults(dev)
                if not d["hostname"]:
                    raise ValueError("主机名为空")
                cfg = build_switch_config(template_cache[tname], dev)
                return i, d["hostname"], tname, cfg, None
            except Exception as e:
                return i, dev.get("hostname") or dev.get("esn") or "?", "", "", str(e)

        ok, fail = 0, 0
        failures = []
        results = {}
        workers = min(8, os.cpu_count() or 4)
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
            futs = [pool.submit(_build, i) for i in range(len(self.devices))]
            done = 0
            for fut in concurrent.futures.as_completed(futs):
                i, host, tname, cfg, err = fut.result()
                done += 1
                if err:
                    fail += 1
                    failures.append((host, err))
                else:
                    results[i] = (host, tname, cfg)
                    ok += 1
                if done % 50 == 0 or done == len(self.devices):
                    self.gen_progress.config(value=done)
                    self.gen_progress_label.config(text=f"{done}/{len(self.devices)}")
                    self.root.update_idletasks()

        # 写文件（保持设备原顺序）
        for i in range(len(self.devices)):
            if i not in results:
                continue
            host, tname, cfg = results[i]
            fname = host + ".cfg"
            try:
                with open(os.path.join(out_dir, fname), "w", encoding="utf-8", newline="\n") as f:
                    f.write(cfg)
                self._gen_log(f"  ✓ {fname}  [{tname}] ({len(cfg.splitlines())} 行)")
            except OSError as e:
                fail += 1
                failures.append((fname, str(e)))
        for host, err in failures:
            self._gen_log(f"  ✗ {host}: {err}")
        if failures:
            try:
                with open(os.path.join(out_dir, "failures.csv"), "w", encoding="utf-8-sig",
                          newline="") as f:
                    f.write("设备,错误\n")
                    for a, b in failures:
                        f.write(f"{a},{b.replace(chr(10), ' ')}\n")
                self._gen_log(f"  ⚠ 失败明细已导出: failures.csv ({len(failures)} 项)")
            except Exception:
                pass
        # ESN 映射表 / 独立配置文件（由厂商预设 + 可覆盖参数决定）
        bootfile = self.gen_bootfile_var.get().strip()
        if self.gen_mapping_var.get() and bootfile:
            try:
                lswnet = build_lswnet(self.devices,
                                      self.ztp_vrp_var.get().strip(),
                                      self.ztp_patch_var.get().strip())
                with open(os.path.join(out_dir, bootfile), "w", encoding="utf-8", newline="\n") as f:
                    f.write(lswnet)
                mapped = len([l for l in lswnet.splitlines() if l])
                self._gen_log(f"  ✓ {bootfile}  ({mapped} 条映射)")
                no_esn = [d.get("hostname") or d.get("esn") or "?"
                          for d in self.devices if not (d.get("esn") or "").strip()]
                if no_esn:
                    self._gen_log(f"  ⚠ {len(no_esn)} 台设备缺 ESN，未写入映射表（设备无法自动匹配配置）: "
                                  f"{', '.join(no_esn[:5])}{' …' if len(no_esn) > 5 else ''}")
            except Exception as e:
                self._gen_log(f"  ✗ {bootfile}: {e}")
                fail += 1
        else:
            self._gen_log("  ℹ 独立配置文件模式（通用/思科）: 不生成映射表")
            self._gen_log("    在 ⑤ 选「ISC dhcpd / dnsmasq」生成 DHCP 配置，按 MAC 绑定下发专属 cfg")
        self.gen_progress.config(value=len(self.devices))
        self.gen_progress_label.config(text=f"{len(self.devices)}/{len(self.devices)}")
        self._gen_log(f"完成: 成功 {ok} / 失败 {fail}")
        # 归档模式 → TFTP 根目录联动到本次实际输出目录（否则交换机下载不到新 cfg）
        if self.gen_archive_var.get():
            self.tftp_root_var.set(out_dir)
            if self.tftp.running:
                self._gen_log("⚠ TFTP 正在运行，根目录仍指向旧目录；已把「TFTP 根目录」更新为新归档目录，"
                              "请停止并重新启动 TFTP 生效")
            else:
                self._gen_log(f"已同步 TFTP 根目录 → {out_dir}（启动 Server 即用）")
        self._last_gen_dir = out_dir
        self._set_status(f"生成完成: 成功 {ok} / 失败 {fail}")

    # ── ④ TFTP / tftp tab ──
    def _build_tftp_tab(self):
        bar = ttk.LabelFrame(self.tab_tftp, text="TFTP 服务器（asyncio 异步引擎，千并发）", padding=8)
        bar.pack(fill=tk.X, padx=10, pady=(10, 6))
        ttk.Label(bar, text="根目录:").pack(side=tk.LEFT)
        self.tftp_root_var = tk.StringVar(value=self.output_dir)
        ttk.Entry(bar, textvariable=self.tftp_root_var, width=46).pack(side=tk.LEFT, padx=6)
        ttk.Button(bar, text="浏览", command=self._on_browse_tftp_root).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(bar, text="并发上限:").pack(side=tk.LEFT)
        self.tftp_conc_var = tk.StringVar(value=str(self.tftp.max_concurrent))
        ttk.Entry(bar, textvariable=self.tftp_conc_var, width=5).pack(side=tk.LEFT, padx=4)
        self.tftp_start_btn = ttk.Button(bar, text="启动 Server", style="GreenAccent.TButton",
                                         command=self._on_tftp_start)
        self.tftp_start_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.tftp_stop_btn = ttk.Button(bar, text="停止", style="Danger.TButton",
                                        command=self._on_tftp_stop, state=tk.DISABLED)
        self.tftp_stop_btn.pack(side=tk.LEFT)
        ttk.Button(bar, text="生成开局清单", command=self._on_gen_deployed).pack(side=tk.LEFT, padx=(12, 4))
        ttk.Button(bar, text="重置记录", command=self._on_reset_deployed).pack(side=tk.LEFT)

        info = ttk.LabelFrame(self.tab_tftp, text="本机 IP（填到 DHCP option 66）", padding=8)
        info.pack(fill=tk.X, padx=10, pady=(0, 6))
        self.ip_label = ttk.Label(info, text="", foreground="#1e40af", font=("Consolas", 10, "bold"))
        self.ip_label.pack(anchor=tk.W)
        self.tftp_deployed_label = ttk.Label(info, text="已开局 0 台", foreground="#16a34a")
        self.tftp_deployed_label.pack(anchor=tk.W, pady=(4, 0))

        self.tftp_log = scrolledtext.ScrolledText(self.tab_tftp, height=12, font=("Consolas", 9),
                                                  state=tk.DISABLED, bg="#fbfdf9")
        self.tftp_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        self.root.after(2000, self._update_deployed_label)

    def _update_ip_display(self):
        ips = get_local_ips()
        self.ip_label.config(text="   " + "   ".join(ips) if ips else "   未检测到 IPv4 地址")
        # 部署页同步
        if hasattr(self, "tftp_ip_var"):
            if ips:
                self.tftp_ip_var.set(ips[0])

    def _on_browse_tftp_root(self):
        path = filedialog.askdirectory(title="选择 TFTP 根目录", initialdir=self.tftp_root_var.get())
        if path:
            self.tftp_root_var.set(path)

    def _on_gen_deployed(self):
        """生成开局完成清单（deployed_report.txt）到 TFTP 根目录。"""
        with self.tftp._deployed_lock:
            deployed = list(self.tftp.deployed)
        if not deployed:
            messagebox.showinfo("提示", "暂无已开局记录\n（TFTP 成功下发过配置文件才会记录）")
            return
        lines = [
            "# NADT 开局完成清单",
            f"# 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"# 已开局: {len(deployed)} 台",
            "",
        ]
        for ts, fname, ip in deployed:
            host = fname[:-4] if fname.lower().endswith(".cfg") else fname
            dev = next((d for d in self.devices if d.get("hostname") == host), None)
            extra = ""
            if dev:
                extra = f"  ESN={dev.get('esn') or '-'}  类型={dev.get('type') or '-'}"
            else:
                extra = "  （清单中未找到此设备）"
            lines.append(f"[{ts}] {host}  客户端 {ip}{extra}")
        path = os.path.join(self.tftp_root_var.get().strip() or self.tftp.root_dir,
                            "deployed_report.txt")
        try:
            with open(path, "w", encoding="utf-8-sig", newline="\n") as f:
                f.write("\n".join(lines) + "\n")
        except Exception as e:
            messagebox.showerror("写入失败", str(e))
            return
        self._set_status(f"开局清单已生成: {path}")
        try:
            os.startfile(path)
        except Exception:
            pass

    def _on_reset_deployed(self):
        """新一批开局前清空记录。"""
        with self.tftp._deployed_lock:
            n = len(self.tftp.deployed)
        if n and messagebox.askyesno("确认", f"清空 {n} 条已开局记录？（新一批开局前使用）"):
            with self.tftp._deployed_lock:
                self.tftp.deployed.clear()
            self._set_status("已开局记录已清空")
        self._update_deployed_label()

    def _update_deployed_label(self):
        try:
            self.tftp_deployed_label.config(
                text=f"已开局 {len(self.tftp.deployed)} 台（成功下发配置文件数，含重复下载）")
        except Exception:
            pass
        try:
            self.root.after(2000, self._update_deployed_label)
        except Exception:
            pass

    def _on_tftp_start(self):
        root = self.tftp_root_var.get().strip()
        try:
            self.tftp.max_concurrent = max(1, int(self.tftp_conc_var.get().strip()))
        except ValueError:
            self.tftp.max_concurrent = 500
            self.tftp_conc_var.set("500")
        self.tftp.root_dir = root
        self.tftp.start()
        if self.tftp.running:
            self.tftp_start_btn.config(state=tk.DISABLED)
            self.tftp_stop_btn.config(state=tk.NORMAL)
            self._set_status(f"TFTP 运行中: 0.0.0.0:69 root={root}")

    def _on_tftp_stop(self):
        self.tftp.stop()
        self.tftp_start_btn.config(state=tk.NORMAL)
        self.tftp_stop_btn.config(state=tk.DISABLED)
        self._set_status("TFTP 已停止")

    def _tftp_log(self, msg: str):
        """TFTP 日志（TFTP 线程调用 → 主线程刷新，tkinter 线程安全）。"""
        def _append():
            try:
                self.tftp_log.config(state=tk.NORMAL)
                self.tftp_log.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
                self.tftp_log.see(tk.END)
                self.tftp_log.config(state=tk.DISABLED)
            except Exception:
                pass
        try:
            self.root.after(0, _append)
        except Exception:
            pass

    # ── ⑤ 部署说明 / guide tab ──
    def _build_guide_tab(self):
        bar = ttk.LabelFrame(self.tab_guide, text="DHCP option 66/67 配置生成", padding=8)
        bar.pack(fill=tk.X, padx=10, pady=(10, 6))
        ttk.Label(bar, text="TFTP 服务器 IP:").pack(side=tk.LEFT)
        self.tftp_ip_var = tk.StringVar(value="")
        ttk.Entry(bar, textvariable=self.tftp_ip_var, width=16).pack(side=tk.LEFT, padx=4)
        ttk.Label(bar, text="管理VLAN:").pack(side=tk.LEFT, padx=(10, 2))
        self.guide_vlan_var = tk.StringVar(value="11")
        ttk.Entry(bar, textvariable=self.guide_vlan_var, width=6).pack(side=tk.LEFT, padx=4)
        ttk.Label(bar, text="网段:").pack(side=tk.LEFT, padx=(10, 2))
        self.guide_net_var = tk.StringVar(value="192.168.100.0")
        ttk.Entry(bar, textvariable=self.guide_net_var, width=14).pack(side=tk.LEFT, padx=4)
        ttk.Label(bar, text="网关:").pack(side=tk.LEFT, padx=(10, 2))
        self.guide_gw_var = tk.StringVar(value="192.168.100.254")
        ttk.Entry(bar, textvariable=self.guide_gw_var, width=14).pack(side=tk.LEFT, padx=4)
        ttk.Button(bar, text="生成配置", style="Accent.TButton", command=self._on_gen_dhcp).pack(side=tk.LEFT, padx=(10, 0))

        # DHCP 输出风格（厂商无关核心，多种 DHCP 服务器）
        style_bar = ttk.Frame(self.tab_guide)
        style_bar.pack(fill=tk.X, padx=10, pady=(0, 6))
        ttk.Label(style_bar, text="DHCP 输出风格:").pack(side=tk.LEFT)
        self.guide_style_var = tk.StringVar(value=DHCP_STYLES["isc"])
        ttk.Combobox(style_bar, textvariable=self.guide_style_var, state="readonly", width=36,
                     values=list(DHCP_STYLES.values())).pack(side=tk.LEFT, padx=4)
        ttk.Label(style_bar, text="  URL(vrp146):").pack(side=tk.LEFT, padx=(8, 2))
        self.guide_url_var = tk.StringVar(value="sftp://user:pass@x.x.x.x:21")
        ttk.Entry(style_bar, textvariable=self.guide_url_var, width=28).pack(side=tk.LEFT, padx=4)
        ttk.Label(style_bar, text="  DHCP池起止:").pack(side=tk.LEFT, padx=(8, 2))
        self.guide_rng1_var = tk.StringVar(value="")
        ttk.Entry(style_bar, textvariable=self.guide_rng1_var, width=13).pack(side=tk.LEFT, padx=2)
        ttk.Label(style_bar, text="~").pack(side=tk.LEFT)
        self.guide_rng2_var = tk.StringVar(value="")
        ttk.Entry(style_bar, textvariable=self.guide_rng2_var, width=13).pack(side=tk.LEFT, padx=2)

        self.guide_text = scrolledtext.ScrolledText(self.tab_guide, font=("Consolas", 9), wrap=tk.WORD)
        self.guide_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        self.guide_text.insert("1.0", self._guide_static())

    def _on_vendor_change(self, *_):
        """厂商预设联动：引导文件名 / 映射表开关 / DHCP 输出风格。"""
        disp = self.gen_vendor_var.get()
        key = next((k for k, v in VENDOR_PRESETS.items() if v["name"] == disp), None)
        p = VENDOR_PRESETS.get(key)
        if not p:
            return
        self.gen_bootfile_var.set(p["bootfile"])
        self.gen_mapping_var.set(p["mapping"])
        if hasattr(self, "guide_style_var"):
            self.guide_style_var.set(DHCP_STYLES.get(p["dhcp"], "isc"))

    def _on_archive_toggle(self):
        self.gen_archive_hint.set("每次生成存到 时间戳子目录，保留最近 10 次" if self.gen_archive_var.get()
                                  else "直接输出到所选目录（不归档）")

    def _latest_output_dir(self) -> str:
        """实际输出目录：归档模式 → 最新时间戳子目录（并清理超 10 个的旧归档）。

        同分钟重复生成 → 追加序号（_2、_3），避免互相覆盖。
        """
        base = self.output_var.get().strip()
        if self.gen_archive_var.get():
            try:
                os.makedirs(base, exist_ok=True)
                arch = sorted([d for d in os.listdir(base)
                               if re.fullmatch(r"\d{8}_\d{4}(?:_\d+)?", d)])
                for old in arch[:-10]:
                    shutil.rmtree(os.path.join(base, old), ignore_errors=True)
                ts = datetime.now().strftime("%Y%m%d_%H%M")
                if os.path.isdir(os.path.join(base, ts)):
                    n = 2
                    while os.path.isdir(os.path.join(base, f"{ts}_{n}")):
                        n += 1
                    ts = f"{ts}_{n}"
                return os.path.join(base, ts)
            except Exception:
                return base
        return base

    def _on_gen_usb(self):
        """生成 U盘 ZTP 包：usb_config.ini + 配置文件拷贝到 输出目录/usb/。"""
        if not self.devices:
            messagebox.showwarning("提示", "设备清单为空")
            return
        out_dir = self._last_gen_dir or self._latest_output_dir()
        usb_dir = os.path.join(out_dir, "usb")
        try:
            os.makedirs(usb_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("目录错误", str(e))
            return
        ini = build_usb_ini(self.devices,
                            fileserver=self.ztp_server_var.get().strip() or "file:/usb:",
                            system_software=self.ztp_sw_var.get().strip(),
                            system_pat=self.ztp_pat_var.get().strip())
        with open(os.path.join(usb_dir, "usb_config.ini"), "w", encoding="utf-8", newline="\r\n") as f:
            f.write(ini)
        copied, skipped = 0, 0
        for dev in self.devices:
            host = (dev.get("hostname") or "").strip()
            if not host:
                continue
            src = os.path.join(out_dir, host + ".cfg")
            if os.path.isfile(src):
                shutil.copy(src, os.path.join(usb_dir, host + ".cfg"))
                copied += 1
            else:
                skipped += 1
        n_blocks = ini.count("[DEVICE")
        self._gen_log(f"[{datetime.now().strftime('%H:%M:%S')}] U盘 ZTP 包已生成 → {usb_dir}")
        self._gen_log(f"  ✓ usb_config.ini（{n_blocks} 个设备匹配块, FILESERVER={self.ztp_server_var.get().strip() or 'file:/usb:'}）")
        self._gen_log(f"  ✓ 已复制 {copied} 个配置文件" + (f"，{skipped} 台未生成 cfg（先点「生成脚本」）" if skipped else ""))
        self._gen_log("  使用: 把 usb 目录内容拷到 U 盘根目录 → 插到交换机后开机，设备自动按 ESN 匹配并加载配置")
        self._set_status(f"U盘 ZTP 包生成完成: {n_blocks} 台 / 配置文件 {copied} 个")
        try:
            os.startfile(usb_dir)
        except Exception:
            pass

    def _on_gen_dhcp(self):
        ip = self.tftp_ip_var.get().strip() or "x.x.x.x"
        style_disp = self.guide_style_var.get().strip()
        style = next((k for k, v in DHCP_STYLES.items() if v == style_disp), "isc")
        vlan = self.guide_vlan_var.get().strip() or "11"
        net = self.guide_net_var.get().strip() or ""
        gw = self.guide_gw_var.get().strip() or ""
        r1 = self.guide_rng1_var.get().strip()
        r2 = self.guide_rng2_var.get().strip()
        bootfile = self.gen_bootfile_var.get().strip() if hasattr(self, "gen_bootfile_var") else ""
        if style == "isc":
            snippet = build_dhcp_isc(self.devices, ip, bootfile, network=net, gateway=gw,
                                     range_start=r1, range_end=r2)
        elif style == "dnsmasq":
            snippet = build_dhcp_dnsmasq(self.devices, ip, bootfile, network=net, gateway=gw,
                                         range_start=r1, range_end=r2)
        elif style == "vrp146":
            snippet = dhcp_snippet(ip, bootfile or BOOTFILE, vlan, net or "192.168.100.0",
                                   gw or "192.168.100.254", mode="146",
                                   fileserver_url=self.guide_url_var.get().strip())
        else:
            snippet = dhcp_snippet(ip, bootfile or BOOTFILE, vlan, net or "192.168.100.0",
                                   gw or "192.168.100.254", mode="67")
        self.guide_text.delete("1.0", tk.END)
        self.guide_text.insert("1.0", snippet)

    @staticmethod
    def _guide_static() -> str:
        return """============================================================
NADT 自动化开局流程（Zero-Touch Provisioning）
============================================================
【方式 A: DHCP 开局】（有 DHCP 服务器）
1. 在 ① 设备清单 填入交换机 ESN / 主机名 / 类型 / 管理IP
2. 在 ② 配置模板 按需修改模板（{{占位符}} 自动替换）
3. 在 ③ 生成脚本 → 输出目录（.cfg × N + lswnet.cfg）
4. 在 ④ 启动 TFTP 服务器，根目录指向输出目录
5. 记下本机 IP（④ 页显示），填到下方 DHCP 配置 → 生成配置
6. 把 DHCP 配置粘贴到 DHCP 服务器（华为设备/路由器）
7. 交换机出厂/恢复出厂开机 → 自动获取 IP → 下载配置 → 自动开局

【方式 B: U盘 ZTP】（无 DHCP，现场插 U 盘）
1. 同上生成脚本后，在 ③ 页点「生成 U盘 ZTP 包」
2. 把 输出目录/usb/ 里的 usb_config.ini + 全部 .cfg 拷到 U 盘根目录
3. U 盘插到交换机 → 开机 → 设备自动按 ESN/MAC 匹配加载配置
   （堆叠设备: 设备扩展字段填 stack_member_id=成员ID，自动生成 STACK-MEMBER-ID）

说明:
- lswnet.cfg 是 ESN 映射表: esn=序列号;[vrpfile=版本;][patchfile=补丁;]cfgfile=配置;
  可同时下发版本文件/补丁（③ 页填写，或设备扩展字段 vrp_file/patch_file）
- 华为交换机 EasyDeploy 用 DHCP option 66（文件服务器）+ option 146（netfile=lswnet.cfg）
- 交换机管理网段与 DHCP 服务器必须互通，TFTP 端口 69 需放行

常见问题:
- 交换机没拉配置 → 确认 DHCP option 66/146（或 67）已配、服务器可达、文件名一致
- 拉错配置 → 检查 lswnet.cfg / usb_config.ini 里的 ESN 是否与设备序列号一致
  （设备上执行 display esn / display device 查看序列号）
- U盘开局没反应 → 确认 usb_config.ini 在 U 盘根目录、文件名正确、文件为 CRLF 换行
============================================================
"""

    # ── 状态栏 / status bar ──
    def _set_status(self, msg: str):
        self.status_var.set(msg)


def main():
    root = tk.Tk()
    NADTGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
