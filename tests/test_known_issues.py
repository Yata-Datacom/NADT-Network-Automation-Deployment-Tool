"""已知问题（bug）特征化用例 / characterization tests.

这些用例锁定的是「当前有问题的行为」，不是期望行为。
目的是：一旦 nadt.py 里对应 bug 被修好，这里会立刻变红，提醒同步更新用例与报告。
每个用例 docstring 里标了 nadt.py 的行号与现象，详见本次测试交付报告。

修 bug 时请同步更新这些用例（并删掉已修复条目）。
"""

import time


import nadt as N


def test_known_issue_spaced_placeholder_silently_survives():
    """BUG-1  nadt.py:89-94（_subst）+ 250-252（残留检测）

    占位符写成 {{ hostname }}（带空格）时：_subst 的正则 \\{\\{(\\w+)\\}\\} 不匹配，
    残留检测用的也是同一个正则，于是"模板缺变量"保护失效，
    生成的 .cfg 里会原样留下 {{ hostname }} —— 交换机拿到残缺配置且无人报警。
    最小复现见下面两行。
    """
    cfg = N.build_switch_config("sysname {{ hostname }}\nvlan {{mgmt_vlan}}\n",
                                {"hostname": "SW-1-1F", "mgmt_vlan": "11",
                                 "mgmt_ip": "192.0.2.11"})
    assert cfg == "sysname {{ hostname }}\nvlan 11\n"      # ← 期望 throw，实际静默通过


def test_known_issue_inline_directive_not_at_line_start_passes_through():
    """BUG-2  nadt.py:154-165（单行内联块只在行首匹配）+ 250（残留检测抓不到 {{#if）

    行首不是 {{#if/{{#for 的内联块整行原样输出，残留检测只匹配 {{word}} 形式
    （{{#if 不是 \\w+），所以又一次静默产出带模板指令的配置。
    """
    cfg = N.build_switch_config("sysname X{{#if a}}Y{{#endif}}Z\n", {"hostname": "SW-1-1F"})
    assert cfg == "sysname X{{#if a}}Y{{#endif}}Z\n"       # ← 期望渲染/报错，实际原样输出
    out = N.render_template("description {{#for n from 1 to 2}}p{{n}}{{#endfor}}\n", {})
    assert out == "description {{#for n from 1 to 2}}p{{n}}{{#endfor}}\n"


def test_known_issue_else_and_endif_bodies_are_dropped_silently():
    """BUG-3  nadt.py:135-138（{{#else}} / {{#endif}} 判定用 startswith）

    {{#if}} / {{#for}} 指令行带内容会报错，但 {{#else}} BODY / {{#endif}} TAIL
    一律静默丢弃行内其余文本（不是报错），用户不会发现配置行丢了。
    """
    assert N.render_template("{{#if a}}\nA\n{{#else}} 需要保留的文本\nB\n{{#endif}}\n",
                             {"a": ""}) == "B\n"
    assert N.render_template("{{#if a}}\nA\n{{#endif}} 行尾注释\n", {"a": "1"}) == "A\n"


def test_known_issue_mapping_sheet_warning_is_misleading():
    """BUG-4  nadt.py:592-594 早退 vs 611-615 映射表识别

    「无参数列」的提前 return 在映射表识别之前，所以真实 lswnet sheet
    （A 列全是 esn=...;cfgfile=...;）命中的是"无参数列"分支，
    611-615 的专用识别与友好提示对真实数据不可达（仅当第 1 行 C 列有参数名时才可达）。
    """
    warnings = []
    rows = [["【lswnet】"],
            ["esn=0000-1111-2222;cfgfile=SW-1-1F.cfg;"]]
    templates, devices = N.parse_sheet_rows(rows, "lswnet", warnings)
    assert (templates, devices) == ({}, [])                # 结果正确（被跳过）
    assert warnings == ["lswnet: 无参数列（第 1 行 C 列起应为 {参数名}）"]   # ← 提示误导


def test_known_issue_count_expansions_is_not_memory_lightweight():
    """BUG-5  nadt.py:406-424（count_expansions）→ 323-344（_parse_range_values）

    docstring 说"轻量预估（不实际展开）"，但 _parse_range_values 会把区间
    完整物化成 list[str]：{1-1000000} 实测 peak ≈ 63MB、约 0.2s；
    {1-100000000} 会直接把内存打满 —— 防组合爆炸的保护挡不住单个超大区间。
    """
    started = time.time()
    assert N.count_expansions({"hostname": "SW-{1-1000000}"}) == 1_000_000
    assert time.time() - started < 10          # 能算完，但代价是 O(区间长度) 的内存


def test_known_issue_invalid_mac_emits_empty_mac_line(tmp_path):
    """BUG-6  nadt.py:297 + 306-307（build_usb_ini）

    MAC 非空但格式非法时 normalize_mac 返回 ""，代码仍会写出 "MAC=" 空行
    （华为 ini 里是空 MAC 项，可能被设备忽略或解析异常），应整行跳过。
    """
    ini = N.build_usb_ini([{"hostname": "SW-1-1F", "mac": "not-a-mac"}])
    assert "MAC=\r\n" in ini                    # ← 期望完全没有 MAC= 行
    assert "ESN=DEFAULT\r\n" in ini
