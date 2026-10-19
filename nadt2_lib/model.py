"""NADT 2.0 项目模型：一个 .nadt 文件 = 设备 + 模板 + 映射 + 凭据 + 生成设置。"""

from __future__ import annotations

import copy
import base64
import hashlib
import json
import os
import re
import shutil
import tempfile
from dataclasses import dataclass, field
from datetime import datetime

import nadt
from . import __version__, secrets

FORMAT = "nadt-project"
FORMAT_VERSION = 2
PROJECT_EXT = ".nadt"

TEMPLATE_VAR_ALIASES = {
    "sysname": "hostname",
    "ip": "mgmt_ip",
    "wg": "gateway",
    "access": "access_vlan",
    "vlanif": "mgmt_vlan",
}
STD_FIELDS = set(nadt.CSV_HEADER)
COLUMN_TITLES = {
    "hostname": "主机名", "type": "类型", "esn": "ESN", "note": "备注",
    "mgmt_ip": "管理IP", "mgmt_vlan": "管理VLAN", "access_vlan": "接入VLAN",
    "mask": "掩码", "gateway": "网关",
}


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def config_dir() -> str:
    base = os.environ.get("APPDATA") or os.environ.get("XDG_CONFIG_HOME") \
        or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "NADT2")


def recent_path() -> str:
    return os.path.join(config_dir(), "recent.json")


def load_recent() -> list[str]:
    try:
        with open(recent_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
        return [str(x) for x in data if isinstance(x, str)][:10]
    except Exception:
        return []


def add_recent(path: str):
    path = os.path.abspath(path)
    items = [p for p in load_recent() if os.path.normcase(p) != os.path.normcase(path)]
    items.insert(0, path)
    try:
        os.makedirs(config_dir(), exist_ok=True)
        with open(recent_path(), "w", encoding="utf-8") as f:
            json.dump(items[:10], f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def clear_recent():
    try:
        os.remove(recent_path())
    except OSError:
        pass


@dataclass
class Settings:
    filename_rule: str = "hostname"
    extension: str = "cfg"
    output_dir: str = "output"
    template_subdir: bool = True
    timestamp_archive: bool = False
    generate_lswnet: bool = True
    eol: str = "crlf"
    scope: str = "all"              # all | type
    default_template: str = ""
    vendor: str = "generic"
    bootfile: str = nadt.BOOTFILE
    dhcp_style: str = "isc"
    tftp_port: int = 69
    tftp_concurrency: int = 500
    mask_secrets: bool = True
    autosave: bool = True

    def to_dict(self) -> dict:
        return dict(self.__dict__)

    @classmethod
    def from_dict(cls, data: dict | None) -> "Settings":
        s = cls()
        for k, v in (data or {}).items():
            if hasattr(s, k) and v is not None:
                setattr(s, k, v)
        return s


@dataclass
class Credential:
    key: str
    value: str = ""
    label: str = ""
    protected: bool = True

    def to_dict(self) -> dict:
        if self.protected:
            token = secrets.protect(self.value)
        else:
            token = secrets.PLAIN_PREFIX + base64.b64encode(
                self.value.encode("utf-8")).decode("ascii")
        return {"label": self.label, "protected": self.protected, "value": token}

    @classmethod
    def from_dict(cls, key: str, data: dict | None) -> "Credential":
        data = data or {}
        return cls(key=key, value=secrets.unprotect(data.get("value", "")),
                   label=str(data.get("label", "")),
                   protected=bool(data.get("protected", True)))


@dataclass
class Template:
    name: str
    text: str = ""
    type_name: str = ""

    @property
    def stem(self) -> str:
        return os.path.splitext(os.path.basename(self.name))[0] or self.name


@dataclass
class Project:
    name: str = "未命名方案"
    devices: list[dict] = field(default_factory=list)
    templates: dict[str, Template] = field(default_factory=dict)
    type_map: dict[str, str] = field(default_factory=dict)
    variables: dict[str, str] = field(default_factory=dict)
    credentials: dict[str, Credential] = field(default_factory=dict)
    settings: Settings = field(default_factory=Settings)
    path: str | None = None
    created: str = field(default_factory=now_str)
    modified: str = field(default_factory=now_str)
    last_generation: dict = field(default_factory=dict)
    dirty: bool = False
    _undo: list = field(default_factory=list, repr=False)
    _redo: list = field(default_factory=list, repr=False)

    # ── 创建 / 模板 ──
    @classmethod
    def create_default(cls, name: str = "未命名方案") -> "Project":
        p = cls(name=name)
        text = ""
        candidates = [
            os.path.join(nadt.TEMPLATE_DIR, "access-switch.cfg"),
            os.path.join(getattr(nadt, "RESOURCE_DIR", nadt.BASE_DIR),
                         "templates", "access-switch.cfg"),
        ]
        for candidate in candidates:
            try:
                with open(candidate, "r", encoding="utf-8") as f:
                    text = f.read()
                break
            except OSError:
                continue
        if not text:
            text = "# {{hostname}}\n"
        tpl = Template(name="access-switch.cfg", text=text, type_name="generic")
        p.templates[tpl.name] = tpl
        p.settings.default_template = tpl.name
        p.type_map["GENERIC"] = tpl.name
        p.settings.output_dir = "output"
        p.dirty = False
        return p

    def template(self, name: str | None = None) -> Template:
        name = name or self.settings.default_template
        if name in self.templates:
            return self.templates[name]
        if self.templates:
            return next(iter(self.templates.values()))
        raise ValueError("项目没有模板")

    def template_for(self, dev: dict) -> Template:
        devtype = str(dev.get("type") or "").strip()
        target = self.type_map.get(devtype.upper())
        if target in self.templates:
            return self.templates[target]
        return self.template()

    def add_template(self, name: str, text: str = "", type_name: str = "") -> Template:
        name = os.path.basename(name.strip()) or "template.cfg"
        if name in self.templates:
            raise ValueError(f"模板已存在: {name}")
        tpl = Template(name=name, text=text, type_name=type_name or os.path.splitext(name)[0])
        self.templates[name] = tpl
        if type_name:
            self.type_map[type_name.upper()] = name
        if not self.settings.default_template:
            self.settings.default_template = name
        self.mark_dirty()
        return tpl

    def remove_template(self, name: str):
        self.templates.pop(name, None)
        for k in [k for k, v in self.type_map.items() if v == name]:
            del self.type_map[k]
        if self.settings.default_template == name:
            self.settings.default_template = next(iter(self.templates), "")
        self.mark_dirty()

    def rename_template(self, old: str, new: str) -> str:
        new = os.path.basename(new.strip())
        if not new or new == old:
            return old
        if new in self.templates:
            raise ValueError(f"模板已存在: {new}")
        tpl = self.templates.pop(old)
        tpl.name = new
        self.templates[new] = tpl
        for k, v in list(self.type_map.items()):
            if v == old:
                self.type_map[k] = new
        if self.settings.default_template == old:
            self.settings.default_template = new
        self.mark_dirty()
        return new

    # ── 变量 / 渲染 ──
    def base_vars(self) -> dict[str, str]:
        values = {str(k): str(v) for k, v in self.variables.items()}
        for k, cred in self.credentials.items():
            values[k] = cred.value
        return values

    def render(self, dev: dict, template_name: str | None = None) -> str:
        tpl = self.template(template_name)
        return nadt.build_switch_config(tpl.text, dev, extra_vars=self.base_vars())

    def target_devices(self, scope: str | None = None) -> list[dict]:
        scope = scope or self.settings.scope
        if scope != "type":
            return list(self.devices)
        tpl = self.template()
        names = {tpl.stem.lower(), tpl.type_name.lower()}
        return [d for d in self.devices
                if (d.get("type") or "").strip().lower() in ("", *names)]

    @staticmethod
    def template_vars(text: str) -> list[str]:
        return nadt.template_variables(text)

    def columns(self, template_name: str | None = None,
                text: str | None = None) -> list[str]:
        tpl = self.template(template_name)
        params = [p for p in self.template_vars(text if text is not None else tpl.text)
                  if p not in ("hostname", "sysname", "type", "esn", "note")]
        # 项目变量/凭据由项目统一提供，不进入设备表（避免把密码显示成设备列）
        reserved = set(self.variables) | set(self.credentials)
        params = [p for p in params if p not in reserved]
        return list(dict.fromkeys(["hostname", "type", "esn"] + params + ["note"]))

    @staticmethod
    def header(col: str) -> str:
        return COLUMN_TITLES.get(col, "{{" + col + "}}")

    @staticmethod
    def cell_get(dev: dict, col: str) -> str:
        extras = dev.get("_extras") or {}
        if col in ("hostname", "sysname"):
            return str(dev.get("hostname") or extras.get("sysname") or "")
        alias = TEMPLATE_VAR_ALIASES.get(col)
        if alias:
            return str(extras.get(col) or dev.get(alias) or "")
        if col in STD_FIELDS:
            return str(dev.get(col) or "")
        return str(extras.get(col, "") or "")

    @staticmethod
    def cell_set(dev: dict, col: str, value: str):
        value = str(value).replace("\r", "").replace("\n", " ")
        extras = dev.setdefault("_extras", {})
        if col in ("hostname", "sysname"):
            dev["hostname"] = value
            extras["sysname"] = value
            return
        alias = TEMPLATE_VAR_ALIASES.get(col)
        if alias:
            dev[alias] = value
            extras[col] = value
            return
        if col in STD_FIELDS:
            dev[col] = value
            extras[col] = value
            return
        extras[col] = value

    # ── 文件名 ──
    @staticmethod
    def is_bad_filename(name: str) -> bool:
        return bool(re.search(r'[\\/:*?"<>|]', name)) or name.upper() in {
            "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5",
            "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4",
            "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"}

    def filename_base(self, dev: dict, rule: str | None = None) -> str:
        rule = rule or self.settings.filename_rule
        if rule == "__single__":
            return self.template().stem
        if rule in ("hostname", "sysname"):
            val = dev.get("hostname") or (dev.get("_extras") or {}).get("sysname") or ""
        elif rule in TEMPLATE_VAR_ALIASES:
            val = (dev.get("_extras") or {}).get(rule) or \
                dev.get(TEMPLATE_VAR_ALIASES[rule], "")
        elif rule in STD_FIELDS:
            val = dev.get(rule, "")
        else:
            val = (dev.get("_extras") or {}).get(rule, "")
        val = str(val).strip()
        if not val:
            raise ValueError(f"文件名规则 {rule} 的值为空")
        if self.is_bad_filename(val):
            raise ValueError(f"文件名含非法字符: {val}")
        return val

    def resolve_output_dir(self, template_name: str | None = None) -> str:
        out = (self.settings.output_dir or "output").strip()
        if not os.path.isabs(out):
            base = os.path.dirname(self.path) if self.path else os.getcwd()
            out = os.path.join(base, out)
        if self.settings.template_subdir:
            tpl_name = self.template(template_name).stem
            out = os.path.join(out, re.sub(r'[\\/:*?"<>|]+', "_", tpl_name))
        if self.settings.timestamp_archive:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            candidate = os.path.join(out, stamp)
            n = 2
            while os.path.exists(candidate):
                candidate = os.path.join(out, f"{stamp}_{n}")
                n += 1
            out = candidate
        return out

    @staticmethod
    def project_output_base(settings: Settings, project_path: str | None) -> str:
        out = (settings.output_dir or "output").strip()
        if not os.path.isabs(out):
            base = os.path.dirname(project_path) if project_path else os.getcwd()
            out = os.path.join(base, out)
        return out

    # ── 持久化 ──
    def mark_dirty(self):
        self.dirty = True
        self.modified = now_str()

    def to_dict(self) -> dict:
        return {
            "format": FORMAT,
            "format_version": FORMAT_VERSION,
            "app_version": __version__,
            "name": self.name,
            "created": self.created,
            "modified": self.modified,
            "devices": copy.deepcopy(self.devices),
            "templates": {n: {"text": t.text, "type": t.type_name}
                          for n, t in self.templates.items()},
            "type_map": dict(self.type_map),
            "variables": dict(self.variables),
            "credentials": {k: c.to_dict() for k, c in self.credentials.items()},
            "settings": self.settings.to_dict(),
            "last_generation": copy.deepcopy(self.last_generation),
        }

    @classmethod
    def from_dict(cls, data: dict) -> "Project":
        if data.get("format") != FORMAT:
            raise ValueError("不是 NADT 项目文件")
        if int(data.get("format_version", 0)) > FORMAT_VERSION:
            raise ValueError("项目文件由更高版本 NADT 创建，无法打开")
        p = cls(name=str(data.get("name") or "未命名方案"),
                created=str(data.get("created") or now_str()),
                modified=str(data.get("modified") or now_str()))
        p.devices = [dict(d) for d in (data.get("devices") or [])]
        for n, t in (data.get("templates") or {}).items():
            p.templates[n] = Template(name=n, text=str(t.get("text", "")),
                                      type_name=str(t.get("type") or
                                                    os.path.splitext(n)[0]))
        p.type_map = {str(k): str(v) for k, v in (data.get("type_map") or {}).items()}
        p.variables = {str(k): str(v) for k, v in (data.get("variables") or {}).items()}
        p.credentials = {k: Credential.from_dict(k, v)
                         for k, v in (data.get("credentials") or {}).items()}
        p.settings = Settings.from_dict(data.get("settings"))
        if not p.settings.default_template and p.templates:
            p.settings.default_template = next(iter(p.templates))
        p.last_generation = data.get("last_generation") or {}
        p.dirty = False
        return p

    def save(self, path: str | None = None, backup: bool = True):
        path = os.path.abspath(path or self.path or "")
        if not path:
            raise ValueError("未指定项目文件路径")
        if not path.lower().endswith(PROJECT_EXT):
            path += PROJECT_EXT
        os.makedirs(os.path.dirname(path), exist_ok=True)
        if backup and os.path.isfile(path):
            try:
                shutil.copy2(path, path + ".bak")
            except OSError:
                pass
        payload = self.to_dict()
        fd, tmp = tempfile.mkstemp(prefix=".nadt-", suffix=".tmp",
                                   dir=os.path.dirname(path))
        try:
            with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
                f.write("\n")
            os.replace(tmp, path)
        finally:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except OSError:
                pass
        self.path = path
        self.dirty = False
        self.modified = now_str()
        add_recent(path)
        return path

    @classmethod
    def load(cls, path: str) -> "Project":
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        p = cls.from_dict(data)
        p.path = os.path.abspath(path)
        add_recent(p.path)
        return p

    def autosave_path(self) -> str:
        return (self.path or os.path.join(config_dir(), "untitled.nadt")) + ".autosave"

    def write_autosave(self):
        try:
            os.makedirs(os.path.dirname(self.autosave_path()), exist_ok=True)
            with open(self.autosave_path(), "w", encoding="utf-8", newline="\n") as f:
                json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def clear_autosave(self):
        try:
            os.remove(self.autosave_path())
        except OSError:
            pass

    def snapshot(self) -> dict:
        return self.to_dict()

    def restore(self, snap: dict):
        other = Project.from_dict(snap)
        self.devices = copy.deepcopy(other.devices)
        self.templates = copy.deepcopy(other.templates)
        self.type_map = dict(other.type_map)
        self.variables = dict(other.variables)
        self.credentials = copy.deepcopy(other.credentials)
        self.settings = copy.deepcopy(other.settings)
        self.name = other.name
        self.mark_dirty()

    # ── 统计 / 指纹 ──
    def stats(self) -> dict:
        return {
            "devices": len(self.devices),
            "templates": len(self.templates),
            "types": len({d.get("type", "") for d in self.devices if d.get("type")}),
            "credentials": len(self.credentials),
        }

    def config_hash(self, dev: dict) -> str:
        try:
            return hashlib.sha256(self.render(dev).encode("utf-8")).hexdigest()
        except Exception:
            return ""
