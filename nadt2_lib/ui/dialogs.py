"""NADT 2.0 对话框：生成预检、Diff 查看、列映射、批量操作、预览。"""

from __future__ import annotations

import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import nadt
from .. import services
from ..model import Project
from . import theme
from .widgets import DarkText, ScrollableFrame, center_window


class _Dialog(tk.Toplevel):
    def __init__(self, parent, title: str, width: int, height: int):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=theme.BG)
        self.transient(parent)
        self.resizable(False, False)
        center_window(self, width, height)
        self.result = None
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Escape>", lambda _e: self._cancel())

    def _cancel(self):
        self.result = None
        self.destroy()

    def show(self):
        self.grab_set()
        self.wait_window()
        return self.result


class FillDialog(_Dialog):
    def __init__(self, parent, columns: list[str]):
        super().__init__(parent, "批量填充", 420, 230)
        body = ttk.Frame(self, padding=16)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text="把选中行的某一列统一填成同一个值。").pack(anchor="w", pady=(0, 10))
        ttk.Label(body, text="列").pack(anchor="w")
        self.col = tk.StringVar(value=columns[0] if columns else "")
        ttk.Combobox(body, textvariable=self.col, values=columns,
                     state="readonly").pack(fill=tk.X, pady=(2, 10))
        ttk.Label(body, text="值").pack(anchor="w")
        self.value = tk.StringVar()
        entry = ttk.Entry(body, textvariable=self.value)
        entry.pack(fill=tk.X, pady=(2, 12))
        entry.focus_set()
        row = ttk.Frame(body)
        row.pack(fill=tk.X)
        ttk.Button(row, text="取消", style="Tool.TButton", command=self._cancel).pack(side=tk.RIGHT)
        ttk.Button(row, text="填充", style="Accent.TButton",
                   command=self._ok).pack(side=tk.RIGHT, padx=(0, 6))

    def _ok(self):
        self.result = {"column": self.col.get(), "value": self.value.get()}
        self.destroy()


class SequenceDialog(_Dialog):
    def __init__(self, parent, columns: list[str]):
        super().__init__(parent, "批量序列", 480, 390)
        body = ttk.Frame(self, padding=16)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text="按选中行顺序写入数字序列。").pack(anchor="w", pady=(0, 8))
        grid = ttk.Frame(body)
        grid.pack(fill=tk.X)
        self.vars = {}
        fields = [("column", "列", "string", columns[0] if columns else "", columns),
                  ("start", "起始值", "string", "1", None),
                  ("step", "步长", "string", "1", None),
                  ("padding", "补零位数", "string", "0", None),
                  ("prefix", "前缀", "string", "", None),
                  ("suffix", "后缀", "string", "", None)]
        for r, (key, label, _t, default, values) in enumerate(fields):
            ttk.Label(grid, text=label).grid(row=r, column=0, sticky="w", pady=4)
            var = tk.StringVar(value=default)
            if values:
                ttk.Combobox(grid, textvariable=var, values=values,
                             state="readonly").grid(row=r, column=1, sticky="ew", padx=(10, 0))
            else:
                ttk.Entry(grid, textvariable=var).grid(row=r, column=1, sticky="ew", padx=(10, 0))
            self.vars[key] = var
        grid.columnconfigure(1, weight=1)
        ttk.Label(body, text="示例：start=10, step=10, padding=3, prefix=V → V010 / V020 …",
                  style="Muted.TLabel").pack(anchor="w", pady=(10, 12))
        row = ttk.Frame(body)
        row.pack(fill=tk.X)
        ttk.Button(row, text="取消", style="Tool.TButton", command=self._cancel).pack(side=tk.RIGHT)
        ttk.Button(row, text="应用", style="Accent.TButton",
                   command=self._ok).pack(side=tk.RIGHT, padx=(0, 6))

    def _ok(self):
        try:
            result = {
                "column": self.vars["column"].get(),
                "start": int(self.vars["start"].get()),
                "step": int(self.vars["step"].get()),
                "padding": max(0, int(self.vars["padding"].get() or 0)),
                "prefix": self.vars["prefix"].get(),
                "suffix": self.vars["suffix"].get(),
            }
        except ValueError:
            messagebox.showerror("输入错误", "起始值/步长/补零位数必须是整数", parent=self)
            return
        self.result = result
        self.destroy()


class FindReplaceDialog(_Dialog):
    def __init__(self, parent):
        super().__init__(parent, "查找替换", 460, 300)
        body = ttk.Frame(self, padding=16)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text="在所有设备的所有参数列中查找替换。").pack(anchor="w", pady=(0, 8))
        ttk.Label(body, text="查找").pack(anchor="w")
        self.find = tk.StringVar()
        ttk.Entry(body, textvariable=self.find).pack(fill=tk.X, pady=(2, 8))
        ttk.Label(body, text="替换为").pack(anchor="w")
        self.replace = tk.StringVar()
        ttk.Entry(body, textvariable=self.replace).pack(fill=tk.X, pady=(2, 8))
        self.regex = tk.BooleanVar(value=False)
        self.case = tk.BooleanVar(value=False)
        row = ttk.Frame(body)
        row.pack(fill=tk.X)
        ttk.Checkbutton(row, text="正则表达式", variable=self.regex,
                        style="Panel.TCheckbutton").pack(side=tk.LEFT)
        ttk.Checkbutton(row, text="区分大小写", variable=self.case,
                        style="Panel.TCheckbutton").pack(side=tk.LEFT, padx=(12, 0))
        btns = ttk.Frame(body)
        btns.pack(fill=tk.X, pady=(14, 0))
        ttk.Button(btns, text="取消", style="Tool.TButton", command=self._cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="替换", style="Accent.TButton",
                   command=self._ok).pack(side=tk.RIGHT, padx=(0, 6))

    def _ok(self):
        if not self.find.get():
            messagebox.showwarning("提示", "请输入要查找的内容", parent=self)
            return
        self.result = {"find": self.find.get(), "replace": self.replace.get(),
                       "regex": self.regex.get(), "case": self.case.get()}
        self.destroy()


class ColumnMapDialog(_Dialog):
    """CSV/Excel 列 → 项目参数列 映射。"""

    def __init__(self, parent, header: list[str], rows: list[list[str]],
                 columns: list[str], project: Project):
        super().__init__(parent, "导入表格 — 列映射", 760, 620)
        self.header = header
        self.rows = rows
        self.columns = columns
        self.project = project
        self.choices: dict[int, tk.StringVar] = {}

        top = ttk.Frame(self, padding=(16, 14, 16, 4))
        top.pack(fill=tk.X)
        ttk.Label(top, text=f"共 {len(rows)} 行 · {len(header)} 列。选择每列写入哪个项目参数。",
                  style="Panel.TLabel").pack(anchor="w")
        ttk.Label(top, text="未识别的列默认忽略；主机名/ESN/管理IP 等标准字段会自动匹配。",
                  style="PanelMuted.TLabel").pack(anchor="w", pady=(2, 0))

        scroll = ScrollableFrame(self)
        scroll.pack(fill=tk.BOTH, expand=True, padx=16, pady=8)
        grid = scroll.inner

        ttk.Label(grid, text="源列", style="PanelMuted.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(grid, text="前 3 行示例", style="PanelMuted.TLabel").grid(row=0, column=1, sticky="w")
        ttk.Label(grid, text="写入项目参数", style="PanelMuted.TLabel").grid(row=0, column=2, sticky="w")
        display = ["忽略"] + [f"{c}  {Project.header(c)}" for c in columns]
        display_to_key = {d: (None if d == "忽略" else c)
                          for d, c in zip(display, ["__ignore__"] + columns)}
        for i, h in enumerate(header):
            ttk.Label(grid, text=f"{i + 1}. {h or '(空表头)'}").grid(
                row=i + 1, column=0, sticky="w", pady=3)
            sample = " | ".join(str(r[i])[:18] for r in rows[:3] if i < len(r))
            ttk.Label(grid, text=sample[:42], style="Muted.TLabel").grid(
                row=i + 1, column=1, sticky="w", padx=(10, 10))
            key = self._auto_map(h)
            show = next((d for d, k in display_to_key.items() if k == key), "忽略")
            var = tk.StringVar(value=show)
            ttk.Combobox(grid, textvariable=var, values=display, state="readonly",
                         width=24).grid(row=i + 1, column=2, sticky="ew", padx=(10, 0))
            self.choices[i] = var
        grid.columnconfigure(2, weight=1)

        btns = ttk.Frame(self, padding=(16, 4, 16, 14))
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="取消", style="Tool.TButton", command=self._cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="导入", style="Accent.TButton",
                   command=self._ok).pack(side=tk.RIGHT, padx=(0, 6))

    def _auto_map(self, header: str) -> str | None:
        h = header.strip().lower()
        if not h:
            return None
        h = h.replace("{{", "").replace("}}", "").strip()
        for c in self.columns:
            if h == c.lower() or h == Project.header(c).lower():
                return c
        alias = nadt.CSV_ALIASES.get(h)
        if alias in self.columns:
            return alias
        return None

    def _ok(self):
        mapping = {}
        display_to_key = {"忽略": None}
        for c in self.columns:
            display_to_key[f"{c}  {Project.header(c)}"] = c
        for i, var in self.choices.items():
            key = display_to_key.get(var.get())
            if key:
                mapping[i] = key
        if not mapping:
            messagebox.showwarning("提示", "至少映射一列", parent=self)
            return
        self.result = mapping
        self.destroy()


class GenerationDialog(_Dialog):
    """生成预检：变更统计 + 每个文件的 Diff。"""

    STATUS = {"added": "新增", "changed": "修改", "unchanged": "不变", "error": "错误"}

    def __init__(self, parent, plan: services.GenerationPlan, project: Project):
        super().__init__(parent, "生成预检", 980, 680)
        self.plan = plan
        self.project = project
        self.resizable(True, True)
        self.minsize(820, 560)

        top = ttk.Frame(self, padding=(16, 14, 16, 6))
        top.pack(fill=tk.X)
        ttk.Label(top, text="生成预检", style="Panel.TLabel",
                  font=theme.FONT_TITLE).pack(anchor="w")
        ttk.Label(top, text=f"输出目录: {plan.out_dir}", style="PanelMuted.TLabel").pack(anchor="w")
        stats = ttk.Frame(top)
        stats.pack(fill=tk.X, pady=(6, 0))
        for text, color in ((f"新增 {plan.added}", theme.GREEN),
                            (f"修改 {plan.changed}", theme.AMBER),
                            (f"不变 {plan.unchanged}", theme.MUTED),
                            (f"错误 {len(plan.errors)}", theme.RED if plan.errors else theme.MUTED)):
            tk.Label(stats, text=text, bg=theme.PANEL, fg=color,
                     font=("Microsoft YaHei UI", 9, "bold"), padx=10, pady=3).pack(side=tk.LEFT)
        if plan.warnings:
            ttk.Label(top, text="⚠ " + "；".join(plan.warnings[:3]),
                      style="PanelMuted.TLabel").pack(anchor="w", pady=(4, 0))

        split = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        split.pack(fill=tk.BOTH, expand=True, padx=16, pady=6)
        left = ttk.Frame(split)
        right = ttk.Frame(split)
        split.add(left, weight=2)
        split.add(right, weight=3)

        cols = ("name", "status", "add", "del")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", selectmode="browse")
        for c, title, w in (("name", "文件", 220), ("status", "状态", 70),
                            ("add", "+", 50), ("del", "-", 50)):
            self.tree.heading(c, text=title)
            self.tree.column(c, width=w, anchor="w" if c == "name" else "center")
        vsb = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.tag_configure("error", foreground=theme.RED)
        self.tree.tag_configure("changed", foreground=theme.AMBER)
        self.tree.tag_configure("added", foreground=theme.GREEN)
        for i, f in enumerate(plan.files):
            if f.status == "error":
                add = rem = "-"
            else:
                add, rem = services.diff_stats(f.old_content, f.content)
                add = f"+{add}" if add else ""
                rem = f"-{rem}" if rem else ""
            self.tree.insert("", tk.END, iid=str(i),
                             values=(f.name, self.STATUS.get(f.status, f.status), add, rem),
                             tags=(f.status,))
        self.tree.bind("<<TreeviewSelect>>", lambda _e: self._show_diff())

        bar = ttk.Frame(right)
        bar.pack(fill=tk.X, pady=(0, 4))
        self.file_label = tk.StringVar(value="选择左侧文件查看 Diff")
        ttk.Label(bar, textvariable=self.file_label, style="PanelMuted.TLabel").pack(side=tk.LEFT)
        self.mask_btn = ttk.Button(bar, text="显示敏感值", style="Tool.TButton",
                                   command=self._toggle_mask)
        self.mask_btn.pack(side=tk.RIGHT)
        self.diff = DarkText(right)
        self.diff.pack(fill=tk.BOTH, expand=True)
        self._masked = True

        if plan.files:
            first = min(range(len(plan.files)), key=lambda i: plan.files[i].status == "unchanged")
            self.tree.selection_set(str(first))
            self.tree.focus(str(first))

        bottom = ttk.Frame(self, padding=(16, 6, 16, 14))
        bottom.pack(fill=tk.X)
        writable = sum(1 for f in plan.files if not f.error and f.status != "unchanged")
        self.write_btn = ttk.Button(
            bottom, text=f"写入 {writable} 个文件", style="Accent.TButton",
            command=self._write)
        self.write_btn.pack(side=tk.RIGHT)
        if not writable:
            self.write_btn.state(["disabled"])
        ttk.Button(bottom, text="返回修改", style="Tool.TButton",
                   command=self._cancel).pack(side=tk.RIGHT, padx=(0, 8))
        if plan.errors:
            ttk.Label(bottom, text=f"有 {len(plan.errors)} 个错误，错误文件不会写入",
                      style="PanelMuted.TLabel").pack(side=tk.LEFT)

    def _current(self) -> services.OutputFile | None:
        sel = self.tree.selection()
        if not sel:
            return None
        try:
            return self.plan.files[int(sel[0])]
        except (ValueError, IndexError):
            return None

    def _show_diff(self):
        f = self._current()
        if not f:
            return
        self.file_label.set(f"{f.name} · {self.STATUS.get(f.status, f.status)}")
        if f.error:
            self.diff.set(f"生成失败:\n{f.error}")
            return
        if f.status == "unchanged":
            self.diff.set("文件内容没有变化（不会重写）。")
            return
        diff = services.diff_text(f.old_content, f.content, f.name, self.project)
        if not diff:
            diff = f"新文件: {f.name}\n内容 {len(f.content.splitlines())} 行"
            if self._masked:
                diff = services.mask_secrets(diff, self.project)
        self.diff.set(diff, diff=True)

    def _toggle_mask(self):
        self._masked = not self._masked
        old = self.project.settings.mask_secrets
        self.project.settings.mask_secrets = self._masked
        self.mask_btn.configure(text="隐藏敏感值" if self._masked else "显示敏感值")
        self._show_diff()
        self.project.settings.mask_secrets = old

    def _write(self):
        self.result = True
        self.destroy()


class PreviewDialog(_Dialog):
    def __init__(self, parent, project: Project, device_label: str, content: str):
        super().__init__(parent, f"配置预览 — {device_label}", 860, 660)
        self.resizable(True, True)
        self.project = project
        self.content = content
        top = ttk.Frame(self, padding=(14, 12, 14, 4))
        top.pack(fill=tk.X)
        self.mask = tk.BooleanVar(value=project.settings.mask_secrets)
        ttk.Label(top, text=f"{device_label} · {len(content.splitlines())} 行",
                  style="PanelMuted.TLabel").pack(side=tk.LEFT)
        ttk.Checkbutton(top, text="隐藏敏感值", variable=self.mask,
                        style="Panel.TCheckbutton",
                        command=self._render).pack(side=tk.RIGHT)
        self.text = DarkText(self)
        self.text.pack(fill=tk.BOTH, expand=True, padx=14, pady=4)
        btns = ttk.Frame(self, padding=(14, 4, 14, 12))
        btns.pack(fill=tk.X)
        ttk.Button(btns, text="关闭", style="Tool.TButton",
                   command=self._cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="复制", style="Tool.TButton",
                   command=self.text.copy).pack(side=tk.RIGHT, padx=(0, 6))
        ttk.Button(btns, text="另存为…", style="Tool.TButton",
                   command=self._save).pack(side=tk.RIGHT, padx=(0, 6))
        self._render()

    def _render(self):
        text = services.mask_secrets(self.content, self.project) if self.mask.get() \
            else self.content
        self.text.set(text)

    def _save(self):
        path = filedialog.asksaveasfilename(
            parent=self, title="另存配置", defaultextension=".cfg",
            filetypes=[("配置文件", "*.cfg *.txt *.ini *.conf"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8", newline="") as f:
                f.write(self.content)
        except OSError as e:
            messagebox.showerror("保存失败", str(e), parent=self)


class CredentialDialog(_Dialog):
    def __init__(self, parent, key: str = "", label: str = "", value: str = "",
                 protected: bool = True, existing_keys: set[str] | None = None):
        super().__init__(parent, "凭据", 480, 330)
        self.existing = existing_keys or set()
        body = ttk.Frame(self, padding=16)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text="变量名（模板里写 {{变量名}}）").pack(anchor="w")
        self.key = tk.StringVar(value=key)
        ttk.Entry(body, textvariable=self.key).pack(fill=tk.X, pady=(2, 8))
        ttk.Label(body, text="说明（可选）").pack(anchor="w")
        self.label = tk.StringVar(value=label)
        ttk.Entry(body, textvariable=self.label).pack(fill=tk.X, pady=(2, 8))
        ttk.Label(body, text="值").pack(anchor="w")
        self.value = tk.StringVar(value=value)
        self.entry = ttk.Entry(body, textvariable=self.value, show="•")
        self.entry.pack(fill=tk.X, pady=(2, 4))
        self.show = tk.BooleanVar(value=False)
        ttk.Checkbutton(body, text="显示值", variable=self.show,
                        style="Panel.TCheckbutton",
                        command=self._toggle).pack(anchor="w")
        self.protected = tk.BooleanVar(value=protected)
        ttk.Checkbutton(body, text="使用 Windows DPAPI 加密保存", variable=self.protected,
                        style="Panel.TCheckbutton").pack(anchor="w", pady=(2, 8))
        btns = ttk.Frame(body)
        btns.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Button(btns, text="取消", style="Tool.TButton", command=self._cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="保存", style="Accent.TButton",
                   command=self._ok).pack(side=tk.RIGHT, padx=(0, 6))

    def _toggle(self):
        self.entry.configure(show="" if self.show.get() else "•")

    def _ok(self):
        key = self.key.get().strip()
        if not key:
            messagebox.showwarning("提示", "变量名不能为空", parent=self)
            return
        if not re.fullmatch(r"[A-Za-z_]\w*", key):
            messagebox.showwarning("提示", "变量名只能包含字母/数字/下划线", parent=self)
            return
        self.result = {"key": key, "label": self.label.get().strip(),
                       "value": self.value.get(), "protected": self.protected.get()}
        self.destroy()


class VariableDialog(_Dialog):
    def __init__(self, parent, key: str = "", value: str = ""):
        super().__init__(parent, "项目变量", 460, 260)
        body = ttk.Frame(self, padding=16)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text="变量名（模板里写 {{变量名}}）").pack(anchor="w")
        self.key = tk.StringVar(value=key)
        ttk.Entry(body, textvariable=self.key).pack(fill=tk.X, pady=(2, 8))
        ttk.Label(body, text="值").pack(anchor="w")
        self.value = tk.StringVar(value=value)
        ttk.Entry(body, textvariable=self.value).pack(fill=tk.X, pady=(2, 12))
        btns = ttk.Frame(body)
        btns.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Button(btns, text="取消", style="Tool.TButton", command=self._cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="保存", style="Accent.TButton",
                   command=self._ok).pack(side=tk.RIGHT, padx=(0, 6))

    def _ok(self):
        key = self.key.get().strip()
        if not re.fullmatch(r"[A-Za-z_]\w*", key):
            messagebox.showwarning("提示", "变量名只能包含字母/数字/下划线", parent=self)
            return
        self.result = {"key": key, "value": self.value.get()}
        self.destroy()


class DeviceDialog(_Dialog):
    """编辑一台设备的全部标准字段 + 任意扩展字段。"""

    LABELS = {"esn": "ESN 序列号", "hostname": "主机名", "type": "类型",
              "mgmt_ip": "管理IP", "mgmt_vlan": "管理VLAN", "access_vlan": "接入VLAN",
              "mask": "掩码", "gateway": "网关", "note": "备注"}

    def __init__(self, parent, dev: dict, title: str = "设备详情"):
        super().__init__(parent, title, 720, 640)
        self.resizable(True, True)
        self.dev = dev
        body = ttk.Frame(self, padding=16)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text="标准字段",
                  style="Panel.TLabel", font=theme.FONT_H).pack(anchor="w")
        form = ttk.Frame(body)
        form.pack(fill=tk.X, pady=(6, 10))
        self.vars = {}
        items = list(self.LABELS.items())
        for i, (key, label) in enumerate(items):
            r, c = divmod(i, 3)
            ttk.Label(form, text=label).grid(row=r * 2, column=c, sticky="w",
                                             padx=(0, 10), pady=(4, 2))
            var = tk.StringVar(value=str(dev.get(key, "")))
            ttk.Entry(form, textvariable=var).grid(row=r * 2 + 1, column=c, sticky="ew",
                                                   padx=(0, 10), pady=(0, 4))
            self.vars[key] = var
        for c in range(3):
            form.columnconfigure(c, weight=1)
        ttk.Label(body, text="扩展字段 / 扩展变量（每行一个 K:V；模板可用 {{K}} 引用）",
                  style="Panel.TLabel", font=theme.FONT_H).pack(anchor="w")
        ttk.Label(body, text="例如 ntp_server:10.0.0.1、stack_member_id:1、vrp_file:SW.cc",
                  style="PanelMuted.TLabel").pack(anchor="w", pady=(0, 4))
        self.extras = tk.Text(body, height=8, font=theme.FONT_MONO, wrap=tk.NONE,
                              relief=tk.FLAT, borderwidth=1, highlightthickness=1,
                              highlightbackground=theme.LINE)
        self.extras.pack(fill=tk.BOTH, expand=True)
        self.extras.insert("1.0", nadt.format_extras_text(dev.get("_extras") or {}))
        btns = ttk.Frame(body)
        btns.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(btns, text="取消", style="Tool.TButton", command=self._cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="保存", style="Accent.TButton",
                   command=self._ok).pack(side=tk.RIGHT, padx=(0, 6))

    def _ok(self):
        standard = {k: v.get().strip() for k, v in self.vars.items()}
        extras = nadt.parse_extras_text(self.extras.get("1.0", tk.END))
        self.result = {"standard": standard, "extras": extras}
        self.destroy()
