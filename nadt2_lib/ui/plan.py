"""方案页：模板 + 参数表 + 校验 + 生成预检/Diff + 批量操作。"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import nadt
from .. import services
from ..model import Project
from . import theme
from .dialogs import (ColumnMapDialog, DeviceDialog, FillDialog, FindReplaceDialog,
                      GenerationDialog, PreviewDialog, SequenceDialog)
from .widgets import DarkText, LogPanel, PageHeader, Tooltip, make_tree

RULE_LABELS = {
    "hostname": "主机名",
    "esn": "ESN",
    "mgmt_ip": "管理IP",
    "__single__": "单文件（合并）",
}
SCOPE_LABELS = {"all": "全部设备", "type": "仅当前模板类型"}
SCOPE_KEYS = {v: k for k, v in SCOPE_LABELS.items()}
EOL_LABELS = {"crlf": "CRLF (Windows)", "lf": "LF (Unix)"}
EOL_KEYS = {v: k for k, v in EOL_LABELS.items()}


class PlanPage(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.project: Project = app.project
        self._loading = False
        self._editor = None
        self._editor_ctx = None
        self._cursor = (0, 0)
        self._columns: list[str] = []
        self._loaded_text = ""
        self._issues: list[services.Issue] = []
        self._plan: services.GenerationPlan | None = None
        self._batch_menu = tk.Menu(self, tearoff=0)

        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.header = PageHeader(
            self, "方案工作台",
            "选模板 → 填数据 → 校验 → 生成预检（Diff）→ 写盘",
            "Ctrl+S 保存  ·  F6 校验  ·  F5 生成  ·  Ctrl+E 设备详情")
        self.header.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 0))
        self._build_command_bar()
        self._build_main()
        self._build_bottom()
        self.after(180, self._set_sash)

    # ── 顶部 ──
    def _build_command_bar(self):
        card = ttk.LabelFrame(self, text="方案设置", padding=(14, 8, 14, 9))
        card.grid(row=1, column=0, sticky="ew", padx=10, pady=(8, 6))
        row1 = ttk.Frame(card)
        row1.pack(fill=tk.X)
        gen = ttk.Button(row1, text="生成…", style="Accent.TButton",
                         command=self.generate)
        gen.pack(side=tk.RIGHT)
        Tooltip(gen, "生成预检：先看 Diff，再写入  F5")
        val = ttk.Button(row1, text="校验 F6", style="Tool.TButton",
                         command=self.validate_and_show)
        val.pack(side=tk.RIGHT, padx=(0, 4))
        Tooltip(val, "校验项目：重复/缺变量/文件名冲突  F6")
        prev = ttk.Button(row1, text="预览", style="Tool.TButton",
                          command=self.preview_selected)
        prev.pack(side=tk.RIGHT, padx=(0, 8))
        Tooltip(prev, "预览选中设备的最终配置  Ctrl+Enter")
        ttk.Label(row1, text="模板:").pack(side=tk.LEFT)
        self.var_template = tk.StringVar()
        self.cmb_template = ttk.Combobox(row1, textvariable=self.var_template,
                                         state="readonly", width=20)
        self.cmb_template.pack(side=tk.LEFT, padx=(4, 4))
        self.cmb_template.bind("<<ComboboxSelected>>", lambda _e: self._on_template_change())
        save_tpl = ttk.Button(row1, text="保存模板", style="Tool.TButton",
                              command=self.save_template)
        save_tpl.pack(side=tk.LEFT, padx=(0, 2))
        Tooltip(save_tpl, "把右侧模板文本保存到项目")

        tk.Frame(row1, bg=theme.LINE, width=1, height=22).pack(
            side=tk.LEFT, padx=(9, 10), pady=2)
        ttk.Label(row1, text="文件名:").pack(side=tk.LEFT)
        self.var_rule = tk.StringVar()
        self.cmb_rule = ttk.Combobox(row1, textvariable=self.var_rule,
                                     state="readonly", width=13)
        self.cmb_rule.pack(side=tk.LEFT, padx=(4, 8))
        self.cmb_rule.bind("<<ComboboxSelected>>", lambda _e: self._on_setting_changed())
        ttk.Label(row1, text="后缀:").pack(side=tk.LEFT)
        self.var_ext = tk.StringVar()
        cmb_ext = ttk.Combobox(row1, textvariable=self.var_ext, width=6,
                               values=("cfg", "txt", "ini", "conf", "log", "json", "csv"))
        cmb_ext.pack(side=tk.LEFT, padx=(4, 2))
        Tooltip(cmb_ext, "输出文件后缀，例如 cfg / txt / ini")

        tk.Frame(row1, bg=theme.LINE, width=1, height=22).pack(
            side=tk.LEFT, padx=(9, 10), pady=2)
        ttk.Label(row1, text="换行:").pack(side=tk.LEFT)
        self.var_eol = tk.StringVar()
        cmb_eol = ttk.Combobox(row1, textvariable=self.var_eol, state="readonly",
                               width=14, values=list(EOL_LABELS.values()))
        cmb_eol.pack(side=tk.LEFT, padx=(4, 8))
        cmb_eol.bind("<<ComboboxSelected>>", lambda _e: self._on_setting_changed())
        Tooltip(cmb_eol, "CRLF 与原宏脚本一致；Linux 设备可选 LF")

        row2 = ttk.Frame(card)
        row2.pack(fill=tk.X, pady=(9, 0))
        ttk.Label(row2, text="输出目录:").pack(side=tk.LEFT)
        self.var_output = tk.StringVar()
        self.ent_output = ttk.Entry(row2, textvariable=self.var_output, width=20)
        self.ent_output.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 4))
        self.ent_output.bind("<FocusOut>", lambda _e: self._on_setting_changed())
        self.ent_output.bind("<Return>", lambda _e: self._on_setting_changed())
        browse = ttk.Button(row2, text="浏览", style="Tool.TButton",
                            command=self._browse_output)
        browse.pack(side=tk.LEFT, padx=(0, 2))
        Tooltip(browse, "选择输出目录")

        tk.Frame(row2, bg=theme.LINE, width=1, height=22).pack(
            side=tk.LEFT, padx=(9, 10), pady=2)
        self.var_scope = tk.StringVar()
        ttk.Label(row2, text="范围:").pack(side=tk.LEFT)
        cmb_scope = ttk.Combobox(row2, textvariable=self.var_scope, state="readonly",
                                 width=13, values=list(SCOPE_LABELS.values()))
        cmb_scope.pack(side=tk.LEFT, padx=(4, 12))
        cmb_scope.bind("<<ComboboxSelected>>", lambda _e: self._on_scope_change())
        Tooltip(cmb_scope, "多模板项目建议选「仅当前模板类型」")
        self.var_subdir = tk.BooleanVar()
        self.var_archive = tk.BooleanVar()
        self.var_lswnet = tk.BooleanVar()
        checks = [("模板子目录", self.var_subdir, "输出到 <模板名> 子目录"),
                  ("时间戳归档", self.var_archive, "每次生成到时间戳子目录，保留历史"),
                  ("生成 lswnet.cfg", self.var_lswnet, "自动生成华为 ESN 映射表")]
        for text, var, tip in checks:
            cb = ttk.Checkbutton(row2, text=text, variable=var, style="TCheckbutton",
                                 command=self._on_setting_changed)
            cb.pack(side=tk.LEFT, padx=(0, 10))
            Tooltip(cb, tip)

    def _build_main(self):
        pane = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        pane.grid(row=2, column=0, sticky="nsew", padx=10, pady=(0, 4))
        self.pane = pane

        left = ttk.LabelFrame(pane, text="设备参数", padding=(8, 4, 8, 8))
        pane.add(left, weight=3)
        bar = ttk.Frame(left)
        bar.pack(fill=tk.X, pady=(0, 4))
        self.var_count = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.var_count, style="Muted.TLabel").pack(side=tk.LEFT)
        self.var_search = tk.StringVar()
        self.var_search.trace_add("write", lambda *_a: self.refresh_table())
        ttk.Label(bar, text="搜索", style="Muted.TLabel").pack(side=tk.LEFT, padx=(12, 3))
        self.search_entry = ttk.Entry(bar, textvariable=self.var_search, width=13)
        self.search_entry.pack(side=tk.LEFT)
        Tooltip(self.search_entry, "搜索主机名 / ESN / 管理IP")

        self.device_menu = tk.Menu(self, tearoff=0)
        self.device_menu.add_command(label="设备详情（全部字段）…  Ctrl+E",
                                     command=self.edit_device_details)
        self.device_menu.add_command(label="编辑单元格…  F2", command=self._edit_cursor)
        self.device_menu.add_command(label="预览此设备…  Ctrl+Enter",
                                     command=self.preview_selected)
        self.device_menu.add_separator()
        self.device_menu.add_command(label="复制选中行", command=self.duplicate_selected)
        self.device_menu.add_command(label="删除选中行", command=self.delete_devices)

        self.import_menu = tk.Menu(self, tearoff=0)
        self.import_menu.add_command(label="从 CSV / Excel 导入（列映射）…",
                                     command=self.import_table)
        self.import_menu.add_command(label="导入旧 xlsm 工作簿…",
                                     command=self.app.import_legacy)
        self.import_menu.add_separator()
        self.import_menu.add_command(label="导出 xlsm 兼容工作簿…",
                                     command=self.app.export_workbook)

        ops_btn = ttk.Button(bar, text="操作 ▾", style="Bar.TButton",
                             command=self.show_device_menu)
        ops_btn.pack(side=tk.RIGHT, padx=(0, 0))
        Tooltip(ops_btn, "设备详情、复制、删除、预览")
        add_btn = ttk.Button(bar, text="添加", style="Bar.TButton",
                             command=self.add_device)
        add_btn.pack(side=tk.RIGHT, padx=(0, 4))
        Tooltip(add_btn, "添加一台设备（光标停在主机名）")
        tk.Frame(bar, bg=theme.LINE, width=1, height=22).pack(
            side=tk.RIGHT, padx=7, pady=2)
        batch_btn = ttk.Button(bar, text="批量 ▾", style="Bar.TButton",
                               command=self.show_batch_menu)
        batch_btn.pack(side=tk.RIGHT, padx=(4, 0))
        Tooltip(batch_btn, "填充、序列、查找替换、复制行")
        import_btn = ttk.Button(bar, text="导入 ▾", style="Bar.TButton",
                                command=self.show_import_menu)
        import_btn.pack(side=tk.RIGHT, padx=(4, 0))
        Tooltip(import_btn, "从表格/旧工作簿导入，或导出 xlsm")

        self.tree_frame, self.tree = make_tree(left, ["hostname"], height=12)
        self.tree.tag_configure("even", background="white")
        self.tree.tag_configure("odd", background="#f8fafc")
        self.tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self._on_dclick)
        self.tree.bind("<Button-1>", self._on_click)
        self.tree.bind("<Button-3>", self._on_rmb)
        self.tree.bind("<Control-c>", self._copy)
        self.tree.bind("<Control-C>", self._copy)
        self.tree.bind("<Control-v>", self._paste)
        self.tree.bind("<Control-V>", self._paste)
        self.tree.bind("<Delete>", self._clear_cell)
        self.tree.bind("<F2>", self._edit_cursor)
        self.tree.bind("<MouseWheel>", lambda _e: self._commit_edit(), add="+")

        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="编辑单元格…", command=self._edit_cursor)
        self.menu.add_command(label="设备详情（全部字段）…", command=self.edit_device_details)
        self.menu.add_command(label="预览此设备…", command=self.preview_selected)
        self.menu.add_separator()
        self.menu.add_command(label="在下方插入行", command=self.insert_row_below)
        self.menu.add_command(label="复制选中行", command=self.duplicate_selected)
        self.menu.add_command(label="删除选中行", command=self.delete_devices)
        self.menu.add_separator()
        self.menu.add_command(label="复制", command=self._copy)
        self.menu.add_command(label="粘贴", command=self._paste)

        right = ttk.LabelFrame(pane, text="配置模板", padding=(8, 4, 8, 8))
        pane.add(right, weight=2)
        top = ttk.Frame(right)
        top.pack(fill=tk.X, pady=(0, 3))
        self.var_tpl_hint = tk.StringVar(value="")
        ttk.Label(top, textvariable=self.var_tpl_hint, style="Muted.TLabel").pack(side=tk.LEFT)
        refresh_btn = ttk.Button(top, text="刷新参数列", style="Tool.TButton",
                                 command=self.refresh_columns)
        refresh_btn.pack(side=tk.RIGHT, padx=(4, 0))
        Tooltip(refresh_btn, "模板占位符改了之后，重新生成左侧参数列")
        save_btn = ttk.Button(top, text="保存", style="Tool.TButton",
                              command=self.save_template)
        save_btn.pack(side=tk.RIGHT)
        Tooltip(save_btn, "保存模板文本到项目  Ctrl+S 保存整个项目")
        wrap = ttk.Frame(right)
        wrap.pack(fill=tk.BOTH, expand=True)
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        self.text = tk.Text(wrap, font=theme.FONT_MONO, undo=True, wrap=tk.NONE,
                            relief=tk.FLAT, borderwidth=1, highlightthickness=1,
                            highlightbackground=theme.LINE, bg="white", fg=theme.FG)
        vsb = ttk.Scrollbar(wrap, orient=tk.VERTICAL, command=self.text.yview)
        hsb = ttk.Scrollbar(wrap, orient=tk.HORIZONTAL, command=self.text.xview)
        self.text.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.text.bind("<<Modified>>", self._on_text_modified)
        self.var_tpl_hint2 = tk.StringVar(value="")
        ttk.Label(right, textvariable=self.var_tpl_hint2, style="Muted.TLabel").pack(
            anchor="w", pady=(4, 0))

    def _build_bottom(self):
        holder = ttk.Frame(self, height=220)
        holder.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 8))
        holder.grid_propagate(False)
        holder.pack_propagate(False)
        self.bottom = ttk.Notebook(holder)
        self.bottom.pack(fill=tk.BOTH, expand=True)
        # 问题
        issues = ttk.Frame(self.bottom)
        self.bottom.add(issues, text="问题")
        top = ttk.Frame(issues)
        top.pack(fill=tk.X, padx=8, pady=(6, 3))
        self.var_issue = tk.StringVar(value="尚未校验")
        ttk.Label(top, textvariable=self.var_issue, style="Muted.TLabel").pack(side=tk.LEFT)
        ttk.Button(top, text="校验", style="Tool.TButton",
                   command=self.validate_and_show).pack(side=tk.RIGHT)
        self.issue_frame, self.issue_tree = make_tree(
            issues, ["sev", "device", "field", "message"], height=5,
            headings={"sev": "级别", "device": "设备", "field": "字段", "message": "说明"},
            widths={"sev": 60, "device": 100, "field": 100, "message": 700},
            selectmode="browse")
        self.issue_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 6))
        self.issue_tree.tag_configure("error", foreground=theme.RED)
        self.issue_tree.tag_configure("warn", foreground=theme.AMBER)
        self.issue_tree.tag_configure("info", foreground=theme.MUTED)
        self.issue_tree.bind("<Double-1>", lambda _e: self.focus_issue())

        # 变更
        changes = ttk.Frame(self.bottom)
        self.bottom.add(changes, text="变更 / Diff")
        pane = ttk.Panedwindow(changes, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        top_frame = ttk.Frame(pane)
        bottom_frame = ttk.Frame(pane)
        pane.add(top_frame, weight=2)
        pane.add(bottom_frame, weight=3)
        self.change_tree = ttk.Treeview(top_frame, columns=("name", "status", "add", "del"),
                                        show="headings", height=6, selectmode="browse")
        for c, title, w in (("name", "文件", 200), ("status", "状态", 60),
                            ("add", "+", 45), ("del", "-", 45)):
            self.change_tree.heading(c, text=title)
            self.change_tree.column(c, width=w, anchor="w" if c == "name" else "center")
        vsb = ttk.Scrollbar(top_frame, orient=tk.VERTICAL, command=self.change_tree.yview)
        self.change_tree.configure(yscrollcommand=vsb.set)
        self.change_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.change_tree.tag_configure("added", foreground=theme.GREEN)
        self.change_tree.tag_configure("changed", foreground=theme.AMBER)
        self.change_tree.tag_configure("error", foreground=theme.RED)
        self.change_tree.bind("<<TreeviewSelect>>", lambda _e: self._show_change_diff())
        self.change_diff = DarkText(bottom_frame)
        self.change_diff.pack(fill=tk.BOTH, expand=True)
        self.change_diff.set("生成后在这里查看每个文件的 Diff。")

        # 日志
        logf = ttk.Frame(self.bottom)
        self.bottom.add(logf, text="日志")
        self.log = LogPanel(logf, title="方案日志", height=5, hint="生成 / 导入 / 批量操作")
        self.log.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

    # ── 项目加载 ──
    def reload_project(self):
        self.project = self.app.project
        self._loading = True
        try:
            names = list(self.project.templates)
            self.cmb_template["values"] = names
            name = self.app.current_template
            if name not in names:
                name = names[0] if names else ""
                self.app.current_template = name
            self.var_template.set(name)
            self._load_template_text(name)
            self._apply_settings_to_ui()
            self._issues = []
            self.issue_tree.delete(*self.issue_tree.get_children())
            self.var_issue.set("尚未校验")
            self.app.set_issue_count(0, 0)
            self.refresh_columns(refresh_table=False)
            self.refresh_table()
            self._show_plan(None)
            self._loading = False
            if len(self.project.devices) <= 600:
                self.after(80, lambda: self.validate(silent=True))
            else:
                self.var_issue.set("设备较多：按 F6 校验")
        finally:
            self._loading = False
        self.app.refresh_header()

    def on_show(self):
        self.project = self.app.project
        self.refresh_columns()
        if not self._issues:
            self.validate(silent=True)

    def select_template(self, name: str):
        if name in self.project.templates:
            self.var_template.set(name)
            self._on_template_change(force=True)

    def _apply_settings_to_ui(self):
        s = self.project.settings
        self._set_rule_options()
        self.var_ext.set(s.extension)
        self.var_output.set(s.output_dir)
        self.var_scope.set(SCOPE_LABELS.get(s.scope, SCOPE_LABELS["all"]))
        self.var_subdir.set(bool(s.template_subdir))
        self.var_archive.set(bool(s.timestamp_archive))
        self.var_lswnet.set(bool(s.generate_lswnet))
        self.var_eol.set(EOL_LABELS.get(s.eol, EOL_LABELS["crlf"]))

    def _set_rule_options(self):
        vars_ = self.project.template_vars(self.text.get("1.0", "end-1c")) \
            if hasattr(self, "text") else []
        reserved = set(self.project.variables) | set(self.project.credentials)
        keys = ["hostname", "esn", "mgmt_ip"]
        for v in vars_:
            if v not in reserved and v not in ("hostname", "sysname", "type", "esn", "note") \
                    and v not in keys:
                keys.append(v)
        keys.append("__single__")
        self._rule_keys = keys
        labels = [RULE_LABELS.get(k, k) for k in keys]
        self._rule_key_by_label = dict(zip(labels, keys))
        self.cmb_rule["values"] = labels
        current = self.project.settings.filename_rule
        self.var_rule.set(RULE_LABELS.get(current, current))

    def sync_settings(self) -> bool:
        if self._loading:
            return False
        s = self.project.settings
        old = (s.filename_rule, s.extension, s.output_dir, s.scope, s.template_subdir,
               s.timestamp_archive, s.generate_lswnet, s.eol, s.default_template)
        rule = self._rule_key_by_label.get(self.var_rule.get(), self.var_rule.get())
        s.filename_rule = rule or "hostname"
        s.extension = self.var_ext.get().strip().lstrip(".") or "cfg"
        s.output_dir = self.var_output.get().strip() or "output"
        s.scope = SCOPE_KEYS.get(self.var_scope.get(), "all")
        s.template_subdir = bool(self.var_subdir.get())
        s.timestamp_archive = bool(self.var_archive.get())
        s.generate_lswnet = bool(self.var_lswnet.get())
        s.eol = EOL_KEYS.get(self.var_eol.get(), "crlf")
        s.default_template = self.var_template.get() or s.default_template
        new = (s.filename_rule, s.extension, s.output_dir, s.scope, s.template_subdir,
               s.timestamp_archive, s.generate_lswnet, s.eol, s.default_template)
        if new != old:
            self.project.mark_dirty()
            self.app.refresh_header()
            return True
        return False

    def _on_setting_changed(self):
        if self._loading:
            return
        if self.sync_settings():
            self.app.set_status("方案设置已修改")
            if self.bottom.index(self.bottom.select()) == 0:
                self.validate(silent=True)

    def _on_scope_change(self):
        self._on_setting_changed()
        self.refresh_table()

    # ── 模板 ──
    def _load_template_text(self, name: str):
        tpl = self.project.templates.get(name)
        text = tpl.text if tpl else ""
        self.text.delete("1.0", tk.END)
        self.text.insert("1.0", text)
        self.text.edit_reset()
        self.text.edit_modified(False)
        self._loaded_text = text
        self.var_tpl_hint.set(f"{name} · {len(text.splitlines())} 行")

    def _text_dirty(self) -> bool:
        return self.text.get("1.0", "end-1c") != self._loaded_text

    def _on_text_modified(self, _ev=None):
        if not self.text.edit_modified():
            return
        self.text.edit_modified(False)
        if self._text_dirty():
            self.var_tpl_hint.set(f"{self.var_template.get()} · 未保存")

    def _on_template_change(self, force: bool = False):
        if self._loading:
            return
        if not force and self._text_dirty():
            ans = messagebox.askyesnocancel("模板未保存", "当前模板有修改，是否保存到项目？")
            if ans is None:
                self.var_template.set(self.app.current_template)
                return
            if ans:
                self.save_template(silent=True)
        name = self.var_template.get()
        self.app.current_template = name
        self.project.settings.default_template = name
        self.project.mark_dirty()
        self._load_template_text(name)
        self._apply_settings_to_ui()
        self.refresh_columns()
        self.app.refresh_header()
        self.app.set_status(f"当前模板: {name}")
        self.validate(silent=True)

    def save_template(self, silent: bool = False):
        name = self.var_template.get()
        if name not in self.project.templates:
            return
        text = self.text.get("1.0", "end-1c")
        if text == self.project.templates[name].text:
            if not silent:
                self.app.set_status("模板没有变化")
            return
        self.project.templates[name].text = text
        self._loaded_text = text
        self.project.mark_dirty()
        self.var_tpl_hint.set(f"{name} · {len(text.splitlines())} 行")
        self.refresh_columns()
        self.app.refresh_header()
        self.log.log(f"✓ 模板已保存到项目: {name}")
        if not silent:
            self.app.set_status(f"模板已保存: {name}")

    def refresh_columns(self, refresh_table: bool = True):
        text = self.text.get("1.0", "end-1c") if hasattr(self, "text") else ""
        params = self.project.template_vars(text)
        cols = self.project.columns(self.var_template.get(), text=text)
        self._columns = cols
        self.tree.configure(columns=cols)
        for c in cols:
            self.tree.heading(c, text=Project.header(c))
            self.tree.column(c, width=self._col_width(c), minwidth=55,
                             anchor="center" if c in ("type", "mgmt_vlan", "access_vlan") else "w",
                             stretch=False)
        self.tree.xview_moveto(0)
        self._set_rule_options()
        if refresh_table:
            self.refresh_table()
        reserved = set(self.project.variables) | set(self.project.credentials)
        provided = [p for p in params if p in reserved]
        missing = [p for p in params if p not in reserved and p not in
                   ("hostname", "sysname", "vlan_batch", "mgmt_vlan", "mgmt_ip",
                    "mask", "access_vlan", "gateway")]
        bits = [f"{len(params)} 个占位符 · {len(cols)} 列"]
        if provided:
            bits.append("项目提供: " + ", ".join(provided[:5]))
        if missing:
            bits.append("需在表格/项目变量/凭据提供: " + ", ".join(missing[:6]))
        self.var_tpl_hint2.set(" · ".join(bits))

    def _col_width(self, col: str) -> int:
        def dl(s):
            return sum(2 if ord(ch) > 127 else 1 for ch in str(s))
        maxlen = dl(Project.header(col)) + 2
        for dev in self.project.devices[:300]:
            maxlen = max(maxlen, dl(Project.cell_get(dev, col)))
        if col == "hostname":
            return min(max(maxlen * 7, 130), 200)
        if col == "type":
            return 60
        if col == "esn":
            return min(max(maxlen * 8, 145), 210)
        if col in ("mgmt_vlan", "access_vlan"):
            return 82
        if col == "mask":
            return 108
        if col in ("mgmt_ip", "gateway"):
            return 95
        if col == "note":
            return 80
        return min(max(maxlen * 7, 90), 105)

    # ── 表格 ──
    def _visible_indices(self) -> list[int]:
        p = self.project
        if p.settings.scope == "type":
            tpl = p.templates.get(self.var_template.get())
            stems = {tpl.stem.lower() if tpl else "", tpl.type_name.lower() if tpl else ""}
            base = [i for i, d in enumerate(p.devices)
                    if (d.get("type") or "").strip().lower() in ("", *stems)]
        else:
            base = list(range(len(p.devices)))
        kw = self.var_search.get().strip().lower()
        if not kw:
            return base
        out = []
        for i in base:
            d = p.devices[i]
            parts = [d.get("hostname", ""), d.get("esn", ""), d.get("type", ""),
                     d.get("mgmt_ip", ""), d.get("note", "")]
            parts.extend(str(v) for v in (d.get("_extras") or {}).values())
            if kw in " ".join(parts).lower():
                out.append(i)
        return out

    def refresh_table(self):
        if not self._columns:
            return
        self.tree.delete(*self.tree.get_children())
        visible = self._visible_indices()
        for n, i in enumerate(visible):
            dev = self.project.devices[i]
            vals = [Project.cell_get(dev, c) for c in self._columns]
            self.tree.insert("", tk.END, iid=str(i), values=vals,
                             tags=("odd" if n % 2 else "even",))
            if n and n % 300 == 0:
                self.update_idletasks()
        total = len(self.project.devices)
        target = len(self.project.target_devices())
        text = f"设备 {total} 台 · 生成 {target}"
        if len(visible) != target:
            text += f" · 显示 {len(visible)}"
        self.var_count.set(text)
        self._refresh_batch_menu()
        self.after_idle(lambda: self.tree.xview_moveto(0))

    def _refresh_batch_menu(self):
        self._batch_menu.delete(0, tk.END)
        self._batch_menu.add_command(label="填充值…", command=self.batch_fill)
        self._batch_menu.add_command(label="数字序列…", command=self.batch_sequence)
        self._batch_menu.add_command(label="复制选中行", command=self.duplicate_selected)
        self._batch_menu.add_separator()
        self._batch_menu.add_command(label="查找替换…", command=self.find_replace_dialog)
        self._batch_menu.add_command(label="导入表格（列映射）…", command=self.import_table)
        self._batch_menu.add_separator()
        self._batch_menu.add_command(label="删除选中行", command=self.delete_devices)

    def show_batch_menu(self):
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery()
            self._batch_menu.tk_popup(x, y)
        finally:
            self._batch_menu.grab_release()

    def show_import_menu(self):
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery()
            self.import_menu.tk_popup(x, y)
        finally:
            self.import_menu.grab_release()

    def show_device_menu(self):
        try:
            x = self.winfo_pointerx()
            y = self.winfo_pointery()
            self.device_menu.tk_popup(x, y)
        finally:
            self.device_menu.grab_release()

    def focus_search(self):
        self.search_entry.focus_set()
        self.search_entry.select_range(0, tk.END)

    # ── 单元格编辑 ──
    def _pos(self, ev):
        iid = self.tree.identify_row(ev.y)
        col = self.tree.identify_column(ev.x)
        if not iid or not col:
            return None
        try:
            return iid, int(col[1:]) - 1
        except ValueError:
            return None

    def _on_click(self, ev):
        pos = self._pos(ev)
        if pos:
            self._cursor = (int(pos[0]), pos[1])
        return None

    def _on_dclick(self, ev):
        pos = self._pos(ev)
        if not pos:
            return "break"
        self._cursor = (int(pos[0]), pos[1])
        self._begin_edit(pos[0], pos[1])
        return "break"

    def _on_rmb(self, ev):
        pos = self._pos(ev)
        if pos:
            self._cursor = (int(pos[0]), pos[1])
            self.tree.selection_set(pos[0])
        try:
            self.menu.tk_popup(ev.x_root, ev.y_root)
        finally:
            self.menu.grab_release()

    def _edit_cursor(self, _ev=None):
        if not self._columns:
            return "break"
        sel = self.tree.selection()
        iid = sel[0] if sel else str(self._cursor[0])
        self._begin_edit(iid, min(self._cursor[1], len(self._columns) - 1))
        return "break"

    def _begin_edit(self, iid: str, c: int):
        if not self.tree.exists(iid) or c < 0 or c >= len(self._columns):
            return
        self._commit_edit()
        self.tree.see(iid)
        self.update_idletasks()
        bbox = self.tree.bbox(iid, f"#{c + 1}")
        if not bbox:
            return
        x, y, w, h = bbox
        idx = int(iid)
        val = Project.cell_get(self.project.devices[idx], self._columns[c])
        ent = tk.Entry(self.tree, font=theme.FONT, relief=tk.SOLID, borderwidth=1,
                       highlightthickness=1, highlightbackground=theme.ACCENT,
                       bg="white", fg=theme.FG)
        ent.insert(0, val)
        ent.select_range(0, tk.END)
        ent.place(x=x, y=y + 1, width=max(w - 1, 30), height=max(h - 2, 20))
        ent.focus_set()
        self._editor = ent
        self._editor_ctx = (iid, c)
        ent.bind("<Return>", lambda _e: self._edit_return())
        ent.bind("<Tab>", lambda _e: self._edit_tab())
        ent.bind("<Escape>", lambda _e: self._cancel_edit())
        ent.bind("<FocusOut>", lambda _e: self._commit_edit())

    def _edit_return(self):
        self._commit_edit(next_row=True)
        return "break"

    def _edit_tab(self):
        self._commit_edit(next_col=True)
        return "break"

    def _cancel_edit(self):
        ent = self._editor
        self._editor = None
        self._editor_ctx = None
        if ent:
            ent.destroy()
        return "break"

    def _commit_edit(self, next_row=False, next_col=False):
        ent = self._editor
        if ent is None:
            return
        ctx = self._editor_ctx
        self._editor = None
        self._editor_ctx = None
        try:
            val = ent.get()
        finally:
            ent.destroy()
        if not ctx:
            return
        iid, c = ctx
        idx = int(iid)
        if idx >= len(self.project.devices) or c >= len(self._columns):
            return
        col = self._columns[c]
        old = Project.cell_get(self.project.devices[idx], col)
        if val != old:
            self.app.push_undo("编辑单元格")
            Project.cell_set(self.project.devices[idx], col, val)
            self.project.mark_dirty()
            self.update_row(idx)
            self.app.refresh_header()
            self.app.set_status(f"已修改 {col} = {val[:40]}")
        n = len(self.project.devices)
        if next_col:
            nc, ni = c + 1, idx
            if nc >= len(self._columns):
                nc, ni = 0, idx + 1
            if ni < n:
                self.after(10, lambda: self._begin_edit(str(ni), nc))
        elif next_row:
            ni = idx + 1
            if ni < n:
                self.after(10, lambda: self._begin_edit(str(ni), c))

    def update_row(self, idx: int):
        iid = str(idx)
        if self.tree.exists(iid):
            vals = [Project.cell_get(self.project.devices[idx], c) for c in self._columns]
            self.tree.item(iid, values=vals)

    def _copy(self, _ev=None):
        sel = list(self.tree.selection())
        if not sel or not self._columns:
            return "break"
        if len(sel) > 1:
            text = "\n".join(
                "\t".join(Project.cell_get(self.project.devices[int(i)], c)
                          for c in self._columns) for i in sel)
        else:
            r, c = self._cursor
            if not (0 <= r < len(self.project.devices) and 0 <= c < len(self._columns)):
                return "break"
            text = Project.cell_get(self.project.devices[r], self._columns[c])
        self.clipboard_clear()
        self.clipboard_append(text)
        return "break"

    def _paste(self, _ev=None):
        self._commit_edit()
        try:
            data = self.clipboard_get()
        except tk.TclError:
            return "break"
        if not data or not self._columns:
            return "break"
        rows = [line.split("\t") for line in
                data.replace("\r\n", "\n").replace("\r", "\n").rstrip("\n").split("\n")]
        r0, c0 = self._cursor
        r0 = max(0, min(r0, len(self.project.devices)))
        if rows and all(c0 + i < len(self._columns) and
                        rows[0][i].strip() == Project.header(self._columns[c0 + i]).strip()
                        for i in range(min(len(rows[0]), len(self._columns) - c0))):
            rows = rows[1:]
        if not rows:
            return "break"
        self.app.push_undo("粘贴表格")
        added = 0
        for dr, row in enumerate(rows):
            r = r0 + dr
            while r >= len(self.project.devices):
                dev = nadt.derive_defaults({"type": self._template_stem()})
                dev["_extras"] = {"sysname": ""}
                self.project.devices.append(dev)
                added += 1
            for dc, val in enumerate(row):
                c = c0 + dc
                if c >= len(self._columns):
                    break
                Project.cell_set(self.project.devices[r], self._columns[c], val)
        self.project.mark_dirty()
        self.refresh_table()
        self.app.refresh_header()
        self.log.log(f"✓ 粘贴 {len(rows)} 行" + (f"（新增 {added} 台）" if added else ""))
        self.app.set_status(f"已粘贴 {len(rows)} 行")
        return "break"

    def _clear_cell(self, _ev=None):
        r, c = self._cursor
        if not (0 <= r < len(self.project.devices) and 0 <= c < len(self._columns)):
            return "break"
        self.app.push_undo("清空单元格")
        Project.cell_set(self.project.devices[r], self._columns[c], "")
        self.project.mark_dirty()
        self.update_row(r)
        self.app.set_status("已清空单元格")
        return "break"

    # ── 设备增删 ──
    def _template_stem(self) -> str:
        tpl = self.project.templates.get(self.var_template.get())
        return tpl.stem if tpl else "device"

    def edit_device_details(self):
        self._commit_edit()
        idx = self._selected_index()
        if idx is None or idx >= len(self.project.devices):
            messagebox.showinfo("提示", "请先选中一台设备")
            return
        dev = self.project.devices[idx]
        host = dev.get("hostname") or f"第 {idx + 1} 行"
        dlg = DeviceDialog(self, dev, f"设备详情 — {host}")
        r = dlg.show()
        if not r:
            return
        self.app.push_undo("编辑设备详情")
        tpl_vars = self.project.template_vars(self.text.get("1.0", "end-1c"))
        for k, v in r["standard"].items():
            dev[k] = v
        dev["_extras"] = r["extras"]
        for k, v in r["standard"].items():
            if k in tpl_vars:
                dev["_extras"][k] = v
        if "sysname" in tpl_vars:
            dev["_extras"]["sysname"] = dev.get("hostname", "")
        self.project.mark_dirty()
        self.refresh_table()
        self.app.refresh_header()
        self.validate(silent=True)
        self.app.set_status(f"设备详情已更新: {dev.get('hostname') or host}")

    def add_device(self):
        self._commit_edit()
        if self.var_search.get().strip():
            self.var_search.set("")
        self.app.push_undo("添加设备")
        dev = nadt.derive_defaults({"type": self._template_stem()})
        dev["_extras"] = {"sysname": ""}
        self.project.devices.append(dev)
        self.project.mark_dirty()
        self.refresh_table()
        self.app.refresh_header()
        self.tree.see(str(len(self.project.devices) - 1))
        self.tree.selection_set(str(len(self.project.devices) - 1))
        self.after(30, lambda: self._begin_edit(str(len(self.project.devices) - 1), 0))

    def insert_row_below(self):
        self._commit_edit()
        if self.var_search.get().strip():
            self.var_search.set("")
        r = min(self._cursor[0] + 1, len(self.project.devices))
        self.app.push_undo("插入设备")
        dev = nadt.derive_defaults({"type": self._template_stem()})
        dev["_extras"] = {"sysname": ""}
        self.project.devices.insert(r, dev)
        self.project.mark_dirty()
        self.refresh_table()
        self.tree.selection_set(str(r))
        self.app.refresh_header()
        self.after(30, lambda: self._begin_edit(str(r), 0))

    def delete_devices(self):
        sel = list(self.tree.selection())
        if not sel:
            messagebox.showinfo("提示", "请先选中要删除的设备行")
            return
        if not messagebox.askyesno("删除设备", f"删除选中的 {len(sel)} 台设备？"):
            return
        self.app.push_undo("删除设备")
        for iid in sorted((int(x) for x in sel), reverse=True):
            if 0 <= iid < len(self.project.devices):
                self.project.devices.pop(iid)
        self.project.mark_dirty()
        self.refresh_table()
        self.app.refresh_header()
        self.app.set_status(f"已删除 {len(sel)} 台设备")

    def duplicate_selected(self):
        sel = self._target_indices()
        if not sel:
            messagebox.showinfo("提示", "请先选中要复制的设备行")
            return
        self.app.push_undo("复制设备行")
        new = services.duplicate_rows(self.project, sel)
        self.refresh_table()
        self.app.refresh_header()
        if new:
            self.tree.selection_set([str(i) for i in new])
        self.app.set_status(f"已复制 {len(new)} 行，请修改主机名/ESN")

    # ── 批量 ──
    def _target_indices(self) -> list[int]:
        sel = [int(x) for x in self.tree.selection()]
        return sel if sel else self._visible_indices()

    def batch_fill(self):
        if not self._columns:
            return
        sel = self._target_indices()
        dlg = FillDialog(self, self._columns)
        r = dlg.show()
        if not r:
            return
        self.app.push_undo("批量填充")
        n = services.fill_down(self.project, sel, [r["column"]], r["value"])
        self.refresh_table()
        self.app.refresh_header()
        self.app.set_status(f"批量填充 {n} 个单元格")
        self.log.log(f"✓ 批量填充 {n} 个单元格（{r['column']}）")

    def batch_sequence(self):
        if not self._columns:
            return
        sel = self._target_indices()
        dlg = SequenceDialog(self, self._columns)
        r = dlg.show()
        if not r:
            return
        self.app.push_undo("批量序列")
        n = services.sequence(self.project, sel, r["column"], r["start"], r["step"],
                              r["padding"], r["prefix"], r["suffix"])
        self.refresh_table()
        self.app.refresh_header()
        self.app.set_status(f"已写入 {n} 行序列")
        self.log.log(f"✓ 数字序列 {n} 行（{r['column']}）")

    def find_replace_dialog(self):
        dlg = FindReplaceDialog(self)
        r = dlg.show()
        if not r:
            return
        indices = self._visible_indices()
        cols = list(self._columns)
        self.app.push_undo("查找替换")
        n = services.find_replace(self.project, indices, cols, r["find"], r["replace"],
                                  r["regex"], r["case"])
        self.refresh_table()
        self.app.refresh_header()
        self.app.set_status(f"替换 {n} 个单元格")
        self.log.log(f"✓ 查找替换 {n} 个单元格")

    def import_table(self):
        path = filedialog.askopenfilename(
            title="导入表格（CSV / Excel）",
            filetypes=[("表格", "*.csv *.xlsx *.xlsm"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            header, rows = services.read_table(path)
        except Exception as e:
            messagebox.showerror("读取失败", str(e))
            return
        if not rows:
            messagebox.showwarning("没有数据", "表格除表头外没有数据行")
            return
        dlg = ColumnMapDialog(self, header, rows, self._columns, self.project)
        mapping = dlg.show()
        if not mapping:
            return
        self.app.push_undo("导入表格")
        n = services.import_rows(self.project, header, rows, mapping, self.var_template.get())
        self.refresh_table()
        self.app.refresh_header()
        self.validate(silent=True)
        self.app.set_status(f"已导入 {n} 台设备")
        self.log.log(f"✓ 从表格导入 {n} 台设备: {path}")

    # ── 校验 / 定位 ──
    def validate(self, silent=False) -> list[services.Issue]:
        self.sync_settings()
        self.save_template(silent=True)
        issues = services.Validator(self.project, self.var_template.get()).validate()
        self._issues = issues
        self.issue_tree.delete(*self.issue_tree.get_children())
        for i, it in enumerate(issues):
            self.issue_tree.insert("", tk.END, iid=str(i),
                                   values=(it.severity.upper(), it.device_label,
                                           it.field, it.message),
                                   tags=(it.severity,))
        errs = sum(1 for x in issues if x.severity == "error")
        warns = sum(1 for x in issues if x.severity == "warn")
        self.var_issue.set(f"错误 {errs} · 警告 {warns} · 提示 {len(issues) - errs - warns}")
        self.app.set_issue_count(errs + warns, errs)
        if not silent:
            self.bottom.select(0)
            self.app.set_status(f"校验完成: 错误 {errs} / 警告 {warns}")
        return issues

    def validate_and_show(self):
        self.validate(silent=False)

    def focus_issue(self):
        sel = self.issue_tree.selection()
        if not sel:
            return
        try:
            issue = self._issues[int(sel[0])]
        except (ValueError, IndexError):
            return
        if issue.device_index is None:
            self.app.set_status(issue.message)
            return
        idx = issue.device_index
        if idx not in self._visible_indices():
            self.var_search.set("")
            self.var_scope.set(SCOPE_LABELS["all"])
            self.project.settings.scope = "all"
            self.refresh_table()
        iid = str(idx)
        if not self.tree.exists(iid):
            self.refresh_table()
        self.tree.selection_set(iid)
        self.tree.see(iid)
        if issue.field and issue.field in self._columns:
            self._cursor = (idx, self._columns.index(issue.field))
            self.app.set_status(f"{issue.message} — 双击单元格修改")
        else:
            self.app.set_status(issue.message)

    # ── 预览 / 生成 ──
    def _selected_index(self) -> int | None:
        sel = self.tree.selection()
        if sel:
            try:
                return int(sel[0])
            except ValueError:
                return None
        visible = self._visible_indices()
        return visible[0] if visible else None

    def preview_selected(self):
        self.sync_settings()
        self.save_template(silent=True)
        idx = self._selected_index()
        if idx is None or idx >= len(self.project.devices):
            messagebox.showinfo("提示", "还没有可预览的设备")
            return
        dev = self.project.devices[idx]
        try:
            content = self.project.render(dev, self.var_template.get())
        except Exception as e:
            messagebox.showerror("预览失败", str(e))
            self.app.set_status(f"预览失败: {e}")
            return
        label = dev.get("hostname") or dev.get("esn") or f"第 {idx + 1} 行"
        PreviewDialog(self, self.project, label, content).show()

    def generate(self):
        self.sync_settings()
        self.save_template(silent=True)
        issues = self.validate(silent=True)
        errors = [x for x in issues if x.severity == "error"]
        if errors:
            self.bottom.select(0)
            first = "；".join(f"{x.device_label}: {x.message}" for x in errors[:3])
            if not messagebox.askyesno(
                    "发现错误",
                    f"校验发现 {len(errors)} 个错误。\n{first}\n\n仍要生成吗？",
                    default=messagebox.NO):
                self.app.set_status("已取消生成，请先处理问题")
                return
        try:
            plan = services.Generator(self.project, self.var_template.get()).plan()
        except Exception as e:
            messagebox.showerror("生成失败", str(e))
            return
        if not any(not f.error for f in plan.files):
            messagebox.showerror("无法生成", "\n".join(plan.errors[:8]) or "没有可写文件")
            return
        self._plan = plan
        dlg = GenerationDialog(self, plan, self.project)
        if not dlg.show():
            self.app.set_status("已返回修改")
            return
        result = services.Generator(self.project, self.var_template.get(),
                                    plan.out_dir).write(plan)
        self.app.refresh_header()
        self.log.log(f"✓ 写出 {result['written']} 个文件 → {plan.out_dir}")
        for e in result["errors"]:
            self.log.log(f"  ✗ {e}")
        if result["written"]:
            summary = f"新增 {plan.added} / 修改 {plan.changed} / 不变 {plan.unchanged}"
        else:
            summary = f"没有文件变化（{plan.unchanged} 个文件均为最新）"
        self.app.set_status(f"生成完成: {summary} → {plan.out_dir}")
        self._show_plan(plan)
        self.bottom.select(1)
        if hasattr(self.app.pages.get("deploy"), "set_output_dir"):
            self.app.pages["deploy"].set_output_dir(plan.out_dir)
        if result["errors"]:
            messagebox.showwarning("部分失败", "\n".join(result["errors"][:8]))
        else:
            self.log.log(f"完成: {summary}")

    def _show_plan(self, plan: services.GenerationPlan | None):
        self._plan = plan
        self.change_tree.delete(*self.change_tree.get_children())
        self.change_diff.set("生成后在这里查看每个文件的 Diff。")
        if plan is None:
            return
        for i, f in enumerate(plan.files):
            add, rem = (0, 0) if f.error else services.diff_stats(f.old_content, f.content)
            status = {"added": "新增", "changed": "修改", "unchanged": "不变",
                      "error": "错误"}.get(f.status, f.status)
            self.change_tree.insert("", tk.END, iid=str(i),
                                    values=(f.name, status, f"+{add}" if add else "",
                                            f"-{rem}" if rem else ""), tags=(f.status,))
        if plan.files:
            first = next((i for i, f in enumerate(plan.files) if f.changed), 0)
            self.change_tree.selection_set(str(first))
            self.change_tree.see(str(first))
            self._show_change_diff()

    def _show_change_diff(self):
        if not self._plan:
            return
        sel = self.change_tree.selection()
        if not sel:
            return
        try:
            f = self._plan.files[int(sel[0])]
        except (ValueError, IndexError):
            return
        if f.error:
            self.change_diff.set(f"生成失败:\n{f.error}")
            return
        if f.status == "unchanged":
            self.change_diff.set(f"{f.name}: 内容没有变化")
            return
        diff = services.diff_text(f.old_content, f.content, f.name, self.project)
        self.change_diff.set(diff or f"新文件 {f.name}（{len(f.content.splitlines())} 行）", diff=True)

    def _browse_output(self):
        path = filedialog.askdirectory(title="选择输出目录",
                                       initialdir=self.var_output.get() or os.getcwd())
        if path:
            self.var_output.set(path)
            self._on_setting_changed()

    def _set_sash(self):
        try:
            w = self.pane.winfo_width()
            if w > 100:
                self.pane.sashpos(0, max(460, int(w * 0.65)))
        except tk.TclError:
            pass
