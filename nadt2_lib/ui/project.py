"""项目页：项目信息、模板库、凭据、变量、设置、最近项目。"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from .. import services, secrets
from ..model import Credential, load_recent, clear_recent
from . import theme
from .dialogs import CredentialDialog, VariableDialog
from .widgets import PageHeader, ScrollableFrame, make_tree


class ProjectPage(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.project = app.project
        self._loading = False
        self._show_secrets = tk.BooleanVar(value=False)

        header = PageHeader(
            self, "项目设置",
            "一个 .nadt 装下设备、模板、凭据和生成设置",
            "Ctrl+3 切换工作区")
        header.pack(fill=tk.X, padx=12, pady=(10, 0))

        scroll = ScrollableFrame(self)
        scroll.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 10))
        body = scroll.inner
        body.columnconfigure(0, weight=1)

        self._build_info(body)
        self._build_templates(body)
        self._build_credentials(body)
        self._build_variables(body)
        self._build_settings(body)
        self._build_recent(body)

    def _card(self, parent, title: str, subtitle: str = "") -> ttk.LabelFrame:
        card = ttk.LabelFrame(parent, text=title, padding=(12, 8, 12, 10))
        card.pack(fill=tk.X, pady=(0, 8))
        if subtitle:
            ttk.Label(card, text=subtitle, style="PanelMuted.TLabel").pack(anchor="w", pady=(0, 6))
        return card

    # ── 项目信息 ──
    def _build_info(self, parent):
        card = self._card(parent, "基本信息", "一个 .nadt 文件就是一个完整方案：设备、模板、凭据、生成设置。")
        row = ttk.Frame(card)
        row.pack(fill=tk.X)
        ttk.Label(row, text="名称").pack(side=tk.LEFT)
        self.var_name = tk.StringVar()
        self.ent_name = ttk.Entry(row, textvariable=self.var_name, width=30)
        self.ent_name.pack(side=tk.LEFT, padx=(6, 4))
        self.ent_name.bind("<Return>", lambda _e: self.apply_name())
        ttk.Button(row, text="应用", style="Tool.TButton",
                   command=self.apply_name).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(row, text="保存项目", style="Accent.TButton",
                   command=self.app.save_project).pack(side=tk.RIGHT)
        ttk.Button(row, text="另存为…", style="Tool.TButton",
                   command=lambda: self.app.save_project(save_as=True)).pack(
            side=tk.RIGHT, padx=(0, 6))
        self.var_info = tk.StringVar()
        ttk.Label(card, textvariable=self.var_info, style="PanelMuted.TLabel").pack(
            anchor="w", pady=(6, 0))

    # ── 模板库 ──
    def _build_templates(self, parent):
        card = self._card(parent, "模板库", "每个模板 = 一段配置文本；表格列由 {{占位符}} 自动生成。")
        frame, self.tpl_tree = make_tree(
            card, ("name", "type", "lines", "default"), height=3,
            headings={"name": "模板", "type": "类型", "lines": "行数", "default": "默认"},
            widths={"name": 300, "type": 160, "lines": 70, "default": 70},
            selectmode="browse")
        frame.pack(fill=tk.X)
        btns = ttk.Frame(card)
        btns.pack(fill=tk.X, pady=(6, 0))
        for text, cmd in (("新建模板", self.add_template),
                          ("重命名", self.rename_template),
                          ("删除", self.delete_template),
                          ("设为默认", self.set_default_template),
                          ("从工作簿导入…", self.import_templates)):
            ttk.Button(btns, text=text, style="Tool.TButton",
                       command=cmd).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(btns, text="双击模板可在方案页编辑", style="PanelMuted.TLabel").pack(side=tk.RIGHT)
        self.tpl_tree.bind("<Double-1>", lambda _e: self.edit_template())

    # ── 凭据 ──
    def _build_credentials(self, parent):
        card = self._card(
            parent, "凭据",
            f"模板里写 {{{{变量名}}}} 即可引用；保存时使用 {secrets.protection_label()}。")
        frame, self.cred_tree = make_tree(
            card, ("key", "label", "value", "protection"), height=4,
            headings={"key": "变量名", "label": "说明", "value": "值", "protection": "保护"},
            widths={"key": 200, "label": 220, "value": 260, "protection": 100},
            selectmode="browse")
        frame.pack(fill=tk.X)
        btns = ttk.Frame(card)
        btns.pack(fill=tk.X, pady=(6, 0))
        for text, cmd in (("添加/编辑…", self.edit_credential),
                          ("删除", self.delete_credential)):
            ttk.Button(btns, text=text, style="Tool.TButton",
                       command=cmd).pack(side=tk.LEFT, padx=(0, 4))
        ttk.Checkbutton(btns, text="显示敏感值", variable=self._show_secrets,
                        style="Panel.TCheckbutton",
                        command=self.refresh_credentials).pack(side=tk.LEFT, padx=(12, 0))
        self.cred_tree.bind("<Double-1>", lambda _e: self.edit_credential())

    # ── 变量 ──
    def _build_variables(self, parent):
        card = self._card(parent, "项目变量", "非敏感的项目级变量（如 site=示例站点、ntp_server=192.0.2.1）。")
        frame, self.var_tree = make_tree(
            card, ("key", "value"), height=3,
            headings={"key": "变量名", "value": "值"},
            widths={"key": 240, "value": 520}, selectmode="browse")
        frame.pack(fill=tk.X)
        btns = ttk.Frame(card)
        btns.pack(fill=tk.X, pady=(6, 0))
        for text, cmd in (("添加/编辑…", self.edit_variable),
                          ("删除", self.delete_variable)):
            ttk.Button(btns, text=text, style="Tool.TButton",
                       command=cmd).pack(side=tk.LEFT, padx=(0, 4))
        self.var_tree.bind("<Double-1>", lambda _e: self.edit_variable())

    # ── 设置 ──
    def _build_settings(self, parent):
        card = self._card(parent, "项目设置", "生成细节在方案页调整；这里是项目级行为。")
        self.var_autosave = tk.BooleanVar()
        self.var_mask = tk.BooleanVar()
        ttk.Checkbutton(card, text="每 60 秒自动保存恢复副本（.autosave）",
                        variable=self.var_autosave,
                        style="Panel.TCheckbutton",
                        command=self._on_setting).pack(anchor="w")
        ttk.Checkbutton(card, text="预览 / Diff 默认隐藏凭据值",
                        variable=self.var_mask,
                        style="Panel.TCheckbutton",
                        command=self._on_setting).pack(anchor="w", pady=(4, 0))
        ttk.Label(card, text=f"凭据保护: {secrets.protection_label()}",
                  style="PanelMuted.TLabel").pack(anchor="w", pady=(8, 0))

    # ── 最近项目 ──
    def _build_recent(self, parent):
        card = self._card(parent, "最近项目")
        self.recent_list = tk.Listbox(card, height=5, font=theme.FONT_MONO,
                                      bg="white", fg=theme.FG, relief=tk.FLAT,
                                      highlightthickness=1, highlightbackground=theme.LINE)
        self.recent_list.pack(fill=tk.X)
        self.recent_list.bind("<Double-1>", lambda _e: self.open_recent())
        btns = ttk.Frame(card)
        btns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(btns, text="打开选中", style="Tool.TButton",
                   command=self.open_recent).pack(side=tk.LEFT)
        ttk.Button(btns, text="清空最近列表", style="Tool.TButton",
                   command=self.clear_recent).pack(side=tk.LEFT, padx=(6, 0))

    # ── 刷新 ──
    def reload_project(self):
        self.project = self.app.project
        self._loading = True
        try:
            self.var_name.set(self.project.name)
            p = self.project
            self.var_info.set(
                f"路径: {p.path or '未保存'}    ·    创建: {p.created}    ·    "
                f"修改: {p.modified}    ·    设备 {len(p.devices)} / 模板 {len(p.templates)}")
            self.refresh_templates()
            self.refresh_credentials()
            self.refresh_variables()
            self.var_autosave.set(bool(p.settings.autosave))
            self.var_mask.set(bool(p.settings.mask_secrets))
            self.refresh_recent()
        finally:
            self._loading = False

    def refresh_templates(self):
        self.tpl_tree.delete(*self.tpl_tree.get_children())
        default = self.project.settings.default_template
        for i, (name, tpl) in enumerate(self.project.templates.items()):
            self.tpl_tree.insert("", tk.END, iid=name,
                                 values=(name, tpl.type_name, len(tpl.text.splitlines()),
                                         "★" if name == default else ""),
                                 tags=("odd" if i % 2 else "even",))

    def refresh_credentials(self):
        self.cred_tree.delete(*self.cred_tree.get_children())
        for i, (key, cred) in enumerate(sorted(self.project.credentials.items())):
            if self._show_secrets.get():
                value = cred.value if cred.value else "(空)"
            else:
                value = "••••••" if cred.value else "(空)"
            if not cred.protected:
                protection = "明文"
            elif secrets.protection_label().startswith("Windows"):
                protection = "DPAPI"
            else:
                protection = "明文(降级)"
            self.cred_tree.insert("", tk.END, iid=key,
                                  values=(key, cred.label, value, protection),
                                  tags=("odd" if i % 2 else "even",))

    def refresh_variables(self):
        self.var_tree.delete(*self.var_tree.get_children())
        for i, (key, value) in enumerate(sorted(self.project.variables.items())):
            self.var_tree.insert("", tk.END, iid=key, values=(key, value),
                                 tags=("odd" if i % 2 else "even",))

    def refresh_recent(self):
        self.recent_list.delete(0, tk.END)
        for path in load_recent():
            self.recent_list.insert(tk.END, path)

    # ── 操作 ──
    def apply_name(self):
        name = self.var_name.get().strip()
        if not name or name == self.project.name:
            return
        self.app.push_undo("重命名项目")
        self.project.name = name
        self.project.mark_dirty()
        self.app.refresh_header()
        self.app.set_status(f"项目已重命名: {name}")

    def add_template(self):
        name = simpledialog.askstring("新建模板", "模板文件名（自动加 .cfg）:",
                                      initialvalue="new-template.cfg", parent=self)
        if not name:
            return
        if not name.endswith(".cfg"):
            name += ".cfg"
        try:
            self.app.push_undo("新建模板")
            self.project.add_template(name, "# {{hostname}}\n", os.path.splitext(name)[0])
        except Exception as e:
            messagebox.showerror("新建失败", str(e))
            return
        self.app.current_template = name
        self.project.settings.default_template = name
        self.project.mark_dirty()
        self.app.reload_all()
        self.app.pages["plan"].select_template(name)
        self.app.set_status(f"已新建模板: {name}")

    def rename_template(self):
        sel = self.tpl_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选中模板")
            return
        old = sel[0]
        new = simpledialog.askstring("重命名模板", "新文件名:", initialvalue=old, parent=self)
        if not new or new == old:
            return
        if not new.endswith(".cfg"):
            new += ".cfg"
        try:
            self.app.push_undo("重命名模板")
            new = self.project.rename_template(old, new)
        except Exception as e:
            messagebox.showerror("重命名失败", str(e))
            return
        if self.app.current_template == old:
            self.app.current_template = new
        self.project.mark_dirty()
        self.app.reload_all()
        self.app.set_status(f"模板已重命名: {old} → {new}")

    def delete_template(self):
        sel = self.tpl_tree.selection()
        if not sel:
            messagebox.showinfo("提示", "请先选中模板")
            return
        name = sel[0]
        if len(self.project.templates) <= 1:
            messagebox.showwarning("提示", "项目至少保留一个模板")
            return
        if not messagebox.askyesno("删除模板", f"删除模板 {name}？"):
            return
        self.app.push_undo("删除模板")
        self.project.remove_template(name)
        if self.app.current_template == name:
            self.app.current_template = self.project.settings.default_template
        self.app.reload_all()
        self.app.set_status(f"模板已删除: {name}")

    def set_default_template(self):
        sel = self.tpl_tree.selection()
        if not sel:
            return
        name = sel[0]
        self.app.push_undo("设置默认模板")
        self.project.settings.default_template = name
        self.app.current_template = name
        self.project.mark_dirty()
        self.refresh_templates()
        self.app.pages["plan"].select_template(name)
        self.app.refresh_header()
        self.app.set_status(f"默认模板: {name}")

    def edit_template(self):
        sel = self.tpl_tree.selection()
        if sel:
            self.app.current_template = sel[0]
            self.app.pages["plan"].select_template(sel[0])
            self.app.show_page("plan")

    def import_templates(self):
        path = filedialog.askopenfilename(
            title="从工作簿导入模板",
            filetypes=[("Excel 工作簿", "*.xlsm *.xlsx"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            info = services.import_workbook_as_templates(self.project, path)
        except Exception as e:
            messagebox.showerror("导入失败", str(e))
            return
        self.project.mark_dirty()
        self.app.reload_all()
        self.app.set_status(f"已导入 {info['templates']} 个模板")
        if info["warnings"]:
            messagebox.showinfo("导入完成", f"导入 {info['templates']} 个模板。\n\n" +
                                "\n".join(info["warnings"][:6]))

    def edit_credential(self):
        sel = self.cred_tree.selection()
        existing = self.project.credentials.get(sel[0]) if sel else None
        dlg = CredentialDialog(self, key=existing.key if existing else "",
                               label=existing.label if existing else "",
                               value=existing.value if existing else "",
                               protected=existing.protected if existing else True,
                               existing_keys=set(self.project.credentials))
        r = dlg.show()
        if not r:
            return
        key = r["key"]
        self.app.push_undo("保存凭据")
        if existing and key != existing.key:
            self.project.credentials.pop(existing.key, None)
        self.project.credentials[key] = Credential(
            key=key, value=r["value"], label=r["label"], protected=r["protected"])
        self.project.mark_dirty()
        self.refresh_credentials()
        self.app.refresh_header()
        self.app.set_status(f"凭据已保存: {key}")
        self.app.pages["plan"].validate(silent=True)

    def delete_credential(self):
        sel = self.cred_tree.selection()
        if not sel:
            return
        key = sel[0]
        if messagebox.askyesno("删除凭据", f"删除凭据 {key}？"):
            self.app.push_undo("删除凭据")
            self.project.credentials.pop(key, None)
            self.project.mark_dirty()
            self.refresh_credentials()
            self.app.set_status(f"凭据已删除: {key}")

    def edit_variable(self):
        sel = self.var_tree.selection()
        key = sel[0] if sel else ""
        value = self.project.variables.get(key, "") if key else ""
        dlg = VariableDialog(self, key=key, value=value)
        r = dlg.show()
        if not r:
            return
        self.app.push_undo("保存变量")
        if key and r["key"] != key:
            self.project.variables.pop(key, None)
        self.project.variables[r["key"]] = r["value"]
        self.project.mark_dirty()
        self.refresh_variables()
        self.app.refresh_header()
        self.app.set_status(f"变量已保存: {r['key']}")

    def delete_variable(self):
        sel = self.var_tree.selection()
        if not sel:
            return
        key = sel[0]
        if messagebox.askyesno("删除变量", f"删除变量 {key}？"):
            self.app.push_undo("删除变量")
            self.project.variables.pop(key, None)
            self.project.mark_dirty()
            self.refresh_variables()

    def _on_setting(self):
        if self._loading:
            return
        self.project.settings.autosave = bool(self.var_autosave.get())
        self.project.settings.mask_secrets = bool(self.var_mask.get())
        self.project.mark_dirty()
        self.app.set_status("项目设置已更新")

    def open_recent(self):
        sel = self.recent_list.curselection()
        if not sel:
            return
        self.app.open_project(self.recent_list.get(sel[0]))

    def clear_recent(self):
        clear_recent()
        self.refresh_recent()
        self.app._refresh_recent_menu()
