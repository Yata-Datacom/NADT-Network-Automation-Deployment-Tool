"""NADT 2.0 主窗口：项目文件、菜单、导航、状态栏、撤销/自动保存。"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import nadt
from .. import APP_NAME, __version__
from ..model import Project, load_recent
from .. import services
from . import theme
from .deploy import DeployPage
from .plan import PlanPage
from .project import ProjectPage
from .widgets import Tooltip, center_window


class NadtApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_NAME} {__version__}")
        theme.setup(root)
        try:
            icon = os.path.join(nadt.ASSET_DIR, "nadt.ico")
            if os.path.isfile(icon):
                self.root.iconbitmap(default=icon)
        except tk.TclError:
            pass
        center_window(root, 1360, 880)
        root.minsize(1160, 740)

        self.project: Project = Project.create_default()
        self.current_template: str = self.project.settings.default_template
        self._undo: list[tuple[str, dict]] = []
        self._redo: list[tuple[str, dict]] = []
        self._autosave_job = None
        self._issue_count = 0
        self._issue_errors = 0

        self._build_menu()
        self._build_header()
        self._build_body()
        self._build_statusbar()
        self._bind_keys()
        self.reload_all()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(60000, self._autosave_tick)

    # ── 布局 ──
    def _build_header(self):
        header = tk.Frame(self.root, bg=theme.NAVY)
        header.grid(row=0, column=0, sticky="ew")
        brand = tk.Frame(header, bg=theme.NAVY)
        brand.pack(side=tk.LEFT, padx=(18, 0), pady=10)
        logo = tk.Canvas(brand, width=40, height=40, bg=theme.NAVY,
                         highlightthickness=0, bd=0)
        logo.pack(side=tk.LEFT, padx=(0, 11))
        logo.create_polygon(8, 1, 32, 1, 39, 8, 39, 32, 32, 39, 8, 39, 1, 32, 1, 8,
                            smooth=True, splinesteps=18, fill="#1d4ed8", outline="")
        for nx, ny in ((20, 7), (33, 20), (20, 33), (7, 20)):
            logo.create_line(20, 20, nx, ny, fill="#7dd3fc", width=2)
        for nx, ny in ((20, 7), (33, 20), (20, 33), (7, 20)):
            logo.create_oval(nx - 3, ny - 3, nx + 3, ny + 3, fill="#34d399",
                             outline=theme.NAVY)
        logo.create_oval(15, 15, 25, 25, fill="#38bdf8", outline="#e0f2fe", width=1)

        titles = tk.Frame(brand, bg=theme.NAVY)
        titles.pack(side=tk.LEFT)
        tk.Label(titles, text=f"{APP_NAME}", bg=theme.NAVY, fg=theme.NAVY_FG,
                 font=theme.FONT_TITLE).pack(anchor="w")
        self.project_title = tk.Label(titles, text="未命名方案", bg=theme.NAVY,
                                      fg="#94a3b8", font=("Microsoft YaHei UI", 9))
        self.project_title.pack(anchor="w")

        chips = tk.Frame(header, bg=theme.NAVY)
        chips.pack(side=tk.RIGHT, padx=(0, 18))
        self.chip_issue = tk.Label(chips, text="问题 0", bg=theme.NAVY2, fg="#94a3b8",
                                   font=("Microsoft YaHei UI", 8, "bold"), padx=10, pady=4)
        self.chip_issue.pack(side=tk.RIGHT, padx=(6, 0))
        self.chip_tpl = tk.Label(chips, text="模板 1", bg=theme.NAVY2, fg="#e2e8f0",
                                 font=("Microsoft YaHei UI", 8, "bold"), padx=10, pady=4)
        self.chip_tpl.pack(side=tk.RIGHT, padx=(6, 0))
        self.chip_dev = tk.Label(chips, text="设备 0", bg=theme.NAVY2, fg="#e2e8f0",
                                 font=("Microsoft YaHei UI", 8, "bold"), padx=10, pady=4)
        self.chip_dev.pack(side=tk.RIGHT, padx=(6, 0))
        tk.Label(chips, text=f"v{__version__}", bg="#172554", fg="#93c5fd",
                 font=("Microsoft YaHei UI", 8, "bold"), padx=10, pady=4).pack(side=tk.RIGHT)

        actions = tk.Frame(header, bg=theme.NAVY)
        actions.pack(side=tk.RIGHT, padx=(0, 12))
        for text, cmd, tip in (("新建", self.new_project, "新建项目  Ctrl+N"),
                               ("打开", self.open_project, "打开 .nadt 项目  Ctrl+O"),
                               ("保存", self.save_project, "保存项目  Ctrl+S")):
            btn = tk.Button(actions, text=text, command=cmd, bg=theme.NAVY2,
                            fg="#cbd5e1", activebackground="#334155",
                            activeforeground="white", relief=tk.FLAT, bd=0,
                            padx=10, pady=5, font=("Microsoft YaHei UI", 8),
                            cursor="hand2")
            btn.pack(side=tk.LEFT, padx=(0, 5))
            Tooltip(btn, tip)
        tk.Frame(header, bg="#1e3a8a", height=2).pack(side=tk.BOTTOM, fill="x")

    def _build_body(self):
        body = ttk.Frame(self.root)
        body.grid(row=1, column=0, sticky="nsew")
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        rail = tk.Frame(body, bg=theme.NAVY, width=180)
        rail.pack(side=tk.LEFT, fill=tk.Y)
        rail.pack_propagate(False)
        tk.Label(rail, text="工作区", bg=theme.NAVY, fg="#475569",
                 font=("Microsoft YaHei UI", 8), anchor="w").pack(fill="x", padx=16, pady=(16, 6))
        self.nav = {}
        for key, title, sub in (("plan", "方案", "模板 · 数据 · 生成"),
                                ("deploy", "部署", "TFTP · DHCP · USB"),
                                ("project", "项目", "模板 · 凭据 · 设置")):
            btn = theme.NavButton(rail, title, sub, lambda k=key: self.show_page(k), self)
            btn.pack(fill="x", padx=8, pady=3)
            self.nav[key] = btn
        tk.Frame(rail, bg=theme.NAVY).pack(fill=tk.BOTH, expand=True)
        tk.Label(rail, text="NADT 2.0 · 零接触开局", bg=theme.NAVY, fg="#475569",
                 font=("Microsoft YaHei UI", 7)).pack(pady=12)

        content = ttk.Frame(body)
        content.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 0))
        content.grid_rowconfigure(0, weight=1)
        content.grid_columnconfigure(0, weight=1)
        self.pages = {}
        for key, cls in (("plan", PlanPage), ("deploy", DeployPage),
                         ("project", ProjectPage)):
            page = cls(content, self)
            page.grid(row=0, column=0, sticky="nsew")
            self.pages[key] = page
        self.current_page = "plan"
        self.nav["plan"].set_active(True)
        self.pages["plan"].tkraise()

    def _build_statusbar(self):
        bar = tk.Frame(self.root, bg=theme.PANEL, highlightthickness=1,
                       highlightbackground=theme.LINE)
        bar.grid(row=2, column=0, sticky="ew")
        self.status_var = tk.StringVar(value="就绪 · 新建方案，或打开一个 .nadt 项目")
        tk.Label(bar, text="●", bg=theme.PANEL, fg=theme.GREEN,
                 font=("Microsoft YaHei UI", 8), padx=14, pady=6).pack(side=tk.LEFT)
        tk.Label(bar, textvariable=self.status_var, bg=theme.PANEL, fg=theme.FG,
                 font=theme.FONT, anchor="w", padx=2, pady=6).pack(
            side=tk.LEFT, fill="x", expand=True)
        self.path_var = tk.StringVar(value="未保存")
        tk.Label(bar, textvariable=self.path_var, bg=theme.PANEL, fg=theme.MUTED,
                 font=("Consolas", 8), padx=14, pady=6).pack(side=tk.RIGHT)
        tk.Label(bar, text="Ctrl+1/2/3 切换工作区  ·  F6 校验  ·  F5 生成",
                 bg=theme.PANEL, fg=theme.MUTED, font=("Microsoft YaHei UI", 8),
                 padx=10, pady=6).pack(side=tk.RIGHT)

    def _build_menu(self):
        menubar = tk.Menu(self.root, tearoff=0)
        m_project = tk.Menu(menubar, tearoff=0)
        m_project.add_command(label="新建项目…", accelerator="Ctrl+N", command=self.new_project)
        m_project.add_command(label="打开项目…", accelerator="Ctrl+O", command=self.open_project)
        m_project.add_command(label="保存", accelerator="Ctrl+S", command=self.save_project)
        m_project.add_command(label="另存为…", accelerator="Ctrl+Shift+S",
                              command=lambda: self.save_project(save_as=True))
        self.recent_menu = tk.Menu(m_project, tearoff=0)
        m_project.add_cascade(label="最近打开", menu=self.recent_menu)
        m_project.add_separator()
        m_project.add_command(label="导入旧 xlsm 工作簿…", command=self.import_legacy)
        m_project.add_command(label="导出 xlsm 兼容工作簿…", command=self.export_workbook)
        m_project.add_separator()
        m_project.add_command(label="退出", command=self.on_close)
        menubar.add_cascade(label="项目", menu=m_project)

        m_edit = tk.Menu(menubar, tearoff=0)
        m_edit.add_command(label="撤销", accelerator="Ctrl+Z", command=self.undo)
        m_edit.add_command(label="重做", accelerator="Ctrl+Y", command=self.redo)
        m_edit.add_separator()
        m_edit.add_command(label="设备详情（全部字段）…", accelerator="Ctrl+E",
                           command=lambda: self._plan_action("edit_device_details"))
        m_edit.add_command(label="批量操作…",
                           command=lambda: self._plan_action("show_batch_menu"))
        m_edit.add_command(label="查找替换…", accelerator="Ctrl+H",
                           command=lambda: self._plan_action("find_replace_dialog"))
        m_edit.add_command(label="校验项目", accelerator="F6",
                           command=lambda: self._plan_action("validate_and_show"))
        menubar.add_cascade(label="编辑", menu=m_edit)

        m_gen = tk.Menu(menubar, tearoff=0)
        m_gen.add_command(label="预览选中设备", accelerator="Ctrl+Enter",
                          command=lambda: self._plan_action("preview_selected"))
        m_gen.add_command(label="生成…", accelerator="F5",
                          command=lambda: self._plan_action("generate"))
        menubar.add_cascade(label="生成", menu=m_gen)

        m_view = tk.Menu(menubar, tearoff=0)
        m_view.add_command(label="方案工作台", accelerator="Ctrl+1",
                           command=lambda: self.show_page("plan"))
        m_view.add_command(label="部署", accelerator="Ctrl+2",
                           command=lambda: self.show_page("deploy"))
        m_view.add_command(label="项目设置", accelerator="Ctrl+3",
                           command=lambda: self.show_page("project"))
        menubar.add_cascade(label="视图", menu=m_view)

        m_help = tk.Menu(menubar, tearoff=0)
        m_help.add_command(label="快捷键", command=self.show_shortcuts)
        m_help.add_command(label="关于 NADT 2.0", command=self.show_about)
        menubar.add_cascade(label="帮助", menu=m_help)
        self.root.config(menu=menubar)
        self._refresh_recent_menu()

    def _bind_keys(self):
        for seq, fn in (
                ("<Control-n>", lambda _e: self.new_project()),
                ("<Control-o>", lambda _e: self.open_project()),
                ("<Control-s>", lambda _e: self.save_project()),
                ("<Control-Shift-S>", lambda _e: self.save_project(save_as=True)),
                ("<Control-z>", lambda _e: self.undo()),
                ("<Control-Z>", lambda _e: self.undo()),
                ("<Control-y>", lambda _e: self.redo()),
                ("<Control-Y>", lambda _e: self.redo()),
                ("<F5>", lambda _e: self.pages["plan"].generate()),
                ("<F6>", lambda _e: self.pages["plan"].validate_and_show()),
                ("<Control-Return>", lambda _e: self.pages["plan"].preview_selected()),
                ("<Control-h>", lambda _e: self.pages["plan"].find_replace_dialog()),
                ("<Control-f>", lambda _e: self._plan_action("focus_search")),
                ("<Control-e>", lambda _e: self._plan_action("edit_device_details")),
                ("<Control-Key-1>", lambda _e: self.show_page("plan")),
                ("<Control-Key-2>", lambda _e: self.show_page("deploy")),
                ("<Control-Key-3>", lambda _e: self.show_page("project")),
        ):
            self.root.bind_all(seq, fn)

    # ── 页面 / 状态 ──
    def show_page(self, key: str):
        if key not in self.pages:
            return
        self.current_page = key
        for k, btn in self.nav.items():
            btn.set_active(k == key)
        self.pages[key].tkraise()
        if key == "deploy":
            self.pages["deploy"].on_show()
        elif key == "project":
            self.pages["project"].reload_project()
        elif key == "plan":
            self.pages["plan"].on_show()

    def set_status(self, msg: str, kind: str = "info"):
        self.status_var.set(msg)

    def _plan_action(self, name: str):
        self.show_page("plan")
        method = getattr(self.pages["plan"], name, None)
        if method:
            method()

    def refresh_header(self):
        p = self.project
        star = " *" if p.dirty else ""
        self.project_title.config(text=f"{p.name}{star}")
        self.root.title(f"{APP_NAME} — {p.name}{star}")
        self.chip_dev.config(text=f"设备 {len(p.devices)}")
        self.chip_tpl.config(text=f"模板 {len(p.templates)}")
        if self._issue_count:
            if self._issue_errors:
                self.chip_issue.config(text=f"问题 {self._issue_count}", fg="#fca5a5",
                                       bg="#3f1d1d")
            else:
                self.chip_issue.config(text=f"警告 {self._issue_count}", fg="#fcd34d",
                                       bg="#3a2a0a")
        else:
            self.chip_issue.config(text="问题 0", fg="#4ade80", bg="#0f2a1d")
        self.path_var.set(p.path or "未保存")

    def set_issue_count(self, n: int, errors: int = 0):
        self._issue_count = n
        self._issue_errors = errors
        self.refresh_header()

    def reload_all(self):
        self.current_template = (self.project.settings.default_template
                                 or next(iter(self.project.templates), ""))
        for page in self.pages.values():
            page.reload_project()
        self.refresh_header()
        self._refresh_recent_menu()

    # ── 项目生命周期 ──
    def maybe_save(self) -> bool:
        if not self.project.dirty:
            return True
        ans = messagebox.askyesnocancel("未保存的修改", "当前项目有未保存的修改，是否保存？")
        if ans is None:
            return False
        if ans:
            return self.save_project()
        return True

    def new_project(self):
        if not self.maybe_save():
            return
        name = simpledialog.askstring("新建项目", "项目名称:", initialvalue="新方案",
                                      parent=self.root)
        if not name:
            return
        self.project = Project.create_default(name.strip() or "新方案")
        self.current_template = self.project.settings.default_template
        self._undo.clear()
        self._redo.clear()
        self.reload_all()
        self.set_status(f"已新建项目: {self.project.name}")

    def open_project(self, path: str | None = None):
        if not self.maybe_save():
            return
        path = path or filedialog.askopenfilename(
            title="打开 NADT 项目", filetypes=[("NADT 项目", "*.nadt"),
                                               ("所有文件", "*.*")])
        if not path:
            return
        try:
            project = Project.load(path)
        except Exception as e:
            messagebox.showerror("打开失败", str(e))
            return
        autosave = project.autosave_path()
        if os.path.isfile(autosave) and os.path.getmtime(autosave) > os.path.getmtime(path):
            if messagebox.askyesno("发现自动保存", "检测到比项目文件更新的自动保存，是否恢复？"):
                try:
                    with open(autosave, "r", encoding="utf-8") as f:
                        import json
                        project = Project.from_dict(json.load(f))
                    project.path = os.path.abspath(path)
                    project.dirty = True
                except Exception as e:
                    messagebox.showwarning("恢复失败", str(e))
        self.project = project
        self.current_template = (project.settings.default_template
                                 or next(iter(project.templates), ""))
        self._undo.clear()
        self._redo.clear()
        self.reload_all()
        self.set_status(f"已打开: {path}")

    def save_project(self, save_as: bool = False) -> bool:
        try:
            self.pages["plan"].save_template(silent=True)
            self.pages["plan"].sync_settings()
        except Exception:
            pass
        p = self.project
        path = p.path
        if save_as or not path:
            path = filedialog.asksaveasfilename(
                title="保存 NADT 项目", defaultextension=".nadt",
                initialfile=f"{p.name}.nadt",
                filetypes=[("NADT 项目", "*.nadt")])
            if not path:
                return False
        try:
            p.save(path)
        except Exception as e:
            messagebox.showerror("保存失败", str(e))
            return False
        self.refresh_header()
        self.set_status(f"已保存: {p.path}")
        return True

    def import_legacy(self):
        path = filedialog.askopenfilename(
            title="导入旧 xlsm 工作簿",
            filetypes=[("Excel 工作簿", "*.xlsm *.xlsx"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            if self.project.devices:
                ans = messagebox.askyesnocancel(
                    "导入方式", "当前项目已有设备。\n\n是 = 追加到当前项目\n否 = 替换当前项目\n取消 = 放弃")
                if ans is None:
                    return
                replace = not ans
            else:
                replace = False
            self.push_undo("导入旧工作簿")
            if replace:
                self.project.devices = []
                self.project.templates = {}
                self.project.type_map = {}
            info = services.import_legacy_workbook(self.project, path)
        except Exception as e:
            messagebox.showerror("导入失败", str(e))
            return
        self.project.mark_dirty()
        self.current_template = self.project.settings.default_template
        self.reload_all()
        self.pages["plan"].select_template(self.current_template)
        msg = f"已导入 {info['templates']} 个模板 / {info['devices']} 台设备"
        if info["warnings"]:
            msg += f"；{len(info['warnings'])} 条提示"
        self.set_status(msg)
        messagebox.showinfo("导入完成", msg + "\n\n第一个模板已载入方案页。")

    def export_workbook(self):
        path = filedialog.asksaveasfilename(
            title="导出 xlsm 兼容工作簿", defaultextension=".xlsx",
            initialfile=f"{self.project.name}-export.xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not path:
            return
        try:
            services.export_workbook(self.project, path)
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            return
        self.set_status(f"已导出工作簿: {path}")
        try:
            os.startfile(path)
        except OSError:
            pass

    # ── 撤销 / 重做 ──
    def push_undo(self, label: str):
        self._undo.append((label, self.project.snapshot()))
        if len(self._undo) > 40:
            self._undo.pop(0)
        self._redo.clear()

    def undo(self):
        if not self._undo:
            self.set_status("没有可撤销的操作")
            return
        label, snap = self._undo.pop()
        self._redo.append((label, self.project.snapshot()))
        self.project.restore(snap)
        self.reload_all()
        self.set_status(f"已撤销: {label}")

    def redo(self):
        if not self._redo:
            self.set_status("没有可重做的操作")
            return
        label, snap = self._redo.pop()
        self._undo.append((label, self.project.snapshot()))
        self.project.restore(snap)
        self.reload_all()
        self.set_status(f"已重做: {label}")

    # ── 自动保存 / 关闭 ──
    def _autosave_tick(self):
        try:
            if self.project.dirty and self.project.settings.autosave and self.project.path:
                self.project.write_autosave()
        finally:
            self.root.after(60000, self._autosave_tick)

    def on_close(self):
        try:
            self.pages["plan"].save_template(silent=True)
            self.pages["plan"].sync_settings()
        except Exception:
            pass
        if not self.maybe_save():
            return
        try:
            self.pages["deploy"].stop_tftp()
        except Exception:
            pass
        try:
            self.project.clear_autosave()
        except Exception:
            pass
        self.root.destroy()

    def _refresh_recent_menu(self):
        self.recent_menu.delete(0, tk.END)
        items = load_recent()
        if not items:
            self.recent_menu.add_command(label="（暂无）", state=tk.DISABLED)
            return
        for path in items:
            self.recent_menu.add_command(
                label=path, command=lambda p=path: self.open_project(p))

    # ── 帮助 ──
    def show_shortcuts(self):
        messagebox.showinfo(
            "快捷键",
            "Ctrl+N 新建项目\nCtrl+O 打开项目\nCtrl+S 保存 / Ctrl+Shift+S 另存为\n"
            "Ctrl+1 方案 / Ctrl+2 部署 / Ctrl+3 项目\n"
            "Ctrl+Z 撤销 / Ctrl+Y 重做\nF6 校验项目\nF5 生成\nCtrl+Enter 预览选中设备\n"
            "Ctrl+E 设备详情（全部字段）\nCtrl+F 搜索设备\nCtrl+H 查找替换\n"
            "双击单元格编辑；Ctrl+V 粘贴 Excel 区域")

    def show_about(self):
        messagebox.showinfo(
            "关于 NADT 2.0",
            f"NADT {__version__}\nNetwork Automation Deployment Tool\n\n"
            "项目文件 + 模板驱动数据表 + 校验/Diff + 凭据保护\n"
            "核心渲染基于 DHCP option 66/67 + TFTP 的零接触开局。")


def main():
    root = tk.Tk()
    NadtApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
