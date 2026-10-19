"""通用 UI 组件：日志面板、滚动容器、树表、窗口居中。"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from . import theme


class Tooltip:
    """轻量工具提示；不抢焦点，鼠标离开即销毁。"""

    def __init__(self, widget, text: str, delay: int = 420):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._job = None
        self._tip = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _ev=None):
        self._cancel()
        self._job = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._job:
            try:
                self.widget.after_cancel(self._job)
            except tk.TclError:
                pass
            self._job = None

    def _show(self):
        if self._tip or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        except tk.TclError:
            return
        tip = tk.Toplevel(self.widget)
        tip.wm_overrideredirect(True)
        tip.wm_geometry(f"+{x}+{y}")
        frame = tk.Frame(tip, bg="#0f172a", bd=0)
        frame.pack()
        tk.Label(frame, text=self.text, bg="#0f172a", fg="#e2e8f0",
                 font=("Microsoft YaHei UI", 8), padx=9, pady=5,
                 justify="left").pack()
        self._tip = tip

    def _hide(self, _ev=None):
        self._cancel()
        if self._tip:
            try:
                self._tip.destroy()
            except tk.TclError:
                pass
            self._tip = None


class PageHeader(ttk.Frame):
    """页面标题 + 说明 + 右侧快捷键提示 + 底部分隔线。"""

    def __init__(self, parent, title: str, subtitle: str, hint: str = "", **kw):
        super().__init__(parent, **kw)
        self.columnconfigure(1, weight=1)
        tk.Frame(self, bg=theme.ACCENT, width=4, height=40).grid(
            row=0, column=0, sticky="w", padx=(0, 12), pady=(2, 0))
        box = ttk.Frame(self)
        box.grid(row=0, column=1, sticky="w")
        tk.Label(box, text=title, bg=theme.BG, fg=theme.FG,
                 font=theme.FONT_PAGE, anchor="w").pack(anchor="w")
        tk.Label(box, text=subtitle, bg=theme.BG, fg=theme.MUTED,
                 font=("Microsoft YaHei UI", 9), anchor="w").pack(anchor="w")
        if hint:
            tk.Label(self, text=hint, bg=theme.BG, fg=theme.MUTED,
                     font=("Microsoft YaHei UI", 8), anchor="e").grid(
                row=0, column=2, sticky="e", padx=(12, 2))
        ttk.Separator(self, orient=tk.HORIZONTAL).grid(
            row=1, column=0, columnspan=3, sticky="ew", pady=(9, 0))



class LogPanel(ttk.LabelFrame):
    def __init__(self, parent, title="日志", height=6, hint="", **kw):
        super().__init__(parent, text=title, padding=(8, 4, 8, 7), **kw)
        head = ttk.Frame(self)
        head.pack(fill=tk.X, pady=(0, 4))
        if hint:
            ttk.Label(head, text=hint, style="PanelMuted.TLabel").pack(side=tk.LEFT)
        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True)
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        self.text = tk.Text(body, height=height, font=theme.FONT_MONO, wrap=tk.WORD,
                            bg=theme.CONSOLE_BG, fg=theme.CONSOLE_FG,
                            insertbackground=theme.CONSOLE_FG,
                            selectbackground=theme.ACCENT_DARK, selectforeground="white",
                            relief=tk.FLAT, borderwidth=0, highlightthickness=0,
                            padx=10, pady=7, state=tk.DISABLED)
        vsb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.text.yview)
        self.text.configure(yscrollcommand=vsb.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        for tag, color in (("ok", "#4ade80"), ("err", "#f87171"), ("warn", "#fbbf24"),
                           ("info", "#93c5fd"), ("muted", "#94a3b8")):
            self.text.tag_configure(tag, foreground=color)
        ttk.Button(head, text="清空", style="Tool.TButton",
                   command=self.clear).pack(side=tk.RIGHT)
        ttk.Button(head, text="复制", style="Tool.TButton",
                   command=self.copy).pack(side=tk.RIGHT, padx=(0, 4))

    @staticmethod
    def tag_for(msg: str):
        if "✗" in msg:
            return "err"
        if "⚠" in msg:
            return "warn"
        if "✓" in msg:
            return "ok"
        if "ℹ" in msg:
            return "info"
        return None

    def log(self, msg: str, tag=None):
        self.text.config(state=tk.NORMAL)
        self.text.insert(tk.END, msg + "\n", tag or self.tag_for(msg))
        self.text.see(tk.END)
        self.text.config(state=tk.DISABLED)

    def clear(self):
        self.text.config(state=tk.NORMAL)
        self.text.delete("1.0", tk.END)
        self.text.config(state=tk.DISABLED)

    def copy(self):
        try:
            data = self.text.get("1.0", tk.END).strip()
            if data:
                self.text.clipboard_clear()
                self.text.clipboard_append(data)
        except tk.TclError:
            pass

    def get(self) -> str:
        return self.text.get("1.0", tk.END)


class ScrollableFrame(ttk.Frame):
    """可纵向滚动的容器；把内容放到 .inner。"""

    def __init__(self, parent, **kw):
        super().__init__(parent, **kw)
        self.canvas = tk.Canvas(self, bg=theme.BG, highlightthickness=0, bd=0)
        self.vsb = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.inner = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", self._on_inner)
        self.canvas.bind("<Configure>", self._on_canvas)
        self.canvas.bind("<Enter>", lambda _e: self._bind_wheel(True))
        self.canvas.bind("<Leave>", lambda _e: self._bind_wheel(False))

    def _on_inner(self, _ev=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas(self, ev):
        self.canvas.itemconfigure(self._win, width=ev.width)

    def _bind_wheel(self, on: bool):
        if on:
            self.canvas.bind_all("<MouseWheel>", self._wheel)
        else:
            self.canvas.unbind_all("<MouseWheel>")

    def _wheel(self, ev):
        self.canvas.yview_scroll(int(-ev.delta / 120), "units")


def make_tree(parent, columns: list[str], headings: dict[str, str] | None = None,
              widths: dict[str, int] | None = None, height=8, selectmode="extended"):
    frame = ttk.Frame(parent)
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)
    tree = ttk.Treeview(frame, columns=columns, show="headings",
                        selectmode=selectmode, height=height)
    headings = headings or {}
    widths = widths or {}
    for c in columns:
        tree.heading(c, text=headings.get(c, c))
        tree.column(c, width=widths.get(c, 120), anchor=tk.W,
                    stretch=c in ("hostname", "message", "name"))
    vsb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    tree.tag_configure("odd", background="#ffffff")
    tree.tag_configure("even", background="#f8fafd")
    return frame, tree


def center_window(win: tk.Misc, width: int, height: int):
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = max(0, (sw - width) // 2)
    y = max(0, (sh - height) // 2 - 20)
    win.geometry(f"{width}x{height}+{x}+{y}")


class DarkText(ttk.Frame):
    """深色只读代码框（预览/Diff 共用）。"""

    def __init__(self, parent, **kw):
        super().__init__(parent)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.text = tk.Text(self, font=theme.FONT_MONO, wrap=tk.NONE,
                            bg=theme.CONSOLE_BG, fg=theme.CONSOLE_FG,
                            insertbackground=theme.CONSOLE_FG,
                            selectbackground=theme.ACCENT_DARK,
                            selectforeground="white", relief=tk.FLAT,
                            borderwidth=0, highlightthickness=0,
                            padx=10, pady=8, state=tk.DISABLED)
        vsb = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.text.yview)
        hsb = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.text.xview)
        self.text.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.text.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.text.tag_configure("add", foreground="#4ade80")
        self.text.tag_configure("del", foreground="#f87171")
        self.text.tag_configure("hunk", foreground="#93c5fd")

    def set(self, content: str, diff: bool = False):
        self.text.config(state=tk.NORMAL)
        self.text.delete("1.0", tk.END)
        for line in content.splitlines():
            tag = None
            if diff:
                if line.startswith("+") and not line.startswith("+++"):
                    tag = "add"
                elif line.startswith("-") and not line.startswith("---"):
                    tag = "del"
                elif line.startswith("@@"):
                    tag = "hunk"
            self.text.insert(tk.END, line + "\n", tag)
        self.text.config(state=tk.DISABLED)
        self.text.see("1.0")

    def copy(self):
        try:
            data = self.text.get("1.0", tk.END).strip()
            if data:
                self.text.clipboard_clear()
                self.text.clipboard_append(data)
        except tk.TclError:
            pass
