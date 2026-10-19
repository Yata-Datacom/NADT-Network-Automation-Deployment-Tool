"""NADT 2.0 视觉主题。"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

BG = "#f3f5fa"
PANEL = "#ffffff"
PANEL_ALT = "#f8fafd"
NAVY = "#0f172a"
NAVY2 = "#1e293b"
NAVY_FG = "#f8fafc"
FG = "#1e293b"
MUTED = "#64748b"
LINE = "#dde5f0"
ACCENT = "#2563eb"
ACCENT2 = "#3b82f6"
ACCENT_DARK = "#1e40af"
ACCENT_SOFT = "#e8effd"
GREEN = "#16a34a"
GREEN_SOFT = "#dcfce7"
AMBER = "#b45309"
AMBER_SOFT = "#fef3c7"
RED = "#dc2626"
RED_SOFT = "#fee2e2"
CONSOLE_BG = "#0b1220"
CONSOLE_FG = "#cbd5e1"
FONT = ("Microsoft YaHei UI", 9)
FONT_B = ("Microsoft YaHei UI", 9, "bold")
FONT_H = ("Microsoft YaHei UI", 10, "bold")
FONT_PAGE = ("Microsoft YaHei UI", 17, "bold")
FONT_TITLE = ("Microsoft YaHei UI", 15, "bold")
FONT_MONO = ("Consolas", 9)
FONT_MONO_L = ("Consolas", 10)


def setup(root: tk.Tk):
    root.configure(bg=BG)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure(".", background=BG, foreground=FG, bordercolor=LINE,
                    lightcolor=BG, darkcolor=BG, focuscolor=ACCENT)
    style.configure("TFrame", background=BG)
    style.configure("Panel.TFrame", background=PANEL)
    style.configure("TLabel", background=BG, foreground=FG, font=FONT)
    style.configure("Panel.TLabel", background=PANEL, foreground=FG, font=FONT)
    style.configure("Muted.TLabel", background=BG, foreground=MUTED,
                    font=("Microsoft YaHei UI", 8))
    style.configure("PanelMuted.TLabel", background=PANEL, foreground=MUTED,
                    font=("Microsoft YaHei UI", 8))
    style.configure("Title.TLabel", background=PANEL, foreground=FG, font=FONT_TITLE)
    style.configure("TLabelframe", background=BG, bordercolor=LINE, relief="solid", borderwidth=1)
    style.configure("TLabelframe.Label", background=BG, foreground=ACCENT_DARK, font=FONT_H)
    style.configure("TButton", background="#e8edf5", foreground=FG, borderwidth=0,
                    focusthickness=0, padding=(12, 7), font=FONT)
    style.map("TButton", background=[("pressed", ACCENT_DARK), ("active", "#dbe6f8")],
              foreground=[("pressed", "white")])
    style.configure("Tool.TButton", background="#eef2f9", foreground="#334155",
                    padding=(9, 5), font=("Microsoft YaHei UI", 9))
    style.map("Tool.TButton", background=[("pressed", "#dbe6f8"), ("active", "#e2eafb")],
              foreground=[("pressed", ACCENT_DARK)])
    style.configure("Bar.TButton", background="#eef2f9", foreground="#334155",
                    padding=(7, 5), font=("Microsoft YaHei UI", 9))
    style.map("Bar.TButton", background=[("pressed", "#dbe6f8"), ("active", "#e2eafb")],
              foreground=[("pressed", ACCENT_DARK)])
    style.configure("Accent.TButton", background=ACCENT, foreground="white",
                    padding=(16, 8), font=FONT_B)
    style.map("Accent.TButton", background=[("pressed", ACCENT_DARK), ("active", ACCENT2)],
              foreground=[("pressed", "white"), ("disabled", "#bfdbfe")])
    style.configure("Danger.TButton", background=RED, foreground="white",
                    padding=(12, 7), font=FONT_B)
    style.map("Danger.TButton", background=[("pressed", "#991b1b"), ("active", "#ef4444")],
              foreground=[("disabled", "#fecaca")])
    style.configure("DangerGhost.TButton", background=RED_SOFT, foreground="#b91c1c",
                    padding=(10, 6), font=("Microsoft YaHei UI", 9, "bold"))
    style.map("DangerGhost.TButton", background=[("pressed", "#fecaca"), ("active", "#fecaca")])
    style.configure("GreenAccent.TButton", background=GREEN, foreground="white",
                    padding=(16, 8), font=FONT_B)
    style.map("GreenAccent.TButton", background=[("pressed", "#15803d"), ("active", "#4ade80")],
              foreground=[("pressed", "white"), ("disabled", "#bbf7d0")])
    style.configure("TEntry", fieldbackground="white", foreground=FG, bordercolor=LINE,
                    padding=6, insertcolor=ACCENT)
    style.map("TEntry", bordercolor=[("focus", ACCENT)], lightcolor=[("focus", ACCENT)])
    style.configure("TCombobox", fieldbackground="white", foreground=FG, bordercolor=LINE,
                    padding=6, arrowcolor=ACCENT)
    style.map("TCombobox", bordercolor=[("focus", ACCENT)], fieldbackground=[("readonly", "white")])
    style.configure("TCheckbutton", background=BG, foreground=FG, font=FONT, focuscolor=ACCENT)
    style.map("TCheckbutton", background=[("active", BG)])
    style.configure("Panel.TCheckbutton", background=PANEL, foreground=FG, font=FONT,
                    focuscolor=ACCENT)
    style.map("Panel.TCheckbutton", background=[("active", PANEL)])
    style.configure("TNotebook", background=BG, borderwidth=0, tabmargins=(4, 4, 4, 0))
    style.configure("TNotebook.Tab", background="#e6ecf6", foreground=MUTED,
                    padding=(15, 9), borderwidth=0, font=FONT)
    style.map("TNotebook.Tab", background=[("selected", PANEL), ("active", "#eef3fb")],
              foreground=[("selected", ACCENT)], font=[("selected", FONT_H)])
    style.configure("Treeview", background="white", fieldbackground="white",
                    foreground=FG, rowheight=29, borderwidth=0, font=FONT)
    style.configure("Treeview.Heading", background="#eef2f9", foreground="#334155",
                    padding=(9, 8), relief="flat", font=FONT_B)
    style.map("Treeview.Heading", background=[("active", "#e4ebf7")])
    style.map("Treeview", background=[("selected", "#dbeafe")],
              foreground=[("selected", "#0f172a")])
    style.configure("TProgressbar", background=ACCENT, troughcolor="#e2e8f0",
                    borderwidth=0, thickness=9)
    style.configure("Vertical.TScrollbar", background="#cbd5e1", troughcolor=BG,
                    borderwidth=0, arrowsize=12)
    style.configure("Horizontal.TScrollbar", background="#cbd5e1", troughcolor=BG,
                    borderwidth=0, arrowsize=12)
    return style


class NavButton(tk.Frame):
    """左侧导航按钮（自绘，选中时高亮）。"""

    def __init__(self, parent, text: str, sub: str, command, app):
        super().__init__(parent, bg=NAVY, cursor="hand2")
        self.command = command
        self.app = app
        self._active = False
        self.dot = tk.Frame(self, bg=NAVY, width=4)
        self.dot.pack(side=tk.LEFT, fill=tk.Y)
        box = tk.Frame(self, bg=NAVY)
        box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(12, 8), pady=8)
        self.title = tk.Label(box, text=text, bg=NAVY, fg="#e2e8f0",
                              font=("Microsoft YaHei UI", 10, "bold"), anchor="w")
        self.title.pack(fill=tk.X)
        self.sub = tk.Label(box, text=sub, bg=NAVY, fg="#64748b",
                            font=("Microsoft YaHei UI", 8), anchor="w")
        self.sub.pack(fill=tk.X)
        for w in (self, box, self.title, self.sub):
            w.bind("<Button-1>", lambda _e: self.command())
            w.bind("<Enter>", lambda _e: self._hover(True))
            w.bind("<Leave>", lambda _e: self._hover(False))

    def _hover(self, on: bool):
        if self._active:
            return
        bg = "#172033" if on else NAVY
        for w in (self, self.dot) + tuple(self.winfo_children()):
            try:
                w.configure(bg=bg)
            except tk.TclError:
                pass
        for w in self.winfo_children():
            for c in w.winfo_children():
                try:
                    c.configure(bg=bg)
                except tk.TclError:
                    pass

    def set_active(self, active: bool):
        self._active = active
        bg = "#172554" if active else NAVY
        self.dot.configure(bg=ACCENT if active else NAVY)
        self.configure(bg=bg)
        for w in self.winfo_children():
            try:
                w.configure(bg=bg)
            except tk.TclError:
                pass
            for c in w.winfo_children():
                try:
                    c.configure(bg=bg)
                except tk.TclError:
                    pass
        self.title.configure(fg="white" if active else "#e2e8f0")
        self.sub.configure(fg="#93c5fd" if active else "#64748b")
