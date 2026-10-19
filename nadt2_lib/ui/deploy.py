"""部署页：TFTP 服务器、DHCP option 66/67 配置、USB ZTP 包。"""

from __future__ import annotations

import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import nadt
from .. import services
from . import theme
from .widgets import DarkText, LogPanel, PageHeader, Tooltip


class DeployPage(ttk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.project = app.project
        self.server: nadt.TFTPServer | None = None
        self._last_ip = ""
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        header = PageHeader(
            self, "部署",
            "TFTP 下发 · DHCP option 66/67 · EasyDeploy U 盘包",
            "Ctrl+2 切换工作区")
        header.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 0))
        nb = ttk.Notebook(self)
        nb.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.tab_tftp = ttk.Frame(nb)
        self.tab_dhcp = ttk.Frame(nb)
        self.tab_usb = ttk.Frame(nb)
        nb.add(self.tab_tftp, text="TFTP 服务器")
        nb.add(self.tab_dhcp, text="DHCP option 66/67")
        nb.add(self.tab_usb, text="U 盘 ZTP")
        self._build_tftp()
        self._build_dhcp()
        self._build_usb()

    # ── TFTP ──
    def _build_tftp(self):
        bar = ttk.LabelFrame(self.tab_tftp, text="TFTP 服务器（asyncio，多并发）",
                             padding=(12, 8, 12, 9))
        bar.pack(fill=tk.X, padx=10, pady=(10, 6))
        ttk.Label(bar, text="根目录:").pack(side=tk.LEFT)
        self.var_root = tk.StringVar()
        ttk.Entry(bar, textvariable=self.var_root).pack(side=tk.LEFT, fill=tk.X,
                                                       expand=True, padx=(4, 6))
        ttk.Button(bar, text="浏览", style="Tool.TButton",
                   command=self._browse_root).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(bar, text="端口:").pack(side=tk.LEFT)
        self.var_port = tk.StringVar(value="69")
        ttk.Entry(bar, textvariable=self.var_port, width=5).pack(side=tk.LEFT, padx=(3, 8))
        ttk.Label(bar, text="并发:").pack(side=tk.LEFT)
        self.var_conc = tk.StringVar(value="500")
        ttk.Entry(bar, textvariable=self.var_conc, width=5).pack(side=tk.LEFT, padx=(3, 10))
        self.btn_start = ttk.Button(bar, text="启动", style="GreenAccent.TButton",
                                    command=self.start_tftp)
        self.btn_start.pack(side=tk.LEFT)
        Tooltip(self.btn_start, "启动 TFTP 服务器（需要 UDP 端口权限）")
        self.btn_stop = ttk.Button(bar, text="停止", style="Danger.TButton",
                                   command=self.stop_tftp, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=(4, 0))

        info = ttk.LabelFrame(self.tab_tftp, text="状态", padding=(12, 6, 12, 8))
        info.pack(fill=tk.X, padx=10, pady=(0, 6))
        self.var_ip = tk.StringVar(value="检测中…")
        ttk.Label(info, textvariable=self.var_ip, style="Panel.TLabel",
                  font=("Consolas", 10, "bold"), foreground=theme.ACCENT_DARK).pack(anchor="w")
        row = ttk.Frame(info)
        row.pack(fill=tk.X, pady=(4, 0))
        self.var_state = tk.StringVar(value="● 已停止")
        self.lbl_state = ttk.Label(row, textvariable=self.var_state, style="Panel.TLabel",
                                   font=("Microsoft YaHei UI", 9, "bold"),
                                   foreground=theme.MUTED)
        self.lbl_state.pack(side=tk.LEFT)
        self.var_deployed = tk.StringVar(value="已下发 0 个文件")
        ttk.Label(row, textvariable=self.var_deployed, style="PanelMuted.TLabel").pack(
            side=tk.LEFT, padx=(16, 0))
        ttk.Button(row, text="生成开局清单", style="Tool.TButton",
                   command=self.export_report).pack(side=tk.RIGHT)
        ttk.Button(row, text="打开根目录", style="Tool.TButton",
                   command=lambda: self._open_dir(self.var_root.get())).pack(
            side=tk.RIGHT, padx=(0, 6))

        self.tftp_log = LogPanel(self.tab_tftp, title="传输日志", height=12,
                                 hint="客户端请求 / 下发结果（UDP）")
        self.tftp_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

    def _tftp_log(self, msg: str):
        def _append():
            try:
                self.tftp_log.log(f"[{datetime.datetime.now():%H:%M:%S}] {msg}")
                self.var_deployed.set(f"已下发 {len(self.server.deployed) if self.server else 0} 个文件")
            except Exception:
                pass
        try:
            self.after(0, _append)
        except tk.TclError:
            pass

    def start_tftp(self):
        root = self.var_root.get().strip()
        if not root or not os.path.isdir(root):
            messagebox.showwarning("目录不存在", f"TFTP 根目录不存在:\n{root}")
            return
        try:
            port = int(self.var_port.get())
            conc = max(1, int(self.var_conc.get()))
        except ValueError:
            messagebox.showerror("参数错误", "端口/并发必须是整数")
            return
        if self.server:
            self.server.stop()
        self.server = nadt.TFTPServer(root, log_callback=self._tftp_log,
                                      port=port, max_concurrent=conc)
        self.server.start()
        if self.server.running:
            self.btn_start.state(["disabled"])
            self.btn_stop.state(["!disabled"])
            self.var_state.set(f"● 运行中 · 0.0.0.0:{port}")
            self.lbl_state.configure(foreground=theme.GREEN)
            self.app.set_status(f"TFTP 运行中: 0.0.0.0:{port} root={root}")
        else:
            self.var_state.set("● 启动失败（端口被占用？）")
            self.lbl_state.configure(foreground=theme.RED)
            self.app.set_status("TFTP 启动失败")

    def stop_tftp(self):
        if self.server:
            self.server.stop()
        self.btn_start.state(["!disabled"])
        self.btn_stop.state(["disabled"])
        self.var_state.set("● 已停止")
        self.lbl_state.configure(foreground=theme.MUTED)
        self.app.set_status("TFTP 已停止")

    def _browse_root(self):
        path = filedialog.askdirectory(title="选择 TFTP 根目录",
                                       initialdir=self.var_root.get() or os.getcwd())
        if path:
            self.var_root.set(path)

    def _open_dir(self, path: str):
        try:
            os.makedirs(path, exist_ok=True)
            os.startfile(path)
        except OSError as e:
            messagebox.showerror("打开失败", str(e))

    def export_report(self):
        deployed = list(self.server.deployed) if self.server else []
        if not deployed:
            messagebox.showinfo("提示", "还没有下发记录")
            return
        lines = ["# NADT 开局完成清单",
                 f"# 生成时间: {datetime.datetime.now():%Y-%m-%d %H:%M:%S}",
                 f"# 已下发: {len(deployed)} 个文件", ""]
        for ts, fname, ip in deployed:
            host = fname[:-4] if fname.lower().endswith(".cfg") else fname
            dev = next((d for d in self.project.devices if d.get("hostname") == host), None)
            extra = f"  ESN={dev.get('esn') or '-'}" if dev else "  （清单中未找到）"
            lines.append(f"[{ts}] {host}  客户端 {ip}{extra}")
        path = os.path.join(self.var_root.get().strip() or os.getcwd(), "deployed_report.txt")
        try:
            with open(path, "w", encoding="utf-8-sig", newline="\n") as f:
                f.write("\n".join(lines) + "\n")
        except OSError as e:
            messagebox.showerror("写入失败", str(e))
            return
        self.app.set_status(f"开局清单已生成: {path}")
        self._open_dir(os.path.dirname(path))

    # ── DHCP ──
    def _build_dhcp(self):
        card = ttk.LabelFrame(self.tab_dhcp, text="DHCP 配置生成", padding=(12, 8, 12, 9))
        card.pack(fill=tk.X, padx=10, pady=(10, 6))
        row1 = ttk.Frame(card)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text="输出风格:").pack(side=tk.LEFT)
        self.var_style = tk.StringVar()
        styles = list(nadt.DHCP_STYLES.values())
        ttk.Combobox(row1, textvariable=self.var_style, state="readonly",
                     values=styles, width=34).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Label(row1, text="TFTP IP:").pack(side=tk.LEFT)
        self.var_dhcp_ip = tk.StringVar()
        ttk.Entry(row1, textvariable=self.var_dhcp_ip, width=16).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Label(row1, text="管理VLAN:").pack(side=tk.LEFT)
        self.var_vlan = tk.StringVar(value="100")
        ttk.Entry(row1, textvariable=self.var_vlan, width=6).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Label(row1, text="引导文件:").pack(side=tk.LEFT)
        self.var_bootfile = tk.StringVar(value="lswnet.cfg")
        ttk.Entry(row1, textvariable=self.var_bootfile, width=16).pack(side=tk.LEFT, padx=(4, 0))

        row2 = ttk.Frame(card)
        row2.pack(fill=tk.X, pady=(8, 0))
        labels = [("网段", "var_net", "192.168.100.0", 14),
                  ("网关", "var_gw", "192.168.100.254", 14),
                  ("池起始", "var_r1", "", 13), ("池结束", "var_r2", "", 13)]
        for text, attr, default, width in labels:
            ttk.Label(row2, text=text, style="Muted.TLabel").pack(side=tk.LEFT, padx=(0, 3))
            var = tk.StringVar(value=default)
            setattr(self, attr, var)
            ttk.Entry(row2, textvariable=var, width=width).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(row2, text="URL(vrp146):", style="Muted.TLabel").pack(side=tk.LEFT)
        self.var_url = tk.StringVar(value="sftp://user:pass@x.x.x.x:21")
        ttk.Entry(row2, textvariable=self.var_url, width=28).pack(side=tk.LEFT)
        ttk.Button(card, text="生成 DHCP 配置", style="Accent.TButton",
                   command=self.generate_dhcp).pack(anchor="e", pady=(8, 0))

        bar = ttk.Frame(self.tab_dhcp)
        bar.pack(fill=tk.X, padx=10, pady=(0, 4))
        ttk.Button(bar, text="复制", style="Tool.TButton",
                   command=self.copy_dhcp).pack(side=tk.RIGHT)
        ttk.Button(bar, text="另存为…", style="Tool.TButton",
                   command=self.save_dhcp).pack(side=tk.RIGHT, padx=(0, 6))
        self.dhcp_text = DarkText(self.tab_dhcp)
        self.dhcp_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        self._dhcp_raw = ""

    def generate_dhcp(self):
        style_label = self.var_style.get()
        key = next((k for k, v in nadt.DHCP_STYLES.items() if v == style_label), "isc")
        ip = self.var_dhcp_ip.get().strip() or "x.x.x.x"
        devices = self.project.target_devices()
        net = self.var_net.get().strip()
        gw = self.var_gw.get().strip()
        r1 = self.var_r1.get().strip()
        r2 = self.var_r2.get().strip()
        boot = self.var_bootfile.get().strip() or self.project.settings.bootfile
        try:
            if key == "isc":
                text = nadt.build_dhcp_isc(devices, ip, boot, network=net, gateway=gw,
                                           range_start=r1, range_end=r2)
            elif key == "dnsmasq":
                text = nadt.build_dhcp_dnsmasq(devices, ip, boot, network=net, gateway=gw,
                                               range_start=r1, range_end=r2)
            elif key == "vrp146":
                text = nadt.dhcp_snippet(ip, boot, self.var_vlan.get().strip() or "100",
                                         net or "192.168.100.0", gw or "192.168.100.254",
                                         mode="146", fileserver_url=self.var_url.get().strip())
            else:
                text = nadt.dhcp_snippet(ip, boot, self.var_vlan.get().strip() or "100",
                                         net or "192.168.100.0", gw or "192.168.100.254",
                                         mode="67")
        except Exception as e:
            messagebox.showerror("生成失败", str(e))
            return
        self._dhcp_raw = text
        self.dhcp_text.set(services.mask_secrets(text, self.project))
        self.app.set_status(f"DHCP 配置已生成: {style_label}")

    def copy_dhcp(self):
        if not self._dhcp_raw:
            return
        self.clipboard_clear()
        self.clipboard_append(self._dhcp_raw)
        self.app.set_status("DHCP 配置已复制（含真实值）")

    def save_dhcp(self):
        if not self._dhcp_raw:
            messagebox.showinfo("提示", "请先生成 DHCP 配置")
            return
        path = filedialog.asksaveasfilename(
            title="保存 DHCP 配置", defaultextension=".conf",
            filetypes=[("配置文件", "*.conf *.cfg *.txt"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8", newline="") as f:
                f.write(self._dhcp_raw)
        except OSError as e:
            messagebox.showerror("保存失败", str(e))
            return
        self.app.set_status(f"DHCP 配置已保存: {path}")

    # ── USB ──
    def _build_usb(self):
        card = ttk.LabelFrame(self.tab_usb, text="EasyDeploy U 盘 ZTP 包", padding=(12, 8, 12, 9))
        card.pack(fill=tk.X, padx=10, pady=(10, 6))
        self.var_usb_server = tk.StringVar(value="file:/usb:")
        self.var_usb_sw = tk.StringVar()
        self.var_usb_pat = tk.StringVar()
        fields = [("FILESERVER", self.var_usb_server), ("SYSTEM-SOFTWARE(.cc)", self.var_usb_sw),
                  ("SYSTEM-PAT(.pat)", self.var_usb_pat)]
        for i, (label, var) in enumerate(fields):
            r, c = divmod(i, 2)
            ttk.Label(card, text=label).grid(row=r, column=c * 2, sticky="w",
                                             padx=(0, 6), pady=4)
            ttk.Entry(card, textvariable=var).grid(row=r, column=c * 2 + 1, sticky="ew",
                                                   padx=(0, 16), pady=4)
        card.columnconfigure(1, weight=1)
        card.columnconfigure(3, weight=1)
        row = ttk.Frame(self.tab_usb)
        row.pack(fill=tk.X, padx=10, pady=(0, 6))
        ttk.Button(row, text="生成 U 盘 ZTP 包", style="Accent.TButton",
                   command=self.generate_usb).pack(side=tk.LEFT)
        ttk.Label(row, text="生成目录 = 最近一次生成目录 / usb",
                  style="Muted.TLabel").pack(side=tk.LEFT, padx=(10, 0))
        self.usb_log = LogPanel(self.tab_usb, title="USB 包日志", height=8)
        self.usb_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

    def generate_usb(self):
        out_dir = self._last_output_dir()
        if not out_dir:
            messagebox.showwarning("提示", "请先在方案页生成配置")
            return
        usb_dir = os.path.join(out_dir, "usb")
        os.makedirs(usb_dir, exist_ok=True)
        devices = self.project.target_devices()
        ini = nadt.build_usb_ini(devices, self.var_usb_server.get().strip() or "file:/usb:",
                                 self.var_usb_sw.get().strip(),
                                 self.var_usb_pat.get().strip())
        with open(os.path.join(usb_dir, "usb_config.ini"), "w",
                  encoding="utf-8", newline="\r\n") as f:
            f.write(ini)
        copied = skipped = 0
        for dev in devices:
            host = (dev.get("hostname") or "").strip()
            if not host:
                continue
            src = os.path.join(out_dir, host + ".cfg")
            if os.path.isfile(src):
                import shutil
                shutil.copy2(src, os.path.join(usb_dir, host + ".cfg"))
                copied += 1
            else:
                skipped += 1
        self.usb_log.log(f"✓ usb_config.ini（{ini.count('[DEVICE')} 个设备块）")
        self.usb_log.log(f"✓ 已复制 {copied} 个 cfg" +
                         (f"；{skipped} 台未找到 cfg" if skipped else ""))
        self.usb_log.log(f"目录: {usb_dir}")
        self.app.set_status(f"U盘 ZTP 包已生成: {copied} 个配置")
        self._open_dir(usb_dir)

    # ── 生命周期 ──
    def _last_output_dir(self) -> str:
        last = self.project.last_generation.get("out_dir") if self.project.last_generation else ""
        if last and os.path.isdir(last):
            return last
        return ""

    def set_output_dir(self, path: str):
        self.var_root.set(path)
        self.var_dhcp_ip.set(self._first_ip())

    def reload_project(self):
        if self.server and self.server.running:
            self.stop_tftp()
        self.project = self.app.project
        self.var_root.set(self._last_output_dir() or
                          self.project.project_output_base(self.project.settings,
                                                           self.project.path))
        self.var_bootfile.set(self.project.settings.bootfile)
        self.var_style.set(nadt.DHCP_STYLES.get(self.project.settings.dhcp_style,
                                                nadt.DHCP_STYLES["isc"]))
        self.set_output_dir(self.var_root.get())

    def on_show(self):
        self.var_ip.set("本机 IPv4: " + "  ".join(nadt.get_local_ips())
                        if nadt.get_local_ips() else "未检测到 IPv4 地址")
        if not self.var_dhcp_ip.get():
            self.var_dhcp_ip.set(self._first_ip())

    def _first_ip(self) -> str:
        ips = nadt.get_local_ips()
        return ips[0] if ips else ""
