# NADT — Network Automation Deployment Tool（网络自动化交付工具）

[![CI](https://github.com/Yata-Datacom/NADT-Network-Automation-Deployment-Tool/actions/workflows/ci.yml/badge.svg)](https://github.com/Yata-Datacom/NADT-Network-Automation-Deployment-Tool/actions/workflows/ci.yml)
[![Build EXE](https://github.com/Yata-Datacom/NADT-Network-Automation-Deployment-Tool/actions/workflows/build.yml/badge.svg)](https://github.com/Yata-Datacom/NADT-Network-Automation-Deployment-Tool/actions/workflows/build.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](pyproject.toml)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](#)

**v1.12.2** · [CHANGELOG](CHANGELOG.md) · [开发与测试](#开发与测试--development--tests)

> ⚠️ **预览版 / PREVIEW** —— 本工具仍在持续完善中（部分 UI 与文档待打磨、少数场景未覆盖），
> 先放出来供试用与参考。**当前版本 v1.12（预览版）**，欢迎反馈问题。
>
> 使用前请务必：① 把设备清单里的示例口令/凭据换成你自己的；② 在实验环境验证后再用于生产。


A **vendor-neutral, zero-touch provisioning** tool for network switches, built on **DHCP option 66/67 + TFTP**
(Huawei / Cisco / generic presets). A factory-fresh switch gets an address from DHCP, downloads its own config
via option 66/67, applies it, and is ready — no console cable, no manual typing. It ships with a spreadsheet-style
workbench (template text in column A, parameter names in row 1, one device per row → every `.cfg` in one click),
an asyncio **TFTP server**, per-device `.cfg` generation and an **EasyDeploy USB package** builder.

> 基于 **DHCP option 66/67 + TFTP** 的**厂商无关**交换机批量**零接触开局**（Zero-Touch Provisioning）工具（华为/思科/通用）。
> 设备出厂开机 → DHCP 自动获取 IP → 按 option 66/67 下载配置 → 自动应用 → 开局完成。
>
> **推荐入口：◎ 工作表 Sheet 页** —— Excel 宏式表格：A 列写配置模板文本、第 1 行 C 列起写参数名、
> 每行一台设备参数，双击单元格编辑，点「生成脚本」直接出全部 .cfg。**不会用宏也能上手**。

**⬇️ 下载 / Downloads:** [Releases](https://github.com/Yata-Datacom/NADT-Network-Automation-Deployment-Tool/releases)
（Windows 单文件 exe，免安装、无需 Python）

---

## English

### What it does

NADT mass-provisions switches over the network, without ever touching a console:

1. A brand-new switch boots and gets an address from your DHCP server.
2. The DHCP reply carries option 66/67 (`tftp-server-name` / `bootfile-name`) pointing at your TFTP server.
3. The switch downloads **its own** `.cfg` (rendered per device from your template), applies it — and the
   management IP, VLANs, accounts and services are all in place.

The core engine is **vendor-neutral**: a device list + a `{{placeholder}}` template renderer + `.cfg` generation.
Vendors are only presets:

| Preset | Bootstrap file | ESN mapping | DHCP output |
|---|---|---|---|
| **Generic** (recommended) | — | bound by MAC | ISC dhcpd / dnsmasq |
| **Huawei** | `lswnet.cfg` | ✅ EasyDeploy | Huawei VRP option 146 |
| **Cisco** | `network-confg` | — bound by MAC | ISC dhcpd / dnsmasq |

Adding a vendor = one entry in `VENDOR_PRESETS`; the engine stays untouched. The renderer is not
vendor-specific either: any text config (VRP / IOS / XR / NX-OS / …) works with `{{variables}}`.

### Requirements

- Python 3.10+ (tested on Windows)
- Core is standard-library only; `openpyxl` is optional — needed just for the Excel workbench and
  template import: `pip install -e ".[excel]"`

### Quick start

```bash
python -m pip install -e ".[excel]"
python nadt.py            # or the console entry point: nadt
```

…or simply grab the single-file exe from [Releases](https://github.com/Yata-Datacom/NADT-Network-Automation-Deployment-Tool/releases)
(Windows, no Python required).

Then walk the five tabs: **① device list → ② config template → ③ generate scripts → ④ TFTP server → ⑤ deployment notes**.
The recommended entry point is the **◎ Worksheet Sheet** tab — an Excel-macro-style grid where you put the template
text in column A, parameter names from row 1 column C, and one device per row; one click generates every `.cfg`.
No macro skills needed.

### Notes

- This is a **preview** release — please validate in a lab before production use.
- Replace the sample credentials in the device list and templates with your own first.
- In-depth docs (Excel template import, USB ZTP package, `lswnet.cfg` version files, DHCP modes,
  thousands-of-devices scale) are in the Chinese sections below.

## 🏭 厂商无关设计

NADT 核心引擎与厂商无关（设备清单 + 模板渲染 + cfg 生成），华为只是其中一个"预设"：

| 预设 | 引导文件 | ESN 映射表 | DHCP 输出 |
|---|---|---|---|
| **通用 Generic（推荐）** | 不生成 | ❌ 按 MAC 绑定 | ISC dhcpd / dnsmasq |
| **华为 Huawei** | lswnet.cfg | ✅ EasyDeploy | 华为 VRP option 146 |
| **思科 Cisco** | network-confg | ❌ 按 MAC 绑定 | ISC dhcpd / dnsmasq |

- ③ 页选预设自动填参数，**全部可手动覆盖**（引导文件名、映射表开关）
- 通用模式：DHCP 按设备 MAC 绑定下发各自 cfg（`option bootfile-name "主机名.cfg"`）——思科 auto-install 就是这样
- 华为模式：ESN 映射表 lswnet.cfg + option 146 netfile——EasyDeploy 专有玩法
- 新增厂商 = 在 `VENDOR_PRESETS` 加一项，核心引擎不用动
- 模板引擎不限厂家：任何文本配置（VRP/IOS/XR/NX-OS/…）都能用 `{{变量}}` 渲染

## 🚀 快速上手

1. **① 设备清单**：导入 CSV/Excel，或手动添加设备（ESN 序列号、主机名、类型、管理 IP）
2. **② 配置模板**：华为配置模板，`{{占位符}}` 自动替换（可编辑、可保存）
3. **③ 生成脚本**：一键生成每台设备的 `.cfg` + `lswnet.cfg`（ESN 映射表）
4. **④ TFTP 服务器**：启动，根目录指向生成目录
5. **⑤ 部署说明**：填本机 IP 生成 DHCP option 66/67 配置片段 → 贴到 DHCP 服务器

## 📋 设备清单格式（CSV，首行表头）

```csv
esn,hostname,type,mgmt_ip,mgmt_vlan,access_vlan,mask,gateway,note
[SN],SW-1F-01,ACCESS,10.10.10.1,10,20,,,1F弱电井
```

| 字段 | 说明 |
|---|---|
| esn | 设备序列号（华为 `display esn` / 机身铭牌）|
| hostname | 设备名，同时是生成的 cfg 文件名 |
| type | **自由文本**，仅用于「类型→模板映射」选择模板（如 ACCESS/CORE），不参与推导 |
| mgmt_ip | 管理 IP |
| mgmt_vlan / access_vlan | **必填**（模板用 vlan_batch 时）；留空且模板引用会报错 |
| mask / gateway | 掩码默认 255.255.255.0；网关留空自动取管理网段 .254 |

> 额外列任意加：任何列名都会变成模板 `{{变量}}`（见下节模板语言）。

**💡 批量编号**：主机名/管理IP/ESN/扩展字段支持 `{范围}` 语法，一次添加生成多台：

| 写法 | 生成 |
|---|---|
| `SW-{1-254}` | SW-1 … SW-254（254 台，按顺序）|
| `SW-{1,254}` | 仅 SW-1、SW-254（2 台）|
| `SW-{01-03}` | SW-01、SW-02、SW-03（零填充）|
| `10.10.10.{1-3}` | 10.10.10.1 … .3 |

多字段联动一一对应：`SW-{1-3}` + 管理IP `10.10.10.{1-3}` → SW-1↔.1、SW-2↔.2、SW-3↔.3。
生成超过 20 台会先确认；超过 500 台会拒绝（请分批添加）。

**🔍 搜索与分页**：设备表支持关键字搜索（ESN/主机名/IP/类型，无视大小写）和分页
（每页 100/200/500/1000，右下角翻页），几千台设备也能流畅操作。

**✅ 批量校验**：点「校验 Validate」或生成前自动检查——重复 ESN/IP/主机名、IP 格式错误、
缺必填字段（错误级别）；缺 ESN（警告级别）。有错误时生成会先弹确认，防止带着冲突数据开局。

**💾 本地设备库**：设备清单自动保存到 `devices.db`（SQLite），关闭重开自动恢复；
CSV/Excel 导入导出照常可用，`devices.db` 可随时删除（下次添加自动重建）。

**↶ 撤销（Undo）**：工具栏「↶ 撤销」或 Ctrl+Z，误删/误改/清空 20 步内可恢复。

**📦 生成自动归档**：默认每次生成存到 `输出目录/时间戳/` 子目录，保留最近 10 次自动清理；
取消勾选「自动归档」则直接输出到所选目录。

**🏁 TFTP 开局完成清单**：④ 页实时显示"已开局 N 台"；「生成开局清单」输出
`deployed_report.txt`（时间/文件名/客户端 IP + 反查设备清单的 ESN/类型）；
「重置记录」用于新一批开局前清零。

## 📥 Excel 模板导入（超越宏脚本）

支持师傅的 **Excel 宏工作流**（xlsm 模板库格式（模板 sheet + 参数列 + 每行一台设备））：
- ① 页「导入 Excel 模板」→ 选 xlsm/xlsx
- 识别规则：每个 sheet 一个模板（A1 为【模板名】或 A 列含 `{{变量}}`）：
  - **A 列第 2 行起** = 配置模板文本（`#` 分隔行保留）
  - **第 1 行 C 列起** = 参数名（`{{sysname}}`）
  - **第 2 行起每行** = 一台设备（各参数值）
- 导入后：模板进②模板库（类型=sheet名自动映射），设备进①清单（参数进扩展字段），
  直接③生成脚本即可——**不再需要 VBA 宏**
- 自动跳过 主页/模板数据/帮助说明 等非模板 sheet

**列映射导入**：「导入 CSV/Excel」会先弹列映射窗口（自动识别表头，可手动把任意列
指定为主机名/IP/类型/忽略/扩展字段），表头不规范也能导。

**生成 Excel 模板**：① 页「生成 Excel 模板」→ 保存一个 xlsx：sheet1「示例模板」
（A 列配置文本 + 参数名 + 3 台示例设备，改内容就能用），sheet2「操作方式」
（结构/步骤/变量语法/注意事项）。填好后点「导入 Excel 模板」即可导入。

⚠ 导入后建议先点「校验 Validate」：跨 sheet 同名设备、模板变量拼写错误（如原宏里
`{{vlaif1004}}` 与参数 `vlanif1004` 不一致）会明确列出来——宏时代这些错误是静默的。

**扩展字段**：每行一个 `K:V`（回车分隔），支持 `=` / `:` / `：` 三种分隔符，中文 key 也行；
任何 K 都变成模板 `{{K}}`（例如 `ntp_server:10.0.0.1` → `{{ntp_server}}`）。

## 🧩 模板占位符与模板语言

**变量占位符**：`{{hostname}}` `{{vlan_batch}}` `{{mgmt_vlan}}` `{{mgmt_ip}}` `{{mask}}` `{{access_vlan}}` `{{gateway}}`
**任意变量**：设备清单 CSV 的**额外列**、或设备"扩展字段"里的 `key=value` 都会变成模板变量
（例如列 `ntp_server` → 模板里写 `{{ntp_server}}`）。

**模板指令**（复杂模板也能批量精确）：

| 指令 | 说明 | 示例 |
|---|---|---|
| `{{#if 变量}}` ... `{{#else}}` ... `{{#endif}}` | 条件块，变量非空才输出 | 有 NTP 才加 NTP 配置 |
| `{{#if 变量=值}}` | 条件等于 | `{{#if type=CORE}}` |
| `{{#for n from 1 to N}}` ... `{{#endfor}}` | 循环，N 可以是变量 | 批量生成 24/48 个接口块 |

```cfg
# 示例：24/48 口由设备扩展字段 port_count 决定
{{#for n from 1 to {{port_count}}}}
interface GigabitEthernet0/0/{{n}}
 port link-type access
 port default vlan {{access_vlan}}
#
{{#endfor}}
```

**模板库**：`templates/` 目录放多个 `.cfg` 模板，② 页可切换/新建/删除；
**类型→模板映射**（`templates/mapping.json`）：type 字段自由定义（ACCESS / CORE / 按项目任意），
在②页把类型映射到对应模板，生成时自动按类型选模板。

模板文件：`templates/access-switch.cfg`（默认模板提取自真实园区接入交换机开局配置（**已脱敏**：示例口令请部署前修改））

## 🔧 工作原理

```
【DHCP 开局（通用）】设备开机 → DHCP(option 66=文件服务器, option 67=引导文件)
          → 下载配置 → 应用 → 开局完成
          通用/思科: DHCP 按 MAC 绑定，每台拿自己的 主机名.cfg
          华为:      下载 lswnet.cfg ESN 映射表，按序列号找自己的 cfg

【U盘 ZTP 开局（华为 EasyDeploy）】无 DHCP 也能干：
          生成 usb_config.ini（EasyDeploy 格式）→ 拷 U 盘根目录
          → 插交换机开机 → 按 ESN/MAC 匹配 [DEVICEn DESCRIPTION] 块
          → 从 U 盘(file:/usb:)加载 SYSTEM-CONFIG → 开局完成
```

## 💾 U盘 ZTP 包（EasyDeploy USB 开局）

③ 页「生成 U盘 ZTP 包」→ 输出 `usb_config.ini` + 全部 `.cfg` 到 `输出目录/usb/`：

```ini
;BEGIN DC
[GLOBAL CONFIG]
FILESERVER=file:/usb:
[DEVICE0 DESCRIPTION]
ESN=[SN]
DEVICETYPE=DEFAULT
SYSTEM-SOFTWARE=[SOFTWARE].cc   ; 可选，③页 SYSTEM-SOFTWARE 默认
SYSTEM-CONFIG=[HOSTNAME].cfg       ; 自动取设备主机名
SYSTEM-PAT=[PATCH].PAT           ; 可选
STACK-MEMBER-ID=2                                ; 堆叠：扩展字段 stack_member_id
;END DC
```

- 每台设备一个 `[DEVICEn DESCRIPTION]` 块（n 从 0 开始），按 ESN 匹配；无 ESN 的设备用 MAC（设备清单加 mac 列）
- 堆叠设备：扩展字段填 `stack_member_id=成员ID`（多成员在清单里加多行，同一配置不同 ESN/成员ID）
- 不用 DHCP 时 FILESERVER 保持 `file:/usb:`；FILESERVER 也可填 ftp/sftp/tftp 地址走网络
- usb_config.ini 用 CRLF 换行（华为要求），文件名必须叫 `usb_config.ini`，放 U 盘**根目录**

## 📦 lswnet.cfg 版本文件 / 补丁（复杂场景）

实战格式（核心交换机一起刷版本+补丁）：

```
esn=[SN];vrpfile=[SOFTWARE].cc;patchfile=[PATCH].pat;cfgfile=[HOSTNAME].cfg;
```

- ③ 页「lswnet 版本文件 vrpfile / 补丁 patchfile」填全局默认；留空则不含该字段
- 设备扩展字段 `vrp_file` / `patch_file` 可单台覆盖
- 设备清单 CSV 也可直接加 vrp_file / patch_file 列

## 🔌 DHCP 模式（⑤ 部署说明）

- **模式 146（华为 EasyDeploy 标准，默认）**：
  `option 66 ascii sftp://user:pass@ip:port` + `option 146 ascii opervalue=1;delaytime=0;netfile=lswnet.cfg;`
- **模式 67（通用 PXE）**：`option 66 ip-address IP` + `option 67 ascii lswnet.cfg`
- 华为 S 系列交换机 EasyDeploy 认的是 **option 146**；其它厂家（或通用 DHCP 服务器）用 67

## 🏗 超大规模（几千台）能力

| 能力 | 说明 |
|---|---|
| **异步 TFTP 引擎** | asyncio 单事件循环多路复用，支持上千并发传输（对比旧版每请求一线程）；文件内存缓存预热（总量超 200MB 自动改按需加载） |
| **并发上限** | TFTP 页可调（默认 500），超限拒绝并提示客户端重试，防断电恢复时的「惊群效应」压垮服务器 |
| **并行生成** | 多线程渲染 + 进度条实时显示 N/总数；失败自动汇总导出 `failures.csv` |
| **批量校验** | 生成前自动查重复 ESN/IP/主机名等，杜绝千台开局事故 |
| **设备库持久化** | SQLite 存储，重启不丢；搜索/分页支撑大清单操作 |
| **增量部署建议** | 千台场景建议：DHCP option 66 指向多台 TFTP（轮询），或按楼层/区域分段开局错峰 |

**瓶颈提示**：单机 TFTP 受网卡带宽限制（百兆口约 100 台并发饱和），千台同开建议 option 66
轮询 2-3 台服务器，或分批开局（如按楼栋不同 DHCP scope）。

## 📦 打包

```bash
python -m PyInstaller --clean --noconfirm NADT.spec
# 产物: dist/NADT - Network Automation Deployment Tool.exe
# 首次运行自动在 exe 旁生成 templates/ 和 output/
```

## ⚠️ 注意

- TFTP 端口 69 需在防火墙放行；交换机管理网段与 DHCP/TFTP 必须互通
- 仅用于授权设备的开局交付，请勿用于未授权网络
- 老项目：`C:\Users\yata\Documents\coding\批量交换机脚本生成\`（Excel 宏版）

---

## 开发与测试 / Development & Tests

```bash
# 安装（含开发依赖：pytest / ruff / openpyxl）
python -m pip install -e ".[dev]"

# 跑测试（模板引擎 / 批量编号 / 设备校验 / 清单解析 / DHCP / USB 包）
python -m pytest -q

# 静态检查
ruff check .

# 打包单文件 exe（Windows）
python -m PyInstaller --clean --noconfirm NADT.spec
```

- 装好后也可用命令入口启动：`nadt`
- 核心只用标准库；`openpyxl` 是可选依赖（设备清单 xlsx / Excel 模板导入导出），缺失时界面会提示而非崩溃
- CI：`.github/workflows/ci.yml`（Windows 跑 pytest、Linux 跑 ruff）；`.github/workflows/build.yml`（打 tag 或手动触发 → 产出 exe artifact）
- 版本号在 `pyproject.toml` 与 `nadt.__version__` **两处**，必须保持一致（有测试守着）
