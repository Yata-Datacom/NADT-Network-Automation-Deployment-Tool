# NADT 2.0 User Guide

> Applies to **NADT 2.0.1**. The UI is in Chinese; this guide uses the same labels with English explanations.
> Example IPs, ESNs and credentials in screenshots are fictional.

> **Privacy:** do not commit real `.nadt` project files, device lists, screenshots or credentials to a public
> repository. The screenshots in this guide use synthetic demo data.

- [简体中文使用指南](USER_GUIDE.zh-CN.md)
- [README (English)](../README.md) · [README (中文)](../README.zh-CN.md)

---

## 1. What NADT does

NADT turns a spreadsheet-like device list plus a text template into one `.cfg` per switch. The switch can then
be provisioned without a console cable:

1. The switch boots and gets an address from DHCP.
2. DHCP option 66/67 points it at the NADT TFTP server.
3. The switch downloads **its own** config (rendered from `{{placeholders}}`) and applies it.

NADT 2.0 adds a project file, a single plan workspace, validation, per-file diff before writing, batch editing,
and Windows DPAPI credential protection.

---

## 2. Install and launch

### Option A: Windows exe

Download `NADT 2.0.exe` from
[Releases](https://github.com/Yata-Datacom/NADT-Network-Automation-Deployment-Tool/releases/latest), put it in a writable folder, and run it.
No Python installation is required.

> The exe is unsigned. Windows SmartScreen may ask for confirmation.

### Option B: from source

Requirements: **Python 3.10+**, Windows recommended.

```bash
git clone https://github.com/Yata-Datacom/NADT-Network-Automation-Deployment-Tool.git
cd NADT-Network-Automation-Deployment-Tool
python -m pip install -e ".[excel]"
python nadt2.py
```

After installation, these console entries are available:

| Command | What it starts |
| :-- | :-- |
| `nadt2` / `nadt` | NADT 2.0 UI |
| `nadt1` / `python nadt.py` | legacy 1.x UI |

`openpyxl` is optional; it is only needed for Excel import/export.

---

## 3. Your first project

1. Start NADT 2.0.
2. Click **新建** (New) in the top bar or press `Ctrl+N`.
3. Enter a project name.
4. Click **保存** (Save) or press `Ctrl+S`; choose a `.nadt` file.

A `.nadt` file is the whole plan: devices, templates, type mapping, project variables, credentials and
generation settings. It is a readable JSON file.

- Saves are atomic and keep a `.bak` copy.
- An `.autosave` recovery copy is written periodically.
- Recent projects are available under **项目 → 最近打开**.
- When opening a project with a newer autosave, NADT asks whether to restore it.

![Plan workspace](../assets/nadt2-plan.png)

---

## 4. The Plan workspace

### 4.1 Choose a template

Use the **模板** (Template) combo at the top. **保存模板** (Save template) writes the right-hand editor back
to the project.

A template is ordinary text with placeholders:

```text
#
sysname {{hostname}}
#
vlan batch {{vlan_batch}}
#
interface Vlanif{{mgmt_vlan}}
 ip address {{mgmt_ip}} {{mask}}
#
local-user svcd password irreversible-cipher {{svcd_password}}
#
return
```

The standard placeholder names are:

| Placeholder | Meaning |
| :-- | :-- |
| `{{hostname}}` | device name; also the default filename |
| `{{sysname}}` | usually the same as hostname |
| `{{esn}}` | serial number, used by `lswnet.cfg` |
| `{{mgmt_ip}}` / `{{mask}}` / `{{gateway}}` | management addressing |
| `{{mgmt_vlan}}` / `{{access_vlan}}` / `{{vlan_batch}}` | VLAN values |
| `{{type}}` | device type, useful for multi-template projects |
| `{{note}}` | free-form note |

Any other `{{name}}` becomes a template parameter. A parameter is resolved in this order (later wins):

1. standard device fields
2. project variables
3. project credentials (a credential wins over a same-named variable)
4. device extended fields

Credentials and project variables are **not** shown as editable grid columns; they are supplied globally.
Put per-device overrides in **设备详情** (Device details, `Ctrl+E`) under extended fields.

### 4.2 Fill the data grid

- Double-click a cell to edit; `Enter` commits, `Esc` cancels.
- Select a range and `Ctrl+V` to paste from Excel/TSV.
- `Ctrl+C` copies the selection.
- `F2` edits the current cell; `Delete` clears it.
- The **搜索** (Search) box filters by hostname / ESN / management IP.
- **添加** adds a device; **操作 ▾** opens device details, copy, delete and preview actions.

Columns are generated from the current template. If a template does not use `mgmt_ip`, that field is still
editable from **设备详情** and is still used by DHCP / USB features.

### 4.3 Batch operations

Open **批量 ▾** (Batch) in the device toolbar:

| Action | Use |
| :-- | :-- |
| 填充值 | fill the selected/visible cells in one column |
| 数字序列 | generate a numeric sequence with step, padding, prefix/suffix |
| 复制选中行 | duplicate rows; hostnames get `-copy`, ESN is cleared |
| 查找替换 | find/replace with optional regex and case matching |
| 导入表格 | import CSV/Excel with column mapping |
| 删除选中行 | delete selected devices |

### 4.4 Validate

Press `F6` or click **校验 F6**.

NADT checks:

- duplicate hostnames and duplicate ESNs
- placeholders with no value
- empty credentials used by the template
- filename collisions
- invalid output extension
- missing target devices for the selected scope

The result appears in the **问题** (Issues) tab. Double-click a row to jump to the device and field. If the
device is hidden by the current scope, NADT switches back to **全部设备** (All devices) automatically.

### 4.5 Preview and diff

Click **预览** (Preview) to see one device's final config. Sensitive values are masked by default.

Click **生成…** (Generate) to build a **generation plan** without writing anything:

- added / changed / unchanged / error counts
- file list on the left
- per-file unified diff on the right
- **显示敏感值** toggles credential masking

Only after **写入 N 个文件** (Write N files) does NADT write to disk. Unchanged files are not rewritten.
The result is kept in the **变更 / Diff** tab.

![Preflight diff](../assets/nadt2-preflight.png)

### 4.6 Output settings

| Setting | Meaning |
| :-- | :-- |
| 文件名 | filename rule: hostname, ESN, template variable, etc. |
| 后缀 | extension, default `cfg` |
| 换行 | CRLF (Windows, default) or LF |
| 输出目录 | output folder; relative paths resolve against the project file |
| 范围 | only the current template type, or all devices |
| 模板子目录 | write into a subfolder named after the template |
| 时间戳归档 | write into a timestamped subfolder to keep history |
| 生成 lswnet.cfg | write the Huawei EasyDeploy ESN map |

---

## 5. Project settings

Open **项目** (`Ctrl+3`).

![Project settings](../assets/nadt2-project.png)

### 5.1 Basic information

- edit the project name and click **应用**
- **保存项目** writes the `.nadt` file
- **另存为…** saves a copy

### 5.2 Template library

- **新建模板** creates a template
- **重命名** / **删除** manage templates
- **设为默认** selects the template opened by default
- **从工作簿导入…** imports templates from an xlsm/xlsx workbook
- double-click a template to edit it on the Plan page

### 5.3 Credentials and project variables

**凭据** (Credentials) are secrets such as `svcd_password` or `enable_secret`.

- Values are encrypted with **Windows DPAPI (current user)** when saving the project.
- The value is shown as dots; **显示敏感值** temporarily reveals it.
- The credential is never written to a grid column, diff, preview or log in cleartext.
- Moving the project to another Windows user/machine means DPAPI cannot decrypt the value; re-enter it there.

**项目变量** (Project variables) are non-sensitive defaults such as `ntp_server` or `site`. For one-off
overrides, edit the device in **设备详情** and add an extended field with the same name.

### 5.4 Settings

Settings include output folder, filename rule, extension, line endings, scope, lswnet, autosave and secret
masking. They are stored in the project file.

---

## 6. Deploy

Open **部署** (`Ctrl+2`).

### 6.1 TFTP server

1. Check the root directory (usually the generated output folder).
2. Keep UDP port `69` unless you have a reason to change it.
3. Click **启动** (Start). The status changes to running and the local IPv4 addresses are shown.
4. Watch the transfer log; `停止` stops the server.

Notes:

- Port 69 is privileged on some systems; run as administrator if binding fails.
- Allow the exe/Python through Windows Firewall on the TFTP network.
- The server can handle multiple concurrent requests; the default concurrency is 500.

### 6.2 DHCP option 66/67

1. Enter the DHCP/TFTP server IP.
2. Choose the DHCP style: ISC dhcpd, dnsmasq, Huawei VRP option 146, etc.
3. Click generate, then copy or save the snippet.
4. Paste it into the real DHCP server and reload it.

The displayed snippet masks credentials. **复制** (Copy) copies the real values to the clipboard; handle it
accordingly.

### 6.3 EasyDeploy USB package

1. Generate configs first.
2. Open **U 盘 ZTP**.
3. Choose the USB root (or use the default `output/usb`).
4. Fill system software/pattern if needed.
5. Click generate.

NADT writes `usb_config.ini` and copies the `.cfg` files into the USB folder.

---

## 7. Import a legacy xlsm workbook

Existing 1.x / macro workbooks can be imported in one click:

1. **导入 ▾ → 导入旧 xlsm 工作簿…** in the device toolbar, or **项目 → 导入旧 xlsm 工作簿…**
2. Choose the workbook.
3. NADT recognizes template sheets (sheet name = type, A1 title or column A with `{{placeholders}}`) and
   imports templates, devices, parameters and the `lswnet` ESN map.
4. Review **校验 F6**, then generate.

**导出 xlsm 兼容工作簿…** creates a workbook that follows the old layout, useful for teams that still
exchange spreadsheets.

---

## 8. Keyboard shortcuts

| Shortcut | Action |
| :-- | :-- |
| `Ctrl+N` / `Ctrl+O` / `Ctrl+S` | new / open / save project |
| `Ctrl+Shift+S` | save as |
| `Ctrl+1` / `Ctrl+2` / `Ctrl+3` | Plan / Deploy / Project |
| `Ctrl+Z` / `Ctrl+Y` | undo / redo |
| `F5` | generate (preflight) |
| `F6` | validate |
| `Ctrl+Enter` | preview selected device |
| `Ctrl+E` | device details (all fields) |
| `Ctrl+F` | focus device search |
| `Ctrl+H` | find & replace |
| `F2` | edit cell |
| `Ctrl+C` / `Ctrl+V` | copy / paste selection |
| `Delete` | clear cell |

---

## 9. Troubleshooting

| Symptom | Check |
| :-- | :-- |
| Cannot bind TFTP port 69 | run as administrator; check another TFTP service; allow firewall |
| Switch does not download config | DHCP option 66/67, TFTP IP reachability, filename and root directory |
| Credential is empty after moving project | DPAPI is user/machine-bound; re-enter the credential |
| `模板缺少变量` | add the variable to the device, project variables, or credentials; press 刷新参数列 |
| No devices generated | check the **范围** scope and the template type mapping |
| Output path wrong | relative paths resolve against the `.nadt` file; use an absolute path if needed |
| Excel import fails | install `openpyxl`: `pip install -e ".[excel]"` |
| UI looks cramped | minimum window size is 1160×740; maximize for more table columns |

---

## 10. Files, backups and security

| File | Purpose |
| :-- | :-- |
| `*.nadt` | project file (JSON) |
| `*.nadt.bak` | previous version on save |
| `*.nadt.autosave` | periodic recovery copy |
| output folder | generated `.cfg`, `lswnet.cfg`, manifest |
| `%APPDATA%\NADT2` | recent projects / app config |

- Treat the `.nadt` file as sensitive: device names, IPs, notes and template text are plaintext JSON.
- Credentials are DPAPI-protected; they are not portable across Windows users/machines by design.
- Review generated configs and test in a lab before production.
- This is a **preview** release; see the warnings in the README.

---

## 11. Template language reference

NADT uses a deliberately small, line-oriented template language. The output is plain text; NADT does not parse
the vendor CLI.

### 11.1 Variables

```text
sysname {{hostname}}
ip address {{mgmt_ip}} {{mask}}
```

- A variable is replaced with the resolved value.
- Variable names are case-sensitive; keep them consistent between the template and the grid.
- If any `{{variable}}` remains after rendering, generation stops with `模板缺少变量: …`. NADT never writes a
  literal `{{...}}` to a switch.
- A variable can be supplied by a standard device field, an extended field, a project variable, or a credential.
  See section 12 for precedence.

### 11.2 Conditional blocks

```text
{{#if stack_member_id}}
stack member {{stack_member_id}}
{{#endif}}
```

Supported forms:

| Form | Meaning |
| :-- | :-- |
| `{{#if var}}` | render if the value is non-empty |
| `{{#if !var}}` | render if the value is empty |
| `{{#if var=value}}` | render if the value equals `value` |
| `{{#else}}` | alternative branch |
| `{{#endif}}` | end the block |

Directives must be on their own line, or the whole block must be written inline:

```text
{{#if ntp_server}}ntp-service unicast-server {{ntp_server}}{{#else}}undo ntp-service unicast-server{{#endif}}
```

Inline blocks cannot be nested. A directive line with trailing text raises an error instead of silently
dropping the text.

### 11.3 Numeric loops

```text
{{#for n from 1 to {{port_count}}}}
interface GigabitEthernet1/0/{{n}}
 port link-type access
 port default vlan {{access_vlan}}
{{#endfor}}
```

- `from` and `to` accept integers or variables.
- The loop variable is available as `{{n}}` inside the block.
- Inline form: `{{#for n from 1 to 8}}...{{#endfor}}`.
- Loops are intended for port ranges and similar text expansion, not for complex logic.

### 11.4 Whitespace and line endings

- The template is rendered line by line; trailing newlines are preserved.
- The **换行** setting converts output to CRLF (default) or LF.
- Use the preview/diff before writing, especially when a template mixes Windows and Unix conventions.

---

## 12. Fields, aliases and precedence

### 12.1 Standard fields

| Internal field | Common aliases (CSV/Excel) | Used for |
| :-- | :-- | :-- |
| `esn` | `sn`, `serial`, `序列号`, `设备序列号` | `lswnet.cfg`, USB matching, validation |
| `hostname` | `设备名`, `名称`, `主机名`, `设备名称` | filename, `{{hostname}}`, `{{sysname}}` fallback |
| `type` | `型号`, `类型`, `设备类型` | template type mapping / scope |
| `mgmt_ip` | `ip`, `管理IP`, `IP地址`, `管理地址` | `{{mgmt_ip}}`, DHCP/USB helpers |
| `mgmt_vlan` | `vlanif`, `管理VLAN` | `{{mgmt_vlan}}`, `vlan_batch` derivation |
| `access_vlan` | `access`, `接入VLAN`, `用户VLAN` | `{{access_vlan}}`, `vlan_batch` derivation |
| `mask` | `掩码`, `子网掩码` | `{{mask}}`; defaults to `255.255.255.0` |
| `gateway` | `wg`, `网关`, `默认网关` | `{{gateway}}`; defaults to `.254` when derivable |
| `note` | `备注`, `说明`, `位置` | free-form note; not generated into config automatically |

Derived values:

- `vlan_batch` = `mgmt_vlan + " " + access_vlan` unless an extended field overrides it.
- `sysname` = extended `sysname` if present, otherwise `hostname`.
- `gateway` = explicit gateway, otherwise the management IP with the last octet replaced by `254` (when the
  management IP is a valid IPv4 address).
- Any empty `mask` becomes `255.255.255.0`.

### 12.2 Extended fields

Extended fields are arbitrary key/value pairs attached to one device:

- edit them in **设备详情** (`Ctrl+E`) as `key: value` lines;
- import them from extra CSV/Excel columns;
- reference them as `{{key}}` in the template.

Examples:

```text
stack_member_id: 1
port_count: 24
system_software: SWITCH.cc
system_pat: SWITCH.pat
```

### 12.3 Precedence

When the same name exists in multiple places, the later source wins:

1. standard device field
2. project variable
3. project credential (wins over a same-named variable)
4. device extended field

This is why one device can override a project default without duplicating the whole credential table.

### 12.4 Project variables vs. credentials

- Use **凭据** for secrets: `svcd_password`, `enable_secret`, SNMP secrets.
- Use **项目变量** for non-secrets: `site`, `ntp_server`, `dns_server`, `syslog_server`.
- Credentials are not shown as grid columns and are masked in preview/diff/logs.
- If a template references a credential that is empty, validation emits a warning; if it references a name
  that exists nowhere, rendering emits an error.

---

## 13. Worked example

### 13.1 Template

Create `access-demo.cfg` in **项目 → 模板库 → 新建模板**, then paste:

```text
#
sysname {{hostname}}
#
vlan batch {{vlan_batch}}
#
interface Vlanif{{mgmt_vlan}}
 ip address {{mgmt_ip}} {{mask}}
#
{{#if stack_member_id}}
stack member {{stack_member_id}}
{{#endif}}
#
{{#for n from 1 to {{port_count}}}}
interface GigabitEthernet1/0/{{n}}
 port link-type access
 port default vlan {{access_vlan}}
{{#endfor}}
#
local-user demo password irreversible-cipher {{svcd_password}}
#
ntp-service unicast-server {{ntp_server}}
#
return
```

### 13.2 Device data

| hostname | type | esn | mgmt_ip | mgmt_vlan | access_vlan | mask | port_count | stack_member_id |
| :-- | :-- | :-- | :-- | :-- | :-- | :-- | :-- | :-- |
| SW-DEMO-01 | SW | DEMO-ESN-0001 | 192.0.2.11 | 10 | 110 | 255.255.255.0 | 24 | |
| SW-DEMO-02 | SW | DEMO-ESN-0002 | 192.0.2.12 | 10 | 110 | 255.255.255.0 | 24 | |
| SW-DEMO-03 | SW | DEMO-ESN-0003 | 198.51.100.11 | 20 | 120 | 255.255.255.0 | 48 | 1 |

`port_count` and `stack_member_id` are extended fields.

### 13.3 Project-level values

| Type | Name | Example value |
| :-- | :-- | :-- |
| credential | `svcd_password` | a password entered in the credential dialog |
| variable | `ntp_server` | `192.0.2.1` |

### 13.4 Run

1. Select `access-demo.cfg` on **方案**.
2. Set **范围** to **仅当前模板类型**.
3. Set **输出目录** to `C:\NADT-Demo\output`.
4. Press `F6`; fix all errors.
5. Press `F5`; review the diff for `SW-DEMO-01.cfg`.
6. Click **写入 N 个文件**.
7. Start TFTP with the generated folder as root.

Expected files:

```text
C:\NADT-Demo\output\access-demo\
  SW-DEMO-01.cfg
  SW-DEMO-02.cfg
  SW-DEMO-03.cfg
  lswnet.cfg          (if 生成 lswnet.cfg is enabled)
```

---

## 14. Validation and generation reference

### 14.1 Issue severities

| Severity | Meaning | Blocks generation? |
| :-- | :-- | :-- |
| Error | duplicate hostname, missing variable, invalid filename, empty template | yes by default |
| Warning | duplicate ESN, empty credential, no devices, empty project variable | no, but review |
| Info | missing ESN, plaintext credential fallback | no |

### 14.2 Common checks

| Check | What to do |
| :-- | :-- |
| `主机名重复` | make hostnames unique; hostnames also become filenames |
| `ESN 重复` | fix the serial number; duplicate ESNs break mapping tables |
| `模板缺少变量` | add the field to the device, project variable, or credential; use 刷新参数列 |
| `凭据 … 尚未填写` | fill the credential in **项目 → 凭据** |
| `文件名重复` | change the filename rule or make the selected field unique |
| `后缀不合法` | use letters/digits, e.g. `cfg` |
| `当前范围没有设备` | switch **范围** to 全部设备 or fix the type mapping |
| `模板为空` | paste the template text and click 保存模板 |

Double-click an issue to jump to the row and column. If the device is hidden by the scope, NADT switches to
全部设备.

### 14.3 Generation statuses

| Status | Meaning | Written? |
| :-- | :-- | :-- |
| `新增` | file does not exist yet | yes |
| `修改` | file exists and content differs | yes |
| `不变` | file exists and content is identical | no |
| `错误` | rendering or filename error for that device | no |

The manifest under the `.nadt` file records the last run; the diff itself is computed against the actual files
on disk, not just the manifest.

### 14.4 Single-file mode

Select **单文件（合并）** as the filename rule to render every target device into one file (named after the
template). This is useful for a combined “all devices” appendix, but not for DHCP per-device boot files.
`lswnet.cfg` is not generated in this mode.

---

## 15. Deployment recipes

### 15.1 Network prerequisites

- The switch and the TFTP server must be able to reach each other.
- The DHCP server must be able to deliver option 66/67 (or the vendor-specific option) to the switch.
- UDP 69 must reach the machine running the TFTP server.
- If the switch is in a different VLAN, configure DHCP relay/helper on the gateway.

### 15.2 Generic / Cisco-style flow

1. Generate one `.cfg` per device.
2. Start TFTP with the output folder as root.
3. Configure DHCP option 66 = TFTP server IP, option 67 = `<hostname>.cfg` (or use per-host reservations
   with `filename`).
4. Boot the switch; confirm the TFTP log shows the request.

The exact DHCP syntax depends on your server. Treat the generated snippet as a template; check it before
reloading the production DHCP service.

### 15.3 Huawei EasyDeploy flow

1. Generate `.cfg` files with **生成 lswnet.cfg** enabled.
2. Put `lswnet.cfg` and all device `.cfg` files in the TFTP root.
3. Use a DHCP style with option 146 (`vrp146`) or option 67, depending on the design.
4. Verify that the ESN in `lswnet.cfg` matches the switch (`display esn`).
5. Boot the switch and watch the TFTP log.

### 15.4 USB ZTP

1. Generate configs.
2. Open **部署 → U 盘 ZTP**.
3. Generate the package; the default output is `<output>/usb/`.
4. Copy `usb_config.ini` and every `.cfg` to the USB root.
5. Insert the USB stick and boot the switch.

`SYSTEM-SOFTWARE` / `SYSTEM-PAT` can come from device extended fields or the global defaults. Stack members can
use `stack_member_id`.

---

## 16. Scale, backup and migration

### 16.1 Large projects

- The grid is virtualized by page size in the sense that only visible rows are edited directly; thousands of
  rows remain scrollable.
- Type scope lets you generate one template at a time.
- Auto-validation after loading is skipped for very large projects; press `F6` explicitly.
- Use batch fill/sequence instead of editing repeated values one by one.
- Validate output filenames before generating hundreds of files.

### 16.2 Backup checklist

Before a delivery cutover:

- save the `.nadt` file;
- copy the generated output folder (including `lswnet.cfg` if used);
- export an xlsm-compatible workbook if the team still uses Excel;
- record the version (`v2.0.1`) and the generation folder;
- store credentials in a password manager, not in a public repository.

### 16.3 Moving to another machine

1. Copy the `.nadt` file and its output folder.
2. Open the project; re-enter credentials because DPAPI does not transfer across users/machines.
3. Check **项目 → 设置** and adjust the output path if the folder moved.
4. Run `F6` and a preview before generating.

---

## 17. FAQ

**Do I need Python?** No, if you use the exe. Python is only needed for source runs and Excel support.

**Where is the 1.x UI?** Run `python nadt.py` or `nadt1`. The default `nadt`/`nadt2` commands start 2.0.

**Why did my template text disappear?** The editor is saved when you press **保存模板**, `Ctrl+S`, or generate.
If you switch templates with unsaved edits, NADT asks first.

**Why are no files generated?** Check the scope, type mapping, template, and validation. `不变` files are
intentionally not rewritten.

**Why does the diff show a password as dots?** Credential masking is on. Use **显示敏感值** only on a trusted
screen.

**Can I use a `.nadt` file in Git?** It is readable JSON; credentials are DPAPI-protected, but device names and
IPs are not. Keep it in a private repository or outside version control.

**Can I have multiple templates per project?** Yes. Add templates in the project library, set a default, and
use type scope. Devices with an empty `type` are included in the current type scope.

**Why does `lswnet.cfg` skip a device?** The device has no ESN. Add the serial number.

---

## Appendix A. Sample CSV

```csv
esn,hostname,type,mgmt_ip,mgmt_vlan,access_vlan,mask,gateway,note,port_count,stack_member_id
DEMO-ESN-0001,SW-DEMO-01,SW,192.0.2.11,10,110,255.255.255.0,192.0.2.254,示例接入交换机,24,
DEMO-ESN-0002,SW-DEMO-02,SW,192.0.2.12,10,110,255.255.255.0,192.0.2.254,示例接入交换机,24,
DEMO-ESN-0003,SW-DEMO-03,SW,198.51.100.11,20,120,255.255.255.0,198.51.100.254,示例接入交换机,48,1
```

## Appendix B. Glossary

| Term | Meaning |
| :-- | :-- |
| `.nadt` | project file: device data, templates, settings, credentials |
| Plan | the main data/template/generate workspace |
| Scope | which devices are targeted (all vs. current template type) |
| Preflight | generation plan with per-file diff before writing |
| DPAPI | Windows Data Protection API, used for credential encryption |
| lswnet | Huawei EasyDeploy ESN mapping file |
| ZTP | zero-touch provisioning |
| ESN | equipment serial number |
