# -*- mode: python ; coding: utf-8 -*-
# NADT — Network Automation Deployment Tool
# 基于 DHCP option 66/67 的交换机批量自动化开局（onefile, 无控制台）

a = Analysis(
    ['nadt.py'],
    pathex=[],
    binaries=[],
    datas=[('templates/access-switch.cfg', 'templates')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='NADT - Network Automation Deployment Tool',
    icon='assets/nadt.ico',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
