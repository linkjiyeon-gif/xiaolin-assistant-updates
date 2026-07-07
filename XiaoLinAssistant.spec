# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_dynamic_libs
from PyInstaller.utils.hooks import collect_all

datas = [('assets', 'assets'), ('tools', 'tools'), ('version.json', '.')]
binaries = []
hiddenimports = ['pystray', 'pystray._win32', 'pydivert', 'pynput.keyboard._win32', 'pynput.mouse._win32', 'docx', 'openpyxl', 'PyPDF2', 'yaml', 'app.adb_tools', 'app.apk_info', 'app.log_monitor', 'app.crash_analyzer', 'app.localization_checker', 'app.localization_sheet_parser', 'app.config_validator', 'app.config_rule_templates', 'app.config_quick_rules', 'app.admin_utils', 'app.ios_log_tools', 'app.time_tools', 'app.value_config_compare', 'app.update_manager']
binaries += collect_dynamic_libs('pydivert')
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pystray')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['xiaoxin_assistant.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='XiaoLinAssistant',
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
    icon=['assets\\app.ico'],
)
