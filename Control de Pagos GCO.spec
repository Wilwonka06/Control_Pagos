# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('icon.ico', '.'), ('proceso_semanal.py', '.'), ('proceso_mensual.py', '.')]
binaries = []
hiddenimports = ['proceso_semanal', 'proceso_mensual', 'win32com', 'win32com.client', 'win32com.client.gencache', 'win32com.client.CLSIDToClass', 'pythoncom', 'pywintypes', 'win32timezone', 'openpyxl', 'openpyxl.styles', 'openpyxl.styles.fonts', 'openpyxl.styles.borders', 'openpyxl.styles.alignment', 'openpyxl.cell', 'openpyxl.utils', 'openpyxl.utils.dataframe', 'pandas', 'tkcalendar', 'babel.numbers']
tmp_ret = collect_all('win32com')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('tkcalendar')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['inicio_control.py'],
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
    [],
    exclude_binaries=True,
    name='Control de Pagos GCO',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Control de Pagos GCO',
)
