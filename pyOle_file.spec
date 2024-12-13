# -*- mode: python ; coding: utf-8 -*-

import os
import sys
import win32com

a = Analysis(
    ['pyOle.py'],
    pathex=[],
    binaries=[
        (os.path.join(r'C:\WINDOWS\system32', f'pythoncom{sys.version_info.major}{sys.version_info.minor}.dll'), '.'),
        (os.path.join(r'C:\WINDOWS\system32', f'pywintypes{sys.version_info.major}{sys.version_info.minor}.dll'), '.')
    ],
    datas=[('pyOle.py', '.')],
    hiddenimports=[
        'pythoncom',
        'win32com',
        'win32com.server',
        'win32com.server.register',
        'win32com.server.localserver',
        'pywintypes',
        'win32timezone',
        'ctypes',
        'PyPDF2',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='pyOle',
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