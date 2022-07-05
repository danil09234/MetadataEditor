# -*- mode: python ; coding: utf-8 -*-
import sys
import os

from kivy_deps import sdl2, glew
from kivymd import hooks_path as kivymd_hooks_path


block_cipher = pyi_crypto.PyiBlockCipher(key=os.environ.get("GUI_PYINSTALLER_KEY"))
path = os.path.abspath(".")


a = Analysis(
    ['gui.py'],
    pathex=[path],
    binaries=[],
    datas=[
    ('antismirnova.kv', '.'),
    ('kivy_gui_functions.py', '.'),
    ('default_word_app.xml', '.'),
    ('default_word_core.xml', '.'),
    ('images', 'images')
    ],
    hiddenimports=[],
    hookspath=[kivymd_hooks_path],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    *[Tree(p) for p in (sdl2.dep_bins + glew.dep_bins)],
    name='Antismirnova',
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
    icon='images\\app_icon.ico',
)
