# -*- mode: python ; coding: utf-8 -*-
"""Сборка заглушки-установщика (tools/installer_stub.py).

Маленький onefile-exe только из стандартной библиотеки: без Qt, без
иконок — его единственная работа распаковать приклеенный zip и запустить
программу с --setup. К релизу приклеивается обычным сложением файлов
(tools/make_installer.py), поэтому пересобирается каждый раз быстро.

Запуск из корня репозитория:
    pyinstaller --noconfirm --clean tools/installer_stub.spec
Результат: dist_stub/installer_stub.exe
"""
from pathlib import Path

ROOT = Path(SPECPATH).resolve().parent

a = Analysis(
    [str(ROOT / "tools" / "installer_stub.py")],
    pathex=[str(ROOT / "tools")],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "PySide6", "PyQt6", "PyQt5", "numpy", "unittest",
              "pydoc_data"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="installer_stub",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,          # никаких чёрных окон при двойном клике
    disable_windowed_traceback=False,
)
