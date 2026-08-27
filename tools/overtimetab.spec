# -*- mode: python ; coding: utf-8 -*-
"""
Рецепт сборки OVERTIMETAB.

Простыми словами: это инструкция для программы-сборщика
«какие файлы положить в папку с готовой программой».

Раньше в GitHub Actions было сказано «положи весь PySide6».
Из-за этого в zip уезжал браузер, 3D и прочий конструктор Qt,
которым табель не пользуется.

Теперь:
  1) сборщик смотрит Main.py и сам понимает, что нужно;
  2) плюс мы явно называем куски Qt, которые грузятся из QML
     (их в Python-импортах не видно);
  3) потом выкидываем всё лишнее по списку из slim_pyside.py.

Запускается так (из корня репозитория, уже после make_resources
и make_icon):

    pyinstaller --noconfirm --clean tools/overtimetab.spec
"""
import sys
from pathlib import Path

# SPECPATH — папка, где лежит этот файл (tools/). Корень репозитория на уровень выше.
ROOT = Path(SPECPATH).resolve().parent
sys.path.insert(0, str(ROOT / "tools"))

from slim_pyside import (  # noqa: E402
    UNUSED_PYSIDE_MODULES,
    UNUSED_STDLIB,
    filter_toc,
    format_mb,
    is_unused_hiddenimport,
    print_report,
    slim_dist_tree,
    toc_bytes,
)

# ---------------------------------------------------------------------------
# Что Python-код сам не импортирует, но программа всё равно откроет
# в готовом exe. Без этого списка окно может не подняться.
# ---------------------------------------------------------------------------
HIDDENIMPORTS = [
    "resources_rc",            # картинки/QML/шрифты, зашитые внутрь exe
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",       # иконка у часов и меню по правому клику
    "PySide6.QtQml",
    "PySide6.QtQuick",
    "PySide6.QtQuickControls2",
    "PySide6.QtNetwork",       # «не запускай программу дважды»
    "PySide6.QtSvg",           # иконки в папке icons/*.svg
    "PySide6.QtOpenGL",        # рисует современный интерфейс
    "openpyxl",                # выгрузка табеля в Excel
    "win32print",              # список принтеров
    "win32com",
    "win32com.client",         # печать через установленный Excel
    "pythoncom",
    "pywintypes",
    "app_update",
]

EXCLUDES = list(UNUSED_STDLIB) + list(UNUSED_PYSIDE_MODULES)

# Не ставим --collect-all PySide6. Обычных хуков PyInstaller + списка
# выше хватает, чтобы подтянуть QML и плагины окон. Лишнее режем ниже.


_version_json = ROOT / "version.json"
_changelog = ROOT / "CHANGELOG.md"
_datas = []
if _version_json.exists():
    _datas.append((str(_version_json), "."))
if _changelog.exists():
    _datas.append((str(_changelog), "."))

a = Analysis(
    [str(ROOT / "Main.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=_datas,
    hiddenimports=HIDDENIMPORTS,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
)

# Хуки PyInstaller могли всё равно прихватить WebEngine и прочее
# «на всякий случай» — вычищаем после анализа.
before_bin = toc_bytes(a.binaries)
before_data = toc_bytes(a.datas)

a.binaries, dropped_bin = filter_toc(a.binaries)
a.datas, dropped_data = filter_toc(a.datas)
a.hiddenimports = [h for h in a.hiddenimports if not is_unused_hiddenimport(h)]

print_report(dropped_bin, "библиотеки Qt (.dll)")
print_report(dropped_data, "данные Qt (QML, переводы, плагины)")
print(
    "[slim] осталось в сборке: "
    f"библиотеки {format_mb(toc_bytes(a.binaries))} "
    f"(было {format_mb(before_bin)}), "
    f"данные {format_mb(toc_bytes(a.datas))} "
    f"(было {format_mb(before_data)})"
)

pyz = PYZ(a.pure, a.zipped_data)

icon_path = ROOT / "app_icon.ico"

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="OVERTIMETAB",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # сжатие UPX + Qt часто ругает антивирус и роняет запуск
    console=False,      # без чёрного окна консоли
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(icon_path) if icon_path.exists() else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="OVERTIMETAB",
)

# На случай, если хуки PyInstaller всё-таки положили WebEngine в dist.
slim_dist_tree(ROOT / "dist" / "OVERTIMETAB")
