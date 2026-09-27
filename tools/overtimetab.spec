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
    "recovery",               # аварийный восстановитель (первый import Main.py)
    "webapp.server",          # экспериментальный веб-режим (--web)
    "PySide6.QtWebEngineQuick",   # окно веб-версии внутри программы (этап 4)
    "PySide6.QtWebEngineCore",
    "PySide6.QtWebEngineWidgets",  # запасной вкус окна (QWebEngineView),
    "PySide6.QtPositioning",     # зависимость Qt6WebEngineCore (не вырезать!)
    "PySide6.QtPrintSupport",    # зависимость QtWebEngineWidgets (печать из веб-окна)
    "PySide6.QtQuickWidgets",    # зависимость Qt6WebEngineWidgets (веб-окно)
    "methodical_data",        # методички и производственные календари (окно справки)
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
# Веб-версия (этап переезда): статика интерфейса для режима --web
_web_static = ROOT / "webapp" / "static"
if _web_static.exists():
    _datas.append((str(_web_static), "webapp/static"))

# Встроенный веб-интерфейс (этап 4): Chromium-движок из состава PySide6.
# Кладём то, что хуки PyInstaller могут не найти сами: QML-плагин,
# процесс рендера, ресурсы и переводы. Пути определяем по факту —
# сборка идёт и на Windows (dll/exe), и локально на Linux (so).
import PySide6 as _pyside  # noqa: E402
_qt_dir = Path(_pyside.__file__).resolve().parent / "Qt"
_qml_we = _qt_dir / "qml" / "QtWebEngine"
if _qml_we.exists():
    _datas.append((str(_qml_we), "PySide6/Qt/qml/QtWebEngine"))
if (_qt_dir / "resources").exists():
    _datas.append((str(_qt_dir / "resources"), "PySide6/Qt/resources"))
if (_qt_dir / "translations").exists():
    for _tr in (_qt_dir / "translations").glob("qtwebengine*"):
        _datas.append((str(_tr), "PySide6/Qt/translations"))
_bin_extra = []
_le = _qt_dir / "libexec"
if _le.exists():
    for _f in _le.glob("QtWebEngineProcess*"):
        _bin_extra.append((str(_f), "PySide6/Qt/libexec"))
_qb = _qt_dir / "bin"
if _qb.exists():
    for _pat in ("Qt6WebEngine*", "Qt6WebChannel*"):
        for _f in _qb.glob(_pat):
            _bin_extra.append((str(_f), "PySide6/Qt/bin"))
# на части раскладок (Windows-колёса) библиотеки лежат в корне PySide6
_top = Path(_pyside.__file__).resolve().parent
for _pat in ("Qt6WebEngine*.dll", "Qt6WebChannel*.dll", "QtWebEngineProcess*.exe"):
    for _f in _top.glob(_pat):
        _bin_extra.append((str(_f), "PySide6"))

a = Analysis(
    [str(ROOT / "Main.py")],
    pathex=[str(ROOT)],
    binaries=_bin_extra,
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

# PyInstaller 6: a.zipped_data больше нет.
try:
    pyz = PYZ(a.pure, a.zipped_data)
except (TypeError, AttributeError):
    pyz = PYZ(a.pure)

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


# ═══════════════════════════════════════════════════════════════════
# ВСТРОЕННЫЙ ВЕБ-ДВИЖОК: прямой докоп в dist (мимо всех фильтров TOC).
# Фильтры уже дважды отрывали движку зависимости (Positioning, локали);
# здесь мы гарантированно кладём полный набор прямо из PySide6 сборщика.
# ═══════════════════════════════════════════════════════════════════
import shutil as _shutil

_psd = Path(_pyside.__file__).resolve().parent          # PySide6 сборщика
_app = ROOT / "dist" / "OVERTIMETAB"
_internal = _app / "_internal"                          # PyInstaller 6: onedir
_base = _internal if _internal.is_dir() else _app
_dest_ps = _base / "PySide6"

def _copy_file(srcf: Path, dstdir: Path):
    if srcf.is_file():
        dstdir.mkdir(parents=True, exist_ok=True)
        _shutil.copy2(srcf, dstdir / srcf.name)
        return 1
    return 0

def _copy_dir(srcd: Path, dstdir: Path):
    if not srcd.is_dir():
        return 0
    n = 0
    for f in srcd.rglob("*"):
        if f.is_file():
            dstdir.mkdir(parents=True, exist_ok=True)
            _shutil.copy2(f, dstdir / f.relative_to(srcd))
            n += 1
    return n

_web_n = 0
# 1) DLL и процесс рендера — верхний уровень PySide6 (раскладка Windows)
for _pat in ("Qt6WebEngine*.dll", "Qt6WebChannel*.dll", "Qt6Positioning*.dll",
             "QtWebEngineProcess.exe", "opengl32sw.dll", "icudtl.dat"):
    for _f in _psd.glob(_pat):
        _web_n += _copy_file(_f, _dest_ps)
# 1б) Точные зависимости движка (по таблицам импорта колёс PySide6):
# Qt6WebEngineWidgets.dll -> Qt6QuickWidgets + Qt6PrintSupport,
# Qt6WebEngineQuick.dll   -> Qt6WebChannelQuick.
# (libEGL/libGLESv2/d3dcompiler в Qt 6.11 НЕ нужны — в колёсах их нет)
for _pat in ("Qt6QuickWidgets.dll", "Qt6PrintSupport.dll", "Qt6WebChannelQuick.dll"):
    for _f in _psd.glob(_pat):
        _web_n += _copy_file(_f, _dest_ps)
# то же — в раскладке Qt/bin и Qt/libexec (Linux-колёса)
for _d in (_qt_dir / "bin", _qt_dir / "libexec"):
    if _d.is_dir():
        for _pat in ("Qt6WebEngine*.dll", "Qt6WebChannel*.dll", "Qt6Positioning*.dll",
                     "Qt6WebEngine*.so*", "Qt6WebChannel*.so*", "Qt6Positioning*.so*",
                     "QtWebEngineProcess*"):
            for _f in _d.glob(_pat):
                _web_n += _copy_file(_f, _dest_ps / "Qt" / "bin")
# 2) ресурсы Chromium (.pak) — все возможные места
for _cand in (_psd / "Qt" / "resources", _psd / "resources"):
    _web_n += _copy_dir(_cand, _dest_ps / "Qt" / "resources")
# 3) локали движка (переводы интерфейса Chromium)
for _cand in (_psd / "Qt" / "translations" / "qtwebengine_locales",
              _psd / "translations" / "qtwebengine_locales"):
    _web_n += _copy_dir(_cand, _dest_ps / "Qt" / "translations" / "qtwebengine_locales")
# 4) QML-плагин WebEngine
_web_n += _copy_dir(_psd / "Qt" / "qml" / "QtWebEngine",
                    _dest_ps / "Qt" / "qml" / "QtWebEngine")

# 5) манифест: что реально лежит в dist (видно в логе сборки)
_req = ["PySide6/Qt6WebEngineCore.dll", "PySide6/Qt6WebEngineWidgets.dll",
        "PySide6/Qt6WebEngineQuick.dll", "PySide6/Qt6WebChannel.dll",
        "PySide6/Qt6Positioning.dll", "PySide6/Qt6PrintSupport.dll",
        "PySide6/QtWebEngineProcess.exe",
        "PySide6/Qt/resources/qtwebengine_resources.pak",
        "PySide6/Qt6QuickWidgets.dll", "PySide6/Qt6PrintSupport.dll",
        "PySide6/Qt6WebChannelQuick.dll", "PySide6/QtWebEngineCore.pyd",
        "PySide6/QtWebEngineWidgets.pyd"]
print("[webengine] докопано напрямую: %d файлов" % _web_n)
for _r in _req:
    _ok = (_base / _r).exists()
    print("[webengine] %-52s %s" % (_r, "ЕСТЬ" if _ok else "НЕТ !!!"))
