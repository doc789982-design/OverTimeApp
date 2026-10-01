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

from release_url import from_ci_env  # noqa: E402
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
    "PySide6.QtPrintSupport",  # собственная печать бланка без Excel
    "openpyxl",                # выгрузка табеля в Excel
    "print_engine",           # собственная печать бланка (ленивый импорт)
    "win32print",              # список принтеров
    "win32com",
    "win32com.client",         # печать через установленный Excel
    "pythoncom",
    "pywintypes",
    "app_update",
    "recovery",               # аварийный восстановитель (первый import Main.py)
    "methodical_data",        # методички и производственные календари (окно справки)
]

EXCLUDES = list(UNUSED_STDLIB) + list(UNUSED_PYSIDE_MODULES)

# Не ставим --collect-all PySide6. Обычных хуков PyInstaller + списка
# выше хватает, чтобы подтянуть QML и плагины окон. Лишнее режем ниже.


_changelog = ROOT / "CHANGELOG.md"
_datas = []
if _changelog.exists():
    _datas.append((str(_changelog), "."))
# Карта сборок (память проекта о том, что въехало в какую сборку) —
# в каждой сборке рядом с журналом изменений.
_history = ROOT / "BUILD_HISTORY.md"
if _history.exists():
    _datas.append((str(_history), "."))

# version.json внутрь сборки. В AppTheme лежит имя для человека (например
# «BETA.1»), а обновлятор старых сборок понимает только «X.Y.Z-ИМЯ.N».
# Поэтому в пакет кладётся машинная строка 2.0.0-ALPHA.<сборка> (сборки
# нумеруются сквозняком и старым клиентам сравниваются по номеру сборки)
# плюс поле display с настоящим именем — его показывают окна и тосты.
_theme_src = (ROOT / "components" / "AppTheme.qml").read_text(encoding="utf-8")
import json as _json
import re as _re
_m_ver = _re.search(r'appVersion:\s*"([^"]+)"', _theme_src)
_m_bld = _re.search(r"appBuild:\s*(\d+)", _theme_src)
_app_name = _m_ver.group(1) if _m_ver else ""
_app_build = int(_m_bld.group(1)) if _m_bld else 0
_meta_dir = ROOT / "build"
_meta_dir.mkdir(exist_ok=True)
_meta_file = _meta_dir / "version.json"
# Имя архива релиза внутрь version.json: файл читают разные хранилища
# (GitHub, запасной сервер post.mvd.ru), поэтому в url лежит только имя
# архива (OVERTIMETAB_<имя>_<sha7>.zip), а полный адрес склеивается с тем
# хранилищем, где файл лежит. В сборке по ветке имени ещё нет — поле не пишется.
_meta = {
    "name": "OVERTIMETAB",
    "version": "2.0.0-ALPHA.%d" % _app_build,
    "build": _app_build,
    "display": _app_name,
}
_asset_url = from_ci_env(_app_name)
if _asset_url:
    _meta["url"] = _asset_url
_meta_file.write_text(_json.dumps(_meta, ensure_ascii=False, indent=2) + "\n",
                      encoding="utf-8")
_datas.append((str(_meta_file), "."))
print("[версия] имя=%s сборка=%d машинная=2.0.0-ALPHA.%d"
      % (_app_name, _app_build, _app_build))

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

# Манифест файлов сборки: при старте программа сверяет с ним свою папку
# и убирает файлы прошлых версий (установка «поверх» старой — с заменой).
_dist_root = ROOT / "dist" / "OVERTIMETAB"
_manifest = _dist_root / "app_files.txt"
_lines = ["# OVERTIMETAB build %d" % _app_build]
for _p in sorted(_dist_root.rglob("*")):
    if _p.is_file() and _p != _manifest:
        _lines.append(_p.relative_to(_dist_root).as_posix())
_manifest.write_text("\n".join(_lines) + "\n", encoding="utf-8")
print("[манифест] %d файлов сборки %d" % (len(_lines) - 1, _app_build))
