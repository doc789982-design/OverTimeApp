#!/usr/bin/env python3
"""Тест установки «с заменой»: cleanup_stale_files (app_update).

Новая версия, поставленная поверх старой, оставляет файлы, которых в новой
сборке нет. cleanup_stale_files сверяет папку с манифестом app_files.txt
(его кладёт в сборку tools/overtimetab.spec) и вычищает лишнее, не трогая
пользовательское. Здесь моделируются обе ситуации: ручная замена поверх
старой сборки (в папке смесь старых и новых файлов) и повторный запуск
уже почищенной папки (маркер сборки — чистка не повторяется).

Запуск:
    python3 qa/test_cleanup.py            # из корня репозитория
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from pathlib import Path
import tempfile

import app_update


def make_install(tmp: Path):
    """Папка «после установки новой версии поверх старой».

    Манифест новой сборки: exe, dll, PySide6/{Qt6Gui.dll, QtCore.pyd},
    PySide6/Qt/qml/qm.qml; плюс служебные version.json / app_files.txt.
    Мусор от старой: oldlib.dll в корне, stale.dll и пустой oldqm/ в
    PySide6, целая python-папка numpy/, старый base_library.zip.
    Пользовательское: data/base.sqlite, Отчёты/табель.xlsx, заметка.txt.
    """
    root = tmp / "OVERTIMETAB"
    manifest_lines = [
        "# OVERTIMETAB build 226",
        "OVERTIMETAB.exe",
        "python311.dll",
        "PySide6/Qt6Gui.dll",
        "PySide6/QtCore.pyd",
        "PySide6/Qt/qml/qm.qml",
        "version.json",
    ]
    (root / "PySide6" / "Qt" / "qml").mkdir(parents=True)
    for rel in manifest_lines[1:]:
        (root / rel).write_bytes(b"x")
    (root / MANIFEST).write_text("\n".join(manifest_lines) + "\n", encoding="utf-8")

    # мусор прошлой сборки
    (root / "oldlib.dll").write_bytes(b"x")                 # корень, программный
    (root / "base_library.zip").write_bytes(b"x")           # корень, программный
    (root / "PySide6" / "stale.dll").write_bytes(b"x")      # внутри PySide6
    (root / "PySide6" / "oldqm").mkdir()
    (root / "PySide6" / "oldqm" / "x.qm").write_bytes(b"x") # subdir старой сборки
    numpy_dir = root / "numpy"
    numpy_dir.mkdir()
    (numpy_dir / "__init__.py").write_text("# old", encoding="utf-8")
    (numpy_dir / "core.pyd").write_bytes(b"x")

    # пользовательское — трогать нельзя
    (root / "data").mkdir()
    (root / "data" / "base.sqlite").write_bytes(b"x")
    reports = root / "Отчёты"
    reports.mkdir()
    (reports / "табель.xlsx").write_bytes(b"x")
    (root / "заметка.txt").write_text("моя заметка", encoding="utf-8")
    return root


MANIFEST = app_update.MANIFEST_NAME
KEEP = ["OVERTIMETAB.exe", "python311.dll", "PySide6/Qt6Gui.dll",
        "PySide6/QtCore.pyd", "PySide6/Qt/qml/qm.qml", "version.json",
        MANIFEST,
        "data/base.sqlite", "Отчёты/табель.xlsx", "заметка.txt"]
GONE = ["oldlib.dll", "base_library.zip", "PySide6/stale.dll",
        "PySide6/oldqm/x.qm", "PySide6/oldqm", "numpy"]


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="cleanup_test_"))
    root = make_install(tmp)
    install_root = tmp  # install_root() ищет exe — он в root, передаём root

    removed = app_update.cleanup_stale_files(root)
    print("удалено (%d): %s" % (len(removed), sorted(removed)))

    for rel in KEEP:
        assert (root / rel).exists(), "исчезло нужное/пользовательское: %s" % rel
    print("всё нужное и пользовательское на месте ✓")

    gone_ok = []
    for rel in GONE:
        if not (root / rel).exists():
            gone_ok.append(rel)
    assert len(gone_ok) == len(GONE), "не убрано: %s" % (
        [r for r in GONE if (root / r).exists()])
    print("мусор старой сборки убран ✓ (%s)" % ", ".join(gone_ok))

    # маркер: повторный запуск той же сборки ничего не делает
    removed2 = app_update.cleanup_stale_files(root)
    assert removed2 == [], "повторная чистка что-то удалила: %s" % removed2
    print("повторный запуск ничего не трогает ✓")

    # а сборка поновее снова чистит (маркер другой)
    mf = root / MANIFEST
    mf.write_text(mf.read_text(encoding="utf-8").replace("build 226", "build 227"),
                  encoding="utf-8")
    (root / "junk.dll").write_bytes(b"x")
    removed3 = app_update.cleanup_stale_files(root)
    assert "junk.dll" in removed3, "новая сборка не почистила: %s" % removed3
    print("следующая сборка снова чистит ✓")

    print("═══ УСТАНОВКА «С ЗАМЕНОЙ»: ЧИСТКА РАБОТАЕТ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
