#!/usr/bin/env python3
"""Тест автоматического обнаружения баз.

Сценарии с практики:
  1. Базы лежат в папке хранения, но не подключены в программе —
     программа находит их сама (молча, без вопросов).
  2. Путь в config.json битый (файл перенесли руками в папку хранения) —
     программа находит файл по имени и перепривязывает путь.
  3. Файл исчез совсем — база НЕ вычёркивается молча, а показывается
     с пометкой «файл не найден».
  4. Чужой sqlite и пустые файлы в папке хранения не подключаются.
  5. Своя папка хранения (выбранная в настройках) сканируется так же.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen python3 qa/test_db_discovery.py
"""
import json
import os
import sys
import tempfile
from pathlib import Path

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import types

# Main.py импортирует win32print (только Windows) — в песочнице подменяем
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from PySide6.QtCore import QObject, Signal  # noqa: E402

import Main  # noqa: E402  (тяжёлый импорт: PySide6 и всё остальное)
from database import DB  # noqa: E402


class Harness(QObject):
    """Обвязка без тяжёлого Backend: только нужные методы и пути."""

    dbListChanged = Signal()

    def __init__(self, data_dir: Path, app_dir: Path):
        super().__init__()
        self._data_dir = data_dir
        self.config_path = data_dir / "config.json"
        self.app_dir = app_dir
        self._db_list = []
        for name in ("load_databases", "_known_db_dirs", "_find_db_by_name",
                     "_persist_healed_paths", "_scan_db_folder",
                     "_get_db_dir", "add_to_config", "_validate_db_file"):
            setattr(self, name, types.MethodType(getattr(Main.Backend, name), self))


def make_db(path: Path) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    db = DB(str(path))
    db.close()
    return str(path)


def write_config(data_dir: Path, db_paths, last=None, default_db_dir=None):
    ui = {}
    if default_db_dir:
        ui["default_db_dir"] = str(default_db_dir)
    cfg = {"db_paths": [str(p) for p in db_paths],
           "last_db_path": str(last) if last else None, "ui": ui}
    data_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / "config.json").write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def read_paths(data_dir: Path):
    return json.loads(
        (data_dir / "config.json").read_text(encoding="utf-8"))["db_paths"]


def main() -> int:
    tmp = Path(tempfile.mkdtemp(prefix="dbfind_"))
    data_dir = tmp / "OverTimeTab"
    app_dir = tmp / "app"
    app_dir.mkdir()
    dbs_dir = data_dir / "databases"

    # ── 1. автоскан папки хранения: файлы лежат, в конфиге их нет ──
    make_db(dbs_dir / "alpha.sqlite")
    make_db(dbs_dir / "beta.sqlite")
    write_config(data_dir, [])
    h = Harness(data_dir, app_dir)
    added = h._scan_db_folder()
    assert added == 2, added
    names = {e["name"] for e in h._db_list}
    assert len(h._db_list) == 2, h._db_list
    assert all("missing" not in e for e in h._db_list), h._db_list
    saved = read_paths(data_dir)
    assert len(saved) == 2 and any("alpha.sqlite" in p for p in saved), saved
    # повторный скан ничего не добавляет
    assert h._scan_db_folder() == 0
    print("1: базы в папке хранения подключаются молча, без дублей ✓")

    # ── 2. битый путь лечится по имени файла ──
    make_db(dbs_dir / "gamma.sqlite")
    gone = tmp / "old_place" / "gamma.sqlite"     # файл «перенесли руками»
    write_config(data_dir, [gone], last=gone)
    h2 = Harness(data_dir, app_dir)
    h2.load_databases()
    entry = [e for e in h2._db_list if "gamma" in e["path"]]
    assert entry and "missing" not in entry[0], h2._db_list
    assert entry[0]["path"] == str(dbs_dir / "gamma.sqlite"), entry[0]
    saved = read_paths(data_dir)
    assert str(dbs_dir / "gamma.sqlite") in saved, saved      # путь переписан
    assert not any("old_place" in p for p in saved), saved
    cfg = json.loads((data_dir / "config.json").read_text(encoding="utf-8"))
    assert cfg["last_db_path"] == str(dbs_dir / "gamma.sqlite").replace("\\", "/"), cfg
    print("2: перенесённую руками базу программа находит и перепривязывает ✓")

    # ── 3. файл исчез совсем — базы в списке нет (путь в конфиге живёт) ──
    write_config(data_dir, [tmp / "nowhere" / "delta.sqlite"])
    h3 = Harness(data_dir, app_dir)
    h3.load_databases()
    assert h3._db_list == [], h3._db_list
    # путь остаётся в конфиге — выбрать пропавшее нельзя
    assert read_paths(data_dir) == [str(tmp / "nowhere" / "delta.sqlite")]
    # файл вернулся на старое место — база снова в списке
    make_db(tmp / "nowhere" / "delta.sqlite")
    h3b = Harness(data_dir, app_dir)
    h3b.load_databases()
    assert len(h3b._db_list) == 1 and "delta" in h3b._db_list[0]["path"], h3b._db_list
    print("3: пропавшая база не показывается; файл вернулся — база вернулась ✓")

    # ── 4. чужой sqlite и пустые файлы не подключаются ──
    import sqlite3
    data_dir4 = tmp / "check4"
    dbs4 = data_dir4 / "databases"
    dbs4.mkdir(parents=True, exist_ok=True)
    foreign = dbs4 / "foreign.sqlite"
    con = sqlite3.connect(str(foreign))
    con.execute("CREATE TABLE foo (x INT)")
    con.commit()
    con.close()
    (dbs4 / "empty.sqlite").touch()
    (dbs4 / "notes.txt").write_text("не база", encoding="utf-8")
    write_config(data_dir4, [])
    h4 = Harness(data_dir4, app_dir)
    assert h4._scan_db_folder() == 0
    assert h4._db_list == [], h4._db_list
    print("4: чужие и пустые файлы в папке не подключаются ✓")

    # ── 5. своя папка хранения (из настроек) сканируется так же ──
    custom = tmp / "my_databases"
    make_db(custom / "epsilon.sqlite")
    write_config(data_dir, [], default_db_dir=custom)
    h5 = Harness(data_dir, app_dir)
    assert h5._scan_db_folder() == 1
    assert any("epsilon" in e["path"] for e in h5._db_list), h5._db_list
    # и лечение тоже ищет в ней
    gone2 = tmp / "old_place2" / "epsilon.sqlite"
    write_config(data_dir, [gone2], default_db_dir=custom)
    h6 = Harness(data_dir, app_dir)
    h6.load_databases()
    assert any(e["path"] == str(custom / "epsilon.sqlite")
               for e in h6._db_list), h6._db_list
    print("5: выбранная в настройках папка хранения ищется наравне со стандартной ✓")

    # ── 6. неудача открытия базы — сигнал (QML возвращает экран выбора) ──
    class FailHarness(QObject):
        showToast = Signal(str, str)
        databaseOpenFailed = Signal(str)

        def __init__(self):
            super().__init__()
            self.active_db = None
            for name in ("openDatabase", "_validate_db_file"):
                setattr(self, name, types.MethodType(
                    getattr(Main.Backend, name), self))

    fh = FailHarness()
    captured = []
    fh.databaseOpenFailed.connect(lambda msg: captured.append(msg))

    # файла нет
    fh.openDatabase(str(tmp / "no_such" / "x.sqlite"))
    assert captured and "не найден" in captured[0], captured
    # файл повреждён (не SQLite)
    bad = tmp / "broken.sqlite"
    bad.write_bytes(b"not a database at all")
    captured.clear()
    fh.openDatabase(str(bad))
    assert captured, captured
    # чужая база (SQLite, но не табель)
    alien = tmp / "alien.sqlite"
    import sqlite3
    con = sqlite3.connect(str(alien))
    con.execute("CREATE TABLE foo (x INT)")
    con.commit(); con.close()
    captured.clear()
    fh.openDatabase(str(alien))
    assert captured, captured
    print("6: неудача открытия (нет файла, повреждена, чужая) — сигнал об ошибке ✓")

    print("═══ БАЗЫ НАХОДЯТСЯ, ПУТИ ЛЕЧАТСЯ, НЕУДАЧА ВОЗВРАЩАЕТ ВЫБОР ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
