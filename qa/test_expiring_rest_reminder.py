#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Напоминание об отдыхе за прошлый год, который не помещается до 31.12.

Проверяется:
  A. Только прошлогодняя копилка превращается в целые дни: ночные и
     сверхурочные минуты делятся на 8 часов, дни прибавляются.
  B. Рабочие дни считаются с выходными и праздниками производственного
     календаря; ручная правка в базе главнее официального календаря;
     смещённые выходные Вт–Сб не ломаются запасным календарём.
  C. Бэкенд на фиксированную дату находит превышение и равенство, не
     трогает тех, кому дней хватает, уже закончивших службу и ещё не
     принятых; у заканчивающего службу считает дни только до конца службы.
  D. Слот для QML возвращает текст именно за прошлый год и первого
     сотрудника; отдельный тумблер пишется в свой ключ настроек.
  E. QML-контракт: проверка после открытия базы, кнопка «Открыть»,
     штатный текст табеля не затирается своим уведомлением.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_expiring_rest_reminder.py
"""
import os
import sys
import tempfile
import types
from datetime import date
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from PySide6.QtCore import QCoreApplication, QObject  # noqa: E402

import Main as main_module  # noqa: E402
from database import DB  # noqa: E402
from logic import count_workdays_between, previous_year_rest_days  # noqa: E402
from methodical_data import PROD_CALENDARS  # noqa: E402


def make_db(name):
    path = os.path.join(tempfile.mkdtemp(prefix="expiring_rest_"), name)
    return DB(path)


def add_employee(db, last, first="Иван", middle="Иванович", start="2026-01",
                 prev_minutes=0, prev_overtime=0, prev_days=0, group_id=None):
    emp_id = db.add_employee(
        last, first, middle, "капитан", "инженер", start,
        0, 0, 0, prev_minutes, prev_overtime, prev_days, group_id,
    )
    db.conn.commit()
    return emp_id


def test_previous_year_math():
    assert previous_year_rest_days(600, 300, 3) == 4       # 900 мин → 1 день + 3
    assert previous_year_rest_days(480, 479, 0) == 1       # копилки суммируются
    assert previous_year_rest_days(-1000, 480, -5) == 1    # минусы не съедают отдых
    assert previous_year_rest_days(0, 0, 0) == 0
    print("A: прошлогодние часы и дни пересчитываются в дни отдыха ✓")


def test_workday_count():
    db = make_db("calendar.sqlite")
    emp_id = add_employee(db, "Обычный")
    prod_2026 = PROD_CALENDARS[2026]

    # 24, 25, 28, 29, 30 декабря рабочие; 26–27 выходные, 31 — перенос.
    got = count_workdays_between(
        db, emp_id, date(2026, 12, 24), date(2026, 12, 31), prod_2026)
    assert got == 5, got

    # Ручная правка базы главнее производственного календаря.
    db.set_calendar_day_type(date(2026, 12, 31), "work")
    db.conn.commit()
    got = count_workdays_between(
        db, emp_id, date(2026, 12, 24), date(2026, 12, 31), prod_2026)
    assert got == 6, got
    db.set_calendar_day_type(date(2026, 12, 28), "weekend")
    db.conn.commit()
    got = count_workdays_between(
        db, emp_id, date(2026, 12, 24), date(2026, 12, 31), prod_2026)
    assert got == 5, got

    # Смещённые выходные: понедельник выходной, суббота рабочая, хотя
    # официальный календарь для обычной пятидневки говорит наоборот.
    shifted_group = db.add_group("Вт–Сб", False, True)
    shifted_id = add_employee(db, "Смещённый", group_id=shifted_group)
    prod_2027 = PROD_CALENDARS[2027]
    assert count_workdays_between(
        db, shifted_id, date(2027, 1, 11), date(2027, 1, 11), prod_2027) == 0
    assert count_workdays_between(
        db, shifted_id, date(2027, 1, 16), date(2027, 1, 16), prod_2027) == 1
    print("B: выходные, праздники, ручные правки и шаблон Вт–Сб считаются ✓")


def scenario_backend():
    db = make_db("backend.sqlite")
    ids = {}
    # С 20 по 31 декабря 2026 года у обычной пятидневки 8 рабочих дней.
    ids["over"] = add_employee(db, "Иванов", prev_days=9)
    ids["equal"] = add_employee(db, "Петров", "Пётр", "Петрович", prev_days=8)
    ids["hours"] = add_employee(
        db, "Орлов", "Олег", "Олегович",
        prev_minutes=7 * 480, prev_overtime=480, prev_days=1)
    ids["fits"] = add_employee(db, "Сидоров", "Сидор", "Сидорович", prev_days=7)
    ids["short"] = add_employee(db, "Коротков", "Ким", "Кимович", prev_days=3)
    ids["gone"] = add_employee(db, "Ушедший", prev_days=99)
    ids["future"] = add_employee(db, "Будущий", start="2026-12", prev_days=99)

    # Короткову служба заканчивается 23.12: 21, 22 и 23 — ровно 3 рабочих дня.
    db.conn.execute("UPDATE employee SET end_date='2026-12-23' WHERE id=?", (ids["short"],))
    db.conn.execute("UPDATE employee SET end_date='2026-12-19' WHERE id=?", (ids["gone"],))
    db.conn.execute("UPDATE employee SET hire_date='2026-12-25' WHERE id=?", (ids["future"],))
    db.conn.commit()

    backend = main_module.Backend.__new__(main_module.Backend)
    QObject.__init__(backend)
    backend.active_db = db
    return backend, ids


def test_backend_rows_and_slot():
    backend, ids = scenario_backend()
    today = date(2026, 12, 20)
    rows = backend._expiring_rest_rows(today)
    by_id = {row["id"]: row for row in rows}

    assert set(by_id) == {ids["over"], ids["equal"], ids["hours"], ids["short"]}, rows
    assert by_id[ids["over"]]["prev_rest_days"] == 9
    assert by_id[ids["over"]]["workdays_left"] == 8
    assert by_id[ids["over"]]["excess_days"] == 1
    assert by_id[ids["equal"]]["excess_days"] == 0
    assert by_id[ids["hours"]]["prev_rest_days"] == 9
    assert by_id[ids["short"]]["prev_rest_days"] == 3
    assert by_id[ids["short"]]["workdays_left"] == 3
    assert all("2025" not in row["name"] for row in rows)
    print("C: превышение, равенство, сроки службы и состав списка верны ✓")

    class FixedDate(date):
        @classmethod
        def today(cls):
            return cls(2026, 12, 20)

    original_date = main_module.date
    main_module.date = FixedDate
    try:
        info = backend.getExpiringRestReminder()
    finally:
        main_module.date = original_date

    assert info["checked"] is True and info["count"] == 4, info
    assert info["employee_id"] == rows[0]["id"], info
    assert "за 2025 год" in info["message"], info["message"]
    assert "Иванов И.И." in info["message"], info["message"]
    assert "не поместится до 31 декабря" in info["message"], info["message"]

    written = {}
    backend._expiring_rest_reminder_enabled = True
    backend._write_ui_config = lambda key, value: written.update({key: value})
    backend.setExpiringRestReminderEnabled(False)
    assert backend.expiringRestReminderEnabled is False
    assert written == {"expiring_rest_reminder_enabled": False}, written
    print("D: текст слота, первый сотрудник и отдельный тумблер верны ✓")


def test_qml_contract():
    main_qml = (ROOT / "main.qml").read_text(encoding="utf-8")
    settings = (ROOT / "components/SettingsDialog.qml").read_text(encoding="utf-8")
    notification = (ROOT / "components/AppDesktopNotification.qml").read_text(encoding="utf-8")

    assert "expiringRestOpenTimer.restart()" in main_qml
    assert "mainWindow.lastExpiringRestCheckMs = 0" in main_qml
    assert "backend.getExpiringRestReminder()" in main_qml
    assert "pendingAlertEmployeeId" in main_qml
    assert 'showCustom(' in main_qml and '"Открыть"' in main_qml
    assert "backend.expiringRestReminderEnabled" in settings
    assert "backend.setExpiringRestReminderEnabled(checked)" in settings
    assert "readonly property string defaultMessage" in notification
    assert "root.message = root.defaultMessage" in notification
    print("E: QML проверяет после открытия базы и не затирает текст табеля ✓")


def main():
    app = QCoreApplication.instance() or QCoreApplication([])
    _ = app
    test_previous_year_math()
    test_workday_count()
    test_backend_rows_and_slot()
    test_qml_contract()
    print("PASS: напоминание об отдыхе за прошлый год работает")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
