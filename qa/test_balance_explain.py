#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Вклад дежурств в итоги; отдельного окна расшифровки больше нет.

Проверяется:
  A. Вклад дежурств считается тем же кодом, что и итоги: ночные часы
     с вычетом перерыва, один день за нерабочую дату, а не по числу
     дежурств в эту дату.
  B. У сменного графика сверхурочные за месяц считаются сравнением
     отработанного по графику с нормой.
  C. Внеграфика у сменного графика даёт только один день за календарную
     дату — перерывы и повторное дежурство день не умножают.
  D. Интерфейс расшифровки удалён: нет окна, нет нажатий по строкам
     панели и нет соответствующего вызова в Main.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_balance_explain.py
"""
import os
import sys
import tempfile
import types
from datetime import date, datetime, timedelta
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from database import DB  # noqa: E402
from logic import compute_month_summary, get_accrual_details_for_period  # noqa: E402
from utils import month_bounds_dt  # noqa: E402


def make_db(name):
    return DB(os.path.join(tempfile.mkdtemp(prefix="balance_details_"), name))


def test_duty_details():
    db = make_db("main.sqlite")
    emp_id = db.add_employee(
        "Иванов", "Иван", "Иванович", "майор", "инженер", "2026-01",
        600, 0, 600,
        480, 120, 1)
    night_id = db.add_duty(
        emp_id, datetime(2026, 10, 7, 20, 0), datetime(2026, 10, 8, 8, 0),
        "ночное дежурство")
    db.replace_duty_breaks(night_id, [(
        datetime(2026, 10, 8, 2, 0), datetime(2026, 10, 8, 3, 0))])
    weekend_first = db.add_duty(
        emp_id, datetime(2026, 10, 10, 9, 0), datetime(2026, 10, 10, 17, 0),
        "суббота")
    weekend_second = db.add_duty(
        emp_id, datetime(2026, 10, 10, 18, 0), datetime(2026, 10, 10, 21, 0),
        "второе в ту же субботу")

    summary = compute_month_summary(db, emp_id, 2026, 10)
    assert summary["acc_hours"] == 420 and summary["acc_days"] == 1, summary

    start, end = month_bounds_dt(2026, 10)
    totals, details = get_accrual_details_for_period(db, emp_id, start, end)
    by_id = {d["duty_id"]: d for d in details if d.get("kind") == "duty"}
    assert totals["night"] == 420 and totals["days"] == 1, totals
    assert by_id[night_id]["night"] == 420, by_id[night_id]
    assert by_id[weekend_first]["days"] == 1, by_id[weekend_first]
    assert by_id[weekend_second]["days"] == 0, by_id[weekend_second]
    print("A: вклад дежурств совпадает с итогами, суббота считается один раз ✓")


def test_shift_overtime():
    db = make_db("shift.sqlite")
    group_id = db.add_group("Смены", True, False)
    emp_id = db.add_employee(
        "Сменный", "Семён", "Семёнович", "капитан", "инженер", "2026-01",
        0, 0, 0, 0, 0, 0, group_id)
    cur = date(2026, 10, 1)
    while cur.month == 10:
        if cur.weekday() < 5 and cur.day != 1:
            db.set_day_status(emp_id, cur, "Б")
        cur += timedelta(days=1)
    db.add_duty(emp_id, datetime(2026, 10, 1, 8, 0),
                datetime(2026, 10, 1, 20, 0), "смена", True)
    db.conn.commit()

    start, end = month_bounds_dt(2026, 10)
    totals, details = get_accrual_details_for_period(db, emp_id, start, end)
    months = [d for d in details if d.get("kind") == "overtime_month"]
    assert totals["overtime_acc"] == 240, totals
    assert months and months[0]["shift_minutes"] == 720, months
    assert months[0]["norm_minutes"] == 480 and months[0]["overtime"] == 240, months
    print("B: сменные сверхурочные считаются сравнением с нормой месяца ✓")


def test_shift_off_schedule_days_unique():
    db = make_db("shift_off_schedule.sqlite")
    group_id = db.add_group("Смены", True, False)
    emp_id = db.add_employee(
        "Внеграфика", "Виктор", "Викторович", "майор", "инженер", "2026-01",
        0, 0, 0, 0, 0, 0, group_id)
    long_duty = db.add_duty(
        emp_id, datetime(2026, 10, 10, 8, 0), datetime(2026, 10, 11, 8, 0),
        "вне графика с перерывами", False)
    db.replace_duty_breaks(long_duty, [
        (datetime(2026, 10, 10, 12, 0), datetime(2026, 10, 10, 13, 0)),
        (datetime(2026, 10, 10, 18, 0), datetime(2026, 10, 10, 19, 0)),
        (datetime(2026, 10, 11, 2, 0), datetime(2026, 10, 11, 3, 0)),
    ])
    second_duty = db.add_duty(
        emp_id, datetime(2026, 10, 11, 10, 0), datetime(2026, 10, 11, 12, 0),
        "ещё раз в ту же дату", False)
    db.conn.commit()

    summary = compute_month_summary(db, emp_id, 2026, 10)
    assert summary["acc_days"] == 2, summary

    start, end = month_bounds_dt(2026, 10)
    totals, details = get_accrual_details_for_period(db, emp_id, start, end)
    by_id = {d["duty_id"]: d for d in details if d.get("kind") == "duty"}
    assert totals["days"] == 2, totals
    assert by_id[long_duty]["days"] == 2, by_id[long_duty]
    assert by_id[long_duty]["day_dates"] == [date(2026, 10, 10), date(2026, 10, 11)]
    assert by_id[second_duty]["days"] == 0, by_id[second_duty]
    print("C: перерывы внеграфика не размножают дни — одна дата = один день ✓")


def test_explainer_removed():
    panel = (ROOT / "components/AppSummaryPanel.qml").read_text(encoding="utf-8")
    main_qml = (ROOT / "main.qml").read_text(encoding="utf-8")
    backend = (ROOT / "Main.py").read_text(encoding="utf-8")
    assert not (ROOT / "components/BalanceExplainDialog.qml").exists()
    for forbidden in ("TapHandler", "HoverHandler", "openExplain",
                      "balanceExplainDialog", "Нажмите строку"):
        assert forbidden not in panel, forbidden
    assert "BalanceExplainDialog" not in main_qml
    assert "balanceExplainDialog" not in main_qml
    assert "getBalanceExplanation" not in backend
    assert "Почему такая цифра" not in backend
    print("D: окно расшифровки и нажатия по строкам полностью удалены ✓")


def main():
    test_duty_details()
    test_shift_overtime()
    test_shift_off_schedule_days_unique()
    test_explainer_removed()
    print("PASS: вклад дежурств сходится с итогами, расшифровки нет")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
