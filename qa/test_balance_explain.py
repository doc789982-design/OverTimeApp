#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""«Почему такая цифра?» — расшифровка панели «Итоги».

Проверяется:
  A. Вклад дежурств считается тем же кодом, что и итоги: ночные часы
     с вычетом перерыва, один день за нерабочую дату, а не по числу
     дежурств в эту дату.
  B. Окно для ночных часов и дней сходится с панелью и честно показывает
     пересчёт: взятые сверх дневной копилки дни закрыты ночными часами,
     отдельно для этого и прошлого года.
  C. Деньги по приказу видны в сверхурочных, «Всего» раскладывается
     по часовым остаткам и дням той же формулой, что и панель.
  D. У сменного графика сверхурочные объясняются помесячно:
     отработано по графику минус норма.
  E. Внеграфика у сменного графика даёт только один день за календарную
     дату — перерывы и повторное дежурство в ту же дату день не умножают.
  F. QML-контракт: все три графы и «Всего» нажимаемы и ведут в окно.

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

from PySide6.QtCore import QCoreApplication, QObject  # noqa: E402

import Main as main_module  # noqa: E402
from database import DB  # noqa: E402
from logic import compute_month_summary, get_accrual_details_for_period  # noqa: E402
from utils import month_bounds_dt  # noqa: E402


def make_db(name):
    return DB(os.path.join(tempfile.mkdtemp(prefix="balance_explain_"), name))


def make_backend(db, emp_id=1, year=2026, month=10):
    backend = main_module.Backend.__new__(main_module.Backend)
    QObject.__init__(backend)
    backend.active_db = db
    backend._selected_employee_id = emp_id
    backend.current_year = year
    backend.current_month = month
    return backend


def rows_sum(rows, source=None):
    return sum(int(r.get("amount_value", 0) or 0) for r in rows
               if source is None or r.get("source") == source)


def scenario():
    db = make_db("main.sqlite")
    emp_id = db.add_employee(
        "Иванов", "Иван", "Иванович", "майор", "инженер", "2026-01",
        600, 0, 600,          # этот год: 10 ч ночных, 0 дн., 10 ч сверхурочных
        480, 120, 1)          # прошлый: 8 ч, 2 ч, 1 день
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

    db.add_compensation_hours_dayoff(emp_id, date(2026, 10, 12), 120, "")
    db.add_compensation_days_dayoff(
        emp_id, [date(2026, 10, 13), date(2026, 10, 14)], "")
    prev_days_id = db.add_compensation_days_dayoff(
        emp_id, [date(2026, 10, 15), date(2026, 10, 16)], "")
    db.conn.execute(
        "UPDATE compensation SET event_date='1900-01-01', order_date='2026-10-15' "
        "WHERE id=?", (prev_days_id,))
    db.add_compensation_money(
        emp_id, "overtime", 180, None, "77", date(2026, 10, 20), "")
    prev_money_id = db.add_compensation_money(
        emp_id, "overtime", 120, None, "78", date(2026, 10, 21), "")
    db.conn.execute(
        "UPDATE compensation SET event_date='1900-01-01' WHERE id=?",
        (prev_money_id,))
    db.conn.commit()
    return db, emp_id, night_id, weekend_first, weekend_second


def test_details_and_backend():
    db, emp_id, night_id, weekend_first, weekend_second = scenario()
    summary = compute_month_summary(db, emp_id, 2026, 10)
    assert summary["acc_hours"] == 420 and summary["acc_days"] == 1, summary
    assert summary["comp_h_real"] == 600 and summary["comp_d_real"] == 1, summary
    assert summary["comp_h_prev"] == 480 and summary["comp_d_prev"] == 1, summary

    start, end = month_bounds_dt(2026, 10)
    totals, details = get_accrual_details_for_period(db, emp_id, start, end)
    by_id = {d["duty_id"]: d for d in details if d.get("kind") == "duty"}
    assert totals["night"] == 420 and totals["days"] == 1, totals
    assert by_id[night_id]["night"] == 420, by_id[night_id]
    assert by_id[weekend_first]["days"] == 1, by_id[weekend_first]
    assert by_id[weekend_second]["days"] == 0, by_id[weekend_second]
    print("A: вклад дежурств совпадает с итогами, суббота считается один раз ✓")

    backend = make_backend(db, emp_id)
    hours = backend.getBalanceExplanation("hours", "month")
    assert hours["ok"] is True and hours["result_text"].startswith("7 ч."), hours
    assert rows_sum(hours["acc_rows"]) == 420, hours["acc_rows"]
    assert rows_sum(hours["comp_rows"], "current") == -600, hours["comp_rows"]
    assert rows_sum(hours["comp_rows"], "prev") == -480, hours["comp_rows"]
    assert any(r["kind"] == "adjustment" and r["amount_value"] == -480
               for r in hours["comp_rows"])

    days = backend.getBalanceExplanation("days", "month")
    assert rows_sum(days["acc_rows"]) == 1, days["acc_rows"]
    assert rows_sum(days["comp_rows"], "current") == -1, days["comp_rows"]
    assert rows_sum(days["comp_rows"], "prev") == -1, days["comp_rows"]
    assert sum(1 for r in days["comp_rows"] if r["kind"] == "adjustment") == 2
    print("B: ночные и дни сходятся, пересчёт дневной копилки показан отдельно ✓")

    overtime = backend.getBalanceExplanation("overtime", "month")
    assert rows_sum(overtime["comp_rows"], "current") == -180, overtime
    assert rows_sum(overtime["comp_rows"], "prev") == -120, overtime
    assert any("приказ № 77" in r["detail"] for r in overtime["comp_rows"]), overtime

    total = backend.getBalanceExplanation("total", "month")
    assert total["result_text"] == "1 день", total
    included = [r for r in total["total_rows"]
                if r["title"] == "Положительные часы вместе"]
    assert included and included[0]["amount_value"] == 840, total["total_rows"]
    print("C: приказ виден, «Всего» раскладывается по остаткам и формуле ÷480 ✓")


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

    backend = make_backend(db, emp_id)
    info = backend.getBalanceExplanation("overtime", "month")
    assert rows_sum(info["acc_rows"]) == 240, info
    assert "Отработано 12 ч. при норме 8 ч." in info["acc_rows"][0]["detail"], info
    print("D: сменные сверхурочные объяснены сравнением с нормой месяца ✓")


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

    backend = make_backend(db, emp_id)
    days = backend.getBalanceExplanation("days", "month")
    assert rows_sum(days["acc_rows"]) == 2, days["acc_rows"]
    print("E: перерывы внеграфика не размножают дни — одна дата = один день ✓")


def test_qml_contract():
    panel = (ROOT / "components/AppSummaryPanel.qml").read_text(encoding="utf-8")
    dialog = (ROOT / "components/BalanceExplainDialog.qml").read_text(encoding="utf-8")
    main_qml = (ROOT / "main.qml").read_text(encoding="utf-8")
    for bucket in ('"hours"', '"days"', '"overtime"', '"total"'):
        assert bucket in panel, bucket
    assert "balanceExplainDialog.openExplain" in panel
    assert "TapHandler" in panel and "HoverHandler" in panel
    assert "backend.getBalanceExplanation" in dialog
    assert "Почему такая цифра?" in dialog
    assert "BalanceExplainDialog" in main_qml
    print("F: все цифры панели ведут в окно расшифровки ✓")


def main():
    app = QCoreApplication.instance() or QCoreApplication([])
    _ = app
    test_details_and_backend()
    test_shift_overtime()
    test_shift_off_schedule_days_unique()
    test_qml_contract()
    print("PASS: расшифровка остатков сходится с панелью «Итоги»")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
