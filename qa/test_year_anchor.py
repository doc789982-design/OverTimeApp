#!/usr/bin/env python3
"""Тест якоря года у компенсации периодом (logic.compute_month_summary).

Сценарий пользователя: 50 дней компенсации периодом с 1 декабря, период
перешагивает 1 января. Раньше дни следующего года списывались в его же
«текущий» остаток, где переработки прошлого года не видно, — год уходил
в минус и сохранение блокировалось. Теперь списание крепится к году
НАЧАЛА ПЕРИОДА: дни следующего года оплачиваются декабрём того года,
в котором период начался (и для текущего остатка, и для заначки
«за пред. год»).

Проверяется:
  A. 50 дней с 01.12 из текущего остатка — оба года без минусов;
  B. то же из заначки прошлого года («за пред. год») — оба года чистые;
  C. период внутри года — списание по месяцам как раньше (по дням);
  D. одиночный день — без изменений;
  E. старые записи (order_date пуст) — прежнее поведение по дням.

Запуск:
    python3 qa/test_year_anchor.py             # из корня репозитория
"""
import os
import sys
import tempfile
from datetime import date, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from database import DB
from logic import compute_month_summary, validate_non_negative_over_year
import logic


def fresh_db(hire, opening_days=0):
    """База с одним сотрудником; hiring-год открывает указанные дни."""
    path = os.path.join(tempfile.mkdtemp(prefix="anchor_"), "test.sqlite")
    db = DB(path)
    db.add_employee("Тестов", "Тест", "Тестович", "капитан", "инженер",
                    hire, 0, opening_days, 0, 0, 0, 0)
    logic._SUMMARY_CACHE["data"].clear()
    return db


def place_days_period(db, start, n_days, use_prev_year):
    """Повторяет запись saveCompensation: родитель + дни периода."""
    dates = [start + timedelta(days=i) for i in range(n_days)]
    event = "1900-01-01" if use_prev_year else dates[0].isoformat()
    cur = db.conn.execute(
        "INSERT INTO compensation(employee_id,unit,method,amount_days,comment,event_date,order_date) "
        "VALUES (?,?,?,?,?,?,?)",
        (1, "days", "day_off", len(dates), None, event, dates[0].isoformat()))
    cid = cur.lastrowid
    for d in dates:
        db.conn.execute(
            "INSERT INTO comp_day_off_date(compensation_id,employee_id,day_off_date) VALUES (?,?,?)",
            (cid, 1, d.isoformat()))
    return cid


def ok(db, year, what):
    valid, err = validate_non_negative_over_year(db, 1, year)
    assert valid, "%s: %s" % (what, err)


def main() -> int:
    # ── A. Текущий остаток, период через границу года ──
    db = fresh_db("2026-01", opening_days=50)
    place_days_period(db, date(2026, 12, 1), 50, use_prev_year=False)
    ok(db, 2026, "A/2026"); ok(db, 2027, "A/2027")
    dec = compute_month_summary(db, 1, 2026, 12)
    assert dec["end_days"] == 0, "A: декабрь = %s (ожидали 0)" % dec["end_days"]
    jan = compute_month_summary(db, 1, 2027, 1)
    assert jan["end_days"] == 0 and jan["prev_d_start"] == 0, \
        "A: январь-2027 конец=%s заначка=%s" % (jan["end_days"], jan["prev_d_start"])
    days = db.conn.execute("SELECT COUNT(*) c FROM comp_day_off_date").fetchone()["c"]
    assert days == 50, "A: в календаре должно быть 50 дней, а не %s" % days
    print("A: 50 дней с 01.12 из текущего остатка — оба года чистые ✓")

    # ── B. Заначка прошлого года («за пред. год»), тот же период ──
    db = fresh_db("2025-01", opening_days=50)   # к 2026 году лежит в заначке
    place_days_period(db, date(2026, 12, 1), 50, use_prev_year=True)
    ok(db, 2026, "B/2026"); ok(db, 2027, "B/2027")
    dec = compute_month_summary(db, 1, 2026, 12)
    assert dec["prev_d_end"] == 0, "B: заначка в декабре = %s" % dec["prev_d_end"]
    jan = compute_month_summary(db, 1, 2027, 1)
    assert jan["prev_d_start"] == 0 and jan["prev_d_end"] == 0, \
        "B: заначка-2027 = %s..%s" % (jan["prev_d_start"], jan["prev_d_end"])
    print("B: те же 50 дней «за пред. год» — заначка списана в год периода ✓")

    # ── C. Период внутри года: списание по месяцам как раньше ──
    db = fresh_db("2026-01", opening_days=10)
    place_days_period(db, date(2026, 3, 28), 10, use_prev_year=False)  # 4 дня в марте, 6 в апреле
    ok(db, 2026, "C/2026")
    mar = compute_month_summary(db, 1, 2026, 3)
    apr = compute_month_summary(db, 1, 2026, 4)
    assert mar["end_days"] == 6, "C: март = %s (ожидали 6)" % mar["end_days"]
    assert apr["end_days"] == 0, "C: апрель = %s (ожидали 0)" % apr["end_days"]
    print("C: период внутри года — март/апрель по дням, как раньше ✓")

    # ── D. Одиночный день ──
    db = fresh_db("2026-01", opening_days=5)
    place_days_period(db, date(2026, 5, 15), 1, use_prev_year=False)
    ok(db, 2026, "D/2026")
    may = compute_month_summary(db, 1, 2026, 5)
    assert may["end_days"] == 4, "D: май = %s (ожидали 4)" % may["end_days"]
    print("D: одиночный день — списан своим месяцем ✓")

    # ── E. Старые записи без order_date: прежнее поведение ──
    db = fresh_db("2025-01", opening_days=10)   # заначка-2026 = 10
    cur = db.conn.execute(
        "INSERT INTO compensation(employee_id,unit,method,amount_days,event_date) "
        "VALUES (1,'days','day_off',2,'1900-01-01')")
    cid = cur.lastrowid
    for d in (date(2026, 12, 5), date(2027, 1, 5)):
        db.conn.execute(
            "INSERT INTO comp_day_off_date(compensation_id,employee_id,day_off_date) VALUES (?,?,?)",
            (cid, 1, d.isoformat()))
    dec = compute_month_summary(db, 1, 2026, 12)
    jan = compute_month_summary(db, 1, 2027, 1)
    assert dec["prev_d_end"] == 9, "E: декабрь = %s (ожидали 9)" % dec["prev_d_end"]
    assert jan["prev_d_end"] == -1, "E: январь = %s (ожидали −1, старое поведение)" % jan["prev_d_end"]
    print("E: старые записи без якоря — поведение не изменилось ✓")

    print("═══ ЯКОРЬ ГОДА У ПЕРИОДА: РАБОТАЕТ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
