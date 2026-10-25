#!/usr/bin/env python3
"""Тест компенсации «За оба года» (split_both_years + две записи).

Сценарий пользователя: у человека есть переработка и за текущий, и за
прошлый год — раньше приходилось ставить два периода подряд (сначала
«за этот год», потом ещё раз «за пред. год»). Теперь в окне три
переключателя: «За текущий год», «За предыдущий год», «За оба года».
«За оба года» делает один период из всех доступных дней: первая часть
дней списывается из текущего года, остаток — из заначки прошлого
(две записи, дни не пересекаются, каждая копилка считает своё).

Проверяется:
  A. Деление по ёмкостям: дни и часы, края (0, ровно ёмкость, больше);
  B. Случай пользователя: 47 дней при 21 д + 108 ч (тек.) и
     31 д + 128 ч (прошл.) → 34 дня текущим + 13 заначкой,
     оба года чистые, остатки сходятся;
  C. Запрос больше общего пула → валидация блокирует по ночным часам;
  D. Заначка пуста → всё одним периодом текущего года;
  E. Текущий пуст → всё заначкой;
  F. Часы «за оба года»: 5 ч при 2 ч (тек.) и 3 ч (прошл.) → 2+3,
     обе копилки обнуляются.

Запуск:
    python3 qa/test_both_years.py             # из корня репозитория
"""
import os
import sys
import tempfile
from datetime import date, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from database import DB
from logic import (compute_month_summary, validate_non_negative_over_year,
                   split_both_years)
import logic


def fresh_db(hire, opening_minutes=0, opening_days=0,
             prev_minutes=0, prev_days=0):
    path = os.path.join(tempfile.mkdtemp(prefix="both_"), "test.sqlite")
    db = DB(path)
    db.add_employee("Тестов", "Тест", "Тестович", "капитан", "инженер",
                    hire, opening_minutes, opening_days, 0,
                    prev_minutes, 0, prev_days)
    logic._SUMMARY_CACHE["data"].clear()
    return db


def working_days(start, n):
    out, cur = [], start
    while len(out) < n:
        if cur.weekday() < 5:
            out.append(cur)
        cur += timedelta(days=1)
    return out


def place_both_days(db, dates, d1):
    """Две записи, как пишет saveCompensationBoth: сначала текущий год."""
    start = dates[0].isoformat()
    if d1 > 0:
        cur = db.conn.execute(
            "INSERT INTO compensation(employee_id,unit,method,amount_days,event_date,order_date) "
            "VALUES (1,'days','day_off',?,?,?)", (d1, start, start))
        for d in dates[:d1]:
            db.conn.execute(
                "INSERT INTO comp_day_off_date(compensation_id,employee_id,day_off_date) VALUES (?,?,?)",
                (cur.lastrowid, 1, d.isoformat()))
    if len(dates) - d1 > 0:
        cur = db.conn.execute(
            "INSERT INTO compensation(employee_id,unit,method,amount_days,event_date,order_date) "
            "VALUES (1,'days','day_off',?,?,?)", (len(dates) - d1, "1900-01-01", start))
        for d in dates[d1:]:
            db.conn.execute(
                "INSERT INTO comp_day_off_date(compensation_id,employee_id,day_off_date) VALUES (?,?,?)",
                (cur.lastrowid, 1, d.isoformat()))


def ok(db, year, what):
    valid, err = validate_non_negative_over_year(db, 1, year)
    assert valid, "%s: %s" % (what, err)


def main() -> int:
    # ── A. Деление по ёмкостям ──
    cur = {"hours": 108, "days": 21, "overtime": 0}
    prev = {"hours": 128, "days": 31, "overtime": 0}
    assert split_both_years(cur, prev, "days", 47) == (34, 13)   # 21+13 тек., 13 заначкой
    assert split_both_years(cur, prev, "days", 34) == (34, 0)
    assert split_both_years(cur, prev, "days", 20) == (20, 0)
    assert split_both_years(cur, prev, "days", 82) == (34, 48)
    assert split_both_years({"hours": 0, "days": 0}, prev, "days", 47) == (0, 47)
    assert split_both_years(cur, {"hours": 0, "days": 0}, "days", 47) == (34, 13)
    assert split_both_years(cur, prev, "hours", 300) == (300, 0)   # влезает в текущий
    assert split_both_years({"hours": 2}, {"hours": 3}, "hours", 300) == (120, 180)
    assert split_both_years(cur, prev, "overtime", 0) == (0, 0)
    print("A: деление дней и часов по ёмкостям копилок ✓")

    # ── B. Случай пользователя: 47 дней «за оба года» ──
    db = fresh_db("2026-01", opening_minutes=108 * 60, opening_days=21,
                  prev_minutes=128 * 60, prev_days=31)
    days = working_days(date(2026, 11, 16), 47)
    d1, d2 = split_both_years(
        {"hours": 108, "days": 21}, {"hours": 128, "days": 31}, "days", 47)
    assert (d1, d2) == (34, 13)
    place_both_days(db, days, d1)
    ok(db, 2026, "B/2026"); ok(db, 2027, "B/2027")
    nov = compute_month_summary(db, 1, 2026, 11)
    dec = compute_month_summary(db, 1, 2026, 12)
    assert nov["end_days"] == 10 and nov["prev_d_end"] == 31, \
        "B: ноябрь %s/%s" % (nov["end_days"], nov["prev_d_end"])
    assert dec["end_days"] == 0 and dec["end_hours"] == 4 * 60, \
        "B: декабрь текущего %s/%s (ожидали 0 д и 4 ч)" % (dec["end_days"], dec["end_hours"])
    assert dec["comp_d_real"] == 10 and dec["comp_h_real"] == 13 * 480, \
        "B: компенсировано текущего %s/%s (10 дн + 104 ч = 34 дня)" % (dec["comp_d_real"], dec["comp_h_real"])
    assert dec["prev_d_end"] == 18 and dec["prev_h_end"] == 128 * 60, \
        "B: декабрь заначки %s/%s (ожидали 18 д и 128 ч)" % (dec["prev_d_end"], dec["prev_h_end"])
    assert dec["comp_d_prev"] == 13 and dec["comp_h_prev"] == 0, \
        "B: компенсировано заначки %s/%s" % (dec["comp_d_prev"], dec["comp_h_prev"])
    jan = compute_month_summary(db, 1, 2027, 1)
    assert jan["end_days"] == 0 and jan["prev_d_start"] == 0, \
        "B: январь-2027 %s/%s" % (jan["end_days"], jan["prev_d_start"])
    rows = db.conn.execute("SELECT COUNT(*) c FROM comp_day_off_date").fetchone()["c"]
    assert rows == 47, "B: в календаре %s дней (ожидали 47)" % rows
    recs = db.conn.execute(
        "SELECT event_date, amount_days FROM compensation ORDER BY id").fetchall()
    assert len(recs) == 2 and recs[0]["event_date"] == "2026-11-16" \
        and recs[0]["amount_days"] == 34 and recs[1]["event_date"] == "1900-01-01" \
        and recs[1]["amount_days"] == 13, [dict(r) for r in recs]
    print("B: 47 дней за оба года — 34 текущим + 13 заначкой, остатки сходятся ✓")

    # ── C. Запрос больше общего пула — блокируется ──
    db = fresh_db("2026-01", opening_minutes=108 * 60, opening_days=21,
                  prev_minutes=128 * 60, prev_days=31)
    place_both_days(db, working_days(date(2026, 11, 16), 82), 34)
    valid, err = validate_non_negative_over_year(db, 1, 2026)
    assert not valid and "(Ночные)" in err, "C: %s" % err
    print("C: 82 дня при пуле 81 — блокируется по ночным часам ✓")

    # ── D. Заначка пуста — всё текущим годом ──
    db = fresh_db("2026-01", opening_minutes=108 * 60, opening_days=21)
    d1, d2 = split_both_years(
        {"hours": 108, "days": 21}, {"hours": 0, "days": 0}, "days", 20)
    assert (d1, d2) == (20, 0)
    place_both_days(db, working_days(date(2026, 11, 16), 20), d1)
    ok(db, 2026, "D/2026")
    dec = compute_month_summary(db, 1, 2026, 12)
    assert dec["end_days"] == 1 and dec["end_hours"] == 108 * 60, \
        "D: декабрь %s/%s" % (dec["end_days"], dec["end_hours"])
    n = db.conn.execute("SELECT COUNT(*) c FROM compensation").fetchone()["c"]
    assert n == 1, "D: запись должна быть одна, а не %s" % n
    print("D: заначка пуста — один период текущего года ✓")

    # ── E. Текущий пуст — всё заначкой ──
    db = fresh_db("2026-01", prev_minutes=0, prev_days=31)
    d1, d2 = split_both_years(
        {"hours": 0, "days": 0}, {"hours": 0, "days": 31}, "days", 20)
    assert (d1, d2) == (0, 20)
    place_both_days(db, working_days(date(2026, 11, 16), 20), d1)
    ok(db, 2026, "E/2026")
    dec = compute_month_summary(db, 1, 2026, 12)
    assert dec["prev_d_end"] == 11, "E: заначка в декабре %s (ожидали 11)" % dec["prev_d_end"]
    rec = db.conn.execute(
        "SELECT event_date, amount_days FROM compensation").fetchone()
    assert rec["event_date"] == "1900-01-01" and rec["amount_days"] == 20, dict(rec)
    print("E: текущий пуст — всё уходит в заначку ✓")

    # ── F. Часы «за оба года» ──
    assert split_both_years({"hours": 2}, {"hours": 3}, "hours", 300) == (120, 180)
    db = fresh_db("2026-01", opening_minutes=120, prev_minutes=180)
    for ev, minutes in (("2026-05-20", 120), ("1900-01-01", 180)):
        db.conn.execute(
            "INSERT INTO compensation(employee_id,unit,method,event_date,order_date,amount_minutes) "
            "VALUES (1,'hours','day_off',?,?,?)", (ev, "2026-05-20", minutes))
    ok(db, 2026, "F/2026")
    may = compute_month_summary(db, 1, 2026, 5)
    assert may["end_hours"] == 0 and may["prev_h_end"] == 0, \
        "F: май %s/%s (ожидали 0/0)" % (may["end_hours"], may["prev_h_end"])
    print("F: 5 часов за оба года (2 + 3) — обе копилки обнулены ✓")

    print("═══ КОМПЕНСАЦИЯ «ЗА ОБА ГОДА»: ДЕЛИТСЯ И СПИСЫВАЕТСЯ ВЕРНО ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
