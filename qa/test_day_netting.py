#!/usr/bin/env python3
"""Тест закрытия дней компенсации ночными часами (8 ч = день).

Сценарий пользователя: на 16.11.2026 у человека 128 ч и 31 д за 25 год
(итого 47 дней по правилу «дни + часы/8»). Период 47 дней с 16.11.2026
по 28.01.2027 «за пред. год» окно пропускало («Итого: 47 дн.»), а
проверка бэкенда считала копилки раздельно и в декабре ругалась
«превышен лимит остатков прошлого года (Дни)». Теперь дни сверх
дневной копилки списываются из ночных часов — обе стороны считают
одинаково, остатки и графа «Компенсировано» сходятся арифметически.

Проверяется:
  A. Случай пользователя: 47 дней с 16.11 «за пред. год» — ставится,
     заначка обнуляется (декабрь: 0 д и 0 ч), 2027-й чистый;
  B. 48 дней при той же заначке — блокируется по ночным часам;
  C. Текущий год: 34 дня при 21 д и 108 ч — ставится (21 дн + 13 дн
     из часов), в декабре 4 ч остатка;
  D. Текущий год, часов не хватает, — блокируется «ночных часов»;
  E. Дней достаточно — часы не трогаются (как раньше);
  F. Заначка только в часах: 16 дней из 130 ч, остаток 2 ч;
     17-й день не лезет (переводится целыми днями по 8 ч).

Запуск:
    python3 qa/test_day_netting.py             # из корня репозитория
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


def fresh_db(hire, opening_minutes=0, opening_days=0):
    """База с одним сотрудником; вводные остатки года приёма."""
    path = os.path.join(tempfile.mkdtemp(prefix="netting_"), "test.sqlite")
    db = DB(path)
    db.add_employee("Тестов", "Тест", "Тестович", "капитан", "инженер",
                    hire, opening_minutes, opening_days, 0, 0, 0, 0)
    logic._SUMMARY_CACHE["data"].clear()
    return db


def working_days(start, n):
    """Первые n будних дней (пн–пт) от start — как «пропуск нерабочих»."""
    out, cur = [], start
    while len(out) < n:
        if cur.weekday() < 5:
            out.append(cur)
        cur += timedelta(days=1)
    return out


def place_days_period(db, dates, use_prev_year):
    """Повторяет запись saveCompensation: родитель + дни периода."""
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
    # ── A. Случай пользователя: 128 ч + 31 д за 25 год, 47 дней с 16.11 ──
    db = fresh_db("2025-01", opening_minutes=128 * 60, opening_days=31)
    days = working_days(date(2026, 11, 16), 47)      # 11 в ноябре, 23 в декабре, 13 в январе
    assert sum(1 for d in days if d.year == 2026 and d.month == 11) == 11
    assert sum(1 for d in days if d.year == 2026 and d.month == 12) == 23
    assert days[-1].year == 2027
    place_days_period(db, days, use_prev_year=True)
    ok(db, 2026, "A/2026"); ok(db, 2027, "A/2027")
    nov = compute_month_summary(db, 1, 2026, 11)
    dec = compute_month_summary(db, 1, 2026, 12)
    assert nov["prev_d_end"] == 20 and nov["prev_h_end"] == 128 * 60, \
        "A: ноябрь %s/%s" % (nov["prev_d_end"], nov["prev_h_end"])
    assert dec["prev_d_end"] == 0 and dec["prev_h_end"] == 0, \
        "A: декабрь %s/%s (ожидали 0/0)" % (dec["prev_d_end"], dec["prev_h_end"])
    assert dec["comp_d_prev"] == 20, "A: дней в декабре %s (ожидали 20 = 31 − 11)" % dec["comp_d_prev"]
    assert dec["comp_h_prev"] == 128 * 60, "A: часов в декабре %s (ожидали 128 ч)" % dec["comp_h_prev"]
    jan = compute_month_summary(db, 1, 2027, 1)
    assert jan["prev_d_start"] == 0 and jan["prev_h_start"] == 0, "A: заначка-2027 не пуста"
    total = db.conn.execute("SELECT COUNT(*) c FROM comp_day_off_date").fetchone()["c"]
    assert total == 47, "A: в календаре %s дней (ожидали 47)" % total
    print("A: 47 дней с 16.11 из заначки 128 ч + 31 д — ставится целиком, оба года чистые ✓")

    # ── B. 48 дней при той же заначке — не хватает ночных часов ──
    db = fresh_db("2025-01", opening_minutes=128 * 60, opening_days=31)
    place_days_period(db, working_days(date(2026, 11, 16), 48), use_prev_year=True)
    valid, err = validate_non_negative_over_year(db, 1, 2026)
    assert not valid and "(Ночные)" in err, "B: %s" % err
    print("B: 48-й день сверх заначки — блокируется по ночным часам ✓")

    # ── C. Текущий год: 21 д и 108 ч, период 34 дней с 16.11 ──
    db = fresh_db("2026-01", opening_minutes=108 * 60, opening_days=21)
    place_days_period(db, working_days(date(2026, 11, 16), 34), use_prev_year=False)
    ok(db, 2026, "C/2026"); ok(db, 2027, "C/2027")
    nov = compute_month_summary(db, 1, 2026, 11)
    dec = compute_month_summary(db, 1, 2026, 12)
    assert nov["end_days"] == 10 and nov["end_hours"] == 108 * 60, \
        "C: ноябрь %s/%s" % (nov["end_days"], nov["end_hours"])
    assert dec["end_days"] == 0 and dec["end_hours"] == 4 * 60, \
        "C: декабрь %s/%s (ожидали 0 д и 4 ч)" % (dec["end_days"], dec["end_hours"])
    assert dec["comp_d_real"] == 10, "C: дней в декабре %s (ожидали 10)" % dec["comp_d_real"]
    assert dec["comp_h_real"] == 13 * 480, "C: часов в декабре %s (ожидали 13 дн × 8 ч)" % dec["comp_h_real"]
    print("C: текущий год 21 д + 108 ч — 34 дня ставятся (13 дней из часов), остаток 4 ч ✓")

    # ── D. Часов не хватает — блокируется ──
    db = fresh_db("2026-01", opening_minutes=16 * 60, opening_days=21)
    place_days_period(db, working_days(date(2026, 11, 16), 30), use_prev_year=False)
    valid, err = validate_non_negative_over_year(db, 1, 2026)
    assert not valid and "ночных часов" in err, "D: %s" % err
    print("D: 30 дней при 21 д и 16 ч — блокируется «не хватает ночных часов» ✓")

    # ── E. Дней достаточно — часы не трогаются ──
    db = fresh_db("2026-01", opening_minutes=108 * 60, opening_days=21)
    place_days_period(db, working_days(date(2026, 11, 16), 15), use_prev_year=False)
    ok(db, 2026, "E/2026")
    dec = compute_month_summary(db, 1, 2026, 12)
    assert dec["end_days"] == 6 and dec["end_hours"] == 108 * 60, \
        "E: декабрь %s/%s (ожидали 6 д и 108 ч)" % (dec["end_days"], dec["end_hours"])
    assert dec["comp_h_real"] == 0, "E: часы не должны тратиться: %s" % dec["comp_h_real"]
    print("E: дней хватает — списание только днями, как раньше ✓")

    # ── F. Заначка только в часах: целыми днями по 8 ч ──
    db = fresh_db("2025-01", opening_minutes=130 * 60, opening_days=0)
    place_days_period(db, working_days(date(2026, 12, 1), 16), use_prev_year=True)
    ok(db, 2026, "F/16")
    dec = compute_month_summary(db, 1, 2026, 12)
    assert dec["prev_d_end"] == 0 and dec["prev_h_end"] == 2 * 60, \
        "F: декабрь %s/%s (ожидали 0 д и 2 ч)" % (dec["prev_d_end"], dec["prev_h_end"])
    db2 = fresh_db("2025-01", opening_minutes=130 * 60, opening_days=0)
    place_days_period(db2, working_days(date(2026, 12, 1), 17), use_prev_year=True)
    valid, err = validate_non_negative_over_year(db2, 1, 2026)
    assert not valid and "(Ночные)" in err, "F: %s" % err
    print("F: часы переводятся в дни целиком (по 8 ч) — 16 дней из 130 ч можно, 17 нельзя ✓")

    print("═══ ДНИ СВЕРХ КОПИЛКИ ЗАКРЫВАЮТСЯ НОЧНЫМИ ЧАСАМИ: РАБОТАЕТ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
