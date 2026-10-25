#!/usr/bin/env python3
"""Тест пересечения дней компенсаций (UNIQUE в comp_day_off_date).

Сценарий: день сотрудника уже входит в одну компенсацию, и при
сохранении новой (период или одиночный день) с тем же днём SQLite
падал сырой ошибкой «UNIQUE constraint failed: …comp_day_off_date».
Теперь сохранение заранее находит пересечения (DB.find_day_conflicts)
и показывает понятное сообщение (format_day_conflicts): с какими
днями пересеклись и какой компенсацией те заняты.

Проверяется:
  A. Период занял дни — новый период с пересечением видит конфликт
     (даты, границы и размер чужой записи, метка «за пред. год»);
  B. exclude_comp_id: свои дни не мешают (редактирование своей записи);
  C. Непересекающиеся дни — конфликтов нет;
  D. Одиночный день на занятую дату — «уже входит в компенсацию…»;
  E. Сырая вставка дубля по-прежнему невозможна (UNIQUE работает);
  F. Сообщение: больше трёх дат — «и ещё N», две группы — через «;».

Запуск:
    python3 qa/test_day_conflict.py             # из корня репозитория
"""
import os
import sqlite3
import sys
import tempfile
from datetime import date, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from database import DB, format_day_conflicts


def fresh_db():
    path = os.path.join(tempfile.mkdtemp(prefix="conflict_"), "test.sqlite")
    db = DB(path)
    db.add_employee("Тестов", "Тест", "Тестович", "капитан", "инженер",
                    "2025-01", 0, 0, 0, 0, 0, 0)
    return db


def place_period(db, start, n_days, use_prev_year):
    """Период компенсации, как пишет saveCompensation."""
    dates = [start + timedelta(days=i) for i in range(n_days)]
    event = "1900-01-01" if use_prev_year else dates[0].isoformat()
    cur = db.conn.execute(
        "INSERT INTO compensation(employee_id,unit,method,amount_days,comment,event_date,order_date) "
        "VALUES (1,'days','day_off',?,?,?,?)",
        (len(dates), None, event, dates[0].isoformat()))
    cid = cur.lastrowid
    for d in dates:
        db.conn.execute(
            "INSERT INTO comp_day_off_date(compensation_id,employee_id,day_off_date) VALUES (?,?,?)",
            (cid, 1, d.isoformat()))
    return cid


def iso(d):
    return d.isoformat()


def main() -> int:
    # ── A. Пересечение периодов ──
    db = fresh_db()
    cid = place_period(db, date(2026, 11, 16), 47, use_prev_year=True)  # 16.11.2026 – 01.01.2027
    new_dates = [iso(date(2026, 12, 20) + timedelta(days=i)) for i in range(10)]  # 20–29.12
    conflicts = db.find_day_conflicts(1, new_dates)
    assert len(conflicts) == 10, "A: конфликтов %s (ожидали 10)" % len(conflicts)
    c0 = conflicts[0]
    assert c0["date"] == "2026-12-20" and c0["comp_id"] == cid
    assert c0["start"] == "2026-11-16" and c0["end"] == "2027-01-01", c0
    assert c0["days"] == 47 and c0["prev_year"] is True
    msg = format_day_conflicts(conflicts)
    assert "Дни 20.12.2026, 21.12.2026, 22.12.2026 и ещё 7" in msg, msg
    assert "за период 16.11.2026 – 01.01.2027 (47 дн., за пред. год)" in msg, msg
    print("A: пересечение периодов — все 10 дней с границами чужой записи ✓")

    # ── B. Свои дни не мешают ──
    own = [iso(date(2026, 11, 16) + timedelta(days=i)) for i in range(5)]
    assert db.find_day_conflicts(1, own, exclude_comp_id=cid) == [], "B: свои дни мешают"
    print("B: exclude_comp_id — свои дни не считаются конфликтом ✓")

    # ── C. Без пересечения ──
    free = [iso(date(2027, 5, 10) + timedelta(days=i)) for i in range(3)]
    assert db.find_day_conflicts(1, free) == [], "C: ложный конфликт"
    print("C. свободные дни — конфликтов нет ✓")

    # ── D. Одиночный день на занятую дату ──
    conflicts = db.find_day_conflicts(1, ["2026-11-20"])
    assert len(conflicts) == 1
    msg = format_day_conflicts(conflicts)
    assert msg == ("День 20.11.2026 уже входит в компенсацию за период "
                   "16.11.2026 – 01.01.2027 (47 дн., за пред. год)."), msg
    print("D: одиночный день — внятное сообщение с периодом ✓")

    # ── E. Схема по-прежнему не пускает дубль ──
    try:
        db.conn.execute(
            "INSERT INTO comp_day_off_date(compensation_id,employee_id,day_off_date) VALUES (?,?,?)",
            (999, 1, "2026-11-20"))
        raise AssertionError("E: дубль прошёл — UNIQUE сломан")
    except sqlite3.IntegrityError:
        pass
    print("E: дубль в базу по-прежнему не лезет (UNIQUE на месте) ✓")

    # ── F. Две группы и сокращение списка дат ──
    db2 = fresh_db()
    place_period(db2, date(2026, 3, 2), 3, use_prev_year=False)    # 02–04.03, этот год
    place_period(db2, date(2026, 4, 6), 2, use_prev_year=True)     # 06–07.04, за пред. год
    conflicts = db2.find_day_conflicts(
        1, [iso(date(2026, 3, d)) for d in (2, 3, 4)] + ["2026-04-06"])
    assert len(conflicts) == 4, conflicts
    msg = format_day_conflicts(conflicts)
    first, second = msg.split(";\n")
    assert "02.03.2026, 03.03.2026, 04.03.2026" in first and "(3 дн.)" in first, first
    assert "за пред. год" not in first, first   # у этого года метки нет
    assert "06.04.2026 уже входит в компенсацию за период 06.04.2026 – 07.04.2026 (2 дн., за пред. год)" in second, second
    print("F: две группы через «;», сокращение списка дат ✓")

    print("═══ ПЕРЕСЕЧЕНИЕ ДНЕЙ КОМПЕНСАЦИЙ: ПОНЯТНО И БЕЗ падения SQLITE ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
