#!/usr/bin/env python3
"""Тест нормы выработки с даты приема (неполный месяц приема).

Сценарий пользователя: сменщик принят 25-го числа, а норма считалась
за весь месяц — огромная «недоработка», которой по факту нет.
Решение (вариант А): норма месяца приема считается по
производственному календарю, но С ДАТЫ ПРИЕМА — человек обязан
отработать только рабочие дни, идущие после приёма.

Проверяется:
  A. Норма полного месяца (приём 1-го числа) не изменилась —
     по будням 8 ч, без праздников в тестовом календаре;
  B. Приём 25.03.2026 — норма только за хвост месяца;
  C. Приём в выходной (15.02.2026, воскресенье) — с понедельника;
  D. Месяц до приёма — итог по-прежнему нули;
  E. Сменщик, принятый 25-го: с двумя сменами переработка считается
     против маленькой нормы, а не всего месяца;
  F. Старая запись без hire_date (NULL) — фолбэк на 1-е число
     месяца, норма полная;
  G. Норма показывается обрезанной и в итоге месяца (compute_month_summary).

Запуск:
    python3 qa/test_hire_norm.py             # из корня репозитория
"""
import os
import sys
import tempfile
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from database import DB
from logic import compute_month_norm_minutes, compute_month_summary, build_shift_checker
import logic


def fresh_db():
    path = os.path.join(tempfile.mkdtemp(prefix="norm_"), "test.sqlite")
    db = DB(path)
    return db


def expected_norm(year, month, from_day):
    """Ожидаемая норма: будни с from_day до конца месяца по 8 часов
    (в свежей базе нет праздников и правок календаря)."""
    from calendar import monthrange
    total = 0
    for d in range(from_day, monthrange(year, month)[1] + 1):
        if date(year, month, d).weekday() < 5:
            total += 480
    return total


def main() -> int:
    # ── A. Полный месяц: приём 1-го числа ──
    db = fresh_db()
    db.add_employee("Первого", "Числа", "", "", "", "2026-03-01", 0, 0, 0, 0, 0, 0)
    got = compute_month_norm_minutes(db, 1, 2026, 3, lambda d: True)
    assert got == expected_norm(2026, 3, 1), "A: %s != %s" % (got, expected_norm(2026, 3, 1))
    print("A: приём 1-го — норма полного месяца (%d мин) ✓" % got)

    # ── B. Приём 25-го ──
    db = fresh_db()
    db.add_employee("Двадцать", "Пятого", "", "", "", "2026-03-25", 0, 0, 0, 0, 0, 0)
    got = compute_month_norm_minutes(db, 1, 2026, 3, lambda d: True)
    want = expected_norm(2026, 3, 25)
    assert got == want, "B: %s != %s" % (got, want)
    full = expected_norm(2026, 3, 1)
    assert got < full, "B: норма должна быть меньше полной"
    print("B: приём 25.03 — норма хвоста месяца (%d мин вместо %d) ✓" % (got, full))

    # ── C. Приём в выходной ──
    db = fresh_db()
    assert date(2026, 2, 15).weekday() == 6, "15.02.2026 должно быть воскресеньем"
    db.add_employee("Воскресный", "Приём", "", "", "", "2026-02-15", 0, 0, 0, 0, 0, 0)
    got = compute_month_norm_minutes(db, 1, 2026, 2, lambda d: True)
    assert got == expected_norm(2026, 2, 16), "C: %s" % got
    print("C: приём в воскресенье — норма с понедельника ✓")

    # ── D. Месяц до приёма — нули ──
    logic._SUMMARY_CACHE["data"].clear()
    s = compute_month_summary(db, 1, 2026, 1)
    assert s["norm_minutes"] == 0 and s["end_days"] == 0, "D: %s" % s["norm_minutes"]
    print("D: месяц до приёма — итог нули, как раньше ✓")

    # ── E. Сменщик, принятый 25-го: две смены против маленькой нормы ──
    db = fresh_db()
    gid = db.add_group("Сменный", is_shift=True)
    db.add_employee("Сменщик", "Ночной", "", "", "", "2026-03-25",
                    0, 0, 0, 0, 0, 0, group_id=gid)
    assert build_shift_checker(db, 1)(date(2026, 3, 26))
    # две суточных смены по 26.03 и 27.03 (по 24 часа со сменой в 8:00)
    db.add_duty(1, datetime(2026, 3, 26, 8, 0), datetime(2026, 3, 27, 8, 0), "", True)
    db.add_duty(1, datetime(2026, 3, 27, 8, 0), datetime(2026, 3, 28, 8, 0), "", True)
    logic._SUMMARY_CACHE["data"].clear()
    s = compute_month_summary(db, 1, 2026, 3)
    tail = expected_norm(2026, 3, 25)
    assert s["norm_minutes"] == tail, "E: норма %s != %s" % (s["norm_minutes"], tail)
    # переработка = смены (2×24 ч) − норма хвоста, а не всего месяца
    full = expected_norm(2026, 3, 1)
    assert s["end_overtime"] == 2 * 24 * 60 - tail, \
        "E: переработка %s (ожидали смены − норма хвоста = %s)" % (s["end_overtime"], 2 * 24 * 60 - tail)
    assert s["end_overtime"] > 2 * 24 * 60 - full, "E: раньше было бы ещё меньше"
    print("E: сменщик с 25-го — норма хвоста (%d), переработка против неё ✓"
          % s["norm_minutes"])

    # ── F. Старая запись без hire_date — фолбэк на 1-е число ──
    db = fresh_db()
    db.add_employee("Старая", "Запись", "", "", "", "2026-03", 0, 0, 0, 0, 0, 0)
    db.conn.execute("UPDATE employee SET hire_date=NULL WHERE id=1")
    got = compute_month_norm_minutes(db, 1, 2026, 3, lambda d: True)
    assert got == expected_norm(2026, 3, 1), "F: %s" % got
    print("F: hire_date пуст — норма полного месяца (фолбэк на 1-е) ✓")

    # ── G. Пятидневщик: норма — понятие сменщика, здесь 0 и раньше ──
    db = fresh_db()
    db.add_employee("Обычный", "Ежедневщик", "", "", "", "2026-03-25", 0, 0, 0, 0, 0, 0)
    logic._SUMMARY_CACHE["data"].clear()
    s = compute_month_summary(db, 1, 2026, 3)
    assert s["norm_minutes"] == 0, \
        "G: у пятидневщика норма и раньше была 0, а не %s" % s["norm_minutes"]
    print("G: пятидневщик — норма по-прежнему 0 (это понятие сменщика) ✓")

    print("═══ НОРМА С ДАТЫ ПРИЕМА: НЕПОЛНЫЙ МЕСЯЦ СЧИТАЕТСЯ ВЕРНО ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
