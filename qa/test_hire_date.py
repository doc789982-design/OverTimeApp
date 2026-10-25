#!/usr/bin/env python3
"""Тест даты приема вместо месяца приема (миграция + поведение).

Месяц приема («ГГГГ-ММ» в employee.start_month) вырос в дату приема:
новое поле employee.hire_date («ГГГГГ-ММ-ДД»). Требование: у существующих
баз при обновлении у всех записей появляется 1-е число их месяца и НИЧЕГО
не пересчитывается — все расчёты в программе продолжают работать по
месяцу (start_month остаётся столбцом «ГГГГ-ММ», все строковые сравнения
не меняются).

Проверяется:
  A. Новая база: поле hire_date в схеме, add_employee пишет оба поля;
  B. add_employee принимает и дату («2026-03-25»), и старый месяц
     («2026-03» — день 1-е число);
  C. Старая база без столбца: при открытии столбец добавляется и
     заполняется 1-м числом месяца; повторное открытие ничего не меняет;
  D. update_employee согласует месяц и дату при любом вызове;
  E. Ничего не пересчиталось: итоги месяца до и после миграции
     совпадают; список месяца включает принятого в этом же месяце
     (даже с датой 25-го числа) и не включает принятого позже;
  F. Ошибка валидации при кривом формате.

Запуск:
    python3 qa/test_hire_date.py             # из корня репозитория
"""
import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from database import DB
from logic import compute_month_summary
import logic


def fresh_path():
    return os.path.join(tempfile.mkdtemp(prefix="hire_"), "test.sqlite")


def main() -> int:
    # ── A/B. Новая база: оба поля, оба формата ──
    path = fresh_path()
    db = DB(path)
    cols = [r["name"] for r in db.conn.execute("PRAGMA table_info(employee)")]
    assert "hire_date" in cols, "нет столбца hire_date: %s" % cols
    db.add_employee("Старый", "Формат", "", "", "", "2025-01",
                    0, 0, 0, 0, 0, 0)
    db.add_employee("Новый", "Формат", "", "", "", "2026-03-25",
                    0, 0, 0, 0, 0, 0)
    rows = {r["last_name"]: r for r in db.conn.execute(
        "SELECT last_name, start_month, hire_date FROM employee")}
    assert rows["Старый"]["start_month"] == "2025-01" \
        and rows["Старый"]["hire_date"] == "2025-01-01", dict(rows["Старый"])
    assert rows["Новый"]["start_month"] == "2026-03" \
        and rows["Новый"]["hire_date"] == "2026-03-25", dict(rows["Новый"])
    print("A/B: схема с hire_date, месячный и дневной форматы записи ✓")

    # ── C. Старая база без столбца: миграция при открытии ──
    db.conn.commit()
    db.conn.execute("ALTER TABLE employee DROP COLUMN hire_date")
    db.conn.commit()
    db.close()
    db = DB(path)                     # открытие = миграция
    rows = list(db.conn.execute(
        "SELECT start_month, hire_date FROM employee ORDER BY id"))
    assert rows[0]["start_month"] == "2025-01" and rows[0]["hire_date"] == "2025-01-01", \
        dict(rows[0])
    assert rows[1]["start_month"] == "2026-03" and rows[1]["hire_date"] == "2026-03-01", \
        "миграция должна дать 1-е число месяца: %s" % dict(rows[1])
    db.close()
    db = DB(path)                     # идемпотентность
    rows2 = list(db.conn.execute(
        "SELECT start_month, hire_date FROM employee ORDER BY id"))
    assert rows2 == rows, "повторное открытие что-то изменило"
    print("C: старая база — hire_date = 1-е число месяца, идемпотентно ✓")

    # ── D. update_employee согласует поля ──
    db.update_employee(1, hire_date="2027-05-14")
    r = db.conn.execute("SELECT start_month, hire_date FROM employee WHERE id=1").fetchone()
    assert r["start_month"] == "2027-05" and r["hire_date"] == "2027-05-14", dict(r)
    db.update_employee(1, start_month="2028-02")
    r = db.conn.execute("SELECT start_month, hire_date FROM employee WHERE id=1").fetchone()
    assert r["start_month"] == "2028-02" and r["hire_date"] == "2028-02-01", dict(r)
    print("D: обновление любым из полей согласует оба ✓")

    # ── E. Ничего не пересчиталось ──
    # старая база с данными: сотрудник-сменщик с переработкой и компенсацией
    path2 = fresh_path()
    db2 = DB(path2)
    db2.add_employee("Итоги", "Проверка", "", "", "", "2026-01",
                     480, 2, 0, 0, 0, 0)
    db2.conn.execute(
        "INSERT INTO compensation(employee_id,unit,method,amount_days,event_date,order_date) "
        "VALUES (1,'days','day_off',1,'2026-05-10','2026-05-10')")
    db2.conn.execute(
        "INSERT INTO comp_day_off_date(compensation_id,employee_id,day_off_date) "
        "VALUES (1,1,'2026-05-10')")
    logic._SUMMARY_CACHE["data"].clear()
    before = compute_month_summary(db2, 1, 2026, 5)
    # «обновление программы»: имитируем старую базу и миграцию
    db2.conn.commit()
    db2.conn.execute("ALTER TABLE employee DROP COLUMN hire_date")
    db2.conn.commit()
    db2.close()
    db2 = DB(path2)
    logic._SUMMARY_CACHE["data"].clear()
    after = compute_month_summary(db2, 1, 2026, 5)
    assert before == after, "итоги изменились после миграции!\n%s\n%s" % (before, after)
    # принятый в этом же месяце (25-го!) виден в списке месяца
    db2.add_employee("Пятый", "Число", "", "", "", "2026-05-25", 0, 0, 0, 0, 0, 0)
    names = [r["last_name"] for r in db2.list_employees_for_month(2026, 5, True, "")]
    assert "Пятый" in names and "Итоги" in names, names
    # принятый в следующем месяце — не виден
    db2.add_employee("Будущий", "Месяц", "", "", "", "2026-06-01", 0, 0, 0, 0, 0, 0)
    names = [r["last_name"] for r in db2.list_employees_for_month(2026, 5, True, "")]
    assert "Будущий" not in names, names
    print("E: итоги до/после миграции совпадают, списки месяца верны ✓")

    # ── F. Кривой формат ──
    try:
        db2.add_employee("Кривой", "Формат", "", "", "", "март 2026",
                         0, 0, 0, 0, 0, 0)
        raise AssertionError("кривой формат прошёл")
    except ValueError:
        pass
    print("F: кривой формат даты отклоняется ✓")

    print("═══ ДАТА ПРИЕМА ВМЕСТО МЕСЯЦА: МИГРАЦИЯ И ПОВЕДЕНИЕ ЦЕЛЫ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
