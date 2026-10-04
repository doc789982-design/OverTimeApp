#!/usr/bin/env python3
"""Тест ведомости: один денежный приказ — всем сотрудникам сразу.

Приказ по подразделению: шапка (номер, дата, комментарий) вводится
один раз, у каждого сотрудника — свои часы/сверхурочные/дни; при
ошибке окно трясётся и подсвечивает строку виновника.

Проверяется:
  A. Список сотрудников с остатками (moneyOrderEmployees);
  B. Приказ «всем по табелю» одной транзакцией: записи у каждого
     получателя с общим номером и датой; ошибка (не хватает остатка)
     возвращает виновника поимённо и НЕ трогает базу;
  C. Журнал приказов: группировка по номеру+дате (×N), суммы,
     удаление приказа целиком;
  D. Структура QML: окно-ведомость с тряской и подсветкой строк,
     журнал, кнопки, старое одиночное окно удалено.

Запуск:
    python3 qa/test_money_order.py             # из корня репозитория
"""
import json
import os
import sys
import types
from datetime import date, datetime
from pathlib import Path

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

# Main.py импортирует win32print (только Windows) — в песочнице подменяем
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from database import DB
from Main import Backend
from PySide6.QtCore import QObject

OCT = 10


def fresh_db():
    path = os.path.join(os.environ.get("TMPDIR", "/tmp"),
                        "money_order_test.sqlite")
    if os.path.exists(path):
        os.remove(path)
    db = DB(path)
    db.add_employee("Иванов", "Иван", "", "капитан", "инженер",
                    "2020-01", 0, 0, 0, 0, 0, 0)
    db.add_employee("Петров", "Пётр", "", "лейтенант", "инженер",
                    "2020-01", 0, 0, 0, 0, 0, 0)
    # ночные часы появляются из дежурств: 20:00–08:00 на будний день
    db.add_duty(1, datetime(2026, 10, 7, 20, 0), datetime(2026, 10, 8, 8, 0), "")
    db.add_duty(2, datetime(2026, 10, 7, 20, 0), datetime(2026, 10, 8, 8, 0), "")
    db.conn.commit()
    return db


def make_backend(db):
    b = Backend.__new__(Backend)
    QObject.__init__(b)
    b.active_db = db
    b._selected_employee_id = 0
    b.current_year = 2026
    b.current_month = OCT
    b._money_orders = []
    b.refresh_calendar = lambda: None
    b.refresh_employees = lambda: None
    b._defer_year_refresh = lambda: None
    return b


def part_a_employees() -> None:
    db = fresh_db()
    b = make_backend(db)
    emps = b.moneyOrderEmployees()
    assert [e["id"] for e in emps] == [1, 2], emps
    assert "Иванов" in emps[0]["name"] and "инженер" in emps[0]["subtitle"]
    # ночные часы от дежурства — у обоих что-то есть
    by_id = {e["id"]: e for e in emps}
    assert by_id[1]["hours"] >= 4, by_id[1]
    assert by_id[2]["hours"] >= 4, by_id[2]
    print("A: сотрудники с остатками от дежурств ✓")


def part_bc_order_and_journal() -> None:
    db = fresh_db()
    b = make_backend(db)

    # ── успешный приказ «по табелю»: каждому сколько есть ──
    emps = b.moneyOrderEmployees()
    rows = [{"id": e["id"], "hours": e["hours"], "overtime": 0,
             "days": e["days"]} for e in emps]
    bal = {e["id"]: e["hours"] for e in emps}
    res = b.saveMoneyOrder(json.dumps(rows), "245", "2026-10-12", "За октябрь")
    assert res["ok"] is True, res
    assert res["count"] == 2, res
    n = db.conn.execute(
        "SELECT COUNT(*) c FROM compensation WHERE method='money'").fetchone()["c"]
    assert n == 2, n  # по одной записи с часами на сотрудника
    for eid in (1, 2):
        r = db.conn.execute(
            "SELECT order_no, order_date, amount_minutes FROM compensation "
            "WHERE method='money' AND employee_id=?", (eid,)).fetchone()
        assert r["order_no"] == "245" and r["order_date"] == "2026-10-12"
        assert r["amount_minutes"] == bal[eid] * 60, (eid, r["amount_minutes"])

    # ── журнал: один приказ — одна строка с двумя получателями ──
    b.loadMoneyOrders()
    orders = b._money_orders
    assert len(orders) == 1, orders
    o = orders[0]
    assert o["order_no"] == "245" and o["count"] == 2
    assert o["hours"] == sum(bal.values()), o
    assert {e["id"] for e in o["employees"]} == {1, 2}

    # ── ошибка: Петрову просим больше, чем есть; база не трогается ──
    before = db.conn.execute(
        "SELECT COUNT(*) c FROM compensation WHERE method='money'").fetchone()["c"]
    rows = [
        {"id": 1, "hours": 1, "overtime": 0, "days": 0},
        {"id": 2, "hours": bal[2] + 100, "overtime": 0, "days": 0},  # больше, чем есть
    ]
    res = b.saveMoneyOrder(json.dumps(rows), "246", "2026-10-13", "")
    assert res["ok"] is False, res
    assert len(res["errors"]) == 1, res
    err = res["errors"][0]
    assert err["id"] == 2 and "Петров" in err["name"], err
    assert "не хватает часов" in err["message"], err
    after = db.conn.execute(
        "SELECT COUNT(*) c FROM compensation WHERE method='money'").fetchone()["c"]
    assert before == after, "при ошибке база изменилась"

    # ── удаление приказа целиком ──
    b.deleteMoneyOrder("245", "2026-10-12")
    n = db.conn.execute(
        "SELECT COUNT(*) c FROM compensation WHERE method='money'").fetchone()["c"]
    assert n == 0, n
    b.loadMoneyOrders()
    assert b._money_orders == [], b._money_orders
    print("B: приказ по двум сотрудникам одной транзакцией; ошибка — поимённо, "
          "база не тронута; удаление приказа целиком ✓")


def part_d_structure() -> None:
    dlg = open(os.path.join(ROOT, "components", "MoneyOrderDialog.qml"),
               encoding="utf-8").read()
    # построчное редактирование: часы/сверхурочные/дни у каждого
    assert "root.rows[index].hours" in dlg and "root.rows[index].days" in dlg
    # быстрое заполнение
    assert "Всем" in dlg and "Каждому по табелю" in dlg
    # тряска окна и подсветка строк с ошибкой
    assert "shakeAnim" in dlg and "rowErrors" in dlg
    assert "bgDangerSoft" in dlg and "accentDanger" in dlg
    insp = open(os.path.join(ROOT, "components", "MoneyInspector.qml"),
                encoding="utf-8").read()
    assert "backend.moneyOrders" in insp
    assert "deleteMoneyOrder" in insp and "repeatMoneyOrder" in insp
    assert "×" in insp, "нет бейджа с числом получателей"
    left = open(os.path.join(ROOT, "components", "LeftControlPanel.qml"),
                encoding="utf-8").read()
    assert "moneyOrderDialog.openNew()" in left, "нет кнопки приказа в панели"
    main = open(os.path.join(ROOT, "main.qml"), encoding="utf-8").read()
    assert "MoneyOrderDialog" in main and "repeatMoneyOrder" in main
    summ = open(os.path.join(ROOT, "components", "AppSummaryPanel.qml"),
                encoding="utf-8").read()
    assert "loadMoneyOrders" in summ, "кнопка «Деньгами» не открывает журнал"
    # старое одиночное окно удалено
    assert not os.path.exists(os.path.join(ROOT, "components", "MoneyDialog.qml"))
    backend_src = open(os.path.join(ROOT, "Main.py"), encoding="utf-8").read()
    assert "saveMoneyCompList" not in backend_src
    assert "loadMoneyComps" not in backend_src and "deleteMoneyComp" not in backend_src
    print("C: окно-ведомость (построчно + тряска + подсветка), журнал приказов, "
          "кнопки, старое окно удалено ✓")


def main() -> int:
    part_a_employees()
    part_bc_order_and_journal()
    part_d_structure()
    print("═══ ВЕДОМОСТЬ: ОДИН ПРИКАЗ — ВСЕМ СОТРУДНИКАМ СРАЗУ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
