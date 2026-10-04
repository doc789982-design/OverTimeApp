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
     удаление приказа целиком; страховки бэкенда (пустой номер,
     битая дата, пустой список получателей);
  D. Структура QML: окно-ведомость (галочка «все», одиночный
     режим из карточки, обязательные №/дата, карточки-строки,
     «сверх.» всегда в «Доступно», журнал), старое окно удалено.

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

    # ── страховки: пустой номер, битая дата, пустой список ──
    emps = b.moneyOrderEmployees()
    rows = [{"id": e["id"], "hours": e["hours"], "overtime": 0,
             "days": e["days"]} for e in emps]
    bad = b.saveMoneyOrder(json.dumps(rows), "  ", "2026-10-12", "")
    assert bad["ok"] is False and "номер" in bad["errors"][0]["message"], bad
    bad = b.saveMoneyOrder(json.dumps(rows), "245", "12.10.2026", "")
    assert bad["ok"] is False and "дату" in bad["errors"][0]["message"], bad
    bad = b.saveMoneyOrder("[]", "245", "2026-10-12", "")
    assert bad["ok"] is False and "суммы" in bad["errors"][0]["message"], bad
    bad = b.saveMoneyOrder(
        json.dumps([{"id": emps[0]["id"], "hours": 0, "overtime": 0, "days": 0}]),
        "245", "2026-10-12", "")
    assert bad["ok"] is False, "нулевые суммы не должны проводиться"
    print("B0: пустой №/дата/список отклоняются бэкендом ✓")

    # ── успешный приказ «по табелю»: каждому сколько есть ──
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
    # быстрое заполнение: кнопки-глаголы
    assert "По табелю" in dlg and "Одинаково всем" in dlg
    assert "Раздать выбранным" in dlg
    # тряска — ШТАТНАЯ AppDialog.shake(), не самописная анимация
    assert "root.shake()" in dlg and "shakeAnim" not in dlg
    assert "root.scrollToBottom()" in dlg
    # подсветка строк с ошибкой
    assert "rowErrors" in dlg and "bgDangerSoft" in dlg
    assert "accentDanger" in dlg
    # грабли сборки 257: ListView и Layout-и в contentArea AppDialog
    # дают пустой список и разъехавшиеся подписи — только прямые дети
    # (Repeater в Column, Row с явными ширинами)
    assert "ListView" not in dlg, "ListView в AppDialog — список пустой"
    assert "RowLayout" not in dlg and "ColumnLayout" not in dlg, \
        "Layout-и растягивают поля — подписи разъезжаются"
    assert "Repeater" in dlg and "colNum" in dlg
    # сборка 259: галочка «все», одиночный режим, обязательные №/дата
    assert "openForEmployee" in dlg and "singleMode" in dlg
    assert "toggleRow" in dlg, "переключение строки живёт в корне (не в делегате)"
    assert "Выбрать всех или снять выделение" in dlg
    assert "Укажите номер и дату приказа" in dlg and "hasError" in dlg
    assert "Журнал приказов" in dlg and "requestJournal" in dlg
    # «Доступно»: сверхурочные видны ВСЕГДА (без условия >0)
    assert "сверх." in dlg and "balOvertime > 0" not in dlg
    # строки — карточки как числа месяца, без наведения
    assert "AppTheme.bgCell" in dlg
    # required index у делегата (иначе root.rows[index] падает)
    assert "required property int index" in dlg
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
    assert "onRequestJournal" in main, "журнал не подключён к ведомости"
    summ = open(os.path.join(ROOT, "components", "AppSummaryPanel.qml"),
                encoding="utf-8").read()
    assert "openForEmployee" in summ, \
        "кнопка «Деньгами» в карточке не открывает приказ этому сотруднику"
    # старое одиночное окно удалено
    assert not os.path.exists(os.path.join(ROOT, "components", "MoneyDialog.qml"))
    backend_src = open(os.path.join(ROOT, "Main.py"), encoding="utf-8").read()
    assert "saveMoneyCompList" not in backend_src
    assert "loadMoneyComps" not in backend_src and "deleteMoneyComp" not in backend_src
    print("D: окно-ведомость (Repeater + штатная тряска + подсветка), "
          "журнал приказов, кнопки, старое окно удалено ✓")


def main() -> int:
    part_a_employees()
    part_bc_order_and_journal()
    part_d_structure()
    print("═══ ВЕДОМОСТЬ: ОДИН ПРИКАЗ — ВСЕМ СОТРУДНИКАМ СРАЗУ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
