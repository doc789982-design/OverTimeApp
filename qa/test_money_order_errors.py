#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Свои суммы и подсветка виновника в денежной компенсации (сборка 303).

A. Бэкенд: в ошибке списания названо ПОЛЕ (часы/дни/сверхурочные),
   у одного сотрудника собираются несколько нехваток, виновников
   может быть несколько.
B. Ведомость: при ошибке поле виновника вспыхивает красной рамкой
   (как у текстовых полей формы: 3 секунды, потом плавно гаснет),
   повторная попытка вспыхивает снова; сообщения одного сотрудника
   собраны вместе; верхняя строка называет всех виновников.
C. Журнал приказов: карточка показывает сумму ВЫБРАННОГО сотрудника
   (с именем), если он получатель; не получал — общую по приказу.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_money_order_errors.py
"""
import json
import os
import sys
import time
import types
from datetime import datetime

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

# Main.py импортирует win32print (только Windows) — в песочнице подменяем
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from PySide6.QtCore import QObject, Property, Signal, Slot
from PySide6.QtQuick import QQuickItem  # регистрирует конвертер item*

from database import DB
from Main import Backend

OCT = 10


def fresh_db():
    path = os.path.join(os.environ.get("TMPDIR", "/tmp"),
                        "money_order_errors_test.sqlite")
    if os.path.exists(path):
        os.remove(path)
    db = DB(path)
    db.add_employee("Иванов", "Иван", "", "капитан", "инженер",
                    "2026-01", 0, 0, 0, 10 * 60, 6 * 60, 3)
    db.add_employee("Петров", "Пётр", "", "лейтенант", "инженер",
                    "2026-01", 0, 0, 0, 10 * 60, 6 * 60, 3)
    # ночные часы от дежурства 20:00–08:00
    db.add_duty(1, datetime(2026, 10, 7, 20, 0), datetime(2026, 10, 8, 8, 0), "")
    db.add_duty(2, datetime(2026, 10, 7, 20, 0), datetime(2026, 10, 8, 8, 0), "")
    db.conn.commit()
    return db


def make_backend(db):
    b = Backend.__new__(Backend)
    from PySide6.QtCore import QObject
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


def part_a_backend_fields() -> None:
    db = fresh_db()
    b = make_backend(db)
    emps = b.moneyOrderEmployees()
    bal = {e["id"]: e for e in emps}

    # у Иванова просим больше и часов, и сверхурочных — две нехватки,
    # каждая со своим полем
    rows = [
        {"id": 1, "hours": bal[1]["hours"] + 5,
         "overtime": bal[1]["overtime"] + 7, "days": 0},
        {"id": 2, "hours": bal[2]["days"] and 1 or 1, "overtime": 0, "days": 0},
    ]
    res = b.saveMoneyOrder(json.dumps(rows), "301", "2026-10-13", "", 0)
    assert res["ok"] is False, res
    errs = [e for e in res["errors"] if e["id"] == 1]
    assert len(errs) == 2, res["errors"]
    fields = {e.get("field") for e in errs}
    assert fields == {"hours", "overtime"}, fields

    # двое виновников: у каждого своя нехватка со своим полем
    rows = [
        {"id": 1, "hours": 1, "overtime": 0, "days": bal[1]["days"] + 2},
        {"id": 2, "hours": bal[2]["hours"] + 100, "overtime": 0, "days": 0},
    ]
    res = b.saveMoneyOrder(json.dumps(rows), "302", "2026-10-13", "", 0)
    assert res["ok"] is False, res
    by_id = {}
    for e in res["errors"]:
        by_id.setdefault(e["id"], []).append(e.get("field"))
    assert by_id.get(1) == ["days"], by_id
    assert by_id.get(2) == ["hours"], by_id
    print("A: в ошибке — поле, нехватки собираются, виновников несколько ✓")


# ────────────────────────────────────────────────────────────────
# B. Ведомость: вспышка поля виновника (настоящее окно)
# ────────────────────────────────────────────────────────────────

WRAPPER_B = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 1200; height: 800; visible: true
    AppUI.MoneyOrderDialog { id: dlg }
    Component.onCompleted: dlg.openNew()
}
"""

EMPS = [
    {"id": 1, "name": "Иванов Иван", "subtitle": "капитан — инженер",
     "hours": 8, "overtime": 4, "days": 2},
    {"id": 2, "name": "Петров Пётр Петрович", "subtitle": "",
     "hours": 4, "overtime": 0, "days": 0},
    {"id": 3, "name": "Сидоров Сидор", "subtitle": "",
     "hours": 0, "overtime": 0, "days": 1},
]


FAIL_COUNT = {"n": 0}


class StubOrderBackend(QObject):
    """Стаб уровня модуля (локальные классы PySide6 регистрирует
    ненадёжно — окно рендерится без данных)."""

    isDarkChanged = Signal()

    def __init__(self):
        super().__init__()
        self._dark = False

    def getDark(self):
        return self._dark

    def setDark(self, v):
        self._dark = bool(v)

    isDarkTheme = Property(bool, getDark, setDark, notify=isDarkChanged)

    @Slot(result="QVariant")
    def moneyOrderEmployees(self):
        return EMPS

    @Slot(str, str, str, str, int, result="QVariant")
    def saveMoneyOrder(self, payload, order_no, order_date_iso,
                       comment, source_mode):
        FAIL_COUNT["n"] += 1
        return {"ok": False, "errors": [
            {"id": 1, "name": "Иванов Иван", "field": "overtime",
             "message": "не хватает сверхурочных: доступно 4, запрошено 40"},
            {"id": 1, "name": "Иванов Иван", "field": "hours",
             "message": "не хватает часов: доступно 8, запрошено 80"},
            {"id": 3, "name": "Сидоров Сидор", "field": "days",
             "message": "не хватает дней: доступно 1, запрошено 5"},
        ]}


def part_b_dialog_flash(app) -> None:
    from PySide6.QtQml import QQmlApplicationEngine, QJSValue

    def to_var(v):
        return v.toVariant() if isinstance(v, QJSValue) else v

    def walk_items(item, out=None):
        """Визуальное дерево: делегаты Repeater НЕ попадают в
        QObject-дерево findChildren — только в childItems."""
        out = out if out is not None else []
        if not isinstance(item, QQuickItem):
            return out
        out.append(item)
        for ch in item.childItems():
            walk_items(ch, out)
        return out

    def all_items(win):
        res = []
        for w2 in win.findChildren(QObject):
            pass
        # contentItem окна + визуальные дети
        ci = win.contentItem() if hasattr(win, "contentItem") else None
        if ci is not None:
            walk_items(ci, res)
        return res

    stub = StubOrderBackend()
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()          # ссылка обязательна (GC)
    ctx.setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_test_money_errors.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER_B)
    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "окно не загрузилось"
        win = engine.rootObjects()[0]

        def settle(s):
            end = time.time() + s
            while time.time() < end:
                app.processEvents()
                time.sleep(0.01)
            for _ in range(30):
                app.processEvents()

        settle(0.6)
        dlg = next(o for o in win.findChildren(QObject)
                   if o.metaObject().className().startswith("MoneyOrderDialog"))

        # номер и дата приказа — валидные, чтобы дошло до сохранения.
        # Поля ищем в визуальном дереве (строки — делегаты Repeater)
        def find_fields():
            items = all_items(win)
            return [it for it in items
                    if "AppTextField" in it.metaObject().className()
                    or "AppDateField" in it.metaObject().className()]

        # суммы в строки: заполнение строк перестраивает делегаты,
        # поэтому значения ставим в сам массив строк (поля проверим после)
        full_rows = []
        for e in EMPS:
            full_rows.append({
                "id": e["id"], "name": e["name"], "subtitle": e["subtitle"],
                "checked": True,
                "hours": 80 if e["id"] == 1 else 0,
                "overtime": 40 if e["id"] == 1 else 0,
                "days": 5 if e["id"] == 3 else 0,
                "balHours": e["hours"], "balOvertime": e["overtime"],
                "balDays": e["days"],
                "prevHours": 0, "prevOvertime": 0, "prevDays": 0,
            })
        dlg.setProperty("rows", full_rows)
        settle(0.3)

        # номер и дата приказа — валидные, чтобы дошло до сохранения
        for f in find_fields():
            lbl = f.property("label") or ""
            if lbl == "Номер приказа":
                f.setProperty("text", "77")
            elif lbl == "Дата приказа":
                f.setProperty("selectedDate", "2026-10-13")
        settle(0.2)

        dlg.accepted.emit()
        settle(0.4)

        # поля строк СВЕЖИЕ (делегаты могли пересобраться): без подписи,
        # по три на сотрудника в порядке строк
        row_fields = [f for f in find_fields()
                      if (f.property("label") or "") == ""]
        assert len(row_fields) == 9, "ожидалось 3 строки × 3 поля, есть %d" % len(row_fields)
        rows3 = [row_fields[i:i + 3] for i in range(0, 9, 3)]
        # внутри строки: часы, дни, сверхурочные
        ivan = {"hours": rows3[0][0], "days": rows3[0][1], "overtime": rows3[0][2]}
        petr = {"hours": rows3[1][0], "days": rows3[1][1], "overtime": rows3[1][2]}
        sidr = {"hours": rows3[2][0], "days": rows3[2][1], "overtime": rows3[2][2]}

        # вспыхнули именно поля виновника, у невиновного — чисто
        assert ivan["overtime"].property("hasError") is True, "сверхурочные Иванова не вспыхнули"
        assert ivan["hours"].property("hasError") is True, "часы Иванова не вспыхнули"
        assert ivan["days"].property("hasError") is False, "дни Иванова вспыхнули зря"
        for k, f in petr.items():
            assert f.property("hasError") is False, "Петров подсвечен зря (%s)" % k
        assert sidr["days"].property("hasError") is True, "дни Сидорова не вспыхнули"
        assert sidr["hours"].property("hasError") is False

        # сообщения одного сотрудника собраны вместе
        row_errors = to_var(dlg.property("rowErrors"))
        msg1 = row_errors.get("1") or row_errors.get(1)
        assert msg1 and "не хватает часов" in msg1 and "не хватает сверхурочных" in msg1, \
            "сообщения Иванова не собраны: %r" % msg1

        # верхняя строка называет ВСЕХ виновников
        top = dlg.property("topError")
        assert "Иванов" in top and "Сидоров" in top, top
        print("B: поля виновников вспыхивают рамкой, все виновники названы ✓")

        # ── вспышка гаснет, повторная попытка вспыхивает снова ──
        settle(3.8)   # 3 секунды удержания + затухание
        assert ivan["overtime"].property("hasError") is False, "вспышка не погасла"
        dlg.accepted.emit()
        settle(0.4)
        # делегаты те же (строки не менялись) — читаем те же поля
        assert ivan["overtime"].property("hasError") is True, \
            "повторная попытка не вспыхнула (тот же текст ошибки)"
        assert FAIL_COUNT["n"] == 2, FAIL_COUNT
        print("   повторная попытка с тем же текстом вспыхивает снова ✓")
    finally:
        try:
            os.remove(wrapper)
        except OSError:
            pass


# ────────────────────────────────────────────────────────────────
# C. Журнал: карточка показывает сумму выбранного сотрудника
# ────────────────────────────────────────────────────────────────

WRAPPER_C = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 700; height: 600; visible: true
    AppUI.MoneyInspector { id: dlg }
    Component.onCompleted: dlg.open()
}
"""

ORDERS = [
    {
        "order_no": "245", "order_date": "2026-10-12", "date": "12.10.2026",
        "comment": "", "source_mode": 0, "count": 2,
        "hours": 120, "overtime": 0, "days": 0,
        "employees": [
            {"id": 1, "name": "Иванов Иван Иванович",
             "hours": 40, "overtime": 0, "days": 0},
            {"id": 2, "name": "Петров Пётр Петрович",
             "hours": 80, "overtime": 0, "days": 0},
        ],
    },
    {
        "order_no": "246", "order_date": "2026-10-13", "date": "13.10.2026",
        "comment": "", "source_mode": 0, "count": 1,
        "hours": 8, "overtime": 2, "days": 0,
        "employees": [
            {"id": 1, "name": "Иванов Иван Иванович",
             "hours": 8, "overtime": 2, "days": 0},
        ],
    },
]


class StubJournalBackend(QObject):
    moneyOrdersChanged = Signal()
    selectedEmployeeChanged = Signal()

    def __init__(self):
        super().__init__()
        self._orders = ORDERS
        self._sel = 1

    def getOrders(self):
        return self._orders

    moneyOrders = Property(list, getOrders, constant=True)

    def getSel(self):
        return self._sel

    def setSel(self, v):
        self._sel = int(v)
        self.selectedEmployeeChanged.emit()

    selectedEmployeeId = Property(int, getSel, setSel,
                                  notify=selectedEmployeeChanged)


def part_c_inspector_amounts(app) -> None:
    from PySide6.QtQml import QQmlApplicationEngine

    stub = StubJournalBackend()
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()          # ссылка обязательна (GC)
    ctx.setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_test_money_journal.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER_C)
    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "окно не загрузилось"
        win = engine.rootObjects()[0]

        def settle(s=0.5):
            end = time.time() + s
            while time.time() < end:
                app.processEvents()
                time.sleep(0.01)
            for _ in range(30):
                app.processEvents()

        settle()

        def card_texts():
            # карточки — делегаты Repeater: ищем в визуальном дереве
            # (обход в порядке следования, не стеком)
            res = []
            ci = win.contentItem() if hasattr(win, "contentItem") else None
            if ci is None:
                return res
            def rec(it):
                if it.objectName() == "orderCardAmount":
                    res.append(it.property("text"))
                for ch in it.childItems():
                    rec(ch)
            rec(ci)
            return res

        # выбран Иванов (получатель обоих приказов) — его суммы, с именем
        texts = card_texts()
        assert len(texts) == 2, texts
        assert texts[0] == "Иванов И. И.: 40 ч", texts[0]
        # одиночный приказ — без подписи (сумма и так его)
        assert texts[1] == "8 ч · 2 ч сверх.", texts[1]

        # выбран Петров — его сумма из общего приказа
        stub.setSel(2)
        settle(0.2)
        texts = card_texts()
        assert texts[0] == "Петров П. П.: 80 ч", texts[0]
        # приказ №246 Петрова не касается — общая сумма, как и раньше
        assert texts[1] == "8 ч · 2 ч сверх.", texts[1]

        # выбран никто (0) — общие суммы по приказу
        stub.setSel(0)
        settle(0.2)
        texts = card_texts()
        assert texts[0] == "120 ч", texts[0]
        print("C: журнал показывает сумму выбранного сотрудника, "
              "чужие приказы — общую ✓")
    finally:
        try:
            os.remove(wrapper)
        except OSError:
            pass


def main() -> int:
    from PySide6.QtGui import QGuiApplication
    part_a_backend_fields()
    app = QGuiApplication([])
    part_b_dialog_flash(app)
    part_c_inspector_amounts(app)
    print("═══ ДЕНЕЖНАЯ КОМПЕНСАЦИЯ: СВОИ СУММЫ И ВИНОВНИК ВИДЕН ✓ ═══")
    return 0


if __name__ == "__main__":
    rc = main()
    sys.stdout.flush()
    sys.stderr.flush()
    os._exit(rc)
