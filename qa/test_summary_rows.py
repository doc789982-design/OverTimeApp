#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Панель «Итоги» строками: было → отработано → компенсировано → осталось.

Проверяется:
  A. Backend отдаёт панели отдельные числа этого года и остаток
     прошлого года, а также подписи с настоящими датами периода.
  B. В настоящей панели видны заголовки колонок и подписанные
     значения (+16, −8), остаток прошлого года — словами, без скобок.
  C. Нажатие на строку параметра и на «Всего» открывает расшифровку.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_summary_rows.py
"""
import os
import sys
import tempfile
import time
import types
from datetime import date, datetime
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from PySide6.QtCore import QObject, Property  # noqa: E402

import Main as main_module  # noqa: E402
from database import DB  # noqa: E402
from logic import compute_month_summary  # noqa: E402


def make_backend(db):
    backend = main_module.Backend.__new__(main_module.Backend)
    QObject.__init__(backend)
    backend.active_db = db
    backend._selected_employee_id = 1
    backend.current_year = 2026
    backend.current_month = 10
    return backend


def test_backend_values():
    db = DB(os.path.join(tempfile.mkdtemp(prefix="summary_rows_"), "test.sqlite"))
    # add_employee: ночные минуты, дни, сверхурочные; затем прошлый год
    # в том же порядке, что принят в database.py
    emp_id = db.add_employee(
        "Иванов", "Иван", "", "майор", "инженер", "2026-01",
        600, 2, -240,
        1200, 600, 1)
    assert emp_id == 1
    db.add_duty(1, datetime(2026, 10, 7, 20, 0), datetime(2026, 10, 8, 8, 0), "")
    db.add_compensation_hours_dayoff(1, date(2026, 10, 12), 120, "", unit="hours")
    db.conn.commit()

    backend = make_backend(db)
    backend.refresh_calendar()
    raw = compute_month_summary(db, 1, 2026, 10)
    s = backend._month_summary
    assert s["period_start_label"] == "на 1 октября", s
    assert s["period_flow_label"] == "за октябрь", s
    assert s["period_end_label"] == "на 31 октября", s
    assert s["prev_year_label"] == "2025 год", s
    pairs = [
        ("hours_start_value", "start_hours"), ("hours_acc_value", "acc_hours"),
        ("hours_comp_value", "comp_h_real"), ("hours_end_value", "end_hours"),
        ("hours_prev_end_value", "prev_h_end"),
        ("days_start_value", "start_days"), ("days_acc_value", "acc_days"),
        ("days_comp_value", "comp_d_real"), ("days_end_value", "end_days"),
        ("days_prev_end_value", "prev_d_end"),
        ("overtime_start_value", "start_overtime"),
        ("overtime_acc_value", "acc_overtime"),
        ("overtime_comp_value", "comp_o_real"),
        ("overtime_end_value", "end_overtime"),
        ("overtime_prev_end_value", "prev_o_end"),
    ]
    for new_key, raw_key in pairs:
        assert s[new_key] == int(raw[raw_key] or 0), (new_key, s[new_key], raw[raw_key])

    backend._refresh_yearly_panorama_impl()
    ys = backend._year_summary
    dec = compute_month_summary(db, 1, 2026, 12)
    assert ys["period_start_label"] == "на 1 января", ys
    assert ys["period_flow_label"] == "за 2026 год", ys
    assert ys["period_end_label"] == "на 31 декабря", ys
    assert ys["hours_end_value"] == int(dec["end_hours"] or 0), ys
    assert ys["days_end_value"] == int(dec["end_days"] or 0), ys
    assert ys["overtime_end_value"] == int(dec["end_overtime"] or 0), ys
    print("A: backend отдаёт отдельные числа и настоящие даты периода ✓")


MONTH = {
    "is_shift": True,
    "period_start_label": "на 1 октября", "period_flow_label": "за октябрь",
    "period_end_label": "на 31 октября", "prev_year_label": "2025 год",
    "hours_start_value": 720, "hours_acc_value": 960, "hours_comp_value": 480,
    "hours_end_value": 1200, "hours_prev_end_value": 1200,
    "days_start_value": 2, "days_acc_value": 3, "days_comp_value": 1,
    "days_end_value": 4, "days_prev_end_value": 0,
    "overtime_start_value": -240, "overtime_acc_value": 1200,
    "overtime_comp_value": 0, "overtime_end_value": 960,
    "overtime_prev_end_value": 600,
    "norm_minutes": "168 ч.", "shift_night": "48 ч.", "shift_holiday": "16 ч.",
    "total_days": 12,
}
YEAR = dict(MONTH)


class StubBackend(QObject):
    isDarkTheme = Property(bool, lambda self: True, constant=True)
    reducedMotion = Property(bool, lambda self: False, constant=True)

    @Property(int, constant=True)
    def selectedEmployeeId(self):
        return 1

    @Property(dict, constant=True)
    def monthSummary(self):
        return MONTH

    @Property(dict, constant=True)
    def yearSummary(self):
        return YEAR

    @Property(str, constant=True)
    def currentPeriodText(self):
        return "Октябрь 2026"

    def loadMoneyOrders(self):
        pass


WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 980; height: 380; visible: true
    color: "#111111"
    Item {
        id: balanceExplainDialog
        objectName: "explainStub"
        property string lastBucket: ""
        property string lastScope: ""
        function openExplain(bucket, scope) { lastBucket = bucket; lastScope = scope }
    }
    Item { id: moneyInspector; function show() {} }
    AppUI.AppSummaryPanel {
        anchors.left: parent.left; anchors.right: parent.right; anchors.top: parent.top
        anchors.margins: 20
    }
}
"""


def test_qml_rows(app):
    from PySide6.QtCore import QPoint, QPointF, Qt
    from PySide6.QtQml import QQmlApplicationEngine
    from PySide6.QtQuick import QQuickWindow
    from PySide6.QtTest import QTest

    stub = StubBackend()
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()
    ctx.setContextProperty("backend", stub)
    wrapper = ROOT / "_test_summary_rows.qml"
    wrapper.write_text(WRAPPER, encoding="utf-8")

    def settle(seconds):
        end = time.time() + seconds
        while time.time() < end:
            app.processEvents()
            time.sleep(0.01)

    try:
        engine.load(str(wrapper))
        assert engine.rootObjects(), "панель не загрузилась"
        settle(2.0)
        qwin = max(
            [w for w in app.allWindows() if isinstance(w, QQuickWindow)],
            key=lambda w: w.width() * w.height())
        panel = next(
            o for o in qwin.findChildren(QObject)
            if o.metaObject().className().startswith("AppSummaryPanel"))
        texts = []

        def walk(item):
            if "Text" in item.metaObject().className():
                try:
                    value = item.property("text")
                except Exception:
                    value = None
                if value:
                    texts.append((str(value), item))
            for child in item.childItems():
                walk(child)

        # Делегаты Repeater надёжно находятся по видимому дереву,
        # а не по QObject-родству.
        walk(panel)

        by_text = {}
        for value, obj in texts:
            by_text.setdefault(value, []).append(obj)
        for needed in (
                "БЫЛО", "ОТРАБОТАНО", "КОМПЕНСИРОВАНО", "ОСТАЛОСЬ",
                "на 1 октября", "за октябрь", "на 31 октября",
                "12 ч.", "+16 ч.", "−8 ч.", "20 ч.",
                "2025 год: 20 ч.", "Всего", "12 дней"):
            assert needed in by_text, (needed, sorted(by_text))
        assert not any("(" in value and ")" in value for value in by_text), sorted(by_text)
        assert "?" not in by_text

        # Порядок в строке ночных часов: было → отработано → компенсировано → осталось
        xs = []
        for value in ("12 ч.", "+16 ч.", "−8 ч.", "20 ч."):
            item = by_text[value][0]
            xs.append(item.mapToItem(None, QPointF(0, 0)).x())
        assert xs == sorted(xs), xs
        print("B: колонки и подписанные цифры читаются одной строкой, скобок нет ✓")

        explain = next(
            o for o in qwin.findChildren(QObject)
            if o.objectName() == "explainStub")

        def click_item(item):
            pt = item.mapToItem(
                None, QPointF(item.width() / 2, item.height() / 2))
            QTest.mouseClick(
                qwin, Qt.LeftButton, Qt.NoModifier,
                QPoint(int(pt.x()), int(pt.y())))
            settle(0.25)

        click_item(by_text["20 ч."][0])
        assert explain.property("lastBucket") == "hours", explain.property("lastBucket")
        assert explain.property("lastScope") == "month", explain.property("lastScope")
        click_item(by_text["12 дней"][0])
        assert explain.property("lastBucket") == "total", explain.property("lastBucket")
        print("C: строка параметра и «Всего» открывают расшифровку ✓")
    finally:
        try:
            wrapper.unlink()
        except OSError:
            pass


def main():
    from PySide6.QtWidgets import QApplication

    app = QApplication.instance() or QApplication([])
    test_backend_values()
    test_qml_rows(app)
    print("PASS: панель «Итоги» читается строками")
    return 0


if __name__ == "__main__":
    rc = main()
    sys.stdout.flush()
    os._exit(rc)
