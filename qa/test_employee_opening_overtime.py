#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Начальный минус разрешён только в «сверх нормы».

Проверяется:
  A. Backend при добавлении и редактировании сотрудника сохраняет
     отрицательные сверхурочные за этот и прошлый год, а ночные
     часы и дни ниже нуля не пропускает.
  B. Эти минусы честно попадают в итоги первого месяца учёта.
  C. Поле ввода с allowNegative принимает минус, обычное числовое — нет;
     в карточке сотрудника allowNegative стоит только на сверхурочных.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_employee_opening_overtime.py
"""
import os
import sys
import tempfile
import time
import types
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from PySide6.QtCore import QObject, Property  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

import Main as main_module  # noqa: E402
from database import DB  # noqa: E402
from logic import compute_month_summary  # noqa: E402


WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 420; height: 240; visible: true
    Column {
        anchors.fill: parent; anchors.margins: 16; spacing: 12
        AppUI.AppTextField { width: parent.width; label: "Обычное"; numericOnly: true }
        AppUI.AppTextField { width: parent.width; label: "С минусом"; numericOnly: true; allowNegative: true }
    }
}
"""


class StubBackend(QObject):
    isDarkTheme = Property(bool, lambda self: True, constant=True)
    reducedMotion = Property(bool, lambda self: False, constant=True)


def make_backend(db):
    backend = main_module.Backend.__new__(main_module.Backend)
    QObject.__init__(backend)
    backend.active_db = db
    backend.current_group_id = None
    backend._selected_employee_id = 0
    backend.refresh_employees = lambda: None
    backend.refresh_calendar = lambda: None
    return backend


def test_backend_and_summary():
    db = DB(os.path.join(tempfile.mkdtemp(prefix="opening_overtime_"), "test.sqlite"))
    backend = make_backend(db)
    backend.saveEmployee(
        "Иванов", "Иван", "", "майор", "инженер", "2026-01-01",
        -120, -180, -2,
        -60, -240, -3)
    emp = db.get_employee(1)
    assert emp is not None
    assert int(emp["opening_minutes"]) == 0, dict(emp)
    assert int(emp["opening_overtime_minutes"]) == -180, dict(emp)
    assert int(emp["opening_days"]) == 0, dict(emp)
    assert int(emp["prev_opening_minutes"]) == 0, dict(emp)
    assert int(emp["prev_opening_overtime_minutes"]) == -240, dict(emp)
    assert int(emp["prev_opening_days"]) == 0, dict(emp)

    summary = compute_month_summary(db, 1, 2026, 1)
    assert summary["start_overtime"] == -180, summary
    assert summary["prev_o_start"] == -240, summary
    assert summary["end_overtime"] == -180, summary
    assert summary["prev_o_end"] == -240, summary
    assert summary["start_hours"] == 0 and summary["start_days"] == 0, summary
    print("A–B: минус сверхурочных сохранён и виден в итогах, ночные/дни не ушли ниже нуля ✓")

    backend.updateEmployee(
        1, "Иванов", "Иван", "", "майор", "инженер", "2026-01-01",
        -999, -60, -9,
        -999, -120, -9)
    emp = db.get_employee(1)
    assert int(emp["opening_minutes"]) == 0 and int(emp["opening_days"]) == 0
    assert int(emp["prev_opening_minutes"]) == 0 and int(emp["prev_opening_days"]) == 0
    assert int(emp["opening_overtime_minutes"]) == -60
    assert int(emp["prev_opening_overtime_minutes"]) == -120
    print("A2: редактирование действует по тем же правилам ✓")


def test_qml_fields(app):
    from PySide6.QtQml import QQmlApplicationEngine
    from PySide6.QtQuick import QQuickWindow

    dialog_qml = (ROOT / "components/EmpDialog.qml").read_text(encoding="utf-8")
    assert dialog_qml.count("allowNegative: true") == 2, dialog_qml.count("allowNegative: true")
    assert "id: empOpenOvertime" in dialog_qml and "id: empPrevOpenOvertime" in dialog_qml

    stub = StubBackend()
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()
    ctx.setContextProperty("backend", stub)
    wrapper = ROOT / "_test_opening_overtime.qml"
    wrapper.write_text(WRAPPER, encoding="utf-8")
    try:
        engine.load(str(wrapper))
        assert engine.rootObjects(), "окно не загрузилось"
        wins = [w for w in app.allWindows() if isinstance(w, QQuickWindow)]
        assert wins, "QQuickWindow не найден"
        win = max(wins, key=lambda w: w.width() * w.height())
        end = time.time() + 0.5
        while time.time() < end:
            app.processEvents()
            time.sleep(0.01)

        fields = {}
        ci = win.contentItem()

        def walk(item):
            if "AppTextField" in item.metaObject().className():
                fields[item.property("label")] = item
            for child in item.childItems():
                walk(child)

        walk(ci)
        assert "Обычное" in fields and "С минусом" in fields, fields.keys()
        fields["Обычное"].setProperty("text", "-12")
        fields["С минусом"].setProperty("text", "-12")
        app.processEvents()
        assert fields["Обычное"].property("acceptableInput") is False
        assert fields["С минусом"].property("acceptableInput") is True
        print("C: минус принимает только специально разрешённое поле ✓")
    finally:
        try:
            wrapper.unlink()
        except OSError:
            pass


def main():
    app = QApplication.instance() or QApplication([])
    test_backend_and_summary()
    test_qml_fields(app)
    print("PASS: начальный минус возможен только в сверхурочных")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
