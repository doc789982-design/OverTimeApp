#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Метки календаря при наведении разворачиваются в пилюльки.

Проверяется:
  A. Количество времени пишется полными словами: 1 час, 3 часа,
     8 часов, минуты не теряются.
  B. Backend кладёт в каждый день подпись компенсации: ночные часы,
     сверхурочные и дни — по дате, где стоит метка «В».
  C. В настоящем календаре при наведении круг «В» и круги Б/О/К
     растут только влево: правый край не двигается, внутри —
     «8 часов», «больничный», «отпуск», «командировка»; курсор
     ушёл — пилюлька снова круг.
  D. В панели «Итоги» вопросительных знаков больше нет, а нажатие
     на графы (расшифровка цифры) осталось.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_calendar_badge_pills.py
"""
import calendar as _cal
import os
import sys
import tempfile
import time
import types
from datetime import date
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from PySide6.QtCore import QObject, Property, Signal  # noqa: E402

import Main as main_module  # noqa: E402
from database import DB  # noqa: E402
from utils import fmt_minutes_ru_full  # noqa: E402


def test_words():
    assert fmt_minutes_ru_full(60) == "1 час"
    assert fmt_minutes_ru_full(180) == "3 часа"
    assert fmt_minutes_ru_full(480) == "8 часов"
    assert fmt_minutes_ru_full(90) == "1 час 30 минут"
    assert fmt_minutes_ru_full(30) == "30 минут"
    assert fmt_minutes_ru_full(61) == "1 час 1 минута"
    print("A: время полными словами — 1 час, 3 часа, 8 часов ✓")


def make_backend(db):
    backend = main_module.Backend.__new__(main_module.Backend)
    QObject.__init__(backend)
    backend.active_db = db
    backend._selected_employee_id = 1
    backend.current_year = 2026
    backend.current_month = 10
    return backend


def test_backend_labels():
    db = DB(os.path.join(tempfile.mkdtemp(prefix="badge_pills_"), "test.sqlite"))
    emp_id = db.add_employee(
        "Иванов", "Иван", "", "майор", "инженер", "2026-01",
        0, 0, 0, 0, 0, 0)
    assert emp_id == 1
    db.add_compensation_hours_dayoff(1, date(2026, 10, 5), 480, "", unit="hours")
    db.add_compensation_hours_dayoff(1, date(2026, 10, 6), 180, "", unit="overtime")
    db.add_compensation_hours_dayoff(1, date(2026, 10, 7), 60, "", unit="hours")
    db.add_compensation_days_dayoff(1, [date(2026, 10, 8), date(2026, 10, 9)], "")
    db.conn.commit()

    backend = make_backend(db)
    backend.refresh_calendar()
    days = {d["date_str"]: d for d in backend._calendar_days}
    expected = {
        "2026-10-05": "8 часов",
        "2026-10-06": "3 часа",
        "2026-10-07": "1 час",
        "2026-10-08": "1 день",
        "2026-10-09": "1 день",
    }
    for date_str, label in expected.items():
        row = days[date_str]
        assert row["has_comp"] is True, row
        assert row["comp_label"] == label, row
    assert days["2026-10-12"]["comp_label"] == ""
    print("B: подпись компенсации лежит в дне календаря ✓")


# ── настоящий календарь на стабе ──────────────────────────────────

COMP_DATE = "2026-10-05"
STATUS_DATES = {
    "2026-10-12": ("Б", "больничный"),
    "2026-10-13": ("О", "отпуск"),
    "2026-10-14": ("К", "командировка"),
}


def month_days():
    days = []
    for week in _cal.Calendar(firstweekday=0).monthdatescalendar(2026, 10):
        for d in week:
            iso = d.isoformat()
            days.append({
                "date_str": iso, "day_number": d.day,
                "is_current_month": d.month == 10,
                "is_before_hire": False, "is_after_end": False,
                "is_weekend": d.weekday() >= 5, "is_holiday": False,
                "is_pre_holiday": False,
                "status": STATUS_DATES.get(iso, ("", ""))[0],
                "has_comp": iso == COMP_DATE,
                "comp_label": "8 часов" if iso == COMP_DATE else "",
                "duties": [],
            })
    return days


DAYS = month_days()


class StubBackend(QObject):
    calendarDaysChanged = Signal()
    periodChanged = Signal()
    isDarkTheme = Property(bool, lambda self: True, constant=True)
    reducedMotion = Property(bool, lambda self: False, constant=True)

    @Property(list, notify=calendarDaysChanged)
    def calendarDays(self):
        return DAYS

    @Property(str, notify=periodChanged)
    def currentPeriodText(self):
        return "Октябрь 2026"

    @Property(int, constant=True)
    def selectedEmployeeId(self):
        return 1

    @Property(bool, constant=True)
    def isSelectedEmployeeShiftedWeekends(self):
        return False

    @Property(list, constant=True)
    def dayDuties(self):
        return []

    @Property(list, constant=True)
    def dayComps(self):
        return []

    @Property(list, constant=True)
    def yearlyData(self):
        return []

    def loadDayDetails(self, date_str):
        pass

    def executeHotkey(self, seq, date_str):
        pass

    def handleClipboard(self, action, date_str):
        pass

    def jumpToMonth(self, m):
        pass

    def setYear(self, y):
        pass


WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 1400; height: 900; visible: true
    AppUI.CalendarWorkspace { anchors.fill: parent }
}
"""


def badge_body(cell, ref_name):
    item = cell.property(ref_name)
    assert item is not None, ref_name
    return item, item.childItems()[0]


def test_qml_pills(app):
    from PySide6.QtCore import QCoreApplication, QEvent, QPointF
    from PySide6.QtGui import QHoverEvent
    from PySide6.QtQml import QQmlApplicationEngine
    from PySide6.QtQuick import QQuickWindow

    stub = StubBackend()
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()
    ctx.setContextProperty("backend", stub)
    wrapper = ROOT / "_test_badge_pills.qml"
    wrapper.write_text(WRAPPER, encoding="utf-8")

    def settle(seconds):
        end = time.time() + seconds
        while time.time() < end:
            app.processEvents()
            time.sleep(0.01)

    try:
        engine.load(str(wrapper))
        assert engine.rootObjects(), "окно не загрузилось"
        settle(0.5)
        stub.calendarDaysChanged.emit()
        settle(1.2)
        stub.calendarDaysChanged.emit()
        settle(0.4)

        wins = [w for w in app.allWindows() if isinstance(w, QQuickWindow)]
        assert wins, "QQuickWindow не найден"
        qwin = max(wins, key=lambda w: w.width() * w.height())
        ws = None
        for obj in qwin.findChildren(QObject):
            if hasattr(obj, "findDayCell"):
                ws = obj
                break
        assert ws is not None, "CalendarWorkspace не найден"

        def hover_at(pt, old=None):
            ev = QHoverEvent(QEvent.HoverMove, pt, old or QPointF(pt.x() - 1, pt.y()))
            QCoreApplication.sendEvent(qwin, ev)
            settle(0.35)

        def check_pill(date_str, ref_name, letter, label):
            cell = ws.findDayCell(date_str)
            assert cell is not None, date_str
            item, body = badge_body(cell, ref_name)
            assert item.property("fullLabel") == label, item.property("fullLabel")
            assert body.width() <= 19, body.width()
            right_before = body.mapToItem(None, QPointF(body.width(), 0)).x()
            texts = {}
            for child in body.childItems():
                try:
                    if child.property("text"):
                        texts[child.property("text")] = child
                except Exception:
                    pass
            assert letter in texts and label in texts, texts.keys()

            center = body.mapToItem(None, QPointF(body.width() / 2, body.height() / 2))
            hover_at(center)
            assert item.property("expanded") is True, date_str
            assert body.width() > 30, body.width()
            right_after = body.mapToItem(None, QPointF(body.width(), 0)).x()
            assert abs(right_after - right_before) <= 1.5, (right_before, right_after)
            assert texts[label].property("opacity") > 0.9
            assert texts[letter].property("opacity") < 0.1

            hover_at(QPointF(2, 2), center)
            assert item.property("expanded") is False, date_str
            assert body.width() <= 19, body.width()

        check_pill(COMP_DATE, "compBadgeItemRef", "В", "8 часов")
        for date_str, (letter, label) in STATUS_DATES.items():
            check_pill(date_str, "statusBarItemRef", letter, label)
        print("C: «В» показывает «8 часов», Б/О/К — полные слова, рост только влево ✓")
    finally:
        try:
            wrapper.unlink()
        except OSError:
            pass


def test_summary_no_questions():
    panel = (ROOT / "components/AppSummaryPanel.qml").read_text(encoding="utf-8")
    assert 'text: "?"' not in panel
    assert "onTapped: root.openExplain(card.bucket)" in panel
    assert 'onTapped: root.openExplain("total")' in panel
    print("D: вопросиков нет, расшифровка по нажатию осталась ✓")


def main():
    from PySide6.QtWidgets import QApplication

    app = QApplication.instance() or QApplication([])
    test_words()
    test_backend_labels()
    test_qml_pills(app)
    test_summary_no_questions()
    print("PASS: метки календаря разворачиваются в пилюльки")
    return 0


if __name__ == "__main__":
    rc = main()
    sys.stdout.flush()
    # PySide6 6.12 падает при финализации интерпретатора после QML-сцены
    # (none_dealloc на выходе) — выходим сразу после проверок
    os._exit(rc)
