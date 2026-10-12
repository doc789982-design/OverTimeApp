#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сцены внутри окон: интеграция режиссёра с реальными окнами (батч 301).

Три сцены, каждая на своём окне:
  A. Производственный календарь — смена года в открытом окне: подмена
     отложена до конца выхода, новое содержание видно, картинка меняется.
  B. Настройки — смена раздела: индекс доезжает, видимость возвращается,
     картинка меняется.
  C. Календарь месяца — переключение «год ↔ месяц»: индекс и видимость
     возвращаются (вход месяца — чистое проявление, каскад несут ячейки).
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from PySide6.QtCore import QObject, Property, Signal, Slot, QMetaObject, Q_ARG
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuick import QQuickWindow


def wait_until(app, cond, timeout=4.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for _ in range(5):
            app.processEvents()
        if cond():
            return True
        time.sleep(0.02)
    return cond()


def find_child(win, className_part):
    return next(o for o in win.findChildren(QObject)
                if className_part in o.metaObject().className())


def quick_window(app, engine):
    """rootObjects может отдать QWindow-обёртку — берём настоящий QQuickWindow."""
    win = engine.rootObjects()[0]
    if not isinstance(win, QQuickWindow):
        wins = [w for w in app.allWindows() if isinstance(w, QQuickWindow)]
        assert wins, "нет QQuickWindow"
        win = wins[-1]
    return win


def grab_bytes(win):
    img = win.grabWindow()
    return (img.width(), img.height(), img.constBits().tobytes())


class ThemeStub(QObject):
    """Общий минимум: тема + производственный календарь (настоящие данные)."""
    isDarkChanged = Signal()

    def __init__(self):
        super().__init__()
        self._dark = False

    def getDark(self):
        return self._dark

    def setDark(self, v):
        self._dark = bool(v)

    isDarkTheme = Property(bool, getDark, setDark, notify=isDarkChanged)

    @Slot("QVariant", result="QVariant")
    def getProdCalendar(self, year):
        try:
            from methodical_data import PROD_CALENDARS
            return PROD_CALENDARS.get(int(year)) or {}
        except Exception:
            return {}


def scene_prod_calendar(app, engine) -> bool:
    wrapper = os.path.join(ROOT, "_test_scene_prodcal.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write("""import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 940; height: 700; visible: true
    AppUI.ProdCalendarDialog { id: dlg }
    function openY(y) { dlg.openYear(y) }
}
""")
    try:
        engine.load(wrapper)
        win = quick_window(app, engine)
        dlg = find_child(win, "ProdCalendarDialog")
        QMetaObject.invokeMethod(win, "openY", Q_ARG("QVariant", 2026))
        assert wait_until(app, lambda: dlg.property("opened")), "календарь не открылся"
        assert dlg.property("calYear") == 2026
        before = grab_bytes(win)

        # смена года в открытом окне — сцена
        QMetaObject.invokeMethod(win, "openY", Q_ARG("QVariant", 2027))
        app.processEvents()
        still_old = (dlg.property("calYear") == 2026)
        assert wait_until(app, lambda: dlg.property("calYear") == 2027), "год не сменился"
        director = find_child(win, "AppSceneDirector")
        column = director.property("target")
        assert column is not None, "режиссёр не привязан к колонке"
        assert wait_until(app, lambda: column.property("opacity") == 1.0), \
            "новый год не проявился до конца"
        after = grab_bytes(win)
        assert still_old, "подмена года применилась мгновенно, без сцены"
        assert before[2] != after[2], "картинка года не изменилась"
        print("A: год в открытом календаре меняется сценой, видимость возвращается ✓")
        return True
    finally:
        os.remove(wrapper)


def scene_settings(app, engine) -> bool:
    class SettingsStub(ThemeStub):
        @Property(str, constant=True)
        def activeDepartmentName(self):
            return ""

        @Property(list, constant=True)
        def dbList(self):
            return []

        @Property("QVariant", constant=True)
        def departmentData(self):
            return {}

        @Property(bool, constant=True)
        def reminderEnabled(self):
            return False

        @Property(bool, constant=True)
        def startupUpdateEnabled(self):
            return True

        @Property(str, constant=True)
        def timeInputMode(self):
            return "slider"

        @Property(bool, constant=True)
        def updateBusy(self):
            return False

        @Property(bool, constant=True)
        def updateReady(self):
            return False

        @Property(str, constant=True)
        def updateStatusText(self):
            return ""

        @Property(str, constant=True)
        def updateVersion(self):
            return ""

        @Property(list, constant=True)
        def hotkeysList(self):
            return []

        @Slot()
        def toggleTheme(self):
            pass

        @Slot(str)
        def setTimeInputMode(self, m):
            pass

        @Slot(bool)
        def setReminderEnabled(self, v):
            pass

        @Slot(bool)
        def setStartupUpdateEnabled(self, v):
            pass

    engine2 = QQmlApplicationEngine()
    stub_b = SettingsStub()   # ссылку держим: GC уносит стаб — и контекст пустеет
    engine2.rootContext().setContextProperty("backend", stub_b)
    wrapper = os.path.join(ROOT, "_test_scene_settings.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write("""import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 1100; height: 800; visible: true
    AppUI.SettingsDialog { id: dlg }
    function openT(i) { dlg.openTab(i) }
}
""")
    try:
        engine2.load(wrapper)
        win = quick_window(app, engine2)
        dlg = find_child(win, "SettingsDialog")
        dlg.open()
        assert wait_until(app, lambda: dlg.property("opened")), "настройки не открылись"
        before = grab_bytes(win)

        QMetaObject.invokeMethod(win, "openT", Q_ARG("QVariant", 2))
        app.processEvents()
        still_old = (dlg.property("currentIndex") in (None, 0))
        director = find_child(win, "AppSceneDirector")
        stack = director.property("target")
        assert wait_until(app, lambda: stack.property("currentIndex") == 2), \
            "раздел не сменился"
        assert wait_until(app, lambda: stack.property("opacity") == 1.0), \
            "раздел не проявился до конца"
        after = grab_bytes(win)
        assert before[2] != after[2], "картинка раздела не изменилась"
        print("B: раздел настроек меняется сценой, видимость возвращается ✓")
        print("   (мгновенность до сцены:", still_old, ")")
        return True
    finally:
        os.remove(wrapper)


def scene_calendar(app, engine) -> bool:
    import calendar as _cal

    def month_days(year, month):
        days = []
        for week in _cal.Calendar(firstweekday=0).monthdatescalendar(year, month):
            for d in week:
                days.append({
                    "date_str": d.isoformat(), "day_number": d.day,
                    "is_current_month": d.month == month and d.year == year,
                    "is_before_hire": False, "is_after_end": False,
                    "is_weekend": d.weekday() >= 5, "is_holiday": False,
                    "is_pre_holiday": False, "status": "", "has_comp": False,
                    "duties": [],
                })
        return days

    DAYS = month_days(2026, 10)

    class CalendarStub(ThemeStub):
        @Property(list, constant=True)
        def calendarDays(self):
            return DAYS

        @Property(int, constant=True)
        def selectedEmployeeId(self):
            return 1

        @Property(bool, constant=True)
        def isSelectedEmployeeShiftedWeekends(self):
            return False

        @Property(str, constant=True)
        def currentPeriodText(self):
            return "Октябрь 2026"

        @Property(list, constant=True)
        def dayDuties(self):
            return []

        @Property(list, constant=True)
        def dayComps(self):
            return []

        @Property(list, constant=True)
        def yearlyData(self):
            return []

        @Slot(str)
        def loadDayDetails(self, d):
            pass

        @Slot(int)
        def jumpToMonth(self, m):
            pass

        @Slot(int)
        def setYear(self, y):
            pass

    engine2 = QQmlApplicationEngine()
    stub_c = CalendarStub()   # ссылку держим: GC уносит стаб — и контекст пустеет
    engine2.rootContext().setContextProperty("backend", stub_c)
    wrapper = os.path.join(ROOT, "_test_scene_calendar.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write("""import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 1200; height: 800; visible: true
    AppUI.CalendarWorkspace { id: ws; anchors.fill: parent }
}
""")
    try:
        engine2.load(wrapper)
        win = quick_window(app, engine2)
        tabs = find_child(win, "AppPeriodTabs")
        director = find_child(win, "AppSceneDirector")
        stack = director.property("target")
        assert stack is not None, "режиссёр не привязан к стеку календаря"
        assert stack.property("currentIndex") == 0
        before = grab_bytes(win)

        # месяц → год
        tabs.yearViewClicked.emit()
        assert wait_until(app, lambda: stack.property("currentIndex") == 1), \
            "годовой вид не включился"
        assert wait_until(app, lambda: stack.property("opacity") == 1.0), \
            "годовой вид не проявился"
        year_img = grab_bytes(win)
        assert before[2] != year_img[2], "картинка годового вида не изменилась"

        # год → месяц (март: каскад ячеек после чистого проявления)
        tabs.monthClicked.emit(3)
        assert wait_until(app, lambda: stack.property("currentIndex") == 0), \
            "месячный вид не вернулся"
        assert wait_until(app, lambda: stack.property("opacity") == 1.0), \
            "месячный вид не проявился"
        print("C: календарь месяц ↔ год — сцены доезжают, видимость возвращается ✓")
        return True
    finally:
        os.remove(wrapper)


def main() -> int:
    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    stub = ThemeStub()        # ссылку держим: GC уносит стаб — и контекст пустеет
    engine.rootContext().setContextProperty("backend", stub)

    ok_a = scene_prod_calendar(app, engine)
    ok_b = scene_settings(app, engine)
    ok_c = scene_calendar(app, engine)
    assert ok_a and ok_b and ok_c
    print("═══ СЦЕНЫ В ОКНАХ: РЕЖИССЁР РАБОТАЕТ ═══")
    return 0


if __name__ == "__main__":
    rc = main()
    sys.stdout.flush()
    sys.stderr.flush()
    os._exit(rc)
