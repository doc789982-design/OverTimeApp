#!/usr/bin/env python3
"""Тест волны меток по выделенным дням.

Статус применили сразу к нескольким дням (выделение протяжкой) —
метки собираются по очереди, в порядке выделения: читается как
мазок кисти по дням. Одиночный день собирается как раньше, смена
месяца гасит волну (там играет вход сетки), карта задержек живёт
недолго и стирается сама.

Проверяется:
  A. Карта задержек: меню готовит «день → задержка» ДО обновления
     данных, шаг 40 мс, порядок — хронологический;
  B. Волна: ранние метки собираются раньше поздних; в середине
     волны поздняя ещё не началась, ранняя уже на месте;
  C. Все метки волны в итоге собраны и видны на экране (пиксели);
  D. Одиночный день без волны — собирается сразу, как раньше;
  E. Смена месяца — волна не играет поверх входа сетки;
  F. Карта стирается сама после самой длинной волны.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 python3 qa/test_badge_wave.py
"""
import calendar as _cal
import os
import sys
import time
import types

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from PySide6.QtCore import QObject, Property, Signal  # noqa: E402
from PySide6.QtGui import QGuiApplication  # noqa: E402
from PySide6.QtQml import QQmlApplicationEngine  # noqa: E402
from PySide6.QtQuick import QQuickItem  # noqa: E402  (без него нет конвертеров Quick-типов)


def month_days(year, month, statuses=None):
    """Сетка месяца как строит refresh_calendar (недели с понедельника)."""
    statuses = statuses or {}
    days = []
    for week in _cal.Calendar(firstweekday=0).monthdatescalendar(year, month):
        for d in week:
            days.append({
                "date_str": d.isoformat(), "day_number": d.day,
                "is_current_month": d.month == month and d.year == year,
                "is_before_hire": False, "is_after_end": False,
                "is_weekend": d.weekday() >= 5, "is_holiday": False,
                "is_pre_holiday": False,
                "status": statuses.get(d.isoformat(), ""),
                "has_comp": False,
                "duties": [],
            })
    return days


DAYS = month_days(2026, 10)          # текущие дни стаба
PERIOD = "Октябрь 2026"

# Волна: 7 дней, 5–11 октября (как протяжка с понедельника по воскресенье)
WAVE_DATES = ["2026-10-%02d" % d for d in range(5, 12)]
LAST_DELAY = (len(WAVE_DATES) - 1) * 40      # 240 мс у последнего дня


class StubBackend(QObject):
    """Стаб на уровне модуля: календарь и период обновляемые."""

    calendarDaysChanged = Signal()
    periodChanged = Signal()

    @Property(list, notify=calendarDaysChanged)
    def calendarDays(self):
        return DAYS

    @Property(str, notify=periodChanged)
    def currentPeriodText(self):
        return PERIOD

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
    id: w
    width: 1200; height: 800; visible: true
    AppUI.CalendarWorkspace {
        id: workspace
        anchors.fill: parent
    }
}
"""


def find_workspace(win):
    for obj in win.findChildren(QObject):
        if hasattr(obj, "findDayCell"):
            return obj
    return None


def badge_body(cell):
    item = cell.property("statusBarItemRef")
    return item.childItems()[0] if item else None


def main() -> int:
    global DAYS, PERIOD

    qml_path = os.path.join(ROOT, "_render_wave.qml")
    with open(qml_path, "w", encoding="utf-8") as f:
        f.write(WRAPPER)

    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()          # ссылка обязательна (см. test_day_select)
    stub = StubBackend()
    ctx.setContextProperty("backend", stub)

    def settle(seconds):
        t = time.time() + seconds
        while time.time() < t:
            app.processEvents()
            time.sleep(0.01)

    try:
        engine.load(qml_path)
        assert engine.rootObjects(), "CalendarWorkspace не загрузился"
        settle(0.5)
        win = engine.rootObjects()[0]
        ws = find_workspace(win)
        assert ws is not None, "CalendarWorkspace не найден"

        # Прогрев, как в живом приложении: первый сигнал календаря —
        # это «смена» (сотрудник 0 → 1), играет вход месяца; второй —
        # уже обновление данных, флаг анимации сетки снимается
        stub.calendarDaysChanged.emit()
        settle(1.2)                     # вход месяца отыграет
        stub.calendarDaysChanged.emit()
        app.processEvents()
        assert ws.property("needAnimation") is False

        cells = {d: ws.findDayCell(d) for d in WAVE_DATES}
        assert all(c is not None for c in cells.values()), "ячейки не нашлись"

        # ── A. Карта задержек: меню готовит её до обновления данных ──
        ws.prepareBadgeWave(WAVE_DATES)
        delays = ws.property("badgeWaveDelays").toVariant()
        assert delays == {d: i * 40 for i, d in enumerate(WAVE_DATES)}, delays
        print("A: карта задержек — день → позиция * 40 мс, "
              "хронологический порядок ✓")

        # ── B. Волна: метки собираются по очереди ──
        DAYS = month_days(2026, 10,
                          statuses={d: "Б" for d in WAVE_DATES})
        stub.calendarDaysChanged.emit()
        app.processEvents()             # пересборка + отложенные решения

        bodies = {d: badge_body(cells[d]) for d in WAVE_DATES}
        assert all(b is not None for b in bodies.values()), "метки не создались"
        settle(0.12)                    # первый день успел, последний (240 мс) нет
        first_scale = bodies[WAVE_DATES[0]].property("scale")
        last_scale = bodies[WAVE_DATES[-1]].property("scale")
        assert first_scale > 0.9, \
            f"первая метка не собирается: scale={first_scale}"
        assert abs(last_scale - 0.55) < 0.01, \
            f"последняя метка не ждёт очереди: scale={last_scale}"
        print(f"B: волна — первая scale={first_scale:.2f}, "
              f"последняя ждёт (scale={last_scale:.2f}) ✓")

        # ── C. Все метки в итоге собраны и видны на экране ──
        settle(0.6)                     # 240 мс + 250 мс сборки
        for d in WAVE_DATES:
            sc = bodies[d].property("scale")
            op = bodies[d].property("opacity")
            assert sc > 0.999 and op > 0.999, \
                f"метка {d} не собралась: scale={sc}, opacity={op}"

        from PySide6.QtCore import QPointF
        from PySide6.QtQuick import QQuickWindow
        qwin = next(x for x in app.topLevelWindows()
                    if isinstance(x, QQuickWindow))
        img = qwin.grabWindow()
        assert img.width() > 0 and img.height() > 0, "пустой снимок окна"
        pt = bodies[WAVE_DATES[-1]].mapToScene(
            QPointF(bodies[WAVE_DATES[-1]].width() / 2,
                    bodies[WAVE_DATES[-1]].height() / 2))
        px = img.pixelColor(int(pt.x()), int(pt.y()))
        assert px.red() > px.blue() + 8, \
            f"последняя метка волны не видна на экране: {px.name()}"
        print(f"C: все метки собраны, последняя видна на экране "
              f"(пиксель {px.name()}) ✓")

        # ── D. Одиночный день без волны — как раньше, сразу ──
        settle(0.6)                     # карта волны стирается (проверим в F)
        DAYS = month_days(2026, 10,
                          statuses={**{d: "Б" for d in WAVE_DATES},
                                    "2026-10-12": "О"})
        stub.calendarDaysChanged.emit()
        app.processEvents()
        body12 = badge_body(ws.findDayCell("2026-10-12"))
        assert body12 is not None, "метка одиночного дня не создалась"
        assert body12.property("scale") < 0.9, \
            "одиночная метка не собирается сразу"
        settle(0.5)
        assert body12.property("scale") > 0.999
        print("D: одиночный день собирается сразу, без очереди ✓")

        # ── E. Смена месяца: волна не играет поверх входа сетки ──
        ws.prepareBadgeWave(["2026-11-%02d" % d for d in range(2, 9)])
        DAYS = month_days(2026, 11,
                          statuses={"2026-11-%02d" % d: "О"
                                    for d in range(2, 9)})
        PERIOD = "Ноябрь 2026"
        stub.periodChanged.emit()
        stub.calendarDaysChanged.emit()
        app.processEvents()
        body_nov = badge_body(ws.findDayCell("2026-11-02"))
        assert body_nov is not None, "ноябрьская метка не создалась"
        assert body_nov.property("scale") > 0.999 and \
            body_nov.property("opacity") > 0.999, \
            "при смене месяца метка не должна ждать очередь — вход сетки"
        print("E: смена месяца — метки на месте сразу, волна погашена ✓")

        # ── F. Карта стирается сама ──
        settle(1.6)
        leftovers = ws.property("badgeWaveDelays").toVariant()
        assert not leftovers, f"карта задержек не стёрлась: {leftovers}"
        print("F: карта задержек стирается сама после волны ✓")

        print("═══ ВОЛНА МЕТОК ПО ВЫДЕЛЕННЫМ ДНЯМ ═══")
    finally:
        try:
            os.remove(qml_path)
        except OSError:
            pass

    return 0


if __name__ == "__main__":
    rc = main()
    sys.stdout.flush()
    # PySide6 6.12 падает при финализации интерпретатора после QML-сцены
    # (none_dealloc на выходе) — выходим сразу после проверок
    os._exit(rc)
