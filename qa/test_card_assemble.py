#!/usr/bin/env python3
"""Тест сборки новых карточек в календаре.

Добавленное дежурство, метка статуса и метка компенсации появляются
мягко: проявление и подъём (карточка дежурства) / проявление и рост
(круглые метки). Пара к растворению при удалении. Гасится при смене
месяца/сотрудника — там вход уже показывает шахматная анимация сетки.

Проверяется:
  A. Новое дежурство: карточка начинается прозрачной и собравшейся
     (подъём 8 px), через 300+ мс — на месте;
  B. Старые карточки того же дня при обновлении НЕ анимируются заново;
  C. Метка статуса и метка «В»: растут из полупрозрачного состояния;
  D. Смена месяца: карточки и метки появляются сразу, без второй
     анимации поверх шахматного входа сетки;
  E. Дежурство, исчезнувшее и вернувшееся (отмена действия),
     собирается снова — память ячейки о показанных id чистится.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 python3 qa/test_card_assemble.py
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


def month_days(year, month, duties=None, statuses=None):
    """Сетка месяца как строит refresh_calendar (недели с понедельника)."""
    duties = duties or {}
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
                "duties": duties.get(d.isoformat(), []),
            })
    return days


OCT = month_days(2026, 10)
NOV = month_days(2026, 11)
DUTY_101 = {"id": 101, "text": "Дн 08:00–20:00", "is_shift": True}
DUTY_102 = {"id": 102, "text": "Н 20:00–08:00", "is_shift": False}
DUTY_201 = {"id": 201, "text": "Дн 08:00–20:00", "is_shift": True}

DAYS = OCT          # текущие дни стаба
PERIOD = "Октябрь 2026"


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


def as_list(value):
    try:
        return list(value.toVariant())
    except AttributeError:
        return list(value)


def find_workspace(win):
    for obj in win.findChildren(QObject):
        if hasattr(obj, "findDayCell"):
            return obj
    return None


def iter_items(item):
    for c in item.childItems():
        yield c
        yield from iter_items(c)


def chip_body(cell, duty_id):
    """Тело карточки дежурства (внутренний Rectangle) по id.

    Ходим по дереву элементов: алиас dutyColItemRef указывает на
    QQuickColumn — тип приватный, PySide6 его не конвертирует.
    """
    for it in iter_items(cell):
        did = it.property("dutyId")
        if did is not None and int(did) == duty_id:
            return it.childItems()[0]
    return None


def badge_body(cell, alias):
    item = cell.property(alias)
    return item.childItems()[0] if item else None


def main() -> int:
    global DAYS, PERIOD

    qml_path = os.path.join(ROOT, "_render_assemble.qml")
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
        assert ws.property("needAnimation") is False, \
            "после обновления данных флаг анимации сетки должен сняться"

        # ── A+B. Добавили дежурство в день и статус в соседний ──
        DAYS = month_days(2026, 10,
                          duties={"2026-10-05": [DUTY_101]},
                          statuses={"2026-10-06": "Б"})
        DAYS = [dict(d, has_comp=(d["date_str"] == "2026-10-07")) for d in DAYS]
        stub.calendarDaysChanged.emit()
        app.processEvents()             # синхронная пересборка + старт анимаций

        cell5 = ws.findDayCell("2026-10-05")
        cell6 = ws.findDayCell("2026-10-06")
        cell7 = ws.findDayCell("2026-10-07")
        body101 = chip_body(cell5, 101)
        status_body = badge_body(cell6, "statusBarItemRef")
        comp_body = badge_body(cell7, "compBadgeItemRef")
        assert body101 is not None, "карточка дежурства не создалась"
        assert status_body is not None, "метка статуса не создалась"
        assert comp_body is not None, "метка компенсации не создалась"
        op_start = body101.property("opacity")
        sc_start = status_body.property("scale")
        sc_comp = comp_body.property("scale")
        assert op_start < 0.9, f"карточка не собирается: opacity={op_start}"
        assert sc_start < 0.9, f"метка не собирается: scale={sc_start}"
        assert sc_comp < 0.9, f"метка компенсации не собирается: scale={sc_comp}"
        settle(0.5)
        assert body101.property("opacity") > 0.999, "карточка не собралась до конца"
        assert status_body.property("scale") > 0.999, "метка не выросла до конца"
        assert comp_body.property("scale") > 0.999, "метка компенсации не выросла"
        assert body101.property("opacity") == 1.0
        print(f"A: дежурство собирается (старт opacity={op_start:.2f}), метки "
              f"растут (статус scale={sc_start:.2f}, «В» scale={sc_comp:.2f}) ✓")

        # ── F. Рендер: собранное видно на экране (не только свойства) ──
        from PySide6.QtCore import QPointF
        from PySide6.QtQuick import QQuickWindow
        qwin = next(x for x in app.topLevelWindows() if isinstance(x, QQuickWindow))
        img = qwin.grabWindow()
        assert img.width() > 0 and img.height() > 0, "пустой снимок окна"

        def center_color(item):
            pt = item.mapToScene(QPointF(item.width() / 2, item.height() / 2))
            return img.pixelColor(int(pt.x()), int(pt.y()))

        px_chip = center_color(body101)
        px_status = center_color(status_body)
        px_comp = center_color(comp_body)
        # чип сменный — синий, статус «Б» — красный, «В» — бирюзовый:
        # проверяем направление цвета, точный оттенок зависит от подложки
        assert px_chip.blue() > px_chip.red() + 8, \
            f"чип не синий на экране: {px_chip.name()}"
        assert px_status.red() > px_status.blue() + 8, \
            f"метка «Б» не красная на экране: {px_status.name()}"
        assert px_comp.blue() > px_comp.red() + 8, \
            f"метка «В» не бирюзовая на экране: {px_comp.name()}"
        print(f"F: рендер — чип {px_chip.name()}, «Б» {px_status.name()}, "
              f"«В» {px_comp.name()} ✓")

        # ── B. Вторая карточка в тот же день: старая не мигает ──
        DAYS = month_days(2026, 10,
                          duties={"2026-10-05": [DUTY_101, DUTY_102]},
                          statuses={"2026-10-06": "Б"})
        DAYS = [dict(d, has_comp=(d["date_str"] == "2026-10-07")) for d in DAYS]
        stub.calendarDaysChanged.emit()
        app.processEvents()
        body101b = chip_body(cell5, 101)
        body102 = chip_body(cell5, 102)
        assert body101b is not None and body102 is not None, "карточки не создались"
        assert body101b.property("opacity") > 0.999, \
            "старая карточка заново анимируется при обновлении данных"
        assert body102.property("opacity") < 0.9, "новая карточка не собирается"
        settle(0.5)
        assert body102.property("opacity") > 0.999
        print("B: при добавлении второй карточки старая стоит на месте ✓")

        # ── D. Смена месяца: без второй анимации поверх входа сетки ──
        settle(0.15)                    # память ячейки сверится (50 мс)
        DAYS = month_days(2026, 11,
                          duties={"2026-11-03": [DUTY_201]},
                          statuses={"2026-11-04": "О"})
        PERIOD = "Ноябрь 2026"
        stub.periodChanged.emit()
        stub.calendarDaysChanged.emit()
        app.processEvents()
        celln3 = ws.findDayCell("2026-11-03")
        celln4 = ws.findDayCell("2026-11-04")
        body201 = chip_body(celln3, 201)
        status_body_n = badge_body(celln4, "statusBarItemRef")
        assert body201 is not None and status_body_n is not None, "ноябрь не собрался"
        assert body201.property("opacity") > 0.999, \
            "при смене месяца карточка не должна играть вторую анимацию"
        assert status_body_n.property("scale") > 0.999, \
            "при смене месяца метка не должна играть вторую анимацию"
        print("D: смена месяца — карточки и метки появляются сразу ✓")

        # ── E. Исчезло и вернулось (отмена действия) — собирается снова ──
        settle(1.0)                     # вход ноября доиграет, память сверится
        DAYS = month_days(2026, 11, duties={"2026-11-03": []},
                          statuses={"2026-11-04": "О"})
        stub.calendarDaysChanged.emit()
        settle(0.15)                    # свера памяти: id 201 забыт
        DAYS = month_days(2026, 11,
                          duties={"2026-11-03": [DUTY_201]},
                          statuses={"2026-11-04": "О"})
        stub.calendarDaysChanged.emit()
        app.processEvents()
        body201b = chip_body(celln3, 201)
        assert body201b is not None, "вернувшаяся карточка не создалась"
        assert body201b.property("opacity") < 0.9, \
            "вернувшееся дежурство (отмена действия) не собирается"
        settle(0.5)
        assert body201b.property("opacity") > 0.999
        print("E: дежурство после отмены действия собирается снова ✓")

        print("═══ КАРТОЧКИ СОБИРАЮТСЯ ═══")
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
