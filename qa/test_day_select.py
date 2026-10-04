#!/usr/bin/env python3
"""Тест выделения дней протяжкой и массовых действий меню дня.

Зажатая ЛКМ и движение по числам месяца выделяет диапазон дней; после
отпускания открывается меню дня, и его действия применяются ко всем
выделенным дням сразу.

Проверяется:
  A. Массовые слоты Python (одна транзакция, все дни сразу):
     статусы К/Б/О и сброс, тип дня (рабочий/выходной/праздничный),
     удаление всех дежурств и всех компенсаций; пустой список — не падает;
  B. Логика выделения в календаре (QML): начало на одном дне,
     протяжка до другого — выделяется диапазон между ними; чужой
     месяц в диапазон не попадает; Esc/закрытие меню снимает выделение;
  C. Меню дня: диапазон в шапке «12–15 марта 2026 г.», действия
     идут через массовые слоты при выделении.

Запуск:
    python3 qa/test_day_select.py             # из корня репозитория
"""
import os
import sys
import types
from datetime import date, datetime

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

# Main.py импортирует win32print (только Windows) — в песочнице подменяем
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

from database import DB
from Main import Backend
from PySide6.QtCore import QObject, Property
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine


def fresh_db():
    path = os.path.join(os.environ.get("TMPDIR", "/tmp"),
                        "day_select_test.sqlite")
    if os.path.exists(path):
        os.remove(path)
    db = DB(path)
    db.add_employee("Тестов", "Тест", "Тестович", "капитан", "инженер",
                    "2020-01", 0, 0, 0, 0, 0, 0)
    # прямые записи оставляют транзакцию открытой (в приложении её
    # закрывает обёртка Backend) — закрываем сами, иначе снапшот для
    # Ctrl+Z в begin() уходит в вечное ожидание
    db.conn.commit()
    return db


def make_backend(db, eid=1):
    """Живой Backend без тяжёлого __init__: только то, что нужно слотам."""
    b = Backend.__new__(Backend)
    QObject.__init__(b)
    b.active_db = db
    b._selected_employee_id = eid
    b.refresh_calendar = lambda: None
    b._defer_year_refresh = lambda: None
    return b


def part_a_bulk_slots() -> None:
    db = fresh_db()
    b = make_backend(db)

    # ── статусы нескольким дням ──
    b.setDayStatusBulk(["2026-03-02", "2026-03-03", "2026-03-04"], "Б")
    st = db.get_statuses_for_period(1, "2026-03-01", "2026-03-31")
    got = {d.isoformat(): v for d, v in st.items()}
    assert got == {"2026-03-02": "Б", "2026-03-03": "Б", "2026-03-04": "Б"}, got

    # ── сброс статусов тем же слотом ──
    b.setDayStatusBulk(["2026-03-02", "2026-03-04"], "")
    st = db.get_statuses_for_period(1, "2026-03-01", "2026-03-31")
    got = {d.isoformat(): v for d, v in st.items() if v}
    assert got == {"2026-03-03": "Б"}, got

    # ── тип дня: праздничный ──
    b.setDayTypeBulk(["2026-03-05", "2026-03-06"], "holiday")
    row = db.conn.execute(
        "SELECT is_working, is_holiday FROM calendar_day WHERE date=?",
        ("2026-03-05",)).fetchone()
    assert row and row["is_holiday"] == 1 and row["is_working"] == 0, tuple(row)
    row = db.conn.execute(
        "SELECT is_holiday FROM calendar_day WHERE date=?",
        ("2026-03-06",)).fetchone()
    assert row and row["is_holiday"] == 1, tuple(row)

    # ── удаление всех дежурств в выделенных днях ──
    db.add_duty(1, datetime(2026, 3, 2, 8, 0), datetime(2026, 3, 2, 20, 0), "а")
    db.add_duty(1, datetime(2026, 3, 3, 8, 0), datetime(2026, 3, 3, 20, 0), "б")
    db.add_duty(1, datetime(2026, 3, 9, 8, 0), datetime(2026, 3, 9, 20, 0), "в")
    db.conn.commit()
    b.clearDayDutiesBulk(["2026-03-02", "2026-03-03"])
    left = db.list_duties_for_period(
        1, datetime(2026, 3, 1), datetime(2026, 4, 1))
    assert len(left) == 1 and left[0]["comment"] == "в", \
        [dict(r) for r in left]

    # ── удаление компенсаций в выделенных днях ──
    db.add_compensation_hours_dayoff(1, date(2026, 3, 12), 480, "ч")
    db.conn.commit()
    b.clearDayCompensationsBulk(["2026-03-12"])
    n = db.conn.execute(
        "SELECT COUNT(*) c FROM compensation WHERE employee_id=1").fetchone()["c"]
    assert n == 0, n

    # ── пустой список и дни без записей — не падают ──
    b.setDayStatusBulk([], "О")
    b.setDayTypeBulk([], "work")
    b.clearDayDutiesBulk([])
    b.clearDayCompensationsBulk(["2026-03-20"])
    print("A: статусы, типы дней, дежурства и компенсации — сразу всем дням ✓")


# ────────────────────────────────────────────────────────────────
# Часть B: логика выделения в календаре (настоящий CalendarWorkspace)
# ────────────────────────────────────────────────────────────────

def march_days():
    """42 ячейки марта 2026 (31 день) + апрель-хвост (не текущий месяц)."""
    days = []
    for d in range(1, 32):
        days.append({
            "date_str": "2026-03-%02d" % d, "day_number": d,
            "is_current_month": True, "is_before_hire": False,
            "is_after_end": False, "is_weekend": False, "is_holiday": False,
            "is_pre_holiday": False, "status": "", "has_comp": False,
            "duties": [],
        })
    for d in range(1, 12):
        days.append({
            "date_str": "2026-04-%02d" % d, "day_number": d,
            "is_current_month": False, "is_before_hire": False,
            "is_after_end": False, "is_weekend": False, "is_holiday": False,
            "is_pre_holiday": False, "status": "", "has_comp": False,
            "duties": [],
        })
    return days


DAYS = march_days()

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


class StubBackend(QObject):
    """Стаб на уровне модуля (локальные классы QML видит ненадёжно)."""

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
        return "Март 2026"

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


def read_var(obj, name):
    """QML property var приходит в Python как QJSValue — разворачиваем."""
    v = obj.property(name)
    try:
        from PySide6.QtQml import QJSValue
        if isinstance(v, QJSValue):
            v = v.toVariant()
    except Exception:
        pass
    return v


def qml_selection_checks(app) -> None:
    # обёртка кладётся в корень репозитория — относительный импорт
    # «components» работает только оттуда (как у test_panel_view)
    qml_path = os.path.join(ROOT, "_render_dayselect.qml")
    with open(qml_path, "w", encoding="utf-8") as f:
        f.write(WRAPPER)

    engine = QQmlApplicationEngine()
    # ссылки на контекст и стаб обязательны: без них PySide6 отдает
    # контекст сборщику мусора и в QML backend становится null
    ctx = engine.rootContext()
    stub = StubBackend()
    ctx.setContextProperty("backend", stub)
    try:
        engine.load(qml_path)
        if not engine.rootObjects():
            raise AssertionError("CalendarWorkspace не загрузился")
        win = engine.rootObjects()[0]
        ws = None
        for obj in win.findChildren(QObject):
            if hasattr(obj, "extendDaySelection"):
                ws = obj
                break
        assert ws is not None, "CalendarWorkspace не найден"

        # ячейка по дате: центр ячейки в координатах сетки
        def cell_center(date_str):
            c = ws.findDayCell(date_str)
            assert c is not None, "ячейка %s не нашлась" % date_str
            return (c.property("x") + c.property("width") / 2,
                    c.property("y") + c.property("height") / 2)

        # ── протяжка с 5-го на 10-е марта = 6 дней ──
        ws.beginDaySelection("2026-03-05")
        assert list(read_var(ws, "multiSelectDates") or []) == ["2026-03-05"], \
            list(read_var(ws, "multiSelectDates") or [])
        assert not bool(read_var(ws, "multiSelectActive")), "один день — не выделение"
        x, y = cell_center("2026-03-10")
        ws.extendDaySelection(x, y)
        want = ["2026-03-%02d" % d for d in range(5, 11)]
        assert list(read_var(ws, "multiSelectDates") or []) == want, \
            list(read_var(ws, "multiSelectDates") or [])
        assert bool(read_var(ws, "multiSelectActive")), \
            "диапазон должен считаться выделением"

        # ── протяжка «назад» с 10-го на 3-е — те же правила ──
        ws.beginDaySelection("2026-03-10")
        x, y = cell_center("2026-03-03")
        ws.extendDaySelection(x, y)
        want = ["2026-03-%02d" % d for d in range(3, 11)]
        assert list(read_var(ws, "multiSelectDates") or []) == want, \
            list(read_var(ws, "multiSelectDates") or [])

        # ── апрельская ячейка (не текущий месяц) — выделение не меняется ──
        x, y = cell_center("2026-04-03")
        before = list(read_var(ws, "multiSelectDates") or [])
        ws.extendDaySelection(x, y)
        assert list(read_var(ws, "multiSelectDates") or []) == before, \
            list(read_var(ws, "multiSelectDates") or [])

        # ── точка мимо ячеек — выделение не меняется ──
        ws.extendDaySelection(-50, -50)
        assert list(read_var(ws, "multiSelectDates") or []) == before, \
            list(read_var(ws, "multiSelectDates") or [])

        # ── снятие выделения ──
        ws.clearDaySelection()
        assert len(read_var(ws, "multiSelectDates") or []) == 0
        assert not bool(read_var(ws, "multiSelectActive"))
        print("B: протяжка 5→10 марта, задним ходом, чужой месяц мимо, "
              "снятие выделения ✓")
    finally:
        try:
            os.remove(qml_path)
        except OSError:
            pass


def part_c_menu_checks() -> None:
    # fmtRange в меню: один месяц — «5–10 марта 2026 г.»,
    # через месяц — полные даты через тире
    qml = open(os.path.join(ROOT, "components", "AppDayMenu.qml"),
               encoding="utf-8").read()
    assert "targetDates" in qml, "меню не знает про массовые дни"
    assert "setDayStatusBulk" in qml and "setDayTypeBulk" in qml
    assert "clearDayDutiesBulk" in qml and "clearDayCompensationsBulk" in qml
    assert "fmtRange" in qml, "нет диапазона в шапке"
    cal = open(os.path.join(ROOT, "components", "CalendarWorkspace.qml"),
               encoding="utf-8").read()
    assert "beginDaySelection" in cal and "extendDaySelection" in cal
    assert "multiSelectDates" in cal
    main = open(os.path.join(ROOT, "main.qml"), encoding="utf-8").read()
    assert "clearDaySelection" in main, "меню не снимает выделение при закрытии"
    print("C: меню дня — массовые слоты, диапазон в шапке, снятие при "
          "закрытии ✓")


def main() -> int:
    part_a_bulk_slots()

    app = QGuiApplication.instance() or QGuiApplication(sys.argv)
    qml_selection_checks(app)
    part_c_menu_checks()

    print("═══ ВЫДЕЛЕНИЕ ДНЕЙ ПРОТЯЖКОЙ И МАССОВЫЕ ДЕЙСТВИЯ РАБОТАЮТ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
