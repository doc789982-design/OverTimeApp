#!/usr/bin/env python3
"""Тест общей картины месяца (полоса дней у каждого сотрудника).

Пока сотрудник не выбран, карточка каждого человека в списке
продолжается строкой мини-ячеек дней текущего месяца — те же цвета и
метки, что у календаря. Включается тумблером в «Внешнем виде».

Проверяется:
  A. Построитель данных (Main.Backend.refresh_team_grid): по сотруднику
     на строку — только дни текущего месяца; выходные, праздники,
     статусы, дежурства (текст как в календаре), компенсации, запертые
     дни до приёма; при выбранном сотруднике не строится;
  B. Живой рендер: EmployeeListPanel в режиме команды рисует по 31
     мини-ячейке на сотрудника (TeamDayCell);
  C. Тумблер: по умолчанию включён, слот пишет в настройки; маскот
     и ряд «Пн..Вс» скрываются, список расширяется на окно.

Запуск:
    python3 qa/test_team_strip.py             # из корня репозитория
"""
import os
import sys
import types
from datetime import date, datetime
from pathlib import Path

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

OCT = 10


def fresh_db():
    path = os.path.join(os.environ.get("TMPDIR", "/tmp"),
                        "team_strip_test.sqlite")
    if os.path.exists(path):
        os.remove(path)
    db = DB(path)
    db.add_employee("Иванов", "Иван", "", "капитан", "инженер",
                    "2020-01", 0, 0, 0, 0, 0, 0)
    db.add_employee("Петров", "Пётр", "", "лейтенант", "инженер",
                    "2026-10", 0, 0, 0, 0, 0, 0)
    # прямые записи оставляют транзакцию открытой (в приложении её
    # закрывает обёртка Backend) — закрываем сами
    db.conn.commit()
    return db


def make_backend(db):
    """Живой Backend без тяжёлого __init__: только то, что нужно полосе."""
    b = Backend.__new__(Backend)
    QObject.__init__(b)
    b.active_db = db
    b._selected_employee_id = 0
    b.current_year = 2026
    b.current_month = OCT
    b._team_grid = []
    b._team_strip_enabled = True
    b.config_path = Path(os.environ.get("TMPDIR", "/tmp")) / "team_strip_cfg.json"
    if b.config_path.exists():
        os.remove(b.config_path)
    b._employee_list = [
        {"id": -1, "is_header": True, "name": "Все"},
        {"id": 1, "is_header": False, "hire_date": "2020-01-01",
         "end_date": ""},
        {"id": 2, "is_header": False, "hire_date": "2026-10-10",
         "end_date": ""},
    ]
    return b


def part_a_builder() -> None:
    db = fresh_db()
    db.set_day_status(1, date(2026, OCT, 5), "Б")
    db.add_duty(1, datetime(2026, OCT, 7, 8, 0), datetime(2026, OCT, 7, 20, 0), "")
    db.add_compensation_hours_dayoff(1, date(2026, OCT, 12), 480, "")
    db.conn.commit()

    b = make_backend(db)
    b.refresh_team_grid()

    grid = b._team_grid
    assert [g["id"] for g in grid] == [1, 2], "заголовки групп не должны попасть"
    for g in grid:
        assert len(g["days"]) == 31, "в октябре 31 день"

    days1 = {d["date_str"]: d for d in grid[0]["days"]}
    # 3 октября 2026 — суббота
    assert days1["2026-10-03"]["is_weekend"] is True
    assert days1["2026-10-05"]["is_weekend"] is False
    # статус, дежурство, компенсация
    assert days1["2026-10-05"]["status"] == "Б"
    duty = days1["2026-10-07"]["duties"]
    assert len(duty) == 1 and duty[0]["text"] == "08:00-20:00", duty
    assert days1["2026-10-12"]["has_comp"] is True
    assert days1["2026-10-01"]["has_comp"] is False

    # у принятого 10-го числа дни до приёма заперты
    days2 = {d["date_str"]: d for d in grid[1]["days"]}
    assert days2["2026-10-05"]["is_before_hire"] is True
    assert days2["2026-10-10"]["is_before_hire"] is False
    assert days2["2026-10-20"]["is_before_hire"] is False

    # при выбранном сотруднике полоса не строится
    b._selected_employee_id = 1
    b._team_grid = []
    b.refresh_team_grid()
    assert b._team_grid == []

    # тумблер: по умолчанию включён, слот пишет в настройки
    assert b.teamStripEnabled is True
    b.setTeamStripEnabled(False)
    assert b.teamStripEnabled is False
    assert '"team_month_strip": false' in b.config_path.read_text(
        encoding="utf-8"), "настройка не записана"
    b.setTeamStripEnabled(True)
    assert b.teamStripEnabled is True
    print("A: полоса строится по всем сотрудникам, тумблер пишет настройки ✓")


# ────────────────────────────────────────────────────────────────
# Часть B: живой рендер полосы
# ────────────────────────────────────────────────────────────────

def stub_days():
    days = []
    for d in range(1, 32):
        days.append({
            "date_str": "2026-10-%02d" % d, "day_number": d,
            "is_weekend": d % 7 in (3, 4), "is_holiday": False,
            "status": "Б" if d == 5 else "",
            "has_comp": d == 12, "duties": [{"id": 1, "text": "08:00-20:00",
                                             "is_shift": True}] if d == 7 else [],
            "is_before_hire": False, "is_after_end": False,
        })
    return days


def emp(idx):
    return {"id": idx, "is_header": False, "name": "Иванов Иван",
            "subtitle": "капитан — инженер", "is_active": True,
            "inactive_reason": "", "has_overtime": False,
            "shift_minutes": 0, "norm_minutes": 0,
            "last_name": "Иванов", "first_name": "Иван", "middle_name": "",
            "rank": "капитан", "position": "инженер", "start_month": "2020-01",
            "hire_date": "2020-01-01", "end_date": "",
            "opening_minutes": 0, "opening_overtime": 0, "opening_days": 0,
            "prev_opening_minutes": 0, "prev_opening_overtime": 0,
            "prev_opening_days": 0, "group_id": 1}


EMPLOYEES = [{"id": 0, "is_header": True, "name": "Первая группа"}] + \
            [emp(1), emp(2)]
TEAM_GRID = [{"id": 1, "days": stub_days()}, {"id": 2, "days": stub_days()}]

WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI

ApplicationWindow {
    id: w
    width: 1200; height: 800; visible: true
    AppUI.EmployeeListPanel {
        id: panel
        anchors.fill: parent
    }
}
"""


class StubBackend(QObject):
    """Стаб на уровне модуля (локальные классы QML видит ненадёжно)."""

    @Property(list, constant=True)
    def employeeList(self):
        return EMPLOYEES

    @Property(int, constant=True)
    def selectedEmployeeId(self):
        return 0

    @Property(bool, constant=True)
    def teamStripEnabled(self):
        return True

    @Property(list, constant=True)
    def teamMonthGrid(self):
        return TEAM_GRID

    @Property(int, constant=True)
    def updateChromeExtra(self):
        return 0

    def setSearchText(self, t):
        pass

    def setActiveOnly(self, v):
        pass

    def selectEmployee(self, i):
        pass

    def reorderEmployees(self, a, b, c):
        pass


def part_b_render(app) -> None:
    qml_path = os.path.join(ROOT, "_render_teamstrip.qml")
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
        assert engine.rootObjects(), "EmployeeListPanel не загрузился"
        for _ in range(6):
            app.processEvents()
        win = engine.rootObjects()[0]

        # мини-ячейки: по 31 на сотрудника (делегаты Repeater из Python
        # не видны — спрашиваем у самой панели)
        panel = None
        for o in win.findChildren(QObject):
            if hasattr(o, "teamStripInfo"):
                panel = o
                break
        assert panel is not None, "панель не нашлась"
        info = panel.teamStripInfo()
        if hasattr(info, "toVariant"):
            from PySide6.QtQml import QJSValue
            if isinstance(info, QJSValue):
                info = info.toVariant()
        assert info["cells"] == 62, \
            "ожидали 62 мини-ячейки (2 × 31), есть %s" % info
        assert info["cellWidth"] > 4, "ячейки сжались в ноль: %s" % info
        print("B: полоса нарисована — 62 мини-ячейки (2 × 31), "
              "ширина %.1f px ✓" % info["cellWidth"])
    finally:
        try:
            os.remove(qml_path)
        except OSError:
            pass


def part_c_structure() -> None:
    panel = open(os.path.join(ROOT, "components", "EmployeeListPanel.qml"),
                 encoding="utf-8").read()
    assert "teamMode" in panel and "TeamDayCell" in panel
    assert "backend.teamMonthGrid" in panel
    settings = open(os.path.join(ROOT, "components", "SettingsDialog.qml"),
                    encoding="utf-8").read()
    assert "backend.teamStripEnabled" in settings, "нет тумблера в настройках"
    assert "setTeamStripEnabled" in settings
    cal = open(os.path.join(ROOT, "components", "CalendarWorkspace.qml"),
               encoding="utf-8").read()
    # маскот остаётся только при выключенной полосе
    assert "selectedEmployeeId === 0 && !backend.teamStripEnabled" in cal
    main = open(os.path.join(ROOT, "main.qml"), encoding="utf-8").read()
    assert "workspaceRoot.teamMode" in main, "список не расширяется"
    print("C: тумблер в «Внешнем виде», маскот и «Пн..Вс» уступают место, "
          "список расширяется ✓")


def main() -> int:
    part_a_builder()

    app = QGuiApplication.instance() or QGuiApplication(sys.argv)
    part_b_render(app)
    part_c_structure()

    print("═══ ОБЩАЯ КАРТИНА МЕСЯЦА У КАЖДОГО СОТРУДНИКА РАБОТАЕТ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
