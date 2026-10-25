#!/usr/bin/env python3
"""Рендер окон, затронутых унификацией батча 297: оболочки, шапки, поля, меню.

Снимки в qa/render/refactor_*.png. Проверяем глазами: модальные окна
на общей оболочке (AppModalShell), общая шапка (AppDialogHeader),
поля на общей базе (AppFieldBase), меню на общей базе (AppMenu).

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen QSG_RASTER_BACKEND=1 \
    python3 qa/render_refactor.py
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
OUT = os.path.join(ROOT, "qa", "render")

from PySide6.QtCore import QObject, Property, Signal, Slot
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuick import QQuickWindow


class StubBackend(QObject):
    """Минимальный backend: календарь и рекомендации — настоящие данные."""

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

    @Slot(result="QVariant")
    def getRecoSections(self):
        try:
            from methodical_data import RECO_SECTIONS
            return list(RECO_SECTIONS)
        except Exception:
            return []

    # меню дня: вызовы без ответа, меню рисуется по своим умолчаниям
    @Slot(str)
    def loadDayDetails(self, d):
        pass

    @Slot(str, str)
    def setDayType(self, d, t):
        pass

    @Slot(str, str)
    def setDayStatus(self, d, s):
        pass


# Каждая сцена — отдельная обёртка (движок перезагружаем, чтобы окна
# не мешали друг другу). Обёртки живут в КОРНЕ репозитория — иначе
# синглтон AppTheme не увидит контекстное свойство backend.
SCENES = {
    # 1. Подтверждение (опасное) — теперь обычный AppDialog
    "confirm_danger": """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 700; height: 480; visible: true
    AppUI.AppConfirmDialog {
        id: dlg
        dialogTitle: "Удалить группу?"
        dialogMessage: "В группе 12 сотрудников. Действие нельзя отменить после закрытия окна."
        confirmLabel: "Удалить"
        dangerMode: true
    }
    Timer { interval: 400; running: true; onTriggered: dlg.open() }
}
""",
    # 2. Подтверждение (спокойное) — синяя кнопка
    "confirm_safe": """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 700; height: 480; visible: true
    AppUI.AppConfirmDialog {
        id: dlg
        dialogTitle: "Открыть ведомость?"
        dialogMessage: "Ведомость за текущий период уже сформирована."
        confirmLabel: "Открыть"
        dangerMode: false
    }
    Timer { interval: 400; running: true; onTriggered: dlg.open() }
}
""",
    # 3. Форма с полями: заполненное, пустое обязательное, дата; вспышка ошибки
    "fields": """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 700; height: 520; visible: true
    AppUI.AppDialog {
        id: dlg
        title: "Новый сотрудник"
        width: 460
        AppUI.AppTextField { id: f1; width: parent.width; label: "Фамилия"; text: "Иванов" }
        AppUI.AppTextField { id: f2; width: parent.width; label: "Имя"; isRequired: true }
        AppUI.AppDateField { id: f3; width: parent.width; label: "Дата приема"; isRequired: true }
        AppUI.AppTextField { id: f4; width: parent.width; label: "Часы (числовое)"; numericOnly: true; text: "120" }
    }
    Timer {
        interval: 400; running: true
        onTriggered: { dlg.open(); f2.flashError() }
    }
}
""",
    # 4. Большое окно (настройки/печать) — AppLargeModal на оболочке
    "large_modal": """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 900; height: 620; visible: true
    AppUI.AppLargeModal {
        id: dlg
        title: "Настройки"
        width: 720; height: 520
        Rectangle {
            width: parent.width; height: parent.height
            color: AppUI.AppTheme.bgInput
            radius: AppUI.AppTheme.radiusMedium
            Text {
                anchors.centerIn: parent
                text: "Контент большого окна"
                color: AppUI.AppTheme.textSecondary
                font.pixelSize: AppUI.AppTheme.sizeH4
            }
        }
    }
    Timer { interval: 400; running: true; onTriggered: dlg.open() }
}
""",
    # 5. Производственный календарь — общая шапка + годы в слоте
    "prod_calendar": """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 1100; height: 780; visible: true
    AppUI.ProdCalendarDialog { id: dlg }
    Timer { interval: 400; running: true; onTriggered: dlg.openYear(2026) }
}
""",
    # 6. Методические рекомендации — читалка с оглавлением
    "reco": """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 1100; height: 780; visible: true
    AppUI.RecoDialog { id: dlg }
    Timer { interval: 400; running: true; onTriggered: dlg.openReco() }
}
""",
    # 7. Обычное меню на базе AppMenu
    "menu": """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 500; height: 400; visible: true
    AppUI.AppMenu {
        id: m
        AppUI.AppMenuItem { text: "Открыть базу"; iconSource: "../icons/database.svg" }
        AppUI.AppMenuItem { text: "Настройки"; iconSource: "../icons/settings.svg" }
        AppUI.AppMenuSeparator {}
        AppUI.AppMenuItem { text: "Удалить"; isDanger: true; iconSource: "../icons/trash.svg" }
    }
    Timer { interval: 400; running: true; onTriggered: m.popup(60, 60) }
}
""",
    # 8. Меню дня — шапка с датой и статусы
    "day_menu": """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 500; height: 520; visible: true
    AppUI.AppDayMenu {
        id: m
        targetDate: "2026-10-11"
        targetDates: ["2026-10-11"]
    }
    Timer { interval: 400; running: true; onTriggered: m.popup(60, 60) }
}
""",
}


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    app = QGuiApplication([])
    shots = []
    for name, qml in SCENES.items():
        engine = QQmlApplicationEngine()
        ctx = engine.rootContext()
        # держим ссылку: без неё PySide6 соберёт заглушку, и backend
        # в QML станет null (как у окна печати в этом же харнессе)
        stub = StubBackend()
        ctx.setContextProperty("backend", stub)
        wrapper = os.path.join(ROOT, "_render_refactor.qml")
        with open(wrapper, "w", encoding="utf-8") as f:
            f.write(qml)
        try:
            engine.load(wrapper)
            assert engine.rootObjects(), "сцена %s не загрузилась" % name
            win = engine.rootObjects()[0]
            if not isinstance(win, QQuickWindow):
                wins = [w for w in app.allWindows() if isinstance(w, QQuickWindow)]
                assert wins, "нет QQuickWindow в сцене %s" % name
                win = wins[-1]
            # ждём открытия и успокоения анимаций
            deadline = time.time() + 2.5
            while time.time() < deadline:
                app.processEvents()
                time.sleep(0.02)
            img = win.grabWindow()
            path = os.path.join(OUT, "refactor_%s.png" % name)
            img.save(path)
            assert os.path.getsize(path) > 5000, "пустой снимок: %s" % name
            shots.append(path)
            print("сцена %s: %dx%d" % (name, img.width(), img.height()))
        finally:
            try:
                os.remove(wrapper)
            except OSError:
                pass
            engine.deleteLater()
            app.processEvents()
    print("снимки: " + ", ".join(os.path.basename(s) for s in shots))
    print("═══ ОКНА УНИФИКАЦИИ ОТРЕНДЕРЕНЫ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
