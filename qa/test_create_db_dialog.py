#!/usr/bin/env python3
"""Тест окна создания базы с полями ответственного лица (батч 298).

Проверяется: поля ответственного лица уходят в createNewDatabase вместе
с названием; без названия база не создаётся (ошибка в окне, не тост).
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from PySide6.QtCore import QObject, Property, Signal, Slot
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuick import QQuickItem

WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 700; height: 600; visible: true
    AppUI.CreateDbDialog { id: dlg }
    Timer { interval: 200; running: true; onTriggered: dlg.open() }
}
"""


class StubBackend(QObject):
    """Запоминает, с чем позвали createNewDatabase."""
    isDarkChanged = Signal()
    calls = []

    def __init__(self):
        super().__init__()
        self._dark = False

    def getDark(self):
        return self._dark

    def setDark(self, v):
        self._dark = bool(v)

    isDarkTheme = Property(bool, getDark, setDark, notify=isDarkChanged)

    @Slot(str, str, str, str, str, str)
    def createNewDatabase(self, name, pos, rank, last, first, mid):
        StubBackend.calls.append((name, pos, rank, last, first, mid))


def wait_until(app, cond, timeout=3.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for _ in range(5):
            app.processEvents()
        if cond():
            return True
        time.sleep(0.02)
    return cond()


def fields_of(win):
    """AppTextField окна по label, затем по порядку."""
    return [it for it in win.findChildren(QQuickItem)
            if "AppTextField" in it.metaObject().className()]


def main() -> int:
    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    stub = StubBackend()
    engine.rootContext().setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_test_createdb.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER)
    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "окно не загрузилось"
        win = engine.rootObjects()[0]
        dlg = next(o for o in win.findChildren(QObject)
                   if "CreateDbDialog" in o.metaObject().className())
        # поля лежат в дереве ещё до открытия — ждём именно открытия
        assert wait_until(app, lambda: dlg.property("opened")), "диалог не открылся"
        flds = fields_of(win)
        labels = [f.property("label") for f in flds]
        for want in ("Название отдела", "Должность", "Звание",
                     "Фамилия", "Имя", "Отчество"):
            assert want in labels, "нет поля «%s» (есть %s)" % (want, labels)
        print("A: все шесть полей в окне ✓")

        # ── B. Пустое название: база не создаётся, окно остаётся ──
        # сигнал эмитим вручную — это то же, что кнопка «Создать»
        dlg.accepted.emit()
        app.processEvents()
        assert not StubBackend.calls, "база создалась без названия!"
        assert dlg.property("opened"), "окно закрылось без названия"
        print("B: без названия база не создаётся, окно на месте ✓")

        # ── C. Заполнено: все шесть значений уходят в базу ──
        values = dict(zip(labels, ["1 Отдел", "Начальник отдела",
                                   "Подполковник полиции",
                                   "Иванов", "Пётр", "Сергеевич"]))
        for f in flds:
            f.setProperty("text", values[f.property("label")])
        dlg.accepted.emit()
        app.processEvents()
        assert wait_until(app, lambda: len(StubBackend.calls) == 1), \
            "createNewDatabase не вызван"
        got = StubBackend.calls[0]
        want = ("1 Отдел", "Начальник отдела", "Подполковник полиции",
                "Иванов", "Пётр", "Сергеевич")
        assert got == want, "в базу ушло %r" % (got,)
        assert wait_until(app, lambda: not dlg.property("opened")), "окно не закрылось"
        print("C: ответственное лицо уходит в базу вместе с названием ✓")

        print("═══ ОКНО СОЗДАНИЯ БАЗЫ: ПОЛЯ ОТВЕТСТВЕННОГО ЛИЦА РАБОТАЮТ ═══")
        return 0
    finally:
        try:
            os.remove(wrapper)
        except OSError:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
