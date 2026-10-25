#!/usr/bin/env python3
"""Тест диалога подтверждения после пересадки на AppDialog (батч 297).

Проверяется контракт, который был у старого самодельного Popup и обязан
сохраниться: Enter/клик — accepted и окно закрылось; Esc/«Отмена»/крестик —
rejected и закрылось; фокус при открытии — на кнопке подтверждения
(Enter/Space нативно); dangerMode красит кнопку.
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from PySide6.QtCore import Qt, QTimer, QObject, Property, Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuick import QQuickItem
from PySide6.QtTest import QTest


class StubBackend(QObject):
    """Минимальный backend: AppTheme спрашивает только тему."""
    isDarkChanged = Signal()

    def __init__(self):
        super().__init__()
        self._dark = False

    def getDark(self):
        return self._dark

    def setDark(self, v):
        self._dark = bool(v)

    isDarkTheme = Property(bool, getDark, setDark, notify=isDarkChanged)

WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 700; height: 480; visible: true
    property int acc: 0
    property int rej: 0
    AppUI.AppConfirmDialog {
        id: dlg
        dialogTitle: "Удалить группу?"
        dialogMessage: "Действие нельзя отменить."
        confirmLabel: "Удалить"
        dangerMode: true
        onAccepted: w.acc++
        onRejected: w.rej++
    }
    Timer { interval: 200; running: true; onTriggered: dlg.open() }
}
"""


def wait_until(app, cond, timeout=3.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for _ in range(5):
            app.processEvents()
        if cond():
            return True
        time.sleep(0.02)
    return cond()


def buttons_of(win):
    """Все AppButton окна (у диалога это «Отмена» и «Удалить»)."""
    return [it for it in win.findChildren(QQuickItem)
            if "AppButton" in it.metaObject().className()]


def main() -> int:
    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    stub = StubBackend()   # держим ссылку: иначе GC съест и backend станет null
    engine.rootContext().setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_test_confirm.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER)
    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "окно не загрузилось"
        win = engine.rootObjects()[0]

        def get_dlg():
            # попап — QObject (QQuickPopup), не QQuickItem
            for obj in win.findChildren(QObject):
                if "AppConfirmDialog" in obj.metaObject().className():
                    return obj
            return None

        assert wait_until(app, lambda: get_dlg() is not None), "диалог не открылся"
        dlg = get_dlg()

        # ── A. Открытие и фокус на подтверждении ──
        assert wait_until(app, lambda: dlg.property("opened")), "не открылся"
        assert wait_until(app, lambda: win.property("acc") == 0), "лишний accepted"
        btns = buttons_of(win)
        assert len(btns) == 2, "кнопок %d, ожидали 2 (Отмена + Удалить)" % len(btns)
        acc_btn = next((b for b in btns if b.property("variant") == "danger"), None)
        rej_btn = next((b for b in btns if b.property("variant") == "secondary"), None)
        assert acc_btn is not None and rej_btn is not None, "нет danger/secondary кнопок"
        assert wait_until(app, lambda: acc_btn.property("activeFocus") is True), \
            "фокус не на кнопке подтверждения"
        print("A: открылся, фокус на «Удалить», кнопка danger ✓")

        # ── B. Клик по подтверждению: accepted + закрылось ──
        acc_btn.click()
        assert wait_until(app, lambda: win.property("acc") == 1), "accepted не пришёл"
        assert wait_until(app, lambda: not dlg.property("opened")), "не закрылся"
        print("B: клик «Удалить» → accepted, окно закрылось ✓")

        # ── C. Повторное открытие, Enter с клавиатуры ──
        QTimer.singleShot(0, dlg.open)
        assert wait_until(app, lambda: dlg.property("opened")), "не открылся повторно"
        assert wait_until(app, lambda: acc_btn.property("activeFocus") is True)
        QTest.keyClick(win, Qt.Key_Return)
        assert wait_until(app, lambda: win.property("acc") == 2), "Enter не подтвердил"
        assert wait_until(app, lambda: not dlg.property("opened")), "Enter не закрыл"
        print("C: Enter → accepted, окно закрылось ✓")

        # ── D. Esc → rejected + закрылось ──
        QTimer.singleShot(0, dlg.open)
        assert wait_until(app, lambda: dlg.property("opened"))
        QTest.keyClick(win, Qt.Key_Escape)
        assert wait_until(app, lambda: win.property("rej") == 1), "Esc не отменил"
        assert wait_until(app, lambda: not dlg.property("opened")), "Esc не закрыл"
        print("D: Esc → rejected, окно закрылось ✓")

        # ── E. Кнопка «Отмена» → rejected + закрылось ──
        QTimer.singleShot(0, dlg.open)
        assert wait_until(app, lambda: dlg.property("opened"))
        rej_btn.click()
        assert wait_until(app, lambda: win.property("rej") == 2), "Отмена не отменила"
        assert wait_until(app, lambda: not dlg.property("opened")), "Отмена не закрыла"
        print("E: «Отмена» → rejected, окно закрылось ✓")

        print("═══ ДИАЛОГ ПОДТВЕРЖДЕНИЯ: КОНТРАКТ СОХРАНЁН ═══")
        return 0
    finally:
        try:
            os.remove(wrapper)
        except OSError:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
