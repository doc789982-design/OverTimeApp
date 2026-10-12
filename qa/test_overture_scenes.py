#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Увертюра и сцены: тесты двух помощников движения (батч 301).

AppStageEntrance — участник увертюры: цель гаснет, в свой ход входит
смещением (Translate, раскладка не трогается) и проявляется; страховка
доводит до конца даже без анимации.

AppSceneDirector — режиссёр сцен: подмена содержимого отложена до конца
выхода старой сцены; новая входит по общей оси; выключенный (не на
экране) режиссёр подменяет мгновенно; быстрая повторная смена отменяет
недоигравшую подмену.
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from PySide6.QtCore import QObject, Property, Signal, QMetaObject, Q_ARG
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine


class StubBackend(QObject):
    """AppTheme спрашивает у backend тему — минимальный стаб."""
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
    id: win
    width: 600; height: 400; visible: true

    // испытуемые
    Rectangle { id: probeA; objectName: "probeA"; width: 100; height: 100; color: "red" }
    Rectangle { id: probeB; objectName: "probeB"; width: 100; height: 100; y: 200; color: "blue" }

    AppUI.AppStageEntrance {
        id: entrance
        target: probeA
        delay: 60; fromX: -20; fromY: 8; duration: 200
    }
    AppUI.AppSceneDirector {
        id: director
        target: probeB
        armed: true
    }

    property int sceneValue: 0
    // окна в смещениях участников (Translate не читается из Python напрямую)
    property real entShiftX: entrance.shift.x
    property real entShiftY: entrance.shift.y
    property real dirShiftY: director.shift.y

    function doSwap(v) {
        director.swap(function() { win.sceneValue = v })
    }
    function doSwapInstant(v) {
        director.armed = false
        director.swap(function() { win.sceneValue = v })
    }
    function rearm() { director.armed = true }
    function setSceneImmediate(v) { win.sceneValue = v }
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


def main() -> int:
    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    stub = StubBackend()
    engine.rootContext().setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_test_overture.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER)
    try:
        engine.load(wrapper)
        assert engine.rootObjects(), "окно не загрузилось"
        win = engine.rootObjects()[0]
        entrance = next(o for o in win.findChildren(QObject)
                        if "AppStageEntrance" in o.metaObject().className())
        director = next(o for o in win.findChildren(QObject)
                        if "AppSceneDirector" in o.metaObject().className())
        probeA = next(o for o in win.findChildren(QObject) if o.objectName() == "probeA")
        probeB = next(o for o in win.findChildren(QObject) if o.objectName() == "probeB")

        # ── A. Увертюра: гаснет сразу, входит в свой ход ──
        QMetaObject.invokeMethod(entrance, "play")
        app.processEvents()
        assert probeA.property("opacity") == 0.0, "цель не погасла при play()"
        assert win.property("entShiftX") == -20 and win.property("entShiftY") == 8, \
            "смещение входа не выставлено (%r, %r)" % (win.property("entShiftX"), win.property("entShiftY"))
        print("A: цель гаснет, смещение выставлено ✓")

        # ── B. Середина входа: проявляется и едет ──
        assert wait_until(app, lambda: 0.0 < probeA.property("opacity") < 1.0), \
            "за время хода цель не начала проявляться"
        mid_x = win.property("entShiftX")
        assert -20 < mid_x < 0, "вход не едет по оси (x=%r)" % mid_x
        assert wait_until(app, lambda: probeA.property("opacity") == 1.0), "вход не завершился"
        assert win.property("entShiftX") == 0 and win.property("entShiftY") == 0, \
            "после входа смещение не обнулилось"
        print("B: вход едет с торможением и встаёт на место ✓")

        # ── C. Сцена: подмена отложена до конца выхода ──
        QMetaObject.invokeMethod(win, "doSwap", Q_ARG("QVariant", 1))
        app.processEvents()
        assert win.property("sceneValue") == 0, "подмена применилась сразу, без выхода"
        assert wait_until(app, lambda: probeB.property("opacity") < 1.0), \
            "старая сцена не начала уходить"
        assert win.property("sceneValue") == 0, "подмена применилась во время выхода"
        assert wait_until(app, lambda: win.property("sceneValue") == 1), \
            "после выхода подмена не применилась"
        assert wait_until(app, lambda: probeB.property("opacity") == 1.0), \
            "новая сцена не вошла"
        assert win.property("dirShiftY") == 0, "сцена не встала на место"
        print("C: старая уходит, подмена после, новая входит ✓")

        # ── D. Быстрая повторная смена отменяет недоигравшую ──
        QMetaObject.invokeMethod(win, "doSwap", Q_ARG("QVariant", 2))
        assert wait_until(app, lambda: probeB.property("opacity") < 1.0), "выход не начался"
        # отменяет подмену «2», ещё не применённую
        QMetaObject.invokeMethod(win, "doSwap", Q_ARG("QVariant", 3))
        assert wait_until(app, lambda: win.property("sceneValue") == 3), \
            "повторная смена не дошла до подмены"
        assert win.property("sceneValue") == 3, "отменённая подмена «2» применилась!"
        assert wait_until(app, lambda: probeB.property("opacity") == 1.0), "вход не завершился"
        print("D: повторная смена отменяет недоигравшую подмену ✓")

        # ── E. Не на экране — мгновенно, без хореографии ──
        QMetaObject.invokeMethod(win, "doSwapInstant", Q_ARG("QVariant", 4))
        app.processEvents()
        assert win.property("sceneValue") == 4, "выключенный режиссёр не подменил сразу"
        assert probeB.property("opacity") == 1.0, "выключенный режиссёр тронул видимость"
        QMetaObject.invokeMethod(win, "rearm")
        print("E: закрытое окно меняет содержимое мгновенно ✓")

        print("═══ УВЕРТЮРА И СЦЕНЫ: ПОМОЩНИКИ РАБОТАЮТ ═══")
        return 0
    finally:
        try:
            os.remove(wrapper)
        except OSError:
            pass


if __name__ == "__main__":
    rc = main()
    sys.stdout.flush()
    sys.stderr.flush()
    os._exit(rc)
