#!/usr/bin/env python3
"""Анимация загрузки обновления: облако, превращающееся в кольцо
(components/UpdateCloudRing.qml, по присланному образцу-гифке).

Проверяем по фазам цикла (5220 мс, тайминги образца):
  - clock 0    — ОБЛАКО: силуэт с буграми (радиальный профиль
    «неровный»: std > 8), шире высокого (аспект > 1.2), внутри
    маленькое облачко (пиксели у центра);
  - clock 1900 — КОЛЬЦО: профиль ровный (std < 3), аспект ≈ 1,
    у центра пусто (облачко растворилось);
  - clock 900 и 3400 — середины морфов (профиль между);
  - подпись (caption) рисуется ПОД холстом;
  - цикл живёт сам (clock идёт вперёд при running).

Интеграция в кнопку обновления шапки (UpdateDownloadButton):
  - загрузка  → облако-морф видно, дуга прогресса скрыта,
    подсказка «Загрузка обновления… N%»;
  - готово    → облако скрыто, кольцо-дуга на месте.

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_update_loader.py
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "qa", "render")

import numpy as np

from PySide6.QtCore import QObject, Property, Signal, Slot
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuick import QQuickItem, QQuickWindow

CAPTION = "Загрузка обновления… 45%"

WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 220; height: 240; visible: true
    color: "#FFFFFF"
    AppUI.UpdateCloudRing {
        id: ring
        x: 20; y: 20
        diameter: 160
        running: false
        clock: 0
        caption: "%s"
    }
}
""" % CAPTION


class StubBackend(QObject):
    """Свойства с notify — кнопка на них переключает состояния."""

    remoteUpdateAvailableChanged = Signal()
    remoteDownloadingChanged = Signal()
    remoteDownloadProgressChanged = Signal()
    updateReadyChanged = Signal()

    def __init__(self):
        super().__init__()
        self._avail = False
        self._down = True
        self._prog = 45
        self._ready = False

    @Property(bool, notify=remoteUpdateAvailableChanged)
    def remoteUpdateAvailable(self):
        return self._avail

    @Property(bool, notify=remoteDownloadingChanged)
    def remoteDownloading(self):
        return self._down

    @Property(int, notify=remoteDownloadProgressChanged)
    def remoteDownloadProgress(self):
        return self._prog

    @Property(bool, notify=updateReadyChanged)
    def updateReady(self):
        return self._ready

    def set_state(self, down=None, ready=None):
        if down is not None and down != self._down:
            self._down = down
            self.remoteDownloadingChanged.emit()
        if ready is not None and ready != self._ready:
            self._ready = ready
            self.updateReadyChanged.emit()

    @Slot()
    def checkAllUpdateSources(self):
        pass

    @Slot()
    def startRemoteDownload(self):
        pass

    @Slot()
    def applyReadyUpdate(self):
        pass


def settle(app, sec):
    deadline = time.time() + sec
    while time.time() < deadline:
        app.processEvents()
        time.sleep(0.02)


def to_img(win, x0, y0, x1, y1):
    img = win.grabWindow()
    arr = np.frombuffer(img.constBits(), dtype=np.uint8).reshape(
        img.height(), img.width(), 4)[:, :, :3].copy()
    return arr[y0:y1, x0:x1]


def radial(arr):
    """маска != белого → профиль радиусов от центра (72 луча)."""
    m = np.abs(arr.astype(int) - 255).sum(axis=2) > 60
    h, w = m.shape
    cx, cy = w / 2, h / 2
    prof = []
    for a in range(72):
        ang = np.deg2rad(a * 5)
        rr = 0
        for r in range(4, int(min(cx, cy)) - 2):
            x, y = int(cx + r * np.cos(ang)), int(cy + r * np.sin(ang))
            if m[y, x]:
                rr = r
        prof.append(rr)
    return np.array(prof, dtype=float), m


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    wrapper = os.path.join(ROOT, "_render_updring.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER)
    engine.load(wrapper)
    assert engine.rootObjects(), "окно не загрузилось"
    win = engine.rootObjects()[0]
    settle(app, 1.2)

    ring = None
    for o in win.findChildren(QObject):
        if o.metaObject().className().startswith("UpdateCloudRing"):
            ring = o
            break
    assert ring is not None, "компонент не найден"

    def phase(clock):
        ring.setProperty("clock", clock)
        settle(app, 0.35)
        return to_img(win, 20, 20, 180, 180)

    # ── 1. облако ──
    arr = phase(0)
    prof, m = radial(arr)
    ys, xs = np.where(m)
    aspect = (xs.max() - xs.min()) / float(ys.max() - ys.min())
    assert prof.std() > 8, "силуэт облака слишком ровный: std=%.1f" % prof.std()
    assert aspect > 1.2, "облако не шире высокого: %.2f" % aspect
    inner = int(m[60:100, 60:100].sum())
    assert inner > 40, "внутри облака нет облачка: %d px" % inner
    print("1. Облако: бугристый силуэт (std=%.1f, аспект %.2f), "
          "внутри облачко (%d px) ✓" % (prof.std(), aspect, inner))

    # ── 2. кольцо ──
    arr = phase(1900)
    prof, m = radial(arr)
    ys, xs = np.where(m)
    aspect = (xs.max() - xs.min()) / float(ys.max() - ys.min())
    assert prof.std() < 3, "кольцо неровное: std=%.1f" % prof.std()
    assert abs(aspect - 1.0) < 0.06, "кольцо не круг: аспект %.2f" % aspect
    inner = int(m[60:100, 60:100].sum())
    assert inner < 20, "внутри кольца что-то осталось: %d px" % inner
    print("2. Кольцо: ровный круг (std=%.1f, аспект %.2f), "
          "облачко растворилось ✓" % (prof.std(), aspect))

    # ── 3. морфы посередине: профиль между облаком и кольцом ──
    for c, name in ((900, "облако→кольцо"), (3400, "кольцо→облако")):
        arr = phase(c)
        prof, _ = radial(arr)
        assert 2 < prof.std() < 16, \
            "%s: странный профиль посередине морфа (std=%.1f)" % (name, prof.std())
        print("3. Морф %s на середине: промежуточный силуэт (std=%.1f) ✓"
              % (name, prof.std()))

    # ── 4. подпись под холстом ──
    cap = None
    for o in ring.findChildren(QObject):
        if o.metaObject().className() == "QQuickText":
            cap = o
            break
    assert cap is not None, "подпись не найдена"
    assert str(cap.property("text")) == CAPTION, cap.property("text")
    assert float(cap.property("y")) >= 160, \
        "подпись не под холстом: y=%s" % cap.property("y")
    assert cap.property("visible") is True
    print("4. Подпись статуса — под анимацией ✓")

    # ── 5. цикл живёт сам ──
    ring.setProperty("caption", "")
    ring.setProperty("running", True)
    settle(app, 0.5)
    c1 = float(ring.property("clock"))
    settle(app, 0.5)
    c2 = float(ring.property("clock"))
    assert c1 > 0 and c2 > c1, "часы цикла не идут: %.1f → %.1f" % (c1, c2)
    ring.setProperty("running", False)
    print("5. Цикл крутится сам (clock %.0f → %.0f мс) и останавливается ✓"
          % (c1, c2))

    # ── 6. кнопка обновления: загрузка → облако-морф, готово → кольцо ──
    engine2 = QQmlApplicationEngine()
    stub = StubBackend()
    engine2.rootContext().setContextProperty("backend", stub)
    wrapper2 = os.path.join(ROOT, "_render_updbtn.qml")
    with open(wrapper2, "w", encoding="utf-8") as f:
        f.write("""
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 200; height: 60; visible: true
    Item {
        x: 10; y: 7; width: 46; height: 46
        AppUI.UpdateDownloadButton { anchors.fill: parent }
    }
}
""")
    engine2.load(wrapper2)
    assert engine2.rootObjects(), "кнопка не загрузилась"
    win2 = engine2.rootObjects()[0]
    settle(app, 1.0)     # дебаунс 80 мс + анимации слоёв

    def find_btn():
        for o in win2.findChildren(QObject):
            if o.metaObject().className().startswith("UpdateDownloadButton"):
                return o
        return None

    btn = find_btn()
    assert btn is not None
    cloud = None
    for o in btn.findChildren(QObject):
        if str(o.property("objectName") or "") == "updCloudRing":
            cloud = o
    assert cloud is not None, "в кнопке нет облака-морфа"
    # канвасы: первый — внутри облака-морфа, второй — дуга-кольцо
    canvases = [o for o in btn.findChildren(QObject)
                if o.metaObject().className() == "QQuickCanvasItem"]
    assert len(canvases) == 2, "ожидали два канваса: %d" % len(canvases)
    arcs = canvases[-1]
    assert arcs.property("visible") is False, \
        "дуга прогресса не скрылась при загрузке"
    assert cloud.property("visible") is True and \
        cloud.property("running") is True, "морф не включился при загрузке"
    tip = ""
    for o in btn.findChildren(QObject):
        if o.metaObject().className().startswith("AppToolTip"):
            tip = str(o.property("text") or "")
    assert "45%" in tip and "Загрузка" in tip, "подсказка без процентов: %r" % tip
    print("6а. Кнопка при загрузке: облако-морф работает, "
          "подсказка «Загрузка обновления… 45%» ✓")

    stub.set_state(down=False, ready=True)
    settle(app, 1.0)
    assert cloud.property("visible") is False, "морф не скрылся после загрузки"
    assert arcs.property("visible") is True, \
        "кольцо-дуга не вернулось в состоянии «готово»"
    print("6б. Кнопка в «готово»: облако скрылось, кольцо на месте ✓")

    for w_ in (wrapper, wrapper2):
        try:
            os.remove(w_)
        except OSError:
            pass
    print("═══ ОБЛАКО→КОЛЬЦО: АНИМАЦИЯ ЗАГРУЗКИ ОБНОВЛЕНИЯ ✓ ═══")
    return 0


if __name__ == "__main__":
    sys.exit(main())
