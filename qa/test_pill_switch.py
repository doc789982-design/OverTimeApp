#!/usr/bin/env python3
"""Большой переключатель-пилюля (components/AppPillSwitch.qml) —
перенос CSS-свотча. Главный эффект: КРУЖОК РАСШИРЯЕТСЯ — ножка
старого состояния растёт (600 мс) и заливает пилюлю цветом
нового; прежнее поле мгновенно сжимается в кружок, но лежит ПОД
волной и проявляется как новая ножка только В КОНЦЕ (смена
слоёв с задержкой 600 мс, transition z-index 0s .6s образца).

Проверяется:
  1. при открытии НЕТ холостой анимации: волна уже в покое
     (scale = cover точно, а не «доезжает»);
  2. покой (выкл): тёмное поле, светлая ножка справа НАД полем;
     размеры по умолчанию «округлые» — пилюля 64×32 (2:1, как
     252×126 образца);
  3. клик: тёмная волна сжалась МГНОВЕННО и УШЛА ПОД растущую
     (z: светлая сверху), центр ещё тёмный — расширение видно;
  4. конец (600 мс): светлая залила пилюлю, слои ПОМЕНЯЛИСЬ —
     тёмная ножка проявилась сверху слева; фон после 480 мс
     снова светлый;
  5. выключение: зеркально — тёмная растёт сверху, в конце
     светлая ножка справа; картина покоя вернулась;
  6. цвета — палитра темы (#2D3B45 / #FFFFFF);
  7. снимок-лента фаз: qa/render/pill_switch.png.

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_pill_switch.py
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

from PySide6.QtCore import QObject, QMetaObject
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine
from PySide6.QtQuick import QQuickWindow

WRAPPER = """
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 500; height: 300; visible: true
    color: "#F5F5F5"
    AppUI.AppPillSwitch {
        id: sw
        objectName: "sw"
        x: 100; y: 60
        pillWidth: 252
        pillHeight: 126
        withShadow: true
    }
    AppUI.AppPillSwitch {
        id: sw2
        objectName: "sw2"
        x: 100; y: 220
    }
}
"""

DARK = (0x2D, 0x3B, 0x45)
LIGHT = (0xFF, 0xFF, 0xFF)


def settle(app, sec):
    deadline = time.time() + sec
    while time.time() < deadline:
        app.processEvents()
        time.sleep(0.01)


def near(p, ref, tol=40):
    return abs(p[0]-ref[0]) + abs(p[1]-ref[1]) + abs(p[2]-ref[2]) < tol


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    wrapper = os.path.join(ROOT, "_render_pillswitch.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write(WRAPPER)
    engine.load(wrapper)
    assert engine.rootObjects(), "окно не загрузилось"
    settle(app, 1.0)
    win = next(x for x in app.topLevelWindows()
               if isinstance(x, QQuickWindow))

    def find_switch(name):
        for o in win.findChildren(QObject):
            if str(o.property("objectName") or "") == name:
                return o
        return None

    sw = find_switch("sw")
    sw2 = find_switch("sw2")
    assert sw is not None and sw2 is not None

    rects = [o for o in sw.findChildren(QObject)
             if o.metaObject().className().startswith("QQuickRectangle")]
    pill = [r for r in rects
            if abs(float(r.property("width") or 0) - 252) < 1
            and abs(float(r.property("height") or 0) - 126) < 1][0]
    knob = 126 * 80 / 126
    waves = [r for r in rects
             if abs(float(r.property("width") or 0) - knob) < 1]
    assert len(waves) == 2, "ожидали две волны-ножки: %d" % len(waves)
    dark = [r for r in waves if float(r.property("x") or 0) < 100][0]
    light = [r for r in waves if float(r.property("x") or 0) > 100][0]
    cover = float(sw.property("coverScale"))

    def snap():
        img = win.grabWindow()
        arr = np.frombuffer(img.constBits(), dtype=np.uint8).reshape(
            img.height(), img.width(), 4)[:, :, :3].copy()
        return arr[:, :, ::-1]      # BGR32 → RGB

    # большой свитч: предмет 36 центрирует пилюлю 126 → верх (60-45)
    PX, PY = 100, 60 - 45

    def px(arr, x, y):
        return tuple(int(v) for v in arr[PY + y, PX + x])

    # ── 1. при открытии нет холостой анимации ──
    assert abs(float(dark.property("scale")) - cover) < 0.001, \
        "волна «доезжает» после открытия: %s" % dark.property("scale")
    assert abs(float(light.property("scale")) - 1.0) < 0.001
    # «округлые» пропорции по умолчанию: 64×32 = 2:1, как 252×126
    assert abs(float(sw2.property("pillWidth")) - 64) < 0.01
    assert abs(float(sw2.property("pillHeight")) - 32) < 0.01
    assert abs(float(sw2.property("implicitWidth")) - 64) < 0.01
    print("1. Открытие без холостой анимации; умолчание — пилюля 64×32 "
          "(2:1, пропорции образца) ✓")

    # ── 2. покой (выкл) ──
    arr = snap()
    assert near(px(arr, 126, 63), DARK), "центр не тёмный: %r" % (px(arr, 126, 63),)
    assert near(px(arr, 189, 63), LIGHT), "ножка справа не белая: %r" % (px(arr, 189, 63),)
    assert int(light.property("z")) > int(dark.property("z")), \
        "ножка (светлая) должна быть над полем"
    print("2. Покой (выкл): тёмное поле, светлая ножка справа над ним ✓")

    # ── 3. клик: сжатие мгновенно и ПОД волной, расширение видно ──
    fired = []
    sw.toggled.connect(lambda: fired.append(1))
    QMetaObject.invokeMethod(sw, "toggle")
    settle(app, 0.08)
    assert fired and sw.property("checked") is True
    assert abs(float(dark.property("scale")) - 1.0) < 0.05, \
        "тёмная не сжалась мгновенно: %s" % dark.property("scale")
    assert 1.0 < float(light.property("scale")) < cover, \
        "светлая не растёт: %s" % light.property("scale")
    # ГЛАВНОЕ: растущая волна СВЕРХУ, сжавшийся круг ПОД ней —
    # глазу видно только расширение, ножка проявится в конце
    assert int(light.property("z")) > int(dark.property("z")), \
        "растущая волна не сверху: z=%s vs %s" % (light.property("z"),
                                                  dark.property("z"))
    arr = snap()
    assert near(px(arr, 126, 63), DARK), \
        "центр зальлся раньше времени (нет эффекта расширения)"
    print("3. Клик: сжатие мгновенно и ПОД волной; видно только "
          "расширение светлого круга ✓")

    # ── 4. конец анимации: слои поменялись, ножка проявилась ──
    settle(app, 0.7)      # ~780 мс от клика
    assert abs(float(light.property("scale")) - cover) < 0.01
    assert abs(float(dark.property("scale")) - 1.0) < 0.01
    assert int(dark.property("z")) > int(light.property("z")), \
        "ножка (тёмная) не проявилась в конце (слои не поменялись)"
    arr = snap()
    assert near(px(arr, 126, 63), LIGHT), "центр не белый: %r" % (px(arr, 126, 63),)
    assert near(px(arr, 63, 63), DARK), "ножка слева не тёмная: %r" % (px(arr, 63, 63),)
    pill_color = pill.property("color")
    assert near(tuple(int(v) for v in pill_color.toTuple()[:3]), LIGHT), \
        "фон пилюли не вернулся к светлому: %r" % (pill_color.toTuple(),)
    print("4. Конец: волна залила пилюлю, тёмная ножка проявилась "
          "сверху слева, фон выдержан 480 мс ✓")

    # ── 5. выключение: зеркально ──
    frames_on = snap()
    QMetaObject.invokeMethod(sw, "toggle")
    settle(app, 0.08)
    assert abs(float(light.property("scale")) - 1.0) < 0.05, \
        "светлая не сжалась мгновенно"
    assert 1.0 < float(dark.property("scale")) < cover, "тёмная не растёт"
    assert int(dark.property("z")) > int(light.property("z")), \
        "при выключении растущая (тёмная) волна не сверху"
    settle(app, 0.75)
    arr = snap()
    assert near(px(arr, 126, 63), DARK), "обратно: центр не тёмный"
    assert near(px(arr, 189, 63), LIGHT), "обратно: ножка справа не белая"
    assert int(light.property("z")) > int(dark.property("z")), \
        "обратно: светлая ножка не сверху"
    print("5. Выключение зеркально: тёмная растёт сверху, в конце "
          "светлая ножка справа ✓")

    # ── 6. цвета ──
    assert str(sw.property("darkColor").name()).lower() == "#2d3b45"
    assert str(sw.property("lightColor").name()).lower() == "#ffffff"
    print("6. Цвета сторон — палитра темы (#2D3B45 / #FFFFFF) ✓")

    # ── 7. снимок-лента фаз ──
    from PIL import Image, ImageDraw
    QMetaObject.invokeMethod(sw, "toggle")     # включаем
    settle(app, 0.28)
    mid = snap()                               # расширение в пути
    settle(app, 0.55)
    on = snap()
    QMetaObject.invokeMethod(sw, "toggle")     # выключаем
    settle(app, 0.28)
    mid_off = snap()
    settle(app, 0.55)
    off = snap()
    tiles = [("выключен", off), ("включается", mid),
             ("включён", on), ("выключается", mid_off)]
    tw, th = 252, 126
    strip = Image.new("RGB", (tw * 4 + 50, th + 34), "#F5F5F5")
    d = ImageDraw.Draw(strip)
    for i, (name, a) in enumerate(tiles):
        im = Image.fromarray(a[PY:PY + th, PX:PX + tw])
        x = 10 + i * (tw + 10)
        strip.paste(im, (x, 10))
        d.text((x + tw // 2 - 34, th + 18), name, fill="#5F6B74")
    shot = os.path.join(OUT, "pill_switch.png")
    strip.save(shot)
    assert os.path.getsize(shot) > 4000, "пустой снимок"
    print("7. Снимок-лента фаз: qa/render/pill_switch.png ✓")

    os.remove(wrapper)
    print("═══ ПИЛЮЛЯ: КРУЖОК РАСШИРЯЕТСЯ, НОЖКА ПРОЯВЛЯЕТСЯ В КОНЦЕ ✓ ═══")
    return 0


if __name__ == "__main__":
    sys.exit(main())
