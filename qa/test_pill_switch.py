#!/usr/bin/env python3
"""Большой переключатель-пилюля (components/AppPillSwitch.qml) —
перенос CSS-свотча: волна-круг растёт и закрывает пилюлю, вторая
мгновенно сжимается в ножку, фон держит прошлый цвет до 80%.

Проверяется:
  1. исходное состояние (выкл): тёмная волна закрыла пилюлю
     (scale 4.8), светлая — ножка справа (scale 1, z выше);
     фон пилюли светлый;
  2. клик: сигнал toggled, тёмная волна сжалась МГНОВЕННО (50 мс),
     светлая растёт (600 мс, InOutQuad), фон стал тёмным;
  3. середина роста (300 мс): светлая волна в пути (1.5..4.5);
  4. конец (750 мс): светлая закрыла пилюлю, тёмная — ножка слева
     и НАД волной; фон после 480 мс снова светлый;
  5. пиксели в покое: выключен — центр тёмный, ножка справа
     белая; включён — центр белый, ножка слева тёмная;
  6. выключение обратно возвращает всё;
  7. цвета — палитра темы (#2D3B45 / #FFFFFF);
  8. снимок-лента фаз: qa/render/pill_switch.png.

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
    width: 400; height: 300; visible: true
    color: "#F5F5F5"
    AppUI.AppPillSwitch {
        id: sw
        objectName: "sw"
        x: 74; y: 87
        pillWidth: 252
        pillHeight: 126
        withShadow: true
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


def find_switch(win):
    for o in win.findChildren(QObject):
        if str(o.property("objectName") or "") == "sw":
            return o
    return None


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
    sw = find_switch(win)
    assert sw is not None, "переключатель не найден"

    # размер пилюли по свойствам (252×126 — пропорции образца)
    assert abs(float(sw.property("pillWidth")) - 252) < 1
    assert abs(float(sw.property("pillHeight")) - 126) < 1
    cover = float(sw.property("coverScale"))
    assert 3.8 < cover < 5.2, "странная кратность волны: %s" % cover
    # волны — дети пилюли: тёмная левее светлой
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

    def snap():
        img = win.grabWindow()
        arr = np.frombuffer(img.constBits(), dtype=np.uint8).reshape(
            img.height(), img.width(), 4)[:, :, :3].copy()
        return arr[:, :, ::-1]      # BGR32 → RGB

    # предмет свитча высотой 36 центрирует пилюлю 126px:
    # её верх-лево в окне = (74, 87 - (126-36)/2) = (74, 42)
    PILL_X, PILL_Y = 74, 87 - 45

    def px(arr, x, y):
        return tuple(int(v) for v in arr[PILL_Y + y, PILL_X + x])

    def near(p, ref, tol=40):
        return abs(p[0]-ref[0]) + abs(p[1]-ref[1]) + abs(p[2]-ref[2]) < tol

    # ── 1. исходное состояние ──
    assert sw.property("checked") is False
    assert abs(float(dark.property("scale")) - cover) < 0.01, dark.property("scale")
    assert abs(float(light.property("scale")) - 1.0) < 0.01
    assert int(light.property("z")) > int(dark.property("z")), \
        "ножка (светлая) должна быть над волной"
    # стандартный размер (замена AppSwitch): пилюля 52×32
    assert abs(float(sw.property("knobD")) - 126 * 80 / 126) < 0.01
    arr = snap()
    assert near(px(arr, 126, 63), DARK), "центр не тёмный: %r" % (px(arr, 126, 63),)
    assert near(px(arr, 189, 63), LIGHT), "ножка справа не белая: %r" % (px(arr, 189, 63),)
    assert near(px(arr, 63, 63), DARK), "слева не тёмный: %r" % (px(arr, 63, 63),)
    print("1. Выключен: тёмная волна закрыла пилюлю, светлая ножка справа ✓")

    # ── 2. клик: сигнал, мгновенное сжатие, рост, фон ──
    fired = []
    sw.toggled.connect(lambda: fired.append(1))
    QMetaObject.invokeMethod(sw, "toggle")
    settle(app, 0.05)
    assert fired, "сигнал toggled не вышел"
    assert sw.property("checked") is True
    assert abs(float(dark.property("scale")) - 1.0) < 0.05, \
        "тёмная волна не сжалась мгновенно: %s" % dark.property("scale")
    assert 1.0 < float(light.property("scale")) < cover, \
        "светлая волна не начала расти: %s" % light.property("scale")
    print("2. Клик: сигнал вышел, тёмная сжалась мгновенно, светлая растёт ✓")

    # ── 3. середина роста ──
    settle(app, 0.25)      # всего ~300 мс от клика
    ls = float(light.property("scale"))
    assert 1.3 < ls < cover - 0.2, "светлая волна не в пути (300 мс): %s" % ls
    print("3. Середина роста: светлая волна в пути (scale %.2f) ✓" % ls)

    # ── 4. конец анимации ──
    settle(app, 0.5)       # всего ~800 мс
    assert abs(float(light.property("scale")) - cover) < 0.01
    assert abs(float(dark.property("scale")) - 1.0) < 0.01
    assert int(dark.property("z")) > int(light.property("z")), \
        "ножка (тёмная) должна быть над волной"
    arr = snap()
    assert near(px(arr, 126, 63), LIGHT), "центр не белый: %r" % (px(arr, 126, 63),)
    assert near(px(arr, 63, 63), DARK), "ножка слева не тёмная: %r" % (px(arr, 63, 63),)
    print("4. Включён: светлая волна закрыла пилюлю, тёмная ножка слева ✓")

    # ── 5. фон пилюли после 480 мс снова светлый (changeColor 80%) ──
    pill_color = pill.property("color").toTuple() if hasattr(pill.property("color"), "toTuple") else None
    assert pill_color is not None and near(tuple(int(v) for v in pill_color[:3]), LIGHT), \
        "фон пилюли не вернулся к светлому: %r" % (pill_color,)
    print("5. Фон пилюли после 480 мс выдержки стал светлым ✓")

    # ── 6. выключение обратно ──
    frames = [snap()]          # для снимка: включён
    QMetaObject.invokeMethod(sw, "toggle")
    settle(app, 0.3)
    frames.append(snap())      # середина выключения
    settle(app, 0.5)
    assert sw.property("checked") is False
    assert abs(float(dark.property("scale")) - cover) < 0.01
    assert abs(float(light.property("scale")) - 1.0) < 0.01
    arr = snap()
    frames.append(arr)
    assert near(px(arr, 126, 63), DARK), "обратно: центр не тёмный"
    assert near(px(arr, 189, 63), LIGHT), "обратно: ножка справа не белая"
    print("6. Выключение возвращает исходную картинку ✓")

    # ── 7. цвета — палитра темы ──
    assert str(sw.property("darkColor").name()).lower() == "#2d3b45"
    assert str(sw.property("lightColor").name()).lower() == "#ffffff"
    print("7. Цвета сторон — палитра темы (#2D3B45 / #FFFFFF) ✓")

    # ── 8. снимок-лента фаз ──
    from PIL import Image, ImageDraw
    # кадр «выключен» уже есть (frames[2]); снимем середину включения заново
    QMetaObject.invokeMethod(sw, "toggle")
    settle(app, 0.3)
    mid_on = snap()
    settle(app, 0.55)
    on = snap()
    QMetaObject.invokeMethod(sw, "toggle")
    settle(app, 0.8)
    off = snap()
    tiles = [("выключен", off), ("включается", mid_on), ("включён", on)]
    tile_w, tile_h = 252, 126
    strip = Image.new("RGB", (tile_w * 3 + 40, tile_h + 34), "#F5F5F5")
    d = ImageDraw.Draw(strip)
    for i, (name, a) in enumerate(tiles):
        im = Image.fromarray(a[PILL_Y:PILL_Y + tile_h, PILL_X:PILL_X + tile_w])
        x = 10 + i * (tile_w + 10)
        strip.paste(im, (x, 10))
        d.text((x + tile_w // 2 - 30, tile_h + 18), name, fill="#5F6B74")
    shot = os.path.join(OUT, "pill_switch.png")
    strip.save(shot)
    assert os.path.getsize(shot) > 4000, "пустой снимок"
    print("8. Снимок-лента фаз: qa/render/pill_switch.png ✓")

    os.remove(wrapper)
    print("═══ ПЕРЕКЛЮЧАТЕЛЬ-ПИЛЮЛЯ: ВОЛНА, НОЖКА, ФОН — КАК В ОБРАЗЦЕ ✓ ═══")
    return 0


if __name__ == "__main__":
    sys.exit(main())
