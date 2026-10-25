#!/usr/bin/env python3
"""Проверка внешнего вида окна печати по пикселям (песочница).

Рендерит AppPrintDialog скриптом qa/render_print_dialog.py и проверяет
снимки: карточка по центру, цвета светлой/тёмной темы, акцентная кнопка,
затемнение правой колонки при выборе «Экспортировать в Excel»,
появление полей диапазона.

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \\
    QSG_RASTER_BACKEND=1 python3 qa/test_print_dialog_view.py
"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RENDER = os.path.join(ROOT, "qa", "render")


def run_render():
    env = dict(os.environ)
    env.setdefault("QT_QPA_PLATFORM", "offscreen")
    env.setdefault("QSG_RASTER_BACKEND", "1")
    r = subprocess.run([sys.executable, os.path.join(ROOT, "qa",
                        "render_print_dialog.py")],
                       capture_output=True, text=True, env=env, cwd=ROOT)
    assert r.returncode == 0, "рендер упал:\n" + r.stdout[-800:] + r.stderr[-800:]


def load(name):
    from PySide6.QtGui import QImage
    img = QImage(os.path.join(RENDER, "print_dialog_%s.png" % name))
    assert not img.isNull(), "нет снимка %s" % name
    return img


def near(c, target, tol=28):
    return all(abs(c[i] - target[i]) <= tol for i in range(3))


def ink(img, x0, y0, x1, y1, bg, tol=40):
    """Доля не-фоновых пикселей в прямоугольнике."""
    n = hit = 0
    for y in range(y0, y1, 2):
        for x in range(x0, x1, 2):
            c = (img.pixelColor(x, y).red(), img.pixelColor(x, y).green(),
                 img.pixelColor(x, y).blue())
            n += 1
            if not near(c, bg, tol):
                hit += 1
    return hit / max(1, n)


def modal_color(img, x0, y0, x1, y1):
    """Самый частый цвет области — цвет подложки без текста и границ."""
    counts = {}
    for y in range(y0, y1):
        for x in range(x0, x1):
            c = (img.pixelColor(x, y).red(), img.pixelColor(x, y).green(),
                 img.pixelColor(x, y).blue())
            counts[c] = counts.get(c, 0) + 1
    return max(counts, key=counts.get)


def has_accent(img, x0, y0, x1, y1):
    """Есть ли фирменный синий (кнопка «Отправить на печать»)."""
    for y in range(y0, y1, 3):
        for x in range(x0, x1, 3):
            c = img.pixelColor(x, y)
            if c.blue() > 140 and c.blue() > c.red() + 50 and c.green() > 80:
                return True
    return False


def main() -> int:
    run_render()
    light, dark = load("light"), load("dark")
    excel, rng = load("excel"), load("range")
    W, H = light.width(), light.height()
    assert (W, H) == (920, 660), "размер окна %dx%d" % (W, H)

    # окно (фон под модальным затемнением) и карточка — светлая тема.
    # Оверлей bgOverlay = rgba(32,33,36,0.4) поверх #E4E7EA → ≈ #96989A
    corner = light.pixelColor(24, 24)
    assert near((corner.red(), corner.green(), corner.blue()), (150, 152, 155)), \
        "фон окна под затемнением не тот: %s" % corner.name()
    card = modal_color(light, 140, 165, 300, 200)
    assert near(card, (255, 255, 255), 12), "карточка не белая: %s" % str(card)

    # заголовок карточки — есть чернила (текст «Настройки печати»)
    title_ink = ink(light, W // 2 - 150, 100, W // 2 + 150, 140, (255, 255, 255))
    assert title_ink > 0.01, "нет текста заголовка (%.3f)" % title_ink

    # акцентная кнопка в левой колонке
    assert has_accent(light, 130, 250, 460, 420), "нет синей кнопки печати"

    # диапазон: в правой колонке появились поля — чернил стало больше
    right = (480, 150, 790, 480)
    bg_card = (255, 255, 255)
    ink_light = ink(light, *right, bg_card)
    ink_range = ink(rng, *right, bg_card)
    assert ink_range > ink_light + 0.005, "поля диапазона не появились (%.3f → %.3f)" % (
        ink_light, ink_range)

    # Excel: правая колонка затемнена — чернил сильно меньше
    ink_excel = ink(excel, *right, bg_card)
    assert ink_excel < ink_light / 2, "правая колонка не затемнилась (%.3f)" % ink_excel

    # тёмная тема: оверлей чёрный 0.6 поверх #121212 → ≈ #070707,
    # карточка — bgModal #333333
    corner_d = dark.pixelColor(24, 24)
    assert near((corner_d.red(), corner_d.green(), corner_d.blue()), (7, 7, 7), 12), \
        "фон тёмного окна не затемнён: %s" % corner_d.name()
    card_d = modal_color(dark, 140, 165, 300, 200)
    assert near(card_d, (51, 51, 51), 12), "карточка не тёмная: %s" % str(card_d)
    assert has_accent(dark, 130, 250, 460, 420), "нет синей кнопки в тёмной теме"

    print("светлая: карточка/фон/заголовок/кнопка ✓")
    print("диапазон: поля появились (чернила %.3f → %.3f) ✓" % (ink_light, ink_range))
    print("Excel: правая колонка затемнена (%.3f) ✓" % ink_excel)
    print("тёмная: карточка %s на фоне %s, кнопка ✓" % (str(card_d), str((corner_d.red(), corner_d.green(), corner_d.blue()))))
    print("═══ ВНЕШНИЙ ВИД ОКНА ПЕЧАТИ: ВСЁ НА МЕСТЕ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
