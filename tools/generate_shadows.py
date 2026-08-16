# -*- coding: utf-8 -*-
"""
Генератор готовых картинок теней (как в Telegram/Discord).

Вместо того чтобы заставлять видеокарту ВЫЧИСЛЯТЬ размытие тени
в реальном времени (DropShadow), мы один раз рисуем тень здесь,
сохраняем в PNG, а программа просто "приклеивает" картинку.

Запуск:  python tools/generate_shadows.py
Результат: папка shadows/ с файлами shadow_l1.png ... shadow_l5.png
           и shadow_knob.png (для кружка переключателя).

Чистый Python, никаких библиотек не нужно.
"""
import os
import struct
import zlib

# Уровни теней — те же цифры, что были в AppTheme.qml:
#   (размытие, скругление углов элемента)
LEVELS = {
    1: (3, 8),
    2: (6, 8),
    3: (8, 8),
    4: (12, 16),
    5: (16, 8),
}

SS = 2  # суперсэмплинг (рисуем в 2 раза крупнее, потом уменьшаем — гладкие края)


def rounded_rect_alpha(w, h, r):
    """Маска прямоугольника со скруглёнными углами: 1.0 внутри, 0.0 снаружи."""
    a = [[0.0] * w for _ in range(h)]
    for y in range(h):
        for x in range(w):
            # расстояние до "внутреннего" прямоугольника
            dx = max(r - x, x - (w - 1 - r), 0)
            dy = max(r - y, y - (h - 1 - r), 0)
            if dx * dx + dy * dy <= r * r:
                a[y][x] = 1.0
    return a


def box_blur_h(src, radius):
    h = len(src)
    w = len(src[0])
    out = [[0.0] * w for _ in range(h)]
    div = 2 * radius + 1
    for y in range(h):
        row = src[y]
        s = sum(row[0] if i < 0 else row[min(i, w - 1)] for i in range(-radius, radius + 1))
        for x in range(w):
            out[y][x] = s / div
            i_add = min(x + radius + 1, w - 1)
            i_sub = max(x - radius, 0)
            s += row[i_add] - row[i_sub]
    return out


def box_blur_v(src, radius):
    h = len(src)
    w = len(src[0])
    out = [[0.0] * w for _ in range(h)]
    div = 2 * radius + 1
    for x in range(w):
        s = sum(src[0][x] if i < 0 else src[min(i, h - 1)][x] for i in range(-radius, radius + 1))
        for y in range(h):
            out[y][x] = s / div
            i_add = min(y + radius + 1, h - 1)
            i_sub = max(y - radius, 0)
            s += src[i_add][x] - src[i_sub][x]
    return out


def blur(src, radius):
    """Три коробочных размытия подряд ≈ гауссово (так делают все)."""
    r = max(1, radius)
    for _ in range(3):
        src = box_blur_v(box_blur_h(src, r), r)
    return src


def downsample(src, factor):
    h = len(src) // factor
    w = len(src[0]) // factor
    out = [[0.0] * w for _ in range(h)]
    f2 = factor * factor
    for y in range(h):
        for x in range(w):
            s = 0.0
            for dy in range(factor):
                for dx in range(factor):
                    s += src[y * factor + dy][x * factor + dx]
            out[y][x] = s / f2
    return out


def write_png(path, alpha):
    """Сохраняем чёрный PNG с картой прозрачности."""
    h = len(alpha)
    w = len(alpha[0])
    raw = b""
    for y in range(h):
        raw += b"\x00"  # фильтр строки
        for x in range(w):
            a = max(0, min(255, int(round(alpha[y][x] * 255))))
            raw += bytes((0, 0, 0, a))  # чёрный + альфа

    def chunk(tag, data):
        c = struct.pack(">I", len(data)) + tag + data
        c += struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        return c

    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))
    png += chunk(b"IDAT", zlib.compress(raw, 9))
    png += chunk(b"IEND", b"")
    with open(path, "wb") as f:
        f.write(png)


def make_level(level, blur_r, corner_r, out_dir):
    # Отступ вокруг фигуры, чтобы тени было куда "расплыться"
    pad = blur_r * 2 + 2
    # Центральная растягиваемая зона
    center = 8
    size = 2 * (pad + corner_r) + center

    s = SS
    # ВАЖНО: фигура рисуется В ЦЕНТРЕ картинки с отступом pad со всех
    # сторон — именно в этот отступ и "расплывается" размытие.
    # (Раньше фигура занимала весь файл и тень выглядела как плита.)
    full = size * s
    inner_w = full - 2 * pad * s
    shape = rounded_rect_alpha(inner_w, inner_w, corner_r * s)
    big = [[0.0] * full for _ in range(full)]
    off = pad * s
    for y in range(inner_w):
        row_src = shape[y]
        row_dst = big[y + off]
        for x in range(inner_w):
            row_dst[x + off] = row_src[x]

    big = blur(big, max(1, int(round(blur_r * 0.55 * s))))
    small = downsample(big, s)

    path = os.path.join(out_dir, f"shadow_l{level}.png")
    write_png(path, small)
    print(f"  {path}  {size}x{size}px  (граница для BorderImage: {pad + corner_r}px)")
    return pad, corner_r


def make_knob(out_dir):
    """Круглая тень для кружка переключателя (AppSwitch)."""
    blur_r = 5
    pad = blur_r * 2 + 2
    d = 24  # диаметр кружка
    size = d + 2 * pad
    s = SS
    big = rounded_rect_alpha(size * s, size * s, (d // 2 + pad) * s)  # почти круг
    # вырезаем именно круг: скругление = половина стороны
    big = rounded_rect_alpha(size * s, size * s, size * s // 2)
    # но нам нужен круг диаметром d в центре: перерисуем аккуратно
    big = [[0.0] * (size * s) for _ in range(size * s)]
    cx = cy = size * s / 2.0
    rr = d * s / 2.0
    for y in range(size * s):
        for x in range(size * s):
            dx = x + 0.5 - cx
            dy = y + 0.5 - cy
            if dx * dx + dy * dy <= rr * rr:
                big[y][x] = 1.0
    big = blur(big, max(1, int(round(blur_r * 0.55 * s))))
    small = downsample(big, s)
    path = os.path.join(out_dir, "shadow_knob.png")
    write_png(path, small)
    print(f"  {path}  {size}x{size}px")


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = os.path.join(root, "shadows")
    os.makedirs(out_dir, exist_ok=True)
    print("Генерирую тени:")
    for level, (blur_r, corner_r) in LEVELS.items():
        make_level(level, blur_r, corner_r, out_dir)
    make_knob(out_dir)
    print("Готово.")


if __name__ == "__main__":
    main()
