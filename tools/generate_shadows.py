# -*- coding: utf-8 -*-
"""
Генератор готовых картинок теней (как в Telegram/Discord).

Вместо того чтобы заставлять видеокарту ВЫЧИСЛЯТЬ размытие тени
в реальном времени (DropShadow), мы один раз рисуем тень здесь,
сохраняем в PNG, а программа просто "приклеивает" картинку.

Запуск:  python tools/generate_shadows.py
Результат: папка shadows/ с файлами shadow_l1.png ... shadow_l5.png
           и shadow_knob.png (для кружка переключателя),
           плюс неоморфный набор: soft_l1/l2, inset_l1, soft_knob
           (и их _dark-версии) для точечных «выпуклых» элементов.

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



# ════════════════════════════════════════════════════════════════
# НЕОМОРФНАЯ СЕКЦИЯ (soft UI) — точечные элементы интерфейса.
# Классические shadow_l*.png выше НЕ меняются: диалоги, меню и
# всплывающие окна сохраняют прежний вид. Здесь же — отдельный
# набор картинок для «выпуклых» и «вдавленных» деталей:
#   soft_l1/l2      — выпуклые тени (карточка сотрудника, плашки
#                     иконок, крупные панели)
#   inset_l1        — вдавленная рамка (чекбоксы)
#   soft_knob       — круглая выпуклая тень для кружка переключателя
# ════════════════════════════════════════════════════════════════

SOFT_LEVELS = {
    1: (5, 12),   # мелкие элементы: карточки списка, плашки иконок
    2: (8, 12),   # крупные панели
}

SOFT_LIGHT = {
    "hi": (255, 255, 255, 0.95),   # светлая тень (сверху-слева)
    "lo": (163, 177, 200, None),   # тёмная тень (снизу-справа)
    "lo_scale": {1: 0.62, 2: 0.68},
}

SOFT_DARK = {
    "hi": (255, 255, 255, 0.08),   # лишь намёк на светлый контур
    "lo": (7, 10, 14, None),       # почти чёрная тень
    "lo_scale": {1: 0.74, 2: 0.82},
}

INSET_LEVELS = {
    1: (3, 8),     # чекбоксы и небольшие «ямки»
}

INSET_LIGHT = {"lo": (163, 177, 200, 0.65), "hi": (255, 255, 255, 0.85)}
INSET_DARK  = {"lo": (5, 8, 12, 0.85),      "hi": (255, 255, 255, 0.07)}


def write_png_rgba(path, pixels):
    """pixels[y][x] = (r, g, b, a_0_1). Сохраняем как RGBA PNG."""
    h = len(pixels)
    w = len(pixels[0])
    raw = b""
    for y in range(h):
        raw += b"\x00"  # фильтр строки
        for x in range(w):
            r, g, b, a = pixels[y][x]
            raw += bytes((int(r), int(g), int(b), max(0, min(255, int(round(a * 255))))))

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


def shift_layer(src, dx, dy):
    """Сдвиг карты: положительное — вправо-вниз, отрицательное — влево-вверх."""
    h = len(src)
    w = len(src[0])
    out = [[0.0] * w for _ in range(h)]
    for y in range(h):
        sy = y - dy
        if 0 <= sy < h:
            row_o = out[y]
            row_s = src[sy]
            for x in range(w):
                sx = x - dx
                if 0 <= sx < w:
                    row_o[x] = row_s[sx]
    return out


def make_soft(level, blur_r, corner_r, out_dir, theme, suffix):
    """Выпуклая неоморфная тень: один контур, два сдвига (свет и тьма)."""
    pad = blur_r * 2 + 2
    center = 8
    size = 2 * (pad + corner_r) + center

    s = SS
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

    ring = blur(big, max(1, int(round(blur_r * 0.55 * s))))
    # Нормализация: у прямой кромки размытие даёт ~0.5 — поднимаем до 1.0
    ring = [[min(1.0, ring[y][x] * 2.0) for x in range(full)] for y in range(full)]

    shift = max(1, int(round(blur_r * 0.5))) * s
    layer_hi = shift_layer(ring, -shift, -shift)   # свет — влево-вверх
    layer_lo = shift_layer(ring, shift, shift)     # тень — вправо-вниз
    # Чистим центр фигуры ПОСЛЕ сдвигов: элемент закрывает его собой
    layer_hi = [[layer_hi[y][x] * (1.0 - big[y][x]) for x in range(full)] for y in range(full)]
    layer_lo = [[layer_lo[y][x] * (1.0 - big[y][x]) for x in range(full)] for y in range(full)]

    hi_r, hi_g, hi_b, hi_a = theme["hi"]
    lo_r, lo_g, lo_b = theme["lo"][0], theme["lo"][1], theme["lo"][2]
    lo_a = theme["lo_scale"].get(level, 0.65)

    small = []
    for y in range(0, full, s):
        row = []
        for x in range(0, full, s):
            la = 0.0
            da = 0.0
            for dy in range(s):
                for dx in range(s):
                    la += layer_hi[y + dy][x + dx]
                    da += layer_lo[y + dy][x + dx]
            la = (la / (s * s)) * hi_a
            da = (da / (s * s)) * lo_a
            a = da + la * (1.0 - da)
            if a <= 0.002:
                row.append((0.0, 0.0, 0.0, 0.0))
            else:
                r = (lo_r * da + hi_r * la * (1.0 - da)) / a
                g = (lo_g * da + hi_g * la * (1.0 - da)) / a
                b = (lo_b * da + hi_b * la * (1.0 - da)) / a
                row.append((r, g, b, a))
        small.append(row)

    path = os.path.join(out_dir, f"soft_l{level}{suffix}.png")
    write_png_rgba(path, small)
    return path, size, pad + corner_r


def make_inset_frame(level, blur_r, corner_r, out_dir, colors, suffix):
    """Вдавленная рамка: гало от кромок ВНУТРЬ фигуры,
    тёмная кромка сверху-слева, светлая — снизу-справа."""
    pad_in = 2
    border = pad_in + blur_r * 2
    center = 8
    size = 2 * (border + corner_r) + center

    s = SS
    full = size * s
    shape = rounded_rect_alpha(full - 2 * pad_in * s, full - 2 * pad_in * s, corner_r * s)
    mask = [[0.0] * full for _ in range(full)]
    off = pad_in * s
    for y in range(full - 2 * off):
        row_src = shape[y]
        row_dst = mask[y + off]
        for x in range(full - 2 * off):
            row_dst[x + off] = row_src[x]

    inv = [[1.0 - mask[y][x] for x in range(full)] for y in range(full)]
    halo = blur(inv, max(1, int(round(blur_r * 0.85 * s))))
    halo = [[min(1.0, halo[y][x] * 2.0) * mask[y][x] for x in range(full)]
            for y in range(full)]

    lo_r, lo_g, lo_b, lo_a = colors["lo"]
    hi_r, hi_g, hi_b, hi_a = colors["hi"]

    small = []
    for y in range(0, full, s):
        row = []
        for x in range(0, full, s):
            h_sum = 0.0
            wlo_sum = 0.0
            whi_sum = 0.0
            n = s * s
            for dy in range(s):
                for dx in range(s):
                    h_sum += halo[y + dy][x + dx]
                    ny = (y + dy) / (full - 1.0)
                    nx = (x + dx) / (full - 1.0)
                    wlo_sum += max(1.0 - ny, 1.0 - nx) ** 1.5
                    whi_sum += max(ny, nx) ** 1.5
            h_val = h_sum / n
            if h_val <= 0.004:
                row.append((0.0, 0.0, 0.0, 0.0))
                continue
            w_lo = wlo_sum / n
            w_hi = whi_sum / n
            da = h_val * lo_a * w_lo
            la = h_val * hi_a * w_hi
            a = da + la * (1.0 - da)
            if a <= 0.002:
                row.append((0.0, 0.0, 0.0, 0.0))
            else:
                r = (lo_r * da + hi_r * la * (1.0 - da)) / a
                g = (lo_g * da + hi_g * la * (1.0 - da)) / a
                b = (lo_b * da + hi_b * la * (1.0 - da)) / a
                row.append((r, g, b, a))
        small.append(row)

    path = os.path.join(out_dir, f"inset_l{level}{suffix}.png")
    write_png_rgba(path, small)
    return path, size, border


def make_soft_knob(out_dir, theme):
    """Круглая выпуклая тень для кружка переключателя (двухцветная)."""
    d = 20            # диаметр кружка
    pad = 10          # запас на размытие
    blur_r = 4
    size = d + 2 * pad
    s = SS
    full = size * s
    big = [[0.0] * full for _ in range(full)]
    cx = cy = full / 2.0
    rr = d * s / 2.0
    for y in range(full):
        for x in range(full):
            ddx = x + 0.5 - cx
            ddy = y + 0.5 - cy
            if ddx * ddx + ddy * ddy <= rr * rr:
                big[y][x] = 1.0
    ring = blur(big, max(1, int(round(blur_r * 0.55 * s))))
    ring = [[min(1.0, ring[y][x] * 2.0) * (1.0 - big[y][x]) for x in range(full)]
            for y in range(full)]

    shift = max(1, int(round(blur_r * 0.5))) * s
    layer_hi = shift_layer(ring, -shift, -shift)
    layer_lo = shift_layer(ring, shift, shift)

    hi_r, hi_g, hi_b, hi_a = theme["hi"]
    lo_r, lo_g, lo_b = theme["lo"][0], theme["lo"][1], theme["lo"][2]
    lo_a = theme["lo_scale"][1]

    small = []
    for y in range(0, full, s):
        row = []
        for x in range(0, full, s):
            la = 0.0
            da = 0.0
            for dy in range(s):
                for dx in range(s):
                    la += layer_hi[y + dy][x + dx]
                    da += layer_lo[y + dy][x + dx]
            la = (la / (s * s)) * hi_a
            da = (da / (s * s)) * lo_a
            a = da + la * (1.0 - da)
            if a <= 0.002:
                row.append((0.0, 0.0, 0.0, 0.0))
            else:
                r = (lo_r * da + hi_r * la * (1.0 - da)) / a
                g = (lo_g * da + hi_g * la * (1.0 - da)) / a
                b = (lo_b * da + hi_b * la * (1.0 - da)) / a
                row.append((r, g, b, a))
        small.append(row)

    for suffix in ("", "_dark"):
        path = os.path.join(out_dir, f"soft_knob{suffix}.png")
        # тёмная версия — те же карты, другие цвета
        if suffix == "_dark":
            hi_r, hi_g, hi_b, hi_a = SOFT_DARK["hi"]
            lo_r, lo_g, lo_b = SOFT_DARK["lo"][0], SOFT_DARK["lo"][1], SOFT_DARK["lo"][2]
            lo_a = SOFT_DARK["lo_scale"][1]
            small = []
            for y in range(0, full, s):
                row = []
                for x in range(0, full, s):
                    la = 0.0
                    da = 0.0
                    for dy in range(s):
                        for dx in range(s):
                            la += layer_hi[y + dy][x + dx]
                            da += layer_lo[y + dy][x + dx]
                    la = (la / (s * s)) * hi_a
                    da = (da / (s * s)) * lo_a
                    a = da + la * (1.0 - da)
                    if a <= 0.002:
                        row.append((0.0, 0.0, 0.0, 0.0))
                    else:
                        r = (lo_r * da + hi_r * la * (1.0 - da)) / a
                        g = (lo_g * da + hi_g * la * (1.0 - da)) / a
                        b = (lo_b * da + hi_b * la * (1.0 - da)) / a
                        row.append((r, g, b, a))
                small.append(row)
        write_png_rgba(path, small)
        print(f"  {path}  {size}x{size}px (рамка {pad + d // 2}px)")


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = os.path.join(root, "shadows")
    os.makedirs(out_dir, exist_ok=True)
    print("Генерирую тени:")
    for level, (blur_r, corner_r) in LEVELS.items():
        make_level(level, blur_r, corner_r, out_dir)
    make_knob(out_dir)

    print("Неоморфные тени (точечные элементы):")
    for level, (blur_r, corner_r) in SOFT_LEVELS.items():
        p, size, border = make_soft(level, blur_r, corner_r, out_dir, SOFT_LIGHT, "")
        print(f"  {p}  {size}x{size}px  (рамка {border}px)")
        p, _, _ = make_soft(level, blur_r, corner_r, out_dir, SOFT_DARK, "_dark")
        print(f"  {p}")
    for level, (blur_r, corner_r) in INSET_LEVELS.items():
        p, size, border = make_inset_frame(level, blur_r, corner_r, out_dir, INSET_LIGHT, "")
        print(f"  {p}  {size}x{size}px  (рамка {border}px)")
        p, _, _ = make_inset_frame(level, blur_r, corner_r, out_dir, INSET_DARK, "_dark")
        print(f"  {p}")
    make_soft_knob(out_dir, SOFT_LIGHT)
    print("Готово.")


if __name__ == "__main__":
    main()
