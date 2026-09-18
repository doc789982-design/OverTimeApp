# -*- coding: utf-8 -*-
"""
Генератор готовых картинок теней — НЕОМОРФИЗМ (soft UI).

Вместо того чтобы заставлять видеокарту ВЫЧИСЛЯТЬ размытие тени
в реальном времени (DropShadow), мы один раз рисуем тень здесь,
сохраняем в PNG, а программа просто "приклеивает" картинку.
Стоимость — как у обычной картинки, то есть почти ноль.

Рецепт неоморфизма:
  * ВЫПУКЛЫЙ элемент = две тени сразу в одном PNG:
      светлая (белая) сверху-слева, тёмная (серо-синяя) снизу-справа.
  * ВДАВЛЕННЫЙ элемент (поля ввода, нажатые кнопки) = рамка
      с точным наоборот: тёмная кромка сверху-слева, светлая снизу-справа.
  * Тёмная тема считается по другому рецепту: чёрные тени +
      едва заметный светлый контур (белые тени на тёмном не видны),
      поэтому для каждой картинки два файла: обычный и _dark.

Запуск:  python tools/generate_shadows.py
Результат: папка shadows/:
    shadow_l1..l5.png      — выпуклые тени, светлая тема
    shadow_l1..l5_dark.png — выпуклые тени, тёмная тема
    inset_l1,l2.png        — вдавленные рамки, светлая тема
    inset_l1,l2_dark.png   — вдавленные рамки, тёмная тема
    shadow_knob.png        — круглая тень для кружка (переключатели)
    _preview.png           — контрольный лист для просмотра глазами

Чистый Python, никаких библиотек не нужно.
"""
import os
import struct
import zlib

# ════════════════════════════════════════════════════════════════
# ГЕОМЕТРИЯ. Выпуклые уровни 1..5 — те же размеры, что были раньше
# (pad и рамка BorderImage в AppShadow.qml остаются без изменений).
# ════════════════════════════════════════════════════════════════
LEVELS = {
    1: (5, 8),
    2: (8, 8),
    3: (11, 8),
    4: (16, 16),
    5: (20, 8),
}

# Вдавленные рамки: (размытие, скругление). 1 — поля ввода и мелочи,
# 2 — крупные вдавленные зоны.
INSET_LEVELS = {
    1: (4, 8),
    2: (7, 12),
}

SS = 2  # суперсэмплинг (рисуем в 2 раза крупнее, потом уменьшаем — гладкие края)

# ════════════════════════════════════════════════════════════════
# ЦВЕТА РЕЦЕПТА
#   light/dark = (r, g, b, сила) для светлого и тёмного мазка.
#   Для выпуклых теней сила тёмного мазка растёт с уровнем:
#   мелочь — деликатно, модальные окна — плотно.
# ════════════════════════════════════════════════════════════════
LIGHT_THEME = {
    "hi":   (255, 255, 255, 0.95),   # светлая тень (у кромки элемента)
    "lo":   (163, 177, 200, None),   # тёмная тень (сила берётся из карты ниже)
    "lo_scale": {1: 0.62, 2: 0.68, 3: 0.72, 4: 0.78, 5: 0.82},
}

DARK_THEME = {
    "hi":   (255, 255, 255, 0.08),   # лишь намёк на светлый контур
    "lo":   (7, 10, 14, None),       # почти чёрная тень
    "lo_scale": {1: 0.74, 2: 0.82, 3: 0.88, 4: 0.94, 5: 1.00},
}

INSET_LIGHT = {"lo": (163, 177, 200, 0.65), "hi": (255, 255, 255, 0.85)}
INSET_DARK  = {"lo": (5, 8, 12, 0.85),      "hi": (255, 255, 255, 0.07)}


# ════════════════════════════════════════════════════════════════
# Базовая математика (та же, что была в первой версии скрипта)
# ════════════════════════════════════════════════════════════════
def rounded_rect_alpha(w, h, r):
    """Маска прямоугольника со скруглёнными углами: 1.0 внутри, 0.0 снаружи."""
    a = [[0.0] * w for _ in range(h)]
    for y in range(h):
        for x in range(w):
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


def write_png_alpha(path, alpha):
    """Чёрный PNG с картой прозрачности (для shadow_knob — как раньше)."""
    h = len(alpha)
    w = len(alpha[0])
    write_png_rgba(path, [[(0, 0, 0, alpha[y][x]) for x in range(w)] for y in range(h)])


# ════════════════════════════════════════════════════════════════
# ВЫПУКЛАЯ ТЕНЬ: один размытый контур + два сдвига (свет и тьма)
# ════════════════════════════════════════════════════════════════
def make_raised(level, blur_r, corner_r, out_dir, theme, suffix):
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

    # Сдвиг светлого и тёмного мазков (в финальных пикселях)...
    shift = max(1, int(round(blur_r * 0.5))) * s
    layer_hi = shift_layer(ring, -shift, -shift)   # свет — влево-вверх
    layer_lo = shift_layer(ring, shift, shift)     # тень — вправо-вниз
    # ...и только теперь чистим центр фигуры: элемент закрывает его собой,
    # но если родитель прозрачен, грязное пятно под ним никому не нужно
    layer_hi = [[layer_hi[y][x] * (1.0 - big[y][x]) for x in range(full)] for y in range(full)]
    layer_lo = [[layer_lo[y][x] * (1.0 - big[y][x]) for x in range(full)] for y in range(full)]

    hi_r, hi_g, hi_b, hi_a = theme["hi"]
    lo_r, lo_g, lo_b = theme["lo"][0], theme["lo"][1], theme["lo"][2]
    lo_a = theme["lo_scale"].get(level, 0.7)

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

    path = os.path.join(out_dir, f"shadow_l{level}{suffix}.png")
    write_png_rgba(path, small)
    return path, size, pad + corner_r


# ════════════════════════════════════════════════════════════════
# ВДАВЛЕННАЯ РАМКА: гало от краёв ВНУТРЬ фигуры,
# тёмная кромка сверху-слева, светлая — снизу-справа
# ════════════════════════════════════════════════════════════════
def make_inset(level, blur_r, corner_r, out_dir, colors, suffix):
    pad_in = 2                       # фигура чуть отступает от края файла
    border = pad_in + blur_r * 2     # ширина зоны гало (рамка BorderImage)
    center = 8
    size = 2 * (border + corner_r) + center

    s = SS
    full = size * s
    # Фигура: скруглённый прямоугольник с отступом pad_in от краёв
    shape = rounded_rect_alpha(full - 2 * pad_in * s, full - 2 * pad_in * s, corner_r * s)
    mask = [[0.0] * full for _ in range(full)]
    off = pad_in * s
    for y in range(full - 2 * off):
        row_src = shape[y]
        row_dst = mask[y + off]
        for x in range(full - 2 * off):
            row_dst[x + off] = row_src[x]

    # Инвертированная маска, размытая и обрезанная обратно внутрь фигуры
    inv = [[1.0 - mask[y][x] for x in range(full)] for y in range(full)]
    halo = blur(inv, max(1, int(round(blur_r * 0.85 * s))))
    # Нормализация по кромке (см. выпуклые тени) и обрезка внутрь фигуры
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
                    ny = (y + dy) / (full - 1.0)   # 0 — верх, 1 — низ
                    nx = (x + dx) / (full - 1.0)   # 0 — лево, 1 — право
                    wlo_sum += max(1.0 - ny, 1.0 - nx) ** 1.5
                    whi_sum += max(ny, nx) ** 1.5
            h_val = h_sum / n
            if h_val <= 0.004:
                row.append((0.0, 0.0, 0.0, 0.0))
                continue
            w_lo = wlo_sum / n                     # верхне-левая кромка
            w_hi = whi_sum / n                     # нижне-правая кромка
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


# ════════════════════════════════════════════════════════════════
# Круглая тень для кружка переключателя (без изменений, как раньше)
# ════════════════════════════════════════════════════════════════
def make_knob(out_dir):
    blur_r = 5
    pad = blur_r * 2 + 2
    d = 24
    size = d + 2 * pad
    s = SS
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
    write_png_alpha(path, small)
    print(f"  {path}")


# ════════════════════════════════════════════════════════════════
# КОНТРОЛЬНЫЙ ЛИСТ: все тени на своих фонах — посмотреть глазами
# ════════════════════════════════════════════════════════════════
def make_preview(out_dir):
    def load_rgba(path):
        with open(path, "rb") as f:
            data = f.read()
        # Минимальный PNG-ридер (наши файлы: 8-bit RGBA, без interlace)
        pos = 8
        w = h = 0
        idat = b""
        while pos < len(data):
            length = struct.unpack(">I", data[pos:pos + 4])[0]
            tag = data[pos + 4:pos + 8]
            body = data[pos + 8:pos + 8 + length]
            if tag == b"IHDR":
                w, h = struct.unpack(">II", body[:8])
            elif tag == b"IDAT":
                idat += body
            pos += 12 + length
        raw = zlib.decompress(idat)
        px = []
        p = 0
        prev = [0] * (w * 4)
        for y in range(h):
            ftype = raw[p]
            p += 1
            line = list(raw[p:p + w * 4])
            p += w * 4
            # фильтры 0..4
            for i in range(w * 4):
                a = line[i - 4] if i >= 4 else 0
                b = prev[i]
                c = prev[i - 4] if i >= 4 else 0
                if ftype == 1:
                    line[i] = (line[i] + a) & 0xFF
                elif ftype == 2:
                    line[i] = (line[i] + b) & 0xFF
                elif ftype == 3:
                    line[i] = (line[i] + (a + b) // 2) & 0xFF
                elif ftype == 4:
                    pp = a + b - c
                    pa, pb, pc = abs(pp - a), abs(pp - b), abs(pp - c)
                    pr = a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)
                    line[i] = (line[i] + pr) & 0xFF
            prev = line
            row = []
            for x in range(w):
                i = x * 4
                row.append((line[i], line[i + 1], line[i + 2], line[i + 3]))
            px.append(row)
        return px, w, h

    def paste(canvas, img, cx, cy):
        px, w, h = img
        for yy in range(h):
            ty = cy - h // 2 + yy
            if not (0 <= ty < len(canvas)):
                continue
            for xx in range(w):
                tx = cx - w // 2 + xx
                if not (0 <= tx < len(canvas[0])):
                    continue
                r, g, b, a = px[yy][xx]
                af = a / 255.0
                cr, cg, cb = canvas[ty][tx]
                canvas[ty][tx] = (
                    int(r * af + cr * (1 - af)),
                    int(g * af + cg * (1 - af)),
                    int(b * af + cb * (1 - af)),
                )

    def fill_round(canvas, x0, y0, w, h, r, color):
        mask = rounded_rect_alpha(w, h, r)
        for y in range(h):
            for x in range(w):
                if mask[y][x] > 0.5:
                    canvas[y0 + y][x0 + x] = color

    W, H = 940, 560
    light_bg = (233, 237, 243)   # #E9EDF3
    dark_bg = (39, 44, 52)       # #272C34
    canvas = [[light_bg if x < W // 2 else dark_bg for x in range(W)] for y in range(H)]

    # Ряд 1: выпуклые L1..L5 (элемент 64×64, скругление 16)
    for i, lvl in enumerate([1, 2, 3, 4, 5]):
        cx = 90 + i * 80
        for half, suffix, bg in ((0, "", light_bg), (1, "_dark", dark_bg)):
            base = half * (W // 2)
            img = load_rgba(os.path.join(out_dir, f"shadow_l{lvl}{suffix}.png"))
            paste(canvas, img, base + cx, 90)
            fill_round(canvas, base + cx - 32, 90 - 32, 64, 64, 16, bg)

    # Ряд 2: вдавленные рамки L1/L2 на плашках 120×56 и 160×56
    for i, lvl in enumerate([1, 2]):
        cx = 90 + i * 140
        w_pl = 120 + i * 40
        for half, suffix, bg in ((0, "", light_bg), (1, "_dark", dark_bg)):
            base = half * (W // 2)
            img = load_rgba(os.path.join(out_dir, f"inset_l{lvl}{suffix}.png"))
            paste(canvas, img, base + cx, 240)
            fill_round(canvas, base + cx - w_pl // 2, 240 - 28, w_pl, 56, 16, bg)

    # Ряд 3: «кнопка» — выпуклая L1 + вдавленная L1 рядом (эффект нажатия)
    for half, suffix, bg in ((0, "", light_bg), (1, "_dark", dark_bg)):
        base = half * (W // 2)
        raised = load_rgba(os.path.join(out_dir, f"shadow_l1{suffix}.png"))
        pressed = load_rgba(os.path.join(out_dir, f"inset_l1{suffix}.png"))
        paste(canvas, raised, base + 100, 420)
        fill_round(canvas, base + 100 - 45, 420 - 16, 90, 32, 16, bg)
        paste(canvas, pressed, base + 260, 420)
        fill_round(canvas, base + 260 - 45, 420 - 16, 90, 32, 16, bg)

    write_png_rgba(os.path.join(out_dir, "_preview.png"),
                   [[(c[0], c[1], c[2], 1.0) for c in row] for row in canvas])
    print("  _preview.png (контрольный лист)")


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = os.path.join(root, "shadows")
    os.makedirs(out_dir, exist_ok=True)
    print("Генерирую неоморфные тени:")
    for level, (blur_r, corner_r) in LEVELS.items():
        p, size, border = make_raised(level, blur_r, corner_r, out_dir, LIGHT_THEME, "")
        print(f"  {p}  {size}x{size}px  (рамка {border}px)")
        p, _, _ = make_raised(level, blur_r, corner_r, out_dir, DARK_THEME, "_dark")
        print(f"  {p}")
    for level, (blur_r, corner_r) in INSET_LEVELS.items():
        p, size, border = make_inset(level, blur_r, corner_r, out_dir, INSET_LIGHT, "")
        print(f"  {p}  {size}x{size}px  (рамка {border}px)")
        p, _, _ = make_inset(level, blur_r, corner_r, out_dir, INSET_DARK, "_dark")
        print(f"  {p}")
    make_knob(out_dir)
    make_preview(out_dir)
    print("Готово.")


if __name__ == "__main__":
    main()
