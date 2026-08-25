# -*- coding: utf-8 -*-
"""
Конвертирует app_icon.png в app_icon.ico (все нужные размеры одним файлом).
PyInstaller для иконки exe на Windows хочет именно .ico.

Запуск:  python tools/make_icon.py   (нужен пакет Pillow)
"""
import sys
from pathlib import Path
from PIL import Image

# Кириллица в консоли сборочной машины (cp1252) — принудительный UTF-8
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent

def main():
    src = ROOT / "app_icon.png"
    dst = ROOT / "app_icon.ico"
    img = Image.open(src).convert("RGBA")
    # Приводим к квадрату (на всякий случай)
    side = max(img.size)
    square = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    square.paste(img, ((side - img.width) // 2, (side - img.height) // 2))
    square.save(dst, sizes=[(16, 16), (24, 24), (32, 32), (48, 48),
                            (64, 64), (128, 128), (256, 256)])
    print(f"app_icon.ico готов ({dst.stat().st_size / 1024:.0f} КБ)")

if __name__ == "__main__":
    main()
