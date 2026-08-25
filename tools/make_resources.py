# -*- coding: utf-8 -*-
"""
Генератор ресурсов для сборки в .exe.

Собирает все нужные файлы (интерфейс, иконки, шрифты, тени, шаблон Excel)
в resources.qrc и компилирует их в resources_rc.py через pyside6-rcc.
После этого PyInstaller кладёт resources_rc.py внутрь exe, и программа
работает из ОДНОГО файла — вокруг exe не должно быть никаких папок.

Запуск:  python tools/make_resources.py   (из корня репозитория)
"""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# Что кладём внутрь exe. Папки берутся целиком (кроме мусора),
# отдельные файлы — поимённо.
DIRS = ["components", "icons", "fonts", "shadows"]
ROOT_FILES = ["main.qml", "Template.xlsx", "app_icon.png"]

# Какие расширения из папок забираем (всё остальное — мимо кассы:
# скрипты и заготовки для ИИ в exe не попадают)
INCLUDE_EXT = {".qml", ".svg", ".png", ".jpg", ".ttf", ".otf"}

def collect_files():
    files = []
    for name in ROOT_FILES:
        p = ROOT / name
        if p.exists():
            files.append(p)
        else:
            print(f"ВНИМАНИЕ: не найден {name}")
    for d in DIRS:
        base = ROOT / d
        if not base.exists():
            continue
        for p in sorted(base.rglob("*")):
            if p.is_file() and (p.suffix.lower() in INCLUDE_EXT or p.name == "qmldir"):
                files.append(p)
    return files

def main():
    files = collect_files()
    if not files:
        print("ОШИБКА: не найдено ни одного файла ресурсов")
        sys.exit(1)

    total_kb = sum(f.stat().st_size for f in files) / 1024
    lines = ["<RCC>", '  <qresource prefix="/">']
    for f in files:
        rel = f.relative_to(ROOT).as_posix()
        lines.append(f"    <file>{rel}</file>")
    lines += ["  </qresource>", "</RCC>"]

    qrc_path = ROOT / "resources.qrc"
    qrc_path.write_text("\n".join(lines), encoding="utf-8")
    print(f"resources.qrc: {len(files)} файлов, {total_kb:.0f} КБ")

    # Ищем pyside6-rcc (ставится вместе с PySide6)
    import shutil
    rcc = shutil.which("pyside6-rcc")
    if rcc is None:
        # запасной вариант: лежит рядом с питоном
        candidate = Path(sys.executable).parent / "pyside6-rcc"
        if candidate.exists():
            rcc = str(candidate)
        else:
            print("ОШИБКА: не найден pyside6-rcc. Установите: pip install PySide6")
            sys.exit(1)

    out = ROOT / "resources_rc.py"
    result = subprocess.run([rcc, "-o", str(out), str(qrc_path)])
    if result.returncode != 0:
        print("ОШИБКА компиляции ресурсов")
        sys.exit(1)
    print(f"resources_rc.py: {out.stat().st_size / 1024:.0f} КБ — готово")

if __name__ == "__main__":
    main()
