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

# ВАЖНО: на сборочной машине Windows консоль может быть в cp1252,
# которая не умеет кириллицу — принудительно переводим вывод в UTF-8
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parent.parent

# Что кладём внутрь exe. Папки берутся целиком (кроме мусора),
# отдельные файлы — поимённо.
DIRS = ["components", "icons", "fonts", "shadows"]
ROOT_FILES = ["main.qml", "Template.xlsx", "app_icon.png"]

# Какие расширения из папок забираем (всё остальное — мимо кассы:
# скрипты и заготовки для ИИ в exe не попадают)
INCLUDE_EXT = {".qml", ".svg", ".png", ".jpg", ".ttf", ".otf"}

# Шрифты, которые программа реально открывает (см. AppTheme.qml).
# Остальные файлы в fonts/ — запасные, в exe их не кладём:
# один только GoogleSans.ttf весит почти 5 МБ.
FONT_ALLOW = {
    "Roboto-Regular.ttf",
    "Roboto-Medium.ttf",
    "Roboto-Bold.ttf",
    "RobotoCondensed-Regular.ttf",
    "RobotoCondensed-Bold.ttf",
}

def write_version_json():
    """Кладём version.json рядом с exe, чтобы обновлятор понял номер сборки."""
    theme = ROOT / "components" / "AppTheme.qml"
    text = theme.read_text(encoding="utf-8") if theme.exists() else ""
    m = __import__("re").search(r'appVersion:\s*"([^"]+)"', text)
    version = m.group(1) if m else "dev"
    out = ROOT / "version.json"
    out.write_text(
        '{\n  "name": "OVERTIMETAB",\n  "version": "%s"\n}\n' % version,
        encoding="utf-8",
    )
    print(f"version.json: {version}")
    return version


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
            if not p.is_file():
                continue
            if p.suffix.lower() in {".ttf", ".otf"} and p.name not in FONT_ALLOW:
                print(f"пропуск шрифта (программа его не открывает): {p.name}")
                continue
            if p.suffix.lower() in INCLUDE_EXT or p.name == "qmldir":
                files.append(p)
    return files

def main():
    write_version_json()
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

    # Ищем pyside6-rcc в нескольких местах (Windows/Linux совместимость)
    import shutil
    candidates = []
    found = shutil.which("pyside6-rcc")
    if found:
        candidates.append(found)
    exe_dir = Path(sys.executable).parent
    for name in ("pyside6-rcc", "pyside6-rcc.exe"):
        candidates.append(str(exe_dir / name))
        candidates.append(str(exe_dir / "Scripts" / name))
    rcc = next((c for c in candidates if Path(c).exists()), None)
    if not rcc:
        print("ERROR: pyside6-rcc not found. Run: pip install PySide6")
        sys.exit(1)

    out = ROOT / "resources_rc.py"
    result = subprocess.run([rcc, "-o", str(out), str(qrc_path)])
    if result.returncode != 0:
        print("ОШИБКА компиляции ресурсов")
        sys.exit(1)
    print(f"resources_rc.py: {out.stat().st_size / 1024:.0f} КБ — готово")

if __name__ == "__main__":
    main()
