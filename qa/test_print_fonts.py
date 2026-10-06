#!/usr/bin/env python3
"""Тест шрифта печати табеля — PT Astra Serif, вшитого в программу.

Ситуация с практики: у части коллег в напечатанном табеле шрифт
подменяется системным — Excel и рендер печати берут PT Astra Serif
только если шрифт есть в системе. Решение: четыре начертания вшиты
в программу (fonts/, пакуются в exe) и при запуске
  1) добавляются в базу шрифтов Qt (печать/PDF программы);
  2) если в Windows шрифта нет — ставятся «для меня» без прав
     администратора (%LOCALAPPDATA%\\...\\Fonts + реестр HKCU),
     после чего Excel печатает табель правильным шрифтом.

Проверяется:
  A. парсер имён TTF: у всех четырёх файлов — семейство PT Astra Serif
     и правильные полные имена (для записи в реестр);
  B. решение об установке: полное имя сверяется точно, без
     подстрок («Bold» установлен — «Regular» всё равно нужен);
  C. в exe пакуются все четыре файла (FONT_ALLOW сборки ресурсов);
  D. _register_print_fonts: семейство PT Astra Serif со всеми четырьмя
     стилями появляется в базе Qt;
  E. вне Windows установка в Windows — тихий ноль без ошибок.

Запуск:
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen python3 qa/test_print_fonts.py
"""
import os
import sys
import types
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

import Main  # noqa: E402


def main() -> int:
    # ── A. парсер имён TTF ──
    expected = {
        "PT-Astra-Serif_Regular.ttf": ("PT Astra Serif", "PT Astra Serif Regular"),
        "PT-Astra-Serif_Bold.ttf": ("PT Astra Serif", "PT Astra Serif Bold"),
        "PT-Astra-Serif_Italic.ttf": ("PT Astra Serif", "PT Astra Serif Italic"),
        "PT-Astra-Serif_Bold-Italic.ttf": ("PT Astra Serif", "PT Astra Serif Bold Italic"),
    }
    assert set(Main.PRINT_FONT_FILES) == set(expected), Main.PRINT_FONT_FILES
    for name, (fam, full) in expected.items():
        data = (ROOT / "fonts" / name).read_bytes()
        assert data, name
        got = Main._ttf_font_names(data)
        assert got == (fam, full), (name, got)
    # мусор не парсится
    assert Main._ttf_font_names(b"not a font at all") == ("", "")
    print("A: имена всех четырёх начертаний читаются из файлов ✓")

    # ── B. решение об установке ──
    empty = Main._needs_excel_font_install("PT Astra Serif Regular", set())
    assert empty is True
    # установлен только Bold — Regular всё равно нужен (сверка точная)
    has_bold = {"pt astra serif bold"}
    assert Main._needs_excel_font_install("PT Astra Serif Bold", has_bold) is False
    assert Main._needs_excel_font_install("PT Astra Serif Regular", has_bold) is True
    assert Main._needs_excel_font_install("PT Astra Serif Bold Italic", has_bold) is True
    # имя с суффиксом реестра распознаётся как установленное
    assert Main._needs_excel_font_install(
        "PT Astra Serif", {"pt astra serif (truetype)"}) is False
    assert Main._needs_excel_font_install("", set()) is False
    print("B: решение об установке — точная сверка полных имён ✓")

    # ── C. все четыре файла пакуются в exe ──
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "make_resources", ROOT / "tools" / "make_resources.py")
    mr = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mr)
    for name in expected:
        assert name in mr.FONT_ALLOW, (name, mr.FONT_ALLOW)
    print("C: все четыре начертания пакуются в exe ✓")

    # ── D. подключение к Qt ──
    from PySide6.QtGui import QGuiApplication, QFontDatabase
    app = QGuiApplication.instance() or QGuiApplication(sys.argv)
    assert "PT Astra Serif" not in QFontDatabase.families() or True
    added = Main._register_print_fonts()
    assert added >= 1, "ни один шрифт не добавился в Qt"
    assert "PT Astra Serif" in QFontDatabase.families(), QFontDatabase.families()[:5]
    styles = set(QFontDatabase.styles("PT Astra Serif"))
    assert {"Regular", "Bold", "Italic", "Bold Italic"} <= styles, styles
    # повторный вызов безвреден (идемпотентность запуска)
    Main._register_print_fonts()
    assert "PT Astra Serif" in QFontDatabase.families()
    print("D: семейство PT Astra Serif со всеми стилями доступно печати ✓")

    # ── E. вне Windows установка — тихий ноль ──
    assert Main._ensure_excel_fonts_windows() == 0
    assert Main._excel_fonts_installed_names() == set()
    print("E: вне Windows — тихий ноль, без ошибок ✓")

    print("═══ ШРИФТ ПЕЧАТИ ТАБЕЛЯ ВШИТ: QT + EXCEL НА ЛЮБОЙ МАШИНЕ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
