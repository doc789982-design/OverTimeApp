#!/usr/bin/env python3
"""Тест подбора имени принтера (print_engine._match_printer_name).

Диалог печати получает имена из win32print, а QPrinter.setPrinterName
молча оставляет принтер ПО УМОЛЧАНИЮ, если имя не совпало со списком Qt
(регистр, пробелы, сетевой путь «\\сервер\принтер»). Поэтому перед
установкой имя подбирается по списку Qt: точное совпадение без учёта
регистра, затем по хвосту сетевого пути — в обе стороны.

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs python3 qa/test_printer_match.py
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

BS = chr(92)  # обратный слэш без экранирования в исходнике теста

UNC = BS * 2 + "SRV" + BS + "HP LaserJet 1020"

CASES = [
    # запрошено / доступно / ожидание
    ("HP LaserJet 1020", ["hp laserjet 1020 ", "Xerox"], "hp laserjet 1020 "),
    (UNC, ["HP LaserJet 1020", "X"], "HP LaserJet 1020"),
    ("HP LaserJet 1020", [UNC], UNC),
    ("Нет такого", ["HP", "Xerox"], ""),
    ("", ["HP"], ""),
]


def main() -> int:
    from PySide6.QtWidgets import QApplication  # noqa: F401
    import print_engine as pe

    for requested, available, want in CASES:
        got = pe._match_printer_name(requested, available)
        assert got == want, "подбор %r: получили %r, ожидали %r" % (
            requested, got, want)
    print("подбор имени принтера: %d случаев ок" % len(CASES))
    print("═══ ВЫБОР ПРИНТЕРА: ИМЕНА СОГЛАСОВАНЫ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
