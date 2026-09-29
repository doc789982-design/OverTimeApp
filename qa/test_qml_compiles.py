#!/usr/bin/env python3
"""Компиляция всего main.qml: ловит синтаксические ошибки QML.

Проверяет main.qml и все компоненты через QQmlComponent: синтаксис и
разрешение типов (ошибки биндингов к backend при стабе — не ошибки).
Выловила реальный баг: лишняя скобка после вырезания кнопки экспорта.

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \\
    QSG_RASTER_BACKEND=1 python3 qa/test_qml_compiles.py
"""
import os
import sys

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main() -> int:
    from PySide6.QtGui import QGuiApplication
    from PySide6.QtQml import QQmlApplicationEngine, QQmlComponent

    app = QGuiApplication([])
    engine = QQmlApplicationEngine()
    comp = QQmlComponent(engine, os.path.join(ROOT, "main.qml"))
    if comp.isError():
        for err in comp.errors():
            print("ОШИБКА:", err.toString())
        return 1
    print("main.qml и все компоненты компилируются без ошибок")
    print("═══ QML ЦЕЛ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
