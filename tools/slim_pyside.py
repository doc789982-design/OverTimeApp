# -*- coding: utf-8 -*-
"""
Фильтр «лишнего Qt» для сборки OVERTIMETAB.

Зачем это нужно простыми словами
--------------------------------
PySide6 — это огромный конструктор. В нём есть браузер, 3D, камера,
карты, Bluetooth и ещё куча вещей, которых в табеле нет.

Если сказать сборщику «положи весь PySide6», в папку с программой
уедет всё это добро — сотни мегабайт.

Этот файл — список того, что программе НЕ нужно.
Сборщик сначала забирает Qt целиком (чтобы ничего нужного не забыть),
а потом по этому списку выкидывает мусор.

Что оставляем (этим программа реально пользуется):
  окна и кнопки, QML-интерфейс, календарь, тени/размытие,
  иконки SVG, диалоги «открыть файл», сеть только для
  «не запускай программу дважды».

Что выкидываем: браузер Qt, 3D, видео, графики, Bluetooth,
карты, PDF-читалку, дизайнер форм и прочие неиспользуемые модули.

Запускать руками не нужно — его подхватывает tools/overtimetab.spec.
"""
from __future__ import annotations

from pathlib import Path


# ---------------------------------------------------------------------------
# Модули Python, которые нельзя даже начинать собирать.
# Если сборщик увидит «import PySide6.QtWebEngine» — он потащит
# за ним ещё ~150 МБ. Мы говорим: этого импорта в программе нет.
# ---------------------------------------------------------------------------
UNUSED_PYSIDE_MODULES = (
    "PySide6.QtWebEngine",
    "PySide6.QtWebEngineCore",
    "PySide6.QtWebEngineWidgets",
    "PySide6.QtWebEngineQuick",
    "PySide6.Qt3DAnimation",
    "PySide6.Qt3DCore",
    "PySide6.Qt3DExtras",
    "PySide6.Qt3DInput",
    "PySide6.Qt3DLogic",
    "PySide6.Qt3DRender",
    "PySide6.QtCharts",
    "PySide6.QtDataVisualization",
    "PySide6.QtGraphs",
    "PySide6.QtGraphsWidgets",
    "PySide6.QtMultimedia",
    "PySide6.QtMultimediaWidgets",
    "PySide6.QtBluetooth",
    "PySide6.QtNfc",
    "PySide6.QtPositioning",
    "PySide6.QtLocation",
    "PySide6.QtSensors",
    "PySide6.QtSerialPort",
    "PySide6.QtSerialBus",
    "PySide6.QtRemoteObjects",
    "PySide6.QtScxml",
    "PySide6.QtStateMachine",
    "PySide6.QtTextToSpeech",
    "PySide6.QtWebChannel",
    "PySide6.QtWebSockets",
    "PySide6.QtWebView",
    "PySide6.QtPdf",
    "PySide6.QtPdfWidgets",
    "PySide6.QtQuick3D",
    "PySide6.QtSpatialAudio",
    "PySide6.QtHttpServer",
    "PySide6.QtDesigner",
    "PySide6.QtHelp",
    "PySide6.QtTest",
    "PySide6.QtSql",           # база у нас через обычный sqlite3, не через Qt
    "PySide6.QtPrintSupport",  # печать идёт через Excel, не через Qt
    "PySide6.QtUiTools",
    "PySide6.QtXml",
    "PySide6.QtDBus",
    "PySide6.QtExampleIcons",
)

# Лишние куски обычного Python, которые сборщик любит прихватить «на всякий».
UNUSED_STDLIB = (
    "tkinter",
    "turtle",
    "unittest",
    "test",
    "pydoc",
    "doctest",
    "pdb",
    "idlelib",
    "lib2to3",
    "ensurepip",
    "venv",
    "xmlrpc",
    "http.server",
    "matplotlib",
    "numpy",
    "pandas",
    "PIL",
    "Pillow",
    "pytest",
)

# ---------------------------------------------------------------------------
# Куски путей/имён файлов, из-за которых файл выкидывается.
# Сравниваем в нижнем регистре, со слэшами «/».
# Пишем достаточно длинные слова, чтобы случайно не задеть нужный файл.
# ---------------------------------------------------------------------------
_DROP_PATH_PARTS = (
    # Браузер Qt (самый жирный кусок, часто 100–180 МБ один только он)
    "/qtwebengine",
    "qt6webengine",
    "qtwebengineprocess",
    "icudtl.dat",
    # 3D
    "qt63d",
    "qt6quick3d",
    "/qt3d",
    "/qtquick3d",
    # Медиа / камера / звук
    "qt6multimedia",
    "/qtmultimedia",
    "qt6spatialaudio",
    # Графики и визуализация
    "qt6charts",
    "/qtcharts",
    "qt6datavisualization",
    "/qtdatavisualization",
    "qt6graphs",
    "/qtgraphs",
    # Железо и сети, которых нет в табеле
    "qt6bluetooth",
    "/qtbluetooth",
    "qt6nfc",
    "/qtnfc",
    "qt6positioning",
    "/qtpositioning",
    "qt6location",
    "/qtlocation",
    "qt6sensors",
    "/qtsensors",
    "qt6serialport",
    "/qtserialport",
    "qt6serialbus",
    "/qtserialbus",
    "qt6remoteobjects",
    "/qtremoteobjects",
    "qt6scxml",
    "/qtscxml",
    "qt6statemachine",
    "/qtstatemachine",
    "qt6texttospeech",
    "/qttexttospeech",
    "qt6webchannel",
    "/qtwebchannel",
    "qt6websockets",
    "/qtwebsockets",
    "qt6webview",
    "/qtwebview",
    "qt6pdf",
    "/qtpdf",
    "/qtquick/pdf",
    "qt6httpserver",
    # Инструменты разработчика Qt, в готовой программе не нужны
    "qt6designer",
    "qt6help",
    "qt6test",
    "/qttest",
    "qt6sql",
    "qt6printsupport",
    "qt6uitools",
    "qt6xml",
    "qt6dbus",
    "qt6lottie",
    "virtualkeyboard",
    # Стили кнопок, которыми программа не пользуется.
    # В Main.py жёстко выставлен стиль Basic — остальные темы Qt не грузятся.
    "/qtquick/controls/material",
    "/qtquick/controls/imagine",
    "/qtquick/controls/universal",
    "/qtquick/controls/fusion",
    "/qtquick/controls/fluentwinui3",
    "/qtquick/controls/windows",
    "/qtquick/controls/ios",
    "/qtquick/controls/macos",
    "/qtquick/controls/designer",
    "/qtquick/controls/nativestyle",
    "qt6quickcontrols2material",
    "qt6quickcontrols2imagine",
    "qt6quickcontrols2universal",
    "qt6quickcontrols2fusion",
    "qt6quickcontrols2fluent",
    "qt6quickcontrols2windows",
    "qt6quicktimeline",
    "/qtquick/scene2d",
    "/qtquick/scene3d",
    "/qtquick/particles",
    "/qtquick/timeline",
    "/qtqml/xmllistmodel",
    # Плагины Qt, которые табель не открывает
    "/plugins/sqldrivers",
    "/plugins/multimedia",
    "/plugins/position",
    "/plugins/sensors",
    "/plugins/canbus",
    "/plugins/qmltooling",
    "/plugins/geometryloaders",
    "/plugins/renderers",
    "/plugins/sceneparsers",
    "/plugins/assetimporters",
    "/plugins/webview",
    "/plugins/networkinformation",
    "/plugins/designer",
    "/plugins/help",
    # Обвязка для компиляции биндингов — в exe не нужна
    "/pyside6/include",
    "/pyside6/metatypes",
    "/pyside6/typesystems",
    "/pyside6/glue",
    "/pyside6/doc",
    "/pyside6/examples",
    "/pyside6/scripts",
)

def _norm(path_like) -> str:
    """Путь к одному виду: маленькие буквы и прямые слэши."""
    return str(path_like or "").replace("\\", "/").lower()


def is_unused_hiddenimport(name: str) -> bool:
    """Этот python-модуль табель не импортирует — в сборку не кладём."""
    n = name or ""
    for prefix in UNUSED_PYSIDE_MODULES:
        if n == prefix or n.startswith(prefix + "."):
            return True
    return False


def should_keep(src, dest="") -> bool:
    """
    Оставить файл в сборке или выкинуть.

    src  — откуда файл на диске сборщика
    dest — куда его хотели положить внутри папки программы
    """
    blob = _norm(src) + " | " + _norm(dest)

    # Сначала режем явно лишнее (браузер, 3D, камера…).
    # Список специально длинный и конкретный, чтобы не задеть
    # нужные Qt6Quick / Qt6Core / SVG / тени.
    for part in _DROP_PATH_PARTS:
        if part in blob:
            return False

    # Переводы Qt: оставляем только русский и английский
    # (диалог «Открыть файл» берёт подписи отсюда).
    if "/translations/" in blob:
        return blob.endswith("_ru.qm") or blob.endswith("_en.qm")

    # Отладочные символы и библиотеки для компилятора C++ — пользователю не нужны
    if blob.endswith(".pdb") or blob.endswith(".lib") or blob.endswith(".exp"):
        return False

    return True


def filter_toc(toc):
    """
    toc — список файлов сборщика: (источник, назначение, тип).
    Возвращает (оставшееся, выброшенное).
    """
    kept, dropped = [], []
    for entry in toc:
        src = entry[0] if entry else ""
        dest = entry[1] if len(entry) > 1 else ""
        if should_keep(src, dest):
            kept.append(entry)
        else:
            dropped.append(entry)
    return kept, dropped


def toc_bytes(toc) -> int:
    """Сколько весят исходные файлы из списка. Для красивого отчёта в логе."""
    total = 0
    for entry in toc:
        try:
            total += Path(entry[0]).stat().st_size
        except Exception:
            pass
    return total


def format_mb(num_bytes: int) -> str:
    return f"{num_bytes / (1024 * 1024):.1f} МБ"


def print_report(dropped, title: str) -> None:
    """Печатает, сколько выкинули и какие файлы были самыми жирными."""
    if not dropped:
        print(f"[slim] {title}: выкидывать было нечего")
        return
    sized = []
    for entry in dropped:
        try:
            sized.append((Path(entry[0]).stat().st_size, entry[0]))
        except Exception:
            sized.append((0, entry[0]))
    sized.sort(reverse=True)
    total = sum(s for s, _ in sized)
    print(f"[slim] {title}: выкинуто {len(sized)} файлов, {format_mb(total)}")
    print("[slim]   самые большие из выкинутых:")
    for size, name in sized[:12]:
        print(f"[slim]     {format_mb(size):>8}  {Path(name).name}")
