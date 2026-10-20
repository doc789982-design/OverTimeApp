#!/usr/bin/env python3
"""Сквозной тест кнопки обновления в шапке (сборка 306).

Жалоба юзера: «после скачивания новой версии нажатие для обновления
самой программы не срабатывает; кнопка не видит архивы с обновой,
если какие-то проблемы со скачиванием; методы проверок — такие же,
как при запуске программы».

Корни (все три воспроизведены тестом):
  1. У Backend.applyReadyUpdate потерялся @Slot() (снят в 299-м при
     добавлении журнала попыток) — QML НЕ МОЖЕТ вызвать недекорированный
     метод: клик «Обновить» умирал с TypeError молча.
  2. Сканер локальных обновлений не заглядывал в папку update_download —
     полностью скачанный архив (сеть подвела ПОСЛЕ скачивания) был
     невидим для проверки по кнопке.
  3. Проверка с недоступной сетью стирала уже найденное обновление:
     «не смогли проверить» приравнивалось к «обновления нет».

Части:
  A — регрессия @Slot: каждый backend.X(…) из QML — слот (ловит класс
      бага 299 по всем кнопкам, не только этой);
  B — сеть мертва, архив лежит в update_download: проверка кнопкой
      находит его, готовит, кнопка показывает «готово»;
  C — найденное обновление не теряется при сбое сети (но честно
      исчезает, когда сервер доступен и говорит «не новее»);
  D — настоящий путь целиком на живом локальном сервере: проверка →
      скачивание с прогрессом → подготовка → «Обновить» доходит до
      запуска установки (без выхода из процесса);
  E — клики по настоящей QML-кнопке: покой → проверка, доступно →
      скачивание, готово → установка.

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen QSG_RASTER_BACKEND=1 \
    python3 qa/test_update_button_flow.py
"""
import http.server
import json
import os
import re
import shutil
import socket
import sys
import threading
import time
import zipfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import types
# Main.py импортирует win32print (только Windows) — в песочнице подменяем
sys.modules.setdefault("win32print", types.ModuleType("win32print"))

import app_update  # noqa: E402
import Main as MainMod  # noqa: E402
from Main import Backend  # noqa: E402

CUR_VER = "2.0.0-beta.305"
CUR_BUILD = 305
NEW_VER = "2.0.0-beta.306"
NEW_BUILD = 306
NEW_ZIP_NAME = "OVERTIMETAB_BETA.1_306test.zip"


# ---------------------------------------------------------------------
# Инфраструктура
# ---------------------------------------------------------------------

KEEP_ALIVE = []  # ссылки на Qt-объекты до конца процесса


def make_package_zip(path: Path, version: str, build: int) -> None:
    """Настоящий по форме zip обновления: exe + version.json в корне."""
    meta = {"name": "OVERTIMETAB", "version": version,
            "build": build, "display": "BETA.1"}
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("OVERTIMETAB.exe", b"MZ dummy update")
        z.writestr("version.json", json.dumps(meta))
        z.writestr("_internal/some.dll", b"dll bytes")


def make_app_dir(base: Path) -> Path:
    """Собранная «программа»: exe + version.json (+ _internal как у PyInstaller 6)."""
    app_dir = base / "OVERTIMETAB"
    (app_dir / "_internal").mkdir(parents=True, exist_ok=True)
    (app_dir / "OVERTIMETAB.exe").write_bytes(b"MZ dummy current")
    meta = {"name": "OVERTIMETAB", "version": CUR_VER,
            "build": CUR_BUILD, "display": "BETA.1"}
    (app_dir / "version.json").write_text(
        json.dumps(meta), encoding="utf-8")
    return app_dir


def make_backend(app_dir: Path, data_dir: Path) -> Backend:
    b = Backend.__new__(Backend)
    from PySide6.QtCore import QObject
    QObject.__init__(b)
    b.app_dir = Path(app_dir)
    b._data_dir = Path(data_dir)
    b.config_path = Path(data_dir) / "config.json"
    b.active_db = None
    b._app_version = CUR_VER
    b._app_build = CUR_BUILD
    b._update_url = ""
    b._update_ready = False
    b._update_version = ""
    b._update_status = ""
    b._update_source = ""
    b._update_busy = False
    b._update_checking = False
    b._pending_notify = False
    b._background_auto = False
    b._staging_silent = False
    b._remote_info = None
    b._remote_update_available = False
    b._remote_downloading = False
    b._remote_download_progress = 0
    b._remote_active_url = ""
    b._last_remote_error = ""
    b._remote_saw_reachable = False
    return b


def pump(app, until, timeout=15.0):
    """Крутит цикл событий, пока until() не станет истинным."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        app.processEvents()
        if until():
            return True
        time.sleep(0.01)
    app.processEvents()
    return until()


class ToastCatcher:
    def __init__(self, backend):
        self.toasts = []
        backend.showToast.connect(self._on)

    def _on(self, text, kind):
        self.toasts.append((text, kind))


def quiet_remote_urls(b):
    """Удалённые источники — один заведомо мёртвый localhost."""
    b._remote_candidate_urls = lambda: ["http://127.0.0.1:1/"]


# ---------------------------------------------------------------------
# A. Регрессия @Slot: всё, что QML зовёт у backend, — слот
# ---------------------------------------------------------------------

def part_a_slot_regression() -> None:
    qml_calls = set()
    for qml in list((ROOT / "components").glob("*.qml")) + [ROOT / "main.qml"]:
        text = qml.read_text(encoding="utf-8", errors="replace")
        # комментарии срезаем: «// Парсим backend.foo (формат…)» — не вызов
        text = re.sub(r"//[^\n]*", "", text)
        qml_calls.update(re.findall(r"\bbackend\.([A-Za-z_]\w*)\s*\(", text))
    assert qml_calls, "не нашли ни одного вызова backend.* в QML — тест слеп"

    src = (ROOT / "Main.py").read_text(encoding="utf-8")
    # пары «декоратор-строка, def-строка» внутри класса Backend
    decorated = set(re.findall(r"@[Ss]lot\([^)]*\)\s*\n\s*def\s+(\w+)", src))
    plain_defs = set(re.findall(r"^\s*def\s+(\w+)", src, re.MULTILINE))

    missing = sorted(n for n in qml_calls
                     if n in plain_defs and n not in decorated)
    assert not missing, (
        "QML вызывает методы Backend без @Slot — клики по ним мертвы: "
        + ", ".join(missing))

    # прямой виновник жалобы — дважды: исходно и в мета-объекте
    assert "applyReadyUpdate" in decorated, "applyReadyUpdate без @Slot"
    mo = Backend.staticMetaObject
    assert mo.indexOfMethod("applyReadyUpdate()") >= 0, \
        "applyReadyUpdate не виден мета-объекту — QML не вызовет"
    print("A: все %d вызовов backend.* из QML — слоты; "
          "applyReadyUpdate в мета-объекте ✓" % len(qml_calls))


# ---------------------------------------------------------------------
# B. Сеть мертва, архив уже скачан: проверка кнопкой видит его
# ---------------------------------------------------------------------

def part_b_offline_archive(app) -> None:
    ws = Path("/tmp/upd_flow_ws/b")
    shutil.rmtree(ws, ignore_errors=True)
    app_dir = make_app_dir(ws)
    data_dir = ws / "data"
    data_dir.mkdir(parents=True)

    # сеть подвела ПОСЛЕ скачивания: архив уже в update_download
    dl = app_dir / "update_download"
    dl.mkdir(parents=True, exist_ok=True)
    make_package_zip(dl / NEW_ZIP_NAME, NEW_VER, NEW_BUILD)

    # до фикса сканер update_download не видел
    found = app_update.pick_best_update(app_dir, CUR_VER, CUR_BUILD)
    assert found, "сканер не видит скачанный архив в update_download"
    assert found.get("already_staged") is False, found
    assert int(found.get("build") or 0) == NEW_BUILD, found

    b = make_backend(app_dir, data_dir)
    toasts = ToastCatcher(b)
    quiet_remote_urls(b)
    b.checkAllUpdateSources()   # та же последовательность, что при запуске
    assert pump(app, lambda: b.updateReady), "кнопка не пришла в «готово»"
    assert "306" in (b.updateVersion or ""), b.updateVersion
    staged = app_update.staged_info(app_dir)
    assert staged and int(staged.get("build") or 0) == NEW_BUILD, staged
    # без сети и без повторного скачивания: пакет подготовлен из локального
    # архива, папка скачивания очищена
    assert not (dl / NEW_ZIP_NAME).exists(), "скачанный zip не убран после подготовки"
    assert not any(t[1] == "error" for t in toasts.toasts), toasts.toasts
    print("B: мёртвая сеть, архив в update_download — проверка кнопкой "
          "нашла, подготовила, «готово к установке» ✓")


# ---------------------------------------------------------------------
# C. Сбой сети не прячет найденное обновление
# ---------------------------------------------------------------------

def part_c_keep_found(app) -> None:
    ws = Path("/tmp/upd_flow_ws/c")
    shutil.rmtree(ws, ignore_errors=True)
    app_dir = make_app_dir(ws)
    data_dir = ws / "data"
    data_dir.mkdir(parents=True)
    b = make_backend(app_dir, data_dir)
    quiet_remote_urls(b)

    info = {"name": "OVERTIMETAB", "version": NEW_VER, "build": NEW_BUILD,
            "url": NEW_ZIP_NAME, "display": "BETA.1"}

    # 1) сеть работала: проверка нашла обновление
    orig = app_update.fetch_update_info
    app_update.fetch_update_info = lambda url, timeout=15.0: (dict(info), "")
    try:
        b.checkAllUpdateSources()
        assert pump(app, lambda: b.remoteUpdateAvailable), "обновление не найдено"
    finally:
        app_update.fetch_update_info = orig
    assert b.remoteUpdateAvailable

    # 2) сеть умерла: повторная проверка не должна прятать найденное
    app_update.fetch_update_info = lambda url, timeout=15.0: (
        None, "Не удалось прочитать version.json (нет сети).")
    try:
        b.checkAllUpdateSources()
        assert pump(app, lambda: not b.updateChecking), "проверка не завершилась"
        assert b.remoteUpdateAvailable, (
            "сбой сети спрятал уже найденное обновление")
        assert b._remote_info and b._remote_info.get("url") == NEW_ZIP_NAME
    finally:
        app_update.fetch_update_info = orig

    # 3) сервер доступен и отвечает «у вас уже последняя» (валидный json
    #    с нашей же версией) — теперь честно прячем
    app_update.fetch_update_info = lambda url, timeout=15.0: (
        {"name": "OVERTIMETAB", "version": CUR_VER, "build": CUR_BUILD,
         "url": "x.zip", "display": "BETA.1"}, "")
    try:
        b.checkAllUpdateSources()
        assert pump(app, lambda: not b.updateChecking), "проверка не завершилась"
        assert not b.remoteUpdateAvailable, (
            "доступный сервер сказал «не новее», а кнопка всё ещё предлагает")
    finally:
        app_update.fetch_update_info = orig
    print("C: сбой сети не прячет найденное; «не новее» — прячет ✓")


# ---------------------------------------------------------------------
# D. Настоящий путь на живом локальном сервере
# ---------------------------------------------------------------------

def _serve(directory: Path):
    class Handler(http.server.SimpleHTTPRequestHandler):
        def __init__(self, *a, **kw):
            super().__init__(*a, directory=str(directory), **kw)

        def log_message(self, *a):
            pass

    sock = socket.socket()
    sock.bind(("127.0.0.1", 0))
    port = sock.getsockname()[1]
    sock.close()
    server = http.server.ThreadingHTTPServer(("127.0.0.1", port), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    return server, f"http://127.0.0.1:{port}/"


def part_d_full_flow(app) -> None:
    ws = Path("/tmp/upd_flow_ws/d")
    shutil.rmtree(ws, ignore_errors=True)
    app_dir = make_app_dir(ws)
    data_dir = ws / "data"
    data_dir.mkdir(parents=True)
    site = ws / "site"
    site.mkdir()
    make_package_zip(site / NEW_ZIP_NAME, NEW_VER, NEW_BUILD)
    (site / "version.json").write_text(json.dumps({
        "name": "OVERTIMETAB", "version": NEW_VER, "build": NEW_BUILD,
        "url": NEW_ZIP_NAME, "display": "BETA.1",
    }), encoding="utf-8")
    server, base_url = _serve(site)
    try:
        b = make_backend(app_dir, data_dir)
        toasts = ToastCatcher(b)
        b._remote_candidate_urls = lambda: [base_url]

        # проверка кнопкой — те же методы, что при запуске (8 c на источник)
        b.checkAllUpdateSources()
        assert pump(app, lambda: b.remoteUpdateAvailable), "сервер не дал обновление"
        assert b._remote_info.get("base_url") == base_url, b._remote_info

        # скачивание с живого HTTP-сервера
        b.startRemoteDownload()
        assert pump(app, lambda: not b.remoteDownloading and b.updateReady), \
            "скачивание/подготовка не дошли до «готово»"
        staged = app_update.staged_info(app_dir)
        assert staged and int(staged.get("build") or 0) == NEW_BUILD, staged
        assert b._remote_download_progress == 100, b._remote_download_progress

        # клик «Обновить»: запуск установки, но процесс не убиваем —
        # подменяем шаг запуска и ловим вызов на месте
        called = {}
        orig_swap = app_update.launch_file_swap
        def fake_swap(src, dst, pid):
            called["src"] = str(src)
            called["dst"] = str(dst)
            called["pid"] = int(pid)
            raise RuntimeError("stop before exit")
        app_update.launch_file_swap = fake_swap
        try:
            from PySide6.QtCore import QMetaObject
            ok = QMetaObject.invokeMethod(b, "applyReadyUpdate")
            assert ok, "applyReadyUpdate не вызван через мета-объект"
        finally:
            app_update.launch_file_swap = orig_swap
        app.processEvents()
        assert called.get("src") == staged["root"], called
        assert called.get("dst") == str(app_update.install_root(app_dir)), called
        assert called.get("pid") == os.getpid(), called
        # попытка записана в журнал (защита от петли) и озвучена тостом
        assert b._apply_attempts(NEW_BUILD) == 1
        assert any("Не удалось запустить обновление" in t[0] for t in toasts.toasts)
        print("D: живой сервер → проверка → скачивание → подготовка → "
              "«Обновить» дошёл до запуска установки ✓")
    finally:
        server.shutdown()


# ---------------------------------------------------------------------
# E. Клики по настоящей QML-кнопке
# ---------------------------------------------------------------------

def part_e_qml_clicks(app) -> None:
    from PySide6.QtCore import QObject, Property, Signal, Slot
    from PySide6.QtQml import QQmlApplicationEngine
    from PySide6.QtTest import QTest
    from PySide6.QtCore import Qt, QPointF

    class StubBackend(QObject):
        remoteUpdateAvailableChanged = Signal()
        remoteDownloadingChanged = Signal()
        remoteDownloadProgressChanged = Signal()
        updateReadyChanged = Signal()
        updateCheckingChanged = Signal()
        reducedMotionChanged = Signal()

        def __init__(self):
            super().__init__()
            self._reduced = False
            self._avail = False
            self._down = False
            self._prog = 0
            self._ready = False
            self._chk = False
            self.calls = []

        remoteUpdateAvailable = Property(
            bool, lambda s: s._avail, notify=remoteUpdateAvailableChanged)
        remoteDownloading = Property(
            bool, lambda s: s._down, notify=remoteDownloadingChanged)
        remoteDownloadProgress = Property(
            int, lambda s: s._prog, notify=remoteDownloadProgressChanged)
        updateReady = Property(
            bool, lambda s: s._ready, notify=updateReadyChanged)
        updateChecking = Property(
            bool, lambda s: s._chk, notify=updateCheckingChanged)
        reducedMotion = Property(
            bool, lambda s: s._reduced, notify=reducedMotionChanged)

        FIELD = {"remoteUpdateAvailable": "_avail",
                 "remoteDownloading": "_down",
                 "remoteDownloadProgress": "_prog",
                 "updateReady": "_ready",
                 "updateChecking": "_chk"}

        def _set(self, prop, val):
            setattr(self, self.FIELD[prop], val)
            getattr(self, prop + "Changed").emit()

        @Slot()
        def checkAllUpdateSources(self):
            self.calls.append("check")

        @Slot()
        def startRemoteDownload(self):
            self.calls.append("download")

        @Slot()
        def applyReadyUpdate(self):
            self.calls.append("apply")

    stub = StubBackend()
    engine = QQmlApplicationEngine()
    engine.rootContext().setContextProperty("backend", stub)
    KEEP_ALIVE.append((engine, stub))
    wrapper = ROOT / "_render_upd_flow.qml"
    wrapper.write_text("""
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    width: 140; height: 120; visible: true
    Row {
        anchors.centerIn: parent
        AppUI.UpdateDownloadButton { height: 46 }
    }
}
""", encoding="utf-8")
    try:
        engine.load(str(wrapper))
        assert engine.rootObjects(), "кнопка не загрузилась"
        win = engine.rootObjects()[0]

        from PySide6.QtQuick import QQuickItem

        item = None
        for _ in range(50):
            item = win.findChild(QQuickItem, "updateButton")
            if item:
                break
            app.processEvents()
            time.sleep(0.02)
        assert item, "кнопка с objectName updateButton не найдена"

        def click():
            pos = item.mapToScene(QPointF(23, 23)).toPoint()
            QTest.mouseClick(win, Qt.MouseButton.LeftButton,
                             Qt.KeyboardModifier.NoModifier, pos)
            for _ in range(10):
                app.processEvents()

        # покой → проверка
        click()
        assert stub.calls == ["check"], stub.calls
        # во время проверки клик молчит (значок уже крутится)
        stub._set("updateChecking", True)
        app.processEvents()
        assert item.property("mode") == "checking", item.property("mode")
        click()
        assert stub.calls == ["check"], stub.calls
        stub._set("updateChecking", False)
        # доступно → скачивание
        stub._set("remoteUpdateAvailable", True)
        app.processEvents()
        assert item.property("mode") == "available", item.property("mode")
        click()
        assert stub.calls == ["check", "download"], stub.calls
        # во время загрузки клик молчит
        stub._set("remoteDownloading", True)
        stub._set("remoteDownloadProgress", 40)
        app.processEvents()
        assert item.property("mode") == "downloading", item.property("mode")
        click()
        assert stub.calls == ["check", "download"], stub.calls
        stub._set("remoteDownloading", False)
        # готово → установка
        stub._set("updateReady", True)
        app.processEvents()
        assert item.property("mode") == "ready", item.property("mode")
        assert item.property("clickable") is True
        click()
        assert stub.calls == ["check", "download", "apply"], stub.calls
        stub._set("updateReady", False)
        stub._set("remoteUpdateAvailable", False)
        stub._set("updateChecking", True)
        app.processEvents()
        assert item.property("mode") == "checking"
        stub._set("updateChecking", False)
        app.processEvents()
        assert item.property("mode") == "idle", item.property("mode")
        print("E: клики QML-кнопки: покой→проверка, доступно→скачать, "
              "загрузка молчит, готово→установить ✓")
    finally:
        win.close()
        wrapper.unlink(missing_ok=True)


def main() -> int:
    from PySide6.QtGui import QGuiApplication
    app = QGuiApplication([])

    part_a_slot_regression()
    part_b_offline_archive(app)
    part_c_keep_found(app)
    part_d_full_flow(app)
    part_e_qml_clicks(app)

    print("═══ КНОПКА ОБНОВЛЕНИЯ: ЦЕПОЧКА ЦЕЛИКОМ РАБОТАЕТ ═══")
    return 0


if __name__ == "__main__":
    sys.exit(main())
