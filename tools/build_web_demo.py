# -*- coding: utf-8 -*-
"""
Собирает СТАТИЧЕСКУЮ демо-версию веб-интерфейса для GitHub Pages.

Зачем: превью и локальный режим --web зависят от живого сервера.
GitHub Pages — это просто URL: страница и все данные вшиты в файлы,
«не запуститься» там нечему.

Как: поднимается настоящий движок (webapp.server) на отдельном порту
со свежей демо-базой, и у него снимаются ровно те ответы, которые
страница запрашивает при работе (bootstrap/month/year/day за 2025–2027).
Ни одной формулы в JS: всё посчитано движком заранее.

Запуск из корня репозитория:
    python3 tools/build_web_demo.py
Результат: docs/ (index.html, app.css, app.js, demo-data.js)
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PORT = 8189
YEARS = (2025, 2026, 2027)
BASE = f"http://127.0.0.1:{PORT}"
OUT = ROOT / "docs"


def wait_up(timeout=30.0):
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            urllib.request.urlopen(BASE + "/api/bootstrap?year=2026", timeout=2)
            return True
        except Exception:
            time.sleep(0.25)
    return False


def get(u: str):
    return json.loads(urllib.request.urlopen(BASE + u, timeout=30).read())


def main() -> int:
    db_path = Path(tempfile.gettempdir()) / "overtimetab_web_demo_build.db"
    if db_path.exists():
        db_path.unlink()

    # свой экземпляр движка: свежая демо-база, отдельный порт
    code = (
        "import sys; sys.path.insert(0, %r)\n"
        "import webapp.server as ws\n"
        "ws.DEMO_DB = %r\n"
        "ws.PORT = %d\n"
        "ws.main()\n" % (str(ROOT), str(db_path), PORT)
    )
    proc = subprocess.Popen([sys.executable, "-u", "-c", code], cwd=str(ROOT))
    try:
        if not wait_up():
            print("сервер демо не поднялся")
            return 1
        print("движок поднялся, снимаю ответы…")

        data = {}
        emps = []
        for y in YEARS:
            b = get(f"/api/bootstrap?year={y}")
            data[f"/api/bootstrap?year={y}"] = b
            if not emps:
                emps = [e["id"] for e in b["employees"]]

        n = 0
        for emp in emps:
            for y in YEARS:
                data[f"/api/year?emp={emp}&year={y}"] = get(f"/api/year?emp={emp}&year={y}")
                for m in range(1, 13):
                    data[f"/api/month?emp={emp}&year={y}&month={m}"] = \
                        get(f"/api/month?emp={emp}&year={y}&month={m}")
                    n += 1
            print("  сотрудник %s: месяцы и годы сняты" % emp)

        # дни: инспектор открывается кликом по любой дате
        import calendar as cal_lib
        for emp in emps:
            for y in YEARS:
                for m in range(1, 13):
                    for d in range(1, cal_lib.monthrange(y, m)[1] + 1):
                        u = f"/api/day?emp={emp}&date={y:04d}-{m:02d}-{d:02d}"
                        data[u] = get(u)
            print("  сотрудник %s: дни сняты" % emp)

        print("всего ответов: %d" % len(data))
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()
        if db_path.exists():
            db_path.unlink()

    # ── вывод в docs/ ─────────────────────────────────────────
    if OUT.exists():
        shutil.rmtree(OUT)
    OUT.mkdir()

    demo_js = OUT / "demo-data.js"
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    demo_js.write_text(
        "// Сгенерировано tools/build_web_demo.py — снимок ответов движка\n"
        "// (демо-база, годы %s). НЕ редактировать руками.\n"
        "window.DEMO_DATA = %s;\n" % (list(YEARS), payload),
        encoding="utf-8")

    shutil.copy(ROOT / "webapp" / "static" / "app.css", OUT / "app.css")
    shutil.copy(ROOT / "webapp" / "static" / "app.js", OUT / "app.js")

    html = (ROOT / "webapp" / "static" / "index.html").read_text(encoding="utf-8")
    html = html.replace(
        '<script src="app.js"></script>',
        '<script src="demo-data.js"></script>\n<script src="app.js"></script>')
    html = html.replace(
        "</body>",
        '<div style="position:fixed;bottom:10px;left:50%;transform:translateX(-50%);'
        'z-index:999;background:#334155;color:#fff;font:12px Segoe UI,Arial,sans-serif;'
        'padding:6px 14px;border-radius:999px;opacity:.92;box-shadow:0 2px 8px rgba(0,0,0,.25)">'
        'Демо-снимок (GitHub Pages) — записи отключены</div>\n</body>')
    (OUT / "index.html").write_text(html, encoding="utf-8")

    total = sum(f.stat().st_size for f in OUT.iterdir())
    print("docs/: %s (demo-data %.1f МБ, всего %.1f МБ; gzip на Pages ~%.0f КБ)" % (
        sorted(f.name for f in OUT.iterdir()),
        demo_js.stat().st_size / 1e6, total / 1e6,
        total * 0.1 / 1024))
    return 0


if __name__ == "__main__":
    sys.exit(main())
