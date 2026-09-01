# -*- coding: utf-8 -*-
"""
Готовит файлы для веб-хранилища обновлений.

Программа умеет обновляться через интернет: в настройках задают адрес
хранилища (например https://site.ru/updates/), где лежат:
    version.json                 — «витрина»: версия + где взять архив
    OVERTIMETAB_<версия>_<sha>.zip — сам архив новой версии

Этот скрипт по готовому zip делает server-side version.json (с ссылкой на
архив и его контрольной суммой sha256) и печатает, что и куда загрузить.

Запуск из корня репозитория:
    python tools/make_web_release.py path/to/OVERTIMETAB_2.0.0-ALPHA.20_abc1234.zip
    python tools/make_web_release.py --local path/to/archive.zip --version 2.0.0-ALPHA.20 --build 126
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
THEME = ROOT / "components" / "AppTheme.qml"


def read_identity() -> tuple[str, int]:
    text = THEME.read_text(encoding="utf-8")
    m = re.search(r'appVersion:\s*"([^"]+)"', text)
    if not m:
        sys.exit("Не нашли appVersion в AppTheme.qml")
    bm = re.search(r"appBuild:\s*(\d+)", text)
    return m.group(1), int(bm.group(1)) if bm else 0


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description="Сгенерировать version.json для веб-хранилища")
    parser.add_argument("zip", nargs="?", help="путь к архиву OVERTIMETAB_*.zip")
    parser.add_argument("--version", help="версия (по умолчанию из AppTheme.qml)")
    parser.add_argument("--build", type=int, help="сборка (по умолчанию из AppTheme.qml)")
    parser.add_argument("--out", help="куда писать version.json (по умолчанию рядом с zip)")
    args = parser.parse_args()

    if not args.zip:
        sys.exit("Укажите путь к zip, например: python tools/make_web_release.py path/to/OVERTIMETAB_*.zip")
    zip_path = Path(args.zip)
    if not zip_path.is_file():
        sys.exit(f"Нет файла {zip_path}")

    version = args.version
    build = args.build if args.build is not None else 0
    if not version:
        version, build = read_identity()

    # normalize "scan" version (2.0.0-ALPHA.20) same way the app reads it
    sys.path.insert(0, str(ROOT))
    try:
        from app_update import scan_version
        scan = scan_version(version, build)
    except Exception:
        scan = version

    payload = {
        "name": "OVERTIMETAB",
        "version": scan,
        "url": zip_path.name,
        "sha256": sha256_of(zip_path),
    }
    if int(build or 0) > 0:
        payload["build"] = int(build)
    if version and version != scan:
        payload["display"] = version

    out = Path(args.out) if args.out else zip_path.with_name("version.json")
    out.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    print(f"Версия: {version}" + (f" · сборка {build}" if build else ""))
    print(f"Архив:  {zip_path.name} ({zip_path.stat().st_size / (1024*1024):.1f} МБ)")
    print(f"version.json записан: {out}")
    print()
    print("Загрузите на сайт в одну папку эти два файла:")
    print(f"  {zip_path.name}")
    print(f"  {out.name}")
    print("В настройках программы укажите адрес этой папки (пример: https://site.ru/updates).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
