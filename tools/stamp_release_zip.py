# -*- coding: utf-8 -*-
"""
Дописывает version.json в zip GitHub Release.

Сборка Windows на Actions кладёт номер версии только внутрь exe/ресурсов.
Уже установленная программа смотрит в архиве файл version.json —
не находит и пишет «в архиве нет номера версии». Кнопка обновления
тогда не появляется.

Запуск из корня, когда zip уже висит на релизе:
    python tools/stamp_release_zip.py
    python tools/stamp_release_zip.py v2.0.0-ALPHA.28
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
THEME = ROOT / "components" / "AppTheme.qml"
CHANGELOG = ROOT / "CHANGELOG.md"


def read_identity() -> tuple[str, int]:
    text = THEME.read_text(encoding="utf-8")
    m = re.search(r'appVersion:\s*"([^"]+)"', text)
    if not m:
        sys.exit("Не нашли appVersion в AppTheme.qml")
    bm = re.search(r"appBuild:\s*(\d+)", text)
    return m.group(1), int(bm.group(1)) if bm else 0


def read_version() -> str:
    return read_identity()[0]


def version_payload(version: str, build: int = 0) -> bytes:
    payload = {"name": "OVERTIMETAB", "version": version}
    if int(build or 0) > 0:
        payload["build"] = int(build)
    return (json.dumps(payload, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def _norm(name: str) -> str:
    return name.replace("\\", "/").lstrip("/").lower()


def stamp_zip(zip_path: Path, version: str, build: int = 0) -> list[str]:
    """Добавляет version.json в корень архива и в _internal/, если их ещё нет."""
    payload = version_payload(version, build)
    added: list[str] = []
    with zipfile.ZipFile(zip_path, "r") as zf:
        names = {_norm(n) for n in zf.namelist()}
    need = []
    if "version.json" not in names:
        need.append("version.json")
    if "_internal/version.json" not in names:
        need.append("_internal/version.json")
    if CHANGELOG.exists() and "changelog.md" not in names:
        need.append("CHANGELOG.md")
    if not need:
        return added
    try:
        with zipfile.ZipFile(zip_path, "a") as zf:
            for name in need:
                if name == "CHANGELOG.md":
                    zf.writestr(name, CHANGELOG.read_bytes())
                else:
                    zf.writestr(name, payload)
                added.append(name)
        return added
    except Exception as exc:
        print(f"append не вышел ({exc}), переписываем архив…")
    tmp = zip_path.with_suffix(".stamped.zip")
    with zipfile.ZipFile(zip_path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for info in zin.infolist():
            if info.is_dir():
                continue
            zout.writestr(info, zin.read(info.filename))
        for name in need:
            if name == "CHANGELOG.md":
                zout.writestr(name, CHANGELOG.read_bytes())
            else:
                zout.writestr(name, payload)
            added.append(name)
    tmp.replace(zip_path)
    return added


def run_gh(args: list[str], check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(["gh", *args], cwd=str(ROOT), text=True, capture_output=True, check=check)


def main() -> int:
    if len(sys.argv) >= 3 and sys.argv[1] == "--local":
        zip_path = Path(sys.argv[2])
        if not zip_path.is_file():
            sys.exit(f"Нет файла {zip_path}")
        ver, bld = read_identity()
        added = stamp_zip(zip_path, ver, bld)
        if added:
            print("дописали:", ", ".join(added))
        else:
            print("version.json уже был в архиве")
        return 0
    version, build = read_identity()
    tag = sys.argv[1] if len(sys.argv) > 1 else f"v{version}"
    print(f"релиз:  {tag}")
    print(f"версия: {version}" + (f" · сборка {build}" if build else ""))

    view = run_gh(["release", "view", tag, "--json", "assets"], check=False)
    if view.returncode != 0:
        sys.stderr.write(view.stderr or "релиз не найден\n")
        return 1
    assets = json.loads(view.stdout).get("assets") or []
    zips = [a for a in assets if str(a.get("name") or "").lower().endswith(".zip")]
    if not zips:
        sys.exit("На релизе ещё нет zip — подождите сборку Windows")
    asset_name = zips[0]["name"]
    print(f"архив:  {asset_name}")

    with tempfile.TemporaryDirectory(prefix="ot-stamp-") as td:
        dest = Path(td)
        r = run_gh(
            ["release", "download", tag, "-p", asset_name, "-D", str(dest)],
            check=False,
        )
        if r.returncode != 0:
            sys.stderr.write(r.stderr or r.stdout or "не скачали zip\n")
            return r.returncode
        zip_path = dest / asset_name
        if not zip_path.exists():
            found = list(dest.glob("*.zip"))
            if not found:
                sys.exit("Скачали релиз, но zip не нашли")
            zip_path = found[0]
        added = stamp_zip(zip_path, version, build)
        if not added:
            print("version.json уже был в архиве — ничего не меняли")
            return 0
        print("дописали:", ", ".join(added))
        up = run_gh(
            ["release", "upload", tag, str(zip_path), "--clobber"],
            check=False,
        )
        if up.returncode != 0:
            sys.stderr.write(up.stderr or up.stdout or "не залили zip\n")
            return up.returncode
        print("залили обратно на", tag)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
