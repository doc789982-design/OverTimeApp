#!/usr/bin/env python3
"""Тест гарантии version.json в релизном zip (tools/stamp_zip_ci.py, батч 299).

Сборки 294–298 уходили в релиз без номера версии в архиве — клиенты
отказывались обновляться («В архиве нет номера версии») и могли
зациклиться на скачивании. Проверяем: скрипт докладывает version.json
в корень и _internal любого архива, заменяет чужой, не трогает
правильный и не падает без архива; итоговый архив читается обновлятором.
"""
import json
import subprocess
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import app_update  # noqa: E402

SCRIPT = ROOT / "tools" / "stamp_zip_ci.py"


def run_stamp(workspace: Path) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, str(SCRIPT), str(workspace)],
                          capture_output=True, text=True, encoding="utf-8")


def make_zip(path: Path, with_root=False, with_internal=False,
             root_version=None, internal_version=None):
    meta = {"name": "OVERTIMETAB"}
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("OVERTIMETAB.exe", b"MZ dummy")
        if with_root:
            m = dict(meta, version=root_version, build=1)
            z.writestr("version.json", json.dumps(m))
        if with_internal:
            m = dict(meta, version=internal_version, build=2)
            z.writestr("_internal/version.json", json.dumps(m))
        z.writestr("_internal/PySide6/Qt6Core.dll", b"dll bytes")


def read_meta(path: Path, entry: str) -> dict:
    with zipfile.ZipFile(path, "r") as z:
        return json.loads(z.read(entry).decode("utf-8"))


def main() -> int:
    ws = Path("/tmp/zip_stamp_ws")
    if ws.exists():
        for p in ws.iterdir():
            p.unlink()
    ws.mkdir(parents=True, exist_ok=True)

    # ожидаемое содержимое — из репозитория (build/version.json или AppTheme)
    proc_probe = run_stamp(ws)  # архива нет — просто прогон импорта
    assert proc_probe.returncode == 0, proc_probe.stderr

    identity = json.loads((ROOT / "build" / "version.json").read_text(encoding="utf-8")) \
        if (ROOT / "build" / "version.json").exists() else None
    assert identity and identity.get("version") and identity.get("build"), \
        "нет build/version.json — запустите спеку или make_resources"
    want_ver = identity["version"]
    want_bld = int(identity["build"])

    # ── A. Архив без version.json: докладываем в корень и _internal ──
    zp = ws / "OVERTIMETAB_BETA.1_abc1234.zip"
    make_zip(zp)
    r = run_stamp(ws)
    assert r.returncode == 0, r.stdout + r.stderr
    m_root = read_meta(zp, "version.json")
    m_int = read_meta(zp, "_internal/version.json")
    assert m_root["version"] == want_ver and int(m_root["build"]) == want_bld, m_root
    assert m_int["version"] == want_ver, m_int
    assert app_update._looks_like_app_version(m_root["version"]), \
        "машинная строка должна выглядеть как X.Y.Z-версия"
    print("A: архив без версии — version.json в корне и _internal, машина %s ✓" % want_ver)

    # ── B. Правильный архив — нетронут (контроль читается) ──
    r = run_stamp(ws)
    assert r.returncode == 0, r.stdout + r.stderr
    assert "архив в порядке" in r.stdout, r.stdout
    print("B: правильный архив не трогаем ✓")

    # ── C. Чужой version.json в корне — заменён на правильный ──
    make_zip(zp, with_root=True, with_internal=True,
             root_version="0.0.1-ancient.1", internal_version="0.0.1-ancient.1")
    r = run_stamp(ws)
    assert r.returncode == 0, r.stdout + r.stderr
    assert read_meta(zp, "version.json")["version"] == want_ver
    assert read_meta(zp, "_internal/version.json")["version"] == want_ver
    # и содержимое архива не потерялось при пересборке
    with zipfile.ZipFile(zp, "r") as z:
        names = {n.replace("\\", "/") for n in z.namelist()}
        assert "OVERTIMETAB.exe" in names and "_internal/PySide6/Qt6Core.dll" in names, names
    print("C: чужая версия в корне — заменена, файлы архива целы ✓")

    # ── D. Архива нет — не ошибка (шаг мог поменяться) ──
    zp.unlink()
    r = run_stamp(ws)
    assert r.returncode == 0 and "нечего проверять" in r.stdout, r.stdout + r.stderr
    print("D: без архива — тихий пропуск ✓")

    # ── E. Итоговый архив читается обновлятором (как клиентом) ──
    zp = ws / "OVERTIMETAB_BETA.1_abc1234.zip"
    make_zip(zp, with_internal=True, internal_version=want_ver)
    r = run_stamp(ws)
    assert r.returncode == 0, r.stdout + r.stderr
    info = app_update.describe_package(zp)
    assert info.get("version") == want_ver and int(info.get("build") or 0) == want_bld, info
    # и старый строковый клиент видит обновление с любой древней версии
    assert app_update.is_newer(want_ver, "2.0.0-ALPHA.293", want_bld, 293)
    assert app_update.is_newer(want_ver, "BETA.1", want_bld, 293)
    print("E: обновлятор читает архив, старые клиенты видят обновление ✓")

    print("═══ ZIP РЕЛИЗА: ВЕРСИЯ ГАРАНТИРОВАНА ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
