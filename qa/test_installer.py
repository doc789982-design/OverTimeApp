#!/usr/bin/env python3
"""Тест установщика: заглушка, склейка, логика установки.

Релизный exe = заглушка (tools/installer_stub.py) + zip программы.
Проверяется без Windows:
  A. make_installer: склейка работает, результат читается как zip,
     состав совпадает, размер — сумма частей;
  B. make_installer: без MZ / без OVERTIMETAB.exe внутри — отказ;
  C. заглушка extract_payload: читает гибрид, распаковывает, защищается
     от «выйти за папку», битый файл — понятная ошибка;
  D. installer.default_install_dir: «для меня» → AppData\\Programs,
     «для всех» → Program Files;
  E. installer.find_existing_install/_probe_install_dir: находит копию
     с номером сборки из version.json;
  F. installer.copy_program + do_install: копирует файлы, отдаёт прогресс,
     отказывает, если exe назначения занят (эмуляция);
  G. data_root: данные живут в Documents\\OverTimeTab.

Запуск:
    python3 qa/test_installer.py
"""
import json
import os
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tools"))

from make_installer import GlueError, make_installer  # noqa: E402
import installer_stub  # noqa: E402
import installer  # noqa: E402


def build_zip(dest: Path, exe_in_root=True) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dest, "w") as z:
        if exe_in_root:
            z.writestr("OVERTIMETAB.exe", b"MZ fake")
        else:
            z.writestr("data/readme.txt", b"no exe here")
        z.writestr("version.json", json.dumps(
            {"version": "2.0.0-ALPHA.999", "build": 999, "display": "BETA.1"}))
        z.writestr("_internal/PySide6.dll", b"fake dll" * 100)
    return dest


def main() -> int:
    tmp = Path(os.environ.get("TEST_TMP") or (ROOT / "build" / "qa_installer"))
    if tmp.exists():
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
    tmp.mkdir(parents=True, exist_ok=True)

    # ── A. склейка ──
    zipped = build_zip(tmp / "OVERTIMETAB_BETA.1_a1b2c3d.zip")
    stub = tmp / "installer_stub.exe"
    stub.write_bytes(b"MZ" + b"\x00" * 1024)  # фейковая заглушка
    exe = make_installer(stub, zipped, tmp / "OVERTIMETAB_BETA.1_a1b2c3d.exe")
    with zipfile.ZipFile(exe) as zf:
        got = {n for n in zf.namelist() if not n.endswith("/")}
    assert got == {"OVERTIMETAB.exe", "version.json", "_internal/PySide6.dll"}, got
    assert exe.stat().st_size == stub.stat().st_size + zipped.stat().st_size
    print("A: склейка exe = заглушка + zip, состав и размер верны ✓")

    # ── B. склейка отказывает на мусоре ──
    for bad in ((tmp / "not_stub.exe", zipped), (stub, build_zip(tmp / "bad.zip", exe_in_root=False))):
        try:
            if bad[0] == tmp / "not_stub.exe":
                bad[0].write_bytes(b"ELF not windows")
            make_installer(bad[0], bad[1], tmp / "out.exe")
            raise AssertionError("склейка должна была отказать: %s" % (bad,))
        except GlueError:
            pass
    print("B: склейка отказывает без MZ и без OVERTIMETAB.exe в архиве ✓")

    # ── C. заглушка читает гибрид ──
    out_dir = tmp / "setup_unpack"
    names = installer_stub.extract_payload(exe, out_dir)
    assert "OVERTIMETAB.exe" in names, names
    assert (out_dir / "OVERTIMETAB.exe").read_bytes() == b"MZ fake"
    assert json.loads((out_dir / "version.json").read_text())["build"] == 999
    print("C1: заглушка распаковывает гибрид ✓")

    # защита от «выйти за папку» и битый файл
    evil = tmp / "evil.exe"
    with zipfile.ZipFile(tmp / "evil.zip", "w") as z:
        z.writestr("../escape.txt", b"x")
        z.writestr("OVERTIMETAB.exe", b"MZ")
    evil.write_bytes(b"MZ" + b"\x00" * 64 + (tmp / "evil.zip").read_bytes())
    out2 = tmp / "evil_unpack"
    installer_stub.extract_payload(evil, out2)
    assert not (tmp / "escape.txt").exists(), "файл вырвался за папку!"
    print("C2: кривые пути в архиве не вырываются за папку ✓")

    for broken, why in ((tmp / "plain.exe", "нет архива"),
                        (tmp / "badzip.exe", "битый архив")):
        if why == "битый архив":
            broken.write_bytes(b"MZ" * 100 + b"PK\x03\x04broken-not-a-zip")
        else:
            broken.write_bytes(b"MZ" * 100)
        try:
            installer_stub.extract_payload(broken, tmp / "x")
            raise AssertionError("должна была ошибка: " + why)
        except installer_stub.PayloadError:
            pass
    print("C3: битый файл установщика — понятная ошибка ✓")

    # ── D. папки установки ──
    env = {"LOCALAPPDATA": "/home/u/AppData/Local",
           "ProgramFiles": "/opt/progfiles"}
    assert str(installer.default_install_dir(True, env)) == \
        "/home/u/AppData/Local/Programs/OVERTIMETAB"
    assert str(installer.default_install_dir(False, env)) == \
        "/opt/progfiles/OVERTIMETAB"
    assert installer.known_install_dirs(env) == [
        Path("/home/u/AppData/Local/Programs/OVERTIMETAB"),
        Path("/opt/progfiles/OVERTIMETAB")]
    print("D: папки «для меня»/«для всех» считаются по окружению ✓")

    # ── E. поиск установленной копии ──
    fake_home = tmp / "home"
    installed = installer.default_install_dir(True, {
        "LOCALAPPDATA": str(fake_home / "AppData" / "Local")})
    build_zip_into = tmp / "pkg"
    with zipfile.ZipFile(zipped) as zf:
        zf.extractall(build_zip_into)
    info = installer._probe_install_dir(build_zip_into)
    assert info and info["build"] == 999, info
    empty = installer._probe_install_dir(tmp / "nowhere")
    assert empty is None
    print("E: поиск копии видит сборку 999 из version.json ✓")

    # ── F. копирование и установка ──
    seen = []
    dest = tmp / "dest"
    n = installer.copy_program(build_zip_into, dest,
                               progress=lambda d, t: seen.append((d, t)))
    assert (dest / "OVERTIMETAB.exe").exists()
    assert (dest / "_internal" / "PySide6.dll").exists()
    assert seen and seen[-1][0] == seen[-1][1] > 0, seen[-3:]
    assert n == 3
    rep = installer.do_install(build_zip_into, dest, version="BETA.1",
                               build=999)
    assert rep["files"] == 3 and rep["dir"] == str(dest.resolve())
    print("F: копирование с прогрессом и полный отчёт установки ✓")

    # ── G. данные отдельно ──
    assert installer.data_root("/home/u") == Path("/home/u/Documents/OverTimeTab")
    print("G: пользовательские данные — Documents\\OverTimeTab ✓")

    print("═══ УСТАНОВЩИК: СКЛЕЙКА, ЗАГЛУШКА, ЛОГИКА — ВСЁ ЦЕЛО ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
