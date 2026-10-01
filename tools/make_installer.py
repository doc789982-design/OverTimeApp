#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Склейка установщика: заглушка + zip программы = один exe.

Итоговый файл релиза OVERTIMETAB_<версия>_<sha>.exe — это простое
побайтовое сложение:

    installer_stub.exe  +  OVERTIMETAB_<...>.zip  →  OVERTIMETAB_<...>.exe

Zip-формат читается с конца файла, PyInstaller-заглушка ищет свой архив
полным сканом с конца — оба формата в одном файле не мешают друг другу.

Запуск (из корня репозитория, на CI — после сборки zip):
    python tools/make_installer.py --stub build_stub/installer_stub.exe \
        --zip OVERTIMETAB_BETA.1_a1b2c3d.zip --out OVERTIMETAB_BETA.1_a1b2c3d.exe

Скрипт проверяет результат: у exe должен читаться приклеенный zip с
OVERTIMETAB.exe внутри, размер — ровно сумма частей. Это часть релизной
цепочки: без установщика релиз не может быть опубликован (см. RELEASING.md
и проверку ассетов в tools/publish_release.py).
"""
from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path

# Windows-консоль (и GitHub Actions) по умолчанию использует cp1251/cp1252 —
# русские буквы в print падают с UnicodeEncodeError ЕЩЁ ДО выхода из скрипта,
# хотя вся работа уже сделана. Приводим потоки к UTF-8.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


class GlueError(Exception):
    pass


def make_installer(stub: Path, zip_path: Path, out: Path) -> Path:
    stub, zip_path, out = Path(stub), Path(zip_path), Path(out)
    if not stub.is_file():
        raise GlueError("Нет заглушки: %s" % stub)
    if not zip_path.is_file():
        raise GlueError("Нет архива программы: %s" % zip_path)
    if stub.read_bytes()[:2] != b"MZ":
        raise GlueError("Заглушка не похожа на exe (нет MZ): %s" % stub)

    with zipfile.ZipFile(zip_path, "r") as zf:
        names = [n for n in zf.namelist() if not n.endswith("/")]
        if not any(n.replace("\\", "/").lower().endswith("overtimetab.exe")
                   for n in names):
            raise GlueError("В архиве нет OVERTIMETAB.exe — клеить нельзя: %s"
                            % zip_path)
        zip_names = set(names)

    out.parent.mkdir(parents=True, exist_ok=True)
    with open(out, "wb") as dst, open(stub, "rb") as s, open(zip_path, "rb") as z:
        while True:
            chunk = s.read(1 << 20)
            if not chunk:
                break
            dst.write(chunk)
        while True:
            chunk = z.read(1 << 20)
            if not chunk:
                break
            dst.write(chunk)

    # проверка результата: приклеенный zip читается и совпадает с исходным
    try:
        with zipfile.ZipFile(out, "r") as zf:
            got = {n for n in zf.namelist() if not n.endswith("/")}
    except zipfile.BadZipFile:
        raise GlueError("Склеенный exe не читается как zip — сборка битая")
    if got != zip_names:
        raise GlueError("Состав архива в exe не совпал с исходным zip")
    if out.stat().st_size != stub.stat().st_size + zip_path.stat().st_size:
        raise GlueError("Размер exe не равен сумме частей")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="Склеить установщик из заглушки и zip")
    ap.add_argument("--stub", required=True, help="installer_stub.exe")
    ap.add_argument("--zip", required=True, help="zip программы (OVERTIMETAB_*.zip)")
    ap.add_argument("--out", required=True, help="итоговый exe (OVERTIMETAB_*.exe)")
    args = ap.parse_args()
    try:
        out = make_installer(Path(args.stub), Path(args.zip), Path(args.out))
    except GlueError as e:
        print("ОШИБКА: %s" % e, file=sys.stderr)
        return 1
    print("установщик: %s (%.1f МБ = заглушка %.1f + архив %.1f)"
          % (out.name,
             out.stat().st_size / 1e6,
             Path(args.stub).stat().st_size / 1e6,
             Path(args.zip).stat().st_size / 1e6))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
