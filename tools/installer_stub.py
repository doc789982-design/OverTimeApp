#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Заглушка-установщик OVERTIMETAB (то, что человек запускает двойным кликом).

Файл релиза OVERTIMETAB_<версия>_<sha>.exe устроен так:

    [эта заглушка, собранная PyInstaller] + [обычный zip программы в конце]

Человек запускает exe → работает эта заглушка. Она НЕ умеет ничего, кроме
распаковки: находит приклеенный в конец zip (zip-формат читается с конца
файла, что приклеено перед ним — не важно), распаковывает программу во
временную папку и запускает её с ключом --setup. Дальше мастер установки —
это уже сама программа (полноценный интерфейс в её стиле), см. Main.py.

Установленная программа при автообновлении качает НЕ этот exe, а zip —
поэтому здесь нет никакой логики обновления.

Только стандартная библиотека — заглушка собирается в маленький onefile-exe
(~8 МБ), без Qt. Сборка: pyinstaller --noconfirm --clean tools/installer_stub.spec
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


class PayloadError(Exception):
    """Приклеенный архив не читается или в нём нет программы."""


def extract_payload(exe_path: Path, dest_dir: Path) -> list[str]:
    """Распаковывает приклеенный zip из exe_path в dest_dir.

    Возвращает список имён файлов. Zip-ридер сам находит архив по концу
    файла, поэтому заглушка не мешает. Если архива нет или в нём нет
    OVERTIMETAB.exe — PayloadError с понятным текстом.
    """
    exe_path = Path(exe_path)
    if not exe_path.is_file():
        raise PayloadError("Не найден файл установщика: %s" % exe_path)
    try:
        with zipfile.ZipFile(exe_path, "r") as zf:
            names = [n for n in zf.namelist() if not n.endswith("/")]
            if not names:
                raise PayloadError("Файл установщика повреждён (пустой архив).")
            if not any(n.replace("\\", "/").lower().endswith("overtimetab.exe")
                       for n in names):
                raise PayloadError(
                    "Файл установщика повреждён: в нём нет OVERTIMETAB.exe. "
                    "Скачайте установщик заново.")
            dest_dir.mkdir(parents=True, exist_ok=True)
            for n in names:
                # защита от «выйти за папку» в кривом архиве
                rel = n.replace("\\", "/").lstrip("/")
                if not rel or ".." in rel.split("/"):
                    continue
                target = dest_dir / rel
                target.parent.mkdir(parents=True, exist_ok=True)
                with zf.open(n) as src, open(target, "wb") as out:
                    while True:
                        chunk = src.read(1 << 16)
                        if not chunk:
                            break
                        out.write(chunk)
            return names
    except zipfile.BadZipFile:
        raise PayloadError(
            "Не удалось прочитать данные программы — файл установщика "
            "повреждён или скачан не полностью. Скачайте его заново.")


def run_setup(extract_dir: Path) -> int:
    """Запускает распакованную программу в режиме мастера установки."""
    exe = Path(extract_dir) / "OVERTIMETAB.exe"
    if not exe.is_file():
        raise PayloadError("После распаковки не оказалось OVERTIMETAB.exe.")
    proc = subprocess.Popen([str(exe), "--setup"], cwd=str(extract_dir))
    return proc.pid


def _show_error(text: str) -> None:
    """Окно ошибки без всяких GUI-библиотек — MessageBox через ctypes."""
    print(text, file=sys.stderr)
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            None, text, "OVERTIMETAB — установка", 0x00000010)  # MB_ICONERROR
    except Exception:
        pass  # не-Windows (тесты): текст уже в stderr


def main() -> int:
    try:
        exe_path = Path(sys.executable if getattr(sys, "frozen", False)
                        else __file__)
        # уникальная папка: два установщика рядом не подерутся
        dest = Path(tempfile.mkdtemp(prefix="OVERTIMETAB_setup_"))
        extract_payload(exe_path, dest)
        run_setup(dest)
        return 0
    except PayloadError as e:
        _show_error(str(e))
        return 1
    except Exception as e:
        _show_error("Неожиданная ошибка установщика: %s" % e)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
