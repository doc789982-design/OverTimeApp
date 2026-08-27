# -*- coding: utf-8 -*-
"""
Подмена вызова GitHub Actions:

    pyinstaller --collect-all PySide6 Main.py

на рецепт tools/overtimetab.spec. Сам YAML на этой ветке
менять нельзя (у токена нет права на workflows), поэтому
make_resources.py ставит tools/ci_bin первым в PATH.

Аргументы командной строки намеренно игнорируются: в YAML
как раз --collect-all, из-за него zip раздувается до ~280 МБ.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SPEC = ROOT / "tools" / "overtimetab.spec"


def main() -> int:
    if not SPEC.is_file():
        print(f"[slim] нет файла {SPEC}", file=sys.stderr)
        return 1
    print("[slim] GitHub вызвал pyinstaller с --collect-all — собираем tools/overtimetab.spec")
    if len(sys.argv) > 1:
        print("[slim] исходные аргументы:", " ".join(sys.argv[1:]))
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        str(SPEC),
    ]
    return subprocess.call(cmd, cwd=str(ROOT))


if __name__ == "__main__":
    raise SystemExit(main())
