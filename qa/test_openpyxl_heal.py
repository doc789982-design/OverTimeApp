#!/usr/bin/env python3
"""Тест самолечения numpy-конфликта openpyxl (utils.import_openpyxl).

На чужих машинах бывает битая/нестандартная numpy: openpyxl при импорте
обращается к numpy.short и падает с AttributeError — из-за этого ломаются
и Excel-экспорт («Для экспорта нужен openpyxl»), и печать
(«движок печати не смог: module 'numpy' has no attribute 'short'»).

Проверяется в отдельных процессах:
  1. воспроизведение: фальшивая numpy без short роняет обычный import openpyxl;
  2. лечение: utils.import_openpyxl() в тех же условиях импортирует openpyxl,
     numpy для openpyxl отключается (NUMPY=False), load_workbook работает;
  3. чистая среда: import_openpyxl() работает как обычный импорт.

Запуск:
    python3 qa/test_openpyxl_heal.py            # из корня репозитория
"""
import os
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

FAKE_NUMPY = "__version__ = '9.9.9-fake'\n# нет short/ushort/… — как битая установка\n"


def run(code, extra_env=None):
    env = dict(os.environ)
    if extra_env:
        env.update(extra_env)
    return subprocess.run([sys.executable, "-c", code], capture_output=True,
                          text=True, cwd=ROOT, env=env)


def main() -> int:
    fake_dir = tempfile.mkdtemp(prefix="fake_numpy_")
    os.makedirs(os.path.join(fake_dir, "numpy"), exist_ok=True)
    with open(os.path.join(fake_dir, "numpy", "__init__.py"), "w",
              encoding="utf-8") as f:
        f.write(FAKE_NUMPY)
    pp = {"PYTHONPATH": fake_dir}

    # 1) воспроизведение: обычный импорт openpyxl падает
    r = run("import openpyxl", pp)
    assert r.returncode != 0, "поломка не воспроизвелась — тест ничего не значит"
    assert "attribute 'short'" in r.stderr or "attribute \"short\"" in r.stderr, r.stderr[-300:]
    print("воспроизведение: import openpyxl падает (%s)" %
          r.stderr.strip().splitlines()[-1])

    # 2) лечение: import_openpyxl работает, numpy отключён для openpyxl
    r = run(
        "import sys; sys.path.insert(0, '.')\n"
        "from utils import import_openpyxl\n"
        "openpyxl = import_openpyxl()\n"
        "from openpyxl.compat.numbers import NUMPY\n"
        "assert not NUMPY, 'numpy должен быть отключён для openpyxl'\n"
        "from openpyxl import load_workbook\n"
        "print('ok', openpyxl.__version__)\n",
        pp)
    assert r.returncode == 0, r.stderr[-500:]
    print("лечение: import_openpyxl работает, numpy для openpyxl отключён")

    # 3) чистая среда: без numpy всё как обычно
    r = run(
        "import sys; sys.path.insert(0, '.')\n"
        "from utils import import_openpyxl\n"
        "m = import_openpyxl()\n"
        "import openpyxl\n"
        "assert openpyxl is m\n"
        "print('ok', openpyxl.__version__)\n")
    assert r.returncode == 0, r.stderr[-500:]
    print("чистая среда: import_openpyxl работает как обычный импорт")

    print("═══ NUMPY-КОНФЛИКТ: САМОЛЕЧЕНИЕ РАБОТАЕТ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
