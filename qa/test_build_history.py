#!/usr/bin/env python3
"""Тест карты сборок (BUILD_HISTORY.md + tools/build_history.py).

Карта «сборка → что въехало» строится из истории git и живёт в
BUILD_HISTORY.md — в репозитории (видна на GitHub) и в каждой сборке.
Зачем: данные о прошлых сборках не должны зависеть от памяти
конкретного ассистента — файл прочитает любой человек или ИИ.
publish_release.py отказывается публиковать релиз, если текущей
сборки нет в карте.

Проверяется:
  A. Инструмент согласован с git: --check проходит;
  B. Карта покрывает историю: 152+ сборок, от 86-й до текущей;
  C. Текущая сборка (appBuild из AppTheme) описана в карте;
  D. Каждая пометка <!--b:N--> в CHANGELOG имеет раздел в карте;
  E. Защита публикации: history_has_build(текущая) = True,
     history_has_build(99999) = False.

Запуск:
    python3 qa/test_build_history.py             # из корня репозитория
"""
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))

from publish_release import history_has_build, read_identity


def main() -> int:
    hist_path = os.path.join(ROOT, "BUILD_HISTORY.md")
    assert os.path.exists(hist_path), "нет BUILD_HISTORY.md"
    hist = open(hist_path, encoding="utf-8").read()

    # ── A. Согласованность с git ──
    r = subprocess.run([sys.executable,
                        os.path.join(ROOT, "tools", "build_history.py"),
                        "--check"], capture_output=True, text=True, cwd=ROOT)
    assert r.returncode == 0, "--check провалился:\n" + r.stdout + r.stderr
    print("A: карта соответствует истории git ✓")

    # ── B. Покрытие ──
    builds = [int(m.group(1)) for m in re.finditer(r"^## Сборка (\d+)", hist, re.M)]
    assert len(builds) >= 152, "сборок в карте: %d" % len(builds)
    assert min(builds) <= 86 and max(builds) >= 241, (min(builds), max(builds))
    assert 241 in builds and 226 in builds and 210 in builds
    print("B: %d сборок, от %d до %d ✓" % (len(builds), min(builds), max(builds)))

    # ── C. Текущая сборка описана ──
    _, app_build = read_identity()
    assert history_has_build(app_build), "сборка %d отсутствует в карте" % app_build
    print("C: текущая сборка %d описана в карте ✓" % app_build)

    # ── D. Пометки чейнджлога имеют разделы ──
    changelog = open(os.path.join(ROOT, "CHANGELOG.md"), encoding="utf-8").read()
    tagged = {int(m.group(1)) for m in re.finditer(r"<!--b:(\d+)-->", changelog)}
    missing = sorted(b for b in tagged if ("## Сборка %d" % b) not in hist)
    assert not missing, "пометки без раздела в карте: %s" % missing
    print("D: все %d помеченных сборок имеют разделы в карте ✓" % len(tagged))

    # ── E. Защита публикации ──
    assert history_has_build(app_build) and not history_has_build(99999)
    print("E: защита публикации срабатывает (есть текущая, нет пустышки) ✓")

    print("═══ КАРТА СБОРОК: ГЕНЕРИРУЕТСЯ, ПОКРЫВАЕТ, ЗАЩИЩАЕТ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
