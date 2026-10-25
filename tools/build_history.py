#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Карта сборок: какая сборка что въехало (восстановление из истории git).

Каждый батч проекта — отдельный коммит с номером сборки в теме
(«…; сборка 231») и записью в CHANGELOG.md. Этот инструмент проходит
историю git и строит BUILD_HISTORY.md — карту «сборка → дата → коммит →
записи журнала». Она же ставит невидимые пометки сборок (<!--b:N-->)
в CHANGELOG.md для окна «Что нового».

Зачем: чтобы данные о прошлых сборках не жили только в памяти
ассистента — файл лежит в репозитории (виден на GitHub), попадает в
сборку и читается любым человеком или ИИ, который продолжит работу.

Запуск из корня репозитория:
    python tools/build_history.py --write      # пересобрать BUILD_HISTORY.md
    python tools/build_history.py --tag        # разметить CHANGELOG.md
    python tools/build_history.py --write --tag
    python tools/build_history.py --check      # выход 1, если файл устарел

Порядок при выпуске сборки: сначала закоммитить саму сборку
(запись в CHANGELOG + appBuild), затем запустить --write --tag
и закоммитить обновлённые файлы. publish_release.py откажется
публиковать релиз, если текущей сборки нет в BUILD_HISTORY.md.
"""
from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
HISTORY = ROOT / "BUILD_HISTORY.md"
CHANGELOG = ROOT / "CHANGELOG.md"

# «сборка 231» / «Сборка 212:» / «сборка 87» (именительный падеж)
_BUILD_RE = re.compile(r"[Сс]борка\s*[:№]?\s*(\d+)")
# коммиты version.json: «сборки 235» (родительный падеж)
_VJSON_BUILD_RE = re.compile(r"[Сc]борки\s+(\d+)")
# запись журнала: «- **Заголовок.**» (пометка сборки в конце не нужна)
_ENTRY_RE = re.compile(r"^\+- \*\*(.+?)\*\*")
_TAG_RE = re.compile(r"<!--\s*b\s*:\s*(\d+)\s*-->")


def git(*args: str) -> str:
    return subprocess.check_output(
        ["git", *args], cwd=str(ROOT), text=True, errors="replace")


def strip_tag(text: str) -> str:
    return _TAG_RE.sub("", text).strip()


def subject_build(subject: str) -> int:
    """Номер сборки из темы коммита (0 — не удалось)."""
    s = subject or ""
    if s.startswith("version.json"):
        m = _VJSON_BUILD_RE.search(s)
        if m:
            return int(m.group(1))
    m = _BUILD_RE.search(s)
    return int(m.group(1)) if m else 0


def collect() -> dict:
    """Проходит историю git (от старых к новым) и собирает карту сборок.

    Возвращает {build: {"date", "commits": [(sha, subject)], "titles": []}}.
    Запись журнала относится к той сборке, в чьём коммите она ВПЕРВЫЕ
    появилась (правило первого вхождения: позже текст могли править).
    """
    out: dict[int, dict] = {}
    title_to_build: dict[str, int] = {}

    log = git("log", "--reverse", "--format=%H%x09%ad%x09%s", "--date=short")
    for line in log.splitlines():
        parts = line.split("\t", 2)
        if len(parts) != 3:
            continue
        sha, date, subject = parts
        build = subject_build(subject)
        if not build:
            continue
        b = out.setdefault(build, {"date": date, "commits": [], "titles": []})
        if not b["commits"]:
            b["date"] = date
        b["commits"].append((sha[:7], subject.strip()))

        # записи журнала, добавленные этим коммитом
        diff = git("show", sha, "--format=", "--unified=0", "--", "CHANGELOG.md")
        for dl in diff.splitlines():
            if dl.startswith("+++") or not dl.startswith("+"):
                continue
            m = _ENTRY_RE.match(dl)
            if not m:
                continue
            title = strip_tag(m.group(1)).strip()
            if not title or title in title_to_build:
                continue
            title_to_build[title] = build
            if title not in b["titles"]:
                b["titles"].append(title)
    return out


def render(builds: dict) -> str:
    head = (
        "# История сборок OVERTIMETAB\n\n"
        "Карта «сборка → что въехло». Строится из истории git инструментом\n"
        "`tools/build_history.py --write --tag` (он же ставит невидимые пометки\n"
        "сборок `<!--b:N-->` в CHANGELOG.md для окна «Что нового»). Файл лежит\n"
        "в репозитории, виден на GitHub и попадает в каждую сборку — данные о\n"
        "прошлых сборках не зависят от памяти конкретного ассистента.\n\n"
        "Правило сопровождения: каждая новая сборка обязана быть описана здесь,\n"
        "иначе `tools/publish_release.py` не опубликует релиз. Порядок: коммит\n"
        "сборки → `python tools/build_history.py --write --tag` → коммит этих\n"
        "файлов → публикация.\n\n"
        "Сборки без записей журнала — служебные (правки сборочной обвязки и т.п.).\n"
        "Записи, добавленные до появления этой карты одним пакетом, остались\n"
        "без номера сборки и в окне «Что нового» не показываются.\n"
    )
    lines = [head]
    for build in sorted(builds, reverse=True):
        b = builds[build]
        sha = b["commits"][0][0] if b["commits"] else ""
        lines.append("\n## Сборка %d — %s — %s\n" % (build, b["date"], sha))
        for sha_i, subj in b["commits"]:
            lines.append("- коммит %s: %s" % (sha_i, subj))
        if b["titles"]:
            lines.append("- записи журнала:")
            for t in b["titles"]:
                lines.append("  - **%s**" % t)
        else:
            lines.append("- служебная сборка (записей в журнале нет)")
    return "\n".join(lines) + "\n"


def tag_changelog(builds: dict) -> int:
    """Ставит <!--b:N--> у записей CHANGELOG.md по карте. Возвращает число новых пометок."""
    with open(CHANGELOG, encoding="utf-8", newline="") as f:
        text = f.read()
    title_to_build = {}
    for build, b in builds.items():
        for t in b["titles"]:
            title_to_build.setdefault(t, build)

    lines = text.split("\n")
    tagged = 0
    unmatched = []
    for n, line in enumerate(lines):
        if not line.startswith("- **"):
            continue
        m = re.match(r"^- \*\*(.+?)\*\*", line)
        if not m:
            continue
        title = strip_tag(m.group(1)).strip()
        build = title_to_build.get(title)
        if build is None:
            unmatched.append(title)
            continue
        if "<!--b:" in line:
            continue
        lines[n] = line.rstrip() + " <!--b:%d-->" % build
        tagged += 1
    with open(CHANGELOG, "w", encoding="utf-8", newline="") as f:
        f.write("\n".join(lines))
    if unmatched:
        print("без номера сборки осталось %d записей (древние/переформулированные):"
              % len(unmatched))
        for t in unmatched[:8]:
            print("   - %s" % t[:70])
    return tagged


def main() -> int:
    ap = argparse.ArgumentParser(description="Карта сборок из истории git")
    ap.add_argument("--write", action="store_true", help="переписать BUILD_HISTORY.md")
    ap.add_argument("--tag", action="store_true", help="проставить пометки в CHANGELOG.md")
    ap.add_argument("--check", action="store_true",
                    help="выход 1, если BUILD_HISTORY.md не соответствует истории")
    args = ap.parse_args()

    if not (args.write or args.tag or args.check):
        ap.print_help()
        return 0

    builds = collect()
    body = render(builds)

    if args.check:
        if not HISTORY.exists():
            print("BUILD_HISTORY.md отсутствует — запустите --write")
            return 1
        with open(HISTORY, encoding="utf-8", newline="") as f:
            current = f.read()
        if current != body:
            print("BUILD_HISTORY.md устарел — запустите --write --tag и закоммитьте")
            return 1
        print("BUILD_HISTORY.md актуален: %d сборок" % len(builds))
        return 0

    if args.write:
        with open(HISTORY, "w", encoding="utf-8", newline="") as f:
            f.write(body)
        print("BUILD_HISTORY.md: %d сборок (от %s до %s)"
              % (len(builds), min(builds), max(builds)))
    if args.tag:
        n = tag_changelog(builds)
        print("CHANGELOG.md: новых пометок %d" % n)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
