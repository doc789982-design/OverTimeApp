#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Проверка стиля записей журнала изменений (CHANGELOG.md).

Правила — в начале самого CHANGELOG.md (раздел «Как писать записи»).
Коротко: обычные слова без программистских терминов, одно–три
предложения, не длиннее 420 знаков, пометка сборки в конце строки.

Проверку использует tools/publish_release.py: выпуск сборки с кривой
записью невозможен. Правила живут в репозитории и действуют для
любого, кто ведёт журнал, — человека или ИИ, — независимо от памяти
конкретного собеседника.

Запуск (проверить журнал прямо сейчас):
    python3 tools/changelog_style.py
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CHANGELOG = ROOT / "CHANGELOG.md"

# Максимальная длина записи (знаков) и число предложений
MAX_LEN = 420
MAX_SENTENCES = 3

# Слова, которых не должно быть в записях: программистский жаргон,
# непонятный обычному человеку, и мета-темы (записи о самом журнале,
# установщике, внутреннем устройстве — см. правила в CHANGELOG.md).
# Сравнение без учёта регистра, по подстроке.
DENY_WORDS = [
    # программистский жаргон
    "релиз", "мастер", "ассет", "репозитор", "коммит", "воркфлоу",
    "workflow", "гибрид", "заглушк", "бэкенд", "фронтенд", "пайплайн",
    "ассистент", "innosetup", "inno setup", "pyinstaller", "github",
    "api", "виджет", "фич", "деплой",
    # мета-темы и внутреннее устройство
    "установщик", "чейнджлог", "журнал изменений", "журнала измен",
    "журнале измен", "журналу измен", "карта сборок",
    "build_history", "кэш", "кэшир", "рендер", "отрисов", "манифест",
    "миграци", "тени",
]

_ENTRY_RE = re.compile(r"^- \*\*(.+?)\*\*(.*?)<!--b:(\d+)-->$")


def entry_problems(title: str, body: str) -> list[str]:
    """Проблемы одной записи (пустой список — запись в порядке)."""
    out = []
    text = (body or "").strip()
    low = (title + " " + text).lower()   # заголовок проверяем вместе с текстом
    for w in DENY_WORDS:
        if w in low:
            out.append("слово «%s» — жаргон или мета-тема, подберите обычное" % w)
    if len(text) > MAX_LEN:
        out.append("длина %d знаков (лимит %d) — сократите" % (len(text), MAX_LEN))
    sentences = [s for s in re.split(r"(?<=[.!?])\s+", text) if s.strip()]
    if len(sentences) > MAX_SENTENCES:
        out.append("предложений %d (лимит %d) — оставьте суть" % (len(sentences), MAX_SENTENCES))
    return out


def text_problems(changelog_text: str) -> list[str]:
    """Проблемы всех записей файла: список строк «сборка N · заголовок: …»."""
    out = []
    for line in changelog_text.split("\n"):
        m = _ENTRY_RE.match(line)
        if not m:
            continue
        title, body, build = m.group(1), m.group(2), m.group(3)
        for p in entry_problems(title, body):
            out.append("сборка %s · %s: %s" % (build, title, p))
    return out


def main() -> int:
    text = CHANGELOG.read_text(encoding="utf-8")
    problems = text_problems(text)
    if problems:
        print("Журнал не проходит проверку стиля (правила — в начале CHANGELOG.md):")
        for p in problems:
            print("  - " + p)
        return 1
    n = sum(1 for line in text.split("\n") if _ENTRY_RE.match(line))
    if n == 0:
        print("НЕ НАЙДЕНО ни одной записи — проверка ничего не проверила!",
              file=sys.stderr)
        return 1
    print("стиль записей в порядке: %d записей проверено" % n)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
