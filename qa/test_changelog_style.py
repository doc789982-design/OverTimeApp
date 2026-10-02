#!/usr/bin/env python3
"""Тест правил журнала изменений (CHANGELOG.md + tools/changelog_style.py).

Правила живут в начале CHANGELOG.md («Как писать записи») и действуют
для любого, кто ведёт журнал, — человека или ИИ. Выпуск сборки с
кривой записью невозможен: tools/publish_release.py проверяет стиль
перед публикацией.

Проверяется:
  A. правила описаны в шапке CHANGELOG.md (жаргон под запретом,
     лимит длины, пометки сборок, ссылка на проверяльщик);
  B. все записи текущего журнала проходят проверку стиля;
  C. проверяльщик ловит нарушения: жаргон, длину, число предложений;
  D. publish_release.py отказывает в публикации при плохом стиле
     (в коде есть вызов changelog_style_ok до создания релиза).

Запуск:
    python3 qa/test_changelog_style.py
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

from changelog_style import DENY_WORDS, MAX_LEN, entry_problems, text_problems


def main() -> int:
    text = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")

    # ── A. правила в шапке ──
    head = text[:text.find("### Добавили")]
    for marker in ("Как писать записи", "обычный человек", "<!--b:N-->",
                   "changelog_style", "соседней сборке", "установщике"):
        assert marker in head, "в шапке CHANGELOG.md нет: " + marker
    for word in ("релиз", "мастер"):
        assert word in head, "шапка должна называть запрещённые слова"
    # журнал плоский: единственный ## — правила, заголовков версий нет
    h2 = [l for l in text.split("\n") if l.startswith("## ")]
    assert h2 == ["## Как писать записи"], "лишние заголовки: %s" % h2
    assert "Предыдущие версии" not in text and "## BETA" not in text
    print("A: правила в шапке, журнал плоский — без заголовков версий ✓")

    # ── B. текущий журнал чист ──
    problems = text_problems(text)
    assert not problems, "\n".join(problems)
    n = sum(1 for line in text.split("\n") if line.startswith("- **")
            and "<!--b:" in line)
    assert n >= 40, "записей с пометками подозрительно мало: %d" % n
    print("B: все %d записей проходят проверку стиля ✓" % n)

    # ── C. ловит нарушения ──
    cases = [
        ("Мастер установки качает релиз с GitHub.", "жаргон"),
        ("Слово " * 90 + "и ещё немного в конце.", "длина"),
        ("Раз. Два. Три. Четыре. Пять.", "предложения"),
    ]
    for body, why in cases:
        assert entry_problems("X", body), "не поймал: " + why
    assert entry_problems("X", " Окно стало понятнее.") == []
    # мета-темы ловятся и в заголовке
    assert entry_problems("Установщик программы", " Ставит программу.") != []
    assert entry_problems("X", " Записи в журнале изменений переписаны.") != []
    # каждое запрещённое слово ловится
    for w in DENY_WORDS:
        assert text_problems("- **X.** Есть %s здесь. <!--b:1-->" % w), w
    print("C: жаргон, длина и число предложений ловятся (%d слов в списке) ✓"
          % len(DENY_WORDS))

    # ── D. ворота публикации ──
    pub = (ROOT / "tools" / "publish_release.py").read_text(encoding="utf-8")
    assert "def changelog_style_ok" in pub and "changelog_style_ok()" in pub, \
        "publish_release не проверяет стиль журнала"
    # вызов стоит ДО публикации (до создания/редактирования релиза)
    assert pub.find("changelog_style_ok()") < pub.find("release_exists(tag)"), \
        "проверка стиля должна срабатывать до работы с релизом"
    print("D: публикация блокируется при плохом стиле ✓")

    print("═══ ЖУРНАЛ ИЗМЕНЕНИЙ: ПРАВИЛА ЖИВУТ В РЕПОЗИТОРИИ И ПРОВЕРЯЮТСЯ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
