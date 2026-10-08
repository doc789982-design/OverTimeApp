#!/usr/bin/env python3
"""Тест окна «Что нового» по сборкам (только свежие записи).

Сейчас после обновления окно показывало ВЕСЬ раздел текущей версии —
все записи за всю историю. Теперь каждая запись в CHANGELOG.md несёт
невидимую пометку сборки (<!--b:239-->), а окно показывает записи
только тех сборок, что появились после последней виденной:
обновление на соседнюю сборку — одна порция изменений; перепрыг через
несколько (230 → 239) — всё накопившееся, сгруппированное по сборкам.

Проверяется:
  A. Парсер: пометки читаются, из текста записей вырезаются;
  B. Интервалы: (238, 239] — одна сборка; (230, 239] — блоки без
     «пустых» сборок; (224, 240] — перепрыг через несколько сборок;
     ничего;
  C. Группировка: сверху новее, секции на месте (237 — «Починили»);
  D. Ключ старого формата «BETA.1+239» → 239 (конфиги прошлых сборок);
  E. Заметки релиза (publish_release.changelog_section) без пометок;
  F. Сводка для QML: одна порция на всё обновление — записи всех
     сборок интервала в общих разделах, без заголовков сборок;
     окно без «После обновления…» и с общим заголовком «Что нового
     в версии …»;
  G. Обновление без записей в журнале (правки внутренние) — окно
     показывает стандартную фразу, а не молчит.

Запуск:
    python3 qa/test_whats_new.py             # из корня репозитория
"""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))

import app_update
from publish_release import changelog_section, read_identity


def main() -> int:
    text = open(os.path.join(ROOT, "CHANGELOG.md"), encoding="utf-8").read()

    # ── A. Парсер ──
    bullets = app_update.parse_changelog_bullets(text)
    tagged = [b for b in bullets if b["build"] > 0]
    builds = sorted({b["build"] for b in tagged})
    _, cur_build = read_identity()          # текущая сборка — из AppTheme
    want = [210, 213, 214, 218, 219, 221, 222, 224, 225, 227, 228,
            231, 232, 233, 235, 236, 237, 238, 239]
    assert set(want) <= set(builds), "не хватает исторических сборок: %s" % (
        sorted(set(want) - set(builds)))
    # всё, что выше 240 — свежие батчи (242, 243, …): их может быть много
    extra = set(builds) - set(want)
    assert all(b > 240 for b in extra), "неожиданные сборки: %s" % sorted(extra)
    for b in tagged:
        assert "<!--" not in b["text"], "пометка осталась в тексте: %s" % b["text"][:60]
    assert any("За оба года" in b["text"] for b in tagged if b["build"] == 235)
    assert any("повреждённой numpy" in b["text"] for b in tagged if b["build"] == 225)
    # мета-записи (установщик, сам журнал, карта сборок) в окне не показываются
    for b in tagged:
        assert "установщ" not in b["text"].lower(), b["text"][:60]
    print("A: %d записей помечены (%d сборок, включая ранние), тексты чистые ✓"
          % (len(tagged), len(builds)))

    # ── B. Интервалы ──
    one = app_update.changelog_for_builds(text, 238, 239)
    assert len(one) == 1 and one[0]["build_num"] == 239, one
    assert any("увольнения" in t for t in one[0]["added"]), one[0]["added"]
    many = app_update.changelog_for_builds(text, 230, 239)
    assert [b["build_num"] for b in many] == \
        [239, 238, 237, 236, 235, 233, 232, 231], [b["build_num"] for b in many]
    # перепрыг с 224-й: мелкие сборки (226, 230, 234) записей не имеют —
    # их блоков нет, бабушка видит только настоящие изменения
    jump = app_update.changelog_for_builds(text, 224, 240)
    assert [b["build_num"] for b in jump] == \
        [239, 238, 237, 236, 235, 233, 232, 231, 228, 227, 225], \
        [b["build_num"] for b in jump]
    b225 = [b for b in jump if b["build_num"] == 225][0]
    assert any("numpy" in x for sec in ("added", "changed", "fixed")
               for x in b225[sec]), b225
    allv = app_update.changelog_for_builds(text, 0, cur_build)
    assert len(allv) == len(set(builds)), (len(allv), len(set(builds)))
    assert app_update.changelog_for_builds(text, 239, 239) == []
    # интервал, где записей точно нет (за пределами всех сборок)
    assert app_update.changelog_for_builds(text, 400, 500) == []
    print("B: 238→239 один; 230→239 восемь; 224→240 — одиннадцать блоков ✓")

    # ── C. Группировка и секции ──
    b237 = [b for b in many if b["build_num"] == 237][0]
    assert any("Вкладки месяцев" in t for t in b237["fixed"]), b237["fixed"]
    b239 = many[0]
    assert any("увольнения и перевода" in t for t in b239["added"])
    assert many[0]["build_num"] > many[1]["build_num"], "сверху новее"
    print("C: сверху новее, секции на месте (237 — Починили) ✓")

    # ── D. Ключ старого формата ──
    assert app_update.build_from_version_key("BETA.1+239") == 239
    assert app_update.build_from_version_key("BETA.1") == 0
    assert app_update.build_from_version_key("") == 0
    assert app_update.build_from_version_key("2.0.0-ALPHA.20+228") == 228
    print("D: старый ключ «ВЕРСИЯ+сборка» читается ✓")

    # ── E. Заметки релиза без пометок ──
    notes = changelog_section("BETA.1")
    assert "<!--" not in notes, "пометка попала в текст релиза"
    assert "За оба года" in notes
    assert app_update.strip_build_tags("текст <!--b:239--> конец") == "текст  конец"
    print("E: заметки релиза чистые, strip_build_tags работает ✓")

    # ── F. Сводка для QML: одна порция на всё обновление ──
    jump = app_update.changelog_for_builds(text, 224, 240)
    summary = app_update.whats_new_qml(jump)
    assert len(summary) == 1, "сводка должна быть одна, а не по сборкам"
    s = summary[0]
    # записи разных сборок интервала — в общих разделах
    assert "За оба года" in s["addedText"] and "Выборочная печать" in s["addedText"]
    assert "numpy" in s["addedText"], s["addedText"][:200]
    assert not s["version"], "заголовков сборок больше нет"
    single = app_update.whats_new_qml(one)
    assert "увольнения" in single[0]["addedText"]
    # окно: общий заголовок с версией, без «После обновления…»
    qml_src = open(os.path.join(ROOT, "components", "WhatsNewDialog.qml"),
                   encoding="utf-8").read()
    assert "После обновления" not in qml_src, "строка «После обновления…» вернулась"
    assert "Что нового в версии" in qml_src and "appVersionFull" in qml_src
    assert 'text: "Версия "' not in qml_src, "заголовки блоков версий вернулись"
    print("F: сводка единая (224→240 в четырёх разделах), окно чистое ✓")

    # ── G. Обновление без записей — стандартная фраза ──
    synthetic = "### Поменяли\n- **А.** Одна запись. <!--b:300-->\n"
    assert app_update.changelog_for_builds(synthetic, 300, 305) == [], \
        "в интервале (300, 305] записей быть не должно"
    fb = app_update.whats_new_fallback()
    assert len(fb) == 1, "фраза должна быть одной порцией"
    assert fb[0]["hasChanged"] and not fb[0]["hasAdded"], fb[0]
    assert "Мелкие улучшения" in fb[0]["changedText"], fb[0]["changedText"]
    assert not fb[0]["version"], "у фразы не должно быть заголовка сборки"
    # Main применяет фразу, но только к свежему (непрочитанному) обновлению
    main_src = open(os.path.join(ROOT, "Main.py"), encoding="utf-8").read()
    assert "whats_new_fallback()" in main_src, \
        "Main не показывает стандартную фразу пустого обновления"
    assert "_whats_new_fresh" in main_src, \
        "фраза обязана появляться только у свежего обновления"
    print("G: пустое обновление → стандартная фраза ✓")

    print("═══ «ЧТО НОВОГО» ПОКАЗЫВАЕТ ТОЛЬКО СВЕЖИЕ СБОРКИ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
