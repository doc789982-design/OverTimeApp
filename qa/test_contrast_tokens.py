#!/usr/bin/env python3
"""Контраст цветовых токенов против WCAG AA (сборка 308).

Ревью по навыку better-colors (jakubkrehel/skills) нашло: приглушённые
подписи textTertiary не держали 4.5:1 (худшее — 2.69:1 на тёмном окне),
а красные тексты ошибок (accentDanger как цвет ТЕКСТА) — 4.2:1 на белых
карточках. Токены выправлены и появился отдельный textDanger для текста.

Этот тест не даёт токенам снова «поплыть»: каждая пара текст×поверхность,
которая реально рендерится, перемеряется по WCAG (относительная яркость
по sRGB). Правило: текст — 4.5:1, рамки фокуса — 3:1.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def lum(hexc: str) -> float:
    h = hexc.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    def f(c):
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def ratio(fg: str, bg: str) -> float:
    l1, l2 = sorted([lum(fg), lum(bg)], reverse=True)
    return (l1 + 0.05) / (l2 + 0.05)


def parse_tokens() -> dict:
    """{имя: {"dark": "#hex", "light": "#hex"}} из AppTheme.qml."""
    text = (ROOT / "components" / "AppTheme.qml").read_text(encoding="utf-8")
    tokens = {}
    for m in re.finditer(
            r'property\s+color\s+(\w+)\s*:\s*isDark\s*\?\s*"#([0-9A-Fa-f]{6})"\s*:\s*"#([0-9A-Fa-f]{6})"',
            text):
        tokens[m.group(1)] = {"dark": "#" + m.group(2), "light": "#" + m.group(3)}
    return tokens


def main() -> int:
    t = parse_tokens()
    need_names = {"textPrimary", "textSecondary", "textTertiary", "textDanger",
                  "textOnAccent", "bgBase", "bgPanel", "bgModal", "bgSurface",
                  "bgCell", "accentBrand", "accentSuccess", "accentDanger",
                  "borderFocus"}
    missing = need_names - set(t)
    assert not missing, "нет токенов в AppTheme.qml: %s" % sorted(missing)

    checks = []  # (описание, fg, bg, норма)

    # Текст на поверхностях, где он реально рендерится (обе темы)
    surfaces = ["bgBase", "bgPanel", "bgModal", "bgSurface", "bgCell"]
    for theme in ("light", "dark"):
        for text_tok in ("textPrimary", "textSecondary", "textTertiary",
                         "textDanger", "textBrand"):
            for surf in surfaces:
                checks.append((f"{theme} {text_tok} на {surf}",
                               t[text_tok][theme], t[surf][theme], 4.5))
        # акцент ЗАЛИВКОЙ как текст — только реальные пары: вкладки и пункты
        # списков живут на базовом фоне и всплывающих панелях
        for surf in ("bgBase", "bgElevated"):
            checks.append((f"{theme} accentBrand текстом на {surf}",
                           t["accentBrand"][theme], t[surf][theme], 4.5))
        # суммы приказов — бирюзовым на карточках
        checks.append((f"{theme} accentTeal текстом на bgElevated",
                       t["accentTeal"][theme], t["bgElevated"][theme], 4.5))
        # текст на акцентных заливках (кнопки, плашки)
        for acc in ("accentBrand", "accentSuccess", "accentDanger"):
            checks.append((f"{theme} textOnAccent на {acc}",
                           t["textOnAccent"][theme], t[acc][theme], 4.5))
        # рамка фокуса обязана быть видна (3:1 против границы/фона)
        for surf in ("bgBase", "bgModal", "bgPanel"):
            checks.append((f"{theme} borderFocus на {surf}",
                           t["borderFocus"][theme], t[surf][theme], 3.0))

    failed = []
    for name, fg, bg, need in checks:
        r = ratio(fg, bg)
        if r < need:
            failed.append(f"{name}: {r:.2f}:1 < {need}:1 ({fg} на {bg})")
    assert not failed, "пары ниже нормы:\n  " + "\n  ".join(failed)

    print(f"Контраст: {len(checks)} пар — все выше нормы "
          f"(текст 4.5:1, фокус 3:1), обе темы ✓")
    return 0


if __name__ == "__main__":
    sys.exit(main())
