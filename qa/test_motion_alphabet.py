#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ДИРИЖЁР: машинная проверка алфавита движения.

Правило канона (components/AppTheme.qml, раздел «АНИМАЦИИ И ФИЗИКА»):
длительности и кривые движения в QTL программы берутся ТОЛЬКО из
AppTheme (durMicro/durFast/durNormal/durStandard/durSlow/durStagger/
durShake, easeColor/easeEnter/easeExit/easeStandard, riseOffset,
slideOffset). Голые числа в duration: и прямые Easing.* допустимы
только в файлах-исключениях ниже — каждое исключение согласовано
отдельно (маскот, шторка выбора базы, радар, искры…).

Новый файл с голой длительностью или кривой мимо алфавита — падение
этого теста: либо возьми токен, либо обоснуй исключение и внеси файл
в реестр ниже с причиной.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# ── РЕЕСТР ИСКЛЮЧЕНИЙ (файл → причина; только согласанная пластика) ──
WHITELIST = {
    "main.qml":                  "хореография шторки выбора базы (600/400/180 мс, свои кривые) и шаг увертюры",
    "AppEmptyMascot.qml":        "живой маскот: фоновые петли дыхания 2–4 с",
    "AppEnterAnimation.qml":     "канонический вход сетки месяца: MD3 emphasized-decelerate 350 мс",
    "AppFieldBase.qml":          "медленное угасание подписи ошибки поля (700 мс, InQuad)",
    "AppSummaryPanel.qml":       "одометр итоговых часов (собственная механика перекатывания)",
    "EmployeeListPanel.qml":     "скольжения панели поиска и карточек (260/160, OutCubic)",
    "GroupSidebar.qml":          "переупорядочивание групп: переходы ListView move (180, OutCubic)",
    "AppSwitch.qml":             "механика тумблера (120, OutCubic)",
    "AppRipple.qml":             "волна нажатия (собственное затухание)",
    "AppCursorDelegate.qml":     "мигание текстового курсора",
    "ToastHost.qml":             "механика тостов: высадка и полёт",
    "AppDialog.qml":             "встряска окна: кривые Quad внутри шагов durShake",
    "AppDesktopNotification.qml": "выезд системного уведомления с отскоком",
    "AnimatedThemeButton.qml":   "фирменная кнопка темы: лучи и морф луны",
    "AnimatedFilterButton.qml":  "фирменная кнопка фильтра с отскоком",
    "AnimatedPrintButton.qml":   "фирменная кнопка печати",
    "AnimatedSettingsButton.qml": "фирменная кнопка настроек",
    "UpdateDownloadButton.qml":  "состояния кнопки обновления: кольцо и отскоки иконок",
    "UpdateRadar.qml":           "пульс радара обновления (1,6 с)",
    "HolidaySparkles.qml":       "искры праздников (фоновая петля)",
    "PreHolidaySparkle.qml":     "искра предпраздничного дня",
    "WhatsNewStage.qml":         "актёры виньеток «Что нового» и курсор-плеер",
    "WhatsNewVignette.qml":      "плеер тактов виньетки",
    "ThanosEffect.qml":          "распыление удалений",
    "ThemeTransition.qml":       "перекраска темы (медленная, согласована)",
    "LeftControlPanel.qml":      "волна по буквам логотипа (фоновая петля 2 с)",
    "SmoothFlickable.qml":       "физика инерции прокрутки",
    "SmoothListView.qml":        "физика инерции прокрутки",
}

# Токены, которые обязан иметь алфавит (сам AppTheme)
REQUIRED_TOKENS = [
    "durMicro", "durFast", "durNormal", "durStandard", "durSlow",
    "durStagger", "durShake",
    "easeColor", "easeEnter", "easeExit", "easeStandard",
    "riseOffset", "slideOffset", "scaleActive",
]

RE_DURATION = re.compile(r"duration\s*:\s*\d")
RE_EASING = re.compile(r"Easing\.[A-Z]")


def strip_comments(text: str) -> str:
    """Убирает // комментарии (многострочных блочных в QML почти нет,
    а в реестре выше они не фигурируют)."""
    return "\n".join(line.split("//", 1)[0] for line in text.splitlines())


def main() -> int:
    failures = []

    # 1. Алфавит на месте
    theme = (ROOT / "components" / "AppTheme.qml").read_text(encoding="utf-8")
    for token in REQUIRED_TOKENS:
        if not re.search(r"property\s+\w+\s+" + token + r"\s*:", theme):
            failures.append(f"AppTheme.qml: в алфавите нет токена {token}")

    # 2. Реестр не ссылается на несуществующие файлы
    for name in WHITELIST:
        if not (ROOT / "components" / name).exists() and name != "main.qml":
            failures.append(f"Реестр исключений ссылается на несуществующий файл: {name}")

    # 3. Голые длительности и кривые мимо алфавита
    qml_files = sorted(ROOT.glob("components/*.qml")) + [ROOT / "main.qml"]
    offenders = []
    for path in qml_files:
        rel = path.name
        if rel == "AppTheme.qml":
            continue
        body = strip_comments(path.read_text(encoding="utf-8"))
        hits = []
        for i, line in enumerate(body.splitlines(), 1):
            if rel in WHITELIST:
                continue
            if RE_DURATION.search(line):
                hits.append(f"  {rel}:{i}  голая длительность: {line.strip()}")
            if RE_EASING.search(line):
                hits.append(f"  {rel}:{i}  кривая мимо алфавита: {line.strip()}")
        offenders.extend(hits)

    if offenders:
        failures.append("Движение мимо алфавита (возьми токен AppTheme или обоснуй исключение):")
        failures.extend(offenders)

    if failures:
        print("═══ ДИРИЖЁР: НАРУШЕНИЯ АЛФАВИТА ═══")
        for f in failures:
            print(f)
        return 1

    total = len(WHITELIST)
    print(f"═══ ДИРИЖЁР: АЛФАВИТ ЧИСТ ═══")
    print(f"Токены алфавита на месте ({len(REQUIRED_TOKENS)} шт.), "
          f"движение мимо AppTheme отсутствует.")
    print(f"Реестр согласованных исключений: {total} файлов "
          "(маскот, шторка, радар, искры, физика жестов…).")
    return 0


if __name__ == "__main__":
    rc = main()
    if "-k" not in sys.argv:
        sys.exit(rc)
    sys.exit(0)
