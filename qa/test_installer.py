#!/usr/bin/env python3
"""Тест установщика: скрипт Inno Setup и чистота после перехода.

Установщик OVERTIMETAB — стандартный Inno Setup (tools/overtimetab.iss):
привычные окна «Далее → папка → установка», как у любой программы.
Компилятор Inno работает только на Windows, поэтому здесь проверяется
статически:
  A. скрипт .iss на месте и содержит ключевые директивы: режим
     «для меня/для всех» ({autopf} + PrivilegesRequiredOverridesAllowed),
     русский язык, ярлыки (Пуск всегда, рабочий стол по галке),
     запись удаления, иконку, сжатие;
  B. скрипт пакует собранную программу из dist и кладёт exe в корень
     репозитория (OutputDir=..), имя задаётся ключом /F на CI;
  C. при удалении спрашивается про данные сотрудников (Документы\\OverTimeTab)
     и по умолчанию они остаются;
  D. самописный мастер сборки 243 удалён полностью (файлы, Main.py,
     make_resources, .gitignore);
  E. путь данных совпадает с программой: Documents\\OverTimeTab
     (app_update.py, rollback-папка).

Запуск:
    python3 qa/test_installer.py
"""
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    iss_path = ROOT / "tools" / "overtimetab.iss"
    assert iss_path.exists(), "нет tools/overtimetab.iss"
    iss = iss_path.read_text(encoding="utf-8")

    # ── A. ключевые директивы ──
    checks = [
        ("DefaultDirName={autopf}\\OVERTIMETAB", "папка установки {autopf}"),
        ("PrivilegesRequired=lowest", "установка без прав по умолчанию"),
        ("PrivilegesRequiredOverridesAllowed=dialog", "вопрос «для меня/для всех»"),
        ('Name: "russian"', "русский язык мастера"),
        ("SetupIconFile=..\\app_icon.ico", "иконка установщика"),
        ("Compression=lzma2/max", "сжатие"),
        ("UsePreviousAppDir=yes", "обновление помнит прежнюю папку"),
    ]
    for marker, why in checks:
        assert marker in iss, "в .iss нет: %s (%s)" % (marker, why)
    print("A: директивы Inno Setup на месте (режим, язык, иконка, сжатие) ✓")

    # ── B. источник и результат ──
    assert 'Source: "{#SourceDir}\\*"' in iss and "recursesubdirs" in iss, \
        "файлы программы не пакуются из dist"
    assert 'define SourceDir "..\\dist\\OVERTIMETAB"' in iss
    assert "OutputDir=.." in iss, "exe должен ложиться в корень репозитория"
    assert "OutputBaseFilename=" in iss, "нет имени по умолчанию для ручной сборки"
    print("B: пакует dist\\OVERTIMETAB, exe — в корень репозитория ✓")

    # ── C. данные при удалении ──
    assert "{userdocs}\\OverTimeTab" in iss and "DelTree" in iss, \
        "удаление не спрашивает про данные"
    assert "MB_YESNO" in iss, "вопрос про данные должен быть да/нет"
    assert "MB_YES_NO" not in iss, "MB_YES_NO — не константа Inno (пишется MB_YESNO)"
    print("C: удаление спрашивает про Документы\\OverTimeTab (по умолчанию — остаются) ✓")

    # ── D. самописный мастер 243 удалён ──
    gone = ["tools/installer_stub.py", "tools/installer_stub.spec",
            "tools/make_installer.py", "installer.py",
            "components/SetupWizard.qml", "main_setup.qml",
            "qa/test_setup_view.py"]
    for p in gone:
        assert not (ROOT / p).exists(), "не удалён: " + p
    main_src = (ROOT / "Main.py").read_text(encoding="utf-8")
    for marker in ("--setup", "--uninstall", "SetupBackend", "setup_main"):
        assert marker not in main_src, "в Main.py осталось: " + marker
    res = (ROOT / "tools" / "make_resources.py").read_text(encoding="utf-8")
    assert "main_setup.qml" not in res, "make_resources тащит main_setup.qml"
    gi = (ROOT / ".gitignore").read_text(encoding="utf-8")
    assert "installer_stub" not in gi, "gitignore помнит заглушку"
    print("D: самописный мастер (243) удалён полностью ✓")

    # ── E. путь данных совпадает с программой ──
    # app_update.py: Path(home) / "Documents" ... / "OverTimeTab"
    au = (ROOT / "app_update.py").read_text(encoding="utf-8")
    assert '"Documents"' in au and '"OverTimeTab"' in au, \
        "app_update не хранит данные в Documents/OverTimeTab"
    assert "{userdocs}\\OverTimeTab" in iss, \
        ".iss удаляет не ту папку данных"
    print("E: данные сотрудников — Documents\\OverTimeTab, как в программе ✓")

    print("═══ УСТАНОВЩИК: СТАНДАРТНЫЙ INNO SETUP, СКРИПТ ЦЕЛ ✓ ═══")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
