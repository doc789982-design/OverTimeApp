#!/usr/bin/env python3
"""Восстанавливает /tmp/stubs — стабы системных библиотек для Qt в песочнице.

Песочница Arena пересоздаётся: пропадают PySide6, системные библиотеки и /tmp.
Здесь собираются пустышки (gcc -shared) для библиотек, которых нет в системе,
но которые тянет PySide6. Символы берутся из фактических импортов библиотек Qt
(objdump -T): стаб экспортирует их с нужными версиями, тела — пустые.
Offscreen-платформа эти библиотеки по-настоящему не вызывает.

Использование:
    pip install --break-system-packages PySide6   # если модуля нет
    python3 tools/make_sandbox_libs.py            # создаст /tmp/stubs
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen python3 <тест>
"""
import re
import subprocess
import sys
from pathlib import Path

STUB_DIR = Path("/tmp/stubs")

PYSIDE = Path("/usr/local/lib/python3.11/dist-packages/PySide6")

# имя файла-стаба → (префиксы символов, версии из objdump)
STUBS = {
    "libdbus-1.so.3": (["dbus_"], ["LIBDBUS_1_3"]),
    "libEGL.so.1": (["egl", "EGL"], []),
    "libGL.so.1": (["gl", "GL"], []),
    "libxkbcommon.so.0": (["xkb_"], []),
}


def qt_libraries():
    libs = list((PYSIDE / "Qt" / "lib").glob("libQt6*.so.6"))
    libs += list(PYSIDE.glob("*.abi3.so"))
    libs += list((PYSIDE / "Qt" / "plugins").rglob("*.so"))
    return libs


def undefined_symbols():
    """{(префикс): {(символ, версия)}} по всем UND-символам библиотек Qt."""
    found = {name: set() for name in STUBS}
    ver_re = re.compile(r"\(([\w.]+)\)\s+(\S+)$")
    for lib in qt_libraries():
        r = subprocess.run(["objdump", "-T", str(lib)], capture_output=True, text=True)
        for ln in r.stdout.splitlines():
            if "UND" not in ln:
                continue
            m = ver_re.search(ln)
            if m:                       # «(ВЕРСИЯ) символ»
                sym, ver = m.group(2), m.group(1)
            else:                       # без версии: последний токен строки
                parts = ln.split()
                sym, ver = (parts[-1], "") if len(parts) >= 5 else ("", "")
            if not sym:
                continue
            for name, (prefixes, _) in STUBS.items():
                if sym.startswith(tuple(prefixes)):
                    found[name].add((sym, ver))
    return found


def build_stub(name, symbols):
    """Собирает один стаб: пустые тела + version-script с версиями."""
    src = STUB_DIR / (name + ".c")
    ver_file = STUB_DIR / (name + ".ver")
    src.write_text("".join("void %s(void) {}\n" % s for s, _ in sorted(symbols)),
                   encoding="utf-8")
    # группируем символы по версиям
    groups = {}
    for s, v in symbols:
        groups.setdefault(v or "__none__", []).append(s)
    script = ""
    for v, ss in groups.items():
        if v == "__none__":
            # анонимный узел (базовая версия): без имени перед скобкой
            script += "{\n  global:\n"
            script += "".join("    %s;\n" % s for s in sorted(ss))
            script += "  local: *;\n};\n"
        else:
            script += "%s {\n  global:\n" % v
            script += "".join("    %s;\n" % s for s in sorted(ss))
            script += "  local: *;\n};\n"
    ver_file.write_text(script, encoding="utf-8")
    out = STUB_DIR / name
    r = subprocess.run(
        ["gcc", "-shared", "-fPIC", "-o", str(out), str(src),
         "-Wl,--version-script=" + str(ver_file)],
        capture_output=True, text=True,
    )
    if r.returncode != 0:
        print("не собрали %s: %s" % (name, r.stderr.strip()), file=sys.stderr)
        return False
    return True


def main() -> int:
    STUB_DIR.mkdir(parents=True, exist_ok=True)
    found = undefined_symbols()
    for name, symbols in found.items():
        if not symbols:
            print("стаб %s: не потребовался" % name)
            continue
        if build_stub(name, symbols):
            print("стаб %s: %d символов" % (name, len(symbols)))
    print("готово: LD_LIBRARY_PATH=%s" % STUB_DIR)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
