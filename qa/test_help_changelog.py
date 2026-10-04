#!/usr/bin/env python3
"""Сборка 267: чейнджлог в справке, «Деньгами» → журнал, редактор группы.

  A. Справка (HelpDialog): карточка «Что нового» с кнопкой —
     клик даёт сигнал requestWhatsNew (main.qml ведёт к окну).
  B. Окно «Что нового» (WhatsNewDialog): showChangelog() открывает
     ВЕСЬ список изменений версии (backend.whatsNewAll), даже когда
     «после обновления» уже видели (whatsNew пуст); закрытие
     сбрасывает режим и отмечает просмотр.
  C. Группа (AddGroupDialog): режим редактирования — editGroup
     заполняет поля, титул «Редактирование группы», кнопка
     «Сохранить», сохранение зовёт backend.updateGroup; создание —
     как раньше, backend.createGroup.
  D. Исходники: ПКМ по группе — «Редактировать группу» вместо
     тумблера выходных; «Деньгами» в карточке — штатная кнопка
     (AppButton), открывающая журнал приказов.

Запуск (песочница):
    LD_LIBRARY_PATH=/tmp/stubs QT_QPA_PLATFORM=offscreen \
    QSG_RASTER_BACKEND=1 OVERTIMETAB_SANDBOX_FONTS=1 \
    python3 qa/test_help_changelog.py
"""
import os
import sys
import time

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
os.environ.setdefault("QSG_RASTER_BACKEND", "1")
os.environ.setdefault("OVERTIMETAB_SANDBOX_FONTS", "1")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtCore import QObject, Slot, Property
from PySide6.QtGui import QGuiApplication
from PySide6.QtQml import QQmlApplicationEngine, QJSValue

WHATS_NEW_ALL = [{
    "addedText": "**Пробная запись.** Проверка полного списка изменений.",
    "fixedText": "",
}]


class StubBackend(QObject):
    def __init__(self):
        super().__init__()
        self.calls = []

    # WhatsNewDialog (свойства — их читает QML без скобок)
    @Property(list, constant=True)
    def whatsNew(self):
        return []          # «после обновления» уже видели

    @Property(list, constant=True)
    def whatsNewAll(self):
        self.calls.append("whatsNewAll")
        return WHATS_NEW_ALL

    @Slot()
    def ackWhatsNew(self):
        self.calls.append("ack")

    # AddGroupDialog
    @Slot(str, bool, bool)
    def createGroup(self, name, is_shift, shifted):
        self.calls.append(("create", name, is_shift, shifted))
        return None

    @Slot(int, str, bool, bool)
    def updateGroup(self, gid, name, is_shift, shifted):
        self.calls.append(("update", gid, name, is_shift, shifted))
        return None


def to_var(v):
    return v.toVariant() if isinstance(v, QJSValue) else v


def settle(app, sec=0.5):
    deadline = time.time() + sec
    while time.time() < deadline:
        app.processEvents()
        time.sleep(0.02)


def find_by(root, cls=None, object_name=None, text=None, label=None):
    out = []
    for o in root.findChildren(QObject):
        if cls and not o.metaObject().className().startswith(cls):
            continue
        if object_name is not None and str(o.property("objectName") or "") != object_name:
            continue
        if text is not None and str(o.property("text") or "") != text:
            continue
        if label is not None and str(o.property("label") or "") != label:
            continue
        out.append(o)
    return out


def main() -> int:
    app = QGuiApplication([])
    stub = StubBackend()
    engine = QQmlApplicationEngine()
    ctx = engine.rootContext()
    ctx.setContextProperty("backend", stub)
    wrapper = os.path.join(ROOT, "_render_help_changelog.qml")
    with open(wrapper, "w", encoding="utf-8") as f:
        f.write("""
import QtQuick
import QtQuick.Controls
import "components" as AppUI
ApplicationWindow {
    id: w
    width: 900; height: 700; visible: true
    AppUI.HelpDialog      { id: helpDialog }
    AppUI.WhatsNewDialog  { id: whatsNewDialog }
    AppUI.AddGroupDialog  { id: addGroupDialog }
}
""")
    engine.load(wrapper)
    assert engine.rootObjects(), "окна не загрузились"
    settle(app, 1.2)
    # при запуске файлом rootObjects[0] бывает голым QWindow —
    # берём настоящее QQuickWindow из topLevelWindows
    from PySide6.QtQuick import QQuickWindow
    win = next(x for x in app.topLevelWindows()
               if isinstance(x, QQuickWindow))

    help_dlg = find_by(win, "HelpDialog")[0]
    new_dlg = find_by(win, "WhatsNewDialog")[0]
    group_dlg = find_by(win, "AddGroupDialog")[0]

    # ── A. справка: карточка «Что нового» и кнопка ──
    texts = [str(o.property("text") or "")
             for o in help_dlg.findChildren(QObject) if o.property("text")]
    assert any(t == "Что нового" for t in texts), "нет карточки «Что нового»"
    btns = find_by(help_dlg, "AppButton", object_name="whatsNewButton")
    assert btns, "нет кнопки чейнджлога в справке"
    fired = []
    help_dlg.requestWhatsNew.connect(lambda: fired.append(1))
    from PySide6.QtCore import QMetaObject
    QMetaObject.invokeMethod(btns[0], "clicked")
    settle(app, 0.2)
    assert fired, "кнопка не сигналит requestWhatsNew"
    print("A. Справка: карточка «Что нового», кнопка сигналит ✓")

    # ── B. «Что нового»: весь чейнджлог по кнопке ──
    assert not to_var(new_dlg.property("visible"))
    QMetaObject.invokeMethod(new_dlg, "showChangelog")
    settle(app, 1.5)
    assert to_var(new_dlg.property("visible")), "окно не открылось"
    assert new_dlg.property("showAll") is True, "не включён режим полного списка"
    assert any(c == "whatsNewAll" for c in stub.calls), "не читан полный список"
    # делегаты Repeater ищем обходом item-дерева: findChildren
    # (QObject) их не видит (попап + Repeater)
    from PySide6.QtQuick import QQuickItem

    def walk_texts(item, acc):
        if item.metaObject().className() == "QQuickText":
            acc.append(str(item.property("text") or ""))
        for c in item.childItems():
            walk_texts(c, acc)

    body = []
    roots = [ri for ri in win.findChildren(QQuickItem)
             if ri.parentItem() is None]
    assert roots, "нет корневых айтемов — окно не построилось"
    for ri in roots:
        walk_texts(ri, body)
    assert any("Пробная запись" in t for t in body), \
        "запись полного списка не показана: %r" % (body[:6],)
    QMetaObject.invokeMethod(new_dlg, "close")
    settle(app, 1.0)
    assert any(c == "ack" for c in stub.calls), "закрытие не отметило просмотр"
    assert new_dlg.property("showAll") is False, "режим полного списка не сброшен"
    print("B. «Что нового»: полный список открылся и закрылся корректно ✓")

    # ── C. группа: редактирование ──
    assert str(group_dlg.property("title")) == "Новая группа"
    assert str(group_dlg.property("acceptText")) == "Создать"
    from PySide6.QtCore import Q_ARG
    QMetaObject.invokeMethod(group_dlg, "editGroup",
                             Q_ARG("QVariant", 5), Q_ARG("QVariant", "Смена 2"),
                             Q_ARG("QVariant", True), Q_ARG("QVariant", True))
    settle(app, 0.5)
    assert int(group_dlg.property("editGroupId")) == 5
    assert str(group_dlg.property("title")) == "Редактирование группы"
    assert str(group_dlg.property("acceptText")) == "Сохранить"
    name_f = find_by(group_dlg, "AppTextField", label="Название группы")[0]
    assert str(name_f.property("text")) == "Смена 2", name_f.property("text")
    switches = find_by(group_dlg, "AppSwitch")
    assert len(switches) == 2, "в окне нет двух тумблеров (график, выходные)"
    assert all(s.property("checked") is True for s in switches), \
        "editGroup не включил тумблеры"
    # смена имени и сохранение → updateGroup
    name_f.setProperty("text", "Смена 2А")
    QMetaObject.invokeMethod(group_dlg, "accepted")
    settle(app, 0.3)
    upd = [c for c in stub.calls if isinstance(c, tuple) and c[0] == "update"]
    assert upd and upd[0][1:] == (5, "Смена 2А", True, True), stub.calls
    # создание по-прежнему зовёт createGroup
    group_dlg.setProperty("editGroupId", 0)
    QMetaObject.invokeMethod(group_dlg, "open")
    settle(app, 0.3)
    name_f = find_by(group_dlg, "AppTextField", label="Название группы")[0]
    name_f.setProperty("text", "Ночная")
    QMetaObject.invokeMethod(group_dlg, "accepted")
    settle(app, 0.3)
    cre = [c for c in stub.calls if isinstance(c, tuple) and c[0] == "create"]
    assert cre and cre[0][1:] == ("Ночная", False, False), stub.calls
    print("C. Группа: редактирование обновляет, создание создаёт ✓")

    # ── D. исходники: меню группы и «Деньгами» ──
    sidebar = open(os.path.join(ROOT, "components", "GroupSidebar.qml"),
                   encoding="utf-8").read()
    assert "Редактировать группу" in sidebar, "нет пункта редактирования группы"
    assert 'text: "Смещённые выходные"' not in sidebar and \
        'text: "Обычные выходные"' not in sidebar, \
        "тумблер выходных не убран из меню"
    assert "setGroupShiftedWeekends" not in sidebar, \
        "меню всё ещё дёргает setGroupShiftedWeekends"
    assert "addGroupDialog.editGroup(" in sidebar, "меню не зовёт редактор"
    panel = open(os.path.join(ROOT, "components", "AppSummaryPanel.qml"),
                 encoding="utf-8").read()
    i_btn = panel.find('objectName: "moneyOrderBtn"')
    assert i_btn > 0, "у «Деньгами» нет objectName"
    assert panel.rfind("AppButton {", 0, i_btn) > panel.rfind("Rectangle {", 0, i_btn), \
        "«Деньгами» — не штатная кнопка"
    assert "ruble.svg" in panel and "moneyInspector.show()" in panel
    main_qml = open(os.path.join(ROOT, "main.qml"), encoding="utf-8").read()
    assert "onRequestWhatsNew: whatsNewDialog.showChangelog()" in main_qml, \
        "кнопка справки не подключена к окну чейнджлога"
    print("D. Меню группы и «Деньгами» ведут куда надо ✓")

    os.remove(wrapper)
    print("═══ ЧЕЙНДЛОГ В СПРАВКЕ · ДЕНЬГАМИ→ЖУРНАЛ · РЕДАКТОР ГРУППЫ ✓ ═══")
    return 0


if __name__ == "__main__":
    sys.exit(main())
