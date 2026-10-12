import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

// ============================================================
// ДИАЛОГ-ФОРМА — базовое окно программы: шапка, прокручиваемый
// контент, подвал с кнопками. Механика появления/затемнения/
// страховки живёт в AppModalShell. Дочерние элементы падают
// в прокручиваемую зону (default property dialogContent).
// ============================================================
AppModalShell {
    id: root

    // --- Настройки интерфейса окна ---
    property string title: "Заголовок"
    property string acceptText: "Сохранить"
    property string rejectText: "Отмена"
    property string acceptVariant: "primary"
    property string rejectVariant: "secondary"
    property bool showFooter: true
    property bool showAccept: true
    property bool showReject: true

    // --- Сигналы ---
    signal accepted()
    signal rejected()

    // Подтверждение закрывает окно сразу (окна «Вы уверены?»);
    // формы по умолчанию закрывают себя сами в onAccepted
    property bool closeOnAccept: false
    function _accept() {
        root.accepted()
        if (root.closeOnAccept) root.close()
    }

    // Зона для вставки контента
    default property alias dialogContent: contentArea.data

    width: 380
    // По умолчанию окно подстраивает высоту под содержимое (не выше экрана).
    // Если задать heightFraction > 0 — фиксированная высота = доля высоты экрана.
    property real heightFraction: 0
    property real effectiveHeight: heightFraction > 0
        ? Math.max(200, (ApplicationWindow.window ? ApplicationWindow.window.height : 720) * heightFraction)
        : Math.min(mainLayout.implicitHeight + 40, ApplicationWindow.window ? ApplicationWindow.window.height - 100 : 800)
    height: effectiveHeight

    // Тряска при ошибке
    property real baseShakeX: 0
    function shake() { if (!shakeAnimation.running) { baseShakeX = root.x; shakeAnimation.start() } }

    // Прокрутка содержимого к низу: сообщения об ошибках живут в конце
    // колонки, и без прокрутки окно трясётся «непонятно из-за чего» —
    // текст остаётся за нижним краем.
    function scrollToBottom() {
        let fl = scrollArea.contentItem
        if (fl) fl.contentY = Math.max(0, fl.contentHeight - fl.height)
    }

    // Прокрутка к конкретному элементу (сообщению об ошибке): окно
    // показывает именно его — и если он выше края, и если ниже.
    // Позиции полей узнаём ПОСЛЕ раскладки: в момент клика высоты могут
    // быть ещё не пересчитаны, и адрес прокрутки выйдет устаревшим.
    // Поэтому пробуем следующим кадром (Qt.callLater) и подстраховываем
    // коротким таймером — для окна это мгновенно.
    function _scrollItemIntoView(item) {
        let fl = scrollArea.contentItem
        if (!fl || !item) return
        let pos = item.mapToItem(fl, 0, 0).y
        let h = item.height
        if (pos < fl.contentY) {
            fl.contentY = Math.max(0, pos - AppTheme.spaceM)
        } else if (pos + h > fl.contentY + fl.height) {
            fl.contentY = Math.max(0, pos + h - fl.height + AppTheme.spaceM)
        }
    }

    Timer {
        id: scrollRetry
        interval: 120
        repeat: false
        property var target: null
        onTriggered: if (target) root._scrollItemIntoView(target)
    }

    function scrollToItem(item) {
        if (!item) return
        // Обычный случай — раскладка давно готова: скроллим сразу,
        // а следующим кадром и коротким таймером перепроверяем
        // (на случай, если высоты полей ещё пересчитывались).
        root._scrollItemIntoView(item)
        Qt.callLater(function() { root._scrollItemIntoView(item) })
        scrollRetry.target = item
        scrollRetry.restart()
    }
    SequentialAnimation {
        id: shakeAnimation
        NumberAnimation { target: root; property: "x"; to: baseShakeX + 10; duration: 50; easing.type: Easing.OutQuad }
        NumberAnimation { target: root; property: "x"; to: baseShakeX - 10; duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: root; property: "x"; to: baseShakeX + 8;  duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: root; property: "x"; to: baseShakeX - 8;  duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: root; property: "x"; to: baseShakeX + 4;  duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: root; property: "x"; to: baseShakeX;      duration: 50; easing.type: Easing.OutQuad }
    }

    contentItem: ColumnLayout {
        // Фокус сюда: активным его делает попап (обёртка QQuickPopupItem
        // держит фокус-скоуп), иначе Keys не увидят клавиши.
        // Enter — «Сохранить» (когда фокус не в поле ввода: там Enter
        // завершает ввод поля, как принято), Esc — «Отмена».
        // Раньше Esc закрывал окно молча, не сообщая «отмены».
        focus: true
        Keys.onReturnPressed: if (root.showAccept) root._accept()
        Keys.onEnterPressed: if (root.showAccept) root._accept()
        Keys.onEscapePressed: { root.rejected(); root.close() }
        id: mainLayout
        spacing: 0

        // ================= ШАПКА =================
        AppDialogHeader {
            Layout.fillWidth: true
            title: root.title
            onCloseRequested: { root.rejected(); root.close() }
        }

        Rectangle {
            Layout.fillWidth: true
            height: 1
            color: AppTheme.borderDivider
        }

        // ================= КОНТЕНТ (скроллируемый) =================
        ScrollView {
            id: scrollArea
            Layout.fillWidth: true
            Layout.fillHeight: true
            clip: true
            ScrollBar.horizontal.policy: ScrollBar.AlwaysOff
            ScrollBar.vertical.policy: ScrollBar.AsNeeded

            Column {
                id: contentArea
                width: scrollArea.width - (AppTheme.spaceL * 2)
                x: AppTheme.spaceL
                spacing: AppTheme.spaceM
                topPadding: AppTheme.spaceM
                bottomPadding: AppTheme.spaceM
            }
        }

        Rectangle {
            Layout.fillWidth: true
            height: 1
            color: AppTheme.borderDivider
        }

        // ================= ПОДВАЛ (кнопки внизу, как в дизайн-системе) =================
        Item {
            visible: root.showFooter
            Layout.fillWidth: true
            Layout.preferredHeight: 60

            Row {
                anchors.horizontalCenter: parent.horizontalCenter
                anchors.verticalCenter: parent.verticalCenter
                spacing: AppTheme.spaceM

                AppButton {
                    visible: root.showReject
                    text: root.rejectText
                    variant: root.rejectVariant
                    onClicked: { root.rejected(); root.close() }
                }
                AppButton {
                    id: acceptBtn
                    visible: root.showAccept
                    text: root.acceptText
                    variant: root.acceptVariant
                    onClicked: root._accept()
                }
            }
        }
    }

    // Автофокус первого поля: открыв форму, можно сразу печатать
    // (значение поля выделяется — ввод заменит его)
    property bool autoFocusFirst: true
    // Фокус сразу на кнопке подтверждения (окна «Вы уверены?»):
    // Enter/Space работают нативно, кольцо фокуса подсвечивает кнопку
    property bool focusAcceptOnOpen: false

    function findField(item) {
        if (!item || !item.visible || !item.enabled) return null
        if (item.isFormField === true) return item
        var kids = item.children
        for (var i = 0; i < kids.length; ++i) {
            var k = kids[i]
            if (!k || k.visible === undefined) continue   // не визуальный объект
            var r = findField(k)
            if (r) return r
        }
        return null
    }
    function focusFirstField() {
        var f = findField(contentItem)
        if (f) f.forceActiveFocus()
    }

    // Функции показа (оставляем как были)
    function showAt(callerItem, mouseX, mouseY) {
        root.morphOpen = false
        var clickPoint = callerItem.mapToItem(null, mouseX, mouseY)
        root.x = Math.max(AppTheme.spaceL, Math.min(clickPoint.x, ApplicationWindow.window.width - root.width - AppTheme.spaceL))
        root.y = Math.max(AppTheme.spaceL, Math.min(clickPoint.y + AppTheme.spaceM, ApplicationWindow.window.height - root.height - AppTheme.spaceL))
        root.open()
    }

    // Открытие после «морфинга»: диалог появляется сразу в заданном прямоугольнике,
    // без анимации масштаба — продолжает рост ячейки дня как одно целое окно.
    function openMorph(x, y, w, h) {
        root.morphOpen = true
        root.x = x
        root.y = y
        root.width = w
        root.height = h
        root.open()
    }

    function showCentered() {
        if (ApplicationWindow.window) {
            root.x = (ApplicationWindow.window.width - root.width) / 2
            root.y = (ApplicationWindow.window.height - root.height) / 2
        }
        root.open()
    }

    // При открытии раскладка ещё может не успеть досчитаться: если выставлять
    // высоту по implicitHeight сразу, окно получается короче контента и нижние
    // поля (в т.ч. тумблер «Период» и кнопки) уходят за край. Дожидаемся кадра,
    // перечитываем реальную высоту под контент и не даём окну вылезти за экран.
    onOpened: {
        // автофокус первого поля — после пересчёта геометрии
        if (root.autoFocusFirst) Qt.callLater(focusFirstField)
        if (root.focusAcceptOnOpen) Qt.callLater(function() { acceptBtn.forceActiveFocus() })
        Qt.callLater(function() {
            // Высота снова становится привязкой и дорастает вместе
            // с контентом: тексты «Что нового» досчитываются дольше
            // одного кадра, и окно, застывшее на раннем значении,
            // оказывалось меньше содержимого
            root.height = Qt.binding(function() { return root.effectiveHeight })
            if (ApplicationWindow.window) {
                var maxY = ApplicationWindow.window.height - root.height - AppTheme.spaceL
                if (root.y > maxY)
                    root.y = Math.max(AppTheme.spaceL, maxY)
                var maxX = ApplicationWindow.window.width - root.width - AppTheme.spaceL
                if (root.x > maxX)
                    root.x = Math.max(AppTheme.spaceL, maxX)
            }
        })
    }
}
