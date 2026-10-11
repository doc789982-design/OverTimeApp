import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

// ============================================================
// УНИВЕРСАЛЬНОЕ ДИАЛОГ ПОДТВЕРЖДЕНИЯ
// Обычный AppDialog: вопрос в шапке, иконка и описание в центре,
// опасная кнопка выделена красным. Enter — подтвердить,
// Esc/крестик — отменить.
// ============================================================
AppDialog {
    id: root

    property string dialogTitle: "Вы уверены?"
    property string dialogMessage: ""
    property string confirmLabel: "Подтвердить"
    property bool dangerMode: true   // true = красная кнопка (удаление и т.п.)

    title: root.dialogTitle
    closeOnAccept: true
    acceptText: root.confirmLabel
    acceptVariant: root.dangerMode ? "danger" : "primary"
    rejectText: "Отмена"
    rejectVariant: "secondary"

    width: 420
    anchors.centerIn: parent

    // Нет полей ввода — фокус сразу на кнопке подтверждения:
    // Enter/Space работают нативно, кольцо фокуса подсвечивает её
    autoFocusFirst: false
    focusAcceptOnOpen: true

    // ИКОНКА В КРУЖКЕ
    Rectangle {
        width: 44
        height: 44
        radius: AppTheme.radiusPill
        anchors.horizontalCenter: parent.horizontalCenter
        color: root.dangerMode ? AppTheme.bgDangerSoft : AppTheme.bgBrandSoft

        IconImage {
            anchors.centerIn: parent
            source: root.dangerMode ? "../icons/trash.svg" : "../icons/help.svg"
            width: AppTheme.iconLarge
            height: AppTheme.iconLarge
            color: root.dangerMode ? AppTheme.accentDanger : AppTheme.accentBrand
        }
    }

    // ОПИСАНИЕ
    Text {
        width: parent.width
        text: root.dialogMessage
        visible: root.dialogMessage !== ""
        color: AppTheme.textSecondary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        horizontalAlignment: Text.AlignHCenter
        wrapMode: Text.WordWrap
        lineHeight: 1.25
    }


    function showCentered() {
        root.open()
    }
}
