import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

// ============================================================
// ШАПКА ОКНА — единая у всех окон на AppModalShell:
// заголовок слева (конденсатный жирный, как дата в меню дня),
// место для доп. элементов справа (переключатель годов у календа-
// ря) и крестик. Высота 60, отбивку снизу рисуетcaller.
// ============================================================
Item {
    id: root

    property string title: ""
    // Доп. элементы шапки (справа, перед крестиком)
    default property alias headerExtras: extrasSlot.data
    signal closeRequested()

    implicitHeight: 60
    height: 60

    Text {
        anchors.left: parent.left
        anchors.leftMargin: AppTheme.spaceL
        anchors.verticalCenter: parent.verticalCenter
        anchors.right: parent.right
        anchors.rightMargin: 56
        text: root.title
        color: AppTheme.textPrimary
        // Тот же конденсатный жирный шрифт, что и дата в меню дня
        font.family: AppTheme.fontCondensed
        font.pixelSize: AppTheme.sizeH4
        font.weight: AppTheme.weightBold
        visible: text !== ""
        elide: Text.ElideRight
        verticalAlignment: Text.AlignVCenter
    }

    // Слот для доп. элементов: год-переключатель календаря и т.п.
    // Контент прижимается вправо сам (anchors.right: parent.right),
    // ширина слота нулевая — он лишь задаёт точку отсчёта
    Item {
        id: extrasSlot
        anchors.right: closeButton.left
        anchors.rightMargin: AppTheme.spaceXS
        anchors.verticalCenter: parent.verticalCenter
    }

    Rectangle {
        id: closeButton
        width: 32; height: 32; radius: AppTheme.radiusPill
        anchors.right: parent.right
        anchors.rightMargin: AppTheme.spaceM
        anchors.verticalCenter: parent.verticalCenter

        color: closeHov.pressed ? AppTheme.statePress : (closeHov.containsMouse ? AppTheme.stateHover : "transparent")
        IconImage { anchors.centerIn: parent; source: "../icons/close.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: AppTheme.textSecondary }

        // Клавиатура: таб, Enter/Пробел; имя для экранного доступа
        activeFocusOnTab: true
        Keys.onReturnPressed: root.closeRequested()
        Keys.onEnterPressed: root.closeRequested()
        Keys.onSpacePressed: root.closeRequested()
        Accessible.role: Accessible.Button
        Accessible.name: "Закрыть окно"

        // рамка фокуса с клавиатуры — как у кнопок формы
        Rectangle {
            anchors.fill: parent
            anchors.margins: -AppTheme.focusOffset
            radius: AppTheme.radiusPill
            color: "transparent"
            border.color: AppTheme.borderFocus
            border.width: AppTheme.focusWidth
            opacity: closeButton.activeFocus ? 1.0 : 0.0
            Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro } }
        }

        HoverHandler { cursorShape: Qt.PointingHandCursor }
        MouseArea { id: closeHov; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: root.closeRequested() }
    }
}
