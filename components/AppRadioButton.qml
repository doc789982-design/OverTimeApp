import QtQuick
import QtQuick.Controls

// ============================================================
// РАДИОКНОПКА ДИЗАЙН-СИСТЕМЫ
//
// Круг с точкой по образцу AppCheckBox: та же зона отклика 40×40
// с ховером и волной, то же кольцо фокуса, тот же текст справа.
// В отличие от CheckBox, сама себя не переключает: клик только
// зовёт toggled(), а состояние (checked) полностью управляется
// снаружи — группа «текущий / предыдущий / оба года» живёт в
// одном свойстве окна, и исключительность выбирает оно.
// ============================================================
Item {
    id: control

    property string text: ""
    property bool checked: false
    property color activeColor: AppTheme.accentBrand
    property color inactiveColor: AppTheme.textSecondary
    property color dotColor: AppTheme.textOnAccent

    signal toggled()

    implicitHeight: 36
    // x индикатора входит в ширину: без него лейбл короче на 2px
    // и текст вечно обрезался многоточием («Текущий год…»)
    implicitWidth: indicatorRow.x + indicatorRow.width + AppTheme.spaceM + label.implicitWidth

    opacity: control.enabled ? 1.0 : AppTheme.alphaDisabled
    scale: ma.pressed ? AppTheme.scaleActive : 1.0

    Behavior on scale { NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeStandard } }
    Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeColor } }

    Item {
        id: indicatorRow
        width: 20
        height: 36
        x: 2

        // Зона отклика 40×40 (touch target из MD3): ховер, нажатие
        // и волна работают по всей зоне, а не только по кругу 20px.
        Item {
            id: touchZone
            anchors.centerIn: ring
            width: 40
            height: 40

            Rectangle {
                anchors.fill: parent
                radius: AppTheme.radiusPill
                color: ma.pressed ? AppTheme.statePress
                       : (ma.containsMouse ? AppTheme.stateHover : "transparent")
                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
            }

            AppRipple {
                pressed: ma.pressed
                maskRadius: width / 2
                rippleColor: AppTheme.isDark ? Qt.rgba(77/255, 154/255, 238/255, 0.20)
                                             : Qt.rgba(3/255, 116/255, 181/255, 0.16)
            }
        }

        // Круг (основа) — 20px, обводка 2px по MD3
        Rectangle {
            id: ring
            anchors.centerIn: parent
            width: 20
            height: 20
            radius: 10
            color: control.checked ? control.activeColor : "transparent"
            border.color: control.checked ? control.activeColor : control.inactiveColor
            border.width: 2

            Behavior on color { ColorAnimation { duration: AppTheme.durFast } }
            Behavior on border.color { ColorAnimation { duration: AppTheme.durFast } }
        }

        // Точка выбора: выпрыгивает из центра, как галочка чекбокса
        Rectangle {
            anchors.centerIn: ring
            width: 10
            height: 10
            radius: 5
            color: control.dotColor

            scale: control.checked ? 1.0 : 0.0
            opacity: control.checked ? 1.0 : 0.0
            Behavior on scale { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeEnter } }
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
        }

        // Кольцо фокуса (Accessibility)
        Rectangle {
            anchors.fill: ring
            anchors.margins: -AppTheme.focusOffset - AppTheme.focusWidth
            radius: 10 + AppTheme.focusOffset

            color: "transparent"
            border.color: AppTheme.borderFocus
            border.width: AppTheme.focusWidth

            opacity: ma.activeFocus ? 1.0 : 0.0
            Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeColor } }
        }
    }

    Text {
        id: label
        x: indicatorRow.x + indicatorRow.width + AppTheme.spaceM
        anchors.verticalCenter: parent.verticalCenter
        width: Math.max(0, control.width - x)

        text: control.text
        color: control.checked ? AppTheme.textPrimary : AppTheme.textSecondary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        font.weight: control.checked ? AppTheme.weightBold : AppTheme.weightMedium
        elide: Text.ElideRight
    }

    HoverHandler { cursorShape: Qt.PointingHandCursor }
    MouseArea {
        id: ma
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: Qt.PointingHandCursor
        onClicked: if (control.enabled) control.toggled()
    }
}
