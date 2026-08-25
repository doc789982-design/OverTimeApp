import QtQuick
import QtQuick.Controls
import Qt5Compat.GraphicalEffects

Switch {
    id: control

    // Цвета
    property color trackActiveColor: AppTheme.accentBrand
    property color trackInactiveColor: AppTheme.bgDisabled
    property color knobActiveColor: AppTheme.textOnAccent
    property color knobInactiveColor: AppTheme.bgElevated

    implicitHeight: 36
    focusPolicy: Qt.StrongFocus

    opacity: enabled ? 1.0 : AppTheme.alphaDisabled
    Behavior on opacity { NumberAnimation { duration: 180 } }

    indicator: Item {
        implicitWidth: 44      // ← увеличил для лучшей пропорции
        implicitHeight: 20
        x: control.leftPadding
        anchors.verticalCenter: parent.verticalCenter

        // ==================== ТРЕК ====================
        Rectangle {
            id: track
            anchors.fill: parent
            radius: 10
            color: control.checked ? control.trackActiveColor : control.trackInactiveColor

            Behavior on color {
                ColorAnimation { duration: 200; easing.type: Easing.OutCubic }
            }

            // Hover / Pressed layer
            Rectangle {
                anchors.fill: parent
                radius: 10
                color: "black"
                opacity: control.pressed ? 0.11 : control.hovered ? 0.06 : 0
                Behavior on opacity { NumberAnimation { duration: 110 } }
            }
        }

        // ==================== КРУЖОК ====================
        Rectangle {
            id: knob

            width: 24
            height: 24
            radius: 12
            anchors.verticalCenter: parent.verticalCenter

            // Идеально выверенные позиции
            x: control.checked ? parent.width - width + 1 : 1   // +1/-1 для красивого overhang

            color: control.checked ? control.knobActiveColor : control.knobInactiveColor

            // Размер (20 → 24)
            scale: control.checked ? 1.0 : 20 / 24

            // Тень-картинка для кружка (легко для видеокарты)
            Image {
                z: -1
                anchors.centerIn: parent
                anchors.verticalCenterOffset: control.checked ? 2 : 2
                width: parent.width + 20
                height: parent.height + 20
                source: "../shadows/shadow_knob.png"
                opacity: control.checked ? 0.26 : 0.20
                smooth: true
            }

            // Анимации
            Behavior on x {
                NumberAnimation {
                    duration: 260
                    easing.type: Easing.OutCubic
                }
            }

            Behavior on scale {
                NumberAnimation {
                    duration: 260
                    easing.type: Easing.OutCubic
                }
            }

            // Нажатие
            scale: control.pressed ? (control.checked ? 1.13 : 1.19) : (control.checked ? 1.0 : 20/24)

            Behavior on scale {
                NumberAnimation {
                    duration: control.pressed ? 85 : 260
                    easing.type: control.pressed ? Easing.OutQuad : Easing.OutCubic
                }
            }
        }

        // ==================== ФОКУС ====================
        Rectangle {
            anchors.fill: parent
            anchors.margins: -7
            radius: 14
            color: "transparent"
            border.color: AppTheme.borderFocus
            border.width: AppTheme.focusWidth
            opacity: control.visualFocus ? 1 : 0
            Behavior on opacity { NumberAnimation { duration: 170 } }
        }
    }

    // ==================== ТЕКСТ ====================
    contentItem: Text {
        text: control.text
        color: control.checked ? AppTheme.textPrimary : AppTheme.textSecondary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        font.weight: AppTheme.weightMedium
        verticalAlignment: Text.AlignVCenter
        leftPadding: control.indicator.width + AppTheme.spaceM
    }
}