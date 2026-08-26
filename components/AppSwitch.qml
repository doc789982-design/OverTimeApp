import QtQuick
import QtQuick.Controls

Switch {
    id: control

    implicitHeight: 36
    padding: 0
    spacing: 0
    focusPolicy: Qt.StrongFocus
    opacity: enabled ? 1.0 : AppTheme.alphaDisabled
    Behavior on opacity { NumberAnimation { duration: AppTheme.durNormal } }

    indicator: Item {
        implicitWidth: 44
        implicitHeight: 24
        x: control.leftPadding
        anchors.verticalCenter: parent.verticalCenter

        Rectangle {
            id: track
            anchors.fill: parent
            radius: AppTheme.radiusPill
            color: control.checked
                   ? AppTheme.accentBrand
                   : (AppTheme.isDark ? "#3A3F46" : AppTheme.borderInput)
            border.width: control.checked ? 0 : 1
            border.color: AppTheme.isDark ? AppTheme.borderInput : AppTheme.borderDisabled

            Behavior on color { ColorAnimation { duration: AppTheme.durNormal; easing.type: AppTheme.easeColor } }
            Behavior on border.width { NumberAnimation { duration: AppTheme.durFast } }
        }

        Rectangle {
            anchors.fill: parent
            radius: AppTheme.radiusPill
            color: control.pressed ? AppTheme.statePress
                   : (control.hovered ? AppTheme.stateHover : "transparent")
        }

        Rectangle {
            id: thumb
            width: 20
            height: 20
            radius: 10
            anchors.verticalCenter: parent.verticalCenter
            x: control.checked ? parent.width - width - 2 : 2
            color: "#FFFFFF"
            border.width: control.checked ? 0 : 1
            border.color: AppTheme.isDark ? AppTheme.borderInput : AppTheme.borderDisabled

            Behavior on x {
                NumberAnimation { duration: AppTheme.durNormal; easing.type: AppTheme.easeStandard }
            }
            Behavior on border.width { NumberAnimation { duration: AppTheme.durFast } }
        }

        Rectangle {
            anchors.fill: parent
            anchors.margins: -4
            radius: AppTheme.radiusPill
            color: "transparent"
            border.color: AppTheme.borderFocus
            border.width: AppTheme.focusWidth
            opacity: control.visualFocus ? 1 : 0
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
        }
    }

    contentItem: Text {
        text: control.text
        color: AppTheme.textPrimary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        font.weight: AppTheme.weightMedium
        verticalAlignment: Text.AlignVCenter
        leftPadding: control.indicator.width + AppTheme.spaceM
    }
}
