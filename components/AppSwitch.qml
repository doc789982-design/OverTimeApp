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
            // Неоморфизм: выключенная дорожка — «утопленная» ямка в поверхности
            // (вдавленная рамка-пилюля); включённая — залита цветом
            color: control.checked
                   ? AppTheme.accentBrand
                   : (AppTheme.isDark ? "#22272E" : "#E0E6EF")
            border.width: 0

            AppInsetShadow { level: 2; visible: !control.checked }

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
            border.width: 0

            // Неоморфизм: кружок «выдавлен» — круглая двухцветная тень-картинка
            BorderImage {
                z: -1
                anchors.fill: parent
                anchors.margins: -10
                source: AppTheme.isDark ? "../shadows/soft_knob_dark.png" : "../shadows/soft_knob.png"
                border.left: 20; border.right: 20; border.top: 20; border.bottom: 20
                smooth: true
                cache: true
            }

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
