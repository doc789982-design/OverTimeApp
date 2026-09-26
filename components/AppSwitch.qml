import QtQuick
import QtQuick.Controls

// ============================================================
// ПЕРЕКЛЮЧАТЕЛЬ — геометрия и поведение Material Design 3
//
// Трек 52×32. Бегунок живой: 16px в выключенном состоянии,
// 24px во включённом, 28px под пальцем — и всё это анимируется
// (250 мс, кривая emphasized). Цвета: выключенный бегунок — цвет
// текста темы (аналог onSurface в MD3), включённый — цвет текста
// на бренде (в тёмной теме тёмный, в светлой белый, как у кнопок).
// Контур выключенного трека — 2px (по MD3, было 1px).
// ============================================================
Switch {
    id: control

    implicitHeight: 36
    padding: 0
    spacing: 0
    focusPolicy: Qt.StrongFocus
    opacity: enabled ? 1.0 : AppTheme.alphaDisabled
    Behavior on opacity { NumberAnimation { duration: AppTheme.durNormal } }

    indicator: Item {
        implicitWidth: 52
        implicitHeight: 32
        x: control.leftPadding
        anchors.verticalCenter: parent.verticalCenter

        Rectangle {
            id: track
            anchors.fill: parent
            radius: AppTheme.radiusPill
            color: control.checked ? AppTheme.accentBrand
                   : (AppTheme.isDark ? "#3A3F46" : AppTheme.bgCell)
            border.width: control.checked ? 0 : 2
            border.color: AppTheme.textSecondary

            Behavior on color { ColorAnimation { duration: 250; easing.type: AppTheme.easeColor } }
            Behavior on border.width { NumberAnimation { duration: AppTheme.durFast } }
        }

        Rectangle {
            anchors.fill: parent
            radius: AppTheme.radiusPill
            color: control.pressed ? AppTheme.statePress
                   : (control.hovered ? AppTheme.stateHover : "transparent")
        }

        // Волна от нажатия — по всей зоне трека (MD3 ripple)
        AppRipple {
            rippleColor: control.checked
                          ? (AppTheme.isDark ? Qt.rgba(11/255, 31/255, 51/255, 0.16)
                                             : Qt.rgba(1, 1, 1, 0.16))
                          : (AppTheme.isDark ? Qt.rgba(77/255, 154/255, 238/255, 0.20)
                                             : Qt.rgba(3/255, 116/255, 181/255, 0.16))
        }

        Rectangle {
            id: thumb
            // центры бегунка: 14px от левого края (выкл) и 38px (вкл)
            property real centerX: control.checked ? parent.width - 14 : 14
            // 16 (выкл) -> 24 (вкл), под пальцем подрастает до 28
            width: control.pressed ? 28 : (control.checked ? 24 : 16)
            height: width
            radius: width / 2
            anchors.verticalCenter: parent.verticalCenter
            x: centerX - width / 2
            color: control.checked ? AppTheme.textOnAccent : AppTheme.textPrimary

            Behavior on width { NumberAnimation { duration: 200; easing.type: AppTheme.easeStandard } }
            Behavior on x { NumberAnimation { duration: 250; easing.type: AppTheme.easeStandard } }
            Behavior on color { ColorAnimation { duration: 250; easing.type: AppTheme.easeColor } }
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
