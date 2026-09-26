import QtQuick
import QtQuick.Controls

// ============================================================
// ПЕРЕКЛЮЧАТЕЛЬ — геометрия Material Design 3, темп десктопный
//
// Трек 52×32. Бегунок живой: 16px в выключенном состоянии,
// 24px во включённом, 28px под курсором. Переезд — на пружине
// (резвый старт, лёгкий перелёт), рост — 120 мс. Цвета: выключенный
// бегунок — цвет текста темы (onSurface в MD3), включённый — цвет
// текста на бренде (в тёмной теме тёмный, в светлой белый, как у
// кнопок). Контур выключенного трека — 2px (по MD3, было 1px).
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

            Behavior on color { ColorAnimation { duration: 180; easing.type: AppTheme.easeColor } }
            Behavior on border.width { NumberAnimation { duration: AppTheme.durFast } }
        }

        Rectangle {
            anchors.fill: parent
            radius: AppTheme.radiusPill
            color: control.pressed ? AppTheme.statePress
                   : (control.hovered ? AppTheme.stateHover : "transparent")
        }

        // Волна от нажатия — пилюля по форме трека, растёт из центра
        AppRipple {
            rippleShape: 1
            rippleColor: control.checked
                          ? (AppTheme.isDark ? Qt.rgba(11/255, 31/255, 51/255, 0.16)
                                             : Qt.rgba(1, 1, 1, 0.16))
                          : (AppTheme.isDark ? Qt.rgba(77/255, 154/255, 238/255, 0.20)
                                             : Qt.rgba(3/255, 116/255, 181/255, 0.16))
        }

        // Бегунок: позиция — отдельным позиционером по ЦЕНТРУ (14 выкл / 38 вкл),
        // чтобы рост размера не дёргал цель пружины каждый кадр
        Item {
            id: thumbPos
            x: control.checked ? parent.width - 14 : 14
            y: parent.height / 2

            // резвый старт и лёгкий перелёт — «инерция» механического
            // переключателя (в MD3 переезд описан пружиной)
            Behavior on x { SpringAnimation { spring: 5.0; damping: 0.35; mass: 1.3 } }

            Rectangle {
                id: thumb
                anchors.centerIn: parent
                // 16 (выкл) -> 24 (вкл), под курсором подрастает до 28
                width: control.pressed ? 28 : (control.checked ? 24 : 16)
                height: width
                radius: width / 2
                color: control.checked ? AppTheme.textOnAccent : AppTheme.textPrimary

                Behavior on width { NumberAnimation { duration: 120; easing.type: Easing.OutCubic } }
                Behavior on color { ColorAnimation { duration: 180; easing.type: AppTheme.easeColor } }
            }
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
