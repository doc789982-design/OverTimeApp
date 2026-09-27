import QtQuick
import QtQuick.Controls
import QtQuick.Shapes // Нужен для красивой галочки!

CheckBox {
    id: control

    // Цвета по умолчанию из темы
    property color activeColor: AppTheme.accentBrand
    // невыбранная обводка — как onSurfaceVariant в MD3 (контрастнее бледной рамки полей)    property color inactiveColor: AppTheme.textSecondary
    property color checkColor: AppTheme.textOnAccent
    
    implicitHeight: 36 
    focusPolicy: Qt.StrongFocus

    // Физика (Прозрачность и Вжатие)
    opacity: control.enabled ? 1.0 : AppTheme.alphaDisabled
    scale: control.pressed ? AppTheme.scaleActive : 1.0
    
    Behavior on scale { NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeStandard } }
    Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeColor } }

    indicator: Item {
        implicitWidth: 20
        implicitHeight: 20
        x: control.leftPadding
        anchors.verticalCenter: parent.verticalCenter

        // ==========================================
        // 0. ЗОНА ОТКЛИКА 40×40 (touch target из MD3):
        // ховер, нажатие и волна работают по всей зоне,
        // а не только по квадратику 18px
        // ==========================================
        Item {
            id: touchZone
            anchors.centerIn: box
            width: 40
            height: 40

            Rectangle {
                anchors.fill: parent
                radius: AppTheme.radiusPill
                color: control.pressed ? AppTheme.statePress
                       : (control.hovered ? AppTheme.stateHover : "transparent")
                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
            }

            AppRipple {
                // круг из центра зоны, расходится по всей зоне отклика;
                // драйвер — состояние pressed самого чекбокса
                pressed: control.pressed
                rippleColor: AppTheme.isDark ? Qt.rgba(77/255, 154/255, 238/255, 0.20)
                                             : Qt.rgba(3/255, 116/255, 181/255, 0.16)
            }
        }

        // ==========================================
        // 1. КВАДРАТ (Основа) — 18px, скругление 2, обводка 2px по MD3
        // ==========================================
        Rectangle {
            id: box
            anchors.centerIn: parent
            width: 18
            height: 18
            radius: 2

            // Если выбран - заливаем брендом, если нет - прозрачный
            color: control.checked ? control.activeColor : "transparent"
            border.color: control.checked ? control.activeColor : control.inactiveColor
            border.width: 2

            Behavior on color { ColorAnimation { duration: AppTheme.durFast } }
            Behavior on border.color { ColorAnimation { duration: AppTheme.durFast } }
        }

        // ==========================================
        // 2. ГАЛОЧКА (Красивая отрисовка векторами)
        // ==========================================
        Shape {
            anchors.fill: parent
            visible: control.checked
            opacity: control.checked ? 1.0 : 0.0
            
            // Анимация масштаба: галочка выпрыгивает из центра
            scale: control.checked ? 1.0 : 0.5
            Behavior on scale { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeEnter } }
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }

            ShapePath {
                strokeColor: control.checkColor
                strokeWidth: 2
                capStyle: ShapePath.RoundCap
                joinStyle: ShapePath.RoundJoin
                fillColor: "transparent"

                // Идеальные пропорции галочки внутри квадрата 18x18
                startX: 4.5
                startY: 9
                PathLine { x: 8; y: 12.5 }
                PathLine { x: 13.5; y: 5.5 }
            }
        }

        // ==========================================
        // 3. КОЛЬЦО ФОКУСА (Accessibility)
        // ==========================================
        Rectangle {
            anchors.fill: parent
            anchors.margins: -AppTheme.focusOffset - AppTheme.focusWidth
            radius: 2 + AppTheme.focusOffset
            
            color: "transparent"
            border.color: AppTheme.borderFocus
            border.width: AppTheme.focusWidth
            
            opacity: control.visualFocus ? 1.0 : 0.0
            Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeColor } }
        }
    }

    // ==========================================
    // 4. ТЕКСТ СПРАВА ОТ ЧЕКБОКСА
    // ==========================================
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