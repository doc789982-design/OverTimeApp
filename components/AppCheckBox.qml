import QtQuick
import QtQuick.Controls
import QtQuick.Shapes // Нужен для красивой галочки!

CheckBox {
    id: control

    // Цвета по умолчанию из темы
    property color activeColor: AppTheme.accentBrand
    property color inactiveColor: AppTheme.borderInput
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
        // 1. КВАДРАТ (Основа)
        // ==========================================
        Rectangle {
            anchors.fill: parent
            radius: AppTheme.radiusSmall // 4px (строгий квадрат с легким скруглением)
            
            // Если выбран - заливаем брендом, если нет - прозрачный
            color: control.checked ? control.activeColor : "transparent"
            border.color: control.checked ? control.activeColor : control.inactiveColor
            border.width: 1
            
            Behavior on color { ColorAnimation { duration: AppTheme.durFast } }
            Behavior on border.color { ColorAnimation { duration: AppTheme.durFast } }

            // Слой ховера (Слегка затемняет/высветляет квадратик при наведении)
            Rectangle {
                anchors.fill: parent
                radius: parent.radius
                color: control.hovered ? AppTheme.stateHover : "transparent"
            }
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

                // Идеальные пропорции галочки внутри квадрата 20x20
                startX: 5
                startY: 10
                PathLine { x: 9; y: 14 }
                PathLine { x: 15; y: 6 }
            }
        }

        // ==========================================
        // 3. КОЛЬЦО ФОКУСА (Accessibility)
        // ==========================================
        Rectangle {
            anchors.fill: parent
            anchors.margins: -AppTheme.focusOffset - AppTheme.focusWidth
            radius: AppTheme.radiusSmall + AppTheme.focusOffset
            
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