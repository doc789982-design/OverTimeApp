import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects // Для тени Level 1

Button {
    id: control
    
    // --- Настройки кнопки ---
    property string variant: "primary" // primary, success, danger, secondary, ghost
    property string iconSource: "" 
    
    // Строгая высота по дизайн-системе
    implicitHeight: 36 
    
    // Ширина = ширина контента + системные отступы (минимум 100px)
    implicitWidth: Math.max(100, contentRow.implicitWidth + (AppTheme.spaceM * 2))

    // Отключаем системный фокус Qt, рисуем свой
    focusPolicy: Qt.StrongFocus

    // ==========================================
    // 1. АНИМАЦИЯ ВЖАТИЯ (Scale Physics)
    // ==========================================
    scale: control.pressed ? AppTheme.scaleActive : 1.0
    Behavior on scale { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeStandard } }
    
    // Прозрачность для отключенной кнопки
    opacity: control.enabled ? 1.0 : AppTheme.alphaDisabled
    Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeColor } }

    // ==========================================
    // 2. ЦВЕТОВЫЕ ФУНКЦИИ
    // ==========================================
    function getVariantBgColor() {
        if (variant === "primary")   return AppTheme.accentBrand
        if (variant === "success")   return AppTheme.accentSuccess
        if (variant === "danger")    return AppTheme.accentDanger
        
        // МАГИЯ ЗДЕСЬ: Secondary кнопка теперь полностью прозрачная!
        if (variant === "secondary") return "transparent"   
        if (variant === "ghost")     return "transparent"
        
        return AppTheme.accentBrand
    }

    function getVariantTextColor() {
        if (variant === "primary" || variant === "success" || variant === "danger") 
            return AppTheme.textOnAccent 
            
        if (variant === "secondary" || variant === "ghost") 
            return AppTheme.textPrimary  
            
        return AppTheme.textPrimary
    }

    function getVariantBorderColor() {
        if (variant === "secondary") return AppTheme.borderInput
        return "transparent"
    }

    // ==========================================
    // 3. ФОН И СОСТОЯНИЯ
    // ==========================================
    background: Item {
        anchors.fill: parent

        Rectangle {
            id: bgRect
            anchors.fill: parent
            color: control.getVariantBgColor()
            radius: AppTheme.radiusMedium 
            
            border.color: control.getVariantBorderColor()
            border.width: 1

            Behavior on color { ColorAnimation { duration: AppTheme.durFast } }
            Behavior on border.color { ColorAnimation { duration: AppTheme.durFast } }

            // МАГИЯ 2: Тень-картинка; для Ghost и Secondary кнопок отключена!
            AppShadow { level: 1; visible: control.variant !== "ghost" && control.variant !== "secondary" }

            // МАГИЯ: Слой состояния (Hover / Press)
            // Он ложится поверх базового цвета, делая синий - темно-синим, а зеленый - темно-зеленым!
            Rectangle {
                anchors.fill: parent
                radius: parent.radius
                color: control.pressed ? AppTheme.statePress : (control.hovered ? AppTheme.stateHover : "transparent")
                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
            }
        }

        // ==========================================
        // 4. КОЛЬЦО ФОКУСА (Accessibility)
        // ==========================================
        Rectangle {
            anchors.fill: parent
            anchors.margins: -AppTheme.focusOffset - AppTheme.focusWidth
            radius: AppTheme.radiusMedium + AppTheme.focusOffset
            
            color: "transparent"
            border.color: AppTheme.borderFocus
            border.width: AppTheme.focusWidth
            
            opacity: control.visualFocus ? 1.0 : 0.0
            Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeColor } }
        }
    }

    // ==========================================
    // 5. КОНТЕНТ (Иконка + Текст)
    // ==========================================
    contentItem: Item {
        anchors.fill: parent 
        
        Row {
            id: contentRow
            anchors.centerIn: parent 
            spacing: AppTheme.spaceS
            
            IconImage {
                visible: control.iconSource !== ""
                source: control.iconSource
                width: AppTheme.iconMedium
                height: AppTheme.iconMedium
                color: control.getVariantTextColor()
                anchors.verticalCenter: parent.verticalCenter
            }

            Text {
                text: control.text
                color: control.getVariantTextColor()
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                font.weight: AppTheme.weightMedium // Полужирный для кнопок
                anchors.verticalCenter: parent.verticalCenter
            }
        }
    }
}