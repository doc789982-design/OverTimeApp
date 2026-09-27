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
    implicitHeight: 40 // высота кнопки по Material Design 3
    
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
            radius: AppTheme.radiusPill // «стадион» по Material Design 3
            
            border.color: control.getVariantBorderColor()
            border.width: 1

            Behavior on color { ColorAnimation { duration: AppTheme.durFast } }
            Behavior on border.color { ColorAnimation { duration: AppTheme.durFast } }

            // МАГИЯ 2: мягкая тень по форме «стадиона». PNG-тень из AppShadow
            // нарисована для прямоугольных карточек и не совпадает с полностью
            // круглой кнопкой, поэтому собираем тень из полупрозрачных пилюль
            // (стоимость — как у обычных Rectangle). Ghost и Secondary — без тени.
            Repeater {
                model: 4
                Rectangle {
                    z: -1
                    radius: AppTheme.radiusPill
                    anchors.horizontalCenter: parent.horizontalCenter
                    anchors.verticalCenter: parent.verticalCenter
                    anchors.verticalCenterOffset: index + 1
                    width: parent.width + index * 2
                    height: parent.height + index * 2
                    color: "#000000"
                    objectName: "btnShadowPill"
                    // в тёмной теме тени не используются (Workday Canvas):
                    // глубину даёт более светлая заливка, а не чёрный ореол
                    opacity: (AppTheme.isDark ? 0.0 : 0.05) * (1.0 - index * 0.22)
                    visible: !AppTheme.isDark && control.variant !== "ghost" && control.variant !== "secondary"
                }
            }

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
            radius: AppTheme.radiusPill
            
            color: "transparent"
            border.color: AppTheme.borderFocus
            border.width: AppTheme.focusWidth
            
            opacity: control.visualFocus ? 1.0 : 0.0
            Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeColor } }
        }

        // Волна от нажатия (ripple по MD3) — верхний слой фона, контент
        // кнопки рисуется выше — волна под текстом, как в MD3. Круг из
        // центра расходится по всей кнопке и срезается её формой.
        // Драйвер — состояние pressed самой кнопки (паттерн
        // Material-стиля Qt), клику ничего не мешает. На залитых кнопках —
        // цвет текста кнопки, на контурных и ghost — бренд.
        AppRipple {
            pressed: control.pressed
            maskRadius: height / 2   // стадион кнопки
            rippleColor: (control.variant === "primary" || control.variant === "success" || control.variant === "danger")
                          ? (AppTheme.isDark ? Qt.rgba(11/255, 31/255, 51/255, 0.16)
                                             : Qt.rgba(1, 1, 1, 0.16))
                          : (AppTheme.isDark ? Qt.rgba(77/255, 154/255, 238/255, 0.20)
                                             : Qt.rgba(3/255, 116/255, 181/255, 0.16))
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