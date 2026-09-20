import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

TextField {
    id: root

    property string label: ""
    property bool isRequired: false
    property color cutoutColor: AppTheme.bgModal
    property bool numericOnly: false   // true = поле принимает только цифры
    property bool isFloated: root.text.length > 0 || root.activeFocus

    // Защита числовых полей: буквы физически невозможно ввести
    validator: root.numericOnly ? digitsOnly : null
    RegularExpressionValidator {
        id: digitsOnly
        regularExpression: /^[0-9]{0,5}$/
    }

    implicitHeight: 44 
    Layout.fillWidth: true
    
    leftPadding: AppTheme.spaceM
    rightPadding: AppTheme.spaceM
    verticalAlignment: TextInput.AlignVCenter
    
    color: root.enabled ? AppTheme.textPrimary : AppTheme.textDisabled
    font.family: AppTheme.fontFamily
    font.pixelSize: AppTheme.sizeBody
    
    // ==========================================
    // МАГИЯ ПЛАВНОЙ ПОДСКАЗКИ
    // Мы смешиваем цвет текста со 100% прозрачностью (Qt.rgba)
    // ==========================================
    placeholderTextColor: {
        if (root.activeFocus && floatingLabel.y < 0) {
            return AppTheme.textTertiary; // Цвет виден
        } else {
            // Тот же цвет, но с альфа-каналом 0.0 (полностью прозрачный)
            return Qt.rgba(AppTheme.textTertiary.r, AppTheme.textTertiary.g, AppTheme.textTertiary.b, 0.0);
        }
    }
    
    // Плавная анимация изменения цвета (Fade-эффект)
    Behavior on placeholderTextColor { 
        ColorAnimation { duration: AppTheme.durNormal; easing.type: Easing.OutQuad } 
    }
    
    focusPolicy: Qt.StrongFocus
    cursorDelegate: AppCursorDelegate {}

    // ==========================================
    // 1. РАМКА
    // ==========================================
    background: Rectangle {
        color: "transparent"
        radius: AppTheme.radiusMedium
        
        border.color: !root.enabled ? AppTheme.borderDisabled :
                      (root.activeFocus ? AppTheme.borderFocus : 
                      (root.hovered ? AppTheme.textSecondary : AppTheme.borderInput))
        
        border.width: root.activeFocus ? AppTheme.focusWidth : 1
        Behavior on border.color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    // ==========================================
    // 2. ИДЕАЛЬНЫЙ ЛАСТИК (Eraser)
    // ==========================================
    Rectangle {
        color: root.cutoutColor
        x: floatingLabel.x - 4
        y: -2 
        height: 4 
        width: (floatingLabel.width * floatingLabel.scale) + 8
        
        opacity: root.isFloated ? 1.0 : 0.0
        Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
    }

    // ==========================================
    // 3. ПЛАВАЮЩИЙ ЛЕЙБЛ
    // ==========================================
    Row {
        id: floatingLabel
        x: AppTheme.spaceS
        
        y: root.isFloated ? -(height * 0.75) / 2 : (root.height - height) / 2
        scale: root.isFloated ? 0.75 : 1.0
        transformOrigin: Item.TopLeft
        
        Behavior on y { NumberAnimation { duration: AppTheme.durNormal; easing.type: Easing.OutCubic } }
        Behavior on scale { NumberAnimation { duration: AppTheme.durNormal; easing.type: Easing.OutCubic } }

        spacing: AppTheme.spaceMicro

        Text {
            text: root.label
            color: !root.enabled ? AppTheme.textDisabled : 
                   (root.activeFocus ? AppTheme.accentBrand : AppTheme.textSecondary)
                   
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody 
            font.weight: root.isFloated ? AppTheme.weightMedium : AppTheme.weightRegular
            
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
        }
        
        Text {
            visible: root.isRequired
            text: "*"
            color: root.enabled ? AppTheme.accentDanger : AppTheme.textDisabled 
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody 
        }
    }
}