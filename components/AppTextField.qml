import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

TextField {
    id: root

    property string label: ""
    property bool isRequired: false
    property color cutoutColor: AppTheme.bgModal
    property bool numericOnly: false   // true = поле принимает только цифры
    property bool hasError: false       // true = показать ошибку (вспышка рамки)

    // КРАСНАЯ РАМКА ОШИБКИ — ВСПЫШКА: flashError() зажигает рамку
    // на 3 секунды, потом она ПЛАВНО гаснет — не висит, пока не
    // введёшь значение. Повторное «Сохранить» — вспышка снова.
    // Ввод значения гасит рамку сразу.
    property real errGlow: 0           // 0..1 — сила красной рамки
    function flashError() {
        root.hasError = true
        errFade.stop()
        errLight.restart()
        errHold.restart()
    }
    onHasErrorChanged: {
        if (root.hasError) {
            errFade.stop()
            errLight.restart()
            errHold.restart()
        } else {
            errHold.stop()
            errFade.restart()          // плавное затухание
        }
    }
    PropertyAnimation { id: errLight; target: root; property: "errGlow";
                        to: 1; duration: AppTheme.durFast }
    PropertyAnimation { id: errFade; target: root; property: "errGlow";
                        to: 0; duration: 700; easing.type: Easing.InQuad }
    Timer { id: errHold; interval: 3000; onTriggered: root.hasError = false }
    onTextEdited: if (root.hasError) root.hasError = false
    property bool isFormField: true   // автофокус диалога ищет такие
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

    // Выделение как в зрелых приложениях: брендовый цвет (не системный
    // синий) и явная работа мышью; в числовых полях двойной клик и так
    // выделяет всё число (текст без разделителей — одно слово)
    selectByMouse: true
    selectionColor: AppTheme.accentBrand
    selectedTextColor: AppTheme.textOnAccent

    // Tab/автофокус: всё значение выделяется — ввод сразу заменяет его
    onActiveFocusChanged: {
        if (activeFocus) {
            // Tab/автофокус: значение выделяется целиком
            root._focusFresh = true
            Qt.callLater(function() { root.selectAll(); root._focusFresh = false })
        }
    }

    // Клик по полю, только что получившему фокус (в т.ч. самим кликом —
    // TextInput ставит фокус ДО сигнала pressed): выделяем всё значение,
    // ввод сразу заменит его. Первый клик — выделение, повторный по
    // сфокусированному полю — позиция курсора (нативное поведение).
    property bool _focusFresh: false
    property bool _clickGuard: false
    onPressed: (mouse) => {
        if (root._focusFresh) {
            root.selectAll()
            mouse.accepted = true
            root._clickGuard = true
        }
    }
    // отпускание после клик-выделения не должно двигать курсор
    onReleased: (mouse) => {
        if (root._clickGuard) {
            mouse.accepted = true
            root._clickGuard = false
            // базовый обработчик мог сбить выделение — восстанавливаем
            // уже после полной обработки отпускания
            Qt.callLater(function() { root.selectAll() })
        }
    }

    // ==========================================
    // 1. РАМКА
    // ==========================================
    background: Rectangle {
        color: "transparent"
        radius: AppTheme.radiusMedium
        
        border.color: {
            let base = !root.enabled ? AppTheme.borderDisabled :
                       (root.activeFocus ? AppTheme.borderFocus :
                       (root.hovered ? AppTheme.textSecondary : AppTheme.borderInput))
            if (root.errGlow <= 0.001) return base
            let d = AppTheme.accentDanger
            return Qt.rgba(base.r + (d.r - base.r) * root.errGlow,
                           base.g + (d.g - base.g) * root.errGlow,
                           base.b + (d.b - base.b) * root.errGlow, 1)
        }
        
        border.width: (root.activeFocus || root.errGlow > 0.01) ? AppTheme.focusWidth : 1
        Behavior on border.color { ColorAnimation { duration: AppTheme.durMicro } }
        Behavior on border.width { NumberAnimation { duration: AppTheme.durMicro } }
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