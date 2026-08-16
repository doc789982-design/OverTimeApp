import QtQuick
import QtQuick.Controls
import Qt5Compat.GraphicalEffects

Popup {
    id: root
    
    // Автоматическое позиционирование в правом нижнем углу экрана
    x: parent ? parent.width - width - AppTheme.spaceL : 0
    y: parent ? parent.height - height - AppTheme.spaceL : 0
    
    modal: false
    focus: false
    closePolicy: Popup.NoAutoClose
    
    // Z-Index: Поверх окон и тултипов
    z: AppTheme.zToast
    
    // Ширина подстраивается под текст, плюс ровные системные отступы
    width: toastText.implicitWidth + (AppTheme.spaceL * 2)
    height: 50
    
    property string toastType: "success"
    property string message: ""

    // ==========================================
    // 1. ФОН И ТЕНЬ (Level 5)
    // ==========================================
    background: Rectangle {
        color: AppTheme.bgElevated
        radius: AppTheme.radiusMedium
        
        // Цвет рамки зависит от типа уведомления
        border.color: root.toastType === "error" ? AppTheme.accentDanger : AppTheme.accentSuccess
        border.width: 1
        
        // Тень-картинка вместо вычисляемой (Level 5)
        AppShadow { level: 5 }
        
        // Тонкая цветная полоска слева (Фишка профессиональных тоастов)
        Rectangle {
            anchors.left: parent.left
            anchors.top: parent.top
            anchors.bottom: parent.bottom
            width: 4
            radius: AppTheme.radiusMedium
            // Чтобы левые углы были круглыми, а правые острыми:
            Rectangle {
                anchors.right: parent.right
                anchors.top: parent.top
                anchors.bottom: parent.bottom
                width: 2
                color: root.toastType === "error" ? AppTheme.accentDanger : AppTheme.accentSuccess
            }
            color: root.toastType === "error" ? AppTheme.accentDanger : AppTheme.accentSuccess
        }
    }

    // ==========================================
    // 2. КОНТЕНТ (Текст)
    // ==========================================
    contentItem: Text {
        id: toastText
        anchors.centerIn: parent
        text: root.message
        color: AppTheme.textPrimary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        font.weight: AppTheme.weightMedium
    }

    // ==========================================
    // 3. ТАЙМЕР И АНИМАЦИИ (Motion)
    // ==========================================
    Timer {
        id: toastTimer
        interval: 3000 // 3 секунды до закрытия
        repeat: false
        onTriggered: root.close()
    }

    // Анимация выпрыгивания: Slide Up + Fade
    enter: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeColor }
            NumberAnimation { 
                property: "y"; 
                from: root.parent.height; 
                to: root.parent.height - root.height - AppTheme.spaceL; 
                duration: AppTheme.durStandard; 
                // Оставляем микро-пружинку (OutBack) ТОЛЬКО для Тоастов, чтобы привлечь периферийное зрение!
                easing.type: Easing.OutBack 
            }
        }
    }
    
    // Анимация затухания: Slide Down + Fade
    exit: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: AppTheme.durFast; easing.type: AppTheme.easeColor }
            NumberAnimation { 
                property: "y"; 
                from: root.y; 
                to: root.y + AppTheme.slideOffset; 
                duration: AppTheme.durFast; 
                easing.type: AppTheme.easeExit 
            }
        }
    }

    // ==========================================
    // 4. ГЛАВНАЯ ФУНКЦИЯ
    // ==========================================
    function show(msg, type) {
        root.message = msg
        root.toastType = type || "success"
        
        if (root.opened) {
            toastTimer.restart()
        } else {
            root.open()
            toastTimer.start()
        }
    }
}