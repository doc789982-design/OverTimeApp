import QtQuick
import QtQuick.Controls
import Qt5Compat.GraphicalEffects

Menu {
    id: control
    
    // Z-Слой: Поверх модальных окон, но под тултипами
    z: AppTheme.zDropdown

    // ==========================================
    // 1. АНИМАЦИИ (Slide + Fade)
    // ==========================================
    enter: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
            NumberAnimation { property: "scale"; from: 0.95; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
        }
    }
    exit: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: AppTheme.durMicro; easing.type: AppTheme.easeExit }
            NumberAnimation { property: "scale"; from: 1.0; to: 0.95; duration: AppTheme.durMicro; easing.type: AppTheme.easeExit }
        }
    }

    // ==========================================
    // 2. ФОН И ТЕНЬ (Level 2)
    // ==========================================
    background: Rectangle {
        implicitWidth: 220
        color: AppTheme.bgElevated
        radius: AppTheme.radiusMedium
        border.color: AppTheme.borderDivider
        border.width: 1

        // Тень-картинка вместо вычисляемой (Level 2)
        AppShadow { level: 2 }
    }
}