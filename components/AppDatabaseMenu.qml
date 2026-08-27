import QtQuick
import QtQuick.Controls
import Qt5Compat.GraphicalEffects
import "."

Menu {
    id: root
    
    transformOrigin: Item.TopLeft
    padding: 0
    topPadding: AppTheme.spaceXS
    bottomPadding: AppTheme.spaceXS
    leftPadding: AppTheme.spaceXXS
    rightPadding: AppTheme.spaceXXS
    
    // Используем AppMenuItem как делегат для единообразного стиля
    delegate: AppMenuItem {}
    
    background: Rectangle {
        implicitWidth: 200
        color: AppTheme.bgElevated
        border.color: AppTheme.borderDivider
        border.width: 1
        radius: AppTheme.radiusMedium
        
        // Тень-картинка вместо вычисляемой (Level 3)
        AppShadow { level: 3 }
    }
    
    enter: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
            NumberAnimation { property: "scale"; from: 0.95; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
        }
    }
    
    exit: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
            NumberAnimation { property: "scale"; from: 1.0; to: 0.95; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
        }
    }
}