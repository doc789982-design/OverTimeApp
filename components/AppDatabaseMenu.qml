import QtQuick
import QtQuick.Controls
import Qt5Compat.GraphicalEffects
import "."

Menu {
    id: root
    
    topPadding: AppTheme.spaceXS
    bottomPadding: AppTheme.spaceXS
    
    // Используем AppMenuItem как делегат для единообразного стиля
    delegate: AppMenuItem {}
    
    background: Rectangle {
        implicitWidth: 260
        color: AppTheme.bgElevated
        border.color: AppTheme.borderDivider
        border.width: 1
        radius: AppTheme.radiusMedium
        
        layer.enabled: true
        layer.effect: DropShadow {
            transparentBorder: true
            color: AppTheme.shadowColor
            radius: AppTheme.shadowL3Blur
            verticalOffset: AppTheme.shadowL3Y
            samples: 17
        }
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