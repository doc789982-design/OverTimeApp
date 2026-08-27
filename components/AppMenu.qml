import QtQuick
import QtQuick.Controls
import Qt5Compat.GraphicalEffects

Menu {
    id: control

    z: AppTheme.zDropdown
    transformOrigin: Item.TopLeft
    padding: 0
    topPadding: AppTheme.spaceXS
    bottomPadding: AppTheme.spaceXS
    leftPadding: AppTheme.spaceXXS
    rightPadding: AppTheme.spaceXXS

    enter: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
            NumberAnimation { property: "scale"; from: 0.96; to: 1.0; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
        }
    }
    exit: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: AppTheme.durMicro; easing.type: AppTheme.easeExit }
            NumberAnimation { property: "scale"; from: 1.0; to: 0.96; duration: AppTheme.durMicro; easing.type: AppTheme.easeExit }
        }
    }

    background: Rectangle {
        implicitWidth: 200
        color: AppTheme.bgElevated
        radius: AppTheme.radiusMedium
        border.color: AppTheme.borderDivider
        border.width: 1
        AppShadow { level: 2 }
    }
}
