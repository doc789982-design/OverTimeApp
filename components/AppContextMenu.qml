import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

Menu {
    id: root
    
    topPadding: AppTheme.spaceXS
    bottomPadding: AppTheme.spaceXS
    
    delegate: MenuItem {
        id: menuItem
        
        contentItem: Row {
            spacing: AppTheme.spaceS
            leftPadding: AppTheme.spaceM
            rightPadding: AppTheme.spaceM
            
            // Иконка (если есть)
            IconImage {
                visible: menuItem.icon.source != ""
                source: menuItem.icon.source
                width: AppTheme.iconMedium
                height: AppTheme.iconMedium
                color: menuItem.highlighted ? AppTheme.textOnSoft : AppTheme.textSecondary
                anchors.verticalCenter: parent.verticalCenter
            }
            
            // Текст
            Text {
                text: menuItem.text
                color: menuItem.highlighted ? AppTheme.accentBrand : AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
                font.weight: menuItem.highlighted ? AppTheme.weightBold : AppTheme.weightMedium
                anchors.verticalCenter: parent.verticalCenter
            }
        }
        
        background: Rectangle {
            implicitWidth: 220
            implicitHeight: 36
            color: menuItem.highlighted ? AppTheme.stateHover : "transparent"
            radius: AppTheme.radiusSmall
            
            Behavior on color {
                ColorAnimation { duration: AppTheme.durMicro }
            }
        }
    }
    
    background: Rectangle {
        color: "#FF0000"
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