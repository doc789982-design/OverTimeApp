import QtQuick
import QtQuick.Controls
import Qt5Compat.GraphicalEffects

Item {
    id: root
    implicitHeight: 36
    
    property int startMinutes: 480  
    property int endMinutes: 1200   
    property string draggingHandle: "none"

    function formatTime(mins) {
        var h = Math.floor(mins / 60)
        var m = Math.floor(mins % 60)
        return (h < 10 ? "0" + h : h.toString()) + ":" + (m < 10 ? "0" + m : m.toString())
    }

    // ==========================================
    // 1. ФОНОВЫЙ ТРЕК
    // ==========================================
    Rectangle {
        id: track
        anchors.centerIn: parent
        width: parent.width - 24 
        height: 4 
        color: AppTheme.borderDivider // Строгий системный серый
        radius: 2

        Rectangle { 
            visible: root.startMinutes <= root.endMinutes
            color: AppTheme.accentBrand
            height: parent.height; radius: 2
            x: (root.startMinutes / 1440) * track.width
            width: ((root.endMinutes - root.startMinutes) / 1440) * track.width 
        }
        Rectangle { 
            visible: root.startMinutes > root.endMinutes
            color: AppTheme.accentBrand
            height: parent.height; radius: 2
            x: 0; width: (root.endMinutes / 1440) * track.width 
        }
        Rectangle { 
            visible: root.startMinutes > root.endMinutes
            color: AppTheme.accentBrand
            height: parent.height; radius: 2
            x: (root.startMinutes / 1440) * track.width; width: track.width - x 
        }
    }

    // ==========================================
    // 2. ПОЛЗУНОК НАЧАЛА
    // ==========================================
    Rectangle {
        id: handleStart
        width: 16; height: 16; radius: AppTheme.radiusPill
        
        // Цвет: синий, если тянем, иначе белый
        color: root.draggingHandle === "start" ? AppTheme.accentBrand : AppTheme.bgSurface
        border.color: AppTheme.accentBrand
        border.width: root.draggingHandle === "start" ? 0 : 2
        
        // Squish анимация
        scale: root.draggingHandle === "start" ? AppTheme.scaleActive : (mouseArea.containsMouse ? 1.1 : 1.0)
        
        Behavior on scale { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeStandard } }
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

        x: track.x + (root.startMinutes / 1440) * track.width - (width / 2)
        y: track.y + (track.height / 2) - (height / 2)

        // ТУЛТИП
        Rectangle {
            property bool isVisible: root.draggingHandle === "start" || mouseArea.containsMouse
            
            width: timeTextStart.width + 12
            height: 22
            radius: AppTheme.radiusSmall
            color: AppTheme.bgElevated
            border.color: AppTheme.borderDivider
            border.width: 1
            
            anchors.horizontalCenter: parent.horizontalCenter
            anchors.bottom: parent.top
            anchors.bottomMargin: isVisible ? AppTheme.spaceS : 0 
            
            opacity: isVisible ? 1.0 : 0.0
            
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeColor } }
            Behavior on anchors.bottomMargin { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeEnter } }

            layer.enabled: true
            layer.effect: DropShadow {
                transparentBorder: true
                color: AppTheme.shadowColor
                radius: AppTheme.shadowL5Blur
                verticalOffset: AppTheme.shadowL5Y
                samples: 17
            }

            Text {
                id: timeTextStart
                anchors.centerIn: parent
                text: root.formatTime(root.startMinutes)
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeSmall
                font.weight: AppTheme.weightMedium
            }
        }
    }

    // ==========================================
    // 3. ПОЛЗУНОК КОНЦА
    // ==========================================
    Rectangle {
        id: handleEnd
        width: 16; height: 16; radius: AppTheme.radiusPill
        
        color: root.draggingHandle === "end" ? AppTheme.accentBrand : AppTheme.bgSurface
        border.color: AppTheme.accentBrand
        border.width: root.draggingHandle === "end" ? 0 : 2
        
        scale: root.draggingHandle === "end" ? AppTheme.scaleActive : (mouseArea.containsMouse ? 1.1 : 1.0)
        
        Behavior on scale { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeStandard } }
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

        x: track.x + (root.endMinutes / 1440) * track.width - (width / 2)
        y: track.y + (track.height / 2) - (height / 2)

        // ТУЛТИП
        Rectangle {
            property bool isVisible: root.draggingHandle === "end" || mouseArea.containsMouse
            
            width: timeTextEnd.width + 12
            height: 22
            radius: AppTheme.radiusSmall
            color: AppTheme.bgElevated
            border.color: AppTheme.borderDivider
            border.width: 1
            
            anchors.horizontalCenter: parent.horizontalCenter
            anchors.bottom: parent.top
            anchors.bottomMargin: isVisible ? AppTheme.spaceS : 0 
            
            opacity: isVisible ? 1.0 : 0.0
            
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeColor } }
            Behavior on anchors.bottomMargin { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeEnter } }

            layer.enabled: true
            layer.effect: DropShadow {
                transparentBorder: true
                color: AppTheme.shadowColor
                radius: AppTheme.shadowL5Blur
                verticalOffset: AppTheme.shadowL5Y
                samples: 17
            }

            Text {
                id: timeTextEnd
                anchors.centerIn: parent
                text: root.formatTime(root.endMinutes)
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeSmall
                font.weight: AppTheme.weightMedium
            }
        }
    }

    // ==========================================
    // 4. ЛОГИКА МЫШИ
    // ==========================================
    MouseArea {
        id: mouseArea
        anchors.fill: parent
        hoverEnabled: true

        onPressed: function(mouse) {
            var clickMins = (mapToItem(track, mouse.x, mouse.y).x / track.width) * 1440
            var distStart = Math.abs(clickMins - root.startMinutes)
            var distEnd = Math.abs(clickMins - root.endMinutes)
            
            if (distStart <= distEnd) { 
                root.draggingHandle = "start"
                root.updateTime(mapToItem(track, mouse.x, mouse.y).x, "start") 
            } else { 
                root.draggingHandle = "end"
                root.updateTime(mapToItem(track, mouse.x, mouse.y).x, "end") 
            }
        }
        
        onPositionChanged: function(mouse) {
            if (root.draggingHandle === "start") root.updateTime(mapToItem(track, mouse.x, mouse.y).x, "start")
            else if (root.draggingHandle === "end") root.updateTime(mapToItem(track, mouse.x, mouse.y).x, "end")
        }
        
        onReleased: function() { 
            root.draggingHandle = "none" 
        }
    }

    function updateTime(pixelX, handleType) {
        if (track.width === 0) return
        var newMins = (pixelX / track.width) * 1440
        if (newMins < 0) newMins = 0
        if (newMins > 1440) newMins = 1440
        
        newMins = Math.round(newMins / 15) * 15
        if (newMins === 1440) newMins = 0 
        
        if (handleType === "start") root.startMinutes = newMins
        else root.endMinutes = newMins
    }
}