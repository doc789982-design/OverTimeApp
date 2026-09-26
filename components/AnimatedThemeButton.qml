import QtQuick
import QtQuick.Controls

Item {
    id: root
    
    implicitWidth: 32
    implicitHeight: 32

    property color iconColor: AppTheme.textSecondary
    property color iconHoverColor: AppTheme.textPrimary
    property int bgRadius: AppTheme.radiusPill

    property bool isDark: backend.isDarkTheme

    signal clicked()

    Rectangle {
        anchors.fill: parent
        radius: root.bgRadius
        color: mouseArea.pressed ? AppTheme.statePress : (mouseArea.containsMouse ? AppTheme.stateHover : "transparent")
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    Canvas {
        id: canvas
        width: AppTheme.iconLarge
        height: AppTheme.iconLarge
        anchors.centerIn: parent
        
        property real morphProgress: isDark ? 1.0 : 0.0 
        property real rayOffset: 0 
        property color drawColor: mouseArea.containsMouse ? root.iconHoverColor : root.iconColor

        onMorphProgressChanged: requestPaint()
        onRayOffsetChanged: requestPaint()
        onDrawColorChanged: requestPaint()

        Behavior on morphProgress { NumberAnimation { duration: 600; easing.type: Easing.InOutBack } }
        Behavior on rayOffset { NumberAnimation { duration: AppTheme.durSlow; easing.type: AppTheme.easeEnter } }

        onPaint: {
            var ctx = getContext("2d"); ctx.clearRect(0, 0, width, height)
            let cx = width / 2; let cy = height / 2; let baseRadius = 4.5 + (4.0 * morphProgress)
            ctx.fillStyle = drawColor; ctx.strokeStyle = drawColor; ctx.lineCap = "round"; ctx.lineWidth = 2.5
            
            ctx.beginPath(); ctx.arc(cx, cy, baseRadius, 0, Math.PI * 2); ctx.fill()

            if (morphProgress > 0) {
                ctx.globalCompositeOperation = "destination-out"
                ctx.beginPath()
                let cutX = cx + (4.0 * morphProgress); let cutY = cy - (4.0 * morphProgress); let cutRadius = baseRadius - (0.5 * morphProgress) + 0.2 
                ctx.arc(cutX, cutY, cutRadius, 0, Math.PI * 2); ctx.fill()
                ctx.globalCompositeOperation = "source-over"
            }

            if (morphProgress < 1.0) {
                ctx.globalAlpha = 1.0 - (morphProgress * 1.5); if (ctx.globalAlpha < 0) ctx.globalAlpha = 0
                let rayLength = 2.5; let rayStart = 6.8 - rayOffset; let rayEnd = rayStart + rayLength
                let actualStart = rayStart - (morphProgress * 4.0); let actualEnd = rayEnd - (morphProgress * 4.0)

                if (actualEnd > actualStart) {
                    for (let i = 0; i < 8; i++) {
                        let angle = (i * Math.PI) / 4 - (Math.PI / 2)
                        let x1 = cx + Math.cos(angle) * actualStart; let y1 = cy + Math.sin(angle) * actualStart
                        let x2 = cx + Math.cos(angle) * actualEnd;   let y2 = cy + Math.sin(angle) * actualEnd
                        ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke()
                    }
                }
                ctx.globalAlpha = 1.0
            }
        }
    }

    SequentialAnimation {
        id: moonWobble
        loops: Animation.Infinite
        NumberAnimation { target: canvas; property: "rotation"; to: -15; duration: AppTheme.durSlow; easing.type: Easing.InOutSine }
        NumberAnimation { target: canvas; property: "rotation"; to: 5; duration: AppTheme.durSlow; easing.type: Easing.InOutSine }
        NumberAnimation { target: canvas; property: "rotation"; to: 0; duration: AppTheme.durSlow; easing.type: Easing.OutSine }
    }

    SequentialAnimation {
        id: sunPulse
        loops: Animation.Infinite
        NumberAnimation { target: canvas; property: "rayOffset"; to: 1.5; duration: 600; easing.type: Easing.InOutSine }
        NumberAnimation { target: canvas; property: "rayOffset"; to: 0; duration: 600; easing.type: Easing.InOutSine }
    }

    MouseArea {
        id: mouseArea
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: Qt.PointingHandCursor
        
        onEntered: { if (root.isDark) moonWobble.restart(); else sunPulse.restart() }
        onExited: { moonWobble.stop(); sunPulse.stop(); canvas.rotation = 0; canvas.rayOffset = 0 }
        onClicked: { moonWobble.stop(); sunPulse.stop(); canvas.rotation = 0; canvas.rayOffset = 0; root.clicked(); restartTimer.start() }
    }

    Timer {
        id: restartTimer; interval: 600 
        onTriggered: { if (mouseArea.containsMouse) { if (root.isDark) moonWobble.restart(); else sunPulse.restart() } }
    }
}