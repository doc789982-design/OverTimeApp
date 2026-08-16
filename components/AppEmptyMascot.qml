import QtQuick
import QtQuick.Controls

Item {
    id: root
    width: 200
    height: 200

    // ══════════════════════════════════════════════════
    //  АНИМИРУЕМЫЕ СВОЙСТВА
    // ══════════════════════════════════════════════════
    property real breathY:   0.0
    property real breathS:   1.0
    property real scanX:     0.0
    property real glowPulse: 0.6
    property real tiltAngle: 0.0

    // ОПТИМИЗАЦИЯ: анимации маскота живут только когда он реально
    // на экране. Раньше 5 вечных анимаций крутились даже когда
    // маскот был скрыт или окно свернуто — видеокарта работала впустую.
    readonly property bool animsAlive: root.visible && root.opacity > 0
                                       && root.Window.window !== null
                                       && root.Window.window.visible

    // ── Дыхание ───────────────────────────────────────
    SequentialAnimation on breathY {
        loops: Animation.Infinite
        running: root.animsAlive
        NumberAnimation { to: -6; duration: 3000; easing.type: Easing.InOutSine }
        NumberAnimation { to:  0; duration: 3000; easing.type: Easing.InOutSine }
    }
    SequentialAnimation on breathS {
        loops: Animation.Infinite
        running: root.animsAlive
        NumberAnimation { to: 1.025; duration: 3000; easing.type: Easing.InOutSine }
        NumberAnimation { to: 1.000; duration: 3000; easing.type: Easing.InOutSine }
    }

    // ── Покачивание ───────────────────────────────────
    SequentialAnimation on tiltAngle {
        loops: Animation.Infinite
        running: root.animsAlive
        NumberAnimation { to: -2.5; duration: 4000; easing.type: Easing.InOutSine }
        NumberAnimation { to:  2.5; duration: 4000; easing.type: Easing.InOutSine }
    }

    // ── Пульсация свечения ────────────────────────────
    SequentialAnimation on glowPulse {
        loops: Animation.Infinite
        running: root.animsAlive
        NumberAnimation { to: 1.0; duration: 2000; easing.type: Easing.InOutSine }
        NumberAnimation { to: 0.5; duration: 2000; easing.type: Easing.InOutSine }
    }

    // ── Сканер ────────────────────────────────────────
    SequentialAnimation on scanX {
        loops: Animation.Infinite
        running: root.animsAlive
        NumberAnimation { to:  7; duration: 900; easing.type: Easing.InOutSine }
        NumberAnimation { to: -7; duration: 900; easing.type: Easing.InOutSine }
    }

    // ══════════════════════════════════════════════════
    //  ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ ФОРМЫ ЩИТА
    // ══════════════════════════════════════════════════
    function paintShield(ctx, w, h, fill, stroke, lineW) {
        var r  = 14
        var bh = h * 0.70
        ctx.beginPath()
        ctx.moveTo(r, 0)
        ctx.lineTo(w - r, 0)
        ctx.quadraticCurveTo(w, 0, w, r)
        ctx.lineTo(w, bh)
        ctx.quadraticCurveTo(w, h * 0.87, w / 2, h)
        ctx.quadraticCurveTo(0, h * 0.87, 0, bh)
        ctx.lineTo(0, r)
        ctx.quadraticCurveTo(0, 0, r, 0)
        ctx.closePath()
        ctx.fillStyle = fill
        ctx.fill()
        if (lineW > 0) {
            ctx.strokeStyle = stroke
            ctx.lineWidth   = lineW
            ctx.stroke()
        }
    }

    // ══════════════════════════════════════════════════
    //  КОРЕНЬ МАСКОТА
    // ══════════════════════════════════════════════════
    Item {
        id: shieldRoot
        width:  110
        height: 130
        anchors.centerIn: parent
        y: root.breathY
        scale: root.breathS
        rotation: root.tiltAngle
        transformOrigin: Item.Center

        // ══════════════════════════════════════════════
        //  1. ТЕНЬ
        // ══════════════════════════════════════════════
        Canvas {
            id: shadowCanvas
            width:  shieldRoot.width
            height: shieldRoot.height
            anchors.centerIn: parent
            anchors.verticalCenterOffset: 10
            opacity: 0.15
            z: 0

            onPaint: {
                var ctx = getContext("2d")
                ctx.clearRect(0, 0, width, height)
                root.paintShield(ctx, width, height, "#000000", "#000000", 0)
            }
        }

        // ══════════════════════════════════════════════
        //  2. СВЕЧЕНИЕ
        // ══════════════════════════════════════════════
        Canvas {
            id: glowCanvas
            width:  shieldRoot.width  + 32
            height: shieldRoot.height + 32
            anchors.centerIn: parent
            opacity: root.glowPulse * 0.30
            z: 1

            onPaint: {
                var ctx = getContext("2d")
                ctx.clearRect(0, 0, width, height)
                ctx.save()
                ctx.shadowColor = AppTheme.accentBrand
                ctx.shadowBlur  = 24
                ctx.translate(16, 16)
                root.paintShield(
                    ctx,
                    shieldRoot.width,
                    shieldRoot.height,
                    AppTheme.accentBrand,
                    AppTheme.accentBrand,
                    0
                )
                ctx.restore()
            }

            Connections {
                target: AppTheme
                function onIsDarkChanged() { glowCanvas.requestPaint() }
            }
        }

        // Перерисовка свечения при изменении пульсации
        onOpacityChanged: glowCanvas.requestPaint()

        // ══════════════════════════════════════════════
        //  3. ОСНОВНОЙ ЩИТ
        // ══════════════════════════════════════════════
        Canvas {
            id: shieldCanvas
            width:  shieldRoot.width
            height: shieldRoot.height
            anchors.fill: parent
            z: 2

            onPaint: {
                var ctx = getContext("2d")
                ctx.clearRect(0, 0, width, height)
                root.paintShield(
                    ctx,
                    width,
                    height,
                    AppTheme.bgSurface,
                    AppTheme.accentBrand,
                    2.0
                )
            }

            Connections {
                target: AppTheme
                function onIsDarkChanged() { shieldCanvas.requestPaint() }
            }
        }

        // ══════════════════════════════════════════════
        //  4. ГОРИЗОНТАЛЬНЫЙ АКЦЕНТ
        // ══════════════════════════════════════════════
        Rectangle {
            width:  parent.width * 0.58
            height: 1
            radius: 1
            anchors.horizontalCenter: parent.horizontalCenter
            anchors.top: parent.top
            anchors.topMargin: parent.height * 0.40
            color: AppTheme.accentBrand
            opacity: 0.28
            z: 3
        }

        // ══════════════════════════════════════════════
        //  5. ГЛАЗ-СКАНЕР
        // ══════════════════════════════════════════════
        Item {
            id: eyeGroup
            width:  52
            height: 52
            anchors.horizontalCenter: parent.horizontalCenter
            anchors.top: parent.top
            anchors.topMargin: parent.height * 0.13
            z: 4

            // Внешнее кольцо
            Rectangle {
                id: eyeRing
                anchors.fill: parent
                radius: width / 2
                color: "transparent"
                border.color: AppTheme.accentBrand
                border.width: 1.5
                opacity: 0.55
            }

            // Среднее кольцо (пульсирует)
            Rectangle {
                width:  parent.width  - 12
                height: parent.height - 12
                radius: width / 2
                anchors.centerIn: parent
                color: "transparent"
                border.color: AppTheme.accentBrand
                border.width: 1
                opacity: root.glowPulse * 0.35
            }

            // Зрачок
            Item {
                id: pupilContainer
                anchors.fill: parent

                property real lookX: 0
                property real lookY: 0

                Behavior on lookX {
                    NumberAnimation { duration: 420; easing.type: Easing.OutCubic }
                }
                Behavior on lookY {
                    NumberAnimation { duration: 420; easing.type: Easing.OutCubic }
                }

                Rectangle {
                    id: pupil
                    width:  20
                    height: 20
                    radius: width / 2
                    color: AppTheme.accentBrand
                    anchors.centerIn: parent
                    anchors.horizontalCenterOffset: pupilContainer.lookX
                    anchors.verticalCenterOffset:   pupilContainer.lookY

                    // Блик
                    Rectangle {
                        width: 5; height: 5; radius: 3
                        anchors.top:  parent.top;  anchors.topMargin:  3
                        anchors.left: parent.left; anchors.leftMargin: 3
                        color: "#FFFFFF"
                        opacity: 0.55
                    }
                }
            }

            // Линия сканера
            Rectangle {
                width:  pupil.width - 6
                height: 1.5
                radius: 1
                anchors.horizontalCenter: parent.horizontalCenter
                anchors.verticalCenter:   parent.verticalCenter
                anchors.verticalCenterOffset: root.scanX
                color: "#FFFFFF"
                opacity: 0.48
                z: 5
            }

            // Верхнее веко
            Rectangle {
                id: lidTop
                width:  parent.width + 4
                height: 0
                radius: parent.width / 2
                anchors.top: parent.top
                anchors.topMargin: -2
                anchors.horizontalCenter: parent.horizontalCenter
                color: AppTheme.bgSurface
                z: 6
            }

            // Нижнее веко
            Rectangle {
                id: lidBot
                width:  parent.width + 4
                height: 0
                radius: parent.width / 2
                anchors.bottom: parent.bottom
                anchors.bottomMargin: -2
                anchors.horizontalCenter: parent.horizontalCenter
                color: AppTheme.bgSurface
                z: 6
            }
        }

        // ══════════════════════════════════════════════
        //  6. ПОДПИСЬ
        // ══════════════════════════════════════════════
        Text {
            id: codeLabel
            text: "OverTimeTab"
            font.pixelSize: 9
            font.letterSpacing: 3
            color: AppTheme.accentBrand
            opacity: 0.48
            anchors.horizontalCenter: parent.horizontalCenter
            anchors.top: parent.top
            anchors.topMargin: parent.height * 0.63
            z: 4
        }

        Rectangle {
            width:  52
            height: 1
            radius: 1
            anchors.horizontalCenter: parent.horizontalCenter
            anchors.top: codeLabel.bottom
            anchors.topMargin: 4
            color: AppTheme.accentBrand
            opacity: 0.18
            z: 4
        }

    } // shieldRoot

    // ══════════════════════════════════════════════════
    //  АНИМАЦИИ
    // ══════════════════════════════════════════════════

    // ── Моргание ──────────────────────────────────────
    SequentialAnimation {
        id: blinkAnim
        ParallelAnimation {
            NumberAnimation { target: lidTop; property: "height"; to: 27; duration: 85  }
            NumberAnimation { target: lidBot; property: "height"; to: 27; duration: 85  }
        }
        PauseAnimation { duration: 50 }
        ParallelAnimation {
            NumberAnimation { target: lidTop; property: "height"; to: 0; duration: 130; easing.type: Easing.OutCubic }
            NumberAnimation { target: lidBot; property: "height"; to: 0; duration: 130; easing.type: Easing.OutCubic }
        }
    }

    Timer {
        interval: 3600
        running: root.animsAlive
        repeat: true
        onTriggered: {
            interval = 2600 + Math.random() * 3400
            blinkAnim.start()
        }
    }

    // ── Перевод взгляда ───────────────────────────────
    SequentialAnimation {
        id: lookAnim

        ScriptAction {
            script: {
                pupilContainer.lookX = (Math.random() > 0.5 ? 1 : -1) * (3 + Math.random() * 5)
                pupilContainer.lookY = (Math.random() > 0.5 ? 1 : -1) * (2 + Math.random() * 3)
            }
        }

        PauseAnimation { duration: 800 }

        ParallelAnimation {
            NumberAnimation {
                target: pupilContainer; property: "lookX"
                to: 0; duration: 500; easing.type: Easing.InOutCubic
            }
            NumberAnimation {
                target: pupilContainer; property: "lookY"
                to: 0; duration: 500; easing.type: Easing.InOutCubic
            }
        }
    }

    Timer {
        interval: 4000
        running: root.animsAlive
        repeat: true
        onTriggered: {
            interval = 3200 + Math.random() * 4200
            lookAnim.start()
        }
    }

    // ── Тревога ───────────────────────────────────────
    SequentialAnimation {
        id: alertAnim
        NumberAnimation { target: eyeRing; property: "opacity"; to: 1.0;  duration: 110 }
        NumberAnimation { target: eyeRing; property: "opacity"; to: 0.15; duration: 110 }
        NumberAnimation { target: eyeRing; property: "opacity"; to: 1.0;  duration: 110 }
        NumberAnimation { target: eyeRing; property: "opacity"; to: 0.15; duration: 110 }
        NumberAnimation { target: eyeRing; property: "opacity"; to: 0.55; duration: 350; easing.type: Easing.OutCubic }
    }

    Timer {
        interval: 9000
        running: root.animsAlive
        repeat: true
        onTriggered: {
            interval = 7000 + Math.random() * 7000
            alertAnim.start()
        }
    }

    // ── Сосредоточенность ─────────────────────────────
    SequentialAnimation {
        id: focusAnim

        NumberAnimation {
            target: root; property: "tiltAngle"
            to: -4.5; duration: 650; easing.type: Easing.OutCubic
        }
        ParallelAnimation {
            NumberAnimation { target: pupilContainer; property: "lookX"; to: -4; duration: 350; easing.type: Easing.OutCubic }
            NumberAnimation { target: pupilContainer; property: "lookY"; to:  2; duration: 350; easing.type: Easing.OutCubic }
        }
        PauseAnimation { duration: 1400 }
        ParallelAnimation {
            NumberAnimation { target: root;           property: "tiltAngle"; to: 0; duration: 900; easing.type: Easing.OutElastic }
            NumberAnimation { target: pupilContainer; property: "lookX";     to: 0; duration: 600; easing.type: Easing.InOutCubic }
            NumberAnimation { target: pupilContainer; property: "lookY";     to: 0; duration: 600; easing.type: Easing.InOutCubic }
        }
    }

    Timer {
        interval: 13000
        running: root.animsAlive
        repeat: true
        onTriggered: {
            interval = 10000 + Math.random() * 9000
            focusAnim.start()
        }
    }
}