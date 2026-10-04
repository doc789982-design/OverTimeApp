import QtQuick
import QtQuick.Controls

// ============================================================
// КНОПКА ОБНОВЛЕНИЯ В ШАПКЕ — SVG-иконки, бесшовная хореография
//
// Принцип: окончание анимации одной иконки = начало следующей.
// Все переходы проходят через одну и ту же «позу» — схлопнутую
// в вертикальную линию (xScale = 0): одна иконка складывается
// в неё, следующая из неё раскрывается (с лёгким перелётом).
// Стык состояний непрерывен, без подмен «в лоб».
//
//   покой       — «перезагрузка» (refresh.svg), серая как шапка
//   клик        — прокрут на полный оборот (проверка обновлений)
//   нет обновл. — просто остановилась
//   есть        — прокрутившись, складывается и раскрывается
//                 стрелкой «скачать» (download.svg, зелёная)
//   загрузка    — снова прокрут, «скачать» складывается
//                 и раскрывается кольцом, которое заполняется
//                 прогрессом (тонкая дуга)
//   загружено   — кольцо ДОЗАЛИВАЕТСЯ до полного (его конец)
//                 и зеленеет, внутри вырастает стрелка
//                 «установить» — начало следующего состояния
//                 то же кольцо, разрыва нет
//
// Слоями управляет только switchVis(): у слоёв нет биндингов
// на xScale/opacity (императивные анимации их бы отцепили).
// ============================================================
Item {
    id: root
    width: 46
    height: parent ? parent.height : 46

    // ---- Состояния от Backend ----
    readonly property bool hasUpdate: backend.remoteUpdateAvailable
    readonly property bool downloading: backend.remoteDownloading
    readonly property int progress: backend.remoteDownloadProgress
    readonly property bool ready: backend.updateReady

    readonly property bool idleCheck: !root.hasUpdate && !root.downloading && !root.ready

    property string vis: "refresh"     // refresh | download | ring | install
    property string prevVis: "refresh"

    // ---- Целевой слой по состоянию ----
    readonly property string targetVis: root.downloading ? "ring"
                                       : root.ready ? "install"
                                       : root.hasUpdate ? "download"
                                       : "refresh"

    // ---- Цвет иконок: состояние + ховер ----
    readonly property color stateColor: root.downloading ? AppTheme.accentBrand
                                   : (root.hasUpdate || root.ready) ? AppTheme.accentSuccess
                                   : (hover.containsMouse ? AppTheme.textPrimary : AppTheme.textSecondary)

    readonly property string tipText: {
        if (downloading)
            return "Загрузка обновления… " + progress + "%"
        if (ready)
            return "Обновление загружено. Нажмите, чтобы установить."
        if (hasUpdate)
            return "Доступна новая версия программы. Нажмите, чтобы загрузить."
        return "Проверить обновления"
    }

    // ---- Искры-стрелочки: есть что скачать/установить ----
    UpdateSparkles {
        anchors.centerIn: parent
        z: 10
        active: (root.hasUpdate && !root.downloading) || root.ready
    }

    // ---- Подложка (hover/нажатие) ----
    Rectangle {
        anchors.fill: parent
        color: hover.pressed ? AppTheme.statePress : (hover.containsMouse ? AppTheme.stateHover : "transparent")
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    // ---- Вращающийся контейнер иконок ----
    Item {
        id: spinner
        anchors.centerIn: parent
        width: 28; height: 28
        rotation: 0

        IconImage {
            id: iRefresh
            objectName: "updRefresh"
            anchors.centerIn: parent
            source: "../icons/refresh.svg"
            width: AppTheme.iconMedium + 2; height: AppTheme.iconMedium + 2
            color: root.stateColor
            opacity: 1
            transform: Scale { id: rX; objectName: "updRefreshX"; origin.x: iRefresh.width / 2; origin.y: iRefresh.height / 2; xScale: 1 }
            Behavior on color { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }
        }

        IconImage {
            id: iDownload
            objectName: "updDownload"
            anchors.centerIn: parent
            source: "../icons/download.svg"
            width: AppTheme.iconMedium + 2; height: AppTheme.iconMedium + 2
            color: root.stateColor
            opacity: 0
            transform: Scale { id: dX; objectName: "updDownloadX"; origin.x: iDownload.width / 2; origin.y: iDownload.height / 2; xScale: 0 }
            Behavior on color { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }
        }

        Item {
            id: iRing
            objectName: "updRing"
            anchors.centerIn: parent
            width: 26; height: 26
            opacity: 0
            transform: Scale { id: gX; objectName: "updRingX"; origin.x: iRing.width / 2; origin.y: iRing.height / 2; xScale: 0 }

            // сглаженный прогресс: скачки процентов не щёлкают;
            // в готовности дозаливается до полного — бесшовно
            property real fill: root.downloading ? root.progress / 100 : (root.ready ? 1 : 0)
            Behavior on fill { NumberAnimation { duration: 250; easing.type: Easing.OutCubic } }

            property color arcColor: root.ready ? root.stateColor : AppTheme.accentBrand
            Behavior on arcColor { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }

            Canvas {
                id: arcCanvas
                anchors.fill: parent
                onPaint: {
                    var ctx = getContext("2d")
                    ctx.reset()
                    var w = width, h = height, r = (w - 4) / 2, cx = w / 2, cy = h / 2
                    ctx.lineWidth = 2.4
                    ctx.lineCap = "round"
                    ctx.beginPath()
                    ctx.arc(cx, cy, r, 0, Math.PI * 2)
                    ctx.strokeStyle = Qt.rgba(root.stateColor.r, root.stateColor.g, root.stateColor.b, 0.25)
                    ctx.stroke()
                    ctx.beginPath()
                    var p = Math.min(iRing.fill, 1)
                    if (p > 0.004)
                        ctx.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + Math.PI * 2 * p)
                    ctx.strokeStyle = iRing.arcColor
                    ctx.stroke()
                }
                Connections {
                    target: iRing
                    function onFillChanged() { arcCanvas.requestPaint() }
                    function onArcColorChanged() { arcCanvas.requestPaint() }
                }
                Connections {
                    target: root
                    function onStateColorChanged() { arcCanvas.requestPaint() }
                }
            }
        }

        // «установить»: стрелка, вырастающая ВНУТРИ готового кольца
        IconImage {
            id: iInstall
            objectName: "updInstall"
            anchors.centerIn: parent
            source: "../icons/arrow_down.svg"
            width: AppTheme.iconMedium - 2; height: AppTheme.iconMedium - 2
            color: root.stateColor
            scale: root.vis === "install" ? 1 : 0.3
            opacity: root.vis === "install" ? 1 : 0
            Behavior on scale { SpringAnimation { spring: 5.0; damping: 0.35; mass: 1.1 } }
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
            Behavior on color { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }
        }
    }

    // ---- Анимации слоёв ----
    // складывание в линию — конец состояния
    ParallelAnimation {
        id: rHide
        NumberAnimation { target: rX; property: "xScale"; to: 0; duration: 130; easing.type: Easing.InQuad }
        NumberAnimation { target: iRefresh; property: "opacity"; to: 0; duration: 130; easing.type: Easing.InQuad }
    }
    ParallelAnimation {
        id: dHide
        NumberAnimation { target: dX; property: "xScale"; to: 0; duration: 130; easing.type: Easing.InQuad }
        NumberAnimation { target: iDownload; property: "opacity"; to: 0; duration: 130; easing.type: Easing.InQuad }
    }
    ParallelAnimation {
        id: gHide
        NumberAnimation { target: gX; property: "xScale"; to: 0; duration: 130; easing.type: Easing.InQuad }
        NumberAnimation { target: iRing; property: "opacity"; to: 0; duration: 130; easing.type: Easing.InQuad }
    }
    // раскрытие из линии — начало следующего состояния, с перелётом
    ParallelAnimation {
        id: rShow
        NumberAnimation { target: rX; property: "xScale"; to: 1; duration: 200; easing.type: Easing.OutBack; easing.overshoot: 1.1 }
        NumberAnimation { target: iRefresh; property: "opacity"; to: 1; duration: 150; easing.type: Easing.OutQuad }
    }
    ParallelAnimation {
        id: dShow
        NumberAnimation { target: dX; property: "xScale"; to: 1; duration: 200; easing.type: Easing.OutBack; easing.overshoot: 1.1 }
        NumberAnimation { target: iDownload; property: "opacity"; to: 1; duration: 150; easing.type: Easing.OutQuad }
    }
    ParallelAnimation {
        id: gShow
        NumberAnimation { target: gX; property: "xScale"; to: 1; duration: 200; easing.type: Easing.OutBack; easing.overshoot: 1.1 }
        NumberAnimation { target: iRing; property: "opacity"; to: 1; duration: 150; easing.type: Easing.OutQuad }
    }

    // прокрут на полный оборот с естественным торможением
    NumberAnimation {
        id: spinAnim
        target: spinner
        property: "rotation"
        duration: 750
        easing.type: Easing.InOutCubic
    }

    function spin() {
        spinAnim.to = spinner.rotation + 360
        spinAnim.restart()
    }

    function switchVis(to) {
        if (root.vis === "refresh") rHide.restart()
        else if (root.vis === "download") dHide.restart()
        else if (root.vis === "ring" || root.vis === "install") gHide.restart()
        root.vis = to
        if (to === "refresh") rShow.restart()
        else if (to === "download") dShow.restart()
        else if (to === "ring") gShow.restart()
        // "install" не раскрывает новый слой: кольцо уже на месте
        // (бесшовность), вырастает только стрелка — её ведёт vis
    }

    // Смена состояния бэкенда -> хореография.
    // Антидребезг: бэкенд может сообщить «загрузка кончилась» и «готово»
    // двумя событиями подряд — без задержки между ними проигрался бы
    // лишний промежуточный кадр (кольцо сложилось и раскрылось зря).
    Timer {
        id: stateDebounce
        interval: 80
        onTriggered: {
            var to = root.targetVis
            if (to === root.prevVis) return
            var from = root.prevVis
            root.prevVis = to
            if ((from === "refresh" && to === "download") ||
                (from === "download" && to === "ring")) {
                // «прокручивается, складывается, превращается»
                spin()
                Qt.callLater(function() { root.switchVis(to) })
            } else if (from === "ring" && to === "install") {
                // кольцо дозаливается и зеленеет, стрелка растёт внутри
                root.vis = "install"
            } else {
                root.switchVis(to)
            }
        }
    }
    onTargetVisChanged: stateDebounce.restart()

    // ---- Мышь ----
    HoverHandler { cursorShape: (root.ready || (root.hasUpdate && !root.downloading) || root.idleCheck) }
    MouseArea {
        id: hover
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: (root.ready || (root.hasUpdate && !root.downloading) || root.idleCheck)
                     ? Qt.PointingHandCursor : Qt.ArrowCursor
        onClicked: {
            if (root.ready) {
                backend.applyReadyUpdate()
            } else if (root.hasUpdate && !root.downloading) {
                backend.startRemoteDownload()
            } else if (root.idleCheck) {
                spin()
                backend.checkAllUpdateSources()
            }
        }
    }

    // ---- Тултип ----
    AppToolTip {
        anchors.horizontalCenter: parent.horizontalCenter
        anchors.top: parent.bottom
        anchors.topMargin: AppTheme.spaceXXS
        dropDown: true
        isVisible: hover.containsMouse
        text: root.tipText
    }
}
