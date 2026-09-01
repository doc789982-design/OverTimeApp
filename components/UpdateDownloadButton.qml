import QtQuick
import QtQuick.Controls
import Qt5Compat.GraphicalEffects

// Кнопка загрузки обновления в шапке окна (рядом со справкой).
//
// Состояния (от серого — к загрузке — к готовности):
//   • нет обновления  — серая неактивная иконка загрузки
//   • обновление есть — иконка плавно «дышит» от серого к зелёному
//   • идёт загрузка   — иконка превращается в кружок, заполняющийся синим
//   • загружено       — кружок зеленеет, пульсирует и возвращается к иконке
Item {
    id: root
    width: 46
    height: parent ? parent.height : 46

    // ---- Состояния от Backend ----
    readonly property bool hasUpdate: backend.remoteUpdateAvailable
    readonly property bool downloading: backend.remoteDownloading
    readonly property int progress: backend.remoteDownloadProgress
    readonly property bool ready: backend.updateReady

    // Тултип зависит от состояния
    readonly property string tipText: {
        if (downloading)
            return "Загрузка обновления… " + progress + "%"
        if (ready)
            return "Обновление загружено"
        if (hasUpdate)
            return "Доступна новая версия программы. Нажмите, чтобы загрузить."
        return "Установлена актуальная версия OverTimeTab"
    }

    // ---- Цвета ----
    readonly property color idleColor: AppTheme.textDisabled
    readonly property color availColor: AppTheme.accentSuccess
    readonly property color ringColor: AppTheme.accentBrand

    property color iconColor: idleColor
    Behavior on iconColor { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }

    // Показывать ли кружок вместо иконки.
    // После пульса готовности возвращаемся к иконке (returnedToIcon).
    property bool returnedToIcon: false
    readonly property bool showRing: root.downloading || (root.ready && !root.returnedToIcon)

    // Клик возможен только когда обновление доступно и ещё не загружено
    readonly property bool clickable: root.hasUpdate && !root.downloading && !root.ready

    // Пульсация «обновление доступно»: серая ↔ зелёная плавно
    SequentialAnimation {
        running: root.hasUpdate && !root.downloading && !root.ready
        loops: Animation.Infinite
        ColorAnimation { target: root; property: "iconColor"; to: root.availColor; duration: 1100; easing.type: Easing.InOutQuad }
        ColorAnimation { target: root; property: "iconColor"; to: root.idleColor; duration: 1100; easing.type: Easing.InOutQuad }
    }

    // ---- Подложка (hover/нажатие) ----
    Rectangle {
        anchors.fill: parent
        color: hover.pressed ? AppTheme.statePress : (hover.containsMouse ? AppTheme.stateHover : "transparent")
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    // ---- Иконка загрузки ----
    IconImage {
        id: downloadIcon
        anchors.centerIn: parent
        source: "../icons/export_arrow.svg"
        width: AppTheme.iconMedium + 2
        height: AppTheme.iconMedium + 2
        color: root.iconColor
    }

    // ---- Кружок прогресса ----
    Item {
        id: ring
        anchors.centerIn: parent
        width: 30
        height: 30
        scale: 1.0
        opacity: 0   // в базовом (idle) состоянии кружок скрыт

        // Насколько заполнен круг (0..1)
        property real fill: root.downloading ? (root.progress / 100) : 1.0
        // Цвет дуги: синий во время загрузки, зелёный когда готово
        property color arcColor: root.ready ? root.availColor : root.ringColor
        Behavior on arcColor { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }

        Canvas {
            anchors.fill: parent
            property real p: ring.fill
            property color stroke: ring.arcColor
            onPaint: {
                var ctx = getContext("2d")
                ctx.reset()
                var w = width, h = height, r = (w - 6) / 2, cx = w / 2, cy = h / 2
                ctx.lineWidth = 3
                ctx.lineCap = "round"
                // фоновый серый круг
                ctx.beginPath()
                ctx.arc(cx, cy, r, 0, Math.PI * 2)
                ctx.strokeStyle = root.idleColor
                ctx.stroke()
                // дуга прогресса
                ctx.beginPath()
                ctx.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + Math.PI * 2 * p)
                ctx.strokeStyle = stroke
                ctx.stroke()
            }
            onPChanged: requestPaint()
            onStrokeChanged: requestPaint()
        }

        // Пульс готовности + динамический возврат к иконке загрузки
        SequentialAnimation {
            running: root.ready && root.showRing
            // плавный зелёный пульс: увеличение внутрь и наружу
            NumberAnimation { target: ring; property: "scale"; to: 1.15; duration: 200; easing.type: Easing.OutCubic }
            NumberAnimation { target: ring; property: "scale"; to: 1.0; duration: 200; easing.type: Easing.InCubic }
            // возврат к иконке (showRing → false, кружок исчезает, иконка проявляется)
            ScriptAction { script: { root.returnedToIcon = true } }
        }
    }

    // ---- Состояния видимости иконка/кружок ----
    states: [
        State {
            name: "ringVisible"; when: root.showRing
            PropertyChanges { target: downloadIcon; opacity: 0 }
            PropertyChanges { target: ring; opacity: 1 }
        }
    ]
    transitions: [
        Transition {
            NumberAnimation { properties: "opacity"; duration: 200; easing.type: Easing.InOutQuad }
        }
    ]

    // При старте новой загрузки возвращаем масштаб и флаг
    onDownloadingChanged: {
        if (root.downloading) {
            root.returnedToIcon = false
            ring.scale = 1.0
        }
    }

    // Мышка (enabled всегда — чтобы тултип показывался в любом состоянии)
    MouseArea {
        id: hover
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: root.clickable ? Qt.PointingHandCursor : Qt.ArrowCursor
        onClicked: {
            if (root.clickable)
                backend.startRemoteDownload()
        }
    }

    // Тултип
    AppToolTip {
        anchors.horizontalCenter: parent.horizontalCenter
        anchors.top: parent.bottom
        anchors.topMargin: AppTheme.spaceXXS
        dropDown: true
        isVisible: hover.containsMouse
        text: root.tipText
    }
}
