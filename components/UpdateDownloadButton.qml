import QtQuick
import QtQuick.Controls

// ============================================================
// КНОПКА ОБНОВЛЕНИЯ В ШАПКЕ (рядом со справкой)
//
// Стиль — как у иконок во всей программе: клетка 46px,
// монохромная иконка refresh.svg, ховер затемняет фон.
//
// Состояния (переходы бесшовные, на пружинах):
//   • нет обновления  — иконка как у всех (серая), клик —
//     проверка: иконка делает оборот;
//   • обновление есть — иконка зеленеет и «дышит» (масштаб),
//     вокруг искрят стрелочки;
//   • идёт загрузка   — иконка крутится, вокруг неё дуга
//       прогресса (заполняется плавно, без щелчков);
//   • загружено       — дуга зеленеет и заполняется, иконка
//     пружинисто превращается в галочку.
//
// Прогресс от бэкенда приходит скачками — дуга сглаживает
// его анимацией (Behavior 250 мс), поэтому состояние всегда
// перетекает в следующее, без резких подмен.
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

    // Клик по «праздной» кнопке (нет обновления и ничего не качается) запускает проверку.
    readonly property bool idleCheck: !root.hasUpdate && !root.downloading && !root.ready

    // Тултип зависит от состояния
    readonly property string tipText: {
        if (downloading)
            return "Загрузка обновления… " + progress + "%"
        if (ready)
            return "Обновление загружено. Нажмите, чтобы установить."
        if (hasUpdate)
            return "Доступна новая версия программы. Нажмите, чтобы загрузить."
        return "Проверить обновления"
    }

    // ---- Цвета ----
    readonly property color idleColor: AppTheme.textSecondary   // как у остальных иконок шапки
    readonly property color availColor: AppTheme.accentSuccess
    readonly property color ringColor: AppTheme.accentBrand

    // Цвет иконки: состояние + ховер (всякие переходы — цветовой анимацией)
    property color iconColor: downloading ? ringColor
                       : (hasUpdate || ready) ? availColor
                       : (hover.containsMouse ? AppTheme.textPrimary : idleColor)
    Behavior on iconColor { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }

    // ---- Искры-стрелочки вокруг кнопки (когда есть что скачать) ----
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

    // ---- Иконка (refresh) + галочка (готово): крестятся фейдом, галочка пружинит ----
    Item {
        id: iconWrap
        objectName: "updIconWrap"
        anchors.centerIn: parent
        width: refreshIcon.width
        height: refreshIcon.height

        IconImage {
            id: refreshIcon
            objectName: "updRefresh"
            anchors.centerIn: parent
            source: "../icons/refresh.svg"
            width: AppTheme.iconMedium
            height: AppTheme.iconMedium
            color: root.iconColor
            opacity: root.ready ? 0 : 1
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
        }

        IconImage {
            id: checkIcon
            objectName: "updCheck"
            anchors.centerIn: parent
            source: "../icons/check.svg"
            width: AppTheme.iconMedium
            height: AppTheme.iconMedium
            color: root.availColor
            opacity: root.ready ? 1 : 0
            scale: root.ready ? 1 : 0.4
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
            // пружинистое появление с лёгким перелётом — «инерция»
            Behavior on scale { SpringAnimation { spring: 5.0; damping: 0.35; mass: 1.1 } }
        }

        // «Дыхание», когда обновление доступно: зеленая иконка мягко пульсирует
        SequentialAnimation {
            running: root.hasUpdate && !root.downloading && !root.ready
            loops: Animation.Infinite
            NumberAnimation { target: iconWrap; property: "scale"; to: 1.08; duration: 550; easing.type: Easing.OutCubic }
            NumberAnimation { target: iconWrap; property: "scale"; to: 1.0; duration: 550; easing.type: Easing.InCubic }
        }

        // Кручение во время загрузки: иконка refresh «работает»
        RotationAnimation {
            id: downloadSpin
            target: iconWrap
            property: "rotation"
            from: 0; to: 360
            duration: 1400
            loops: Animation.Infinite
            running: root.downloading
        }

        // Оборот-отклик на клик «проверить обновления»
        RotationAnimation {
            id: checkSpin
            target: iconWrap
            property: "rotation"
            from: 0; to: 360
            duration: 700
            easing.type: Easing.OutCubic
        }
    }

    // ---- Дуга прогресса вокруг иконки ----
    Item {
        id: arcWrap
        objectName: "updArc"
        anchors.centerIn: parent
        width: 32
        height: 32
        opacity: 0
        scale: 0.55

        // Заполнение: скачки прогресса сглаживаются анимацией.
        // До загрузки — 0 (дуга растёт с нуля), в готовности — полная
        property real fill: root.downloading ? (root.progress / 100) : (root.ready ? 1.0 : 0.0)
        Behavior on fill { NumberAnimation { duration: 250; easing.type: Easing.OutCubic } }

        // Цвет дуги: синий при загрузке, зелёный когда готово
        property color arcColor: root.ready ? root.availColor : root.ringColor
        Behavior on arcColor { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }

        // Появление — пружиной с перелётом, уход — фейдом
        states: State {
            when: root.downloading || root.ready
            PropertyChanges { target: arcWrap; opacity: 1; scale: 1 }
        }
        transitions: Transition {
            SpringAnimation { properties: "scale"; spring: 5.0; damping: 0.35; mass: 1.1 }
            NumberAnimation { properties: "opacity"; duration: AppTheme.durFast }
        }

        Canvas {
            id: arcCanvas
            anchors.fill: parent
            property real p: arcWrap.fill
            property color stroke: arcWrap.arcColor
            onPaint: {
                var ctx = getContext("2d")
                ctx.reset()
                var w = width, h = height, r = (w - 4) / 2, cx = w / 2, cy = h / 2
                ctx.lineWidth = 2.5
                ctx.lineCap = "round"
                // фоновый круг
                ctx.beginPath()
                ctx.arc(cx, cy, r, 0, Math.PI * 2)
                ctx.strokeStyle = Qt.rgba(root.idleColor.r, root.idleColor.g, root.idleColor.b, 0.35)
                ctx.stroke()
                // дуга прогресса
                ctx.beginPath()
                ctx.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + Math.PI * 2 * Math.min(p, 1))
                ctx.strokeStyle = stroke
                ctx.stroke()
            }
            onPChanged: requestPaint()
            onStrokeChanged: requestPaint()
        }

        // Пульс готовности: дуга пружинисто вздрагивает и успокаивается
        SequentialAnimation {
            running: root.ready
            SpringAnimation { target: arcWrap; property: "scale"; to: 1.12; spring: 5.0; damping: 0.35; mass: 1.1 }
            SpringAnimation { target: arcWrap; property: "scale"; to: 1.0; spring: 5.0; damping: 0.35; mass: 1.1 }
        }
    }

    // Клик активен, когда обновление доступно (запуск загрузки)
    // или уже загружено (запуск установки) — но не во время загрузки
    readonly property bool clickable: root.hasUpdate && !root.downloading

    // Мышка (enabled всегда — чтобы тултип показывался в любом состоянии)
    MouseArea {
        id: hover
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: (root.clickable || root.idleCheck) ? Qt.PointingHandCursor : Qt.ArrowCursor
        onClicked: {
            if (root.clickable) {
                // Обновление уже скачано — устанавливаем; иначе — начинаем загрузку
                if (root.ready)
                    backend.applyReadyUpdate()
                else
                    backend.startRemoteDownload()
            }
            // Праздная кнопка — клик запускает полную проверку обновлений
            // (локально → адрес из настроек → вшитый GitHub) с оборотом-откликом
            else if (root.idleCheck) {
                checkSpin.restart()
                backend.checkAllUpdateSources()
            }
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
