import QtQuick
import QtQuick.Controls

// ============================================================
// КНОПКА ОБНОВЛЕНИЯ В ШАПКЕ (рядом со справкой)
//
// ИКОНКА — «стрелка вниз на поднос» (download): смысл «скачать
// новую версию», а не «перезагрузить». В покое — как у всех
// иконок шапки: серая, темнеет при наведении.
//
// Состояния (переходы бесшовные, на пружинах):
//   • нет обновления  — просто иконка; клик запускает проверку,
//     иконка приседает и пружинисто возвращается;
//   • обновление есть — иконка зеленеет и «дышит», в углу
//     вспрыгивает зелёная точка-бейдж («есть что скачать»),
//     вокруг искрят стрелочки;
//   • идёт загрузка   — иконка мягко покачивается вниз-вверх,
//     вокруг неё заполняется дуга прогресса (скачки процентов
//     сглаживаются, дуга не щёлкает);
//   • загружено       — дуга зеленеет и заполняется, стрелка
//     пружинисто превращается в галочку.
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

    // Цвет иконки: состояние + ховер (переходы — цветовой анимацией)
    property color iconColor: downloading ? ringColor
                       : (hasUpdate || ready) ? availColor
                       : (hover.containsMouse ? AppTheme.textPrimary : idleColor)
    Behavior on iconColor { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }

    // Показывать бейдж-точку: есть обновление или уже скачано (не во время загрузки)
    readonly property bool badgeShown: (root.hasUpdate || root.ready) && !root.downloading

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

    // ---- Иконка (скачать) + галочка (готово) ----
    Item {
        id: iconWrap
        objectName: "updIconWrap"
        anchors.centerIn: parent
        width: downloadIcon.width
        height: downloadIcon.height

        // покачивание и приседания живут здесь, чтобы не мешать
        // «дыханию» масштаба на iconWrap
        Item {
            id: arrowBob
            objectName: "updArrowBob"
            // НЕ anchors.fill: он заякоривает y и мешает покачиванию
            width: parent.width
            height: parent.height

            IconImage {
                id: downloadIcon
                objectName: "updDownload"
                anchors.centerIn: parent
                source: "../icons/download.svg"
                width: AppTheme.iconMedium + 2
                height: AppTheme.iconMedium + 2
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
        }

        // «Дыхание», когда обновление доступно: иконка мягко пульсирует
        SequentialAnimation {
            running: root.hasUpdate && !root.downloading && !root.ready
            loops: Animation.Infinite
            NumberAnimation { target: iconWrap; property: "scale"; to: 1.08; duration: 550; easing.type: Easing.OutCubic }
            NumberAnimation { target: iconWrap; property: "scale"; to: 1.0; duration: 550; easing.type: Easing.InCubic }
        }

        // Покачивание во время загрузки: стрелка мягко «кланяется» вниз
        SequentialAnimation {
            running: root.downloading
            loops: Animation.Infinite
            NumberAnimation { target: arrowBob; property: "y"; to: 1.6; duration: 380; easing.type: Easing.OutCubic }
            NumberAnimation { target: arrowBob; property: "y"; to: 0; duration: 380; easing.type: Easing.InCubic }
        }

        // Приседание-отклик на клик «проверить обновления»:
        // стрелка dip вниз и пружинисто возвращается
        SequentialAnimation {
            id: checkDip
            NumberAnimation { target: arrowBob; property: "y"; to: 2.5; duration: 110; easing.type: Easing.OutCubic }
            SpringAnimation { target: arrowBob; property: "y"; to: 0; spring: 5.0; damping: 0.35; mass: 1.1 }
        }
    }

    // ---- Бейдж-точка: «есть новая версия» ----
    Rectangle {
        id: badge
        objectName: "updBadge"
        width: 9
        height: 9
        radius: width / 2
        color: root.availColor
        border.color: AppTheme.bgBase     // «вырез» из фона шапки
        border.width: 2
        anchors.right: iconWrap.right
        anchors.top: iconWrap.top
        anchors.rightMargin: -5
        anchors.topMargin: -5
        opacity: root.badgeShown ? 1 : 0
        scale: root.badgeShown ? 1 : 0.2
        Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
        // пружинистое появление с лёгким перелётом
        Behavior on scale { SpringAnimation { spring: 5.0; damping: 0.35; mass: 1.1 } }

        // тихий пульс, пока обновление ждёт клика
        SequentialAnimation {
            running: root.badgeShown && !root.ready
            loops: Animation.Infinite
            NumberAnimation { target: badge; property: "scale"; to: 1.18; duration: 480; easing.type: Easing.OutCubic }
            NumberAnimation { target: badge; property: "scale"; to: 1.0; duration: 480; easing.type: Easing.InCubic }
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
            // (локально → адрес из настроек → вшитый GitHub) с приседанием-откликом
            else if (root.idleCheck) {
                checkDip.restart()
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
