import QtQuick
import QtQuick.Controls

// ============================================================
// КНОПКА ОБНОВЛЕНИЯ В ШАПКЕ — хореография состояний
//
//   покой      — иконка «перезагрузка» (дуга со стрелкой),
//                серая, как остальные иконки шапки
//   клик       — дуга ПРОКРУЧИВАЕТСЯ на полный оборот
//                (проверка обновлений)
//   нет обновл. — прокрутилась и остановилась, ничего не меняется
//   есть обновл. — прокручивается, ВЫПРЯМЛЯЕТСЯ и превращается
//                в стрелку «скачать» (зелёная)
//   загрузка   — снова прокручивается, стрелка СВОРАЧИВАЕТСЯ
//                в кольцо, которое заполняется прогрессом
//   загружено  — кольцо РАЗВОРАЧИВАЕТСЯ в стрелку «установить»
//                (зелёная), клик устанавливает
//
// Все переходы — морфинг одной геометрии (UpdateIcon) на пружинах,
// без подмен иконок: состояние течёт из состояния.
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

    // Праздная кнопка (нет обновления, не качается) — клик = проверка
    readonly property bool idleCheck: !root.hasUpdate && !root.downloading && !root.ready

    // ---- Целевая форма по состоянию ----
    readonly property string shape: root.downloading ? "ring"
                                  : root.ready ? "install"
                                  : root.hasUpdate ? "download"
                                  : "refresh"

    property string prevShape: "refresh"

    // ---- Цвет формы: состояние + ховер ----
    readonly property color stateColor: root.downloading ? AppTheme.accentBrand
                                   : (root.hasUpdate || root.ready) ? AppTheme.accentSuccess
                                   : (hover.containsMouse ? AppTheme.textPrimary : AppTheme.textSecondary)

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

    // ---- Иконка (контейнер вращается при проверке/старте загрузки) ----
    Item {
        id: spinner
        anchors.centerIn: parent
        width: 28
        height: 28
        rotation: 0

        UpdateIcon {
            id: icon
            objectName: "updIcon"
            anchors.fill: parent
            modeA: "refresh"
            modeB: "refresh"
            morph: 0
            progress: root.progress / 100
            iconColor: root.stateColor

            Behavior on iconColor { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }
        }

        // Пульс готовности: кольцо развернулось в стрелку — вздрагивает
        SequentialAnimation {
            id: readyPulse
            SpringAnimation { target: spinner; property: "scale"; to: 1.16; spring: 5.0; damping: 0.35; mass: 1.1 }
            SpringAnimation { target: spinner; property: "scale"; to: 1.0; spring: 5.0; damping: 0.35; mass: 1.1 }
        }
    }

    // ---- Движки переходов ----
    // прокрут на полный оборот с естественным торможением
    NumberAnimation {
        id: spinAnim
        target: spinner
        property: "rotation"
        duration: 750
        easing.type: Easing.InOutCubic
    }
    // морф формы (пружина — с лёгким перелётом, «инерция»)
    SpringAnimation {
        id: morphSpring
        target: icon
        property: "morph"
        to: 1
        spring: 5.0
        damping: 0.4
        mass: 1.0
    }
    // морф после прокрута: сначала оборот, затем выпрямление
    SequentialAnimation {
        id: delayedMorph
        PauseAnimation { duration: 170 }
        ScriptAction {
            script: {
                icon.modeA = root._morphFrom
                icon.modeB = root.shape
                icon.morph = 0
                morphSpring.restart()
            }
        }
    }
    property string _morphFrom: "refresh"

    function beginMorph(fromShape) {
        root._morphFrom = fromShape
        icon.modeA = fromShape
        icon.modeB = root.shape
        icon.morph = 0
        morphSpring.restart()
    }

    onShapeChanged: {
        if (shape === prevShape) return
        var from = prevShape
        prevShape = shape
        // «прокручивается, выпрямляется и превращается»:
        // нашлось обновление или стартовала загрузка — с оборотом
        if ((from === "refresh" && shape === "download") ||
            (from === "download" && shape === "ring")) {
            spinAnim.to = spinner.rotation + 360
            spinAnim.restart()
            root._morphFrom = from
            delayedMorph.restart()
        } else {
            beginMorph(from)
            if (shape === "install") readyPulse.restart()
        }
    }

    // ---- Мышь ----
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
                // прокрут-отклик; состояние сменит бэкенд, когда узнает результат
                spinAnim.to = spinner.rotation + 360
                spinAnim.restart()
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
