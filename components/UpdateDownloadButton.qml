import QtQuick
import QtQuick.Controls

// ============================================================
// КНОПКА ОБНОВЛЕНИЯ В ШАПКЕ — пять состояний, одно место
//
//   покой       — серая «перезагрузка»; клик — проверить обновления
//   проверка    — тот же значок КРУТИТСЯ, пока идёт проверка
//                 (фоновая в том числе: кнопка честно показывает работу)
//   доступно    — синяя стрелка «скачать» + радар
//   загрузка    — синее кольцо наполняется прогрессом; радар скрыт
//   готово      — кольцо полное и зелёное, внутри стрелка «установить»
//
// Все переходы — Behavior: значки перекрёстно растворяются
// (прозрачность + лёгкий рост), анимацию можно прервать и развернуть
// в любую сторону — кнопка не застревает в промежуточной позе.
// Приоритет состояний: загрузка → готово → доступно → проверка → покой.
// Фоновая проверка поверх «доступно» значок НЕ меняет — не мельтешит.
//
// Крутится только значок «перезагрузка», и только в «проверке»;
// остановившись, докручивается до ровного положения.
// ============================================================

    Item {
        id: root
        width: 46
        height: parent ? parent.height : 46
        objectName: "updateButton"

        // Клавиатура: таб останавливается на кнопке, Enter и Пробел нажимают
        activeFocusOnTab: true
        Keys.onReturnPressed: if (root.clickable) root.activate()
        Keys.onEnterPressed: if (root.clickable) root.activate()
        Keys.onSpacePressed: if (root.clickable) root.activate()
        Accessible.role: Accessible.Button
        // тултип и есть честное имя: действие + состояние
        Accessible.name: root.tipText

        // одно действие для мыши и клавиатуры
        function activate() {
            if (root.ready) {
                backend.applyReadyUpdate()
            } else if (root.hasUpdate && !root.downloading) {
                backend.startRemoteDownload()
            } else if (root.idleCheck && !root.checking) {
                backend.checkAllUpdateSources()
            }
        }

    // ---- Состояния от Backend ----
    readonly property bool hasUpdate: backend.remoteUpdateAvailable
    readonly property bool downloading: backend.remoteDownloading
    readonly property int progress: backend.remoteDownloadProgress
    readonly property bool ready: backend.updateReady
    readonly property bool checking: backend.updateChecking

    readonly property bool idleCheck: !root.hasUpdate && !root.downloading && !root.ready

    // ---- Текущее состояние (единственный источник истины) ----
    readonly property string mode: root.downloading ? "downloading"
                                   : root.ready ? "ready"
                                   : root.hasUpdate ? "available"
                                   : root.checking ? "checking"
                                   : "idle"

    // клик что-то делает: установить / скачать / проверить.
    // Во время проверки и загрузки кнопка не откликается — курсор честный.
    readonly property bool clickable: root.ready
                                      || (root.hasUpdate && !root.downloading)
                                      || (root.idleCheck && !root.checking)

    // ---- Цвет значков: состояние + ховер ----
    // До загрузки стрелка синяя — в цвет радара; когда обновление
    // загружено и можно установить — зелёная. Радар в обоих состояниях
    // акцентно-синий: различает их стрелка.
    readonly property color stateColor: root.ready ? AppTheme.accentSuccess
                                   : (root.hasUpdate || root.downloading) ? AppTheme.accentBrand
                                   : (hover.containsMouse ? AppTheme.textPrimary : AppTheme.textSecondary)

    readonly property string tipText: {
        if (downloading)
            return "Загрузка обновления… " + progress + "%"
        if (ready)
            return "Обновление загружено. Нажмите, чтобы установить."
        if (hasUpdate)
            return "Доступна новая версия программы. Нажмите, чтобы загрузить."
        if (checking)
            return "Проверяем обновления…"
        return "Проверить обновления"
    }

    // ---- Радар: есть что скачать/установить ----
    // Точка в правом верхнем углу и исходящие кольца — «нажми на меня».
    // Скрыт, пока идёт загрузка (там говорит кольцо) или проверка.
    UpdateRadar {
        anchors.right: parent.right
        anchors.rightMargin: 4
        anchors.top: parent.top
        anchors.topMargin: 4
        z: 10
        active: (root.hasUpdate && !root.downloading) || root.ready
        waveColor: AppTheme.accentBrand
    }

    // ---- Подложка (hover/нажатие) ----
    Rectangle {
        anchors.fill: parent
        color: hover.pressed ? AppTheme.statePress : (hover.containsMouse ? AppTheme.stateHover : "transparent")
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    // ---- Рамка фокуса с клавиатуры (как у кнопок формы) ----
    Rectangle {
        anchors.fill: parent
        anchors.margins: -AppTheme.focusOffset
        radius: AppTheme.radiusMedium + AppTheme.focusOffset
        color: "transparent"
        border.color: AppTheme.borderFocus
        border.width: AppTheme.focusWidth
        opacity: root.activeFocus ? 1.0 : 0.0
        Behavior on opacity { NumberAnimation { duration: AppTheme.durMicro } }
    }

    // ---- Значки ----
    // Вся клетка сжимается при нажатии (как у обычных кнопок программы).
    Item {
        id: cell
        anchors.centerIn: parent
        width: 28; height: 28
        scale: hover.pressed && root.clickable ? AppTheme.scaleActive : 1.0
        Behavior on scale { NumberAnimation { duration: AppTheme.durFast; easing.type: Easing.OutQuad } }

        // «перезагрузка»: покой и проверка. Крутится, пока идёт проверка.
        IconImage {
            id: iRefresh
            objectName: "updRefresh"
            anchors.centerIn: parent
            source: "../icons/refresh.svg"
            width: AppTheme.iconMedium + 2; height: AppTheme.iconMedium + 2
            color: root.stateColor
            opacity: (root.mode === "idle" || root.mode === "checking") ? 1 : 0
            scale: (root.mode === "idle" || root.mode === "checking") ? 1 : 0.25
            Behavior on color { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
            Behavior on scale { NumberAnimation { duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter } }

            rotation: 0

            // равномерный оборот, пока идёт проверка
            SequentialAnimation {
                id: spin
                running: root.mode === "checking"
                loops: Animation.Infinite
                NumberAnimation {
                    target: iRefresh
                    property: "rotation"
                    from: iRefresh.rotation
                    to: iRefresh.rotation + 360
                    duration: 900
                    easing.type: Easing.Linear
                }
            }
            // проверка кончилась — докрутиться до ровного положения
            // и встать: выход с торможением, без щелчка
            NumberAnimation {
                id: settle
                target: iRefresh
                property: "rotation"
                duration: AppTheme.durNormal
                easing.type: Easing.OutCubic
            }
            Connections {
                target: spin
                function onRunningChanged() {
                    if (!spin.running && iRefresh.rotation > 0) {
                        settle.to = Math.ceil(iRefresh.rotation / 360) * 360
                        settle.restart()
                    }
                }
            }
        }

        // «скачать»: обновление доступно
        IconImage {
            id: iDownload
            objectName: "updDownload"
            anchors.centerIn: parent
            source: "../icons/download.svg"
            width: AppTheme.iconMedium + 2; height: AppTheme.iconMedium + 2
            color: root.stateColor
            opacity: root.mode === "available" ? 1 : 0
            scale: root.mode === "available" ? 1 : 0.25
            Behavior on color { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
            Behavior on scale { NumberAnimation { duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter } }
        }

        // кольцо прогресса: загрузка и готовность
        Item {
            id: iRing
            objectName: "updRing"
            anchors.centerIn: parent
            width: 26; height: 26
            opacity: (root.mode === "downloading" || root.mode === "ready") ? 1 : 0
            scale: (root.mode === "downloading" || root.mode === "ready") ? 1 : 0.25
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
            Behavior on scale { NumberAnimation { duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter } }

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
            opacity: root.mode === "ready" ? 1 : 0
            scale: root.mode === "ready" ? 1 : 0.25
            Behavior on color { ColorAnimation { duration: AppTheme.durFast; easing.type: Easing.InOutQuad } }
            Behavior on opacity { NumberAnimation { duration: AppTheme.durFast } }
            Behavior on scale { NumberAnimation { duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter } }
        }
    }

    // ---- Мышь ----
    HoverHandler { cursorShape: root.clickable ? Qt.PointingHandCursor : Qt.ArrowCursor }
    MouseArea {
        id: hover
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: root.clickable ? Qt.PointingHandCursor : Qt.ArrowCursor
        onClicked: root.activate()
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
