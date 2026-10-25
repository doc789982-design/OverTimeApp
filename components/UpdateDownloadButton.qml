import QtQuick
import QtQuick.Controls

// ============================================================
// КНОПКА ОБНОВЛЕНИЯ В ШАПКЕ — единый «живой» значок (сборка 310)
//
// Вместо набора отдельных картинок значок рисуется из одних и
// тех же элементов: дуги и трёх штрихов с общим узлом. Каждое
// состояние — это просто другая геометрия этих же элементов,
// поэтому иконки не сменяют друг друга, а перетекают:
//
//   покой      — дуга-круг со стрелкой «проверить» (серый);
//   проверка   — тот же значок ровно вращается, пока программа
//                на самом деле проверяет (backend.updateChecking),
//                вращение мягко затухает в исходную позу;
//   доступно   — дуга опускается и становится «лотком», штрихи
//                собираются в стрелку «скачать» (синяя), стрелка
//                едва заметно «дышит» к лотку;
//   загрузка   — лоток поднимается вокруг стрелки кольцом
//                прогресса: дуга заполняется по процентам, а сама
//                стрелка сереет — говорит только синее кольцо;
//   готово     — кольцо дозаливается до полного и зеленеет,
//                стрелка внутри остаётся серой, а кольцо само
//                пускает от себя волну «нажми на меня».
//
// Все параметры фигуры анимируются одновременно (Behavior на
// каждом числе), поэтому переходы — одна непрерывная перестрой-
// ка, без складываний в линию и подмен картинок. Настройка
// «Меньше анимаций» отключает вращение, дыхание и сами переходы.
// Синяя точка-маячок в углу (UpdateRadar) — только у позы
// «доступно»; в «готово» её роль играет само зелёное кольцо.
// Кнопка — только мышь: клавиатурного фокуса и его обводки в
// программе нет (сборка 312).
// ============================================================
Item {
    id: root
    objectName: "updButton"
    width: 46
    height: parent ? parent.height : 46

    // ---- Состояния от Backend ----
    readonly property bool hasUpdate: backend.remoteUpdateAvailable
    readonly property bool downloading: backend.remoteDownloading
    readonly property int progress: backend.remoteDownloadProgress
    readonly property bool ready: backend.updateReady
    // Идёт ли настоящая проверка: вращение значка честно следует
    // за ней, а не крутится «для красоты» по таймеру.
    readonly property bool checking: !!backend.updateChecking

    readonly property bool idleCheck: !root.hasUpdate && !root.downloading && !root.ready
    readonly property bool reduced: !!backend.reducedMotion

    property string vis: "refresh"     // refresh | download | ring | install

    readonly property string targetVis: root.downloading ? "ring"
                                       : root.ready ? "install"
                                       : root.hasUpdate ? "download"
                                       : "refresh"

    // ---- Цвета значка: кольцо и стрелка живут отдельно ----
    // Кольцо (дуга): проверка и «есть обновление» — бренд-синее,
    // «готово» — зелёное, покой — серое как остальная шапка.
    // Стрелка (штрихи): синяя только в «доступно»; пока идёт
    // загрузка и когда всё готово, она серая, «неактивная», —
    // говорит кольцо, а не стрелка (сборка 312).
    readonly property color stateColor: root.ready ? AppTheme.accentSuccess
                                   : (root.hasUpdate || root.downloading || root.checking) ? AppTheme.accentBrand
                                   : (hover.containsMouse ? AppTheme.textPrimary : AppTheme.textSecondary)

    property color glyphColor: root.stateColor
    Behavior on glyphColor { ColorAnimation { duration: root.reduced ? 0 : AppTheme.durSlow; easing.type: Easing.InOutQuad } }

    property color segColor: (root.downloading || root.ready) ? AppTheme.textSecondary : root.stateColor
    Behavior on segColor { ColorAnimation { duration: root.reduced ? 0 : AppTheme.durSlow; easing.type: Easing.InOutQuad } }

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

    // ==========================================================
    // ГЕОМЕТРИЯ ЗНАЧКА (сетка 24×24)
    // Дуга + три штриха из общего узла (joint). Позы состояний —
    // функция applyPose(); каждое число едет к цели своим
    // Behavior, поэтому вся фигура перестраивается как одна.
    // ==========================================================
    property real arcCX: 12
    property real arcCY: 12
    property real arcR: 7.6
    property real arcStart: 35        // градусы, 0 — вправо, по часовой
    property real arcSweep: 270
    property real arcAlpha: 1
    property real trackAlpha: 0       // бледное кольцо под дугой прогресса

    property real jointX: 16.36       // общий узел трёх штрихов
    property real jointY: 5.77
    property real aX: 13.06           // штрих A: крыло «проверить» → шток стрелки
    property real aY: 5.95
    property real aAlpha: 1
    property real bX: 15.40           // штрих B: крыло → левая грань стрелки
    property real bY: 2.61
    property real bAlpha: 1
    property real cX: 16.36           // штрих C: в покое схлопнут в узел,
    property real cY: 5.77            // в стрелке — правая грань
    property real cAlpha: 0

    property real lineW: 2.1
    property real bob: 0              // фаза «дыхания» стрелки, -1..1
    property real bobAmp: 0           // амплитуда дыхания (единицы сетки)

    // Один характер движения для всей фигуры: уверенный старт,
    // мягкая посадка. Дуга прогресса — чуть отзывчивее, чтобы
    // проценты при загрузке не отставали от настоящих.
    Behavior on arcCX      { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on arcCY      { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on arcR       { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on arcStart   { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on arcSweep   { NumberAnimation { duration: root.reduced ? 0 : 280; easing.type: Easing.OutCubic } }
    Behavior on arcAlpha   { NumberAnimation { duration: root.reduced ? 0 : 400; easing.type: Easing.InOutQuad } }
    Behavior on trackAlpha { NumberAnimation { duration: root.reduced ? 0 : 400; easing.type: Easing.InOutQuad } }
    Behavior on jointX     { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on jointY     { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on aX         { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on aY         { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on aAlpha     { NumberAnimation { duration: root.reduced ? 0 : 400; easing.type: Easing.InOutQuad } }
    Behavior on bX         { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on bY         { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on bAlpha     { NumberAnimation { duration: root.reduced ? 0 : 400; easing.type: Easing.InOutQuad } }
    Behavior on cX         { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on cY         { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on cAlpha     { NumberAnimation { duration: root.reduced ? 0 : 400; easing.type: Easing.InOutQuad } }
    Behavior on lineW      { NumberAnimation { duration: root.reduced ? 0 : 540; easing.type: Easing.InOutCubic } }
    Behavior on bobAmp     { NumberAnimation { duration: root.reduced ? 0 : 400; easing.type: Easing.InOutQuad } }

    // Позы состояний. Координаты узла «проверить» — конец дуги
    // (угол -55°), крылья построены по касательной к ней.
    function applyPose(name) {
        root.vis = name
        if (name === "refresh") {
            arcCX = 12; arcCY = 12; arcR = 7.6; arcStart = 35; arcSweep = 270
            arcAlpha = 1; trackAlpha = 0
            jointX = 16.36; jointY = 5.77
            aX = 13.06; aY = 5.95; aAlpha = 1
            bX = 15.40; bY = 2.61; bAlpha = 1
            cX = 16.36; cY = 5.77; cAlpha = 0
            lineW = 2.1; bobAmp = 0
        } else if (name === "download") {
            // дуга спустилась «лотком» под стрелкой
            arcCX = 12; arcCY = 13.4; arcR = 6.8; arcStart = 25; arcSweep = 130
            arcAlpha = 1; trackAlpha = 0
            jointX = 12; jointY = 15.8
            aX = 12; aY = 5.0; aAlpha = 1
            bX = 7.6; bY = 11.4; bAlpha = 1
            cX = 16.4; cY = 11.4; cAlpha = 1
            lineW = 2.2; bobAmp = 0.55
        } else if (name === "ring") {
            // лоток поднялся кольцом прогресса вокруг стрелки
            arcCX = 12; arcCY = 12; arcR = 8.6; arcStart = -90
            arcSweep = root.progress * 3.6
            arcAlpha = 1; trackAlpha = 0.22
            jointX = 12; jointY = 15.4
            aX = 12; aY = 7.4; aAlpha = 1
            bX = 8.3; bY = 11.7; bAlpha = 1
            cX = 15.7; cY = 11.7; cAlpha = 1
            lineW = 2.35; bobAmp = 0.45
        } else if (name === "install") {
            // кольцо полное, стрелка внутри чуть крупнее
            arcCX = 12; arcCY = 12; arcR = 8.6; arcStart = -90; arcSweep = 360
            arcAlpha = 1; trackAlpha = 0
            jointX = 12; jointY = 15.6
            aX = 12; aY = 7.2; aAlpha = 1
            bX = 8.2; bY = 11.8; bAlpha = 1
            cX = 15.8; cY = 11.8; cAlpha = 1
            lineW = 2.35; bobAmp = 0.6
        }
    }

    // Проценты загрузки двигают дугу напрямую (Behavior выше
    // сглаживает шаги, чтобы кольцо не дёргалось).
    onProgressChanged: if (root.vis === "ring") arcSweep = root.progress * 3.6

    // Смена состояния -> перестройка фигуры. Антидребезг 80 мс,
    // как раньше: «загрузка кончилась» и «готово» могут прийти
    // двумя событиями подряд — промежуточная поза не нужна.
    Timer {
        id: stateDebounce
        interval: 80
        onTriggered: if (root.targetVis !== root.vis) root.applyPose(root.targetVis)
    }
    onTargetVisChanged: stateDebounce.restart()
    // Холодный старт сразу в «есть/скачано»: смены состояния не
    // было — запускаем вручную, фигура перестроится из покоя.
    Component.onCompleted: {
        // Любое изменение геометрии/цвета — перерисовать холст.
        var sigs = ["arcCXChanged", "arcCYChanged", "arcRChanged",
                    "arcStartChanged", "arcSweepChanged", "arcAlphaChanged",
                    "trackAlphaChanged", "jointXChanged", "jointYChanged",
                    "aXChanged", "aYChanged", "aAlphaChanged",
                    "bXChanged", "bYChanged", "bAlphaChanged",
                    "cXChanged", "cYChanged", "cAlphaChanged",
                    "lineWChanged", "bobChanged", "bobAmpChanged",
                    "glyphColorChanged", "segColorChanged"]
        for (var i = 0; i < sigs.length; i++)
            root[sigs[i]].connect(function() { glyph.requestPaint() })
        stateDebounce.restart()
    }

    // ---- Вращение при проверке ----
    // Ровный ход, пока идёт проверка; по окончании — плавное
    // торможение точно в исходную позу (кратно полному обороту),
    // перестройка фигуры при этом уже может идти.
    property bool _spinning: false

    NumberAnimation {
        id: spinLoop
        target: glyphItem
        property: "rotation"
        duration: 1000
        easing.type: Easing.Linear
        loops: Animation.Infinite
    }
    NumberAnimation {
        id: spinSettle
        target: glyphItem
        property: "rotation"
        duration: root.reduced ? 0 : 460
        easing.type: Easing.OutCubic
        onFinished: glyphItem.rotation = 0
    }

    function syncSpin() {
        if (root.checking && !root.reduced) {
            spinSettle.stop()
            var base = ((glyphItem.rotation % 360) + 360) % 360
            glyphItem.rotation = base
            spinLoop.from = base
            spinLoop.to = base + 360
            spinLoop.restart()
            root._spinning = true
        } else if (root._spinning) {
            root._spinning = false
            spinLoop.stop()
            var cur = glyphItem.rotation
            var rem = ((cur % 360) + 360) % 360
            spinSettle.to = rem < 8 ? cur - rem : cur + (360 - rem)
            spinSettle.restart()
        }
    }
    onCheckingChanged: syncSpin()
    onReducedChanged: syncSpin()

    // ---- Дыхание стрелки ----
    // Медленное, едва заметное: стрелка тянется к лотку/вниз
    // кольца и возвращается. Только когда есть на что нажимать
    // или идёт загрузка (bobAmp позы), и без «меньше анимаций».
    SequentialAnimation on bob {
        running: !root.reduced && root.bobAmp > 0.01
        loops: Animation.Infinite
        NumberAnimation { to: 1; duration: 1500; easing.type: Easing.InOutSine }
        NumberAnimation { to: -1; duration: 1500; easing.type: Easing.InOutSine }
    }

    // ---- Маячок в углу: только «есть обновление» ----
    // В «готово» угловой точки больше нет: волну пускает само
    // зелёное кольцо (readyHalo ниже).
    UpdateRadar {
        objectName: "updRadar"
        anchors.right: parent.right
        anchors.rightMargin: 4
        anchors.top: parent.top
        anchors.topMargin: 4
        z: 10
        active: root.hasUpdate && !root.downloading && !root.ready
        waveColor: AppTheme.accentBrand
    }

    // ---- Волна готовности: кольцо само себе радар ----
    // Окружность точно по готовому зелёному кольцу расходится
    // наружу и тает — спокойно, как волна маячка. При настройке
    // «меньше анимаций» волны нет: готовности хватает цвета.
    Rectangle {
        id: readyHalo
        objectName: "updHalo"
        anchors.centerIn: parent
        width: glyphItem.width * 0.82
        height: width
        radius: width / 2
        color: "transparent"
        border.color: AppTheme.accentSuccess
        border.width: 1.6
        opacity: 0
        SequentialAnimation on opacity {
            running: root.ready && !root.reduced; loops: Animation.Infinite
            NumberAnimation { to: 0.45; duration: 320; easing.type: Easing.OutQuad }
            NumberAnimation { to: 0.0; duration: 1680; easing.type: Easing.InQuad }
            PauseAnimation { duration: 450 }
        }
        SequentialAnimation on scale {
            running: root.ready && !root.reduced; loops: Animation.Infinite
            NumberAnimation { from: 1.0; to: 1.5; duration: 2000; easing.type: Easing.OutCubic }
            PauseAnimation { duration: 450 }
        }
    }

    // ---- Подложка (hover/нажатие) ----
    Rectangle {
        anchors.fill: parent
        color: hover.pressed ? AppTheme.statePress : (hover.containsMouse ? AppTheme.stateHover : "transparent")
        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
    }

    // ---- Сам значок ----
    Item {
        id: glyphItem
        anchors.centerIn: parent
        width: Math.round(Math.min(root.width, root.height) * 0.65)
        height: width
        scale: hover.pressed ? 0.92 : (hover.containsMouse ? 1.05 : 1)
        Behavior on scale { NumberAnimation { duration: root.reduced ? 0 : AppTheme.durFast; easing.type: AppTheme.easeEnter } }

        Canvas {
            id: glyph
            objectName: "updGlyph"
            anchors.fill: parent
            antialiasing: true
            onWidthChanged: requestPaint()
            onHeightChanged: requestPaint()
            onPaint: {
                var ctx = getContext("2d")
                ctx.reset()
                var s = width / 24
                ctx.scale(s, s)
                ctx.lineCap = "round"
                ctx.lineJoin = "round"
                ctx.lineWidth = root.lineW
                var col = root.glyphColor
                function rgba(a) { return Qt.rgba(col.r, col.g, col.b, a) }
                var scol = root.segColor
                function srgba(a) { return Qt.rgba(scol.r, scol.g, scol.b, a) }
                var yOff = root.bob * root.bobAmp
                var rad = Math.PI / 180

                // бледное кольцо-шкала под дугой прогресса
                if (root.trackAlpha > 0.004) {
                    ctx.beginPath()
                    ctx.arc(root.arcCX, root.arcCY, root.arcR, 0, Math.PI * 2)
                    ctx.strokeStyle = rgba(root.trackAlpha)
                    ctx.stroke()
                }
                // дуга
                if (root.arcAlpha > 0.004 && root.arcSweep > 0.4) {
                    ctx.beginPath()
                    if (root.arcSweep >= 359.6)
                        ctx.arc(root.arcCX, root.arcCY, root.arcR, 0, Math.PI * 2)
                    else
                        ctx.arc(root.arcCX, root.arcCY, root.arcR,
                                root.arcStart * rad,
                                (root.arcStart + root.arcSweep) * rad)
                    ctx.strokeStyle = rgba(root.arcAlpha)
                    ctx.stroke()
                }
                // три штриха из общего узла
                var segs = [[root.aX, root.aY, root.aAlpha],
                            [root.bX, root.bY, root.bAlpha],
                            [root.cX, root.cY, root.cAlpha]]
                var jointAlpha = 0
                for (var i = 0; i < segs.length; i++) {
                    if (segs[i][2] < 0.004)
                        continue
                    jointAlpha = Math.max(jointAlpha, segs[i][2])
                    ctx.beginPath()
                    ctx.moveTo(root.jointX, root.jointY + yOff)
                    ctx.lineTo(segs[i][0], segs[i][1] + yOff)
                    ctx.strokeStyle = srgba(segs[i][2])
                    ctx.stroke()
                }
                // круглая «капля» в узле — стык штрихов без шва
                if (jointAlpha > 0.004) {
                    ctx.beginPath()
                    ctx.arc(root.jointX, root.jointY + yOff, root.lineW / 2, 0, Math.PI * 2)
                    ctx.fillStyle = srgba(jointAlpha)
                    ctx.fill()
                }
            }
        }
    }

    // ---- Действие кнопки (только мышь) ----
    function doClick() {
        if (root.ready) {
            backend.applyReadyUpdate()
        } else if (root.hasUpdate && !root.downloading) {
            backend.startRemoteDownload()
        } else if (root.idleCheck) {
            backend.checkAllUpdateSources()
        }
    }

    HoverHandler { cursorShape: (root.ready || (root.hasUpdate && !root.downloading) || root.idleCheck) }
    MouseArea {
        id: hover
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: (root.ready || (root.hasUpdate && !root.downloading) || root.idleCheck)
                     ? Qt.PointingHandCursor : Qt.ArrowCursor
        onClicked: root.doClick()
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
