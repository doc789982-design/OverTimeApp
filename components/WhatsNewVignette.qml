import QtQuick
import QtQuick.Controls
import "."

// ============================================================
// ЖИВАЯ ВИНЬЕТКА ОКНА «ЧТО НОВОГО»
//
// Сверху окна после обновления появляется маленькая живая сцена,
// показывающая «показывающие» новинки выпуска. Питон отдаёт СПИСОК
// сцен — по одной на каждую новинку (backend.whatsNewScenes).
// Если сцен несколько, они листаются сами: одна проигралась —
// наступает следующая, по кругу. Точки под сценой показывают, какая
// сейчас, и позволяют выбрать любую. Обновлениям без «показывающих»
// новинок достаётся нейтральная сцена: маскот и лёгкое конфетти.
//
// Сам спектакль живёт в WhatsNewStage (актёры + плеер тактов).
// Сцена здесь — только сценарий: список «что за чем происходит».
// ============================================================
Item {
    id: root
    objectName: "whatsNewVignette"

    // Сцены этого обновления (имена, в порядке сборок)
    property var scenes: []
    property int currentIndex: 0
    // Окно открыто и рассказывать есть о чём — только тогда сцена живёт
    property bool active: false

    readonly property string currentScene: scenes.length > 0
                                            ? scenes[currentIndex] : ""
    // Полный прогон сцены + короткий вдох перед следующей
    readonly property int sceneCycleMs: currentScene === "wave"
                                         ? Math.max(3000, stage.totalMs + 400)
                                         : 6000

    visible: active && scenes.length > 0
    width: parent ? parent.width : 0
    height: 132

    onActiveChanged: {
        if (!active) {
            stage.stopPlay()
            sceneAdvance.stop()
            return
        }
        currentIndex = 0
        applyScene()
        vignetteEnter.restart()
    }

    onScenesChanged: if (currentIndex >= scenes.length) currentIndex = 0

    function startAutoAdvance() {
        if (scenes.length > 1 && active) sceneAdvance.restart()
        else sceneAdvance.stop()
    }

    // Выбор сцены по точке (и для проверок)
    function goToScene(i) {
        if (i < 0 || i >= scenes.length || i === currentIndex) return
        currentIndex = i
    }

    onCurrentIndexChanged: applyScene()

    // Запуск сцены: у волны — сценарий тактов, остальные фоновые
    function applyScene() {
        if (currentScene === "wave" && active) {
            stage.steps = waveScript()
            stage.playFromStart()
        } else {
            stage.stopPlay()
        }
        startAutoAdvance()
    }

    // ── СЦЕНАРИЙ «ВОЛНА»: вся история — выделение, выбор статуса
    // в меню, расстановка меток. Одна строка = один такт плеера ──
    function waveScript() {
        return [
            { ms: 200,  run: function() { stage.cursorShow() } },
            { ms: 800,  run: function() { stage.cursorSweepTo(stage.dayCount - 1) } },
            { ms: 250,  run: function() { stage.menuOpen(stage.dayCount - 1) } },
            { ms: 300,  run: function() { stage.cursorToMenu(2) } },
            { ms: 150,  run: function() { stage.menuHighlight(2) } },
            { ms: 260,  run: function() { stage.menuPress(2); stage.cursorPress() } },
            { ms: 250,  run: function() { stage.menuClose(); stage.cursorHide() } },
            { ms: 600,  run: function() { stage.badgesWave() } },
            { ms: 1200, run: function() { } },
            { ms: 300,  run: function() { stage.badgesDissolve(); stage.clearHighlights() } },
            { ms: 600,  run: function() { } }
        ]
    }

    // Авто-листание: сцена проигралась один раз — следующая, по кругу
    Timer {
        id: sceneAdvance
        interval: root.sceneCycleMs
        onTriggered: {
            if (root.scenes.length > 1)
                root.currentIndex = (root.currentIndex + 1) % root.scenes.length
        }
    }

    readonly property bool animsAlive: active && visible
                                        && root.Window.window !== null
                                        && root.Window.window.visible

    // Для проверок: маскот нейтральной сцены. Ищем только среди прямых
    // детей — спуск во внутренности маскота (эффекты) из Python опасен
    function neutralMascot() {
        for (let i = 0; i < neutralScene.children.length; i++) {
            let c = neutralScene.children[i]
            if (c.objectName === "vignetteMascot")
                return c
        }
        return null
    }

    // Появление самой виньетки: проявление и подъём — та же пара,
    // что у собирающихся карточек календаря
    Item {
        id: body
        anchors.fill: parent
        opacity: 0
        transform: Translate { id: bodyLift; y: 8 }

        ParallelAnimation {
            id: vignetteEnter
            NumberAnimation { target: body; property: "opacity"; to: 1; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
            NumberAnimation { target: bodyLift; property: "y"; to: 0; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
        }

        // Сцена стоит на «экране» — спокойной карточке
        Rectangle {
            id: screenBg
            anchors.fill: parent
            radius: AppTheme.radiusMedium
            color: AppTheme.bgSurface
            border.color: AppTheme.borderDivider
            border.width: 1
        }

        // СЦЕНА-ТЕАТР: актёры + плеер тактов (сценарии — выше)
        WhatsNewStage {
            id: stage
            anchors.fill: parent
            visible: root.currentScene === "wave"
        }

        // ============================================================
        // НЕЙТРАЛЬНАЯ СЦЕНА: маскот дышит, вокруг — редкое конфетти
        // в цветах программы (медленный спуск с лёгким покачиванием,
        // без вращения)
        // ============================================================
        Item {
            id: neutralScene
            objectName: "vignetteNeutralScene"
            anchors.fill: parent
            visible: opacity > 0
            opacity: root.currentScene === "neutral" ? 1 : 0
            Behavior on opacity {
                NumberAnimation { duration: AppTheme.durMicro }
            }

            AppEmptyMascot {
                id: mascot
                objectName: "vignetteMascot"
                anchors.centerIn: parent
                scale: 0.5
            }

            Repeater {
                model: 8

                Rectangle {
                    id: piece
                    objectName: "vignetteConfetti"
                    property real baseX: root.width * (0.06 + 0.125 * index)
                    property real sway: (index % 2 === 0) ? -10 : 10
                    property int fallDur: 3400 + index * 280
                    width: 6
                    height: 9
                    radius: 2
                    x: baseX
                    y: -14
                    rotation: (index % 2 === 0) ? -14 : 14
                    color: [AppTheme.accentBrand, AppTheme.accentSuccess,
                            AppTheme.accentWarning, AppTheme.accentTeal,
                            AppTheme.accentPurple][index % 5]
                    opacity: 0

                    SequentialAnimation {
                        running: neutralScene.visible && root.animsAlive
                        loops: Animation.Infinite
                        PauseAnimation { duration: 500 + index * 420 }
                        ParallelAnimation {
                            NumberAnimation { target: piece; property: "y"; from: -14; to: root.height + 14; duration: piece.fallDur; easing.type: Easing.InOutSine }
                            NumberAnimation { target: piece; property: "x"; from: piece.baseX; to: piece.baseX + piece.sway; duration: piece.fallDur; easing.type: Easing.InOutSine }
                            SequentialAnimation {
                                NumberAnimation { target: piece; property: "opacity"; to: 0.7; duration: piece.fallDur * 0.3; easing.type: Easing.OutQuad }
                                NumberAnimation { target: piece; property: "opacity"; to: 0; duration: piece.fallDur * 0.7; easing.type: Easing.InQuad }
                            }
                        }
                    }
                }
            }
        }

        // ============================================================
        // ТОЧКИ: какая сцена сейчас и выбор любой
        // ============================================================
        Row {
            id: dotsRow
            objectName: "vignetteDots"
            visible: root.scenes.length > 1
            anchors.horizontalCenter: parent.horizontalCenter
            anchors.bottom: parent.bottom
            anchors.bottomMargin: 8
            spacing: 6

            Repeater {
                model: root.scenes.length

                Rectangle {
                    id: dot
                    objectName: "vignetteDot"
                    readonly property bool isCurrent: index === root.currentIndex
                    width: isCurrent ? 16 : 6
                    height: 6
                    radius: 3
                    color: isCurrent ? AppTheme.accentBrand
                                     : AppTheme.stateHover
                    Behavior on width {
                        NumberAnimation { duration: AppTheme.durMicro; easing.type: AppTheme.easeEnter }
                    }
                    HoverHandler { cursorShape: Qt.PointingHandCursor }
                    MouseArea {
                        anchors.fill: parent
                        anchors.margins: -8     // щедрая зона нажатия
                        cursorShape: Qt.PointingHandCursor
                        onClicked: root.goToScene(index)
                    }
                }
            }
        }
    }
}
