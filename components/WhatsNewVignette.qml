import QtQuick
import QtQuick.Controls
import "."

// ============================================================
// ЖИВАЯ ВИНЬЕТКА ОКНА «ЧТО НОВОГО»
//
// Сверху окна после обновления появляется маленькая живая сцена,
// показывающая главное изменение выпуска в действии. Питон отдаёт
// СПИСОК сцен — по одной на каждую «показывающую» новинку обновления
// (backend.whatsNewScenes): «wave» — волна меток по дням, выделенным
// протяжкой. Если сцен несколько, они листаются сами: одна проигралась
// — наступает следующая, по кругу. Точки под сценой показывают, какая
// сейчас, и позволяют выбрать любую. Обновлениям без «показывающих»
// новинок достаётся нейтральная сцена: маскот и лёгкое конфетти
// в цветах программы.
//
// Сцена спокойная: те же длительности и кривые, что у остального
// интерфейса (вход — торможение), без вращений и прыжков. Живёт,
// только пока окно открыто.
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
    // Сколько длится один полный прогон сцены — после него авто-листание
    // показывает следующую. Волна — 150 + 800 + 100 + 300 + (80 + 240 +
    // 250) + 1200 + 150 + 700 ≈ 4 с; нейтральная — фоновая, 6 с
    readonly property int sceneCycleMs: currentScene === "wave" ? 4000 : 6000

    visible: active && scenes.length > 0
    width: parent ? parent.width : 0
    height: 132

    onActiveChanged: {
        if (!active) {
            sceneAdvance.stop()
            return
        }
        currentIndex = 0
        waveScene.resetScene()
        vignetteEnter.restart()
        startAutoAdvance()
    }

    onScenesChanged: if (currentIndex >= scenes.length) currentIndex = 0

    function startAutoAdvance() {
        if (scenes.length > 1) sceneAdvance.restart()
        else sceneAdvance.stop()
    }

    // Выбор сцены по точке (и для проверок): перезапускает её сюжет
    function goToScene(i) {
        if (i < 0 || i >= scenes.length || i === currentIndex) return
        currentIndex = i
    }

    onCurrentIndexChanged: {
        if (currentScene === "wave") waveScene.resetScene()
        startAutoAdvance()
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
            id: stage
            anchors.fill: parent
            radius: AppTheme.radiusMedium
            color: AppTheme.bgSurface
            border.color: AppTheme.borderDivider
            border.width: 1
        }

        // ============================================================
        // СЦЕНА «ВОЛНА»: семь дней, курсор проводит по числам,
        // метки отпуска собираются по одной — слева направо
        // ============================================================
        Item {
            id: waveScene
            objectName: "vignetteWaveScene"
            anchors.fill: parent
            // мягкая смена сцен: проявление вместо щелчка
            visible: opacity > 0
            opacity: root.currentScene === "wave" ? 1 : 0
            Behavior on opacity {
                NumberAnimation { duration: AppTheme.durMicro }
            }

            readonly property int cellCount: 7
            readonly property real cellW: 48
            readonly property real cellH: 64
            readonly property real cellGap: 8
            readonly property real stripW: cellCount * cellW + (cellCount - 1) * cellGap
            readonly property real rowY: (height - cellH) / 2

            function cellCenterX(i) {
                return cellsRow.x + i * (cellW + cellGap) + cellW / 2
            }

            // Возврат к началу сюжета (окно открыли заново или вернулись
            // к сцене точкой)
            function resetScene() {
                cursor.sweeping = false
                cursor.opacity = 0
                for (let i = 0; i < cellsRepeater.count; i++) {
                    let cell = cellsRepeater.itemAt(i)
                    if (cell) cell.resetBadge()
                }
            }

            Row {
                id: cellsRow
                x: (parent.width - parent.stripW) / 2
                y: parent.rowY
                spacing: waveScene.cellGap

                Repeater {
                    id: cellsRepeater
                    model: waveScene.cellCount

                    Rectangle {
                        id: miniCell
                        objectName: "vignetteMiniCell"
                        width: waveScene.cellW
                        height: waveScene.cellH
                        radius: AppTheme.radiusMedium
                        // выходные (числа 10 и 11) — как в календаре,
                        // красноватые
                        color: (5 + index) >= 10
                               ? AppTheme.bgDangerSoft : AppTheme.bgCell

                        // Фасады для сюжета сцены
                        function resetBadge() { miniBadge.resetBadge() }
                        function startBadgeIn() { miniBadge.startIn() }
                        function startBadgeOut() { miniBadge.startOut() }

                        // Число дня — маленькое, в правом верхнем углу,
                        // как в настоящих ячейках
                        Text {
                            anchors.top: parent.top
                            anchors.right: parent.right
                            anchors.margins: 5
                            text: 5 + index
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            font.weight: AppTheme.weightBold
                            color: (5 + index) >= 10
                                   ? AppTheme.accentDanger
                                   : AppTheme.textSecondary
                        }

                        // Плёнка выделения: день подсвечен, пока курсор
                        // его прошёл и выделение не снято
                        Rectangle {
                            anchors.fill: parent
                            radius: parent.radius
                            color: AppTheme.stateHover
                            opacity: (cursor.sweeping
                                      && cursor.x + cursor.width / 2
                                         >= waveScene.cellCenterX(index)) ? 1 : 0
                            Behavior on opacity {
                                NumberAnimation { duration: AppTheme.durMicro }
                            }
                        }

                        // Метка отпуска «О» — под числом, как в ячейке
                        Rectangle {
                            id: miniBadge
                            objectName: "vignetteMiniBadge"
                            anchors.right: parent.right
                            anchors.rightMargin: 5
                            y: 26
                            width: 16
                            height: 16
                            radius: AppTheme.radiusPill
                            color: AppTheme.bgWarningSoft
                            opacity: 0
                            scale: 0.55

                            Text {
                                anchors.centerIn: parent
                                text: "О"
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeMicro
                                font.weight: AppTheme.weightBold
                                color: AppTheme.accentWarning
                            }

                            function resetBadge() {
                                stopBadgeAnims()
                                opacity = 0
                                scale = 0.55
                            }
                            function stopBadgeAnims() {
                                badgeIn.stop()
                                badgeOut.stop()
                            }
                            // старт из чужого скоупа: id-шники анимаций —
                            // не свойства, зовём только методы
                            function startIn() {
                                stopBadgeAnims()
                                badgeIn.restart()
                            }
                            function startOut() {
                                stopBadgeAnims()
                                badgeOut.restart()
                            }

                            // Сборка метки: проявление и рост (та же пара,
                            // что у меток календаря), очередь — как в волне
                            SequentialAnimation {
                                id: badgeIn
                                PauseAnimation { duration: 80 + index * 40 }
                                ParallelAnimation {
                                    NumberAnimation { target: miniBadge; property: "opacity"; to: 1; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
                                    NumberAnimation { target: miniBadge; property: "scale"; to: 1; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
                                }
                            }
                            // Растворение: выход — торможение сменяется
                            // разгоном, короче входа
                            ParallelAnimation {
                                id: badgeOut
                                NumberAnimation { target: miniBadge; property: "opacity"; to: 0; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
                                NumberAnimation { target: miniBadge; property: "scale"; to: 0.55; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
                            }
                        }
                    }
                }
            }

            // Курсор: точка, которую «ведут» по числам
            Rectangle {
                id: cursor
                objectName: "vignetteCursor"
                property bool sweeping: false
                width: 12
                height: 12
                radius: 6
                x: waveScene.cellCenterX(0) - width / 2
                y: waveScene.rowY + waveScene.cellH / 2 - height / 2
                color: AppTheme.accentBrand
                opacity: 0
            }

            // Сюжет по кругу: курсор появляется → проводит по дням →
            // отпускает кнопку → метки собираются по одной → пауза →
            // выделение снято, метки растворились → пауза → заново
            SequentialAnimation {
                id: waveLoop
                running: waveScene.visible && root.animsAlive
                loops: Animation.Infinite

                ScriptAction {
                    script: {
                        cursor.x = waveScene.cellCenterX(0) - cursor.width / 2
                        cursor.sweeping = true
                    }
                }
                NumberAnimation { target: cursor; property: "opacity"; to: 0.9; duration: AppTheme.durFast; easing.type: AppTheme.easeEnter }
                NumberAnimation { target: cursor; property: "x"; to: waveScene.cellCenterX(waveScene.cellCount - 1) - cursor.width / 2; duration: 800; easing.type: Easing.InOutSine }
                NumberAnimation { target: cursor; property: "opacity"; to: 0; duration: AppTheme.durMicro; easing.type: AppTheme.easeExit }
                PauseAnimation { duration: 300 }
                ScriptAction {
                    script: {
                        for (let i = 0; i < cellsRepeater.count; i++) {
                            let cell = cellsRepeater.itemAt(i)
                            if (cell) cell.startBadgeIn()
                        }
                    }
                }
                PauseAnimation { duration: 80 + (waveScene.cellCount - 1) * 40 + AppTheme.durStandard }
                PauseAnimation { duration: 1200 }
                ScriptAction {
                    script: {
                        cursor.sweeping = false
                        for (let i = 0; i < cellsRepeater.count; i++) {
                            let cell = cellsRepeater.itemAt(i)
                            if (cell) cell.startBadgeOut()
                        }
                    }
                }
                PauseAnimation { duration: 700 }
            }
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
