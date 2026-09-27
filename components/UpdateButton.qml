import QtQuick

// ============================================================
// КНОПКА ОБНОВЛЕНИЯ — Canvas-морфинг (код пользователя)
//
// Пять состояний, иконка каждого нарисована кодом и перетекает
// в следующее без подмен:
//   покой (круговая стрелка) → проверка (крутится) →
//   есть обновление (стрелка вниз) → загрузка (кольцо с %) →
//   готово (галочка в круге)
//
// Врезана в живой механизм обновления программы: кнопка ничего
// не выдумывает — показывает то, что реально происходит.
//   покой → клик → backend.checkAllUpdateSources()
//   есть  → клик → backend.startRemoteDownload()
//   готово→ клик → backend.applyReadyUpdate() (перезапуск)
//
// Адаптация кода пользователя к программе:
//  • «property real rotation» затирал встроенное вращение Item —
//    счётчик оборотов переименован в spinAngle (та же анимация);
//  • рисунок рассчитан на холст 32×32 — добавлен масштаб fit,
//    чтобы кнопка выглядела одинаково и 64×64, и в шапке 46×36;
//  • цвета из темы программы (тёмная/светлая);
//  • симуляция (Math.random, таймеры) убрана — только реальные
//    состояния; плюс тултип-карточка, как принято в программе.
// ============================================================

Item {
    id: root
    width: 64
    height: 64
    objectName: "updateButton"

    enum State {
        CheckForUpdate,
        Checking,
        UpdateAvailable,
        Downloading,
        ReadyToInstall
    }

    // ---- Живые данные программы ----
    readonly property bool hasUpdate: backend.remoteUpdateAvailable
    readonly property bool downloading: backend.remoteDownloading
    readonly property int progress: backend.remoteDownloadProgress
    readonly property bool ready: backend.updateReady
    readonly property bool busy: backend.updateBusy

    // Что показывать по данным программы
    readonly property int targetState:
        downloading ? UpdateButton.State.Downloading
        : ready ? UpdateButton.State.ReadyToInstall
        : busy ? UpdateButton.State.Checking
        : hasUpdate ? UpdateButton.State.UpdateAvailable
        : UpdateButton.State.CheckForUpdate

    // Что показываем сейчас (этим ведёт хореография морфинга)
    property int currentState: UpdateButton.State.CheckForUpdate
    property real downloadProgress: root.progress / 100.0

    onDownloadProgressChanged:
        if (currentState === UpdateButton.State.Downloading)
            morphCanvas.requestPaint()

    // ---- Цвета: тема программы (тёмная/светлая) ----
    readonly property color primaryColor: AppTheme.accentBrand
    readonly property color accentColor: AppTheme.accentSuccess
    readonly property color downloadColor: AppTheme.accentWarning

    readonly property string tipText: {
        if (downloading)
            return "Загрузка обновления… " + progress + "%"
        if (ready)
            return "Обновление загружено. Нажмите, чтобы установить."
        if (hasUpdate)
            return "Доступна новая версия программы. Нажмите, чтобы загрузить."
        if (busy)
            return "Проверяем обновления…"
        return "Проверить обновления"
    }

    // ---- Подложка (наведение/нажатие) ----
    Rectangle {
        id: background
        anchors.fill: parent
        radius: width / 2
        color: mouseArea.pressed ? Qt.darker(getCurrentColor(), 1.1) : getCurrentColor()
        opacity: mouseArea.containsMouse ? 0.12 : 0.08

        Behavior on color { ColorAnimation { duration: 300 } }
        Behavior on opacity { NumberAnimation { duration: 200 } }
    }

    // ---- Главный Canvas: морфинг-иконка ----
    Canvas {
        id: morphCanvas
        objectName: "updCanvas"
        anchors.centerIn: parent
        width: Math.min(root.width, root.height) - 8
        height: width

        property real morphProgress: 0.0
        property int fromState: UpdateButton.State.CheckForUpdate
        property int toState: UpdateButton.State.CheckForUpdate

        // рисунок рассчитан на холст 32×32 — масштабируем
        readonly property real fit: width / 32

        // счётчик оборотов для проверки обновлений
        property real spinAngle: 0

        NumberAnimation on spinAngle {
            from: 0
            to: 360
            duration: 1000
            loops: Animation.Infinite
            running: currentState === UpdateButton.State.Checking
        }
        onSpinAngleChanged:
            if (currentState === UpdateButton.State.Checking)
                requestPaint()

        onMorphProgressChanged: requestPaint()

        onPaint: {
            var ctx = getContext("2d");
            ctx.reset();
            ctx.strokeStyle = root.getCurrentColor();
            ctx.fillStyle = root.getCurrentColor();
            ctx.lineWidth = 2.5;
            ctx.lineCap = "round";
            ctx.lineJoin = "round";

            if (currentState === UpdateButton.State.Checking) {
                drawCheckIcon(ctx, spinAngle);
            } else if (morphAnimation.running && morphProgress < 1) {
                drawMorphing(ctx, fromState, toState, morphProgress);
            } else {
                switch (currentState) {
                    case UpdateButton.State.CheckForUpdate:
                        drawCheckIcon(ctx, 0);
                        break;
                    case UpdateButton.State.UpdateAvailable:
                        drawDownloadIcon(ctx);
                        break;
                    case UpdateButton.State.Downloading:
                        drawProgressCircle(ctx, root.downloadProgress);
                        break;
                    case UpdateButton.State.ReadyToInstall:
                        drawInstallIcon(ctx);
                        break;
                }
            }
        }

        // Иконка проверки обновлений (круговая стрелка)
        function drawCheckIcon(ctx, rotationAngle) {
            ctx.save();
            ctx.translate(width/2, height/2);
            ctx.scale(fit, fit);
            ctx.rotate(rotationAngle * Math.PI / 180);

            ctx.beginPath();
            ctx.arc(0, 0, 12, 0.3 * Math.PI, 2.2 * Math.PI);
            ctx.stroke();

            ctx.beginPath();
            ctx.moveTo(12 * Math.cos(0.3 * Math.PI), 12 * Math.sin(0.3 * Math.PI));
            ctx.lineTo(12 * Math.cos(0.3 * Math.PI) - 5, 12 * Math.sin(0.3 * Math.PI) - 3);
            ctx.lineTo(12 * Math.cos(0.3 * Math.PI) - 2, 12 * Math.sin(0.3 * Math.PI) + 5);
            ctx.closePath();
            ctx.fill();

            ctx.restore();
        }

        // Иконка загрузки (стрелка вниз)
        function drawDownloadIcon(ctx) {
            ctx.save();
            ctx.translate(width/2, height/2);
            ctx.scale(fit, fit);

            ctx.beginPath();
            ctx.moveTo(0, -10);
            ctx.lineTo(0, 8);
            ctx.stroke();

            ctx.beginPath();
            ctx.moveTo(-6, 2);
            ctx.lineTo(0, 10);
            ctx.lineTo(6, 2);
            ctx.stroke();

            ctx.beginPath();
            ctx.moveTo(-8, 12);
            ctx.lineTo(8, 12);
            ctx.stroke();

            ctx.restore();
        }

        // Круг прогресса загрузки
        function drawProgressCircle(ctx, progress) {
            ctx.save();
            ctx.translate(width/2, height/2);
            ctx.scale(fit, fit);

            var c = root.getCurrentColor();
            ctx.strokeStyle = Qt.rgba(c.r, c.g, c.b, 0.2);
            ctx.lineWidth = 3;
            ctx.beginPath();
            ctx.arc(0, 0, 12, 0, 2 * Math.PI);
            ctx.stroke();

            ctx.strokeStyle = c;
            ctx.lineWidth = 3;
            ctx.beginPath();
            ctx.arc(0, 0, 12, -0.5 * Math.PI, (-0.5 + 2 * progress) * Math.PI);
            ctx.stroke();

            ctx.fillStyle = c;
            ctx.font = "bold 8px sans-serif";
            ctx.textAlign = "center";
            ctx.textBaseline = "middle";
            ctx.fillText(Math.round(progress * 100) + "%", 0, 0);

            ctx.restore();
        }

        // Иконка установки (галочка в круге)
        function drawInstallIcon(ctx) {
            ctx.save();
            ctx.translate(width/2, height/2);
            ctx.scale(fit, fit);

            ctx.beginPath();
            ctx.arc(0, 0, 12, 0, 2 * Math.PI);
            ctx.stroke();

            ctx.beginPath();
            ctx.moveTo(-6, 0);
            ctx.lineTo(-2, 5);
            ctx.lineTo(7, -6);
            ctx.stroke();

            ctx.restore();
        }

        // Морфинг между состояниями
        function drawMorphing(ctx, from, to, progress) {
            ctx.save();
            ctx.translate(width/2, height/2);
            ctx.scale(fit, fit);

            if (from === UpdateButton.State.CheckForUpdate &&
                to === UpdateButton.State.UpdateAvailable) {
                morphCheckToDownload(ctx, progress);
            }
            else if (from === UpdateButton.State.UpdateAvailable &&
                     to === UpdateButton.State.Downloading) {
                morphDownloadToProgress(ctx, progress);
            }
            else if (from === UpdateButton.State.Downloading &&
                     to === UpdateButton.State.ReadyToInstall) {
                morphProgressToInstall(ctx, progress);
            }

            ctx.restore();
        }

        function morphCheckToDownload(ctx, t) {
            var eased = easeInOutCubic(t);

            var arcStart = 0.3 * Math.PI * (1 - eased);
            var arcEnd = 2.2 * Math.PI * (1 - eased);

            if (eased < 0.5) {
                ctx.beginPath();
                ctx.arc(0, 0, 12, arcStart, arcEnd);
                ctx.stroke();

                ctx.globalAlpha = 1 - eased * 2;
                ctx.beginPath();
                ctx.moveTo(12 * Math.cos(0.3 * Math.PI), 12 * Math.sin(0.3 * Math.PI));
                ctx.lineTo(12 * Math.cos(0.3 * Math.PI) - 5, 12 * Math.sin(0.3 * Math.PI) - 3);
                ctx.lineTo(12 * Math.cos(0.3 * Math.PI) - 2, 12 * Math.sin(0.3 * Math.PI) + 5);
                ctx.closePath();
                ctx.fill();
                ctx.globalAlpha = 1;
            } else {
                var t2 = (eased - 0.5) * 2;

                ctx.beginPath();
                ctx.moveTo(0, -10 * t2);
                ctx.lineTo(0, 8 * t2);
                ctx.stroke();

                ctx.globalAlpha = t2;
                ctx.beginPath();
                ctx.moveTo(-6, 2);
                ctx.lineTo(0, 10);
                ctx.lineTo(6, 2);
                ctx.stroke();

                ctx.beginPath();
                ctx.moveTo(-8, 12);
                ctx.lineTo(8, 12);
                ctx.stroke();
                ctx.globalAlpha = 1;
            }
        }

        function morphDownloadToProgress(ctx, t) {
            var eased = easeInOutCubic(t);

            if (eased < 0.5) {
                var t1 = 1 - (eased * 2);
                ctx.globalAlpha = t1;

                ctx.beginPath();
                ctx.moveTo(0, -10 * t1);
                ctx.lineTo(0, 8 * t1);
                ctx.stroke();

                ctx.beginPath();
                ctx.moveTo(-6 * t1, 2 * t1);
                ctx.lineTo(0, 10 * t1);
                ctx.lineTo(6 * t1, 2 * t1);
                ctx.stroke();

                ctx.beginPath();
                ctx.moveTo(-8 * t1, 12);
                ctx.lineTo(8 * t1, 12);
                ctx.stroke();

                ctx.globalAlpha = 1;
            } else {
                var t2 = (eased - 0.5) * 2;
                var c = root.getCurrentColor();

                ctx.strokeStyle = Qt.rgba(c.r, c.g, c.b, 0.2 * t2);
                ctx.lineWidth = 3;
                ctx.beginPath();
                ctx.arc(0, 0, 12 * t2, 0, 2 * Math.PI);
                ctx.stroke();
            }
        }

        function morphProgressToInstall(ctx, t) {
            var eased = easeInOutCubic(t);

            ctx.lineWidth = 3;
            ctx.beginPath();
            ctx.arc(0, 0, 12, 0, 2 * Math.PI);
            ctx.stroke();

            if (eased < 0.5) {
                var t1 = 1 - (eased * 2);
                ctx.fillStyle = root.getCurrentColor();
                ctx.font = "bold 8px sans-serif";
                ctx.textAlign = "center";
                ctx.textBaseline = "middle";
                ctx.globalAlpha = t1;
                ctx.fillText("100%", 0, 0);
                ctx.globalAlpha = 1;
            } else {
                var t2 = (eased - 0.5) * 2;

                ctx.globalAlpha = t2;
                ctx.lineWidth = 2.5;

                var checkProgress = t2;

                ctx.beginPath();
                ctx.moveTo(-6, 0);
                if (checkProgress < 0.5) {
                    var x = -6 + 4 * (checkProgress * 2);
                    var y = 0 + 5 * (checkProgress * 2);
                    ctx.lineTo(x, y);
                } else {
                    ctx.lineTo(-2, 5);
                    var x2 = -2 + 9 * ((checkProgress - 0.5) * 2);
                    var y2 = 5 - 11 * ((checkProgress - 0.5) * 2);
                    ctx.lineTo(x2, y2);
                }
                ctx.stroke();
                ctx.globalAlpha = 1;
            }
        }

        function easeInOutCubic(t) {
            return t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2;
        }
    }

    // ---- Анимация морфинга ----
    NumberAnimation {
        id: morphAnimation
        target: morphCanvas
        property: "morphProgress"
        from: 0
        to: 1
        duration: 600
        easing.type: Easing.InOutCubic

        onFinished: {
            morphCanvas.morphProgress = 0;
            morphCanvas.requestPaint();
        }

        onRunningChanged: {
            if (running) {
                morphCanvas.requestPaint();
            }
        }
    }

    // Смена состояния программы → хореография. Антидребезг: программа
    // может сообщить «загрузка кончилась» и «готово» двумя событиями
    // подряд — без паузы проигрался бы лишний промежуточный кадр.
    Timer {
        id: stateDebounce
        interval: 80
        onTriggered: root.applyTarget()
    }
    onTargetStateChanged: stateDebounce.restart()

    Component.onCompleted: {
        // на старте — сразу нужное состояние, без морфинга
        currentState = targetState;
    }

    function applyTarget() {
        var to = targetState;
        var from = currentState;
        if (to === from)
            return;
        // крутящаяся проверка — та же стрелка, что в покое
        var fromPose = from === UpdateButton.State.Checking
                       ? UpdateButton.State.CheckForUpdate : from;
        currentState = to;
        morphCanvas.requestPaint();
        // морфинг — только по цепочке; обратные и рваные переходы
        // меняются сразу, без трюков
        var chain =
            (fromPose === UpdateButton.State.CheckForUpdate && to === UpdateButton.State.UpdateAvailable) ||
            (fromPose === UpdateButton.State.UpdateAvailable && to === UpdateButton.State.Downloading) ||
            (fromPose === UpdateButton.State.Downloading && to === UpdateButton.State.ReadyToInstall);
        if (chain) {
            morphCanvas.fromState = fromPose;
            morphCanvas.toState = to;
            morphAnimation.restart();
        } else {
            morphAnimation.stop();
            morphCanvas.morphProgress = 0;
        }
    }

    // ---- Мышь ----
    MouseArea {
        id: mouseArea
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: (currentState === UpdateButton.State.CheckForUpdate ||
                      currentState === UpdateButton.State.UpdateAvailable ||
                      currentState === UpdateButton.State.ReadyToInstall)
                     ? Qt.PointingHandCursor : Qt.ArrowCursor

        onClicked: root.handleClick();
    }

    // Ripple-эффект
    Rectangle {
        id: ripple
        width: 0
        height: width
        radius: width / 2
        color: root.getCurrentColor()
        opacity: 0
        anchors.centerIn: parent

        ParallelAnimation {
            id: rippleAnimation
            NumberAnimation {
                target: ripple
                property: "width"
                from: 0
                to: root.width * 1.5
                duration: 400
                easing.type: Easing.OutCubic
            }
            NumberAnimation {
                target: ripple
                property: "opacity"
                from: 0.3
                to: 0
                duration: 400
                easing.type: Easing.OutCubic
            }
        }
    }

    // ---- Тултип-карточка ----
    AppToolTip {
        anchors.horizontalCenter: parent.horizontalCenter
        anchors.top: parent.bottom
        anchors.topMargin: AppTheme.spaceXXS
        dropDown: true
        isVisible: mouseArea.containsMouse
        text: root.tipText
    }

    function getCurrentColor() {
        switch (currentState) {
            case UpdateButton.State.CheckForUpdate:
            case UpdateButton.State.Checking:
                return primaryColor;
            case UpdateButton.State.UpdateAvailable:
            case UpdateButton.State.Downloading:
                return downloadColor;
            case UpdateButton.State.ReadyToInstall:
                return accentColor;
            default:
                return primaryColor;
        }
    }

    function handleClick() {
        rippleAnimation.start();

        switch (currentState) {
            case UpdateButton.State.CheckForUpdate:
                backend.checkAllUpdateSources();
                break;
            case UpdateButton.State.UpdateAvailable:
                backend.startRemoteDownload();
                break;
            case UpdateButton.State.ReadyToInstall:
                backend.applyReadyUpdate();
                break;
        }
    }
}
