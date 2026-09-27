import QtQuick

// ============================================================
// КНОПКА ОБНОВЛЕНИЯ — морфинг-кнопка (код пользователя, волнa 2)
//
// Восемь состояний, пять устойчивых:
//   Idle (дуга) → Checking (крутится) → DownloadReady (стрелка
//   вниз) → Downloading (заливающееся кольцо) → ReadyToInstall
//   (зелёный круг с галочкой)
// и три состояния-перехода, которые рисуют сам морфинг:
//   MorphToDownload / MorphToProgress / MorphToInstall.
// Переходами владеет анимация: морфинг заканчивается —
// устанавливается следующее устойчивое состояние.
//
// Врезано в живой механизм обновления (симуляторы из примера
// убраны): клики зовут backend.checkAllUpdateSources() /
// startRemoteDownload() / applyReadyUpdate(), а рисунок следует
// за реальными свойствами программы. Адаптация к программе:
//  • цвета — из темы (тёмная/светлая), не захардкожены;
//  • холст масштабируется (сетка 44×44) — и 80×80, и шапка 46×36;
//  • forceSpin: клик всегда даёт видимый прокрут, даже если
//    проверка мгновенная;
//  • на время распаковки скачанного (busy) кольцо держится
//    заполненным — галочка вырастает из полного круга;
//  • прогресс сглажен (Behavior), проценты не щёлкают;
//  • тултип-карточка, как принято в программе.
// ============================================================

Item {
    id: root
    width: 80
    height: 80
    objectName: "updateMorphButton"

    enum State {
        Idle,           // Исходное (Проверить)
        Checking,       // Крутится (Идет проверка)
        MorphToDownload,// Морфинг -> Скачать
        DownloadReady,  // Иконка Скачать
        MorphToProgress,// Морфинг -> Прогресс
        Downloading,    // Процесс загрузки (0-100%)
        MorphToInstall, // Морфинг -> Установить
        ReadyToInstall  // Иконка Установки
    }

    // ---- Живые данные программы ----
    readonly property bool hasUpdate: backend.remoteUpdateAvailable
    readonly property bool downloading: backend.remoteDownloading
    readonly property int progress: backend.remoteDownloadProgress
    readonly property bool ready: backend.updateReady
    // updateChecking — живой флаг полной проверки (Main.py)
    readonly property bool checking: backend.updateChecking
    // busy — распаковка/подготовка скачанного обновления
    readonly property bool busy: backend.updateBusy

    // Помним, что загрузка уже шла: пока идёт распаковка
    // (downloading=false, busy=true) — держим кольцо заполненным
    property bool everDownloaded: false
    readonly property bool stagingAfterDownload:
        everDownloaded && busy && !ready && !downloading && hasUpdate

    // Куда стремимся по данным программы
    readonly property int targetState:
        downloading || stagingAfterDownload
            ? UpdateMorphButton.State.Downloading
        : ready ? UpdateMorphButton.State.ReadyToInstall
        : (checking || busy) && !hasUpdate ? UpdateMorphButton.State.Checking
        : hasUpdate ? UpdateMorphButton.State.DownloadReady
        : UpdateMorphButton.State.Idle

    // Что рисуем сейчас (может быть состоянием-морфингом)
    property int currentState: UpdateMorphButton.State.Idle

    // Прогресс 0..1 — сглажен, скачки процентов не щёлкают
    property real downloadProgress: 0.0
    Behavior on downloadProgress {
        NumberAnimation { duration: 250; easing.type: Easing.OutCubic }
    }

    // ---- Цвета: тема программы ----
    readonly property color primaryColor: AppTheme.accentBrand
    readonly property color successColor: AppTheme.accentSuccess

    readonly property string tipText: {
        if (downloading)
            return "Загрузка обновления… " + progress + "%"
        if (stagingAfterDownload)
            return "Готовим обновление…"
        if (ready)
            return "Обновление загружено. Нажмите, чтобы установить."
        if (hasUpdate)
            return "Доступна новая версия программы. Нажмите, чтобы загрузить."
        if (checking)
            return "Проверяем обновления…"
        return "Проверить обновления"
    }

    // ---- Синхронизация прогресса ----
    function syncProgress() {
        downloadProgress = downloading ? progress / 100
                       : (ready || stagingAfterDownload) ? 1
                       : 0
    }
    onProgressChanged: syncProgress()
    onDownloadingChanged: {
        everDownloaded = everDownloaded || downloading
        syncProgress()
    }
    onReadyChanged: syncProgress()
    onStagingAfterDownloadChanged: syncProgress()

    // ---- Подложка (наведение/нажатие) ----
    Rectangle {
        id: bgCircle
        anchors.fill: parent
        radius: width / 2
        color: root.currentState === UpdateMorphButton.State.ReadyToInstall
               ? root.successColor : root.primaryColor
        opacity: mouseArea.containsMouse ? 0.15 : 0.08
        Behavior on color { ColorAnimation { duration: 400 } }
        Behavior on opacity { NumberAnimation { duration: 200 } }
    }

    // ---- Векторный холст для бесшовной мутации ----
    Canvas {
        id: canvas
        objectName: "morphCanvas"
        anchors.fill: parent
        antialiasing: true

        property real morphT: 0.0       // Прогресс анимации мутации (0 -> 1)
        property real spinAngle: 0.0    // Угол вращения при проверке

        // сетка рисунка 44×44 — масштаб под любой размер кнопки
        readonly property real fit: Math.min(width, height) / 44

        onPaint: {
            var ctx = getContext("2d");
            ctx.reset();
            ctx.lineCap = "round";
            ctx.lineJoin = "round";

            var cx = width / 2;
            var cy = height / 2;
            var s = fit;
            var strokeColor = root.currentState === UpdateMorphButton.State.ReadyToInstall
                              ? root.successColor : root.primaryColor;

            ctx.strokeStyle = strokeColor;
            ctx.fillStyle = strokeColor;

            // ----------------------------------------------------
            // 1. СОСТОЯНИЕ: ПРОВЕРКА (Вращающаяся дуга)
            // ----------------------------------------------------
            if (root.currentState === UpdateMorphButton.State.Idle ||
                root.currentState === UpdateMorphButton.State.Checking) {

                ctx.save();
                ctx.translate(cx, cy);
                ctx.scale(s, s);
                ctx.rotate(spinAngle * Math.PI / 180);

                ctx.lineWidth = 3.5;
                ctx.beginPath();
                ctx.arc(0, 0, 18, 0.2 * Math.PI, 1.75 * Math.PI);
                ctx.stroke();

                // Стрелка на конце дуги
                var ax = 18 * Math.cos(0.2 * Math.PI);
                var ay = 18 * Math.sin(0.2 * Math.PI);
                ctx.beginPath();
                ctx.moveTo(ax - 2, ay - 8);
                ctx.lineTo(ax + 2, ay + 2);
                ctx.lineTo(ax - 8, ay + 2);
                ctx.closePath();
                ctx.fill();

                ctx.restore();
            }

            // ----------------------------------------------------
            // 2. МУТАЦИЯ: Проверка -> Скачать
            // ----------------------------------------------------
            else if (root.currentState === UpdateMorphButton.State.MorphToDownload) {
                var t = easeInOutCubic(morphT);

                ctx.save();
                ctx.translate(cx, cy);
                ctx.scale(s, s);
                ctx.lineWidth = 3.5;

                // Дуга распрямляется и превращается в стрелку вниз
                var startAngle = (0.2 * (1 - t)) * Math.PI;
                var endAngle = (1.75 * (1 - t) + 0.5 * t) * Math.PI;
                var radius = 18 * (1 - t);

                if (radius > 2) {
                    ctx.beginPath();
                    ctx.arc(0, 0, radius, startAngle, endAngle);
                    ctx.stroke();
                }

                // Вертикальная линия стрелки (появляется из центра)
                ctx.beginPath();
                ctx.moveTo(0, -18 * t);
                ctx.lineTo(0, 8 * t);
                ctx.stroke();

                // Птичка стрелки вниз
                ctx.beginPath();
                ctx.moveTo(-8 * t, 0);
                ctx.lineTo(0, 8 * t);
                ctx.lineTo(8 * t, 0);
                ctx.stroke();

                // Подставка (платформа)
                ctx.beginPath();
                ctx.moveTo(-12 * t, 16 * t);
                ctx.lineTo(12 * t, 16 * t);
                ctx.stroke();

                ctx.restore();
            }

            // ----------------------------------------------------
            // 3. СОСТОЯНИЕ: СКАЧАТЬ (Иконка готова)
            // ----------------------------------------------------
            else if (root.currentState === UpdateMorphButton.State.DownloadReady) {
                ctx.save();
                ctx.translate(cx, cy);
                ctx.scale(s, s);
                ctx.lineWidth = 3.5;

                // Стержень
                ctx.beginPath();
                ctx.moveTo(0, -16);
                ctx.lineTo(0, 8);
                ctx.stroke();

                // Стрелка
                ctx.beginPath();
                ctx.moveTo(-8, 0);
                ctx.lineTo(0, 8);
                ctx.lineTo(8, 0);
                ctx.stroke();

                // Платформа
                ctx.beginPath();
                ctx.moveTo(-12, 16);
                ctx.lineTo(12, 16);
                ctx.stroke();

                ctx.restore();
            }

            // ----------------------------------------------------
            // 4. МУТАЦИЯ: Скачать -> Круг прогресса
            // ----------------------------------------------------
            else if (root.currentState === UpdateMorphButton.State.MorphToProgress) {
                var t2 = easeInOutCubic(morphT);

                ctx.save();
                ctx.translate(cx, cy);
                ctx.scale(s, s);
                ctx.lineWidth = 3.5;

                // Стрелка сжимается в центр
                ctx.globalAlpha = 1 - t2;
                ctx.beginPath();
                ctx.moveTo(0, -16 * (1 - t2));
                ctx.lineTo(0, 8 * (1 - t2));
                ctx.moveTo(-8 * (1 - t2), 0);
                ctx.lineTo(0, 8 * (1 - t2));
                ctx.lineTo(8 * (1 - t2), 0);
                ctx.stroke();

                // Платформа сворачивается в дугу
                ctx.globalAlpha = 1;
                var rProgress = 20;
                ctx.strokeStyle = Qt.rgba(strokeColor.r, strokeColor.g, strokeColor.b, 0.2);
                ctx.beginPath();
                ctx.arc(0, 0, rProgress, 0, 2 * Math.PI * t2);
                ctx.stroke();

                ctx.restore();
            }

            // ----------------------------------------------------
            // 5. СОСТОЯНИЕ: ЗАГРУЗКА (Заполняющаяся окружность)
            // ----------------------------------------------------
            else if (root.currentState === UpdateMorphButton.State.Downloading) {
                ctx.save();
                ctx.translate(cx, cy);
                ctx.scale(s, s);
                ctx.lineWidth = 4;

                // Фоновое кольцо
                ctx.strokeStyle = Qt.rgba(strokeColor.r, strokeColor.g, strokeColor.b, 0.2);
                ctx.beginPath();
                ctx.arc(0, 0, 20, 0, 2 * Math.PI);
                ctx.stroke();

                // Кольцо прогресса
                ctx.strokeStyle = strokeColor;
                ctx.beginPath();
                ctx.arc(0, 0, 20, -0.5 * Math.PI, (-0.5 + 2 * root.downloadProgress) * Math.PI);
                ctx.stroke();

                ctx.restore();
            }

            // ----------------------------------------------------
            // 6. МУТАЦИЯ: Круг -> Иконка Установки (Галочка)
            // ----------------------------------------------------
            else if (root.currentState === UpdateMorphButton.State.MorphToInstall) {
                var t3 = easeInOutCubic(morphT);

                ctx.save();
                ctx.translate(cx, cy);
                ctx.scale(s, s);
                ctx.lineWidth = 4;

                // Внешний круг
                ctx.beginPath();
                ctx.arc(0, 0, 20, 0, 2 * Math.PI);
                ctx.stroke();

                // Появление галочки в центре
                ctx.beginPath();
                var startX = -8;
                var startY = 0;
                var midX = -2;
                var midY = 6;
                var endX = 8;
                var endY = -6;

                ctx.moveTo(startX, startY);
                if (t3 < 0.5) {
                    var p1 = t3 * 2;
                    ctx.lineTo(startX + (midX - startX) * p1, startY + (midY - startY) * p1);
                } else {
                    var p2 = (t3 - 0.5) * 2;
                    ctx.lineTo(midX, midY);
                    ctx.lineTo(midX + (endX - midX) * p2, midY + (endY - midY) * p2);
                }
                ctx.stroke();

                ctx.restore();
            }

            // ----------------------------------------------------
            // 7. СОСТОЯНИЕ: ГОТОВО К УСТАНОВКЕ
            // ----------------------------------------------------
            else if (root.currentState === UpdateMorphButton.State.ReadyToInstall) {
                ctx.save();
                ctx.translate(cx, cy);
                ctx.scale(s, s);
                ctx.lineWidth = 4;

                // Круг
                ctx.beginPath();
                ctx.arc(0, 0, 20, 0, 2 * Math.PI);
                ctx.stroke();

                // Галочка
                ctx.beginPath();
                ctx.moveTo(-8, 0);
                ctx.lineTo(-2, 6);
                ctx.lineTo(8, -6);
                ctx.stroke();

                ctx.restore();
            }
        }

        function easeInOutCubic(t) {
            return t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2;
        }
    }

    // ----------------------------------------------------
    // АНИМАЦИИ
    // ----------------------------------------------------

    // 1. Непрерывное вращение дуги при проверке.
    //    forceSpin: клик обязан дать видимый отклик — даже если
    //    проверка мгновенная, дуга прокручивается не меньше секунды
    property bool forceSpin: false
    Timer {
        id: minSpinTimer
        interval: 900
        onTriggered: root.forceSpin = false
    }

    NumberAnimation {
        id: spinAnim
        target: canvas
        property: "spinAngle"
        from: 0; to: 360
        duration: 900
        loops: Animation.Infinite
        running: root.currentState === UpdateMorphButton.State.Checking || root.forceSpin
        onRunningChanged: if (!running) canvas.spinAngle = 0
    }

    // 2. Анимация прогресса мутации.
    //    Конец морфинга устанавливает следующее устойчивое
    //    состояние (как в исходном коде), после чего контроллер
    //    догоняет цель, если данные программы уже ушли дальше.
    NumberAnimation {
        id: morphAnim
        target: canvas
        property: "morphT"
        from: 0.0; to: 1.0
        duration: 550
        easing.type: Easing.InOutCubic
        onFinished: {
            if (root.currentState === UpdateMorphButton.State.MorphToDownload) {
                root.currentState = UpdateMorphButton.State.DownloadReady;
            } else if (root.currentState === UpdateMorphButton.State.MorphToProgress) {
                root.currentState = UpdateMorphButton.State.Downloading;
            } else if (root.currentState === UpdateMorphButton.State.MorphToInstall) {
                root.currentState = UpdateMorphButton.State.ReadyToInstall;
            }
            root.applyTarget();
        }
    }

    Connections {
        target: canvas
        function onSpinAngleChanged() { canvas.requestPaint(); }
        function onMorphTChanged() { canvas.requestPaint(); }
    }

    onDownloadProgressChanged: canvas.requestPaint()
    onCurrentStateChanged: canvas.requestPaint()

    // ----------------------------------------------------
    // КОНТРОЛЛЕР: данные программы -> состояния кнопки.
    // Антидребезг 80 мс: «загрузка кончилась» и «готово» могут
    // прийти двумя событиями подряд — без паузы проигрался бы
    // лишний промежуточный кадр.
    // ----------------------------------------------------
    Timer {
        id: stateDebounce
        interval: 80
        onTriggered: root.applyTarget()
    }
    onTargetStateChanged: stateDebounce.restart()

    Component.onCompleted: {
        everDownloaded = ready
        syncProgress()
        currentState = targetState   // старт сразу в нужное состояние
    }

    function applyTarget() {
        var to = targetState;
        var from = currentState;
        if (to === from)
            return;
        // морфинг сейчас доедет и сам вызовет applyTarget
        if (from === UpdateMorphButton.State.MorphToDownload ||
            from === UpdateMorphButton.State.MorphToProgress ||
            from === UpdateMorphButton.State.MorphToInstall)
            return;
        // содержательное состояние — минимальный прокрут не нужен
        if (to !== UpdateMorphButton.State.Idle &&
            to !== UpdateMorphButton.State.Checking) {
            forceSpin = false;
            minSpinTimer.stop();
        }
        // возврат в покой — забыть про прошедшую загрузку
        if (to === UpdateMorphButton.State.Idle)
            everDownloaded = false;
        // цепочка морфингов — только вперёд по соседям
        if ((from === UpdateMorphButton.State.Idle ||
             from === UpdateMorphButton.State.Checking) &&
            to === UpdateMorphButton.State.DownloadReady) {
            currentState = UpdateMorphButton.State.MorphToDownload;
            morphAnim.restart();
            return;
        }
        if (from === UpdateMorphButton.State.DownloadReady &&
            to === UpdateMorphButton.State.Downloading) {
            currentState = UpdateMorphButton.State.MorphToProgress;
            morphAnim.restart();
            return;
        }
        if (from === UpdateMorphButton.State.Downloading &&
            to === UpdateMorphButton.State.ReadyToInstall) {
            currentState = UpdateMorphButton.State.MorphToInstall;
            morphAnim.restart();
            return;
        }
        // остальные переходы (возвраты, сбои) — сразу, без трюков
        morphAnim.stop();
        canvas.morphT = 0;
        currentState = to;
        canvas.requestPaint();
    }

    // ----------------------------------------------------
    // МЫШЬ
    // ----------------------------------------------------
    MouseArea {
        id: mouseArea
        anchors.fill: parent
        hoverEnabled: true
        cursorShape: (root.currentState === UpdateMorphButton.State.Idle ||
                      root.currentState === UpdateMorphButton.State.DownloadReady ||
                      root.currentState === UpdateMorphButton.State.ReadyToInstall)
                     ? Qt.PointingHandCursor : Qt.ArrowCursor

        onClicked: root.handleClick()
    }

    // Что делает клик в каждом состоянии (настоящие действия)
    function handleClick() {
        if (currentState === UpdateMorphButton.State.Idle) {
            // Клик 1: проверка обновлений
            forceSpin = true;
            minSpinTimer.restart();
            backend.checkAllUpdateSources();
        }
        else if (currentState === UpdateMorphButton.State.DownloadReady) {
            // Клик 2: загрузка; морфинг в кольцо запустит
            // контроллер по событию downloading
            backend.startRemoteDownload();
        }
        else if (currentState === UpdateMorphButton.State.ReadyToInstall) {
            // Клик 3: установка (программа перезапустится)
            backend.applyReadyUpdate();
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
}
