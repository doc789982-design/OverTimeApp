import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

Popup {
    id: root

    // --- Настройки интерфейса окна ---
    property string title: "Заголовок"
    property string acceptText: "Сохранить"
    property string rejectText: "Отмена"
    property string acceptVariant: "primary"
    property string rejectVariant: "secondary"
    property bool showFooter: true
    property bool showAccept: true
    property bool showReject: true
    
    // --- Сигналы ---
    signal accepted()
    signal rejected()

    // Режим «морфинга»: диалог открывается сразу в заданном месте
    // без анимации масштаба, чтобы плавно продолжить рост ячейки дня.
    property bool morphOpen: false

    // Зона для вставки контента
    default property alias dialogContent: contentArea.data

    width: 380 
    // Окно не может быть выше экрана, вычитаем 100px для комфортного отступа
    property real effectiveHeight: Math.min(mainLayout.implicitHeight + 40, ApplicationWindow.window ? ApplicationWindow.window.height - 100 : 800)
    height: effectiveHeight
    
    z: AppTheme.zModal
    modal: false 
    dim: false  
    focus: true
    // Закрываем только по Esc / крестику / «Отмена».
    // Клик мимо окна НЕ закрывает форму: случайный щелчок больше не сжигает
    // введённые дежурства, перерывы и балансы (стандарт Apple/Google для форм).
    closePolicy: Popup.CloseOnEscape

    // Анимации появления
    enter: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
            // При морфинге масштаб уже 1.0 (from == to) — анимации роста нет,
            // окно просто продолжает рост ячейки как единое целое.
            NumberAnimation { property: "scale"; from: root.morphOpen ? 1.0 : 0.95; to: 1.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
        }
    }
    exit: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
            NumberAnimation { property: "scale"; from: 1.0; to: 0.95; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
        }
    }

    background: Rectangle {
        color: AppTheme.bgModal 
        radius: AppTheme.radiusModal 
        border.color: AppTheme.borderDivider
        border.width: 1
        
        // Тень-картинка вместо вычисляемой (легко для видеокарты)
        AppShadow { level: 4 }
    }

    // Тряска при ошибке
    property real baseShakeX: 0
    function shake() { if (!shakeAnimation.running) { baseShakeX = root.x; shakeAnimation.start() } }
    SequentialAnimation {
        id: shakeAnimation
        NumberAnimation { target: root; property: "x"; to: baseShakeX + 10; duration: 50; easing.type: Easing.OutQuad }
        NumberAnimation { target: root; property: "x"; to: baseShakeX - 10; duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: root; property: "x"; to: baseShakeX + 8;  duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: root; property: "x"; to: baseShakeX - 8;  duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: root; property: "x"; to: baseShakeX + 4;  duration: 50; easing.type: Easing.InOutQuad }
        NumberAnimation { target: root; property: "x"; to: baseShakeX;      duration: 50; easing.type: Easing.OutQuad }
    }

    contentItem: ColumnLayout {
        id: mainLayout
        spacing: 0

        // ШАПКА
        Item {
            Layout.fillWidth: true
            Layout.preferredHeight: 60
            
            Text {
                anchors.left: parent.left
                anchors.leftMargin: AppTheme.spaceL
                anchors.verticalCenter: parent.verticalCenter
                width: parent.width - 80 
                text: root.title
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeH4 
                font.weight: AppTheme.weightBold
                visible: text !== ""
                elide: Text.ElideRight
            }

            Rectangle {
                width: 32; height: 32; radius: AppTheme.radiusPill
                anchors.right: parent.right
                anchors.rightMargin: AppTheme.spaceM
                anchors.verticalCenter: parent.verticalCenter
                
                color: closeHov.pressed ? AppTheme.statePress : (closeHov.containsMouse ? AppTheme.stateHover : "transparent")
                IconImage { anchors.centerIn: parent; source: "../icons/close.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: AppTheme.textSecondary }
                MouseArea { id: closeHov; anchors.fill: parent; hoverEnabled: true; onClicked: { root.rejected(); root.close() } }
            }
        }

        // КОНТЕНТ (Скроллируемый)
        ScrollView {
            Layout.fillWidth: true
            Layout.preferredHeight: Math.min(
                contentArea.implicitHeight,
                ApplicationWindow.window ? ApplicationWindow.window.height - 280 : 520
            )
            clip: true
            ScrollBar.horizontal.policy: ScrollBar.AlwaysOff

            Column {
                id: contentArea
                width: root.width - (AppTheme.spaceL * 2)
                x: AppTheme.spaceL
                spacing: AppTheme.spaceM
                // Немного отступа сверху для красоты внутри скролла
                topPadding: AppTheme.spaceXS 
                bottomPadding: AppTheme.spaceL

                // ПОДВАЛ (кнопки «Отмена» / «Сохранить») — едут вместе с
                // содержимым в скролле, чтобы при длинном контенте не
                // вылезать за края окна.
                Row {
                    anchors.horizontalCenter: parent.horizontalCenter
                    spacing: AppTheme.spaceM
                    visible: root.showFooter

                    AppButton {
                        visible: root.showReject
                        text: root.rejectText
                        variant: root.rejectVariant
                        onClicked: { root.rejected(); root.close() }
                    }
                    AppButton {
                        visible: root.showAccept
                        text: root.acceptText
                        variant: root.acceptVariant
                        onClicked: root.accepted()
                    }
                }
            }
        }
    }
    
    // Функции показа (оставляем как были)
    function showAt(callerItem, mouseX, mouseY) {
        root.morphOpen = false
        root.height = root.effectiveHeight
        var clickPoint = callerItem.mapToItem(null, mouseX, mouseY)
        root.x = Math.max(AppTheme.spaceL, Math.min(clickPoint.x, ApplicationWindow.window.width - root.width - AppTheme.spaceL))
        root.y = Math.max(AppTheme.spaceL, Math.min(clickPoint.y + AppTheme.spaceM, ApplicationWindow.window.height - root.height - AppTheme.spaceL))
        root.open()
    }

    // Открытие после «морфинга»: диалог появляется сразу в заданном прямоугольнике,
    // без анимации масштаба — продолжает рост ячейки дня как одно целое окно.
    function openMorph(x, y, w, h) {
        root.morphOpen = true
        root.x = x
        root.y = y
        root.width = w
        root.height = h
        root.open()
    }
    
    function showCentered() {
        if (ApplicationWindow.window) {
            root.x = (ApplicationWindow.window.width - root.width) / 2
            root.y = (ApplicationWindow.window.height - root.height) / 2
        }
        root.open()
    }
}