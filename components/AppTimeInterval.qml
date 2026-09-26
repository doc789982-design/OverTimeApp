import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

Item {
    id: root

    property int startMinutes: 480 
    property int endMinutes: 1200  

    width: 280
    height: 60 

    StackLayout {
        anchors.fill: parent
        currentIndex: backend.timeInputMode === "tumbler" ? 1 : 0

        // РЕЖИМ 0: Слайдер
        AppRangeSlider {
            id: rangeSlider
            Layout.fillWidth: true
            Layout.fillHeight: true
            onStartMinutesChanged: { if (root.startMinutes !== startMinutes) root.startMinutes = startMinutes }
            onEndMinutesChanged: { if (root.endMinutes !== endMinutes) root.endMinutes = endMinutes }
        }

        // РЕЖИМ 1: Барабаны
        Row {
            Layout.alignment: Qt.AlignCenter
            spacing: AppTheme.spaceM

            Column {
                spacing: AppTheme.spaceXXS
                Text {
                    text: "НАЧАЛО"
                    color: AppTheme.textSecondary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                    font.letterSpacing: 1
                    font.weight: AppTheme.weightBold // Жирный, как положено мелким подписям
                    anchors.horizontalCenter: parent.horizontalCenter
                }
                AppTumblerTime {
                    id: startTumbler
                    onHoursChanged: root.updateStartFromTumbler()
                    onMinutesChanged: root.updateStartFromTumbler()
                }
            }

            Text {
                text: "—"
                color: AppTheme.borderInput // Разделитель такого же цвета, как рамки
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeH2
                anchors.verticalCenter: parent.verticalCenter
                anchors.verticalCenterOffset: 12 
            }

            Column {
                spacing: AppTheme.spaceXXS
                Text {
                    text: "КОНЕЦ"
                    color: AppTheme.textSecondary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall  
                    font.letterSpacing: 1
                    font.weight: AppTheme.weightBold
                    anchors.horizontalCenter: parent.horizontalCenter
                }
                AppTumblerTime {
                    id: endTumbler
                    onHoursChanged: root.updateEndFromTumbler()
                    onMinutesChanged: root.updateEndFromTumbler()
                }
            }
        }
    }

    function updateStartFromTumbler() {
        var newTotal = startTumbler.hours * 60 + startTumbler.minutes
        if (root.startMinutes !== newTotal) root.startMinutes = newTotal
    }

    function updateEndFromTumbler() {
        var newTotal = endTumbler.hours * 60 + endTumbler.minutes
        if (root.endMinutes !== newTotal) root.endMinutes = newTotal
    }

    // Проталкивание внешних значений в оба режима ввода. Декларативные
    // привязки здесь не годятся: любой ручной ввод (колёсико барабана,
    // перетаскивание ползунка) записывает свойство изнутри и рвёт
    // привязку. После этого окно редактирования другого человека
    // показывало прошлое время, а «Сохранить» без «прокрутки туда-сюда»
    // отправлял старое значение. Теперь внешняя установка времени
    // обновляет барабаны и слайдер всегда.
    function syncValues() {
        startTumbler.setTime(Math.floor(root.startMinutes / 60), root.startMinutes % 60)
        endTumbler.setTime(Math.floor(root.endMinutes / 60), root.endMinutes % 60)
        if (rangeSlider.startMinutes !== root.startMinutes) rangeSlider.startMinutes = root.startMinutes
        if (rangeSlider.endMinutes !== root.endMinutes) rangeSlider.endMinutes = root.endMinutes
    }

    onStartMinutesChanged: syncValues()
    onEndMinutesChanged: syncValues()
    Component.onCompleted: syncValues()
}