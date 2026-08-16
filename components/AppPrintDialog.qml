import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl

AppLargeModal {
    id: root
    width: 680 // Сделали чуть шире для комфорта двух колонок
    height: 440 
    title: "Настройки печати"

    signal printRequested(string printerName, int copies, string pageFrom, string pageTo, string orientation, string paperSize, bool collate)

    onOpened: {
        backend.loadPrinters() 
        var defIdx = 0
        for (var i = 0; i < backend.printerList.length; i++) {
            if (backend.printerList[i] === backend.defaultPrinter) { defIdx = i; break }
        }
        printerCombo.currentIndex = defIdx
        copiesInput.text = "1"
        pageRangeCombo.currentIndex = 0
        pageFromInput.text = ""
        pageToInput.text = ""
    }

    Row {
        anchors.fill: parent
        anchors.margins: AppTheme.spaceL 
        spacing: AppTheme.spaceL

        // ЛЕВАЯ КОЛОНКА
        Column {
            width: (parent.width - AppTheme.spaceL - 1) / 2
            spacing: AppTheme.spaceM

            AppComboBox { id: printerCombo; width: parent.width; label: "Принтер:"; model: backend.printerList }

            RowLayout {
                width: parent.width; spacing: AppTheme.spaceM
                AppTextField { id: copiesInput; Layout.preferredWidth: 80; label: "Копии:"; text: "1"; validator: RegularExpressionValidator { regularExpression: /[1-9][0-9]?/ } }
                AppCheckBox { id: collateCheck; text: "Разобрать по копиям\n(1,2,3  1,2,3)"; checked: true; Layout.alignment: Qt.AlignBottom; enabled: parseInt(copiesInput.text) > 1 }
            }
            
            Item { height: 1; width: 1; Layout.fillHeight: true } 
            
            AppButton {
                text: "Отправить на печать" 
                iconSource: "../icons/print.svg" 
                width: parent.width
                height: 45 // Главная кнопка действия всегда крупнее
                variant: "primary" 
                onClicked: executePrint()
            }
        }

        // РАЗДЕЛИТЕЛЬ (Используем системный цвет)
        Rectangle { width: 1; height: parent.height; color: AppTheme.borderDivider }

        // ПРАВАЯ КОЛОНКА
        Column {
            width: (parent.width - AppTheme.spaceL - 1) / 2
            spacing: AppTheme.spaceM

            AppComboBox {
                id: pageRangeCombo; width: parent.width; label: "Страницы:"
                model: [{ text: "Все страницы", value: "all" }, { text: "Заданный диапазон", value: "range" }]
                textRole: "text"; valueRole: "value"
            }

            RowLayout {
                visible: pageRangeCombo.currentValue === "range"; width: parent.width; spacing: AppTheme.spaceM
                AppTextField { id: pageFromInput; Layout.fillWidth: true; label: "С:"; validator: RegularExpressionValidator { regularExpression: /[1-9][0-9]*/ } }
                AppTextField { id: pageToInput; Layout.fillWidth: true; label: "По:"; validator: RegularExpressionValidator { regularExpression: /[1-9][0-9]*/ } }
            }

            AppComboBox {
                id: orientCombo; width: parent.width; label: "Ориентация:"
                model: [{ text: "Книжная", value: "portrait" }, { text: "Альбомная", value: "landscape" }]
                textRole: "text"; valueRole: "value"; currentIndex: 1 
            }

            AppComboBox {
                id: paperCombo; width: parent.width; label: "Формат бумаги:"
                model: [{ text: "A4 (210 x 297 мм)", value: "A4" }, { text: "A3 (297 x 420 мм)", value: "A3" }]
                textRole: "text"; valueRole: "value"
            }
        }
    }

    function executePrint() {
        var copies = parseInt(copiesInput.text)
        if (isNaN(copies) || copies < 1) copies = 1
        
        root.printRequested(
            printerCombo.currentText, copies, 
            pageRangeCombo.currentValue === "range" ? pageFromInput.text : "", 
            pageRangeCombo.currentValue === "range" ? pageToInput.text : "", 
            orientCombo.currentValue, paperCombo.currentValue, collateCheck.checked
        )
        root.close()
    }
}