import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl

AppLargeModal {
    id: root
    width: 680 // Сделали чуть шире для комфорта двух колонок
    height: 440 
    title: "Настройки печати"

    // Пункт «печати в файл»: экспорт табеля в Excel, как раньше делала
    // отдельная кнопка на панели. Стоит последним в списке принтеров.
    readonly property string excelItem: "Экспортировать в Excel"
    property bool isExcel: printerCombo.currentText === excelItem

    signal printRequested(string printerName, int copies, string pageFrom, string pageTo, string orientation, string paperSize, bool collate)

    onOpened: {
        backend.loadPrinters() 
        var defIdx = 0
        var list = printerCombo.model
        for (var i = 0; i < list.length; i++) {
            if (list[i] === backend.defaultPrinter) { defIdx = i; break }
        }
        printerCombo.currentIndex = defIdx
        copiesInput.text = "1"
        pageRangeCombo.currentIndex = 0
        pageFromInput.text = ""
        pageToInput.text = ""
    }

    // Диапазон заполнен и корректен (С ≥ 1, По ≥ С), либо не используется
    property bool rangeOk: pageRangeCombo.currentValue !== "range" ||
                           (pageFromInput.text !== "" && pageToInput.text !== "" &&
                            parseInt(pageFromInput.text) >= 1 &&
                            parseInt(pageToInput.text) >= parseInt(pageFromInput.text))

    Row {
        anchors.fill: parent
        anchors.margins: AppTheme.spaceL 
        spacing: AppTheme.spaceL

        // ЛЕВАЯ КОЛОНКА
        Column {
            width: (parent.width - AppTheme.spaceL - 1) / 2
            spacing: AppTheme.spaceM

            AppComboBox {
                id: printerCombo
                width: parent.width
                label: "Принтер:"
                model: backend.printerList ? backend.printerList.concat([root.excelItem]) : [root.excelItem]
            }

            RowLayout {
                width: parent.width; spacing: AppTheme.spaceM
                opacity: root.isExcel ? 0.4 : 1.0
                AppTextField { id: copiesInput; Layout.preferredWidth: 80; label: "Копии:"; text: "1"; enabled: !root.isExcel; validator: RegularExpressionValidator { regularExpression: /[1-9][0-9]?/ } }
                AppCheckBox { id: collateCheck; text: "Разобрать по копиям\n(1,2,3  1,2,3)"; checked: true; Layout.alignment: Qt.AlignBottom; enabled: !root.isExcel && parseInt(copiesInput.text) > 1 }
            }
            
            Item { height: 1; width: 1; Layout.fillHeight: true } 
            
            AppButton {
                text: root.isExcel ? "Экспортировать" : "Отправить на печать"
                iconSource: root.isExcel ? "../icons/export_box.svg" : "../icons/print.svg" 
                width: parent.width
                variant: "primary"
                enabled: printerCombo.currentIndex >= 0 && root.rangeOk
                opacity: enabled ? 1.0 : 0.4
                onClicked: executePrint()
            }

            Text {
                visible: !root.rangeOk
                width: parent.width
                wrapMode: Text.WordWrap
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeSmall
                color: AppTheme.accentDanger
                text: "Заполните диапазон: страницы с 1, «по» — не меньше «с»"
            }
        }

        // РАЗДЕЛИТЕЛЬ (Используем системный цвет)
        Rectangle { width: 1; height: parent.height; color: AppTheme.borderDivider }

        // ПРАВАЯ КОЛОНКА
        Column {
            width: (parent.width - AppTheme.spaceL - 1) / 2
            spacing: AppTheme.spaceM
            opacity: root.isExcel ? 0.4 : 1.0

            AppComboBox {
                id: pageRangeCombo; width: parent.width; label: "Страницы:"
                model: [{ text: "Все страницы", value: "all" }, { text: "Заданный диапазон", value: "range" }]
                textRole: "text"; valueRole: "value"
                enabled: !root.isExcel
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
                enabled: !root.isExcel
            }

            AppComboBox {
                id: paperCombo; width: parent.width; label: "Формат бумаги:"
                model: [{ text: "A4 (210 x 297 мм)", value: "A4" }, { text: "A3 (297 x 420 мм)", value: "A3" }]
                textRole: "text"; valueRole: "value"
                enabled: !root.isExcel
            }
        }
    }

    function executePrint() {
        if (!root.rangeOk)
            return
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
