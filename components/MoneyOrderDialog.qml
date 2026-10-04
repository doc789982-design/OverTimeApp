import QtQuick
import QtQuick.Controls
import "."
import "../icons"

// ============================================================
// ВЕДОМОСТЬ: один денежный приказ — всему подразделению сразу.
//
// Из левой панели («₽ Приказ всем») — ведомость на ВСЕХ: шапка
// приказа вводится один раз, у каждого сотрудника галочка и
// СВОИ суммы (часы, сверхурочные, дни). Заполнение: «По табелю»
// (каждому его переработку) или «Одинаково всем…»; галочка в
// заголовке таблицы выбирает и снимает всех сразу.
// Из карточки сотрудника («Деньгами») — тот же приказ, но
// ТОЛЬКО для него: без галочек и без «одинаково всем».
// Номер и дата приказа обязательны. При ошибке окно трясётся
// (штатная тряска AppDialog), а строка сотрудника, у которого
// не хватает остатка, подсвечивается красным с пояснением.
// ============================================================
AppDialog {
    id: root
    width: 640

    // одиночный режим: приказ только для одного сотрудника
    // (открывается кнопкой «Деньгами» из его карточки)
    property bool singleMode: false

    // «Журнал приказов» — история приказов подразделения
    signal requestJournal()

    title: root.singleMode && root.rows.length === 1
           ? "Денежная компенсация — " + root.shortName(root.rows[0].name)
           : "Приказ о денежной компенсации"
    acceptText: "Провести приказ"
    acceptVariant: "primary"
    rejectText: "Закрыть"

    // [{id, name, subtitle, checked, hours, overtime, days, balHours, balOvertime, balDays}]
    property var rows: []
    property var rowErrors: ({})     // id сотрудника → текст ошибки
    property string topError: ""
    property bool showEqual: false   // панель «одинаково всем»
    property int totalEmp: 0
    property int totalHours: 0
    property int totalDays: 0

    // ширины колонок таблицы (одинаково в шапке и строках)
    readonly property int colNum: 56
    readonly property int colBal: 96
    readonly property int nameWidth: width - AppTheme.spaceL * 2 - 28
                                     - colNum * 3 - colBal - 8 * 5 - 16

    function _fromBackend() {
        let list = backend.moneyOrderEmployees()
        let res = []
        for (let i = 0; i < list.length; i++) {
            let e = list[i]
            res.push({
                id: e.id, name: e.name, subtitle: e.subtitle,
                checked: true, hours: 0, overtime: 0, days: 0,
                balHours: e.hours, balOvertime: e.overtime, balDays: e.days,
            })
        }
        return res
    }

    function _resetHead() {
        root.rowErrors = ({})
        root.topError = ""
        root.showEqual = false
        orderNoInput.text = ""
        orderNoInput.hasError = false
        orderCommentInput.text = ""
        orderDateInput.selectedDate = new Date().toISOString().split("T")[0]
        orderDateInput.hasError = false
    }

    function openNew() {
        root.singleMode = false
        root.rows = _fromBackend()
        _resetHead()
        recalcTotals()
        root.showCentered()
    }

    // Приказ только для одного сотрудника — из его карточки
    function openForEmployee(empId) {
        let all = _fromBackend()
        let only = []
        for (let i = 0; i < all.length; i++)
            if (all[i].id === empId) { only.push(all[i]); break }
        root.singleMode = true
        root.rows = only
        _resetHead()
        recalcTotals()
        root.showCentered()
    }

    // «Иванов Иван Иванович» → «Иванов И.» — для заголовка окна
    function shortName(fio) {
        let p = String(fio || "").trim().split(/\s+/)
        if (p.length > 1 && p[1].length > 0)
            return p[0] + " " + p[1][0] + "."
        return p[0] || ""
    }

    // Все ли строки отмечены (галочка «все» в заголовке таблицы)
    function _allChecked() {
        if (root.rows.length === 0) return false
        for (let i = 0; i < root.rows.length; i++)
            if (!root.rows[i].checked) return false
        return true
    }

    // Переключение галочки строки: вызывается из делегата, но живёт
    // в контексте окна — пересоздание делегата ей не страшно
    function toggleRow(i, on) {
        if (i < 0 || i >= root.rows.length) return
        root.rows[i].checked = on
        root.rows = root.rows.slice()
        root.recalcTotals()
    }

    // «Повторить приказ»: получатели и суммы — как в прошлый раз
    // Для проверок в песочнице: геометрия отрисованных строк
    function layoutInfo() { return empTable.layoutInfo() }

    function prefillFromOrder(order) {
        root.singleMode = false
        let current = _fromBackend()
        for (let i = 0; i < current.length; i++) {
            let row = current[i]
            row.checked = false
            for (let j = 0; j < order.employees.length; j++) {
                let e = order.employees[j]
                if (e.id === row.id) {
                    row.checked = true
                    row.hours = e.hours
                    row.overtime = e.overtime
                    row.days = e.days
                    break
                }
            }
        }
        root.rows = current
        root.rowErrors = ({})
        root.topError = ""
        root.showEqual = false
        orderNoInput.text = order.order_no
        orderCommentInput.text = order.comment
        orderDateInput.selectedDate = new Date().toISOString().split("T")[0]
        recalcTotals()
        root.showCentered()
    }

    function recalcTotals() {
        let n = 0, h = 0, d = 0
        for (let i = 0; i < rows.length; i++) {
            let r = rows[i]
            if (!r.checked) continue
            if (r.hours > 0 || r.overtime > 0 || r.days > 0) n++
            h += r.hours + r.overtime
            d += r.days
        }
        totalEmp = n
        totalHours = h
        totalDays = d
    }

    function _toInt(s) {
        let v = parseInt(String(s).replace(/[^\d]/g, ""), 10)
        return isNaN(v) ? 0 : Math.max(0, v)
    }

    onAccepted: {
        // номер и дата приказа обязательны
        orderNoInput.hasError = orderNoInput.text.trim() === ""
        orderDateInput.hasError = !(orderDateInput.selectedDate !== ""
                                     && /^(\d{2})\.(\d{2})\.(\d{4})$/.test(orderDateInput.text.trim()))
        if (orderNoInput.hasError || orderDateInput.hasError) {
            root.rowErrors = ({})
            root.topError = "Укажите номер и дату приказа"
            root.shake()
            return
        }

        let payload = []
        for (let i = 0; i < rows.length; i++) {
            let r = rows[i]
            if (!r.checked) continue
            if (r.hours <= 0 && r.overtime <= 0 && r.days <= 0) continue
            payload.push({ id: r.id, hours: r.hours, overtime: r.overtime, days: r.days })
        }
        if (payload.length === 0) {
            let anyChecked = false
            for (let i = 0; i < rows.length; i++)
                if (rows[i].checked) { anyChecked = true; break }
            root.rowErrors = ({})
            root.topError = anyChecked
                    ? "Укажите, сколько выплатить: часы, сверхурочные или дни отдыха"
                    : "Отметьте сотрудников галочками и укажите суммы"
            root.shake()
            return
        }

        let result = backend.saveMoneyOrder(
            JSON.stringify(payload),
            orderNoInput.text.trim(),
            orderDateInput.selectedDate,
            orderCommentInput.text.trim()
        )
        if (result && result.ok) {
            root.close()
            return
        }
        // Ошибка: окно трясётся (штатная тряска), виновники подсвечены
        let errs = ({})
        let top = ""
        if (result && result.errors) {
            for (let i = 0; i < result.errors.length; i++) {
                let e = result.errors[i]
                if (e.id > 0) errs[e.id] = e.message
                if (top === "") top = e.name !== "" ? (e.name + " — " + e.message) : e.message
            }
        }
        root.rowErrors = errs
        root.topError = top
        root.shake()
        root.scrollToBottom()
    }

    // ── Шапка приказа ──
    Row {
        width: parent.width
        spacing: AppTheme.spaceS

        AppTextField {
            id: orderNoInput
            width: 150
            label: "№ приказа"
            placeholderText: "245"
            maximumLength: 20
            onTextEdited: hasError = false
        }
        AppDateField {
            id: orderDateInput
            width: 190
            label: "Дата приказа"
            onSelectedDateChanged: hasError = false
        }
        AppTextField {
            id: orderCommentInput
            width: parent.width - 150 - 190 - AppTheme.spaceS * 2
            label: "Комментарий"
            placeholderText: "За октябрь"
        }
    }

    // ── Заполнение: понятные кнопки-глаголы ──
    Row {
        id: fillRow
        width: parent.width
        spacing: AppTheme.spaceS

        AppButton {
            id: btnBySheet
            text: "По табелю"
            iconSource: "../icons/edit.svg"
            variant: "secondary"
            height: 40
            onClicked: {
                for (let i = 0; i < root.rows.length; i++) {
                    if (!root.rows[i].checked) continue
                    root.rows[i].hours = root.rows[i].balHours
                    root.rows[i].overtime = root.rows[i].balOvertime
                    root.rows[i].days = root.rows[i].balDays
                }
                root.rows = root.rows.slice()
                root.recalcTotals()
            }
            AppToolTip {
                anchors.horizontalCenter: parent.horizontalCenter
                anchors.top: parent.bottom; anchors.topMargin: AppTheme.spaceXXS
                text: "Подставить каждому выбранному его переработку за месяц"
                isVisible: parent.hovered
            }
        }
        AppButton {
            id: btnEqual
            visible: !root.singleMode
            text: root.showEqual ? "Скрыть" : "Одинаково всем…"
            iconSource: "../icons/money.svg"
            variant: "secondary"
            height: 40
            onClicked: root.showEqual = !root.showEqual
            AppToolTip {
                anchors.horizontalCenter: parent.horizontalCenter
                anchors.top: parent.bottom; anchors.topMargin: AppTheme.spaceXXS
                text: "Одна и та же сумма каждому выбранному"
                isVisible: parent.hovered
            }
        }
        // распорка: «Журнал приказов» прижимается к правому краю
        Item {
            height: 1
            width: parent.width - btnBySheet.width - btnJournal.width
                     - parent.spacing * 2
                     - (btnEqual.visible ? btnEqual.width + parent.spacing : 0)
        }
        AppButton {
            id: btnJournal
            text: "Журнал приказов"
            iconSource: "../icons/clock.svg"
            variant: "secondary"
            height: 40
            onClicked: { root.close(); root.requestJournal() }
            AppToolTip {
                anchors.horizontalCenter: parent.horizontalCenter
                anchors.top: parent.bottom; anchors.topMargin: AppTheme.spaceXXS
                text: "Все прошлые приказы подразделения"
                isVisible: parent.hovered
            }
        }
    }

    // ── Панель «одинаково всем» ──
    Row {
        visible: root.showEqual
        width: parent.width
        spacing: AppTheme.spaceS

        AppTextField {
            id: equalHours
            width: 120
            label: "Часы"
            numericOnly: true
            horizontalAlignment: TextInput.AlignHCenter
        }
        AppTextField {
            id: equalOvertime
            width: 120
            label: "Сверхурочные"
            numericOnly: true
            horizontalAlignment: TextInput.AlignHCenter
        }
        AppTextField {
            id: equalDays
            width: 120
            label: "Дни отдыха"
            numericOnly: true
            horizontalAlignment: TextInput.AlignHCenter
        }
        AppButton {
            text: "Раздать выбранным"
            variant: "primary"
            height: 44
            anchors.verticalCenter: parent.verticalCenter
            onClicked: {
                let h = root._toInt(equalHours.text)
                let o = root._toInt(equalOvertime.text)
                let d = root._toInt(equalDays.text)
                for (let i = 0; i < root.rows.length; i++) {
                    if (!root.rows[i].checked) continue
                    root.rows[i].hours = h
                    root.rows[i].overtime = o
                    root.rows[i].days = d
                }
                root.rows = root.rows.slice()
                root.recalcTotals()
            }
        }
    }

    // ── Заголовок таблицы ──
    Row {
        id: tableHeader
        width: parent.width
        spacing: 8
        leftPadding: 8
        rightPadding: 8

        // галочка «выбрать всех / снять со всех»
        Item {
            width: 28; height: 16
            visible: !root.singleMode && root.rows.length > 0
            AppCheckBox {
                anchors.centerIn: parent
                checked: root._allChecked()
                onToggled: {
                    for (let i = 0; i < root.rows.length; i++)
                        root.rows[i].checked = checked
                    root.rows = root.rows.slice()
                    root.recalcTotals()
                }
                AppToolTip {
                    anchors.horizontalCenter: parent.horizontalCenter
                    anchors.bottom: parent.top; anchors.bottomMargin: AppTheme.spaceXXS
                    text: "Выбрать всех или снять выделение"
                    isVisible: parent.hovered
                }
            }
        }
        Text {
            width: root.nameWidth
            text: "СОТРУДНИК"
            color: AppTheme.textTertiary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeMicro
            font.weight: AppTheme.weightBold
            font.letterSpacing: 1.2
        }
        Text { width: root.colNum; horizontalAlignment: Text.AlignHCenter; text: "ЧАСЫ"; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeMicro; font.weight: AppTheme.weightBold; font.letterSpacing: 1.2 }
        Text { width: root.colNum; horizontalAlignment: Text.AlignHCenter; text: "СВЕРХ."; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeMicro; font.weight: AppTheme.weightBold; font.letterSpacing: 1.2 }
        Text { width: root.colNum; horizontalAlignment: Text.AlignHCenter; text: "ДНИ"; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeMicro; font.weight: AppTheme.weightBold; font.letterSpacing: 1.2 }
        Text { width: root.colBal; horizontalAlignment: Text.AlignHCenter; text: "ДОСТУПНО"; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeMicro; font.weight: AppTheme.weightBold; font.letterSpacing: 1.2 }
    }

    // ── Строки сотрудников ──
    Column {
        id: empTable
        width: parent.width
        spacing: 2
        visible: root.rows.length > 0

        // Для проверок в песочнице: геометрия отрисованных строк
        function layoutInfo() {
            let info = { rows: root.rows.length, heights: [], widths: [],
                         headerFieldX: -1, rowFieldX: -1, dlgW: root.width }
            for (let i = 0; i < children.length; i++) {
                let c = children[i]
                if (c && c.height !== undefined && c.width !== undefined) {
                    info.heights.push(c.height)
                    info.widths.push(c.width)
                }
            }
            if (children.length > 0) {
                let row = children[0]
                if (row.children && row.children.length > 0) {
                    let inner = row.children[0]
                    for (let k = 0; k < inner.children.length; k++) {
                        let ch = inner.children[k]
                        if (ch.width === root.colNum) { info.rowFieldX = ch.x; break }
                    }
                }
            }
            for (let k = 0; k < tableHeader.children.length; k++) {
                let ch = tableHeader.children[k]
                if (ch.width === root.colNum) { info.headerFieldX = ch.x + tableHeader.x; break }
            }
            return info
        }

        Repeater {
            model: root.rows

            Rectangle {
                id: rowRect
                width: parent.width
                required property var modelData
                required property int index
                readonly property var rowData: modelData
                readonly property string errText: root.rowErrors[rowData.id] || ""
                height: errText !== "" ? 86 : 66
                radius: AppTheme.radiusMedium
                color: errText !== "" ? AppTheme.bgDangerSoft : AppTheme.bgCell
                border.width: errText !== "" ? 1 : 0
                border.color: AppTheme.accentDanger
                clip: true
                Behavior on height { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeStandard } }
                Behavior on color { ColorAnimation { duration: AppTheme.durFast } }

                Row {
                    anchors.fill: parent
                    spacing: 8
                    leftPadding: 8
                    rightPadding: 8

                    AppCheckBox {
                        id: rowCheck
                        visible: !root.singleMode
                        checked: rowRect.rowData.checked
                        anchors.verticalCenter: parent.verticalCenter
                        width: 28
                        onToggled: root.toggleRow(index, checked)
                    }

                    Column {
                        width: root.nameWidth
                        anchors.verticalCenter: parent.verticalCenter
                        spacing: 0

                        Text {
                            width: parent.width
                            text: rowRect.rowData.name
                            color: rowRect.rowData.checked ? AppTheme.textPrimary : AppTheme.textDisabled
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: AppTheme.weightMedium
                            elide: Text.ElideRight
                        }
                        Text {
                            width: parent.width
                            visible: rowRect.rowData.subtitle !== ""
                            text: rowRect.rowData.subtitle
                            color: AppTheme.textTertiary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            elide: Text.ElideRight
                        }
                        Text {
                            width: parent.width
                            visible: rowRect.errText !== ""
                            text: rowRect.errText
                            color: AppTheme.accentDanger
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            font.weight: AppTheme.weightBold
                            wrapMode: Text.WordWrap
                        }
                    }

                    AppTextField {
                        width: root.colNum
                        anchors.verticalCenter: parent.verticalCenter
                        text: rowRect.rowData.hours
                        enabled: rowRect.rowData.checked
                        numericOnly: true
                        horizontalAlignment: TextInput.AlignHCenter
                        onEditingFinished: {
                            root.rows[index].hours = root._toInt(text)
                            root.recalcTotals()
                        }
                    }
                    AppTextField {
                        width: root.colNum
                        anchors.verticalCenter: parent.verticalCenter
                        text: rowRect.rowData.overtime
                        enabled: rowRect.rowData.checked
                        numericOnly: true
                        horizontalAlignment: TextInput.AlignHCenter
                        onEditingFinished: {
                            root.rows[index].overtime = root._toInt(text)
                            root.recalcTotals()
                        }
                    }
                    AppTextField {
                        width: root.colNum
                        anchors.verticalCenter: parent.verticalCenter
                        text: rowRect.rowData.days
                        enabled: rowRect.rowData.checked
                        numericOnly: true
                        horizontalAlignment: TextInput.AlignHCenter
                        onEditingFinished: {
                            root.rows[index].days = root._toInt(text)
                            root.recalcTotals()
                        }
                    }

                    Column {
                        width: root.colBal
                        anchors.verticalCenter: parent.verticalCenter
                        spacing: 0
                        Text {
                            width: parent.width
                            horizontalAlignment: Text.AlignHCenter
                            text: rowRect.rowData.balHours + " ч · " + rowRect.rowData.balDays + " д"
                            color: AppTheme.textSecondary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                        }
                        Text {
                            width: parent.width
                            horizontalAlignment: Text.AlignHCenter
                            text: "сверх. " + rowRect.rowData.balOvertime + " ч"
                            color: AppTheme.textTertiary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                        }
                    }
                }
            }
        }
    }

    Text {
        visible: root.rows.length === 0
        width: parent.width
        text: "В подразделении нет активных сотрудников на этот месяц"
        color: AppTheme.textTertiary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        horizontalAlignment: Text.AlignHCenter
        topPadding: AppTheme.spaceL
    }

    // ── Итог ──
    Text {
        width: parent.width
        text: "Выплата: " + root.totalEmp + " сотр. · всего " +
              root.totalHours + " ч · " + root.totalDays + " д"
        color: AppTheme.textSecondary
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        font.weight: AppTheme.weightBold
        horizontalAlignment: Text.AlignRight
    }

    // ── Ошибка (кто именно и что не хватает) ──
    Text {
        visible: root.topError !== ""
        width: parent.width
        text: root.topError
        color: AppTheme.accentDanger
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeBody
        font.weight: AppTheme.weightBold
        wrapMode: Text.WordWrap
    }
}
