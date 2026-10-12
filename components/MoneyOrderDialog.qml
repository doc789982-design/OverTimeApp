import QtQuick
import QtQuick.Controls
import "."

// ============================================================
// ВЕДОМОСТЬ: один денежный приказ — всему подразделению сразу.
//
// Слева, сверху вниз: дата приказа, номер, комментарий.
// Справа от полей: выбор года компенсации (текущий /
// предыдущий / оба), «Компенсировать всё» и «Одинаково всем».
// журнал приказов открывается кнопкой «Деньгами»
// сотрудника — из самой ведомости его больше нет.
// Ниже — список сотрудников: галочка, имя и три суммы
// (часы, дни, сверхурочные). При ошибке окно трясётся (штатная
// тряска AppDialog), поля с ошибкой вспыхивают красной рамкой,
// а строка виновника подсвечивается с пояснением.
//
// Режимы:
//   всем      — «₽» в левой панели или «Новый приказ» в журнале;
//   одному    — ведомость на одного (без галочек);
//   правка    — «Редактировать» из журнала: записи приказа
//               ЗАМЕНЯЮТСЯ, а не прибавляются.
// ============================================================
AppDialog {
    id: root
    width: 640

    property bool singleMode: false
    property bool editMode: false
    property string origOrderNo: ""
    property string origOrderDate: ""

    // За какой год компенсация: 0 — текущий, 1 — предыдущий
    // (заначка), 2 — оба года
    property int sourceMode: 0

    title: root.editMode ? "Редактирование приказа"
           : root.singleMode && root.rows.length === 1
             ? "Денежная компенсация — " + root.shortName(root.rows[0].name)
             : "Приказ о денежной компенсации"
    acceptText: root.editMode ? "Сохранить приказ" : "Провести приказ"
    acceptVariant: "primary"
    rejectText: "Закрыть"

    // [{id, name, subtitle, checked, hours, overtime, days,
    //    balHours, balOvertime, balDays, prevHours, prevOvertime, prevDays}]
    property var rows: []
    property var rowErrors: ({})     // id сотрудника → текст ошибки
    property string topError: ""
    property bool showEqual: false   // панель «одинаково всем»

    // ширины колонок таблицы (одинаково в шапке и строках)
    readonly property int colNum: 56
    // реквизиты — на четверть короче, чтобы правой колонке
    // (журнал, годы, «Компенсировать всё») было вольготно
    readonly property int headLeft: 225
    readonly property int headRight: width - AppTheme.spaceL * 2
                                     - headLeft - AppTheme.spaceL
    // имя сотрудника — от РЕАЛЬНОЙ ширины таблицы (не от ширины
    // окна: у попапа свои поля, из-за этого поля сумм выезжали
    // за правый край карточки на 12px)
    readonly property int nameWidth: empTable.width - 28
                                     - colNum * 3 - 8 * 4 - 16

    function _fromBackend() {
        let list = backend.moneyOrderEmployees()
        let res = []
        for (let i = 0; i < list.length; i++) {
            let e = list[i]
            res.push({
                id: e.id, name: e.name, subtitle: e.subtitle,
                checked: true, hours: 0, overtime: 0, days: 0,
                balHours: e.hours, balOvertime: e.overtime, balDays: e.days,
                prevHours: e.prevHours, prevOvertime: e.prevOvertime,
                prevDays: e.prevDays,
            })
        }
        return res
    }

    // Остатки строки по выбранному году
    function avail(row) {
        if (root.sourceMode === 1)
            return { h: row.prevHours, o: row.prevOvertime, d: row.prevDays }
        if (root.sourceMode === 2)
            return { h: row.balHours + row.prevHours,
                     o: row.balOvertime + row.prevOvertime,
                     d: row.balDays + row.prevDays }
        return { h: row.balHours, o: row.balOvertime, d: row.balDays }
    }

    function _resetHead() {
        root.rowErrors = ({})
        root.topError = ""
        root.showEqual = false
        root.sourceMode = 0
        orderNoInput.text = ""
        orderNoInput.hasError = false
        orderCommentInput.text = ""
        orderDateInput.selectedDate = new Date().toISOString().split("T")[0]
        orderDateInput.hasError = false
    }

    function openNew() {
        root.singleMode = false
        root.editMode = false
        root.rows = _fromBackend()
        _resetHead()
        root.showCentered()
    }

    // Приказ только для одного сотрудника — из его карточки
    function openForEmployee(empId) {
        let all = _fromBackend()
        let only = []
        for (let i = 0; i < all.length; i++)
            if (all[i].id === empId) { only.push(all[i]); break }
        root.singleMode = true
        root.editMode = false
        root.rows = only
        _resetHead()
        root.showCentered()
    }

    // Редактирование приказа из журнала: суммы как в приказе,
    // сохранение ЗАМЕНИТ записи этого приказа
    function editFromOrder(order) {
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
        root.singleMode = false
        root.editMode = true
        root.origOrderNo = order.order_no
        root.origOrderDate = order.order_date
        root.rows = current
        root.rowErrors = ({})
        root.topError = ""
        root.showEqual = false
        root.sourceMode = order.source_mode !== undefined ? order.source_mode : 0
        orderNoInput.text = order.order_no
        orderNoInput.hasError = false
        orderCommentInput.text = order.comment
        orderDateInput.selectedDate = order.order_date
        orderDateInput.hasError = false
        root.showCentered()
    }

    // «Иванов Иван Андреевич» → «Иванов И.А.»
    function shortName(fio) {
        let p = String(fio || "").trim().split(/\s+/)
        let out = p[0] || ""
        for (let i = 1; i < p.length; i++)
            if (p[i].length > 0) out += " " + p[i][0] + "."
        return out
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
    }

    function _toInt(s) {
        let v = parseInt(String(s).replace(/[^\d]/g, ""), 10)
        return isNaN(v) ? 0 : Math.max(0, v)
    }

    // Для проверок в песочнице: геометрия отрисованных строк
    function layoutInfo() { return empTable.layoutInfo() }

    onAccepted: {
        // номер и дата приказа обязательны: вспышка красной рамки
        let noBad = orderNoInput.text.trim() === ""
        let dateBad = !(orderDateInput.selectedDate !== ""
                        && /^\d{2}\.\d{2}\.\d{4}$/.test(orderDateInput.text.trim()))
        if (noBad) orderNoInput.flashError()
        if (dateBad) orderDateInput.flashError()
        if (noBad || dateBad) {
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
                    ? "Укажите, сколько выплатить: часы, дни или сверхурочные"
                    : "Отметьте сотрудников галочками и укажите суммы"
            root.shake()
            return
        }

        let result
        if (root.editMode) {
            result = backend.updateMoneyOrder(
                root.origOrderNo, root.origOrderDate,
                JSON.stringify(payload),
                orderNoInput.text.trim(),
                orderDateInput.selectedDate,
                orderCommentInput.text.trim(),
                root.sourceMode)
        } else {
            result = backend.saveMoneyOrder(
                JSON.stringify(payload),
                orderNoInput.text.trim(),
                orderDateInput.selectedDate,
                orderCommentInput.text.trim(),
                root.sourceMode)
        }
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

    // ── Шапка: слева реквизиты приказа (сверху вниз),
    //    справа — журнал, год компенсации, «Компенсировать всё» ──
    Row {
        width: parent.width
        spacing: AppTheme.spaceL

        Column {
            width: root.headLeft
            spacing: AppTheme.spaceS

            AppDateField {
                id: orderDateInput
                width: parent.width
                label: "Дата приказа"
            }
            AppTextField {
                id: orderNoInput
                width: parent.width
                label: "Номер приказа"
                placeholderText: "123 л/с"
                maximumLength: 20
            }
            AppTextField {
                id: orderCommentInput
                width: parent.width
                label: "Комментарий"
                placeholderText: "Например: за октябрь"
            }
        }

        Column {
            width: root.headRight
            spacing: AppTheme.spaceXS

            // За какой год компенсация
            AppRadioButton {
                text: "Текущий год"
                checked: root.sourceMode === 0
                onToggled: root.sourceMode = 0
            }
            AppRadioButton {
                text: "Предыдущий год"
                checked: root.sourceMode === 1
                onToggled: root.sourceMode = 1
            }
            AppRadioButton {
                text: "Оба года"
                checked: root.sourceMode === 2
                onToggled: root.sourceMode = 2
            }

            AppButton {
                id: btnFillAll
                width: parent.width
                text: "Компенсировать всё"
                iconSource: "../icons/edit.svg"
                variant: "secondary"
                height: 36
                onClicked: {
                    for (let i = 0; i < root.rows.length; i++) {
                        if (!root.rows[i].checked) continue
                        let a = root.avail(root.rows[i])
                        root.rows[i].hours = a.h
                        root.rows[i].overtime = a.o
                        root.rows[i].days = a.d
                    }
                    root.rows = root.rows.slice()
                }
                AppToolTip {
                    anchors.horizontalCenter: parent.horizontalCenter
                    anchors.top: parent.bottom; anchors.topMargin: AppTheme.spaceXXS
                    text: "Подставить каждому выбранному весь его остаток за выбранный год"
                    isVisible: parent.hovered
                }
            }

            AppButton {
                id: btnEqual
                visible: !root.singleMode
                width: visible ? parent.width : implicitWidth
                text: root.showEqual ? "Скрыть" : "Одинаково всем…"
                iconSource: "../icons/money.svg"
                variant: "secondary"
                height: 36
                onClicked: root.showEqual = !root.showEqual
                AppToolTip {
                    anchors.horizontalCenter: parent.horizontalCenter
                    anchors.top: parent.bottom; anchors.topMargin: AppTheme.spaceXXS
                    text: "Одна и та же сумма каждому выбранному"
                    isVisible: parent.hovered
                }
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
        Text { width: root.colNum; horizontalAlignment: Text.AlignHCenter; text: "ДНИ"; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeMicro; font.weight: AppTheme.weightBold; font.letterSpacing: 1.2 }
        Text { width: root.colNum; horizontalAlignment: Text.AlignHCenter; text: "СВЕРХ."; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeMicro; font.weight: AppTheme.weightBold; font.letterSpacing: 1.2 }
    }

    // ── Строки сотрудников ──
    Column {
        id: empTable
        width: parent.width
        spacing: 6
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
                            root.rows = root.rows.slice()
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
                            root.rows = root.rows.slice()
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
                            root.rows = root.rows.slice()
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
