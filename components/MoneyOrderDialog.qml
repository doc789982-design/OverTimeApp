import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import "."

// ============================================================
// ВЕДОМОСТЬ: один денежный приказ — всему подразделению сразу.
//
// Шапка приказа (номер, дата, комментарий) вводится один раз.
// Ниже — список сотрудников: у каждого галочка и СВОИ суммы
// (часы, сверхурочные, дни), которые можно править построчно.
// Заполнение: «всем одинаково» или «каждому по табелю».
// При ошибке окно трясётся, а строка сотрудника, у которого
// не хватает остатка, подсвечивается красным с пояснением.
// ============================================================
AppDialog {
    id: root
    width: 720
    heightFraction: 3 / 4

    title: "Денежная компенсация — приказ по подразделению"
    acceptText: "Провести приказ"
    acceptVariant: "primary"
    rejectText: "Закрыть"

    property var rows: []            // [{id, name, subtitle, checked, hours, overtime, days, balHours, balOvertime, balDays}]
    property var rowErrors: ({})     // id сотрудника → текст ошибки
    property string topError: ""
    property int totalEmp: 0
    property int totalHours: 0
    property int totalDays: 0

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

    function openNew() {
        root.rows = _fromBackend()
        root.rowErrors = ({})
        root.topError = ""
        orderNoInput.text = ""
        orderCommentInput.text = ""
        orderDateInput.selectedDate = new Date().toISOString().split("T")[0]
        recalcTotals()
        root.showCentered()
    }

    // «Повторить приказ»: суммы и получатели — как в прошлый раз
    function prefillFromOrder(order) {
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
        let v = parseInt(String(s).replace(/[^\d-]/g, ""), 10)
        return isNaN(v) ? 0 : Math.max(0, v)
    }

    onAccepted: {
        // Собираем строки и проводим приказ
        let payload = []
        for (let i = 0; i < rows.length; i++) {
            let r = rows[i]
            if (!r.checked) continue
            payload.push({ id: r.id, hours: r.hours, overtime: r.overtime, days: r.days })
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
        // Ошибка: окно трясётся, виновники подсвечиваются
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
        shakeAnim.restart()
    }

    // Тряска окна при ошибке
    SequentialAnimation {
        id: shakeAnim
        alwaysRunToEnd: true
        NumberAnimation { target: contentCol; property: "x"; to: -10; duration: 55 }
        NumberAnimation { target: contentCol; property: "x"; to: 10; duration: 55 }
        NumberAnimation { target: contentCol; property: "x"; to: -7; duration: 50 }
        NumberAnimation { target: contentCol; property: "x"; to: 7; duration: 50 }
        NumberAnimation { target: contentCol; property: "x"; to: 0; duration: 45 }
    }

    onOpened: root.topError = ""

    ColumnLayout {
        id: contentCol
        width: parent.width
        spacing: AppTheme.spaceS

        // ── Шапка приказа ──
        RowLayout {
            Layout.fillWidth: true
            spacing: AppTheme.spaceS

            AppTextField {
                Layout.preferredWidth: 110
                label: "№ приказа"
                placeholderText: "245"
                maximumLength: 20
                id: orderNoInput
            }
            AppDateField {
                Layout.preferredWidth: 160
                label: "Дата"
                id: orderDateInput
            }
            AppTextField {
                Layout.fillWidth: true
                label: "Комментарий"
                placeholderText: "За октябрь"
                id: orderCommentInput
            }
        }

        // ── Быстрое заполнение ──
        RowLayout {
            Layout.fillWidth: true
            spacing: AppTheme.spaceS

            Text {
                text: "Заполнить:"
                color: AppTheme.textSecondary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeBody
            }
            AppTextField {
                Layout.preferredWidth: 74
                label: ""
                placeholderText: "часы"
                id: fillHours
                horizontalAlignment: TextInput.AlignHCenter
            }
            AppTextField {
                Layout.preferredWidth: 74
                label: ""
                placeholderText: "сверх."
                id: fillOvertime
                horizontalAlignment: TextInput.AlignHCenter
            }
            AppTextField {
                Layout.preferredWidth: 74
                label: ""
                placeholderText: "дни"
                id: fillDays
                horizontalAlignment: TextInput.AlignHCenter
            }
            AppButton {
                text: "Всем"
                implicitHeight: 36
                onClicked: {
                    let h = root._toInt(fillHours.text)
                    let o = root._toInt(fillOvertime.text)
                    let d = root._toInt(fillDays.text)
                    for (let i = 0; i < root.rows.length; i++) {
                        if (!root.rows[i].checked) continue
                        root.rows[i].hours = h
                        root.rows[i].overtime = o
                        root.rows[i].days = d
                    }
                    root.rowsChanged()
                    root.recalcTotals()
                }
            }
            Item { Layout.fillWidth: true }
            AppButton {
                text: "Каждому по табелю"
                variant: "secondary"
                implicitHeight: 36
                onClicked: {
                    for (let i = 0; i < root.rows.length; i++) {
                        if (!root.rows[i].checked) continue
                        root.rows[i].hours = root.rows[i].balHours
                        root.rows[i].overtime = root.rows[i].balOvertime
                        root.rows[i].days = root.rows[i].balDays
                    }
                    root.rowsChanged()
                    root.recalcTotals()
                }
            }
        }

        // ── Заголовок таблицы ──
        RowLayout {
            Layout.fillWidth: true
            Layout.leftMargin: AppTheme.spaceS
            Layout.rightMargin: AppTheme.spaceS
            spacing: AppTheme.spaceS

            Text { text: ""; Layout.preferredWidth: 28 }
            Text {
                text: "Сотрудник"
                color: AppTheme.textTertiary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeMicro
                font.weight: AppTheme.weightBold
                Layout.fillWidth: true
            }
            Text { text: "Часы"; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeMicro; font.weight: AppTheme.weightBold; Layout.preferredWidth: 56 }
            Text { text: "Сверх."; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeMicro; font.weight: AppTheme.weightBold; Layout.preferredWidth: 56 }
            Text { text: "Дни"; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeMicro; font.weight: AppTheme.weightBold; Layout.preferredWidth: 56 }
            Text { text: "Остаток"; color: AppTheme.textTertiary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeMicro; font.weight: AppTheme.weightBold; Layout.preferredWidth: 86 }
        }

        // ── Строки сотрудников ──
        ListView {
            id: rowsList
            Layout.fillWidth: true
            Layout.fillHeight: true
            clip: true
            boundsBehavior: Flickable.StopAtBounds
            model: root.rows.length
            spacing: 2

            delegate: Rectangle {
                width: rowsList.width
                height: 62 + (root.rowErrors[root.rows[index].id] !== undefined ? 18 : 0)
                radius: AppTheme.radiusSmall
                color: root.rowErrors[root.rows[index].id] !== undefined
                       ? AppTheme.bgDangerSoft : "transparent"
                border.width: root.rowErrors[root.rows[index].id] !== undefined ? 1 : 0
                border.color: AppTheme.accentDanger
                Behavior on color { ColorAnimation { duration: AppTheme.durFast } }

                RowLayout {
                    anchors.fill: parent
                    anchors.leftMargin: AppTheme.spaceS
                    anchors.rightMargin: AppTheme.spaceS
                    spacing: AppTheme.spaceS

                    AppCheckBox {
                        checked: root.rows[index].checked
                        onClicked: {
                            root.rows[index].checked = checked
                            root.recalcTotals()
                        }
                    }

                    ColumnLayout {
                        Layout.fillWidth: true
                        spacing: 0
                        Text {
                            text: root.rows[index].name
                            color: root.rows[index].checked ? AppTheme.textPrimary : AppTheme.textDisabled
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: AppTheme.weightMedium
                            elide: Text.ElideRight
                            Layout.fillWidth: true
                        }
                        Text {
                            visible: root.rows[index].subtitle !== ""
                            text: root.rows[index].subtitle
                            color: AppTheme.textTertiary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            elide: Text.ElideRight
                            Layout.fillWidth: true
                        }
                        Text {
                            visible: root.rowErrors[root.rows[index].id] !== undefined
                            text: root.rowErrors[root.rows[index].id] || ""
                            color: AppTheme.accentDanger
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            wrapMode: Text.WordWrap
                            Layout.fillWidth: true
                        }
                    }

                    AppTextField {
                        Layout.preferredWidth: 56
                        text: root.rows[index].hours
                        enabled: root.rows[index].checked
                        horizontalAlignment: TextInput.AlignHCenter
                        onEditingFinished: {
                            root.rows[index].hours = root._toInt(text)
                            root.recalcTotals()
                        }
                    }
                    AppTextField {
                        Layout.preferredWidth: 56
                        text: root.rows[index].overtime
                        enabled: root.rows[index].checked
                        horizontalAlignment: TextInput.AlignHCenter
                        onEditingFinished: {
                            root.rows[index].overtime = root._toInt(text)
                            root.recalcTotals()
                        }
                    }
                    AppTextField {
                        Layout.preferredWidth: 56
                        text: root.rows[index].days
                        enabled: root.rows[index].checked
                        horizontalAlignment: TextInput.AlignHCenter
                        onEditingFinished: {
                            root.rows[index].days = root._toInt(text)
                            root.recalcTotals()
                        }
                    }

                    Text {
                        Layout.preferredWidth: 86
                        text: root.rows[index].balHours + " ч · " + root.rows[index].balDays + " д"
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                        horizontalAlignment: Text.AlignHCenter
                    }
                }
            }
        }

        // ── Ошибка сверху (кто именно и что не хватает) ──
        Text {
            visible: root.topError !== ""
            text: root.topError
            color: AppTheme.accentDanger
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody
            wrapMode: Text.WordWrap
            Layout.fillWidth: true
        }

        // ── Итог ──
        Text {
            text: "Выплата: " + root.totalEmp + " сотр. · всего " +
                  root.totalHours + " ч · " + root.totalDays + " д"
            color: AppTheme.textSecondary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeBody
            font.weight: AppTheme.weightBold
            Layout.fillWidth: true
            horizontalAlignment: Text.AlignRight
        }
    }
}
