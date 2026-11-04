import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import Qt5Compat.GraphicalEffects
import QtQuick.Controls.impl

Item {
    id: root
    property bool isYearView: false

    // ════════════════════════════════════════════════════════════════
    // ПАНЕЛЬ «ИТОГИ»: ТРИ ПАРАМЕТРА — ТРИ СТРОКИ
    //
    // Панель читается как короткий расчёт слева направо:
    // было → отработано → компенсировано → осталось.
    // Подписи колонок написаны один раз сверху, главная цифра —
    // всегда в правой колонке «Осталось». Прошлый год не прячется
    // в скобку: под остатком мелко стоит «2025 год: …», и только
    // когда там не ноль. Подробности — по нажатию на строку.
    // ════════════════════════════════════════════════════════════════
    readonly property var summ: backend
            ? ((root.isYearView ? backend.yearSummary : backend.monthSummary) || ({}))
            : ({})
    readonly property bool isShiftMonth: !root.isYearView && root.summ.is_shift === true
    readonly property string periodText: {
        if (!backend || !backend.currentPeriodText) return ""
        if (!root.isYearView) return backend.currentPeriodText
        let parts = String(backend.currentPeriodText).split(" ")
        return parts.length > 1 ? parts[1] + " год" : backend.currentPeriodText
    }
    readonly property string startWhen: root.summ.period_start_label
            || (root.isYearView ? "на 1 января" : "на начало месяца")
    readonly property string flowWhen: root.summ.period_flow_label
            || (root.isYearView ? "за год" : "за месяц")
    readonly property string endWhen: root.summ.period_end_label
            || (root.isYearView ? "на 31 декабря" : "на конец месяца")
    readonly property string prevYearLabel: root.summ.prev_year_label || "прошлый год"
    readonly property int totalDays: root.num("total_days")

    readonly property int colStart: 110
    readonly property int colWorked: 130
    readonly property int colComp: 170
    readonly property int colEnd: 190

    implicitHeight: backend && backend.selectedEmployeeId !== 0 ? mainRect.height : 0

    // Клик по строке или по «Всего» открывает расшифровку этой цифры
    function openExplain(bucket) {
        balanceExplainDialog.openExplain(bucket, root.isYearView ? "year" : "month")
    }

    function num(key) {
        let v = root.summ ? root.summ[key] : undefined
        let n = Number(v)
        return isNaN(n) ? 0 : n
    }

    // Минуты — как в остальной программе: «8 ч.», «8 ч. 30 м.»
    function fmtBucket(bucket, value) {
        let v = Math.trunc(Number(value) || 0)
        if (bucket === "days")
            return v + " д."
        let sign = v < 0 ? "−" : ""
        let a = Math.abs(v)
        let h = Math.trunc(a / 60)
        let m = a % 60
        return sign + h + " ч." + (m > 0 ? " " + m + " м." : "")
    }

    function workedText(bucket, key) {
        let v = root.num(key)
        if (v === 0) return "—"
        return (v > 0 ? "+" : "") + root.fmtBucket(bucket, v)
    }

    function compText(bucket, key) {
        let v = root.num(key)
        if (v === 0) return "—"
        if (v > 0) return "−" + root.fmtBucket(bucket, v)
        return root.fmtBucket(bucket, v)
    }

    function flowColor(value) {
        if (value < 0) return AppTheme.textDanger
        if (value === 0) return AppTheme.textTertiary
        return AppTheme.textSecondary
    }

    // Склонение: 1 день / 2 дня / 5 дней
    function daysWord(n) {
        let a = Math.abs(Math.trunc(Number(n) || 0))
        if (a % 100 >= 11 && a % 100 <= 19) return "дней"
        let d = a % 10
        if (d === 1) return "день"
        if (d >= 2 && d <= 4) return "дня"
        return "дней"
    }

    function secondaryText() {
        if (root.isYearView) {
            return "Больничный " + (root.summ.b_days || "—")
                    + " · Отпуск " + (root.summ.o_days || "—")
                    + " · Командировка " + (root.summ.k_days || "—")
        }
        if (root.isShiftMonth) {
            return "По графику: ночью " + (root.summ.shift_night || "—")
                    + " · в праздники " + (root.summ.shift_holiday || "—")
                    + " · норма месяца " + (root.summ.norm_minutes || "0")
        }
        return ""
    }

    // ==========================================
    // ЗАГОЛОВОК КОЛОНКИ
    // ==========================================
    component ColumnHead: ColumnLayout {
        id: head
        property string title: ""
        property string subtitle: ""
        property int colWidth: 100
        property bool strong: false

        Layout.preferredWidth: head.colWidth
        spacing: 0

        Text {
            Layout.fillWidth: true
            horizontalAlignment: Text.AlignRight
            text: head.title
            color: head.strong ? AppTheme.textSecondary : AppTheme.textTertiary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeMicro
            font.weight: AppTheme.weightBold
            font.letterSpacing: 0.6
            elide: Text.ElideRight
        }
        Text {
            Layout.fillWidth: true
            horizontalAlignment: Text.AlignRight
            text: head.subtitle
            color: AppTheme.textTertiary
            opacity: 0.75
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeMicro
            elide: Text.ElideRight
        }
    }

    // ==========================================
    // ГЛАВНАЯ ПАНЕЛЬ
    // ==========================================
    Rectangle {
        id: mainRect
        visible: backend && backend.selectedEmployeeId !== 0
        anchors.top: parent.top
        anchors.left: parent.left
        anchors.right: parent.right
        height: summaryBody.implicitHeight + AppTheme.spaceM * 2
        radius: AppTheme.radiusLarge
        color: AppTheme.bgSurface
        border.color: AppTheme.borderDivider
        border.width: 1

        AppShadow { level: 1 }

        ColumnLayout {
            id: summaryBody
            anchors.top: parent.top
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.margins: AppTheme.spaceM
            spacing: AppTheme.spaceXS

            // ------------------------------------------
            // ШАПКА: заголовок, период и «Деньгами»
            // ------------------------------------------
            RowLayout {
                Layout.fillWidth: true
                spacing: AppTheme.spaceM

                IconImage {
                    Layout.alignment: Qt.AlignVCenter
                    source: "../icons/overtime.svg"
                    width: 18; height: 18
                    color: AppTheme.textTertiary
                }
                ColumnLayout {
                    spacing: 0
                    Text {
                        text: "Итоги"
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBodyLarge
                        font.weight: AppTheme.weightBold
                    }
                    Text {
                        text: root.periodText
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                    }
                }

                Item { Layout.fillWidth: true }

                // «Деньгами» — кнопка-стадион, как все кнопки программы.
                // Открывает журнал приказов о денежной компенсации.
                AppButton {
                    visible: !root.isYearView
                    objectName: "moneyOrderBtn"
                    text: "Деньгами"
                    iconSource: "../icons/ruble.svg"
                    variant: "secondary"
                    height: 32
                    Layout.alignment: Qt.AlignVCenter
                    onClicked: { backend.loadMoneyOrders(); moneyInspector.show() }
                    AppToolTip {
                        anchors.horizontalCenter: parent.horizontalCenter
                        anchors.bottom: parent.top
                        anchors.bottomMargin: AppTheme.spaceXXS
                        isVisible: parent.hovered
                        text: "Приказы о денежной компенсации"
                    }
                }
            }

            // ------------------------------------------
            // ЗАГОЛОВКИ КОЛОНОК: один раз на всю таблицу
            // ------------------------------------------
            RowLayout {
                Layout.fillWidth: true
                Layout.topMargin: AppTheme.spaceXXS
                spacing: AppTheme.spaceM

                Item {
                    Layout.fillWidth: true
                    Layout.minimumWidth: 150
                }
                ColumnHead {
                    colWidth: root.colStart
                    title: "БЫЛО"
                    subtitle: root.startWhen
                }
                ColumnHead {
                    colWidth: root.colWorked
                    title: "ОТРАБОТАНО"
                    subtitle: root.flowWhen
                }
                ColumnHead {
                    colWidth: root.colComp
                    title: "КОМПЕНСИРОВАНО"
                    subtitle: root.flowWhen
                }
                ColumnHead {
                    colWidth: root.colEnd
                    title: "ОСТАЛОСЬ"
                    subtitle: root.endWhen
                    strong: true
                }
            }

            Rectangle {
                Layout.fillWidth: true
                height: 1
                color: AppTheme.borderDivider
            }

            // ------------------------------------------
            // ТРИ СТРОКИ-ПАРАМЕТРА
            // ------------------------------------------
            Repeater {
                model: [
                    { "title": "Время отдыха за ночные часы", "bucket": "hours",
                      "startKey": "hours_start_value", "accKey": "hours_acc_value",
                      "compKey": "hours_comp_value", "endKey": "hours_end_value",
                      "prevKey": "hours_prev_end_value" },
                    { "title": "Дополнительные дни отдыха", "bucket": "days",
                      "startKey": "days_start_value", "accKey": "days_acc_value",
                      "compKey": "days_comp_value", "endKey": "days_end_value",
                      "prevKey": "days_prev_end_value" },
                    { "title": "Сверхурочные часы", "bucket": "overtime",
                      "startKey": "overtime_start_value", "accKey": "overtime_acc_value",
                      "compKey": "overtime_comp_value", "endKey": "overtime_end_value",
                      "prevKey": "overtime_prev_end_value" }
                ]

                delegate: Item {
                    id: rowItem
                    required property var modelData

                    property string bucket: modelData.bucket
                    property int startValue: root.num(modelData.startKey)
                    property int accValue: root.num(modelData.accKey)
                    property int compValue: root.num(modelData.compKey)
                    property int endValue: root.num(modelData.endKey)
                    property int prevValue: root.num(modelData.prevKey)

                    Layout.fillWidth: true
                    Layout.preferredHeight: prevValue !== 0 ? 52 : 42

                    Rectangle {
                        anchors.fill: parent
                        radius: AppTheme.radiusSmall
                        color: rowHover.hovered ? AppTheme.stateHover : "transparent"
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                    }
                    HoverHandler {
                        id: rowHover
                        cursorShape: Qt.PointingHandCursor
                    }
                    TapHandler {
                        onTapped: root.openExplain(rowItem.bucket)
                    }

                    RowLayout {
                        anchors.fill: parent
                        spacing: AppTheme.spaceM

                        Text {
                            Layout.fillWidth: true
                            Layout.minimumWidth: 150
                            Layout.alignment: Qt.AlignVCenter
                            text: modelData.title
                            color: rowHover.hovered ? AppTheme.textOnSoft : AppTheme.textPrimary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: AppTheme.weightBold
                            wrapMode: Text.WordWrap
                            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        }
                        Text {
                            Layout.preferredWidth: root.colStart
                            Layout.alignment: Qt.AlignVCenter
                            horizontalAlignment: Text.AlignRight
                            text: root.fmtBucket(rowItem.bucket, rowItem.startValue)
                            color: root.flowColor(rowItem.startValue)
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.features: {"tnum": 1}
                            elide: Text.ElideRight
                        }
                        Text {
                            Layout.preferredWidth: root.colWorked
                            Layout.alignment: Qt.AlignVCenter
                            horizontalAlignment: Text.AlignRight
                            text: root.workedText(rowItem.bucket, modelData.accKey)
                            color: root.flowColor(rowItem.accValue)
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.features: {"tnum": 1}
                            elide: Text.ElideRight
                        }
                        Text {
                            Layout.preferredWidth: root.colComp
                            Layout.alignment: Qt.AlignVCenter
                            horizontalAlignment: Text.AlignRight
                            text: root.compText(rowItem.bucket, modelData.compKey)
                            color: root.flowColor(rowItem.compValue)
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.features: {"tnum": 1}
                            elide: Text.ElideRight
                        }
                        ColumnLayout {
                            Layout.preferredWidth: root.colEnd
                            Layout.alignment: Qt.AlignVCenter
                            spacing: 0

                            Text {
                                Layout.fillWidth: true
                                horizontalAlignment: Text.AlignRight
                                text: root.fmtBucket(rowItem.bucket, rowItem.endValue)
                                color: rowItem.endValue < 0
                                        ? AppTheme.textDanger
                                        : (rowHover.hovered ? AppTheme.textOnSoft : AppTheme.textPrimary)
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeH4
                                font.features: {"tnum": 1}
                                font.weight: AppTheme.weightBold
                                elide: Text.ElideRight
                                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                            }
                            Text {
                                visible: rowItem.prevValue !== 0
                                Layout.fillWidth: true
                                horizontalAlignment: Text.AlignRight
                                text: root.prevYearLabel + ": "
                                        + root.fmtBucket(rowItem.bucket, rowItem.prevValue)
                                color: rowItem.prevValue < 0
                                        ? AppTheme.textDanger : AppTheme.textTertiary
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeMicro
                                font.features: {"tnum": 1}
                                elide: Text.ElideRight
                            }
                        }
                    }

                    Rectangle {
                        anchors.left: parent.left
                        anchors.right: parent.right
                        anchors.bottom: parent.bottom
                        height: 1
                        color: AppTheme.borderDivider
                        opacity: 0.65
                    }
                }
            }

            // ------------------------------------------
            // ВСЕГО: итог правой колонки одной строкой
            // ------------------------------------------
            Item {
                Layout.fillWidth: true
                Layout.preferredHeight: 46

                Rectangle {
                    anchors.fill: parent
                    radius: AppTheme.radiusSmall
                    color: totalHover.hovered ? AppTheme.stateHover : "transparent"
                    Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                }
                HoverHandler {
                    id: totalHover
                    cursorShape: Qt.PointingHandCursor
                }
                TapHandler {
                    onTapped: root.openExplain("total")
                }

                RowLayout {
                    anchors.fill: parent
                    spacing: AppTheme.spaceM

                    ColumnLayout {
                        Layout.fillWidth: true
                        Layout.minimumWidth: 150
                        Layout.alignment: Qt.AlignVCenter
                        spacing: 0

                        Text {
                            text: "Всего"
                            color: totalHover.hovered ? AppTheme.textOnSoft : AppTheme.textPrimary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: AppTheme.weightBold
                            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                        }
                        Text {
                            text: "все остатки в днях"
                            color: AppTheme.textTertiary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeMicro
                        }
                    }
                    Item { Layout.preferredWidth: root.colStart }
                    Item { Layout.preferredWidth: root.colWorked }
                    Item { Layout.preferredWidth: root.colComp }
                    Text {
                        Layout.preferredWidth: root.colEnd
                        Layout.alignment: Qt.AlignVCenter
                        horizontalAlignment: Text.AlignRight
                        text: root.totalDays + " " + root.daysWord(root.totalDays)
                        color: root.totalDays < 0
                                ? AppTheme.textDanger
                                : (totalHover.hovered ? AppTheme.textOnSoft : AppTheme.textPrimary)
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeH4
                        font.features: {"tnum": 1}
                        font.weight: AppTheme.weightBold
                        elide: Text.ElideRight
                        Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                    }
                }
            }

            // ------------------------------------------
            // НИЗ: справочное мелким шрифтом + подсказка про нажатие
            // ------------------------------------------
            RowLayout {
                Layout.fillWidth: true
                Layout.topMargin: AppTheme.spaceXXS
                spacing: AppTheme.spaceM

                Text {
                    text: root.secondaryText()
                    visible: text.length > 0
                    color: AppTheme.textTertiary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                    elide: Text.ElideRight
                }
                Item { Layout.fillWidth: true }
                RowLayout {
                    visible: root.isYearView
                    spacing: AppTheme.spaceXXS

                    IconImage {
                        Layout.alignment: Qt.AlignVCenter
                        source: "../icons/ruble.svg"
                        width: 13; height: 13
                        color: AppTheme.textTertiary
                    }
                    Text {
                        Layout.alignment: Qt.AlignVCenter
                        Layout.maximumWidth: 300
                        text: "деньгами " + (root.summ.comp_money || "—")
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                        elide: Text.ElideRight
                    }
                }
                Text {
                    Layout.alignment: Qt.AlignVCenter
                    Layout.maximumWidth: 330
                    text: "Нажмите строку — покажем дежурства и компенсации"
                    color: AppTheme.textTertiary
                    opacity: 0.75
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeMicro
                    elide: Text.ElideRight
                }
            }
        }
    }
}
