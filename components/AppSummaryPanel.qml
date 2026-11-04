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
    // Все строки лежат в ОДНОЙ общей сетке, поэтому колонки имеют
    // общие границы и цифры не могут «съехать» от строки к строке.
    // Прошлый год не прячется в скобку: под остатком мелко стоит
    // «2025 год: …», и только когда там не ноль.
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

    // Колонки цифр: обычные ширины — 110/130/170/190. Если окно
    // узкое и на всё места не хватает, колонки и шрифт просто
    // пропорционально ужимаются: таблица остаётся той же таблицей
    // при любой ширине, ничего за край не вылезает.
    readonly property int titleMinW: 110
    readonly property int colGap: width >= 700 ? AppTheme.spaceM
            : (width >= 540 ? AppTheme.spaceS : AppTheme.spaceXS)
    readonly property real colScale: width > 0
            ? Math.min(1, Math.max(0, width - AppTheme.spaceM * 2
                    - colGap * 4 - titleMinW) / 600)
            : 1
    readonly property int colStart: Math.floor(110 * colScale)
    readonly property int colWorked: Math.floor(130 * colScale)
    readonly property int colComp: Math.floor(170 * colScale)
    readonly property int colEnd: Math.floor(190 * colScale)

    readonly property int titleFont: colScale >= 0.8 ? AppTheme.sizeBody
            : (colScale >= 0.6 ? 13 : 12)
    readonly property int flowFont: colScale >= 0.85 ? AppTheme.sizeBody
            : (colScale >= 0.68 ? 13 : 12)
    readonly property int endFont: colScale >= 0.85 ? AppTheme.sizeH4
            : (colScale >= 0.68 ? 21 : (colScale >= 0.55 ? 18 : 16))
    readonly property int headFont: colScale >= 0.68 ? AppTheme.sizeMicro
            : (colScale >= 0.55 ? 9 : 8)

    implicitHeight: backend && backend.selectedEmployeeId !== 0 ? mainRect.height : 0

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
        Layout.minimumWidth: head.colWidth
        Layout.maximumWidth: head.colWidth
        spacing: 0

        Text {
            Layout.fillWidth: true
            horizontalAlignment: Text.AlignRight
            text: head.title
            color: head.strong ? AppTheme.textSecondary : AppTheme.textTertiary
            font.family: AppTheme.fontFamily
            font.pixelSize: root.headFont
            font.weight: AppTheme.weightBold
            font.letterSpacing: root.colScale >= 0.9 ? 0.6 : 0
            elide: Text.ElideRight
        }
        Text {
            Layout.fillWidth: true
            horizontalAlignment: Text.AlignRight
            text: head.subtitle
            color: AppTheme.textTertiary
            opacity: 0.75
            font.family: AppTheme.fontFamily
            font.pixelSize: root.headFont
            elide: Text.ElideRight
        }
    }

    // Обычная цифра потока: ширина — по своей колонке, выравнивание вправо
    component FlowText: Text {
        id: flow
        property int colWidth: 100

        Layout.preferredWidth: flow.colWidth
        Layout.minimumWidth: flow.colWidth
        Layout.maximumWidth: flow.colWidth
        Layout.fillHeight: true
        horizontalAlignment: Text.AlignRight
        verticalAlignment: Text.AlignVCenter
        font.family: AppTheme.fontFamily
        font.pixelSize: root.flowFont
        font.features: {"tnum": 1}
        elide: Text.ElideRight
    }

    // Правая колонка ответа: остаток этого года и, если есть, прошлого
    component EndCell: ColumnLayout {
        id: endCell
        property string bucket: "hours"
        property int endValue: 0
        property int prevValue: 0

        Layout.preferredWidth: root.colEnd
        Layout.minimumWidth: root.colEnd
        Layout.maximumWidth: root.colEnd
        Layout.alignment: Qt.AlignVCenter
        spacing: 0

        Text {
            Layout.fillWidth: true
            horizontalAlignment: Text.AlignRight
            text: root.fmtBucket(endCell.bucket, endCell.endValue)
            color: endCell.endValue < 0 ? AppTheme.textDanger : AppTheme.textPrimary
            font.family: AppTheme.fontFamily
            font.pixelSize: root.endFont
            font.features: {"tnum": 1}
            font.weight: AppTheme.weightBold
            elide: Text.ElideRight
        }
        Text {
            visible: endCell.prevValue !== 0
            Layout.fillWidth: true
            horizontalAlignment: Text.AlignRight
            text: root.prevYearLabel + ": "
                    + root.fmtBucket(endCell.bucket, endCell.prevValue)
            color: endCell.prevValue < 0 ? AppTheme.textDanger : AppTheme.textTertiary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeMicro
            font.features: {"tnum": 1}
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
            // ОДНА ОБЩАЯ СЕТКА: заголовки, три параметра, «Всего».
            // Границы колонок общие на все строки — съехать нечему.
            // ------------------------------------------
            GridLayout {
                Layout.fillWidth: true
                Layout.topMargin: AppTheme.spaceXXS
                columns: 5
                columnSpacing: root.colGap
                rowSpacing: AppTheme.spaceS

                // Заголовки колонок
                Item {
                    Layout.fillWidth: true
                    Layout.minimumWidth: root.titleMinW
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

                // Время отдыха за ночные часы
                Text {
                    Layout.fillWidth: true
                    Layout.minimumWidth: root.titleMinW
                    Layout.fillHeight: true
                    text: "Время отдыха за ночные часы"
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: root.titleFont
                    font.weight: AppTheme.weightBold
                    verticalAlignment: Text.AlignVCenter
                    wrapMode: Text.WordWrap
                }
                FlowText {
                    colWidth: root.colStart
                    text: root.fmtBucket("hours", root.num("hours_start_value"))
                    color: root.flowColor(root.num("hours_start_value"))
                }
                FlowText {
                    colWidth: root.colWorked
                    text: root.workedText("hours", "hours_acc_value")
                    color: root.flowColor(root.num("hours_acc_value"))
                }
                FlowText {
                    colWidth: root.colComp
                    text: root.compText("hours", "hours_comp_value")
                    color: root.flowColor(root.num("hours_comp_value"))
                }
                EndCell {
                    bucket: "hours"
                    endValue: root.num("hours_end_value")
                    prevValue: root.num("hours_prev_end_value")
                }

                // Дополнительные дни отдыха
                Text {
                    Layout.fillWidth: true
                    Layout.minimumWidth: root.titleMinW
                    Layout.fillHeight: true
                    text: "Дополнительные дни отдыха"
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: root.titleFont
                    font.weight: AppTheme.weightBold
                    verticalAlignment: Text.AlignVCenter
                    wrapMode: Text.WordWrap
                }
                FlowText {
                    colWidth: root.colStart
                    text: root.fmtBucket("days", root.num("days_start_value"))
                    color: root.flowColor(root.num("days_start_value"))
                }
                FlowText {
                    colWidth: root.colWorked
                    text: root.workedText("days", "days_acc_value")
                    color: root.flowColor(root.num("days_acc_value"))
                }
                FlowText {
                    colWidth: root.colComp
                    text: root.compText("days", "days_comp_value")
                    color: root.flowColor(root.num("days_comp_value"))
                }
                EndCell {
                    bucket: "days"
                    endValue: root.num("days_end_value")
                    prevValue: root.num("days_prev_end_value")
                }

                // Сверхурочные часы
                Text {
                    Layout.fillWidth: true
                    Layout.minimumWidth: root.titleMinW
                    Layout.fillHeight: true
                    text: "Сверхурочные часы"
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: root.titleFont
                    font.weight: AppTheme.weightBold
                    verticalAlignment: Text.AlignVCenter
                    wrapMode: Text.WordWrap
                }
                FlowText {
                    colWidth: root.colStart
                    text: root.fmtBucket("overtime", root.num("overtime_start_value"))
                    color: root.flowColor(root.num("overtime_start_value"))
                }
                FlowText {
                    colWidth: root.colWorked
                    text: root.workedText("overtime", "overtime_acc_value")
                    color: root.flowColor(root.num("overtime_acc_value"))
                }
                FlowText {
                    colWidth: root.colComp
                    text: root.compText("overtime", "overtime_comp_value")
                    color: root.flowColor(root.num("overtime_comp_value"))
                }
                EndCell {
                    bucket: "overtime"
                    endValue: root.num("overtime_end_value")
                    prevValue: root.num("overtime_prev_end_value")
                }

                // Всего — итог правой колонки
                ColumnLayout {
                    Layout.fillWidth: true
                    Layout.minimumWidth: root.titleMinW
                    Layout.alignment: Qt.AlignVCenter
                    spacing: 0

                    Text {
                        text: "Всего"
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: root.titleFont
                        font.weight: AppTheme.weightBold
                    }
                    Text {
                        text: "все остатки в днях"
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeMicro
                    }
                }
                Item {
                    Layout.preferredWidth: root.colStart
                    Layout.minimumWidth: root.colStart
                    Layout.maximumWidth: root.colStart
                }
                Item {
                    Layout.preferredWidth: root.colWorked
                    Layout.minimumWidth: root.colWorked
                    Layout.maximumWidth: root.colWorked
                }
                Item {
                    Layout.preferredWidth: root.colComp
                    Layout.minimumWidth: root.colComp
                    Layout.maximumWidth: root.colComp
                }
                Text {
                    Layout.preferredWidth: root.colEnd
                    Layout.minimumWidth: root.colEnd
                    Layout.maximumWidth: root.colEnd
                    Layout.fillHeight: true
                    horizontalAlignment: Text.AlignRight
                    verticalAlignment: Text.AlignVCenter
                    text: root.totalDays + " " + root.daysWord(root.totalDays)
                    color: root.totalDays < 0 ? AppTheme.textDanger : AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: root.endFont
                    font.features: {"tnum": 1}
                    font.weight: AppTheme.weightBold
                    elide: Text.ElideRight
                }
            }

            // ------------------------------------------
            // НИЗ: только справочное мелким шрифтом
            // ------------------------------------------
            RowLayout {
                Layout.fillWidth: true
                Layout.topMargin: AppTheme.spaceXXS
                spacing: AppTheme.spaceM

                Text {
                    text: root.secondaryText()
                    visible: text.length > 0
                    Layout.fillWidth: true
                    color: AppTheme.textTertiary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                    wrapMode: Text.WordWrap
                }
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
            }
        }
    }
}
