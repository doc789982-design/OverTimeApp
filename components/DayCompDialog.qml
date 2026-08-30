import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

// ============================================================
// ОКНО ДОБАВЛЕНИЯ КОМПЕНСАЦИИ
//
// Отдельное окно (не вкладка): собственный компонент поверх общей
// базы AppDialog. Окно дежурства — это уже другой элемент,
// у каждого своя геометрия и свой подбор высоты, поэтому при
// морфинге из меню дня не «перескакивает» с размера другого окна.
// ============================================================
AppDialog {
    id: root
    width: 380
    heightFraction: 2/3

    property string targetDate: ""
    property int editCompId: 0

    // Дата живёт в заголовке окна (слева от крестика закрытия),
    // чтобы в содержимом осталось только само редактирование.
    title: formatBeautifulDate(targetDate)
    acceptText: "Сохранить"
    acceptVariant: "primary"

    onClosed: {
        mainWindow.activeSpotlightCell = null
    }

    function formatBeautifulDate(dateString) {
        if (!dateString) return ""
        let parts = dateString.split("-")
        if (parts.length !== 3) return dateString
        let year = parts[0]
        let month = parseInt(parts[1], 10)
        let day = parseInt(parts[2], 10)
        let months = ["января", "февраля", "марта", "апреля", "мая", "июня",
                      "июля", "августа", "сентября", "октября", "ноября", "декабря"]
        return day + " " + months[month - 1] + " " + year + " г."
    }

    // ==========================================
    // СОДЕРЖИМОЕ: КОМПЕНСАЦИЯ
    // ==========================================
    Column {
        id: compCol
        width: parent.width
        spacing: AppTheme.spaceM

        property int compMode: 0
        property bool isUpdating: false
        property real patternConfidence: -1
        property int patternCycle: 0
        property int patternWorkDays: 0
        property var shiftDates: []

        // ── Переключатель режима ──────────────────
        // Переключатель «Этот день / Период» — всегда доступен
        Row {
            width: parent.width
            spacing: AppTheme.spaceM
            visible: true

            Rectangle {
                width: 120; height: 36
                radius: AppTheme.radiusMedium
                color: compCol.compMode === 0
                    ? AppTheme.stateSelected
                    : (m1Hover.containsMouse ? AppTheme.stateHover : "transparent")
                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                Text {
                    anchors.centerIn: parent
                    text: "Этот день"
                    color: compCol.compMode === 0 ? AppTheme.textOnSoft : AppTheme.textSecondary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    font.weight: compCol.compMode === 0 ? AppTheme.weightBold : AppTheme.weightMedium
                }
                MouseArea {
                    id: m1Hover
                    anchors.fill: parent
                    hoverEnabled: true
                    cursorShape: Qt.PointingHandCursor
                    onClicked: compCol.compMode = 0
                }
            }

            Rectangle {
                width: 120; height: 36
                radius: AppTheme.radiusMedium
                color: compCol.compMode === 1
                    ? AppTheme.stateSelected
                    : (m2Hover.containsMouse ? AppTheme.stateHover : "transparent")
                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                Text {
                    anchors.centerIn: parent
                    text: "Период"
                    color: compCol.compMode === 1 ? AppTheme.textOnSoft : AppTheme.textSecondary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeBody
                    font.weight: compCol.compMode === 1 ? AppTheme.weightBold : AppTheme.weightMedium
                }
                MouseArea {
                    id: m2Hover
                    anchors.fill: parent
                    hoverEnabled: true
                    cursorShape: Qt.PointingHandCursor
                    onClicked: {
                        compCol.compMode = 1
                        compCol.recalcEndFromDays()
                    }
                }
            }
        }

        Rectangle { width: parent.width; height: 1; color: AppTheme.borderDivider }

        // ── Режим 0: ЭТОТ ДЕНЬ ───────────────────
        Column {
            width: parent.width
            spacing: AppTheme.spaceM
            visible: compCol.compMode === 0

            // Список типов зависит от того сменщик или нет
            // Сменщик:     Ночные / Сверх нормы / Дни
            // Пятидневка:  Ночные / Дни
            AppComboBox {
                id: singleCompType
                width: parent.width
                label: "Что списываем?"
                model: backend.isSelectedEmployeeShift
                    ? ["Ночные часы", "Сверх нормы", "Дни"]
                    : ["Ночные часы", "Дни"]
                onCurrentIndexChanged: {
                    singleErrorMsg.visible = false
                    compCol.validateSingleBalance()
                }
            }

            // Барабан с временем — для Ночных и Сверх нормы
            Column {
                width: parent.width
                spacing: AppTheme.spaceXS
                visible: {
                    if (backend.isSelectedEmployeeShift) {
                        // У сменщика: индекс 0 = Ночные, 1 = Сверх нормы, 2 = Дни
                        return singleCompType.currentIndex === 0 || singleCompType.currentIndex === 1
                    } else {
                        // У пятидневщика: индекс 0 = Ночные, 1 = Дни
                        return singleCompType.currentIndex === 0
                    }
                }

                Text {
                    text: "Количество:"
                    color: AppTheme.textSecondary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                }

                AppTumblerTime {
                    id: compTumbler
                    hours: 8
                    minutes: 0
                    onHoursChanged: {
                        singleErrorMsg.visible = false
                        compCol.validateSingleBalance()
                    }
                    onMinutesChanged: {
                        singleErrorMsg.visible = false
                        compCol.validateSingleBalance()
                    }
                }

                // Ошибка проверки остатков для режима «Этот день»
                Text {
                    id: singleErrorMsg
                    visible: false
                    width: parent.width
                    color: AppTheme.accentDanger
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                    wrapMode: Text.WordWrap
                }
            }
        }

        // ── Режим 1: ПЕРИОД ──────────────────────
        Item {
            width: parent.width
            implicitHeight: periodFields.implicitHeight
            visible: compCol.compMode === 1

            Item {
                x: 0; y: 0
                width: 14
                height: parent.height
                clip: true

                property real c1: periodStartInput.y + periodStartInput.height / 2
                property real c2: periodDaysInput.y + periodDaysInput.height / 2
                property real c3: periodEndInput.y + periodEndInput.height / 2

                Rectangle {
                    x: 4; y: parent.c1 - 1
                    width: 20
                    height: parent.c2 - parent.c1 + 2
                    color: "transparent"
                    border.color: AppTheme.borderDivider
                    border.width: 2
                    radius: AppTheme.radiusMedium
                }
                Rectangle {
                    x: 4; y: parent.c2 - 1
                    width: 20
                    height: parent.c3 - parent.c2 + 2
                    color: "transparent"
                    border.color: AppTheme.borderDivider
                    border.width: 2
                    radius: AppTheme.radiusMedium
                }
            }

            Column {
                id: periodFields
                x: 24
                width: parent.width - 24
                spacing: AppTheme.spaceM

                AppDateField {
                    id: periodStartInput
                    width: parent.width
                    label: "Начало периода:"
                    onSelectedDateChanged: {
                        if (periodUseShiftPattern.checked)
                            compCol.recalcShiftPattern()
                        else
                            compCol.recalcEndFromDays()
                    }
                }

                AppTextField {
                    id: periodDaysInput
                    width: parent.width
                    label: "Количество дней:"
                    text: "1"
                    validator: RegularExpressionValidator {
                        regularExpression: /^[1-9][0-9]*$/
                    }
                    onTextChanged: {
                        if (activeFocus && !periodUseShiftPattern.checked)
                            compCol.recalcEndFromDays()
                    }
                }

                AppDateField {
                    id: periodEndInput
                    width: parent.width
                    label: "Конец периода:"
                    onSelectedDateChanged: {
                        if (!compCol.isUpdating && !periodUseShiftPattern.checked)
                            compCol.recalcDaysFromEnd()
                    }
                }

                AppSwitch {
                    id: periodSkipWeekends
                    text: "Пропускать выходные (Сб, Вс)"
                    checked: true
                    visible: !periodUseShiftPattern.checked
                    onCheckedChanged: compCol.recalcEndFromDays()
                }

                AppSwitch {
                    id: periodUseShiftPattern
                    text: "По графику сменности"
                    checked: false
                    visible: backend.isSelectedEmployeeShift
                    onCheckedChanged: {
                        if (checked) {
                            periodSkipWeekends.checked = false
                            compCol.recalcShiftPattern()
                        } else {
                            compCol.patternConfidence = -1
                            compCol.recalcEndFromDays()
                        }
                    }
                }

                Rectangle {
                    id: patternHintBox
                    visible: periodUseShiftPattern.checked && compCol.patternConfidence >= 0
                    width: parent.width
                    height: visible ? patternHintText.implicitHeight + AppTheme.spaceS * 2 : 0
                    radius: AppTheme.radiusMedium
                    color: compCol.patternConfidence > 0.7
                        ? AppTheme.bgSuccessSoft
                        : (compCol.patternConfidence > 0.4
                            ? AppTheme.bgWarningSoft
                            : AppTheme.bgDangerSoft)

                    Text {
                        id: patternHintText
                        anchors.fill: parent
                        anchors.margins: AppTheme.spaceS
                        text: {
                            let c = compCol.patternConfidence
                            if (c < 0) return ""
                            let pct = Math.round(c * 100)
                            let cycle = compCol.patternCycle
                            let work = compCol.patternWorkDays
                            if (c > 0.7)
                                return "✓ Паттерн надёжен (" + pct + "%): цикл " + cycle + " дн., рабочих " + work + " дн."
                            if (c > 0.4)
                                return "⚠ Паттерн определён с оговорками (" + pct + "%): цикл " + cycle + " дн."
                            return "✗ Паттерн ненадёжен (" + pct + "%). Проверьте даты вручную."
                        }
                        color: AppTheme.textPrimary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                        wrapMode: Text.WordWrap
                    }
                }

                Text {
                    id: periodErrorMsg
                    visible: false
                    width: parent.width
                    color: AppTheme.accentDanger
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeSmall
                    wrapMode: Text.WordWrap
                }
            }
        }

        Rectangle { width: parent.width; height: 1; color: AppTheme.borderDivider }

        AppCheckBox {
            id: compPrevYearCheck
            text: "В счет прошлого года"
            width: parent.width
            onCheckedChanged: {
                compCol.validateBalances()
                compCol.validateSingleBalance()
            }
        }

        AppTextField {
            id: compCommentInput
            width: parent.width
            label: "Комментарий:"
            placeholderText: "Например: За работу в праздник"
        }

        // Общая ошибка в окне (а не тост): пустой период, сбой сохранения и т.п.
        Text {
            id: compErrorMsg
            visible: false
            width: parent.width
            color: AppTheme.accentDanger
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall
            wrapMode: Text.WordWrap
        }

        // ==========================================
        // ФУНКЦИИ КАЛЬКУЛЯТОРА
        // ==========================================

        function recalcEndFromDays() {
            if (isUpdating || !periodStartInput.selectedDate) return
            let d = parseInt(periodDaysInput.text)
            if (isNaN(d) || d < 1) return

            let parts = periodStartInput.selectedDate.split("-")
            let cur = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]))

            let added = 1
            let safety = 0
            while (added < d && safety < 365) {
                cur.setDate(cur.getDate() + 1)
                if (periodSkipWeekends.checked) {
                    if (cur.getDay() !== 0 && cur.getDay() !== 6) added++
                } else {
                    added++
                }
                safety++
            }

            let y = cur.getFullYear()
            let m = ("0" + (cur.getMonth() + 1)).slice(-2)
            let day = ("0" + cur.getDate()).slice(-2)

            isUpdating = true
            periodEndInput.selectedDate = y + "-" + m + "-" + day
            isUpdating = false

            validateBalances()
        }

        function recalcDaysFromEnd() {
            if (isUpdating || !periodStartInput.selectedDate || !periodEndInput.selectedDate) return

            let sp = periodStartInput.selectedDate.split("-")
            let ep = periodEndInput.selectedDate.split("-")
            let start = new Date(parseInt(sp[0]), parseInt(sp[1]) - 1, parseInt(sp[2]))
            let end   = new Date(parseInt(ep[0]), parseInt(ep[1]) - 1, parseInt(ep[2]))

            if (end < start) {
                isUpdating = true
                periodDaysInput.text = "1"
                isUpdating = false
                recalcEndFromDays()
                return
            }

            let cur = new Date(start)
            let count = 0
            let safety = 0
            while (cur <= end && safety < 365) {
                if (periodSkipWeekends.checked) {
                    if (cur.getDay() !== 0 && cur.getDay() !== 6) count++
                } else {
                    count++
                }
                cur.setDate(cur.getDate() + 1)
                safety++
            }

            if (count === 0) count = 1

            isUpdating = true
            periodDaysInput.text = count.toString()
            isUpdating = false

            validateBalances()
        }

        function recalcShiftPattern() {
            if (!periodStartInput.selectedDate || !periodEndInput.selectedDate) return

            compCol.patternConfidence = -1

            let result = backend.getShiftDatesForPeriod(
                periodStartInput.selectedDate,
                periodEndInput.selectedDate
            )

            if (result.error) {
                periodErrorMsg.text = "Не удалось определить паттерн: " + result.error
                periodErrorMsg.visible = true
                compCol.shiftDates = []
                compCol.patternConfidence = 0
                return
            }

            compCol.shiftDates = result.dates
            compCol.patternConfidence = result.confidence
            compCol.patternCycle = result.cycle_days || 0
            compCol.patternWorkDays = result.work_days || 0

            isUpdating = true
            periodDaysInput.text = result.dates.length.toString()
            isUpdating = false

            periodErrorMsg.visible = false
            validateBalances()
        }

        function validateBalances() {
            if (compCol.compMode !== 1 || !periodStartInput.selectedDate) {
                periodErrorMsg.visible = false
                return
            }

            let requestedDays = parseInt(periodDaysInput.text)
            if (isNaN(requestedDays) || requestedDays < 1) {
                periodErrorMsg.visible = false
                return
            }

            let selectedYear = parseInt(periodStartInput.selectedDate.split("-")[0]) || new Date().getFullYear()
            let checkYear = compPrevYearCheck.checked ? selectedYear - 1 : selectedYear
            let balances = backend.getAvailableBalances(checkYear)

            let availableDays  = balances["days"]  || 0
            let availableHours = balances["hours"] || 0
            let maxAllowedDays = availableDays + Math.floor(availableHours / 8)

            if (requestedDays > maxAllowedDays) {
                let yearLabel = compPrevYearCheck.checked ? "прошлого года" : "текущего года"
                periodErrorMsg.text = "Не хватает остатков " + yearLabel + "!\n" +
                    "Доступно: " + availableDays + " дн. и " + availableHours +
                    " ч. (Итого: " + maxAllowedDays + " дн.)"
                periodErrorMsg.visible = true
            } else {
                periodErrorMsg.visible = false
            }
        }

        // Проверка остатков для режима «Этот день»: ночные/сверх нормы часы и дни.
        // Показывает ошибку прямо в окне (как в денежной выплате), а не тост.
        function validateSingleBalance() {
            if (compCol.compMode !== 0) {
                singleErrorMsg.visible = false
                return
            }

            let unit = compCol.resolveCompUnit()
            let selectedYear = parseInt(root.targetDate.split('-')[0]) || new Date().getFullYear()
            let checkYear = compPrevYearCheck.checked ? selectedYear - 1 : selectedYear
            let balances = backend.getAvailableBalances(checkYear)
            let yearLabel = compPrevYearCheck.checked ? "прошлого года" : "текущего года"

            if (unit === "days") {
                let requested = 1
                let maxAllowed = balances["days"] || 0
                if (requested > maxAllowed) {
                    singleErrorMsg.text = "Не хватает остатков " + yearLabel + "! Доступно дней отгула: " + maxAllowed + "."
                    singleErrorMsg.visible = true
                } else {
                    singleErrorMsg.visible = false
                }
                return
            }

            // hours / overtime — сравниваем минуты
            let requestedMin = compTumbler.hours * 60 + compTumbler.minutes
            let availableMin = (balances[unit] || 0) * 60
            let label = unit === "overtime" ? "сверх нормы" : "ночных"
            if (requestedMin > availableMin) {
                let availH = Math.floor(availableMin / 60)
                let availM = availableMin % 60
                singleErrorMsg.text = "Не хватает остатков " + yearLabel + "! Доступно " + label +
                    ": " + availH + " ч." + (availM > 0 ? " " + availM + " мин." : "") + "."
                singleErrorMsg.visible = true
            } else {
                singleErrorMsg.visible = false
            }
        }

        // Вспомогательная функция — определяет unit для сохранения
        function resolveCompUnit() {
            if (backend.isSelectedEmployeeShift) {
                // 0 = Ночные, 1 = Сверх нормы, 2 = Дни
                if (singleCompType.currentIndex === 0) return "hours"
                if (singleCompType.currentIndex === 1) return "overtime"
                return "days"
            } else {
                // 0 = Ночные, 1 = Дни
                return singleCompType.currentIndex === 0 ? "hours" : "days"
            }
        }
    }

    // ==========================================
    // ФУНКЦИИ ОТКРЫТИЯ
    // ==========================================

    function prepareForComp(dateStr) {
        root.targetDate = dateStr
        root.editCompId = 0
        compCol.compMode = 0
        singleErrorMsg.visible = false
        compErrorMsg.visible = false
        periodErrorMsg.visible = false
        compCol.patternConfidence = -1
        periodUseShiftPattern.checked = false
        // Сбрасываем на первый пункт
        singleCompType.currentIndex = 0
        periodStartInput.selectedDate = dateStr
        periodDaysInput.text = "1"
        compCol.recalcEndFromDays()
        compCommentInput.text = ""
        compPrevYearCheck.checked = false
    }

    function openForComp(dateStr, callerItem, mouseX, mouseY) {
        prepareForComp(dateStr)
        mainWindow.activeSpotlightCell = callerItem
        root.showAt(callerItem, mouseX, mouseY)
    }

    // Версия для «морфинга»: окно сразу встаёт в нужный прямоугольник,
    // без анимации масштаба и без размытия фона.
    function openForCompMorph(dateStr, x, y, w, h) {
        prepareForComp(dateStr)
        mainWindow.activeSpotlightCell = null
        root.openMorph(x, y, w, h)
    }

    function openForCompEdit(compData, dateStr, callerItem, mouseX, mouseY) {
        root.targetDate = dateStr
        root.editCompId = compData.id

        compCol.compMode = 0
        singleErrorMsg.visible = false
        compErrorMsg.visible = false
        periodErrorMsg.visible = false
        compCol.patternConfidence = -1
        periodUseShiftPattern.checked = false
        compPrevYearCheck.checked = false

        // Восстанавливаем правильный индекс в зависимости от типа и графика
        if (backend.isSelectedEmployeeShift) {
            // 0 = Ночные, 1 = Сверх нормы, 2 = Дни
            if (compData.unit === "hours")    singleCompType.currentIndex = 0
            else if (compData.unit === "overtime") singleCompType.currentIndex = 1
            else singleCompType.currentIndex = 2
        } else {
            // 0 = Ночные, 1 = Дни
            singleCompType.currentIndex = compData.unit === "hours" ? 0 : 1
        }

        if (compData.unit === "hours" || compData.unit === "overtime") {
            compTumbler.hours   = Math.floor(compData.raw_amount / 60)
            compTumbler.minutes = compData.raw_amount % 60
        }

        compCommentInput.text = compData.comment
        mainWindow.activeSpotlightCell = callerItem
        root.showAt(callerItem, mouseX, mouseY)
    }

    // ==========================================
    // СОХРАНЕНИЕ
    // ==========================================
    onAccepted: {
        try {
            if (!root.targetDate) return

            if (compCol.compMode === 1 && periodErrorMsg.visible) {
                root.shake()
                return
            }

            // Режим «Этот день»: проверяем остатки прямо в окне (а не тостом)
            if (compCol.compMode === 0) {
                compCol.validateSingleBalance()
                if (singleErrorMsg.visible) {
                    root.shake()
                    return
                }
            }

            let finalDates = []
            let finalType  = compCol.resolveCompUnit()
            let amount     = "0"

            if (compCol.compMode === 0) {
                // Этот день
                finalDates.push(root.targetDate)

                if (finalType === "days") {
                    amount = "1"
                } else {
                    // hours или overtime — берём из барабана
                    amount = (compTumbler.hours * 60 + compTumbler.minutes).toString()
                }

            } else {
                // Период
                finalType = "days"
                amount = "1"

                if (periodUseShiftPattern.checked && compCol.shiftDates.length > 0) {
                    finalDates = compCol.shiftDates.slice()
                } else {
                    let sp  = periodStartInput.selectedDate.split("-")
                    let ep  = periodEndInput.selectedDate.split("-")
                    let cur = new Date(parseInt(sp[0]), parseInt(sp[1]) - 1, parseInt(sp[2]))
                    let end = new Date(parseInt(ep[0]), parseInt(ep[1]) - 1, parseInt(ep[2]))

                    let safety = 0
                    while (cur <= end && safety < 365) {
                        if (!periodSkipWeekends.checked ||
                            (cur.getDay() !== 0 && cur.getDay() !== 6)) {
                            let y = cur.getFullYear()
                            let m = ("0" + (cur.getMonth() + 1)).slice(-2)
                            let d = ("0" + cur.getDate()).slice(-2)
                            finalDates.push(y + "-" + m + "-" + d)
                        }
                        cur.setDate(cur.getDate() + 1)
                        safety++
                    }
                }
            }

            if (finalDates.length === 0) {
                compErrorMsg.text = "Ошибка: Нет дней в периоде"
                compErrorMsg.visible = true
                root.shake()
                return
            }
            compErrorMsg.visible = false

            let commentTxt = compCommentInput.text || ""

            if (root.editCompId > 0) {
                let amountForEdit = finalType === "days"
                    ? "1"
                    : (compTumbler.hours * 60 + compTumbler.minutes).toString()
                backend.updateCompensation(root.editCompId, root.targetDate,
                                           finalType, amountForEdit,
                                           String(commentTxt), compPrevYearCheck.checked)
            } else {
                backend.saveCompensation(finalDates.join(","), finalType, amount,
                                         String(commentTxt), compPrevYearCheck.checked)
            }

            root.close()

        } catch(e) {
            compErrorMsg.text = "Ошибка: " + e.message
            compErrorMsg.visible = true
            root.shake()
        }
    }
}
