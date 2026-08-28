import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

// ============================================================
// ОКНО ДОБАВЛЕНИЯ ДЕЖУРСТВА
//
// Отдельное окно (не вкладка): собственный компонент поверх общей
// базы AppDialog. Окно компенсации — это уже другой элемент,
// у каждого своя геометрия и свой подбор высоты, поэтому при
// морфинге из меню дня не «перескакивает» с размера другого окна.
// ============================================================
AppDialog {
    id: root
    width: 380
    heightFraction: 2/3

    property string targetDate: ""
    property var activeBreaks: []
    property int editDutyId: 0

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
    // СОДЕРЖИМОЕ: ДЕЖУРСТВО
    // ==========================================
    Column {
        width: parent.width
        spacing: AppTheme.spaceM

        AppTimeInterval {
            id: dutyTimeInput
            anchors.horizontalCenter: parent.horizontalCenter
            startMinutes: 480
            endMinutes: 1200
        }

        AppSwitch {
            id: shiftCheckBox
            text: "Считать по графику сменности"
            checked: true
            anchors.horizontalCenter: parent.horizontalCenter
            visible: backend.isSelectedEmployeeShift
        }

        Rectangle { width: parent.width; height: 1; color: AppTheme.borderDivider }

        Text {
            text: "Перерывы (опционально):"
            color: AppTheme.textSecondary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall
            font.weight: AppTheme.weightBold
        }

        Flow {
            width: parent.width
            spacing: AppTheme.spaceS

            Repeater {
                model: root.activeBreaks
                AppPill {
                    removable: true
                    text: {
                        let sH = Math.floor(modelData.start / 60)
                        let sM = Math.floor(modelData.start % 60)
                        let eH = Math.floor(modelData.end / 60)
                        let eM = Math.floor(modelData.end % 60)
                        return ("0"+sH).slice(-2) + ":" + ("0"+sM).slice(-2) +
                               " — " + ("0"+eH).slice(-2) + ":" + ("0"+eM).slice(-2)
                    }
                    onRemoveClicked: {
                        let arr = root.activeBreaks.slice()
                        arr.splice(index, 1)
                        root.activeBreaks = arr
                    }
                }
            }

            AppPill {
                id: addBreakBtn
                isAction: true
                text: "Добавить"
                onClicked: addBreakDialog.showAt(addBreakBtn, addBreakBtn.width / 2, addBreakBtn.height)
            }
        }

        Rectangle { width: parent.width; height: 1; color: AppTheme.borderDivider }

        AppTextField {
            id: dutyCommentInput
            width: parent.width
            label: "Комментарий:"
            placeholderText: "Необязательно..."
        }
    }

    // ==========================================
    // ФУНКЦИИ ОТКРЫТИЯ
    // ==========================================

    function prepareForDuty(dateStr) {
        root.targetDate = dateStr
        root.editDutyId = 0
        dutyCommentInput.text = ""
        root.activeBreaks = []
    }

    function openForDuty(dateStr, callerItem, mouseX, mouseY) {
        prepareForDuty(dateStr)
        mainWindow.activeSpotlightCell = callerItem
        root.showAt(callerItem, mouseX, mouseY)
    }

    // Версия для «морфинга»: окно сразу встаёт в нужный прямоугольник,
    // без анимации масштаба и без размытия фона.
    function openForDutyMorph(dateStr, x, y, w, h) {
        prepareForDuty(dateStr)
        mainWindow.activeSpotlightCell = null
        root.openMorph(x, y, w, h)
    }

    function openForDutyEdit(dutyData, dateStr, callerItem, mouseX, mouseY) {
        root.targetDate = dateStr
        root.editDutyId = dutyData.id

        let sParts = dutyData.start.split(":")
        dutyTimeInput.startMinutes = parseInt(sParts[0]) * 60 + parseInt(sParts[1])
        let eParts = dutyData.end.split(":")
        dutyTimeInput.endMinutes = parseInt(eParts[0]) * 60 + parseInt(eParts[1])

        shiftCheckBox.checked = dutyData.is_shift
        dutyCommentInput.text = dutyData.comment
        root.activeBreaks = dutyData.breaks || []

        mainWindow.activeSpotlightCell = callerItem
        root.showAt(callerItem, mouseX, mouseY)
    }

    // ==========================================
    // СОХРАНЕНИЕ
    // ==========================================
    onAccepted: {
        try {
            if (!root.targetDate) return

            let sH = Math.floor(dutyTimeInput.startMinutes / 60)
            let sM = dutyTimeInput.startMinutes % 60
            let sStr = ("0" + sH).slice(-2) + ":" + ("0" + sM).slice(-2)

            let eH = Math.floor(dutyTimeInput.endMinutes / 60)
            let eM = dutyTimeInput.endMinutes % 60
            let eStr = ("0" + eH).slice(-2) + ":" + ("0" + eM).slice(-2)

            let breaksArray = []
            for (let i = 0; i < root.activeBreaks.length; i++) {
                let b = root.activeBreaks[i]
                breaksArray.push({
                    "start_h": Math.floor(b.start / 60),
                    "start_m": b.start % 60,
                    "end_h":   Math.floor(b.end / 60),
                    "end_m":   b.end % 60
                })
            }

            let finalShift = backend.isSelectedEmployeeShift ? shiftCheckBox.checked : false

            if (root.editDutyId > 0) {
                backend.updateDuty(root.editDutyId, root.targetDate, sStr, eStr,
                                   finalShift, dutyCommentInput.text, JSON.stringify(breaksArray))
            } else {
                backend.saveDuty(root.targetDate, sStr, eStr,
                                 finalShift, dutyCommentInput.text, JSON.stringify(breaksArray))
            }

            root.close()

        } catch(e) {
            backend.showToast("Ошибка: " + e.message, "error")
        }
    }

    // ==========================================
    // ДИАЛОГ ДОБАВЛЕНИЯ ПЕРЕРЫВА
    // ==========================================
    AppDialog {
        id: addBreakDialog
        parent: Overlay.overlay
        width: 320
        title: ""
        acceptText: "Сохранить"

        AppTimeInterval {
            id: newBreakTimeInput
            anchors.horizontalCenter: parent.horizontalCenter
            startMinutes: 13 * 60
            endMinutes: 14 * 60
        }

        onAccepted: {
            let arr = root.activeBreaks.slice()
            arr.push({ start: newBreakTimeInput.startMinutes, end: newBreakTimeInput.endMinutes })
            arr.sort(function(a, b) { return a.start - b.start })
            root.activeBreaks = arr
            addBreakDialog.close()
        }
    }
}
