import QtQuick
import QtQuick.Controls

AppDialog {
    id: root
    width: 360

    title: "Новая группа"
    acceptText: "Создать"

    onAboutToShow: {
        grpNameInput.text = ""
        grpShiftCheck.checked = false
        grpShiftedCheck.checked = false
        grpErrorMsg.visible = false
    }

    AppTextField {
        id: grpNameInput
        width: parent.width
        label: "Название группы"
        isRequired: true
        placeholderText: "Смена 1"
    }

    AppSwitch {
        id: grpShiftCheck
        text: "Сменный график"
    }

    Column {
        width: parent.width
        spacing: 4

        AppSwitch {
            id: grpShiftedCheck
            text: "Смещённые выходные"
        }
        Text {
            width: parent.width
            text: "Суббота рабочая, понедельник выходной. Воскресенье как обычно. Праздники и дни, которые вы меняете вручную, общие для всех групп."
            color: AppTheme.textTertiary
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall
            wrapMode: Text.WordWrap
        }
    }

    // Ошибка в окне (а не тост)
    Text {
        id: grpErrorMsg
        visible: false
        width: parent.width
        color: AppTheme.accentDanger
        font.family: AppTheme.fontFamily
        font.pixelSize: AppTheme.sizeSmall
        wrapMode: Text.WordWrap
    }

    onAccepted: {
        if (grpNameInput.text.trim() === "") {
            grpErrorMsg.text = "Ошибка: Введите название группы"
            grpErrorMsg.visible = true
            root.shake()
            return
        }
        grpErrorMsg.visible = false

        backend.createGroup(grpNameInput.text, grpShiftCheck.checked, grpShiftedCheck.checked)
        grpNameInput.text = ""
        root.close()
    }
}
