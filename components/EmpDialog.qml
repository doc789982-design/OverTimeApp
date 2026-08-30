import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

AppDialog {
    id: root
    width: 380
    
    property int editId: 0
    property int balanceYear: 0  // 0 = Текущий, 1 = Предыдущий — ПЕРЕНЕСЛИ СЮДА

    property alias lastName: empLastName.text
    property alias firstName: empFirstName.text
    property alias middleName: empMiddleName.text
    property alias rank: empRank.text
    property alias position: empPosition.text
    property alias startMonth: empStartMonth.selectedMonth
    property alias openHours: empOpenHours.text
    property alias openOvertime: empOpenOvertime.text
    property alias openDays: empOpenDays.text
    property alias prevOpenHours: empPrevOpenHours.text
    property alias prevOpenOvertime: empPrevOpenOvertime.text
    property alias prevOpenDays: empPrevOpenDays.text    

    title: root.editId === 0 ? "Новый сотрудник" : "Редактирование"
    acceptText: "Сохранить"

    onAboutToShow: empErrorMsg.visible = false

    Column {
        width: parent.width
        spacing: AppTheme.spaceM

        AppTextField { id: empLastName; width: parent.width; label: "Фамилия"; isRequired: true }
        AppTextField { id: empFirstName; width: parent.width; label: "Имя"; isRequired: true }
        AppTextField { id: empMiddleName; width: parent.width; label: "Отчество" }
        AppTextField { id: empRank; width: parent.width; label: "Звание" }
        AppTextField { id: empPosition; width: parent.width; label: "Должность" }
        AppMonthField { id: empStartMonth; width: parent.width; label: "Месяц приема"; isRequired: true }

        // ==========================================
        // ЗАГОЛОВОК + ПЕРЕКЛЮЧАТЕЛЬ
        // ==========================================
        Row {
            width: parent.width
            spacing: AppTheme.spaceM

            Text {
                text: "Балансы:"
                color: AppTheme.textSecondary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeSmall
                font.weight: AppTheme.weightBold
                anchors.verticalCenter: parent.verticalCenter
            }

            Row {
                spacing: AppTheme.spaceM
                anchors.verticalCenter: parent.verticalCenter

                Rectangle {
                    width: 110
                    height: 36
                    radius: AppTheme.radiusMedium
                    color: root.balanceYear === 0 ? AppTheme.stateSelected : (yearHover1.containsMouse ? AppTheme.stateHover : "transparent")
                    Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                    Text {
                        anchors.centerIn: parent
                        text: "Текущий год"
                        color: root.balanceYear === 0 ? AppTheme.textOnSoft : AppTheme.textSecondary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        font.weight: root.balanceYear === 0 ? AppTheme.weightBold : AppTheme.weightMedium
                    }

                    MouseArea {
                        id: yearHover1
                        anchors.fill: parent
                        hoverEnabled: true
                        cursorShape: Qt.PointingHandCursor
                        onClicked: root.balanceYear = 0
                    }
                }

                Rectangle {
                    width: 130
                    height: 36
                    radius: AppTheme.radiusMedium
                    color: root.balanceYear === 1 ? AppTheme.stateSelected : (yearHover2.containsMouse ? AppTheme.stateHover : "transparent")
                    Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                    Text {
                        anchors.centerIn: parent
                        text: "Предыдущий год"
                        color: root.balanceYear === 1 ? AppTheme.textOnSoft : AppTheme.textSecondary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeBody
                        font.weight: root.balanceYear === 1 ? AppTheme.weightBold : AppTheme.weightMedium
                    }

                    MouseArea {
                        id: yearHover2
                        anchors.fill: parent
                        hoverEnabled: true
                        cursorShape: Qt.PointingHandCursor
                        onClicked: root.balanceYear = 1
                    }
                }
            }
        }

        // ==========================================
        // ПОЛЯ В ОДНУ СТРОКУ
        // ==========================================
        Row {
            width: parent.width
            spacing: AppTheme.spaceS

            AppTextField {
                id: empOpenHours
                numericOnly: true
                width: (parent.width - AppTheme.spaceS * 2) / 3
                label: "Ночные (ч)"
                text: "0"
                visible: root.balanceYear === 0
            }
            AppTextField {
                id: empPrevOpenHours
                numericOnly: true
                width: (parent.width - AppTheme.spaceS * 2) / 3
                label: "Ночные (ч)"
                text: "0"
                visible: root.balanceYear === 1
            }

            AppTextField {
                id: empOpenOvertime
                numericOnly: true
                width: (parent.width - AppTheme.spaceS * 2) / 3
                label: "Сверх нормы (ч)"
                text: "0"
                visible: root.balanceYear === 0
            }
            AppTextField {
                id: empPrevOpenOvertime
                numericOnly: true
                width: (parent.width - AppTheme.spaceS * 2) / 3
                label: "Сверх нормы (ч)"
                text: "0"
                visible: root.balanceYear === 1
            }

            AppTextField {
                id: empOpenDays
                numericOnly: true
                width: (parent.width - AppTheme.spaceS * 2) / 3
                label: "Дни"
                text: "0"
                visible: root.balanceYear === 0
            }
            AppTextField {
                id: empPrevOpenDays
                numericOnly: true
                width: (parent.width - AppTheme.spaceS * 2) / 3
                label: "Дни"
                text: "0"
                visible: root.balanceYear === 1
            }
        }

        // Ошибка в окне (а не тост): обязательные поля
        Text {
            id: empErrorMsg
            visible: false
            width: parent.width
            color: AppTheme.accentDanger
            font.family: AppTheme.fontFamily
            font.pixelSize: AppTheme.sizeSmall
            wrapMode: Text.WordWrap
        }
    }

    onAccepted: {
        if (empLastName.text.trim() === "" || empFirstName.text.trim() === "") {
            empErrorMsg.text = "Ошибка: Заполните обязательные поля (фамилия и имя)"
            empErrorMsg.visible = true
            root.shake()
            return
        }
        empErrorMsg.visible = false

        let cur_mins = parseInt(empOpenHours.text || "0") * 60
        let cur_over = parseInt(empOpenOvertime.text || "0") * 60
        let cur_days = parseInt(empOpenDays.text || "0")
        
        let prev_mins = parseInt(empPrevOpenHours.text || "0") * 60
        let prev_over = parseInt(empPrevOpenOvertime.text || "0") * 60
        let prev_days = parseInt(empPrevOpenDays.text || "0")

        if (root.editId === 0) {
            backend.saveEmployee(
                empLastName.text, empFirstName.text, empMiddleName.text, 
                empRank.text, empPosition.text, empStartMonth.selectedMonth,
                cur_mins, cur_over, cur_days,
                prev_mins, prev_over, prev_days
            )
        } else {
            backend.updateEmployee(
                root.editId, empLastName.text, empFirstName.text, 
                empMiddleName.text, empRank.text, empPosition.text, empStartMonth.selectedMonth,
                cur_mins, cur_over, cur_days,
                prev_mins, prev_over, prev_days
            )
        }
        root.close()
    }
}