import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

Item {
    id: root
    implicitHeight: 60 

    signal monthClicked(int monthNum) 
    signal yearViewClicked()          
    signal yearChanged(int newYear)   

    property bool isYearView: false
    property string currentYearText: backend.currentPeriodText.split(" ")[1]

    // --- ВСПОМОГАТЕЛЬНАЯ ЛОГИКА ПРОВЕРКИ ---
    function isMonthDisabled(monthIdx) {
        if (backend.selectedEmployeeId === 0) return false;
        
        // Разбираем дату приема: "2026-03" -> год 2026, месяц 3
        let start = backend.selectedEmployeeStartMonth.split("-");
        let startYear = parseInt(start[0]);
        let startMonth = parseInt(start[1]);
        let currentYear = parseInt(root.currentYearText);
        let monthNum = monthIdx + 1;

        if (currentYear < startYear) return true;
        if (currentYear === startYear && monthNum < startMonth) return true;
        
        return false;
    }


    RowLayout {
        anchors.fill: parent
        anchors.leftMargin: AppTheme.spaceL; anchors.rightMargin: AppTheme.spaceL
        spacing: AppTheme.spaceL

        // ==========================================
        // 1. ВЫБОР ГОДА
        // ==========================================
        Rectangle {
            Layout.alignment: Qt.AlignVCenter
            width: 80; height: 36; radius: AppTheme.radiusSmall
            color: yearHover.pressed ? AppTheme.statePress : (yearHover.containsMouse ? AppTheme.stateHover : "transparent")
            border.color: yearHover.containsMouse ? AppTheme.borderInput : "transparent"; border.width: 1
            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

            Row {
                anchors.centerIn: parent; spacing: AppTheme.spaceXXS
                Text { text: root.currentYearText; color: AppTheme.textPrimary; font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody; font.weight: AppTheme.weightMedium }
                Text { text: "▾"; color: AppTheme.textSecondary; font.pixelSize: AppTheme.sizeSmall; anchors.verticalCenter: parent.verticalCenter; anchors.verticalCenterOffset: 1 }
            }

            AppMenu {
                id: yearMenu
                Repeater {
                    model: backend.yearList
                    AppMenuItem {
                        text: modelData.toString()
                        
                        // Блокировка годов меньше года приема
                        property bool isYearDisabled: {
                            if (backend.selectedEmployeeId === 0) return false;
                            let startYear = parseInt(backend.selectedEmployeeStartMonth.split("-")[0]);
                            return modelData < startYear;
                        }
                        
                        enabled: !isYearDisabled
                        opacity: isYearDisabled ? 0.3 : 1.0
                        onClicked: root.yearChanged(modelData)
                    }
                }
            }

            MouseArea {
                id: yearHover; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor
                onClicked: yearMenu.popup(parent, 0, parent.height + AppTheme.spaceXXS)
            }
        }

        Item { Layout.fillWidth: true }

        // ==========================================
        // 2. ВКЛАДКИ МЕСЯЦЕВ
        // ==========================================
        Row {
            Layout.alignment: Qt.AlignBottom 
            height: parent.height

            Repeater {
                model: ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек", "Год"]

                Item {
                    width: index === 12 ? 60 : 44; height: parent.height
                    property bool isPseudoMonth: index === 12
                    property int monthNum: index + 1
                    property bool isActive: isPseudoMonth ? root.isYearView : (!root.isYearView && backend.currentPeriodText.startsWith(modelData))
                    
                    // МАГИЯ БЛОКИРОВКИ
                    property bool disabled: !isPseudoMonth && root.isMonthDisabled(index)

                    Rectangle { visible: parent.isPseudoMonth; width: 1; height: 20; color: AppTheme.borderDivider; anchors.left: parent.left; anchors.verticalCenter: parent.verticalCenter }

                    Text {
                        anchors.centerIn: parent
                        text: modelData
                        color: parent.isActive ? AppTheme.accentBrand : 
                               (parent.disabled ? AppTheme.textDisabled : 
                               (tabHover.containsMouse ? AppTheme.textPrimary : AppTheme.textSecondary))
                        font.family: AppTheme.fontFamily; font.pixelSize: AppTheme.sizeBody
                        font.weight: parent.isActive ? AppTheme.weightBold : AppTheme.weightMedium
                        opacity: parent.disabled ? 0.3 : 1.0
                        Behavior on color { ColorAnimation { duration: AppTheme.durFast } }
                    }

                    // Зеленая точка
                    Rectangle {
                        visible: !parent.isPseudoMonth && backend.monthPulse[parent.monthNum] === true && !parent.isActive && !parent.disabled
                        width: 6; height: 6; radius: AppTheme.radiusPill; anchors.top: parent.top; anchors.topMargin: AppTheme.spaceS; anchors.right: parent.right; anchors.rightMargin: AppTheme.spaceXXS; color: AppTheme.accentSuccess
                    }

                    Rectangle {
                        anchors.bottom: parent.bottom; anchors.horizontalCenter: parent.horizontalCenter
                        width: parent.isActive ? parent.width - AppTheme.spaceS : 0; height: 2; color: AppTheme.accentBrand
                        Behavior on width { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeStandard } }
                    }

                    MouseArea {
                        id: tabHover; anchors.fill: parent; hoverEnabled: true
                        // Если месяц заблокирован — клики не проходят
                        enabled: !parent.disabled
                        cursorShape: parent.disabled ? Qt.ArrowCursor : Qt.PointingHandCursor
                        onClicked: {
                            if (parent.isPseudoMonth) root.yearViewClicked()
                            else root.monthClicked(parent.monthNum)
                        }
                    }
                }
            }
        }
    }
}