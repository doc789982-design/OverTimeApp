import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl

Item {
    id: root
    height: 136

    Rectangle {
        anchors.fill: parent
        color: AppTheme.bgPanel 
    }

    Rectangle {
        anchors.fill: parent
        anchors.leftMargin: AppTheme.spaceM
        anchors.rightMargin: AppTheme.spaceM
        anchors.bottomMargin: AppTheme.spaceM 
        anchors.topMargin: AppTheme.spaceS     
        
        radius: AppTheme.radiusLarge
        color: AppTheme.bgElevated
        border.color: AppTheme.borderDivider
        border.width: 1

        ColumnLayout {
            anchors.fill: parent
            anchors.margins: AppTheme.spaceM 
            spacing: AppTheme.spaceS

            RowLayout {
                Layout.fillWidth: true
                spacing: AppTheme.spaceS

                // НАЗВАНИЕ ПОДРАЗДЕЛЕНИЯ (кликабельное)
                Item {
                    id: titleWrapper
                    Layout.fillWidth: true
                    Layout.preferredHeight: 36 
                    Layout.alignment: Qt.AlignVCenter
                    clip: true 

                    HoverHandler { id: titleHover }

                    property real amplitude: titleHover.hovered ? 1.0 : 0.0
                    Behavior on amplitude { 
                        NumberAnimation { duration: AppTheme.durSlow; easing.type: AppTheme.easeStandard } 
                    }

                    property real t: 0.0
                    NumberAnimation on t {
                        from: 0.0
                        to: 1.0
                        duration: 2000 
                        loops: Animation.Infinite
                        running: true 
                    }

                    Row {
                        anchors.verticalCenter: parent.verticalCenter
                        anchors.left: parent.left
                        spacing: 0 

                        Repeater {
                            model: backend.activeDepartmentName ? backend.activeDepartmentName.toUpperCase().split('') : []

                            Item {
                                width: charText.implicitWidth
                                height: charText.implicitHeight

                                Text {
                                    id: charText
                                    anchors.centerIn: parent
                                    text: modelData === " " ? "\u00A0" : modelData
                                    color: AppTheme.textPrimary
                                    font.family: AppTheme.fontCondensed
                                    font.pixelSize: AppTheme.sizeH2
                                    font.weight: AppTheme.weightBlack 
                                    font.letterSpacing: 1.5 
                                    transformOrigin: Item.Center

                                    property real phase: (index * 0.73) % 1.0

                                    rotation: Math.sin((titleWrapper.t + phase) * Math.PI * 2) * 2.5 * titleWrapper.amplitude
                                    anchors.verticalCenterOffset: Math.cos((titleWrapper.t + phase + 0.3) * Math.PI * 4) * 0.2 * titleWrapper.amplitude
                                }
                            }
                        }
                    }
                    
                    MouseArea {
                        anchors.fill: parent
                        cursorShape: Qt.PointingHandCursor
                        onClicked: dbSwitchMenu.popup(titleWrapper, 0, titleWrapper.height + AppTheme.spaceXXS)
                    }
                }

                Item {
                    Layout.preferredWidth: 32
                    Layout.preferredHeight: 32
                    Layout.alignment: Qt.AlignVCenter
                    HoverHandler { id: themeHover } 
                    AnimatedThemeButton {
                        anchors.fill: parent
                        onClicked: {
                            let globalPos = mapToItem(null, width / 2, height / 2)
                            themeTransition.execute(globalPos.x, globalPos.y, function() {
                                backend.toggleTheme()
                            })
                        }
                    }
                    AppToolTip {
                        anchors.horizontalCenter: parent.horizontalCenter
                        anchors.bottom: parent.top
                        anchors.bottomMargin: AppTheme.spaceS
                        text: backend.isDarkTheme ? "Светлая тема" : "Темная тема"
                        isVisible: themeHover.hovered 
                    }
                }

                Item {
                    Layout.preferredWidth: 32
                    Layout.preferredHeight: 32
                    Layout.alignment: Qt.AlignVCenter
                    HoverHandler { id: settingsHover }
                    AnimatedSettingsButton {
                        anchors.fill: parent
                        onClicked: { backend.loadDepartmentData(); backend.loadHotkeys(); settingsDialog.show() }
                    }
                    AppToolTip {
                        anchors.horizontalCenter: parent.horizontalCenter
                        anchors.bottom: parent.top
                        anchors.bottomMargin: AppTheme.spaceS
                        text: "Настройки"
                        isVisible: settingsHover.hovered
                    }
                }
            }

            RowLayout {
                Layout.fillWidth: true
                spacing: AppTheme.spaceS

                AppButton {
                    Layout.fillWidth: true
                    Layout.preferredHeight: 36
                    variant: "secondary" 
                    text: "Сотрудник" 
                    iconSource: "../icons/user_plus.svg" 
                    onClicked: {
                        empDialog.editId = 0
                        empDialog.lastName = ""
                        empDialog.firstName = ""
                        empDialog.middleName = ""
                        empDialog.rank = ""
                        empDialog.position = ""
                        empDialog.openHours = "0"
                        empDialog.openOvertime = "0"
                        empDialog.openDays = "0"
                        empDialog.prevOpenHours = "0"
                        empDialog.prevOpenOvertime = "0"
                        empDialog.prevOpenDays = "0"
                        let currentYear = backend.currentPeriodText.split(" ")[1];
                        if (!currentYear) {
                            currentYear = new Date().getFullYear().toString();
                        }
                        empDialog.startMonth = currentYear + "-01"
                        empDialog.showAt(parent, width / 2, -10)
                    }
                }

                Item {
                    Layout.preferredWidth: 36
                    Layout.preferredHeight: 36
                    HoverHandler { id: printHover }
                    AnimatedPrintButton {
                        anchors.fill: parent
                        onClicked: customPrintDialog.show()
                    }
                    AppToolTip {
                        anchors.horizontalCenter: parent.horizontalCenter
                        anchors.bottom: parent.top
                        anchors.bottomMargin: AppTheme.spaceS
                        text: "Печать"
                        isVisible: printHover.hovered
                    }
                }

                Item {
                    Layout.preferredWidth: 36
                    Layout.preferredHeight: 36
                    HoverHandler { id: exportHover }
                    AnimatedExportButton {
                        anchors.fill: parent
                        onClicked: exportDialog.open()
                    }
                    AppToolTip {
                        anchors.horizontalCenter: parent.horizontalCenter
                        anchors.bottom: parent.top
                        anchors.bottomMargin: AppTheme.spaceS
                        text: "Экспорт"
                        isVisible: exportHover.hovered
                    }
                }
            }
        }
    }
    
    // МЕНЮ БАЗ ДАННЫХ (вне основного Rectangle)
    AppDatabaseMenu {
        id: dbSwitchMenu
        
        Repeater {
            model: backend.dbList
            
            AppMenuItem {
                text: (backend.activeDepartmentName === modelData.name ? "✓ " : "") + modelData.name
                onClicked: backend.openDatabase(modelData.path)
            }
        }
        
        AppMenuSeparator {}
        
        AppMenuItem {
            text: "Управление базами..."
            iconSource: "../icons/settings.svg"
            onClicked: {
                backend.loadDepartmentData()
                backend.loadHotkeys()
                settingsDialog.show()
                settingsDialog.openTab(2)
            }
        }
    }
}