import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

// ============================================================
// ОКНО «МЕТОДИЧЕСКИЕ РЕКОМЕНДАЦИИ»
//
// Читалка в стиле КонсультантПлюс: слева оглавление с быстрыми
// переходами, справа текст. Клик по пункту — мгновенный переход
// к разделу; прокрутка текста подсвечивает текущий пункт.
// Сверху — кнопки производственных календарей (2026/2027).
// ============================================================
Popup {
    id: root

    property string pageTitle: "Методические рекомендации"
    property var reco: ({ title: "", sections: [] })
    property var tocModel: []          // плоское оглавление: {sec, level, label}
    property int currentSection: 0

    signal requestCalendar(int year)

    width: 980
    height: Math.min(
        (ApplicationWindow.window ? ApplicationWindow.window.height : 800) * 0.88,
        (ApplicationWindow.window ? ApplicationWindow.window.height : 800) - 80)
    z: AppTheme.zModal
    modal: true
    dim: true
    focus: true
    closePolicy: Popup.CloseOnEscape | Popup.CloseOnPressOutside

    enter: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
            NumberAnimation { property: "scale"; from: 0.96; to: 1.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
        }
    }
    exit: Transition {
        ParallelAnimation {
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
            NumberAnimation { property: "scale"; from: 1.0; to: 0.96; duration: AppTheme.durFast; easing.type: AppTheme.easeExit }
        }
    }

    background: Rectangle {
        color: AppTheme.bgModal
        radius: AppTheme.radiusModal
        border.color: AppTheme.borderDivider
        border.width: 1
        AppShadow { level: 4 }
    }

    function openReco() {
        reco = backend.getRecoSections()
        var toc = []
        var sections = (reco && reco.sections) ? reco.sections : []
        for (var i = 0; i < sections.length; i++) {
            toc.push({ sec: i, level: 0, label: sections[i].num + ". " + sections[i].title })
            var subs = sections[i].subs || []
            for (var j = 0; j < subs.length; j++)
                toc.push({ sec: i, level: 1, label: subs[j].num + ". " + subs[j].title })
        }
        tocModel = toc
        currentSection = 0
        showCentered()
        contentView.positionViewAtBeginning()
    }

    function jump(sec) {
        currentSection = sec
        contentView.currentIndex = sec
        contentView.positionViewAtIndex(sec, ListView.Beginning)
    }

    function showCentered() {
        if (ApplicationWindow.window) {
            root.x = (ApplicationWindow.window.width - root.width) / 2
            root.y = (ApplicationWindow.window.height - root.height) / 2
        }
        root.open()
    }

    contentItem: Item {
        anchors.fill: parent

        // ================= ШАПКА =================
        Item {
            id: headerBar
            anchors.top: parent.top
            anchors.left: parent.left
            anchors.right: parent.right
            height: 60

            Text {
                anchors.left: parent.left
                anchors.leftMargin: AppTheme.spaceL
                anchors.verticalCenter: parent.verticalCenter
                text: root.pageTitle
                color: AppTheme.textPrimary
                font.family: AppTheme.fontCondensed
                font.pixelSize: AppTheme.sizeH4
                font.weight: AppTheme.weightBold
            }

            Rectangle {
                width: 32; height: 32; radius: AppTheme.radiusPill
                anchors.right: parent.right
                anchors.rightMargin: AppTheme.spaceM
                anchors.verticalCenter: parent.verticalCenter
                color: closeHov.pressed ? AppTheme.statePress : (closeHov.containsMouse ? AppTheme.stateHover : "transparent")
                IconImage { anchors.centerIn: parent; source: "../icons/close.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: AppTheme.textSecondary }
                MouseArea { id: closeHov; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: root.close() }
            }
        }

        Rectangle {
            anchors.left: parent.left; anchors.right: parent.right
            anchors.top: headerBar.bottom
            height: 1
            color: AppTheme.borderDivider
        }

        // ================= ТЕЛО: ОГЛАВЛЕНИЕ + ТЕКСТ =================
        Item {
            anchors.top: headerBar.bottom
            anchors.topMargin: 1
            anchors.left: parent.left
            anchors.right: parent.right
            anchors.bottom: parent.bottom

            // ── ЛЕВАЯ КОЛОНКА: ОГЛАВЛЕНИЕ ──
            Rectangle {
                id: tocPanel
                width: 300
                anchors.left: parent.left
                anchors.top: parent.top
                anchors.bottom: parent.bottom
                anchors.margins: AppTheme.spaceM
                radius: AppTheme.radiusMedium
                color: AppTheme.bgSurface
                border.color: AppTheme.borderDivider
                border.width: 1

                ColumnLayout {
                    anchors.fill: parent
                    anchors.margins: AppTheme.spaceS
                    spacing: AppTheme.spaceS

                    Text {
                        Layout.fillWidth: true
                        Layout.leftMargin: AppTheme.spaceXXS
                        text: "ОГЛАВЛЕНИЕ"
                        color: AppTheme.textTertiary
                        font.family: AppTheme.fontFamily
                        font.pixelSize: AppTheme.sizeSmall
                        font.letterSpacing: 1
                        font.weight: AppTheme.weightBold
                    }

                    // Кнопки производственных календарей
                    Row {
                        Layout.fillWidth: true
                        Layout.leftMargin: AppTheme.spaceXXS
                        spacing: AppTheme.spaceXS

                        Repeater {
                            model: [2026, 2027]
                            delegate: Rectangle {
                                width: (tocPanel.width - AppTheme.spaceS * 2 - AppTheme.spaceXS * 2 - AppTheme.spaceXXS) / 2
                                height: 34
                                radius: AppTheme.radiusMedium
                                color: calHov.containsMouse ? AppTheme.stateHover : AppTheme.bgElevated
                                border.color: AppTheme.borderDivider
                                border.width: 1
                                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                                Row {
                                    anchors.centerIn: parent
                                    spacing: AppTheme.spaceXXS
                                    IconImage { source: "../icons/calendar.svg"; width: AppTheme.iconMedium; height: AppTheme.iconMedium; color: AppTheme.accentBrand; anchors.verticalCenter: parent.verticalCenter }
                                    Text {
                                        anchors.verticalCenter: parent.verticalCenter
                                        text: modelData
                                        color: AppTheme.textPrimary
                                        font.family: AppTheme.fontFamily
                                        font.pixelSize: AppTheme.sizeBody
                                        font.weight: AppTheme.weightBold
                                    }
                                }
                                MouseArea { id: calHov; anchors.fill: parent; hoverEnabled: true; cursorShape: Qt.PointingHandCursor; onClicked: root.requestCalendar(modelData) }
                            }
                        }
                    }

                    Rectangle { Layout.fillWidth: true; height: 1; color: AppTheme.borderDivider; Layout.topMargin: AppTheme.spaceXXS }

                    // Само оглавление
                    ListView {
                        id: tocView
                        objectName: "tocView"
                        Layout.fillWidth: true
                        Layout.fillHeight: true
                        clip: true
                        spacing: 2
                        ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }

                        model: root.tocModel

                        delegate: Rectangle {
                            width: tocView.width
                            height: tocLabel.implicitHeight + 10
                            radius: AppTheme.radiusSmall
                            color: modelData.sec === root.currentSection
                                   ? AppTheme.stateSelected
                                   : (tocHov.containsMouse ? AppTheme.stateHover : "transparent")
                            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                            Text {
                                id: tocLabel
                                anchors.left: parent.left
                                anchors.right: parent.right
                                anchors.verticalCenter: parent.verticalCenter
                                anchors.leftMargin: (modelData.level === 0 ? 8 : 20)
                                anchors.rightMargin: 8
                                text: modelData.label
                                color: modelData.sec === root.currentSection
                                      ? AppTheme.textOnSoft
                                      : (modelData.level === 0 ? AppTheme.textPrimary : AppTheme.textSecondary)
                                font.family: AppTheme.fontFamily
                                font.pixelSize: modelData.level === 0 ? AppTheme.sizeBody : AppTheme.sizeSmall
                                font.weight: modelData.level === 0 ? AppTheme.weightBold : AppTheme.weightMedium
                                wrapMode: Text.WordWrap
                                elide: Text.ElideRight
                                maximumLineCount: 2
                            }

                            MouseArea {
                                id: tocHov
                                anchors.fill: parent
                                hoverEnabled: true
                                cursorShape: Qt.PointingHandCursor
                                onClicked: root.jump(modelData.sec)
                            }
                        }
                    }
                }
            }

            // ── ПРАВАЯ КОЛОНКА: ТЕКСТ ──
            Item {
                anchors.left: tocPanel.right
                anchors.leftMargin: AppTheme.spaceM
                anchors.right: parent.right
                anchors.rightMargin: AppTheme.spaceM
                anchors.top: parent.top
                anchors.bottom: parent.bottom

                ListView {
                    id: contentView
                    objectName: "recoContentView"
                    anchors.fill: parent
                    clip: true
                    spacing: AppTheme.spaceL
                    ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }

                    // Подсветка оглавления при прокрутке текста
                    onContentYChanged: {
                        if (!moving) return
                        var idx = contentView.indexAt(contentView.contentX, contentView.contentY + 16)
                        if (idx >= 0 && idx !== root.currentSection)
                            root.currentSection = idx
                    }

                    header: Item {
                        width: contentView.width
                        height: recoTitleText.implicitHeight + AppTheme.spaceL

                        Text {
                            id: recoTitleText
                            anchors.left: parent.left; anchors.right: parent.right
                            text: root.reco.title || ""
                            color: AppTheme.textTertiary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeSmall
                            wrapMode: Text.WordWrap
                        }
                    }

                    model: (root.reco && root.reco.sections) ? root.reco.sections : []

                    delegate: Column {
                        width: contentView.width
                        spacing: AppTheme.spaceM

                        // Заголовок раздела
                        Text {
                            width: parent.width
                            text: modelData.num + ". " + modelData.title
                            color: AppTheme.textPrimary
                            font.family: AppTheme.fontCondensed
                            font.pixelSize: AppTheme.sizeH5
                            font.weight: AppTheme.weightBold
                            wrapMode: Text.WordWrap
                        }

                        Rectangle { width: parent.width; height: 1; color: AppTheme.borderDivider }

                        // Абзацы раздела
                        Repeater {
                            model: modelData.paras
                            Text {
                                width: parent.width
                                text: modelData
                                color: AppTheme.textPrimary
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeBody
                                lineHeight: 1.35
                                wrapMode: Text.WordWrap
                            }
                        }

                        // Подпункты
                        Repeater {
                            model: modelData.subs
                            Column {
                                width: parent.width
                                spacing: AppTheme.spaceS
                                topPadding: AppTheme.spaceS

                                Text {
                                    width: parent.width
                                    text: modelData.num + ". " + modelData.title
                                    color: AppTheme.accentBrand
                                    font.family: AppTheme.fontFamily
                                    font.pixelSize: AppTheme.sizeBody
                                    font.weight: AppTheme.weightBold
                                    wrapMode: Text.WordWrap
                                }

                                Repeater {
                                    model: modelData.paras
                                    Text {
                                        width: parent.width
                                        text: modelData
                                        color: AppTheme.textPrimary
                                        font.family: AppTheme.fontFamily
                                        font.pixelSize: AppTheme.sizeBody
                                        lineHeight: 1.35
                                        wrapMode: Text.WordWrap
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
