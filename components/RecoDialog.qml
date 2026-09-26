import QtQuick
import QtQuick.Controls
import QtQuick.Layouts

// ============================================================
// ОКНО «МЕТОДИЧЕСКИЕ РЕКОМЕНДАЦИИ»
//
// Читалка в стиле КонсультантПлюс: слева оглавление с быстрыми
// переходами (разделы И подпункты — клик переносит точно к пункту),
// справа — текст, выровненный по ширине, с жирными терминами
// и курсивными ссылками на нормативные акты. Примеры расчётов
// убраны в раскрывающиеся карточки-спойлеры.
// ============================================================
Popup {
    id: root

    property string pageTitle: "Методические рекомендации"
    property var reco: ({ title: "", sections: [] })
    property var atoms: []          // плоская модель контента
    property var tocModel: []       // оглавление: {level, atom, label}
    property int currentAtom: 0

    width: 1000
    height: Math.min(
        (ApplicationWindow.window ? ApplicationWindow.window.height : 800) * 0.9,
        (ApplicationWindow.window ? ApplicationWindow.window.height : 800) - 60)
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
        var a = []
        var toc = []
        var sections = (reco && reco.sections) ? reco.sections : []
        for (var i = 0; i < sections.length; i++) {
            var s = sections[i]
            toc.push({ level: 0, atom: a.length, label: s.num + ". " + s.title })
            a.push({ k: "sec", sec: i })
            for (var j = 0; j < s.subs.length; j++) {
                var sub = s.subs[j]
                toc.push({ level: 1, atom: a.length, label: sub.num + ". " + sub.title })
                a.push({ k: "sub", sec: i, sub: j })
                for (var b = 0; b < sub.blocks.length; b++) {
                    var blk = sub.blocks[b]
                    if (blk.t === "ex") {
                        a.push({ k: "ex", sec: i, sub: j, ex: blk })
                    } else {
                        a.push({ k: blk.t, sec: i, sub: j, x: blk.x })
                    }
                }
            }
        }
        atoms = a
        tocModel = toc
        currentAtom = 0
        showCentered()
        contentView.positionViewAtBeginning()
    }

    function jump(atom) {
        currentAtom = atom
        contentView.currentIndex = atom
        contentView.positionViewAtIndex(atom, ListView.Beginning)
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
                width: 310
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

                    ListView {
                        id: tocView
                        objectName: "tocView"
                        Layout.fillWidth: true
                        Layout.fillHeight: true
                        clip: true
                        spacing: 1
                        ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }

                        model: root.tocModel

                        delegate: Rectangle {
                            width: tocView.width
                            height: tocLabel.implicitHeight + 8
                            radius: AppTheme.radiusSmall
                            color: modelData.atom === root.currentAtom
                                   ? AppTheme.stateSelected
                                   : (tocHov.containsMouse ? AppTheme.stateHover : "transparent")
                            Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }

                            Text {
                                id: tocLabel
                                anchors.left: parent.left
                                anchors.right: parent.right
                                anchors.verticalCenter: parent.verticalCenter
                                anchors.leftMargin: (modelData.level === 0 ? 8 : 22)
                                anchors.rightMargin: 8
                                text: modelData.label
                                color: modelData.atom === root.currentAtom
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
                                onClicked: root.jump(modelData.atom)
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
                    spacing: AppTheme.spaceS
                    boundsBehavior: Flickable.StopAtBounds
                    ScrollBar.vertical: ScrollBar { policy: ScrollBar.AsNeeded }

                    // подсветка оглавления при прокрутке: последний пункт выше края
                    onContentYChanged: {
                        if (!moving) return
                        var idx = contentView.indexAt(contentView.contentX, contentView.contentY + 24)
                        if (idx >= 0) {
                            var cur = 0
                            for (var i = 0; i < root.tocModel.length; i++) {
                                if (root.tocModel[i].atom <= idx) cur = root.tocModel[i].atom
                                else break
                            }
                            if (cur !== root.currentAtom) root.currentAtom = cur
                        }
                    }

                    header: Item {
                        width: contentView.width
                        height: recoTitleText.implicitHeight + AppTheme.spaceM

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

                    model: root.atoms

                    // Один делегат — пять вариантов, включается нужный.
                    // (Inline-компоненты под Loader не видят контекст делегата,
                    // поэтому обычные дети с visible.)
                    delegate: Item {
                        id: atomDel
                        width: contentView.width
                        readonly property var d: modelData
                        height: d.k === "sec" ? secCol.height
                              : d.k === "sub" ? subCol.height
                              : d.k === "p" ? pText.height
                              : d.k === "li" ? liRow.height
                              : exFrame.height

                        // ── заголовок раздела ──
                        Column {
                            id: secCol
                            visible: atomDel.d.k === "sec"
                            width: parent.width
                            spacing: AppTheme.spaceXS
                            topPadding: AppTheme.spaceM

                            Text {
                                width: parent.width
                                text: root.reco.sections[atomDel.d.sec] ? root.reco.sections[atomDel.d.sec].num + ". " + root.reco.sections[atomDel.d.sec].title : ""
                                color: AppTheme.textPrimary
                                font.family: AppTheme.fontCondensed
                                font.pixelSize: AppTheme.sizeH4
                                font.weight: AppTheme.weightBold
                                wrapMode: Text.WordWrap
                            }
                            Text {
                                width: parent.width - AppTheme.spaceL
                                text: root.reco.sections[atomDel.d.sec] ? root.reco.sections[atomDel.d.sec].fullTitle : ""
                                color: AppTheme.textTertiary
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeSmall
                                wrapMode: Text.WordWrap
                            }
                            Rectangle { width: parent.width; height: 2; color: AppTheme.borderDivider }
                        }

                        // ── заголовок подпункта ──
                        Column {
                            id: subCol
                            visible: atomDel.d.k === "sub"
                            width: parent.width
                            spacing: AppTheme.spaceXS
                            topPadding: AppTheme.spaceM

                            Text {
                                width: parent.width
                                text: {
                                    var s = root.reco.sections[atomDel.d.sec]
                                    var sub = s ? s.subs[atomDel.d.sub] : null
                                    return sub ? sub.num + ". " + sub.title : ""
                                }
                                color: AppTheme.accentBrand
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeBodyLarge
                                font.weight: AppTheme.weightBold
                                wrapMode: Text.WordWrap
                            }
                            Rectangle { width: parent.width; height: 1; color: AppTheme.borderDivider; opacity: 0.6 }
                        }

                        // ── абзац ──
                        Text {
                            id: pText
                            visible: atomDel.d.k === "p"
                            width: parent.width
                            text: atomDel.d.x || ""
                            textFormat: Text.RichText
                            horizontalAlignment: Text.AlignJustify
                            wrapMode: Text.WordWrap
                            color: AppTheme.textPrimary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            lineHeight: 1.35
                        }

                        // ── пункт списка ──
                        Row {
                            id: liRow
                            visible: atomDel.d.k === "li"
                            width: parent.width
                            spacing: AppTheme.spaceXS

                            Rectangle {
                                width: 6; height: 6; radius: 3
                                color: AppTheme.accentBrand
                                opacity: 0.7
                                anchors.top: parent.top
                                anchors.topMargin: 8
                            }
                            Text {
                                width: parent.width - 6 - AppTheme.spaceXS
                                text: atomDel.d.x || ""
                                textFormat: Text.RichText
                                horizontalAlignment: Text.AlignJustify
                                wrapMode: Text.WordWrap
                                color: AppTheme.textPrimary
                                font.family: AppTheme.fontFamily
                                font.pixelSize: AppTheme.sizeBody
                                lineHeight: 1.35
                            }
                        }

                        // ── пример-спойлер ──
                        Rectangle {
                            id: exFrame
                            objectName: "exSpoiler"
                            visible: atomDel.d.k === "ex"
                            width: parent.width
                            radius: AppTheme.radiusMedium
                            color: AppTheme.bgSurface
                            border.color: exHov.containsMouse || expanded ? AppTheme.accentBrand : AppTheme.borderDivider
                            border.width: 1
                            clip: true

                            property bool expanded: false
                            height: exHeader.height + (expanded ? exBody.height + AppTheme.spaceS : 0) + AppTheme.spaceS * 2
                            Behavior on height { NumberAnimation { duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter } }
                            Behavior on border.color { ColorAnimation { duration: AppTheme.durMicro } }

                            // фирменная полоса слева
                            Rectangle {
                                anchors.left: parent.left
                                anchors.top: parent.top
                                anchors.bottom: parent.bottom
                                width: 3
                                color: AppTheme.accentBrand
                                opacity: 0.55
                            }

                            Column {
                                anchors.left: parent.left
                                anchors.right: parent.right
                                anchors.margins: AppTheme.spaceM
                                anchors.top: parent.top
                                spacing: AppTheme.spaceS

                                // заголовок-кнопка
                                Item {
                                    id: exHeader
                                    width: parent.width
                                    height: exTitleRow.implicitHeight

                                    RowLayout {
                                        id: exTitleRow
                                        width: parent.width
                                        spacing: AppTheme.spaceXS

                                        Rectangle {
                                            Layout.alignment: Qt.AlignTop
                                            Layout.topMargin: 2
                                            width: exBadge.implicitWidth + 14
                                            height: 20
                                            radius: AppTheme.radiusPill
                                            color: AppTheme.bgBrandSoft
                                            Text {
                                                id: exBadge
                                                anchors.centerIn: parent
                                                text: "ПРИМЕР"
                                                color: AppTheme.accentBrand
                                                font.family: AppTheme.fontFamily
                                                font.pixelSize: 10
                                                font.letterSpacing: 0.8
                                                font.weight: AppTheme.weightBold
                                            }
                                        }

                                        Text {
                                            Layout.fillWidth: true
                                            Layout.alignment: Qt.AlignTop
                                            text: atomDel.d.ex ? atomDel.d.ex.title : ""
                                            color: AppTheme.textPrimary
                                            font.family: AppTheme.fontFamily
                                            font.pixelSize: AppTheme.sizeBody
                                            font.weight: AppTheme.weightBold
                                            wrapMode: Text.WordWrap
                                        }

                                        IconImage {
                                            Layout.alignment: Qt.AlignTop
                                            Layout.topMargin: 3
                                            source: "../icons/chevron_down.svg"
                                            width: AppTheme.iconMedium; height: AppTheme.iconMedium
                                            color: AppTheme.textTertiary
                                            rotation: exFrame.expanded ? 180 : 0
                                            Behavior on rotation { NumberAnimation { duration: AppTheme.durFast; easing.type: AppTheme.easeEnter } }
                                        }
                                    }

                                    MouseArea {
                                        id: exHov
                                        anchors.fill: parent
                                        hoverEnabled: true
                                        cursorShape: Qt.PointingHandCursor
                                        onClicked: exFrame.expanded = !exFrame.expanded
                                    }
                                }

                                // содержимое примера
                                Column {
                                    id: exBody
                                    width: parent.width
                                    spacing: AppTheme.spaceS

                                    Repeater {
                                        model: atomDel.d.ex ? atomDel.d.ex.blocks : []

                                        delegate: Item {
                                            id: exBlock
                                            width: exBody.width
                                            readonly property var bd: modelData
                                            height: bd.t === "li" ? exLiRow.height : exPara.height

                                            Text {
                                                id: exPara
                                                visible: exBlock.bd.t !== "li"
                                                width: parent.width
                                                text: exBlock.bd.x || ""
                                                textFormat: Text.RichText
                                                horizontalAlignment: Text.AlignJustify
                                                wrapMode: Text.WordWrap
                                                color: AppTheme.textPrimary
                                                font.family: AppTheme.fontFamily
                                                font.pixelSize: AppTheme.sizeBody
                                                lineHeight: 1.35
                                            }

                                            Row {
                                                id: exLiRow
                                                visible: exBlock.bd.t === "li"
                                                width: parent.width
                                                spacing: AppTheme.spaceXS
                                                Rectangle {
                                                    width: 6; height: 6; radius: 3
                                                    color: AppTheme.accentWarning
                                                    anchors.top: parent.top
                                                    anchors.topMargin: 8
                                                }
                                                Text {
                                                    width: parent.width - 6 - AppTheme.spaceXS
                                                    text: exBlock.bd.x || ""
                                                    textFormat: Text.RichText
                                                    horizontalAlignment: Text.AlignJustify
                                                    wrapMode: Text.WordWrap
                                                    color: AppTheme.textPrimary
                                                    font.family: AppTheme.fontFamily
                                                    font.pixelSize: AppTheme.sizeBody
                                                    lineHeight: 1.35
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
        }
    }
}
