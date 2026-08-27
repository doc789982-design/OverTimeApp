import QtQuick
import QtQuick.Controls
import QtQuick.Controls.impl

// Полоски как у «Обновить»: та же форма и то же место.
// Фон нейтральный. Новые тосты выезжают снизу, стопка растёт вверх.
// Если висит баннер обновления — тосты над ним (этот хост ставят выше баннера).
Item {
    id: root
    implicitHeight: col.implicitHeight
    height: implicitHeight
    clip: false

    property int barHeight: 52
    property int maxVisible: 3
    property int nextId: 1

    ListModel { id: toasts }

    function show(message, type) {
        var kind = type || "success"
        if (toasts.count >= root.maxVisible)
            toasts.remove(0)
        toasts.append({
            mid: root.nextId,
            message: message || "",
            kind: kind
        })
        root.nextId += 1
    }

    function dismissAt(index) {
        if (index >= 0 && index < toasts.count)
            toasts.remove(index)
    }

    Connections {
        target: backend
        function onShowToast(message, type) {
            root.show(message, type)
        }
    }

    Column {
        id: col
        anchors.left: parent.left
        anchors.right: parent.right
        anchors.bottom: parent.bottom
        spacing: 0

        move: Transition {
            NumberAnimation {
                property: "y"
                duration: AppTheme.durStandard
                easing.type: Easing.OutCubic
            }
        }

        Repeater {
            model: toasts

            delegate: Item {
                id: toastWrap
                width: col.width
                height: leaving ? 0 : root.barHeight
                clip: true

                property bool leaving: false
                property bool risen: false

                Behavior on height {
                    NumberAnimation {
                        duration: 220
                        easing.type: Easing.InCubic
                    }
                }

                Timer {
                    id: lifeTimer
                    interval: model.kind === "error" ? 8000 : 3200
                    running: toastWrap.risen && !toastWrap.leaving && !toastHover.hovered
                    onTriggered: toastWrap.beginLeave()
                }

                function beginLeave() {
                    if (leaving) return
                    leaving = true
                    removeTimer.start()
                }

                Timer {
                    id: removeTimer
                    interval: 230
                    onTriggered: {
                        for (var i = 0; i < toasts.count; i++) {
                            if (toasts.get(i).mid === model.mid) {
                                toasts.remove(i)
                                return
                            }
                        }
                    }
                }

                Rectangle {
                    id: bar
                    width: parent.width - AppTheme.spaceM * 2
                    height: root.barHeight - AppTheme.spaceXS
                    x: AppTheme.spaceM
                    y: toastWrap.risen && !toastWrap.leaving ? 0 : root.barHeight
                    opacity: toastWrap.risen && !toastWrap.leaving ? 1 : 0
                    radius: AppTheme.radiusMedium
                    color: AppTheme.bgElevated
                    border.color: AppTheme.borderDivider
                    border.width: 1

                    Behavior on y {
                        NumberAnimation {
                            duration: 320
                            easing.type: Easing.OutCubic
                        }
                    }
                    Behavior on opacity {
                        NumberAnimation {
                            duration: 240
                            easing.type: AppTheme.easeColor
                        }
                    }

                    Component.onCompleted: Qt.callLater(function() { toastWrap.risen = true })

                    Rectangle {
                        anchors.fill: parent
                        radius: parent.radius
                        visible: pressArea.containsMouse
                        color: Qt.rgba(1, 1, 1, pressArea.pressed ? 0.10 : 0.05)
                    }

                    Row {
                        anchors.verticalCenter: parent.verticalCenter
                        anchors.left: parent.left
                        anchors.leftMargin: AppTheme.spaceM
                        anchors.right: parent.right
                        anchors.rightMargin: AppTheme.spaceM
                        spacing: AppTheme.spaceS

                        Rectangle {
                            width: 28
                            height: 28
                            radius: 14
                            anchors.verticalCenter: parent.verticalCenter
                            color: model.kind === "error"
                                   ? Qt.rgba(AppTheme.accentDanger.r, AppTheme.accentDanger.g, AppTheme.accentDanger.b, 0.18)
                                   : AppTheme.stateHover

                            IconImage {
                                anchors.centerIn: parent
                                source: model.kind === "error" ? "../icons/x_circle.svg" : "../icons/check.svg"
                                width: AppTheme.iconMedium
                                height: AppTheme.iconMedium
                                color: model.kind === "error" ? AppTheme.accentDanger : AppTheme.textPrimary
                            }
                        }

                        Text {
                            anchors.verticalCenter: parent.verticalCenter
                            width: parent.width - 28 - parent.spacing
                            elide: Text.ElideRight
                            wrapMode: Text.NoWrap
                            text: model.message
                            color: AppTheme.textPrimary
                            font.family: AppTheme.fontFamily
                            font.pixelSize: AppTheme.sizeBody
                            font.weight: AppTheme.weightBold
                        }
                    }

                    MouseArea {
                        id: pressArea
                        anchors.fill: parent
                        hoverEnabled: true
                        cursorShape: Qt.PointingHandCursor
                        onClicked: toastWrap.beginLeave()
                    }

                    HoverHandler { id: toastHover }
                }
            }
        }
    }
}
