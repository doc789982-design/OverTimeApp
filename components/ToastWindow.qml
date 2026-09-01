import QtQuick
import QtQuick.Window

// Отдельное всегда-поверх-всего окно для тостов.
//
// Зачем: тосты раньше жили внутри главного окна, поэтому модальные окна
// (настройки, печать и т.п.) перекрывали их. Это окно — собственное native
// окно с флагом WindowStaysOnTopHint: оно всплывает ПОВЕРХ любого окна —
// и модальных диалогов программы, и даже чужих приложений. Системно, без
// правок под каждый конкретный диалог.
Window {
    id: toastWindow

    // Без рамки, не показывается в панели задач, не забирает фокус.
    flags: Qt.FramelessWindowHint
           | Qt.Tool
           | Qt.WindowStaysOnTopHint
           | Qt.WindowDoesNotAcceptFocus
    color: "transparent"
    transparent: true

    // Держимся внизу по центру главного окна и гаснем вместе с ним.
    width: Math.min(520, Math.max(300, mainWindow.width * 0.5))
    height: toasts.implicitHeight
    x: mainWindow.x + Math.round((mainWindow.width - width) / 2)
    y: mainWindow.y + mainWindow.height - height - 12
    visible: mainWindow.visible && mainWindow.width > 0

    ToastHost {
        id: toasts
        anchors.left: parent.left
        anchors.right: parent.right
        anchors.bottom: parent.bottom
    }
}
