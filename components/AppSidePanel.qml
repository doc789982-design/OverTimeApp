import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

Popup {
    id: root

    property string title: "Панель"
    default property alias panelContent: contentArea.data

    // Режим «морфинга»: панель открывается сразу в заданном месте без выезда сбоку.
    property bool morphOpen: false

    width: 450 
    height: ApplicationWindow.window ? ApplicationWindow.window.height : 800
    
    // Прилипаем к правому краю экрана
    x: ApplicationWindow.window ? ApplicationWindow.window.width - width : 0
    y: 0
    
    z: AppTheme.zModal // Правильный слой

    modal: true

    // ==========================================
    // 1. ЗАТЕМНЕНИЕ ФОНА — убрали (окно просто открывается, без затемнения)
    // ==========================================

    // ==========================================
    // 2. АНИМАЦИИ ВЫЕЗДА (Motion System)
    // ==========================================
    enter: Transition {
        // При морфинге from == to — панель не выезжает сбоку, а остаётся на месте
        NumberAnimation { 
            property: "x"
            from: root.morphOpen ? root.x : root.x + root.width
            to: root.x
            duration: AppTheme.durStandard
            easing.type: AppTheme.easeEnter 
        }
    }
    exit: Transition {
        NumberAnimation { 
            property: "x"
            from: root.x
            to: root.morphOpen ? root.x : root.x + root.width
            duration: AppTheme.durFast // Уезжает быстрее, чем выезжает
            easing.type: AppTheme.easeExit 
        }
    }

    // ==========================================
    // 3. ФОН И ТЕНЬ (Surfaces & Elevation)
    // ==========================================
    background: Rectangle {
        color: AppTheme.bgModal
        
        // Скругляем все углы...
        radius: AppTheme.radiusModal
        // ...но заклеиваем правые углы квадратной заплаткой (Sharp)
        Rectangle { 
            width: AppTheme.radiusModal
            height: parent.height
            anchors.right: parent.right
            color: AppTheme.bgModal 
        }
        
        border.color: AppTheme.borderDivider
        border.width: 1

        // Тень-картинка вместо вычисляемой (Level 4)
        AppShadow { level: 4 }
    }

    // ==========================================
    // 4. КОНТЕНТ
    // ==========================================
    contentItem: ColumnLayout {
        anchors.fill: parent
        anchors.margins: AppTheme.spaceL
        spacing: AppTheme.spaceL

        // ШАПКА
        RowLayout {
            Layout.fillWidth: true
            
            Text {
                Layout.fillWidth: true
                text: root.title
                color: AppTheme.textPrimary
                font.family: AppTheme.fontFamily
                font.pixelSize: AppTheme.sizeH2
                font.weight: AppTheme.weightBold // Идеальный жирный текст
            }
            
            // КРЕСТИК ЗАКРЫТИЯ
            Rectangle {
                Layout.preferredWidth: 24
                Layout.preferredHeight: 24
                radius: AppTheme.radiusPill
                
                color: closeHov.pressed ? AppTheme.statePress : (closeHov.containsMouse ? AppTheme.stateHover : "transparent")
                Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                
                IconImage { 
                    anchors.centerIn: parent
                    source: "../icons/close.svg"
                    width: AppTheme.iconSmall
                    height: AppTheme.iconSmall
                    color: AppTheme.textSecondary 
                }
                
                MouseArea { 
                    id: closeHov
                    anchors.fill: parent
                    hoverEnabled: true
                    onClicked: root.close() 
                }
            }
        }

        // КОНТЕНТ (Скроллируемый)
        ScrollView {
            Layout.fillWidth: true
            Layout.fillHeight: true
            clip: true
            ScrollBar.horizontal.policy: ScrollBar.AlwaysOff
            
            Column {
                id: contentArea
                width: parent.width
                spacing: AppTheme.spaceL
            }
        }
    }
    
    function show() {
        if (ApplicationWindow.window) {
            root.height = ApplicationWindow.window.height
            root.x = ApplicationWindow.window.width - root.width
        }
        root.open()
    }

    // Открытие после «морфинга»: панель сразу встаёт в заданный прямоугольник.
    function openMorph(x, y, w, h) {
        root.morphOpen = true
        root.x = x
        root.y = y
        root.width = w
        root.height = h
        root.open()
    }
}