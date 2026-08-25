import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

Popup {
    id: root
    
    property string title: "Заголовок"
    default property alias modalContent: contentArea.data

    anchors.centerIn: parent
    
    modal: true
    dim: true 
    focus: true
    closePolicy: Popup.CloseOnEscape 
    
    z: AppTheme.zModal // Правильный Z-индекс (9000)

    // ==========================================
    // 1. ЗАТЕМНЕНИЕ ФОНА (Overlay)
    // ==========================================
    Overlay.modal: Rectangle {
        color: AppTheme.bgOverlay
        opacity: root.opened ? 1.0 : 0.0
        Behavior on opacity { 
            NumberAnimation { 
                duration: AppTheme.durStandard 
                easing.type: root.opened ? AppTheme.easeEnter : AppTheme.easeExit 
            } 
        }
    }

    // ==========================================
    // 2. АНИМАЦИИ (Premium Pop Effect)
    // ==========================================
    enter: Transition {
        ParallelAnimation {
            // Прозрачность появляется очень быстро (за 150мс), убивая эффект "призрака"
            NumberAnimation { property: "opacity"; from: 0.0; to: 1.0; duration: 150; easing.type: Easing.OutQuad }
            // Масштаб летит плавно и красиво (за 250мс)
            NumberAnimation { property: "scale"; from: 0.85; to: 1.0; duration: AppTheme.durStandard; easing.type: AppTheme.easeEnter }
        }
    }
    exit: Transition {
        ParallelAnimation {
            // Исчезает тоже моментально
            NumberAnimation { property: "opacity"; from: 1.0; to: 0.0; duration: 100; easing.type: Easing.InQuad }
            NumberAnimation { property: "scale"; from: 1.0; to: 0.95; duration: 150; easing.type: AppTheme.easeExit }
        }
    }

    // ==========================================
    // 3. ФОН И ТЕНЬ (Level 4)
    // ==========================================
    background: Rectangle {
        color: AppTheme.bgModal 
        radius: AppTheme.radiusModal
        
        border.color: AppTheme.borderDivider
        border.width: 1
        clip: true
        AppShadow { level: 4 }
    }

    // ==========================================
    // 4. КОНТЕНТ (Шапка + Тело)
    // ==========================================
    contentItem: ColumnLayout {
        anchors.fill: parent
        spacing: 0 

        // ШАПКА ОКНА
        Item {
            Layout.fillWidth: true
            Layout.preferredHeight: 60 // Жесткая высота шапки
            
            RowLayout {
                anchors.fill: parent
                anchors.leftMargin: AppTheme.spaceL
                anchors.rightMargin: AppTheme.spaceL
                
                Text {
                    Layout.fillWidth: true
                    text: root.title
                    color: AppTheme.textPrimary
                    font.family: AppTheme.fontFamily
                    font.pixelSize: AppTheme.sizeH2 
                    font.weight: AppTheme.weightBold // Жирный заголовок H2
                }
                
                // КРЕСТИК ЗАКРЫТИЯ
                Rectangle {
                    Layout.preferredWidth: 32; Layout.preferredHeight: 32; radius: AppTheme.radiusPill
                    
                    // Системный Hover
                    color: closeHov.pressed ? AppTheme.statePress : (closeHov.containsMouse ? AppTheme.stateHover : "transparent")
                    Behavior on color { ColorAnimation { duration: AppTheme.durMicro } }
                    
                    IconImage { 
                        anchors.centerIn: parent; 
                        source: "../icons/close.svg"; 
                        width: AppTheme.iconSmall; 
                        height: AppTheme.iconSmall; 
                        color: AppTheme.textSecondary 
                    }
                    
                    MouseArea { 
                        id: closeHov; 
                        anchors.fill: parent; 
                        hoverEnabled: true; 
                        onClicked: root.close() 
                    }
                }
            }
            
            // Тонкая линия отбивки шапки от контента
            Rectangle {
                anchors.bottom: parent.bottom
                width: parent.width
                height: 1
                color: AppTheme.borderDivider
            }
        }

        // КОНТЕНТ ОКНА (Занимает всё пространство под шапкой)
        Item {
            id: contentArea
            Layout.fillWidth: true
            Layout.fillHeight: true
        }
    }

    function show() {
        root.open()
    }
}