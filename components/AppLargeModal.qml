import QtQuick
import QtQuick.Controls
import QtQuick.Layouts
import QtQuick.Controls.impl
import Qt5Compat.GraphicalEffects

// ============================================================
// БОЛЬШОЕ МОДАЛЬНОЕ ОКНО (настройки, печать): шапка с крупным
// заголовком и контент на всю ширину. Механика появления,
// затемнение и страховка — общие, в AppModalShell.
// ============================================================
AppModalShell {
    id: root
    property string title: "Заголовок"
    default property alias modalContent: contentArea.data
    anchors.centerIn: parent
    // Контент большой — не вылезает за скруглённые углы фона
    clipBackground: true
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
                    
                    HoverHandler { cursorShape: Qt.PointingHandCursor }
                    MouseArea { 
                        id: closeHov; 
                        anchors.fill: parent; 
                        hoverEnabled: true;
                        cursorShape: Qt.PointingHandCursor; 
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