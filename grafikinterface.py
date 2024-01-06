import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QListWidget, QPushButton, QListWidgetItem
from PyQt5 import QtCore
from PyQt5.QtCore import Qt

class FileDropperWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls:
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
            links = []
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    links.append(str(url.toLocalFile()))
                else:
                    links.append(str(url.toString()))
                self.addItems(links)
        else:
            event.ignore

class Kalkulationstool(QWidget):
    def __init__(self, kalk_funktion):
        super().__init__()

        self.setWindowTitle("Kalkulationstool")
        self.setGeometry(100, 100, 600, 400)  # Anfangsposition und Größe des Fensters

        self.initUI(kalk_funktion)

    def initUI(self, kalk_funktion):
        # Hauptlayout
        layout = QVBoxLayout()
        # Überschrift
        label = QLabel("Kalkulationstool")
        label.setAlignment(QtCore.Qt.AlignTop | QtCore.Qt.AlignHCenter)
        label.setStyleSheet("font-weight: bold; font-size: 16px;")
        layout.addWidget(label)
        # Horizontales Layout für die beiden Bereiche
        horizontal_layout = QHBoxLayout()

        # Linker Bereich mit QListWidget
        list_widget = FileDropperWidget()
        horizontal_layout.addWidget(list_widget)

        # Rechter Bereich mit dem Button
        button_layout = QVBoxLayout()
        button_layout.addStretch()  # Füge Strech aus, um den Button nach unten zu zentrieren
        def start_kalkulation(self):
            # Hier kannst du die Funktion hinzufügen, die beim Klick auf den Button ausgeführt werden soll
            items = [list_widget.item(i).text() for i in range(list_widget.count())]
            kalk_funktion(items)
        start_button = QPushButton("Kalkulation starten")
        start_button.clicked.connect(start_kalkulation)
        button_layout.addWidget(start_button, alignment=QtCore.Qt.AlignCenter)
        horizontal_layout.addLayout(button_layout)

        # Füge das horizontale Layout zum Hauptlayout hinzu
        layout.addLayout(horizontal_layout)

        # Setze das Hauptlayout für das Fenster
        self.setLayout(layout)

def eingabefenster_oeffnen(kalk_funktion):
    app = QApplication(sys.argv)
    window = Kalkulationstool(kalk_funktion)
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    def kalk_funktion(items : list[QListWidgetItem]):
        for item in items:
            print("Kalkulation gestartet für ", item)
    eingabefenster_oeffnen(kalk_funktion)