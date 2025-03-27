import sys
from PyQt5.QtWidgets import QApplication, QFileDialog, QWidget

def browse_file():
    app = QApplication(sys.argv)
    widget = QWidget()

    # Open de file dialog
    filepath, _ = QFileDialog.getOpenFileName(widget, "Selecteer een bestand", "", "Excel bestanden (*.xlsx *.xls)")
    
    if filepath:
        print(f"Gekozen bestand: {filepath}")
    else:
        print("Geen bestand geselecteerd.")

browse_file()
