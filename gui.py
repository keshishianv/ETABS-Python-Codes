import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout, QMessageBox, QComboBox
from main import run


class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'ETabs Data'  # Set the title here
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(100, 100, 300, 150)

        self.button = QPushButton('Save As', self)
        self.button.clicked.connect(self.showDialog)

        # Create a dropdown box for unit selection
        self.unitComboBox = QComboBox(self)
        units = [
            "lb_in_F", "lb_ft_F", "kip_in_F", "kip_ft_F",
            "kN_mm_C", "kN_m_C", "kgf_mm_C", "kgf_m_C",
            "N_mm_C", "N_m_C", "Ton_mm_C", "Ton_m_C",
            "kN_cm_C", "kgf_cm_C", "N_cm_C", "Ton_cm_C"
        ]
        self.unitComboBox.addItems(units)

        # Set the default unit index to 4 (kip_ft_F)
        self.unitComboBox.setCurrentIndex(3)  # Index 3 corresponds to "kip_ft_F"

        layout = QVBoxLayout()
        layout.addWidget(self.unitComboBox)
        layout.addWidget(self.button)
        self.setLayout(layout)

        self.show()

    def showDialog(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(
            self,
            "Save As",
            "ETABS DATA",
            "Excel Files (*.xlsx);;All Files (*)",
            options=options
        )
        if fileName:
            if self.is_file_open(fileName):
                QMessageBox.warning(self, "Error",
                                    "The selected file is currently open. Please close it and try again.")
                return

            # Get the selected unit and map it to the corresponding number
            unit_index = self.unitComboBox.currentIndex() + 1  # +1 to match the numbering
            print("Selected file:", fileName)
            print("Selected unit index:", unit_index)

            # Call the run() function from main.py with the selected path and unit index
            run(fileName, unit_index)
            QMessageBox.information(self, "Completed", "Completed!")

    def is_file_open(self, file_path):
        """Check if the file is open by trying to rename it and return False if the file does not exist."""
        if not os.path.exists(file_path):
            return False  # File doesn't exist, return False

        try:
            # Attempt to rename the file
            os.rename(file_path, file_path)
            return False  # File is not open
        except OSError:
            return True  # File is open or locked


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
