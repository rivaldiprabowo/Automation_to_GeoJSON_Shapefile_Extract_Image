# This Python file uses the following encoding: utf-8
import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QPushButton, QHBoxLayout, QWidget
from PyQt6.QtCore import QThread, pyqtSignal
from ui.ui_main import Ui_Main

from src.converter_service import Process

class ConversionThread(QThread):
    progress = pyqtSignal(int)
    log_message = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, file_path=None, directory_path=None, out_directory_path=None, qml_folder=None, batas_wilayah_path=None) -> None:
        super().__init__()
        self.file_path = file_path
        self.directory_path = directory_path
        self.out_directory_path = out_directory_path
        self.qml_folder = qml_folder
        self.batas_wilayah_path = batas_wilayah_path
        self.running = True

    def run(self):
        processor = Process(self.out_directory_path, self.progress.emit)
        try:
            if self.file_path:
                self.log_message.emit(f"Processing file: {self.file_path}")
                processor.process_single_file(
                    self.file_path, 
                    self.log_callback,
                    qml_folder=self.qml_folder,
                    batas_wilayah_path=self.batas_wilayah_path
                )
            elif self.directory_path:
                self.log_message.emit(f"Processing directory: {self.directory_path}")
                processor.process_folder(
                    self.directory_path, 
                    self.log_callback,
                    qml_folder=self.qml_folder,
                    batas_wilayah_path=self.batas_wilayah_path
                )
            self.log_message.emit("Conversion completed!")
        except Exception as e:
            self.log_message.emit(f"Error: {str(e)}")
            import traceback
            self.log_message.emit(traceback.format_exc())
        finally:
            self.finished.emit()

    def log_callback(self, message):
        self.log_message.emit(message)

    def stop(self):
        self.running = False
        self.quit()
        self.wait()

class Main(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_Main()
        self.ui.setupUi(self)

        # Init progressBar value
        self.ui.progressBar.setValue(0)

        # Connect buttons to functions
        self.ui.btnSingleBrowsePath.clicked.connect(self.browse_single_file)
        self.ui.btnBulkBrowseDir.clicked.connect(self.browse_directory)
        self.ui.btnBulkOutDir.clicked.connect(self.browse_out_directory)
        self.ui.btnSingleOutDir.clicked.connect(self.browse_out_directory)
        self.ui.btnConvert.clicked.connect(self.start_conversion)
        self.ui.btnCancel.clicked.connect(self.cancel_conversion)
        self.ui.tabWidget.currentChanged.connect(self.on_tab_changed)

        # Initialize thread reference
        self.conversion_thread = None
        
        # Optional parameters - set to None by default
        self.qml_folder = None
        self.batas_wilayah_path = None
        
        # We'll use the existing UI without adding new buttons for now

    def browse_single_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xls *.xlsx)")
        if file_path:
            self.ui.singleFilePath.setText(file_path)

    def browse_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Directory")
        if directory:
            self.ui.bulkBrowseDir.setText(directory)

    def browse_out_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if directory:
            if self.ui.tabWidget.currentIndex() == 0:
                self.ui.singleOutDir.setText(directory)
            else:
                self.ui.bulkOutDir.setText(directory)

    def on_tab_changed(self, index):
        if index == 0:
            self.ui.bulkBrowseDir.clear()
            self.ui.bulkOutDir.clear()
        else:
            self.ui.singleFilePath.clear()
            self.ui.singleOutDir.clear()

    def start_conversion(self):
        file_path = self.ui.singleFilePath.text().strip()
        directory_path = self.ui.bulkBrowseDir.text().strip()
        if not file_path and not directory_path:
            QMessageBox.warning(self, "Error", "Please select a file or directory first.")
            return

        if self.ui.singleOutDir.text().strip() != "":
            out_directory_path = self.ui.singleOutDir.text().strip()
        elif self.ui.bulkOutDir.text().strip() != "":
            out_directory_path = self.ui.bulkOutDir.text().strip()
        else:
            QMessageBox.warning(self, "Error", "Please select output directory first.")
            return

        # Ask about optional parameters if not set
        if self.qml_folder is None:
            reply = QMessageBox.question(self, "QML Templates", 
                                         "Do you want to specify a QML templates folder?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.qml_folder = QFileDialog.getExistingDirectory(self, "Select QML Templates Directory")
                if self.qml_folder:
                    self.ui.textLog.append(f"QML folder set to: {self.qml_folder}")
        
        if self.batas_wilayah_path is None:
            reply = QMessageBox.question(self, "Boundary File", 
                                        "Do you want to specify a boundary shapefile (batas wilayah)?",
                                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.batas_wilayah_path, _ = QFileDialog.getOpenFileName(self, "Select Batas Wilayah Shapefile", "", "Shapefile (*.shp)")
                if self.batas_wilayah_path:
                    self.ui.textLog.append(f"Batas Wilayah file set to: {self.batas_wilayah_path}")

        self.ui.btnConvert.setEnabled(False)
        self.ui.btnCancel.setEnabled(True)
        self.ui.textLog.append("Starting conversion...")

        # Create and start the conversion thread with additional parameters
        self.conversion_thread = ConversionThread(
            file_path=file_path if file_path else None,
            directory_path=directory_path if directory_path else None,
            out_directory_path=out_directory_path,
            qml_folder=self.qml_folder,
            batas_wilayah_path=self.batas_wilayah_path
        )
        self.conversion_thread.progress.connect(self.ui.progressBar.setValue)
        self.conversion_thread.log_message.connect(self.ui.textLog.append)
        self.conversion_thread.finished.connect(self.conversion_finished)
        self.conversion_thread.start()

    def cancel_conversion(self):
        if self.conversion_thread:
            self.conversion_thread.stop()
            self.ui.textLog.append("Conversion canceled.")
            self.ui.btnConvert.setEnabled(True)
            self.ui.btnCancel.setEnabled(False)

    def conversion_finished(self):
        self.ui.textLog.append("Conversion finished.")
        self.ui.btnConvert.setEnabled(True)
        self.ui.btnCancel.setEnabled(False)
        self.ui.progressBar.setValue(0)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget = Main()
    widget.show()
    sys.exit(app.exec())