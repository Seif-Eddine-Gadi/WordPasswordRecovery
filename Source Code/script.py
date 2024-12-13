import sys
import os
import zipfile
import shutil
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QProgressBar, QFileDialog, QMessageBox, QWidget, QVBoxLayout
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon

class Worker(QThread):
    progress_updated = pyqtSignal(int)  # Signal to update progress bar
    finished = pyqtSignal()  # Signal when processing is finished

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def process_docx(self, file_path, processed_files, total_files):
        try:
            # Check if the file is read-only and make it writable
            if os.access(file_path, os.W_OK) == False:
                os.chmod(file_path, 0o666)  # Make the file writable

            # Rename .docx to .zip
            zip_file = file_path.replace('.docx', '.zip')
            os.rename(file_path, zip_file)

            # Now extract the zip file
            extract_folder = file_path.replace('.docx', '_contents')
            with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                zip_ref.extractall(extract_folder)

            # Find settings.xml in the extracted files
            settings_file = None
            for root, _, files in os.walk(extract_folder):
                for file in files:
                    if file == 'settings.xml':
                        settings_file = os.path.join(root, file)
                        break
                if settings_file:
                    break

            # Modify settings.xml if it exists
            if settings_file:
                with open(settings_file, 'r', encoding='utf-8') as file:
                    content = file.read()

                updated_content = content.replace('w:enforcement="1"', 'w:enforcement="0"')

                with open(settings_file, 'w', encoding='utf-8') as file:
                    file.write(updated_content)

            # Recreate the .docx file
            result_folder = os.path.join(os.path.dirname(file_path), "Result")
            if not os.path.exists(result_folder):
                os.makedirs(result_folder)

            result_docx = os.path.join(result_folder, os.path.basename(file_path))

            # Create a new zip file from the extracted files
            with zipfile.ZipFile(zip_file, 'w') as zip_ref:
                for root, _, files in os.walk(extract_folder):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, extract_folder)
                        zip_ref.write(full_path, arcname)

            # Rename .zip back to .docx and move to Result folder
            os.rename(zip_file, result_docx)

            # Cleanup
            shutil.rmtree(extract_folder)

            # Update progress bar
            processed_files += 1
            self.progress_updated.emit(int(processed_files / total_files * 100))

            return processed_files
        except Exception as e:
            QMessageBox.critical(None, "Error", f"Error processing {os.path.basename(file_path)}: {e}")
            return processed_files

    def run(self):
        try:
            # Process a single .docx file
            docx_file = self.file_path
            total_files = 1
            processed_files = 0

            processed_files = self.process_docx(docx_file, processed_files, total_files)

            self.progress_updated.emit(100)  # Ensure progress bar reaches 100%
            self.finished.emit()  # Emit finished signal

        except Exception as e:
            QMessageBox.critical(None, "Error", f"An error occurred: {e}")
            self.finished.emit()  # Emit finished signal even in case of error


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Microsoft Word Password Recovery")  # Set the window title
        self.setWindowIcon(QIcon("C:/Users/SXB/Desktop/Sifou Code/script.ico"))  # Set the icon for the window
        self.setGeometry(100, 100, 500, 300)

        self.file_path = ""

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # File selection
        self.file_label = QLabel("Select Word File (.docx):", self)
        layout.addWidget(self.file_label)

        self.file_entry = QLineEdit(self)
        self.file_entry.setReadOnly(True)
        layout.addWidget(self.file_entry)

        self.browse_button = QPushButton("Browse", self)
        self.browse_button.clicked.connect(self.select_file)
        layout.addWidget(self.browse_button)

        # Progress bar and Start button
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.start_button = QPushButton("Start", self)
        self.start_button.clicked.connect(self.start_processing)
        layout.addWidget(self.start_button)

        central_widget = QWidget(self)  # Create a central widget
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)  # Set the central widget

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select DOCX File", "", "Word Files (*.docx)")
        if file_path:
            self.file_path = file_path
            self.file_entry.setText(file_path)

    def start_processing(self):
        if not self.file_path:
            QMessageBox.critical(self, "Error", "Please select a DOCX file.")
            return

        self.progress_bar.setValue(0)
        self.worker = Worker(self.file_path)
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def update_progress(self, progress):
        self.progress_bar.setValue(progress)

    def on_finished(self):
        # Gracefully close the app
        QMessageBox.information(self, "Success", "Processing complete.")
        self.close()


if __name__ == '__main__':
    app = QApplication([])  # Create the application object
    window = MainWindow()  # Use MainWindow as the main window
    window.show()  # Show the window
    app.exec_()  # Start the event loop
