import sys
import webbrowser
import ctypes
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit,
    QPushButton, QFileDialog, QVBoxLayout, QHBoxLayout,
    QComboBox, QMessageBox, QFrame, QProgressBar
)
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("OpenConvert.App")

# ================= WORKER THREAD =================
class ConvertWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, path, fmt, save_dir):
        super().__init__()
        self.path = path
        self.fmt = fmt
        self.save_dir = save_dir

    def run(self):
        try:
            for i in range(0, 70, 10):
                self.progress.emit(i)
                self.msleep(120)

            from convertor import convert_file
            output = convert_file(self.path, self.fmt, self.save_dir)

            self.progress.emit(100)
            self.finished.emit(output)
        except Exception as e:
            self.error.emit(str(e))


# ================= MAIN UI =================
class OpenConvert(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OpenConvert")
        self.setFixedSize(620, 540)
        self.setWindowIcon(QIcon("icons/app.png"))
        self.setStyleSheet(self.load_styles())
        self.init_ui()

    def init_ui(self):
        # -------- Header --------
        title = QLabel("OpenConvert")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setObjectName("title")

        subtitle = QLabel("Simple • Fast • Open Source")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setObjectName("subtitle")

        formats = QLabel(
            "Images: PNG, JPG, JPEG, WEBP, BMP, TIFF\n"
            "Documents → PDF: TXT, DOC, DOCX, RTF, ODT, PPT, PPTX, XLS, XLSX, CSV, HTML, MD"
        )
        formats.setAlignment(Qt.AlignmentFlag.AlignCenter)
        formats.setWordWrap(True)
        formats.setObjectName("formats")

        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.HLine)
        divider.setFrameShadow(QFrame.Shadow.Sunken)

        # -------- File Input --------
        self.file_input = QLineEdit()
        self.file_input.setPlaceholderText("Select a file to convert...")
        self.file_input.setMinimumHeight(36)

        browse_btn = QPushButton(" Browse")
        browse_btn.setIcon(QIcon("icons/browse.png"))
        browse_btn.setIconSize(QSize(18, 18))
        browse_btn.setFixedSize(110, 36)
        browse_btn.setToolTip("Select input file")
        browse_btn.clicked.connect(self.browse_file)

        file_layout = QHBoxLayout()
        file_layout.setSpacing(12)
        file_layout.addWidget(self.file_input, stretch=1)
        file_layout.addWidget(browse_btn)

        # -------- Save Location --------
        self.save_input = QLineEdit()
        self.save_input.setPlaceholderText("Save location (default: same as input folder)")
        self.save_input.setMinimumHeight(36)

        save_btn = QPushButton(" Save")
        save_btn.setIcon(QIcon("icons/save.png"))
        save_btn.setIconSize(QSize(18, 18))
        save_btn.setFixedSize(110, 36)
        save_btn.setToolTip("Choose output folder")
        save_btn.clicked.connect(self.choose_save_folder)

        save_layout = QHBoxLayout()
        save_layout.setSpacing(12)
        save_layout.addWidget(self.save_input, stretch=1)
        save_layout.addWidget(save_btn)

        # -------- Output Format --------
        format_label = QLabel("Output Format")
        format_label.setObjectName("section")

        self.format_box = QComboBox()
        self.format_box.setMinimumHeight(34)
        self.format_box.addItems([
            "PDF",
            "PNG", "JPG", "JPEG", "WEBP", "BMP", "TIFF"
        ])

        # -------- Progress --------
        self.progress = QProgressBar()
        self.progress.setFixedHeight(20)
        self.progress.setValue(0)
        self.progress.setFormat("Ready")

        # -------- Convert Button --------
        self.convert_btn = QPushButton(" Convert File")
        self.convert_btn.setIcon(QIcon("icons/convert.png"))
        self.convert_btn.setIconSize(QSize(22, 22))
        self.convert_btn.setFixedHeight(42)
        self.convert_btn.setEnabled(False)
        self.convert_btn.setToolTip("Start conversion")
        self.convert_btn.clicked.connect(self.convert_now)

        # -------- Footer --------
        footer_layout = QHBoxLayout()
        footer_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        footer_layout.setSpacing(14)

        github_btn = QPushButton()
        github_btn.setIcon(QIcon("icons/github.png"))
        github_btn.setFixedSize(28, 28)
        github_btn.setToolTip("Contribute on GitHub")
        github_btn.setFlat(True)
        github_btn.clicked.connect(self.open_contribute)

        about_btn = QPushButton()
        about_btn.setIcon(QIcon("icons/info.png"))
        about_btn.setFixedSize(28, 28)
        about_btn.setToolTip("About OpenConvert")
        about_btn.setFlat(True)
        about_btn.clicked.connect(self.show_about)

        footer_layout.addWidget(github_btn)
        footer_layout.addWidget(about_btn)

        # -------- Main Layout --------
        layout = QVBoxLayout()
        layout.setContentsMargins(32, 26, 32, 26)
        layout.setSpacing(18)

        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(formats)
        layout.addWidget(divider)

        layout.addLayout(file_layout)
        layout.addLayout(save_layout)

        layout.addSpacing(6)
        layout.addWidget(format_label)
        layout.addWidget(self.format_box)

        layout.addSpacing(8)
        layout.addWidget(self.progress)
        layout.addSpacing(12)
        layout.addWidget(self.convert_btn)

        layout.addStretch()
        layout.addLayout(footer_layout)

        self.setLayout(layout)

    # ================= Logic =================
    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select File")
        if path:
            self.file_input.setText(path)
            self.convert_btn.setEnabled(True)

    def choose_save_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.save_input.setText(folder)

    def convert_now(self):
        path = self.file_input.text()
        fmt = self.format_box.currentText()
        save_dir = self.save_input.text().strip()

        self.progress.setValue(0)
        self.progress.setFormat("Converting...")
        self.convert_btn.setEnabled(False)

        self.worker = ConvertWorker(path, fmt, save_dir)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.conversion_finished)
        self.worker.error.connect(self.conversion_error)
        self.worker.start()

    def conversion_finished(self, output):
        self.progress.setValue(100)
        self.progress.setFormat("Completed")
        self.convert_btn.setEnabled(True)
        QMessageBox.information(self, "Success", f"Saved at:\n{output}")

    def conversion_error(self, msg):
        self.progress.setValue(0)
        self.progress.setFormat("Failed")
        self.convert_btn.setEnabled(True)
        QMessageBox.critical(self, "Error", msg)

    # -------- Footer --------
    def open_contribute(self):
        webbrowser.open("https://github.com/yourusername/OpenConvert")

    def show_about(self):
        QMessageBox.information(
            self,
            "About OpenConvert",
            "OpenConvert\n\n"
            "Multi-format file converter\n"
            "Built with Python & PyQt\n\n"
            "Images: PNG, JPG, JPEG, WEBP, BMP, TIFF\n"
            "Documents → PDF: TXT, DOC, DOCX, RTF, ODT, PPT, PPTX, XLS, XLSX, CSV, HTML, MD"
        )

    # ================= Styles =================
    def load_styles(self):
        return """
        QWidget {
            background-color: #1f1f1f;
            color: #ffffff;
            font-family: Segoe UI;
            font-size: 13px;
        }

        QLabel#title {
            font-size: 24px;
            font-weight: bold;
            color: #4fc3f7;
        }

        QLabel#subtitle {
            font-size: 12px;
            color: #bbbbbb;
        }

        QLabel#formats {
            font-size: 10px;
            color: #888888;
        }

        QLabel#section {
            font-size: 13px;
            font-weight: bold;
        }

        QLineEdit, QComboBox {
            background-color: #2b2b2b;
            border: 1px solid #444;
            border-radius: 8px;
            padding: 6px;
        }

        QProgressBar {
            border: 1px solid #444;
            border-radius: 8px;
            text-align: center;
            background: #2b2b2b;
        }

        QProgressBar::chunk {
            background-color: #4fc3f7;
            border-radius: 8px;
        }

        QPushButton {
            background-color: qlineargradient(
                x1:0, y1:0, x2:0, y2:1,
                stop:0 #5fd3ff,
                stop:1 #39b8f2
            );
            color: #000;
            border-radius: 10px;
            padding: 6px 14px;
            font-weight: bold;
            border: none;
        }

        QPushButton:hover {
            background-color: qlineargradient(
                x1:0, y1:0, x2:0, y2:1,
                stop:0 #78ddff,
                stop:1 #52c6f7
            );
        }

        QPushButton:pressed {
            background-color: #2fa6da;
        }

        QPushButton:disabled {
            background-color: #444;
            color: #888;
        }
        """
import ctypes
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("OpenConvert.App")

app = QApplication(sys.argv)
app.setWindowIcon(QIcon("icons/app.png"))

window = OpenConvert()
window.show()
sys.exit(app.exec())
