import os
from openpyxl import load_workbook
from PySide6.QtWidgets import QApplication, QFileDialog, QLabel, QVBoxLayout, QPushButton, QWidget
from PySide6.QtCore import Qt
from PySide6.QtGui import QPalette, QColor


class ExcelUpdaterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Atualizador de Excel")
        self.setFixedSize(400, 300)
        
        # Estilo translúcido e bordas arredondadas
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setStyleSheet("""
            QWidget {
                background-color: rgba(255, 255, 255, 180);
                border-radius: 15px;
            }
            QPushButton {
                background-color: #0078D7;
                color: white;
                border-radius: 10px;
                padding: 5px;
            }
            QPushButton:hover {
                background-color: #005A9E;
            }
        """)

        # Layout e widgets
        layout = QVBoxLayout(self)

        label = QLabel("Selecione a pasta para atualizar arquivos Excel:")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)

        self.select_folder_btn = QPushButton("Selecionar Pasta")
        self.select_folder_btn.clicked.connect(self.select_folder)
        layout.addWidget(self.select_folder_btn)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        self.update_btn = QPushButton("Atualizar Arquivos")
        self.update_btn.clicked.connect(self.update_excel_files)
        self.update_btn.setEnabled(False)
        layout.addWidget(self.update_btn)

        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(self.close_program)
        layout.addWidget(close_btn)

        self.folder_path = ""

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Selecione uma pasta")
        if folder:
            self.folder_path = folder
            self.status_label.setText(f"Pasta selecionada: {folder}")
            self.update_btn.setEnabled(True)

    def update_excel_files(self):
        if not self.folder_path:
            self.status_label.setText("Nenhuma pasta selecionada.")
            return

        try:
            count = 0
            for root, _, files in os.walk(self.folder_path):
                for file in files:
                    if file.endswith(".xlsx"):
                        filepath = os.path.join(root, file)
                        wb = load_workbook(filepath)
                        wb.save(filepath)
                        wb.close()
                        count += 1

            self.status_label.setText(f"{count} arquivos atualizados com sucesso!")
        except Exception as e:
            self.status_label.setText(f"Erro: {str(e)}")

    def close_program(self):
        QApplication.quit()


if __name__ == "__main__":
    app = QApplication([])

    # Personalização global
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(255, 255, 255, 180))
    app.setPalette(palette)

    window = ExcelUpdaterApp()
    window.show()

    app.exec()

# %%
