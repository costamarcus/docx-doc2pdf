from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QProgressBar, QMessageBox
from PyQt5.QtGui import QFont
from docx2pdf import convert
from tkinter import Tk


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Conversor de Word para PDF')
        self.setGeometry(100, 100, 400, 150)

        self.label_word = QLabel('Selecione os arquivos Word:', self)
        self.label_word.move(20, 20)
        self.label_word.setFont(QFont('Arial', 10))

        self.button_word = QPushButton('Selecionar', self)
        self.button_word.move(200, 20)
        self.button_word.clicked.connect(self.get_word_files)

        self.button_start = QPushButton('Iniciar', self)
        self.button_start.move(150, 70)
        self.button_start.clicked.connect(self.start_conversion)

        self.word_files = []

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setGeometry(20, 110, 360, 20)

    def get_word_files(self):
        filenames, _ = QFileDialog.getOpenFileNames(
            self, 'Selecione os arquivos Word', '', 'Word Files (*.docx *.doc)')
        self.word_files = filenames

    def start_conversion(self):
        if not self.word_files:
            QMessageBox.warning(self, 'Nenhum arquivo selecionado',
                                'Por favor, selecione um ou mais arquivos Word para converter.')
            self.button_word.setFocus()
            return

        total_files = len(self.word_files)
        self.progress_bar.setMaximum(total_files)
        self.progress_bar.setValue(0)

        for i, word_file in enumerate(self.word_files):
            pdf_file = word_file.replace('.docx', '.pdf')
            convert(word_file, pdf_file)
            self.progress_bar.setValue(i + 1)

        QMessageBox.information(self, 'Conversão concluída',
                                f'{total_files} arquivos foram convertidos com sucesso.')
        self.close()


app = QApplication([])
window = MainWindow()
window.show()
app.exec_()
