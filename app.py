import sys
import os
import subprocess

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QProgressBar
)
from PySide6.QtCore import QObject, Signal, QThread, Qt, QElapsedTimer


class OCRWorker(QObject):
    progress = Signal(int, int, str)  # current, total, filename
    finished = Signal(dict)
    failed = Signal(str)

    def __init__(self, folder_path, ocr_out, poppler_bin, tesseract_exe, tessdata_dir, dpi=200, first_page_only=True):
        super().__init__()
        self.folder_path = folder_path
        self.ocr_out = ocr_out
        self.poppler_bin = poppler_bin
        self.tesseract_exe = tesseract_exe
        self.tessdata_dir = tessdata_dir
        self.dpi = dpi
        self.first_page_only = first_page_only

    def run(self):
        try:
            # ważne: env dla tesseract języków
            os.environ["TESSDATA_PREFIX"] = self.tessdata_dir

            from ocr_engine import ocr_folder_pdfs

            def cb(filename, current, total):
                # tylko emit sygnału (ZERO UI tutaj)
              #  print("CB:", filename, current, total)
                self.progress.emit(int(current), int(total), str(filename))

            stats = ocr_folder_pdfs(
                pdf_folder=self.folder_path,
                out_txt_folder=self.ocr_out,
                poppler_path=self.poppler_bin,
                tesseract_cmd=self.tesseract_exe,
                dpi=self.dpi,
                first_page_only=self.first_page_only,
                on_progress=cb,
            )
            self.finished.emit(stats)
        except Exception as e:
            self.failed.emit(repr(e))


class FakturyApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Generator Faktur")
        self.setMinimumWidth(520)

        self.folder_path = ""
        self.output_dir = ""

        self.app_dir = os.path.dirname(os.path.abspath(__file__))

        # portable tools
        self.poppler_bin = os.path.join(self.app_dir, "tools", "poppler", "bin")
        self.tesseract_dir = os.path.join(self.app_dir, "tools", "tesseract")
        self.tesseract_exe = os.path.join(self.tesseract_dir, "tesseract.exe")
        self.tessdata_dir = os.path.join(self.tesseract_dir, "tessdata")

        layout = QVBoxLayout()

        self.label = QLabel("Nie wybrano folderu")
        layout.addWidget(self.label)

        self.btn_choose = QPushButton("Wybierz folder z fakturami")
        self.btn_choose.clicked.connect(self.choose_folder)
        layout.addWidget(self.btn_choose)

        self.btn_choose_out = QPushButton("Wybierz folder wyników (.xlsx)")
        self.btn_choose_out.clicked.connect(self.choose_output_folder)
        layout.addWidget(self.btn_choose_out)

        # 2 nowe przyciski
        self.btn_ocr = QPushButton("OCR")
        self.btn_ocr.setEnabled(False)
        self.btn_ocr.clicked.connect(self.run_ocr)

        self.btn_xlsx = QPushButton("Generuj xlsx")
        self.btn_xlsx.setEnabled(True)
        self.btn_xlsx.clicked.connect(self.run_xlsx)

        self.btn_ocr.setVisible(False)
        self.btn_xlsx.setVisible(False)

        # główny przycisk (zostaje na dole) = OCR + xlsx
        self.btn_main = QPushButton("OCR + xlsx")
        self.btn_main.setEnabled(False)
        self.btn_main.clicked.connect(self.run_ocr_and_xlsx)

        layout.addWidget(self.btn_ocr)
        layout.addWidget(self.btn_xlsx)
        layout.addWidget(self.btn_main)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.setLayout(layout)

        # throttling UI update (żeby nie spamować repaint)
        self._ui_timer = QElapsedTimer()
        self._ui_timer.start()
        self._last_progress = -1

        # thread refs
        self.thread = None
        self.worker = None
        self._ocr_stats = None
        self._ocr_error = None
        self._after_ocr = None

    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Wybierz folder z PDF-ami")
        if folder:
            self.folder_path = folder
            self.label.setText(f"Wybrany folder:\n{folder}")
            self.btn_ocr.setEnabled(True)
            self.btn_xlsx.setEnabled(True)
            self.update_buttons_state()

    def choose_output_folder(self):
        out = QFileDialog.getExistingDirectory(self, "Wybierz folder wyników (.xlsx)")
        if out:
            self.output_dir = out
            self.label.setText(f"Wybrany folder:\n{self.folder_path}\n\nFolder wyników:\n{out}")
            self.update_buttons_state()

    def update_buttons_state(self):
        ready = bool(self.folder_path) and bool(self.output_dir)
        self.btn_main.setEnabled(ready)

    def update_counter(self, current, total, filename):
        self.label.setText(
            f"OCR w toku...\n"
            f"Zrobione: {current}/{total}\n"
            f"Plik: {os.path.basename(filename)}"
        )

    def run_ocr(self):
        if not self.folder_path:
            QMessageBox.warning(self, "Błąd", "Nie wybrano folderu z fakturami")
            return

        if not self.output_dir:
            QMessageBox.warning(self, "Błąd", "Nie wybrano folderu wyników (.xlsx)")
            return

        pdf_files = [f for f in os.listdir(self.folder_path) if f.lower().endswith(".pdf")]
        if not pdf_files:
            QMessageBox.warning(self, "Brak plików", "W wybranym folderze nie ma plików PDF")
            return

        # output TXT
        ocr_out = os.path.join(self.folder_path, "ocr_txt")
        os.makedirs(ocr_out, exist_ok=True)

        # poppler exe
        pdftoppm = os.path.join(self.poppler_bin, "pdftoppm.exe")
        pdfinfo = os.path.join(self.poppler_bin, "pdfinfo.exe")

        if not os.path.exists(pdftoppm) or not os.path.exists(pdfinfo):
            QMessageBox.critical(
                self,
                "Brak Popplera",
                f"Nie znaleziono Popplera w:\n{self.poppler_bin}\n\n"
                f"Wymagane pliki:\n- pdfinfo.exe\n- pdftoppm.exe"
            )
            return

        # czy pdfinfo uruchamia się
        try:
            subprocess.run([pdfinfo, "-h"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        except Exception as e:
            QMessageBox.critical(
                self,
                "Poppler nie działa",
                "pdfinfo.exe jest w folderze, ale nie uruchamia się.\n"
                "Najczęstsza przyczyna: brak Microsoft Visual C++ Redistributable.\n\n"
                f"Szczegóły: {repr(e)}"
            )
            return

        if not os.path.exists(self.tesseract_exe):
            QMessageBox.critical(self, "Brak Tesseracta", f"Nie znaleziono tesseract.exe w:\n{self.tesseract_exe}")
            return

        if not os.path.isdir(self.tessdata_dir):
            QMessageBox.critical(self, "Brak tessdata", f"Nie znaleziono folderu tessdata:\n{self.tessdata_dir}")
            return

        # reset buffers
        self._ocr_stats = None
        self._ocr_error = None
        self._ui_timer.restart()
        self._last_progress = -1

        # Thread + worker
        self.thread = QThread()
        self.worker = OCRWorker(
            folder_path=self.folder_path,
            ocr_out=ocr_out,
            poppler_bin=self.poppler_bin,
            tesseract_exe=self.tesseract_exe,
            tessdata_dir=self.tessdata_dir,
            dpi=200,
            first_page_only=True
        )
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)

        def on_progress(current, total, filename):
            self.label.setText(
                f"OCR w toku...\n"
                f"Zrobione: {current}/{total}\n"
                f"Plik: {filename}"
            )

        def on_worker_finished(stats):
            self._ocr_stats = stats
            self.thread.quit()

        def on_worker_failed(msg):
            self._ocr_error = msg
            self.thread.quit()

        def on_thread_finished():
            # tu jesteśmy już w GUI thread
            self.btn_choose.setEnabled(True)
            self.btn_ocr.setEnabled(True)
            self.btn_xlsx.setEnabled(True)
            self.btn_main.setEnabled(True)

            self.progress.setVisible(False)

            if self._ocr_error:
                QMessageBox.critical(self, "Błąd OCR", self._ocr_error)
                return

            stats = self._ocr_stats or {"total": 0, "skipped": 0, "done": 0, "errors": 0}
            QMessageBox.information(
                self,
                "OCR zakończony",
                f"PDF: {stats['total']}\n"
                f"Pominięte (już miały TXT): {stats['skipped']}\n"
                f"Nowo zrobione OCR: {stats['done']}\n"
                f"Błędy: {stats['errors']}\n\n"
                f"TXT zapisane w:\n{ocr_out}"
            )

            if self._after_ocr == "xlsx":
                self._after_ocr = None
                self.run_xlsx()


        # queued connections = brak akcji UI z workera
        self.worker.progress.connect(self.update_counter, Qt.QueuedConnection)
        self.worker.finished.connect(on_worker_finished, Qt.QueuedConnection)
        self.worker.failed.connect(on_worker_failed, Qt.QueuedConnection)
        self.thread.finished.connect(on_thread_finished, Qt.QueuedConnection)

        # cleanup
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.failed.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def run_xlsx(self):
        if not self.folder_path:
            QMessageBox.warning(self, "Błąd", "Nie wybrano folderu")
            return

        # NIE wymagamy OCR. Jeśli txt są – generujemy, jeśli nie – skrypt sam krzyknie.
        try:
            script = os.path.join(self.app_dir, "generate_excel.py")
            subprocess.run(
                [sys.executable, script, "--folder", self.folder_path, "--output", self.output_dir],
                check=True
            )

            QMessageBox.information(self, "Gotowe", f"Wygenerowano plik .xlsx w:\n{self.output_dir}")
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(self, "Błąd generowania xlsx", f"generate_excel.py zakończył się błędem.\n\n{repr(e)}")
        except Exception as e:
            QMessageBox.critical(self, "Błąd", repr(e))

    def run_ocr_and_xlsx(self):
        # Najpierw OCR, potem xlsx
        # Po OCR chcemy odpalić xlsx AUTOMATYCZNIE — więc w run_ocr zrobimy “callback”.
        self._after_ocr = "xlsx"
        self.run_ocr()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FakturyApp()
    window.show()
    sys.exit(app.exec())
