import sys
import os
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QLabel,
    QFileDialog, QProgressBar, QMessageBox, QHBoxLayout, 
    QSpinBox, QCheckBox, QLineEdit, QDateEdit
)
from PySide6.QtCore import QThread, Signal, QDate
from processor import process_csv_files

class Worker(QThread):
    progress = Signal(int)
    finished = Signal(str)

    def __init__(self, folder_path, start_date, params):
        super().__init__()
        self.folder_path = folder_path
        self.start_date = pd.to_datetime(start_date)
        self.params = params

    def run(self):
        output_file = process_csv_files(
            folder_path = self.folder_path, 
            start_date = self.start_date, 
            progress_signal = self.progress, 
            **self.params
            )
        self.finished.emit(output_file)

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("NSE Momentum Stock Selector")
        self.setMinimumWidth(520)

        layout = QVBoxLayout()

        #Folder selection
        self.label = QLabel("Select the Folder")
        layout.addWidget(self.label)

        select_layout = QHBoxLayout()
        self.select_button = QPushButton("Select Folder")
        self.select_button.clicked.connect(self.select_folder)
        select_layout.addWidget(self.select_button)

        self.folder_label = QLabel("No folder selected")
        select_layout.addWidget(self.folder_label)
        layout.addLayout(select_layout)

        #Parameters
        params_layout = QHBoxLayout()
        
        params_layout.addWidget(QLabel("Lookback (months):"))
        self.lookback_spin = QSpinBox()
        self.lookback_spin.setRange(1, 60)
        self.lookback_spin.setValue(12)
        params_layout.addWidget(self.lookback_spin)

        params_layout.addWidget(QLabel("Top K:"))
        self.topk_spin = QSpinBox()
        self.topk_spin.setRange(1, 1000)
        self.topk_spin.setValue(10)
        params_layout.addWidget(self.topk_spin)

        params_layout.addWidget(QLabel("Portfolio Capital"))
        self.capital_edit = QLineEdit("100000")
        params_layout.addWidget(self.capital_edit)

        params_layout.addWidget(QLabel("Benchmark Percentage: "))
        self.benchmark_edit = QLineEdit("6.0")
        params_layout.addWidget(self.benchmark_edit)

        params_layout.addWidget(QLabel("Buffer:"))
        self.buffer_spin = QSpinBox()
        self.buffer_spin.setRange(0, 20)
        self.buffer_spin.setValue(10)
        params_layout.addWidget(self.buffer_spin)

        layout.addLayout(params_layout)

        #Start Date
        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel("Last Date:"))
        
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_date_edit.setMinimumDate(QDate(2016, 4, 1))
        self.start_date_edit.setMaximumDate(QDate.currentDate())
        self.start_date_edit.setDate(QDate.currentDate())
        date_layout.addWidget(self.start_date_edit)

        layout.addLayout(date_layout)

        #Run Button and Progress Bar
        self.run_button = QPushButton("Run Momentum Backtest")
        self.run_button.clicked.connect(self.run_processing)
        self.run_button.setEnabled(False)
        layout.addWidget(self.run_button)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)
        
        self.setLayout(layout)

        self.folder_path = None
        self.worker = None

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.folder_path = folder
            self.label.setText(f"Selected Folder: {folder}")
            self.run_button.setEnabled(True)
    
    def run_processing(self):
        if not self.folder_path:
            QMessageBox.warning(self, "No Folder", "Please select a folder first.")
            return

        try:
            capital = float(self.capital_edit.text())
        except ValueError:
            QMessageBox.warning(self, "Portfolio capital must be numeric.")
            return
        
        params = {
            "lookback_months": int(self.lookback_spin.value()),
            "top_k": int(self.topk_spin.value()),
            "buffer": int(self.buffer_spin.value()),
            "benchmark": float(self.benchmark_edit.text()),
            "portfolio_capital": capital
        }

        qt_date = self.start_date_edit.date()
        start_date = qt_date.toPython()
        
        self.run_button.setEnabled(False)
        self.progress.setValue(0)

        self.worker = Worker(self.folder_path, start_date, params)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self.processing_done)
        self.worker.start()

    def processing_done(self, output_file):
        self.run_button.setEnabled(True)
        QMessageBox.information(self, "Done", f"Processing completed!\nSaved: {output_file}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())