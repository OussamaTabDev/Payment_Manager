import sys
import os
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QTextEdit, QFileDialog,
    QGroupBox, QComboBox, QProgressBar, QMessageBox, QSpinBox
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QFont, QIcon

# Import the main processing module
import payment_processor as processor


class ProcessingThread(QThread):
    """Thread for running the payment processing without blocking UI."""
    progress = pyqtSignal(str)
    finished = pyqtSignal(bool, str)
    
    def __init__(self, parent_file, kid_file, output_file, mode, monthly_fee_a, monthly_fee_b):
        super().__init__()
        self.parent_file = parent_file
        self.kid_file = kid_file
        self.output_file = output_file
        self.mode = mode
        self.monthly_fee_a = monthly_fee_a
        self.monthly_fee_b = monthly_fee_b
    
    def run(self):
        """Run the payment processing."""
        try:
            # Update processor constants
            processor.PARENT_FILE = self.parent_file
            processor.KID_FILE = self.kid_file
            processor.OUTPUT_FILE = self.output_file
            processor.MODE = self.mode
            processor.MONTHLY_FEE_A = self.monthly_fee_a
            processor.MONTHLY_FEE_B = self.monthly_fee_b
            
            self.progress.emit("📂 Loading data...")
            parents_df, kids_df, kids_first_rows, months = processor.load_data(
                self.parent_file, self.kid_file
            )
            self.progress.emit("✅ Data loaded successfully.")
            
            self.progress.emit(f"\n🔧 Running in {self.mode.upper()} mode...")
            kids_df, kids_last_rows, backup_kids_df = processor.filter_dataframe(kids_df, self.mode)
            
            self.progress.emit("\n🔍 Matching kids with parents...")
            combined_df = processor.find_kids_of_parents(parents_df, kids_df, backup_kids_df)
            
            self.progress.emit("\n📊 Creating parent-kid mapping...")
            data_map = processor.get_parent_kid_map(combined_df)
            
            self.progress.emit("\n💰 Calculating payments...")
            amount_map = processor.calculate_months_paid(parents_df)
            
            self.progress.emit("\n📋 Getting kids status...")
            kids_status = processor.get_all_kids_last_updates(self.kid_file, months)
            
            self.progress.emit("\n🧮 Calculating kid payment statuses...")
            kid_payment_status = processor.calculate_kid_payments(
                data_map,
                amount_map,
                {row['kid_name']: {
                    'allocated_amount': 0.0,
                    'class': row['class'],
                    'monthly_fee': processor.get_monthly_fee_for_class(row['class']),
                    'parent': row['parent_name']
                } for _, row in kids_df.iterrows()}
            )
            
            self.progress.emit("\n📝 Updating Excel file...")
            output = processor.update_excel_with_payments(
                kids_df=kids_df,
                kid_payment_status=kid_payment_status,
                kids_status=kids_status,
                months=months,
                kid_file=self.kid_file,
                output_file=self.output_file
            )
            
            self.progress.emit(f"\n✅ Process completed successfully!")
            self.progress.emit(f"📄 Output saved to: {output}")
            self.finished.emit(True, output)
            
        except Exception as e:
            error_msg = f"❌ Error: {str(e)}"
            self.progress.emit(f"\n{error_msg}")
            self.finished.emit(False, error_msg)


class PaymentProcessorGUI(QMainWindow):
    """Main GUI window for payment processing."""
    
    def __init__(self):
        super().__init__()
        self.processing_thread = None
        self.init_ui()
    
    def init_ui(self):
        """Initialize the user interface."""
        self.setWindowTitle("Payment Processing System")
        self.setGeometry(100, 100, 900, 700)
        
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Title
        title = QLabel("🎓 Kids Payment Processing System")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title.setFont(title_font)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(title)
        
        # File selection group
        file_group = self.create_file_selection_group()
        main_layout.addWidget(file_group)
        
        # Settings group
        settings_group = self.create_settings_group()
        main_layout.addWidget(settings_group)
        
        # Output settings group
        output_group = self.create_output_group()
        main_layout.addWidget(output_group)
        
        # Control buttons
        control_layout = QHBoxLayout()
        
        self.process_btn = QPushButton("▶️ Start Processing")
        self.process_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        self.process_btn.clicked.connect(self.start_processing)
        
        self.stop_btn = QPushButton("⏹️ Stop")
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                padding: 10px;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        self.stop_btn.clicked.connect(self.stop_processing)
        
        control_layout.addWidget(self.process_btn)
        control_layout.addWidget(self.stop_btn)
        main_layout.addLayout(control_layout)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Log output
        log_label = QLabel("📋 Processing Log:")
        log_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        main_layout.addWidget(log_label)
        
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #00ff00;
                font-family: 'Courier New', monospace;
                font-size: 11px;
                border: 2px solid #555;
                border-radius: 5px;
                padding: 5px;
            }
        """)
        main_layout.addWidget(self.log_output)
        
        # Status bar
        self.statusBar().showMessage("Ready")
    
    def create_file_selection_group(self):
        """Create file selection group."""
        group = QGroupBox("📁 File Selection")
        layout = QVBoxLayout()
        
        # Parent payments file
        parent_layout = QHBoxLayout()
        parent_label = QLabel("Parent Payments File:")
        parent_label.setMinimumWidth(150)
        self.parent_file_input = QLineEdit("parents_payments.xlsx")
        parent_browse_btn = QPushButton("Browse...")
        parent_browse_btn.clicked.connect(lambda: self.browse_file(self.parent_file_input, "Excel Files (*.xlsx *.xls)"))
        
        parent_layout.addWidget(parent_label)
        parent_layout.addWidget(self.parent_file_input)
        parent_layout.addWidget(parent_browse_btn)
        layout.addLayout(parent_layout)
        
        # Kids list file
        kids_layout = QHBoxLayout()
        kids_label = QLabel("Kids List File:")
        kids_label.setMinimumWidth(150)
        self.kids_file_input = QLineEdit("kids_list.xlsx")
        kids_browse_btn = QPushButton("Browse...")
        kids_browse_btn.clicked.connect(lambda: self.browse_file(self.kids_file_input, "Excel Files (*.xlsx *.xls)"))
        
        kids_layout.addWidget(kids_label)
        kids_layout.addWidget(self.kids_file_input)
        kids_layout.addWidget(kids_browse_btn)
        layout.addLayout(kids_layout)
        
        group.setLayout(layout)
        return group
    
    def create_settings_group(self):
        """Create settings group."""
        group = QGroupBox("⚙️ Processing Settings")
        layout = QVBoxLayout()
        
        # Mode selection
        mode_layout = QHBoxLayout()
        mode_label = QLabel("Processing Mode:")
        mode_label.setMinimumWidth(150)
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["Production", "Test (100 rows)"])
        self.mode_combo.setToolTip("Test mode processes only the first 100 rows")
        
        mode_layout.addWidget(mode_label)
        mode_layout.addWidget(self.mode_combo)
        mode_layout.addStretch()
        layout.addLayout(mode_layout)
        
        # Monthly fees
        fees_layout = QHBoxLayout()
        
        # Fee A
        fee_a_label = QLabel("Monthly Fee (A5-A12):")
        fee_a_label.setMinimumWidth(150)
        self.fee_a_spinbox = QSpinBox()
        self.fee_a_spinbox.setRange(1, 100)
        self.fee_a_spinbox.setValue(25)
        self.fee_a_spinbox.setSuffix(" €")
        
        fees_layout.addWidget(fee_a_label)
        fees_layout.addWidget(self.fee_a_spinbox)
        
        # Fee B
        fee_b_label = QLabel("Monthly Fee (B0-B3):")
        fee_b_label.setMinimumWidth(150)
        self.fee_b_spinbox = QSpinBox()
        self.fee_b_spinbox.setRange(1, 100)
        self.fee_b_spinbox.setValue(15)
        self.fee_b_spinbox.setSuffix(" €")
        
        fees_layout.addWidget(fee_b_label)
        fees_layout.addWidget(self.fee_b_spinbox)
        
        layout.addLayout(fees_layout)
        
        group.setLayout(layout)
        return group
    
    def create_output_group(self):
        """Create output settings group."""
        group = QGroupBox("💾 Output Settings")
        layout = QHBoxLayout()
        
        output_label = QLabel("Output File:")
        output_label.setMinimumWidth(150)
        self.output_file_input = QLineEdit("kids_list_updated.xlsx")
        output_browse_btn = QPushButton("Browse...")
        output_browse_btn.clicked.connect(lambda: self.browse_save_file(self.output_file_input))
        
        layout.addWidget(output_label)
        layout.addWidget(self.output_file_input)
        layout.addWidget(output_browse_btn)
        
        group.setLayout(layout)
        return group
    
    def browse_file(self, line_edit, file_filter):
        """Open file browser dialog."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select File",
            "",
            file_filter
        )
        if file_path:
            line_edit.setText(file_path)
    
    def browse_save_file(self, line_edit):
        """Open save file browser dialog."""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Output File",
            line_edit.text(),
            "Excel Files (*.xlsx)"
        )
        if file_path:
            line_edit.setText(file_path)
    
    def start_processing(self):
        """Start the payment processing."""
        # Validate inputs
        parent_file = self.parent_file_input.text()
        kids_file = self.kids_file_input.text()
        output_file = self.output_file_input.text()
        
        if not os.path.exists(parent_file):
            QMessageBox.warning(self, "Error", f"Parent payments file not found:\n{parent_file}")
            return
        
        if not os.path.exists(kids_file):
            QMessageBox.warning(self, "Error", f"Kids list file not found:\n{kids_file}")
            return
        
        # Clear log
        self.log_output.clear()
        self.log_output.append(f"{'='*60}")
        self.log_output.append(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.log_output.append(f"{'='*60}\n")
        
        # Get settings
        mode = "test" if self.mode_combo.currentIndex() == 1 else "prod"
        monthly_fee_a = self.fee_a_spinbox.value()
        monthly_fee_b = self.fee_b_spinbox.value()
        
        # Update UI state
        self.process_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setVisible(True)
        self.statusBar().showMessage("Processing...")
        
        # Start processing thread
        self.processing_thread = ProcessingThread(
            parent_file, kids_file, output_file, mode, monthly_fee_a, monthly_fee_b
        )
        self.processing_thread.progress.connect(self.update_log)
        self.processing_thread.finished.connect(self.processing_finished)
        self.processing_thread.start()
    
    def stop_processing(self):
        """Stop the processing thread."""
        if self.processing_thread and self.processing_thread.isRunning():
            self.processing_thread.terminate()
            self.processing_thread.wait()
            self.update_log("\n⚠️ Processing stopped by user.")
            self.processing_finished(False, "Stopped by user")
    
    def update_log(self, message):
        """Update the log output."""
        self.log_output.append(message)
        # Auto-scroll to bottom
        self.log_output.verticalScrollBar().setValue(
            self.log_output.verticalScrollBar().maximum()
        )
    
    def processing_finished(self, success, message):
        """Handle processing completion."""
        # Update UI state
        self.process_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.progress_bar.setVisible(False)
        
        if success:
            self.statusBar().showMessage("✅ Processing completed successfully!", 5000)
            QMessageBox.information(
                self,
                "Success",
                f"Payment processing completed successfully!\n\nOutput file: {message}"
            )
        else:
            self.statusBar().showMessage("❌ Processing failed!", 5000)
            if "Stopped by user" not in message:
                QMessageBox.critical(
                    self,
                    "Error",
                    f"Processing failed:\n\n{message}"
                )
        
        self.log_output.append(f"\n{'='*60}")
        self.log_output.append(f"Finished at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.log_output.append(f"{'='*60}")


def main():
    """Main entry point for the GUI application."""
    app = QApplication(sys.argv)
    
    # Set application style
    app.setStyle('Fusion')
    
    # Create and show main window
    window = PaymentProcessorGUI()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()