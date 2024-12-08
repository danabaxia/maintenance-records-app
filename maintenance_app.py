import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QTextEdit, 
                            QPushButton, QTableWidget, QTableWidgetItem, 
                            QDateEdit, QMessageBox)
from PyQt5.QtCore import Qt, QDate
import pandas as pd
from pathlib import Path

class MaintenanceRecordUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Test Equipment Maintenance Records")
        self.setGeometry(100, 100, 800, 600)
        
        # Create main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # Create input form
        form_layout = QVBoxLayout()
        
        # Equipment Details
        equipment_layout = QHBoxLayout()
        equipment_layout.addWidget(QLabel("Equipment ID:"))
        self.equipment_id = QLineEdit()
        equipment_layout.addWidget(self.equipment_id)
        
        equipment_layout.addWidget(QLabel("Equipment Name:"))
        self.equipment_name = QLineEdit()
        equipment_layout.addWidget(self.equipment_name)
        form_layout.addLayout(equipment_layout)
        
        # Maintenance Details
        maintenance_layout = QHBoxLayout()
        maintenance_layout.addWidget(QLabel("Maintenance Date:"))
        self.maintenance_date = QDateEdit()
        self.maintenance_date.setDate(QDate.currentDate())
        maintenance_layout.addWidget(self.maintenance_date)
        
        maintenance_layout.addWidget(QLabel("Technician:"))
        self.technician = QLineEdit()
        maintenance_layout.addWidget(self.technician)
        form_layout.addLayout(maintenance_layout)
        
        # Maintenance Description
        form_layout.addWidget(QLabel("Maintenance Description:"))
        self.description = QTextEdit()
        self.description.setMaximumHeight(100)
        form_layout.addWidget(self.description)
        
        # Buttons
        button_layout = QHBoxLayout()
        self.add_button = QPushButton("Add Record")
        self.add_button.clicked.connect(self.confirm_and_add_record)
        self.clear_button = QPushButton("Clear Form")
        self.clear_button.clicked.connect(self.clear_form)
        self.save_button = QPushButton("Save to Excel")
        self.save_button.clicked.connect(self.save_to_excel)
        self.load_button = QPushButton("Load from Excel")
        self.load_button.clicked.connect(self.load_from_excel)
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.clear_button)
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.load_button)
        form_layout.addLayout(button_layout)
        
        layout.addLayout(form_layout)
        
        # Create table for displaying records
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            "Equipment ID", 
            "Equipment Name", 
            "Maintenance Date", 
            "Technician", 
            "Description"
        ])
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)
        
    def confirm_and_add_record(self):
        if not all([self.equipment_id.text(), 
                   self.equipment_name.text(), 
                   self.technician.text(), 
                   self.description.toPlainText()]):
            QMessageBox.warning(self, "Warning", "Please fill in all fields!")
            return
            
        # Create confirmation message
        confirm_msg = (
            "Please confirm the maintenance record details:\n\n"
            f"Equipment ID: {self.equipment_id.text()}\n"
            f"Equipment Name: {self.equipment_name.text()}\n"
            f"Maintenance Date: {self.maintenance_date.date().toString()}\n"
            f"Technician: {self.technician.text()}\n"
            f"Description: {self.description.toPlainText()}"
        )
        
        reply = QMessageBox.question(self, 'Confirm Record', confirm_msg,
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            self.add_record()
            self.auto_save()
            QMessageBox.information(self, "Success", "Record added and saved automatically!")
            
    def auto_save(self):
        try:
            data = []
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
                data.append(row_data)
            
            df = pd.DataFrame(data, columns=[
                "Equipment ID", 
                "Equipment Name", 
                "Maintenance Date", 
                "Technician", 
                "Description"
            ])
            
            file_name = "maintenance_records.xlsx"
            df.to_excel(file_name, index=False)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to auto-save records: {str(e)}")

    def add_record(self):
        if not all([self.equipment_id.text(), 
                   self.equipment_name.text(), 
                   self.technician.text(), 
                   self.description.toPlainText()]):
            QMessageBox.warning(self, "Warning", "Please fill in all fields!")
            return
            
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        
        self.table.setItem(row_position, 0, 
                          QTableWidgetItem(self.equipment_id.text()))
        self.table.setItem(row_position, 1, 
                          QTableWidgetItem(self.equipment_name.text()))
        self.table.setItem(row_position, 2, 
                          QTableWidgetItem(self.maintenance_date.date().toString()))
        self.table.setItem(row_position, 3, 
                          QTableWidgetItem(self.technician.text()))
        self.table.setItem(row_position, 4, 
                          QTableWidgetItem(self.description.toPlainText()))
        
        self.clear_form()
        
    def clear_form(self):
        self.equipment_id.clear()
        self.equipment_name.clear()
        self.maintenance_date.setDate(QDate.currentDate())
        self.technician.clear()
        self.description.clear()

    def save_to_excel(self):
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "Warning", "No records to save!")
            return

        try:
            data = []
            for row in range(self.table.rowCount()):
                row_data = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
                data.append(row_data)
            
            df = pd.DataFrame(data, columns=[
                "Equipment ID", 
                "Equipment Name", 
                "Maintenance Date", 
                "Technician", 
                "Description"
            ])
            
            file_name = "maintenance_records.xlsx"
            df.to_excel(file_name, index=False)
            QMessageBox.information(self, "Success", f"Records saved to {file_name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save records: {str(e)}")

    def load_from_excel(self):
        try:
            file_name = "maintenance_records.xlsx"
            if not Path(file_name).exists():
                QMessageBox.warning(self, "Warning", "No saved records found!")
                return

            df = pd.read_excel(file_name)
            
            # Clear existing table
            self.table.setRowCount(0)
            
            # Load data into table
            for _, row in df.iterrows():
                row_position = self.table.rowCount()
                self.table.insertRow(row_position)
                for col, value in enumerate(row):
                    self.table.setItem(row_position, col, QTableWidgetItem(str(value)))
                    
            QMessageBox.information(self, "Success", "Records loaded successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load records: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MaintenanceRecordUI()
    window.show()
    sys.exit(app.exec_())