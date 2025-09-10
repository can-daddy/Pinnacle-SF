import sys
import sqlite3
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTableWidget, QTableWidgetItem, QDialog, QFormLayout, QMessageBox, QHeaderView
from PyQt5.QtCore import Qt
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

# Context manager for database
from contextlib import contextmanager

@contextmanager
def get_db_connection():
    conn = sqlite3.connect('school.db')  # Adjust path if needed
    c = conn.cursor()
    try:
        yield conn, c
    finally:
        conn.close()

class Dashboard(QWidget):
    def __init__(self):
        super().__init__()
        self.setStyleSheet("""
            QLabel {
                color: #2D3748;
            }
            QPushButton {
                background-color: #4299E1;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #3182CE;
            }
            QTableWidget {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                border-radius: 5px;
                gridline-color: #E0E3E7;
                selection-background-color: #EDF2F7;
            }
            QTableWidget::item {
                padding: 10px;
                color: #2D3748;
            }
            QHeaderView::section {
                background-color: #F7FAFC;
                padding: 10px;
                border: none;
                border-bottom: 1px solid #E0E3E7;
                font-weight: bold;
                color: #4A5568;
            }
        """)
        main_layout = QHBoxLayout(self)
        main_layout.setSpacing(20)

        # Left column: Pie charts and last 5 inventory table
        left_layout = QVBoxLayout()
        left_layout.setSpacing(15)

        # Pie charts section
        charts_layout = QHBoxLayout()
        self.in_canvas = self.create_pie_canvas("Stock In (Current Month)")
        self.out_canvas = self.create_pie_canvas("Stock Out (Current Month)")
        charts_layout.addWidget(self.in_canvas)
        charts_layout.addWidget(self.out_canvas)
        left_layout.addLayout(charts_layout)

        # Last 5 inventory table
        inventory_title = QLabel("Last 5 Inventory Entries")
        inventory_title.setStyleSheet("font-size: 16px; font-weight: bold;")
        left_layout.addWidget(inventory_title)
        self.inventory_table = QTableWidget()
        self.inventory_table.setColumnCount(7)
        self.inventory_table.setHorizontalHeaderLabels(["ID", "Uniform", "Type", "Quantity", "Price", "Date", "References"])
        self.inventory_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.inventory_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.inventory_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.inventory_table.setShowGrid(True)
        self.inventory_table.verticalHeader().setVisible(False)
        left_layout.addWidget(self.inventory_table)

        # Right column: Uniform list and add button
        right_layout = QVBoxLayout()
        right_layout.setSpacing(15)

        uniform_title = QLabel("Uniform List")
        uniform_title.setStyleSheet("font-size: 16px; font-weight: bold;")
        right_layout.addWidget(uniform_title)
        self.uniform_table = QTableWidget()
        self.uniform_table.setColumnCount(4)
        self.uniform_table.setHorizontalHeaderLabels(["ID", "Name", "Price", "Quantity"])
        self.uniform_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.uniform_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.uniform_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.uniform_table.setShowGrid(True)
        self.uniform_table.verticalHeader().setVisible(False)
        right_layout.addWidget(self.uniform_table)

        add_uniform_btn = QPushButton("Add New Uniform")
        add_uniform_btn.clicked.connect(self.show_add_uniform_dialog)
        right_layout.addWidget(add_uniform_btn, alignment=Qt.AlignBottom)

        main_layout.addLayout(left_layout)
        main_layout.addLayout(right_layout)

        self.load_data()

    def create_pie_canvas(self, title):
        fig = Figure(figsize=(4, 4), dpi=100)
        ax = fig.add_subplot(111)
        ax.set_title(title)
        canvas = FigureCanvas(fig)
        return canvas

    def load_data(self):
        current_month = datetime.now().strftime('%Y-%m')
        
        try:
            with get_db_connection() as (conn, c):
                # Calculate stock in and out for current month
                c.execute("""
                    SELECT SUM(Quantity) 
                    FROM UniformInventory 
                    WHERE Type = 'In' AND strftime('%Y-%m', Date) = ?
                """, (current_month,))
                stock_in = c.fetchone()[0] or 0
                
                c.execute("""
                    SELECT SUM(Quantity) 
                    FROM UniformInventory 
                    WHERE Type = 'Out' AND strftime('%Y-%m', Date) = ?
                """, (current_month,))
                stock_out = c.fetchone()[0] or 0
                
                total = stock_in + stock_out
                in_percentage = (stock_in / total * 100) if total > 0 else 0
                out_percentage = (stock_out / total * 100) if total > 0 else 0
                
                # Update in pie
                fig_in = self.in_canvas.figure
                fig_in.clf()
                ax_in = fig_in.add_subplot(111)
                ax_in.pie([in_percentage, 100 - in_percentage], labels=['In', 'Other'], colors=['#00FF00', '#F0F0F0'], autopct='%1.1f%%')
                ax_in.set_title("Stock In (Current Month)")
                self.in_canvas.draw()
                
                # Update out pie
                fig_out = self.out_canvas.figure
                fig_out.clf()
                ax_out = fig_out.add_subplot(111)
                ax_out.pie([out_percentage, 100 - out_percentage], labels=['Out', 'Other'], colors=['#FF0000', '#F0F0F0'], autopct='%1.1f%%')
                ax_out.set_title("Stock Out (Current Month)")
                self.out_canvas.draw()
                
                # Last 5 inventory - quote "References"
                c.execute("""
                    SELECT ui.UniformInventoryID, u.UniformName, ui.Type, ui.Quantity, ui.Price, ui.Date, ui."References"
                    FROM UniformInventory ui
                    LEFT JOIN Uniform u ON ui.UniformID = u.UniformId
                    ORDER BY ui.Date DESC
                    LIMIT 5
                """)
                rows = c.fetchall()
                self.inventory_table.setRowCount(len(rows))
                for i, row in enumerate(rows):
                    for j, val in enumerate(row):
                        self.inventory_table.setItem(i, j, QTableWidgetItem(str(val)))
                
                # Uniform list
                c.execute("SELECT UniformId, UniformName, UniformPrice, Quantity FROM Uniform")
                rows = c.fetchall()
                self.uniform_table.setRowCount(len(rows))
                for i, row in enumerate(rows):
                    for j, val in enumerate(row):
                        self.uniform_table.setItem(i, j, QTableWidgetItem(str(val)))
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load dashboard data: {str(e)}")

    def show_add_uniform_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add New Uniform")
        layout = QVBoxLayout()
        form_layout = QFormLayout()
        name_edit = QLineEdit()
        price_edit = QLineEdit()
        quantity_edit = QLineEdit()
        form_layout.addRow("Uniform Name:", name_edit)
        form_layout.addRow("Price:", price_edit)
        form_layout.addRow("Quantity:", quantity_edit)
        layout.addLayout(form_layout)
        buttons = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_uniform(dialog, name_edit, price_edit, quantity_edit))
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        buttons.addWidget(save_btn)
        buttons.addWidget(close_btn)
        layout.addLayout(buttons)
        dialog.setLayout(layout)
        dialog.exec_()

    def save_uniform(self, dialog, name_edit, price_edit, quantity_edit):
        name = name_edit.text().strip()
        price_str = price_edit.text().strip()
        quantity_str = quantity_edit.text().strip()
        if not all([name, price_str, quantity_str]):
            QMessageBox.warning(dialog, "Error", "All fields are required")
            return
        try:
            price = float(price_str)
            quantity = int(quantity_str)
            if price <= 0 or quantity <= 0:
                QMessageBox.warning(dialog, "Error", "Price and Quantity must be positive")
                return
        except ValueError:
            QMessageBox.warning(dialog, "Error", "Invalid price or quantity")
            return
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT COUNT(*) FROM Uniform WHERE UniformName = ?", (name,))
                if c.fetchone()[0] > 0:
                    QMessageBox.warning(dialog, "Error", "Duplicate uniform name")
                    return
                c.execute("""INSERT INTO Uniform (UniformName, UniformPrice, Quantity)
                             VALUES (?, ?, ?)""", (name, price, quantity))
                conn.commit()
                QMessageBox.information(dialog, "Success", "Uniform added successfully")
                name_edit.clear()
                price_edit.clear()
                quantity_edit.clear()
                self.load_data()
        except Exception as e:
            QMessageBox.warning(dialog, "Error", f"Failed to add uniform: {str(e)}")

# To integrate in your main window:
# self.body_layout.addWidget(Dashboard())