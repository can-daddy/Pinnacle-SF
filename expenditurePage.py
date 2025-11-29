import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sqlite3
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from xlsxwriter import *
from contextlib import contextmanager
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt5.QtChart import *
from PyQt5.QtPrintSupport import QPrinter

# Enhanced connection function with context manager
@contextmanager
def get_db_connection():
    """Context manager for database connections that handles opening and closing automatically"""
    conn = None
    try:
        conn = sqlite3.connect("Database.db")
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        yield conn, c
    except sqlite3.Error as e:
        print(f"Database connection error: {e}")
        raise
    finally:
        if conn:
            conn.close()

# Helper to ensure stock consistency
def recalculate_supply_stock(supply_id):
    try:
        with get_db_connection() as (conn, c):
            # Calculate total In
            c.execute("SELECT SUM(Quantity) FROM SuppliesInventory WHERE Supplies_ID = ? AND Type = 'In'", (supply_id,))
            total_in = c.fetchone()[0] or 0
            
            # Calculate total Out
            c.execute("SELECT SUM(Quantity) FROM SuppliesInventory WHERE Supplies_ID = ? AND Type = 'Out'", (supply_id,))
            total_out = c.fetchone()[0] or 0
            
            # Net Quantity
            net_qty = total_in - total_out
            
            # Update Supplies table
            c.execute("UPDATE Supplies SET Quantity = ? WHERE SuppliesID = ?", (net_qty, supply_id))
            conn.commit()
            return net_qty
    except Exception as e:
        print(f"Stock recalculation error: {e}")
        raise

def validate_inventory_change(supply_id, new_type=None, new_qty=0, exclude_inv_id=None):
    """
    Checks if a proposed inventory change is valid (prevents negative stock).
    Returns (is_valid, projected_stock).
    """
    try:
        with get_db_connection() as (conn, c):
            # Get all inventory records for this supply
            query = "SELECT Sup_Inv_ID, Type, Quantity FROM SuppliesInventory WHERE Supplies_ID = ?"
            c.execute(query, (supply_id,))
            records = c.fetchall()
            
            total_in = 0
            total_out = 0
            
            # Sum up existing records, excluding the one being modified/deleted
            for r in records:
                if exclude_inv_id and r['Sup_Inv_ID'] == exclude_inv_id:
                    continue
                if r['Type'] == 'In':
                    total_in += r['Quantity']
                elif r['Type'] == 'Out':
                    total_out += r['Quantity']
            
            # Add the new/modified transaction if provided
            if new_type and new_qty is not None:
                if new_type == 'In':
                    total_in += int(new_qty)
                elif new_type == 'Out':
                    total_out += int(new_qty)
            
            return total_in >= total_out, total_in - total_out
    except Exception as e:
        print(f"Validation error: {e}")
        return False, 0

class Expenditure(QTabWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Expenditure Management")
        layout = QVBoxLayout()
        self.tabs = QTabWidget()
        self.tab1 = Supplies()
        self.tab2 = SuppliesInventory()
        self.tab3 = Expense()
        self.tab4 = Assets()
        self.tabs.addTab(self.tab1, "Supplies")
        self.tabs.addTab(self.tab2, "Supplies Inventory")
        self.tabs.addTab(self.tab3, "Expenses")
        self.tabs.addTab(self.tab4, "Assets")
        layout.addWidget(self.tabs)
        self.setLayout(layout)
        
        # Refresh inventory tab when clicked
        self.tabs.currentChanged.connect(self.on_tab_change)

    def on_tab_change(self, index):
        if index == 1: # Supplies Inventory tab
            self.tab2.load_data()
        elif index == 0: # Supplies tab
            self.tab1.load_data()
        elif index == 2: # Expenses tab
            try:
                self.tab3.load_data()
            except Exception:
                pass
        elif index == 3: # Assets tab
            if self.tab4:
                try:
                    self.tab4.load_data()
                except Exception:
                    pass

class Supplies(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Stylesheet
        self.setStyleSheet("""
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
            QLineEdit {
                padding: 8px;
                border: 1px solid #E2E8F0;
                border-radius: 5px;
            }
            QTableWidget {
                border: 1px solid #E2E8F0;
                gridline-color: #E2E8F0;
            }
            QHeaderView::section {
                background-color: #F7FAFC;
                padding: 8px;
                border: none;
                font-weight: bold;
            }
        """)

        self.header = QHBoxLayout()
        self.Heading = QLabel("Supplies")
        self.Heading.setStyleSheet("font-size: 18px; font-weight: bold; color: #2D3748;")
        self.searchbar = QLineEdit()
        self.searchbar.setPlaceholderText("Search Supplies...")
        self.searchbar.textChanged.connect(self.load_data)
        self.addbtn = QPushButton("Add Supply")
        self.addbtn.clicked.connect(self.show_add_supply_dialog)
        
        self.header.addWidget(self.Heading)
        self.header.addStretch()
        self.header.addWidget(self.searchbar)
        self.header.addWidget(self.addbtn)
        layout.addLayout(self.header)
        
        self.table = QTableWidget()
        self.table.setColumnCount(6) # Added ID column (hidden)
        self.table.setHorizontalHeaderLabels(["ID", "Name", "Subcategory", "Quantity", "Unit", "Actions"])
        self.table.hideColumn(0) # Hide ID
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Fixed)
        self.table.setColumnWidth(5, 250)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table)
        
        self.load_data()

    def load_data(self):
        try:
            with get_db_connection() as (conn, c):
                query = """
                    SELECT s.SuppliesID, s.Name, sc.SubCategoryName, s.Quantity, s.Unit 
                    FROM Supplies s
                    LEFT JOIN SubCategory sc ON s.Subcategory = sc.SubCatID
                """
                params = []
                search_text = self.searchbar.text().strip()
                if search_text:
                    query += " WHERE s.Name LIKE ?"
                    params.append(f"%{search_text}%")
                
                c.execute(query, params)
                rows = c.fetchall()
                
                self.table.setRowCount(0)
                for row_idx, row_data in enumerate(rows):
                    self.table.insertRow(row_idx)
                    # Data columns
                    for col_idx, data in enumerate(row_data):
                        self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(data) if data is not None else ""))
                    
                    # Actions
                    action_widget = QWidget()
                    action_layout = QHBoxLayout(action_widget)
                    action_layout.setContentsMargins(5, 2, 5, 2)
                    
                    inv_btn = QPushButton("Inventory")
                    inv_btn.setStyleSheet("background-color: #805AD5;")
                    inv_btn.clicked.connect(lambda _, rid=row_data[0], rname=row_data[1]: self.manage_inventory(rid, rname))
                    
                    edit_btn = QPushButton("Edit")
                    edit_btn.setStyleSheet("background-color: #38A169;")
                    edit_btn.clicked.connect(lambda _, rid=row_data[0]: self.edit_supply(rid))
                    
                    del_btn = QPushButton("Delete")
                    del_btn.setStyleSheet("background-color: #E53E3E;")
                    del_btn.clicked.connect(lambda _, rid=row_data[0]: self.delete_supply(rid))
                    
                    action_layout.addWidget(inv_btn)
                    action_layout.addWidget(edit_btn)
                    action_layout.addWidget(del_btn)
                    self.table.setCellWidget(row_idx, 5, action_widget)
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load data: {e}")

    def show_add_supply_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add New Supply")
        dialog.setMinimumWidth(400)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        
        name_input = QLineEdit()
        dept_combo = QComboBox()
        cat_combo = QComboBox()
        subcat_combo = QComboBox()
        unit_input = QLineEdit()
        
        # Populate Departments
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT DeptID, DeptName FROM Department")
                depts = c.fetchall()
                dept_combo.addItem("Select Department", None)
                for d in depts:
                    dept_combo.addItem(d[1], d[0])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Database Error: {e}")

        # Cascading logic
        def update_categories():
            cat_combo.clear()
            cat_combo.addItem("Select Category", None)
            dept_id = dept_combo.currentData()
            if dept_id:
                try:
                    with get_db_connection() as (conn, c):
                        c.execute("SELECT CatID, CatName FROM Category WHERE Department = ? AND Type='Supplies'", (dept_id,))
                        cats = c.fetchall()
                        for cat in cats:
                            cat_combo.addItem(cat[1], cat[0])
                except Exception as e:
                    print(e)

        def update_subcategories():
            subcat_combo.clear()
            subcat_combo.addItem("Select Subcategory", None)
            cat_id = cat_combo.currentData()
            if cat_id:
                try:
                    with get_db_connection() as (conn, c):
                        c.execute("SELECT SubCatID, SubCategoryName FROM SubCategory WHERE CategoryID = ?", (cat_id,))
                        subcats = c.fetchall()
                        for sub in subcats:
                            subcat_combo.addItem(sub[1], sub[0])
                except Exception as e:
                    print(e)

        dept_combo.currentIndexChanged.connect(update_categories)
        cat_combo.currentIndexChanged.connect(update_subcategories)

        form.addRow("Name:", name_input)
        form.addRow("Department:", dept_combo)
        form.addRow("Category:", cat_combo)
        form.addRow("Subcategory:", subcat_combo)
        form.addRow("Unit:", unit_input)
        
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_supply(dialog, name_input, subcat_combo, unit_input))
        
        layout.addLayout(form)
        layout.addWidget(save_btn)
        dialog.exec_()

    def save_supply(self, dialog, name_input, subcat_combo, unit_input):
        name = name_input.text().strip()
        subcat_id = subcat_combo.currentData()
        unit = unit_input.text().strip()
        
        if not all([name, subcat_id, unit]):
            QMessageBox.warning(dialog, "Missing Data", "Please fill all fields.")
            return
            
        try:
            with get_db_connection() as (conn, c):
                c.execute("INSERT INTO Supplies (Name, Subcategory, Quantity, Unit) VALUES (?, ?, 0, ?)",
                          (name, subcat_id, unit))
                conn.commit()
            dialog.accept()
            self.load_data()
            QMessageBox.information(self, "Success", "Supply added successfully!")
        except Exception as e:
            QMessageBox.critical(dialog, "Error", f"Failed to save: {e}")

    def edit_supply(self, supply_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT Name, Subcategory, Unit FROM Supplies WHERE SuppliesID=?", (supply_id,))
                data = c.fetchone()
                
            if not data:
                return

            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Supply")
            layout = QVBoxLayout(dialog)
            form = QFormLayout()
            
            name_input = QLineEdit(data[0])
            unit_input = QLineEdit(data[2])
            
            form.addRow("Name:", name_input)
            form.addRow("Unit:", unit_input)
            
            save_btn = QPushButton("Update")
            save_btn.clicked.connect(lambda: self.update_supply(dialog, supply_id, name_input, unit_input))
            
            layout.addLayout(form)
            layout.addWidget(save_btn)
            dialog.exec_()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error preparing edit: {e}")

    def update_supply(self, dialog, supply_id, name_input, unit_input):
        try:
            with get_db_connection() as (conn, c):
                c.execute("UPDATE Supplies SET Name=?, Unit=? WHERE SuppliesID=?", 
                          (name_input.text().strip(), unit_input.text().strip(), supply_id))
                conn.commit()
            dialog.accept()
            self.load_data()
        except Exception as e:
            QMessageBox.critical(dialog, "Error", f"Update failed: {e}")

    def delete_supply(self, supply_id):
        reply = QMessageBox.question(self, "Confirm Delete", "Are you sure? This will delete all inventory records for this supply.",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    c.execute("DELETE FROM SuppliesInventory WHERE Supplies_ID=?", (supply_id,))
                    c.execute("DELETE FROM Supplies WHERE SuppliesID=?", (supply_id,))
                    conn.commit()
                self.load_data()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Delete failed: {e}")

    def manage_inventory(self, supply_id, supply_name):
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Inventory: {supply_name}")
        layout = QVBoxLayout(dialog)
        form = QFormLayout()
        
        type_combo = QComboBox()
        type_combo.addItems(["In", "Out"])
        
        qty_input = QLineEdit()
        qty_input.setValidator(QIntValidator(1, 10000))
        
        price_input = QLineEdit()
        price_input.setValidator(QDoubleValidator(0.0, 100000.0, 2))
        
        date_input = QDateEdit()
        date_input.setDate(QDate.currentDate())
        date_input.setCalendarPopup(True)
        
        ref_input = QLineEdit()
        
        form.addRow("Type:", type_combo)
        form.addRow("Quantity:", qty_input)
        form.addRow("Price (Total):", price_input)
        form.addRow("Date:", date_input)
        form.addRow("Reference/Notes:", ref_input)
        
        save_btn = QPushButton("Save Transaction")
        save_btn.clicked.connect(lambda: self.save_inventory(dialog, supply_id, type_combo, qty_input, price_input, date_input, ref_input))
        
        layout.addLayout(form)
        layout.addWidget(save_btn)
        dialog.exec_()

    def save_inventory(self, dialog, supply_id, type_combo, qty_input, price_input, date_input, ref_input):
        itype = type_combo.currentText()
        qty = qty_input.text()
        price = price_input.text()
        date = date_input.date().toString("yyyy-MM-dd")
        ref = ref_input.text()
        
        if not qty or not price:
            QMessageBox.warning(dialog, "Error", "Quantity and Price are required.")
            return
            
        try:
            qty = int(qty)
            price = float(price)
            
            # Validate stock
            is_valid, projected = validate_inventory_change(supply_id, itype, qty)
            if not is_valid:
                QMessageBox.warning(dialog, "Invalid Transaction", 
                                  f"Insufficient stock! This transaction would result in negative inventory.\nCurrent projected stock: {projected}")
                return
            
            with get_db_connection() as (conn, c):
                # 1. Add record
                c.execute("""INSERT INTO SuppliesInventory (Supplies_ID, Type, Quantity, Price, Date, ReferencesNotes)
                             VALUES (?, ?, ?, ?, ?, ?)""",
                          (supply_id, itype, qty, price, date, ref))
                conn.commit()
            
            # 2. Recalculate stock
            recalculate_supply_stock(supply_id)
            
            dialog.accept()
            self.load_data()
            QMessageBox.information(self, "Success", "Inventory updated.")
            
        except Exception as e:
            QMessageBox.critical(dialog, "Error", f"Transaction failed: {e}")

class SuppliesInventory(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Stylesheet
        self.setStyleSheet("""
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
            QLineEdit {
                padding: 8px;
                border: 1px solid #E2E8F0;
                border-radius: 5px;
            }
            QTableWidget {
                border: 1px solid #E2E8F0;
                gridline-color: #E2E8F0;
            }
            QHeaderView::section {
                background-color: #F7FAFC;
                padding: 8px;
                border: none;
                font-weight: bold;
            }
        """)

        self.header = QHBoxLayout()
        self.Heading = QLabel("Supplies Inventory")
        self.Heading.setStyleSheet("font-size: 18px; font-weight: bold; color: #2D3748;")
        self.searchbar = QLineEdit()
        self.searchbar.setPlaceholderText("Search Supplies Inventory...")
        self.searchbar.textChanged.connect(self.load_data)
        self.header.addWidget(self.Heading)
        self.header.addStretch()
        self.header.addWidget(self.searchbar)
        layout.addLayout(self.header)
        
        self.table = QTableWidget()
        self.table.setColumnCount(8) # ID, SupplyID, Name, Type, Quantity, Price, Date, Actions
        self.table.setHorizontalHeaderLabels(["ID", "SupplyID", "Name", "Type", "Quantity", "Price", "Date", "Actions"])
        self.table.hideColumn(0) # Hide Inventory ID
        self.table.hideColumn(1) # Hide Supply ID
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(7, QHeaderView.Fixed)
        self.table.setColumnWidth(7, 180)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table)
        
        self.load_data()

    def load_data(self):
        try:
            with get_db_connection() as (conn, c):
                query = """
                    SELECT si.Sup_Inv_ID, si.Supplies_ID, s.Name, si.Type, si.Quantity, si.Price, si.Date
                    FROM SuppliesInventory si
                    JOIN Supplies s ON si.Supplies_ID = s.SuppliesID
                """
                params = []
                search_text = self.searchbar.text().strip()
                if search_text:
                    query += " WHERE s.Name LIKE ? OR si.ReferencesNotes LIKE ?"
                    params.append(f"%{search_text}%")
                    params.append(f"%{search_text}%")
                
                query += " ORDER BY si.Date DESC"
                
                c.execute(query, params)
                rows = c.fetchall()
                
                self.table.setRowCount(0)
                for row_idx, row_data in enumerate(rows):
                    self.table.insertRow(row_idx)
                    for col_idx, data in enumerate(row_data):
                        self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(data)))
                    
                    # Actions
                    action_widget = QWidget()
                    action_layout = QHBoxLayout(action_widget)
                    action_layout.setContentsMargins(5, 2, 5, 2)
                    
                    edit_btn = QPushButton("Edit")
                    edit_btn.setStyleSheet("background-color: #38A169;")
                    edit_btn.clicked.connect(lambda _, rid=row_data[0]: self.edit_inventory(rid))
                    
                    del_btn = QPushButton("Delete")
                    del_btn.setStyleSheet("background-color: #E53E3E;")
                    del_btn.clicked.connect(lambda _, rid=row_data[0], sid=row_data[1]: self.delete_inventory(rid, sid))
                    
                    action_layout.addWidget(edit_btn)
                    action_layout.addWidget(del_btn)
                    self.table.setCellWidget(row_idx, 7, action_widget)
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load inventory: {e}")

    def edit_inventory(self, inv_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT Supplies_ID, Type, Quantity, Price, Date, ReferencesNotes FROM SuppliesInventory WHERE Sup_Inv_ID=?", (inv_id,))
                data = c.fetchone()
                
            if not data:
                return
            
            supply_id = data[0]

            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Inventory Record")
            layout = QVBoxLayout(dialog)
            form = QFormLayout()
            
            type_combo = QComboBox()
            type_combo.addItems(["In", "Out"])
            type_combo.setCurrentText(data[1])
            
            qty_input = QLineEdit(str(data[2]))
            qty_input.setValidator(QIntValidator(1, 10000))
            
            price_input = QLineEdit(str(data[3]))
            price_input.setValidator(QDoubleValidator(0.0, 100000.0, 2))
            
            date_input = QDateEdit()
            date_input.setDate(QDate.fromString(data[4], "yyyy-MM-dd"))
            date_input.setCalendarPopup(True)
            
            ref_input = QLineEdit(data[5] if data[5] else "")
            
            form.addRow("Type:", type_combo)
            form.addRow("Quantity:", qty_input)
            form.addRow("Price:", price_input)
            form.addRow("Date:", date_input)
            form.addRow("Reference:", ref_input)
            
            save_btn = QPushButton("Update")
            save_btn.clicked.connect(lambda: self.update_inventory(dialog, inv_id, supply_id, type_combo, qty_input, price_input, date_input, ref_input))
            
            layout.addLayout(form)
            layout.addWidget(save_btn)
            dialog.exec_()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error preparing edit: {e}")

    def update_inventory(self, dialog, inv_id, supply_id, type_combo, qty_input, price_input, date_input, ref_input):
        try:
            qty = int(qty_input.text())
            price = float(price_input.text())
            
            # Validate stock
            is_valid, projected = validate_inventory_change(supply_id, type_combo.currentText(), qty, exclude_inv_id=inv_id)
            if not is_valid:
                QMessageBox.warning(dialog, "Invalid Update", 
                                  f"Insufficient stock! This update would result in negative inventory.\nProjected stock: {projected}")
                return

            with get_db_connection() as (conn, c):
                c.execute("""UPDATE SuppliesInventory 
                             SET Type=?, Quantity=?, Price=?, Date=?, ReferencesNotes=?
                             WHERE Sup_Inv_ID=?""",
                          (type_combo.currentText(), qty, price, date_input.date().toString("yyyy-MM-dd"), ref_input.text(), inv_id))
                conn.commit()
            
            # Recalculate stock
            recalculate_supply_stock(supply_id)
            
            dialog.accept()
            self.load_data()
            QMessageBox.information(self, "Success", "Inventory updated.")
        except Exception as e:
            QMessageBox.critical(dialog, "Error", f"Update failed: {e}")

    def delete_inventory(self, inv_id, supply_id):
        reply = QMessageBox.question(self, "Confirm Delete", "Are you sure you want to delete this record?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                # Validate stock before deletion
                is_valid, projected = validate_inventory_change(supply_id, exclude_inv_id=inv_id)
                if not is_valid:
                    QMessageBox.warning(self, "Invalid Delete", 
                                      f"Cannot delete! Removing this record would result in negative inventory.\nProjected stock: {projected}")
                    return

                with get_db_connection() as (conn, c):
                    c.execute("DELETE FROM SuppliesInventory WHERE Sup_Inv_ID=?", (inv_id,))
                    conn.commit()
                
                # Recalculate stock
                recalculate_supply_stock(supply_id)
                
                self.load_data()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Delete failed: {e}")

class Expense(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)
        self.header = QHBoxLayout()
        self.Heading = QLabel("Expenses")
        self.Heading.setStyleSheet("font-size:18px;font-weight:bold;")
        self.searchbar = QLineEdit()
        self.searchbar.setPlaceholderText("Search Expenses...")
        self.searchbar.textChanged.connect(self.load_data)
        self.addbtn = QPushButton("Add Expense")
        self.addbtn.clicked.connect(self.add_expense)
        self.header.addWidget(self.Heading)
        self.header.addStretch()
        self.header.addWidget(self.searchbar)
        self.header.addWidget(self.addbtn)
        layout.addLayout(self.header)

        self.table = QTableWidget()
        # Hidden ExpID, Category, Subcategory, Amount, Date, Description, Actions
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["ID","Category","Subcategory","Amount","Date","Description","Actions"])
        self.table.hideColumn(0)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table)

        self.load_data()

    def _ensure_expense_description_column(self, c):
        # ensure the Expenses table has ExpDescription column; add it if missing
        c.execute("PRAGMA table_info(Expenses)")
        cols = [r[1] for r in c.fetchall()]
        if 'ExpDescription' not in cols:
            try:
                c.execute("ALTER TABLE Expenses ADD COLUMN ExpDescription TEXT")
            except Exception:
                pass

    def load_data(self):
        try:
            with get_db_connection() as (conn, c):
                # Ensure column exists so queries won't fail on older DBs
                self._ensure_expense_description_column(c)

                # Check whether ExpDescription exists now
                c.execute("PRAGMA table_info(Expenses)")
                cols = [r[1] for r in c.fetchall()]
                has_desc = 'ExpDescription' in cols

                if has_desc:
                    query = (
                        "SELECT e.ExpID, c.CatName, sc.SubCategoryName, e.ExpAmount, e.ExpDate, e.ExpDescription "
                        "FROM Expenses e "
                        "LEFT JOIN SubCategory sc ON e.ExpSubCategory = sc.SubCatID "
                        "LEFT JOIN Category c ON sc.CategoryID = c.CatID "
                    )
                else:
                    # select a blank description to keep column positions stable
                    query = (
                        "SELECT e.ExpID, c.CatName, sc.SubCategoryName, e.ExpAmount, e.ExpDate, '' as ExpDescription "
                        "FROM Expenses e "
                        "LEFT JOIN SubCategory sc ON e.ExpSubCategory = sc.SubCatID "
                        "LEFT JOIN Category c ON sc.CategoryID = c.CatID "
                    )

                params = []
                search = self.searchbar.text().strip()
                if search:
                    if has_desc:
                        query += " WHERE c.CatName LIKE ? OR sc.SubCategoryName LIKE ? OR e.ExpDescription LIKE ?"
                        params.extend([f"%{search}%"] * 3)
                    else:
                        query += " WHERE c.CatName LIKE ? OR sc.SubCategoryName LIKE ?"
                        params.extend([f"%{search}%"] * 2)

                query += " ORDER BY e.ExpDate DESC"
                c.execute(query, params)
                rows = c.fetchall()

                self.table.setRowCount(0)
                for idx, row in enumerate(rows):
                    self.table.insertRow(idx)
                    # row: ExpID, CatName, SubCategoryName, ExpAmount, ExpDate, ExpDescription
                    self.table.setItem(idx, 0, QTableWidgetItem(str(row[0] if row[0] is not None else "")))
                    self.table.setItem(idx, 1, QTableWidgetItem(str(row[1] if row[1] else "")))
                    self.table.setItem(idx, 2, QTableWidgetItem(str(row[2] if row[2] else "")))
                    self.table.setItem(idx, 3, QTableWidgetItem(str(row[3] if row[3] is not None else "")))
                    self.table.setItem(idx, 4, QTableWidgetItem(str(row[4] if row[4] else "")))
                    self.table.setItem(idx, 5, QTableWidgetItem(str(row[5] if row[5] else "")))

                    # Actions
                    action_widget = QWidget()
                    action_layout = QHBoxLayout(action_widget)
                    action_layout.setContentsMargins(5, 2, 5, 2)

                    edit_btn = QPushButton("Edit")
                    edit_btn.setStyleSheet("background-color: #38A169; color: white;")
                    edit_btn.clicked.connect(lambda _, rid=row[0]: self.edit_expense(rid))

                    del_btn = QPushButton("Delete")
                    del_btn.setStyleSheet("background-color: #E53E3E; color: white;")
                    del_btn.clicked.connect(lambda _, rid=row[0]: self.delete_expense(rid))

                    action_layout.addWidget(edit_btn)
                    action_layout.addWidget(del_btn)
                    self.table.setCellWidget(idx, 6, action_widget)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load expenses: {e}")

    def add_expense(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add New Expense")
        dialog.setMinimumWidth(420)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()

        category_combo = QComboBox()
        subcat_combo = QComboBox()
        amount_input = QLineEdit()
        amount_input.setValidator(QDoubleValidator(0.0, 1e9, 2))
        date_input = QDateEdit()
        date_input.setDate(QDate.currentDate())
        date_input.setCalendarPopup(True)
        desc_input = QLineEdit()

        # Populate categories of type 'Expenses'
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT CatID, CatName FROM Category WHERE Type='Expenses'")
                cats = c.fetchall()
                category_combo.addItem("Select Category", None)
                for cat in cats:
                    category_combo.addItem(cat[1], cat[0])
        except Exception as e:
            print(f"Error loading categories: {e}")

        def update_subcats():
            subcat_combo.clear()
            subcat_combo.addItem("Select Subcategory", None)
            cat_id = category_combo.currentData()
            if cat_id:
                try:
                    with get_db_connection() as (conn, c):
                        c.execute("SELECT SubCatID, SubCategoryName FROM SubCategory WHERE CategoryID = ?", (cat_id,))
                        subcats = c.fetchall()
                        for s in subcats:
                            subcat_combo.addItem(s[1], s[0])
                except Exception as e:
                    print(e)

        category_combo.currentIndexChanged.connect(update_subcats)

        form.addRow("Category:", category_combo)
        form.addRow("Subcategory:", subcat_combo)
        form.addRow("Amount:", amount_input)
        form.addRow("Date:", date_input)
        form.addRow("Description:", desc_input)

        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_expense(dialog, subcat_combo, amount_input, date_input, desc_input))

        layout.addLayout(form)
        layout.addWidget(save_btn)
        dialog.exec_()

    def save_expense(self, dialog, subcat_combo, amount_input, date_input, desc_input):
        subcat_id = subcat_combo.currentData()
        amount_text = amount_input.text().strip()
        date_text = date_input.date().toString("yyyy-MM-dd")
        desc = desc_input.text().strip()

        if not subcat_id or not amount_text:
            QMessageBox.warning(dialog, "Missing Data", "Please select subcategory and enter amount.")
            return

        try:
            amount = float(amount_text)
            with get_db_connection() as (conn, c):
                # ensure column exists on older DBs
                self._ensure_expense_description_column(c)
                c.execute("INSERT INTO Expenses (ExpSubCategory, ExpAmount, ExpDate, ExpDescription) VALUES (?, ?, ?, ?)",
                          (subcat_id, amount, date_text, desc))
                conn.commit()
            dialog.accept()
            self.load_data()
            QMessageBox.information(self, "Success", "Expense saved.")
        except Exception as e:
            QMessageBox.critical(dialog, "Error", f"Failed to save expense: {e}")

    def edit_expense(self, exp_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT ExpSubCategory, ExpAmount, ExpDate, ExpDescription FROM Expenses WHERE ExpID=?", (exp_id,))
                row = c.fetchone()
            if not row:
                return

            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Expense")
            dialog.setMinimumWidth(420)
            layout = QVBoxLayout(dialog)
            form = QFormLayout()

            category_combo = QComboBox()
            subcat_combo = QComboBox()
            amount_input = QLineEdit(str(row[1]))
            amount_input.setValidator(QDoubleValidator(0.0, 1e9, 2))
            date_input = QDateEdit()
            date_input.setDate(QDate.fromString(row[2], "yyyy-MM-dd"))
            date_input.setCalendarPopup(True)
            desc_input = QLineEdit(row[3] if row[3] else "")

            # Populate categories and subcats, select values
            with get_db_connection() as (conn, c):
                c.execute("SELECT CatID, CatName FROM Category WHERE Type='Expenses'")
                cats = c.fetchall()
                category_combo.addItem("Select Category", None)
                for cat in cats:
                    category_combo.addItem(cat[1], cat[0])

            def update_subcats():
                subcat_combo.clear()
                subcat_combo.addItem("Select Subcategory", None)
                cat_id = category_combo.currentData()
                if cat_id:
                    with get_db_connection() as (conn, c):
                        c.execute("SELECT SubCatID, SubCategoryName FROM SubCategory WHERE CategoryID = ?", (cat_id,))
                        subcats = c.fetchall()
                        for s in subcats:
                            subcat_combo.addItem(s[1], s[0])

            category_combo.currentIndexChanged.connect(update_subcats)

            # find category id for the existing subcat
            existing_subcat = row[0]
            if existing_subcat:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT CategoryID FROM SubCategory WHERE SubCatID = ?", (existing_subcat,))
                    cat_row = c.fetchone()
                    if cat_row:
                        # select category
                        cat_id = cat_row[0]
                        for i in range(category_combo.count()):
                            if category_combo.itemData(i) == cat_id:
                                category_combo.setCurrentIndex(i)
                                break
                        # populate subcats then select
                        update_subcats()
                        for j in range(subcat_combo.count()):
                            if subcat_combo.itemData(j) == existing_subcat:
                                subcat_combo.setCurrentIndex(j)
                                break

            form.addRow("Category:", category_combo)
            form.addRow("Subcategory:", subcat_combo)
            form.addRow("Amount:", amount_input)
            form.addRow("Date:", date_input)
            form.addRow("Description:", desc_input)

            save_btn = QPushButton("Update")
            save_btn.clicked.connect(lambda: self.update_expense(dialog, exp_id, subcat_combo, amount_input, date_input, desc_input))

            layout.addLayout(form)
            layout.addWidget(save_btn)
            dialog.exec_()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error preparing edit: {e}")

    def update_expense(self, dialog, exp_id, subcat_combo, amount_input, date_input, desc_input):
        subcat_id = subcat_combo.currentData()
        amount_text = amount_input.text().strip()
        date_text = date_input.date().toString("yyyy-MM-dd")
        desc = desc_input.text().strip()

        if not subcat_id or not amount_text:
            QMessageBox.warning(dialog, "Missing Data", "Please select subcategory and enter amount.")
            return

        try:
            amount = float(amount_text)
            with get_db_connection() as (conn, c):
                # ensure column exists on older DBs
                self._ensure_expense_description_column(c)
                c.execute("UPDATE Expenses SET ExpSubCategory=?, ExpAmount=?, ExpDate=?, ExpDescription=? WHERE ExpID=?",
                          (subcat_id, amount, date_text, desc, exp_id))
                conn.commit()
            dialog.accept()
            self.load_data()
            QMessageBox.information(self, "Success", "Expense updated.")
        except Exception as e:
            QMessageBox.critical(dialog, "Error", f"Failed to update expense: {e}")

    def delete_expense(self, exp_id):
        reply = QMessageBox.question(self, "Confirm Delete", "Are you sure you want to delete this expense?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    c.execute("DELETE FROM Expenses WHERE ExpID=?", (exp_id,))
                    conn.commit()
                self.load_data()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Delete failed: {e}")

class Assets(QWidget):
    """Simple Assets management UI with CRUD and record popup."""
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)

        self.header = QHBoxLayout()
        self.Heading = QLabel("Assets")
        self.Heading.setStyleSheet("font-size:18px;font-weight:bold;")
        self.searchbar = QLineEdit()
        self.searchbar.setPlaceholderText("Search Assets...")
        self.searchbar.textChanged.connect(self.load_data)
        self.addbtn = QPushButton("Add Asset")
        self.addbtn.clicked.connect(self.add_asset)

        self.header.addWidget(self.Heading)
        self.header.addStretch()
        self.header.addWidget(self.searchbar)
        self.header.addWidget(self.addbtn)
        layout.addLayout(self.header)

        self.table = QTableWidget()
        # Hidden AssetID, Name, Subcategory, Value, Status, Last_update, Actions
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["ID","Name","Subcategory","Value","Status","Last Update","Actions"])
        self.table.hideColumn(0)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        layout.addWidget(self.table)

        self.load_data()

    def load_data(self):
        try:
            with get_db_connection() as (conn, c):
                query = (
                    "SELECT a.AssetID, a.AssetName, sc.SubCategoryName, a.AssetValue, a.Status, a.Last_update "
                    "FROM Assets a LEFT JOIN SubCategory sc ON a.Subcategory = sc.SubCatID"
                )
                params = []
                s = self.searchbar.text().strip()
                if s:
                    query += " WHERE a.AssetName LIKE ? OR sc.SubCategoryName LIKE ? OR a.Status LIKE ?"
                    params.extend([f"%{s}%", f"%{s}%", f"%{s}%"])

                query += " ORDER BY a.Last_update DESC"
                c.execute(query, params)
                rows = c.fetchall()

                self.table.setRowCount(0)
                for idx, row in enumerate(rows):
                    self.table.insertRow(idx)
                    self.table.setItem(idx, 0, QTableWidgetItem(str(row[0] if row[0] is not None else "")))
                    self.table.setItem(idx, 1, QTableWidgetItem(str(row[1] if row[1] else "")))
                    self.table.setItem(idx, 2, QTableWidgetItem(str(row[2] if row[2] else "")))
                    self.table.setItem(idx, 3, QTableWidgetItem(str(row[3] if row[3] is not None else "")))
                    self.table.setItem(idx, 4, QTableWidgetItem(str(row[4] if row[4] else "")))
                    self.table.setItem(idx, 5, QTableWidgetItem(str(row[5] if row[5] else "")))

                    action_widget = QWidget()
                    action_layout = QHBoxLayout(action_widget)
                    action_layout.setContentsMargins(5, 2, 5, 2)

                    rec_btn = QPushButton("Record")
                    rec_btn.setStyleSheet("background-color: #805AD5; color: white;")
                    rec_btn.clicked.connect(lambda _, aid=row[0], aname=row[1]: self.show_records(aid, aname))

                    edit_btn = QPushButton("Edit")
                    edit_btn.setStyleSheet("background-color: #38A169; color: white;")
                    edit_btn.clicked.connect(lambda _, aid=row[0]: self.edit_asset(aid))

                    del_btn = QPushButton("Delete")
                    del_btn.setStyleSheet("background-color: #E53E3E; color: white;")
                    del_btn.clicked.connect(lambda _, aid=row[0]: self.delete_asset(aid))

                    action_layout.addWidget(rec_btn)
                    action_layout.addWidget(edit_btn)
                    action_layout.addWidget(del_btn)
                    self.table.setCellWidget(idx, 6, action_widget)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load assets: {e}")

    def add_asset(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Asset")
        dialog.setMinimumWidth(500)
        layout = QVBoxLayout(dialog)
        form = QFormLayout()

        name_input = QLineEdit()
        dept_combo = QComboBox()
        cat_combo = QComboBox()
        subcat_combo = QComboBox()
        value_input = QLineEdit()
        value_input.setValidator(QDoubleValidator(0.0, 1e12, 2))
        status_combo = QComboBox()
        status_combo.addItems(["Active", "Disposed", "Maintenance", "Unknown"])
        last_update = QDateEdit()
        last_update.setDate(QDate.currentDate())
        last_update.setCalendarPopup(True)

        # Populate Departments
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT DeptID, DeptName FROM Department")
                depts = c.fetchall()
                dept_combo.addItem("Select Department", None)
                for d in depts:
                    dept_combo.addItem(d[1], d[0])
        except Exception:
            pass

        def update_cats():
            cat_combo.clear()
            cat_combo.addItem("Select Category", None)
            dept_id = dept_combo.currentData()
            if dept_id:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT CatID, CatName FROM Category WHERE Department = ? AND Type='Asset'", (dept_id,))
                    cats = c.fetchall()
                    for cat in cats:
                        cat_combo.addItem(cat[1], cat[0])

        def update_subcats():
            subcat_combo.clear()
            subcat_combo.addItem("Select Subcategory", None)
            cat_id = cat_combo.currentData()
            if cat_id:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT SubCatID, SubCategoryName FROM SubCategory WHERE CategoryID = ?", (cat_id,))
                    subs = c.fetchall()
                    for s in subs:
                        subcat_combo.addItem(s[1], s[0])

        dept_combo.currentIndexChanged.connect(update_cats)
        cat_combo.currentIndexChanged.connect(update_subcats)

        form.addRow("Asset Name:", name_input)
        form.addRow("Department:", dept_combo)
        form.addRow("Category:", cat_combo)
        form.addRow("Subcategory:", subcat_combo)
        form.addRow("Value:", value_input)
        form.addRow("Status:", status_combo)
        form.addRow("Last Update:", last_update)

        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_asset(dialog, name_input, subcat_combo, value_input, status_combo, last_update))

        layout.addLayout(form)
        layout.addWidget(save_btn)
        dialog.exec_()

    def save_asset(self, dialog, name_input, subcat_combo, value_input, status_combo, last_update):
        name = name_input.text().strip()
        subcat = subcat_combo.currentData()
        value_text = value_input.text().strip()
        status = status_combo.currentText()
        last_date = last_update.date().toString("yyyy-MM-dd")

        if not all([name, subcat, value_text]):
            QMessageBox.warning(dialog, "Missing Data", "Please fill name, subcategory and value.")
            return

        try:
            value = float(value_text)
            with get_db_connection() as (conn, c):
                c.execute("INSERT INTO Assets (AssetName, Subcategory, AssetValue, Status, Last_update) VALUES (?, ?, ?, ?, ?)",
                          (name, subcat, value, status, last_date))
                conn.commit()
            dialog.accept()
            self.load_data()
            QMessageBox.information(self, "Success", "Asset added.")
        except Exception as e:
            QMessageBox.critical(dialog, "Error", f"Failed to save asset: {e}")

    def edit_asset(self, asset_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT AssetName, Subcategory, AssetValue, Status, Last_update FROM Assets WHERE AssetID=?", (asset_id,))
                row = c.fetchone()
            if not row:
                return

            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Asset")
            dialog.setMinimumWidth(500)
            layout = QVBoxLayout(dialog)
            form = QFormLayout()

            name_input = QLineEdit(row[0])
            dept_combo = QComboBox()
            cat_combo = QComboBox()
            subcat_combo = QComboBox()
            value_input = QLineEdit(str(row[2]))
            value_input.setValidator(QDoubleValidator(0.0, 1e12, 2))
            status_combo = QComboBox()
            status_combo.addItems(["Active", "Disposed", "Maintenance", "Unknown"])
            status_combo.setCurrentText(row[3] if row[3] else "Active")
            last_update = QDateEdit()
            last_update.setDate(QDate.fromString(row[4], "yyyy-MM-dd") if row[4] else QDate.currentDate())
            last_update.setCalendarPopup(True)

            # populate departments and categories similar to add
            try:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT DeptID, DeptName FROM Department")
                    depts = c.fetchall()
                    dept_combo.addItem("Select Department", None)
                    for d in depts:
                        dept_combo.addItem(d[1], d[0])
            except Exception:
                pass

            def update_cats():
                cat_combo.clear()
                cat_combo.addItem("Select Category", None)
                dept_id = dept_combo.currentData()
                if dept_id:
                    with get_db_connection() as (conn, c):
                        c.execute("SELECT CatID, CatName FROM Category WHERE Department = ? AND Type='Asset'", (dept_id,))
                        cats = c.fetchall()
                        for cat in cats:
                            cat_combo.addItem(cat[1], cat[0])

            def update_subcats():
                subcat_combo.clear()
                subcat_combo.addItem("Select Subcategory", None)
                cat_id = cat_combo.currentData()
                if cat_id:
                    with get_db_connection() as (conn, c):
                        c.execute("SELECT SubCatID, SubCategoryName FROM SubCategory WHERE CategoryID = ?", (cat_id,))
                        subs = c.fetchall()
                        for s in subs:
                            subcat_combo.addItem(s[1], s[0])

            dept_combo.currentIndexChanged.connect(update_cats)
            cat_combo.currentIndexChanged.connect(update_subcats)

            # Attempt to pre-select category/department based on existing subcategory
            existing_subcat = row[1]
            if existing_subcat:
                try:
                    with get_db_connection() as (conn, c):
                        c.execute("SELECT CategoryID FROM SubCategory WHERE SubCatID = ?", (existing_subcat,))
                        cres = c.fetchone()
                        if cres:
                            cat_id = cres[0]
                            # find department for this category
                            c.execute("SELECT Department FROM Category WHERE CatID = ?", (cat_id,))
                            dres = c.fetchone()
                            if dres:
                                dept_id = dres[0]
                                for i in range(dept_combo.count()):
                                    if dept_combo.itemData(i) == dept_id:
                                        dept_combo.setCurrentIndex(i)
                                        break
                                update_cats()
                                for j in range(cat_combo.count()):
                                    if cat_combo.itemData(j) == cat_id:
                                        cat_combo.setCurrentIndex(j)
                                        break
                                update_subcats()
                                for k in range(subcat_combo.count()):
                                    if subcat_combo.itemData(k) == existing_subcat:
                                        subcat_combo.setCurrentIndex(k)
                                        break
                except Exception:
                    pass

            form.addRow("Asset Name:", name_input)
            form.addRow("Department:", dept_combo)
            form.addRow("Category:", cat_combo)
            form.addRow("Subcategory:", subcat_combo)
            form.addRow("Value:", value_input)
            form.addRow("Status:", status_combo)
            form.addRow("Last Update:", last_update)

            save_btn = QPushButton("Update")
            save_btn.clicked.connect(lambda: self.update_asset(dialog, asset_id, name_input, subcat_combo, value_input, status_combo, last_update))

            layout.addLayout(form)
            layout.addWidget(save_btn)
            dialog.exec_()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error preparing edit: {e}")

    def update_asset(self, dialog, asset_id, name_input, subcat_combo, value_input, status_combo, last_update):
        name = name_input.text().strip()
        subcat = subcat_combo.currentData()
        value_text = value_input.text().strip()
        status = status_combo.currentText()
        last_date = last_update.date().toString("yyyy-MM-dd")

        if not all([name, subcat, value_text]):
            QMessageBox.warning(dialog, "Missing Data", "Please fill name, subcategory and value.")
            return

        try:
            value = float(value_text)
            with get_db_connection() as (conn, c):
                c.execute("UPDATE Assets SET AssetName=?, Subcategory=?, AssetValue=?, Status=?, Last_update=? WHERE AssetID=?",
                          (name, subcat, value, status, last_date, asset_id))
                conn.commit()
            dialog.accept()
            self.load_data()
            QMessageBox.information(self, "Success", "Asset updated.")
        except Exception as e:
            QMessageBox.critical(dialog, "Error", f"Failed to update asset: {e}")

    def delete_asset(self, asset_id):
        reply = QMessageBox.question(self, "Confirm Delete", "Delete this asset and its records?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    c.execute("DELETE FROM AssetsRecord WHERE AssetsID=?", (asset_id,))
                    c.execute("DELETE FROM Assets WHERE AssetID=?", (asset_id,))
                    conn.commit()
                self.load_data()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Delete failed: {e}")

    def show_records(self, asset_id, asset_name):
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Asset Records: {asset_name}")
        dialog.setMinimumWidth(500)
        layout = QVBoxLayout(dialog)

        list_widget = QListWidget()
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT AssetValue, Status, LastDate FROM AssetsRecord WHERE AssetsID=? ORDER BY LastDate DESC", (asset_id,))
                rows = c.fetchall()
                for r in rows:
                    val = r[0]
                    status = r[1]
                    dt = r[2]
                    item_text = f"{dt} — {status} — Value: {val}"
                    list_widget.addItem(item_text)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load records: {e}")

        layout.addWidget(list_widget)
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        dialog.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Expenditure()
    window.setWindowTitle("Expenditure Management")
    window.resize(800,600)
    window.show()
    sys.exit(app.exec_())