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
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter,landscape, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from utilis import resource_path
import re

DB_NAME = "Database.db"


def initialize_database():
    """
    Check if Database.db exists in the current working directory.
    If not, create it and apply the schema.
    """
    db_path = os.path.join(os.getcwd(), DB_NAME)

    if os.path.exists(db_path):
        # Database already exists → nothing to do
        return

    # Create database and schema
    conn = sqlite3.connect(db_path)
    try:
        cursor = conn.cursor()
        cursor.executescript(SCHEMA_SQL)
        conn.commit()
    finally:
        conn.close()

SCHEMA_SQL = """
    PRAGMA foreign_keys = ON;

    CREATE TABLE IF NOT EXISTS "AcademicYear"(
        "AcademicYearID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "Year" TEXT NOT NULL UNIQUE
    );

    CREATE TABLE IF NOT EXISTS "Class"(
        "ClassID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "ClassName" TEXT NOT NULL,
        "AcademicYear" TEXT NOT NULL,
        FOREIGN KEY("AcademicYear") REFERENCES "AcademicYear"("AcademicYearID")
    );

    CREATE TABLE IF NOT EXISTS "Term"(
        "TermID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "TermName" TEXT NOT NULL,
        "AcademicYear" TEXT NOT NULL,
        "StartDate" TEXT NOT NULL,
        "EndDate" TEXT NOT NULL,
        FOREIGN KEY("AcademicYear") REFERENCES "AcademicYear"("AcademicYearID")
    );

    CREATE TABLE IF NOT EXISTS "Students"(
        "StudentID" TEXT PRIMARY KEY,
        "Student_Name" TEXT NOT NULL,
        "FeePayer" TEXT NOT NULL,
        "PayRate" INTEGER NOT NULL,
        "Parent_Name" TEXT NOT NULL,
        "Parent_Contact" TEXT NOT NULL,
        "Parent_Contact2" TEXT
    );

    CREATE TABLE IF NOT EXISTS "StudentTerm"(
        "StudentTermID" INTEGER PRIMARY KEY,
        "StudentID" TEXT NOT NULL,
        "ClassID" INTEGER NOT NULL,
        "TermID" INTEGER NOT NULL,
        "PayRate" INTEGER NOT NULL CHECK(PayRate BETWEEN 0 AND 100),
        FOREIGN KEY("StudentID") REFERENCES "Students"("StudentID") ON DELETE CASCADE,
        FOREIGN KEY("ClassID") REFERENCES "Class"("ClassID"),
        FOREIGN KEY("TermID") REFERENCES "Term"("TermID"),
        UNIQUE(StudentID, ClassID, TermID)
    );

    CREATE TABLE IF NOT EXISTS "Fees"(
        "FeesID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "Class" TEXT NOT NULL,
        "Term" TEXT NOT NULL,
        "Amount" REAL NOT NULL
    );

    CREATE TABLE IF NOT EXISTS "FeesPayment"(
        "FeesPaymentID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "StudentTermID" INTEGER NOT NULL,
        "FeesID" INTEGER NOT NULL,
        "AmountPaid" REAL NOT NULL,
        "TransactionDate" TEXT NOT NULL,
        "ReferenceNumber" TEXT NOT NULL,
        "DoneOn" TEXT NOT NULL,
        FOREIGN KEY("StudentTermID") REFERENCES "StudentTerm"("StudentTermID") ON DELETE CASCADE,
        FOREIGN KEY("FeesID") REFERENCES "Fees"("FeesID")
    );

    CREATE TABLE IF NOT EXISTS "Uniform"(
        "UniformId" INTEGER PRIMARY KEY AUTOINCREMENT,
        "UniformName" TEXT NOT NULL,
        "UniformPrice" REAL NOT NULL,
        "Quantity" INTEGER NOT NULL
    );

    CREATE TABLE IF NOT EXISTS "UniformInventory"(
        "UniformInventoryID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "UniformID" INTEGER NOT NULL,
        "Type" TEXT NOT NULL CHECK(Type IN ('In','Out')),
        "Quantity" INTEGER NOT NULL CHECK(Quantity > 0),
        "Price" REAL NOT NULL CHECK(Price >= 0),
        "Date" TEXT NOT NULL,
        "ReferencesNotes" TEXT,
        FOREIGN KEY("UniformID") REFERENCES "Uniform"("UniformId")
    );

    CREATE TABLE IF NOT EXISTS "Department"(
        "DeptID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "DeptName" TEXT UNIQUE NOT NULL
    );

    CREATE TABLE IF NOT EXISTS "Category"(
        "CatID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "Type" TEXT NOT NULL CHECK(Type IN ('Supplies','Expenses')),
        "Department" INTEGER NOT NULL,
        "CatName" TEXT NOT NULL,
        UNIQUE(Type, Department, CatName),
        FOREIGN KEY("Department") REFERENCES "Department"("DeptID")
    );

    CREATE TABLE IF NOT EXISTS "SubCategory"(
        "SubCatID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "CategoryID" INTEGER NOT NULL,
        "SubCategoryName" TEXT NOT NULL,
        UNIQUE(CategoryID, SubCategoryName),
        FOREIGN KEY("CategoryID") REFERENCES "Category"("CatID") ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS "Supplies"(
        "SuppliesID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "Name" TEXT NOT NULL,
        "Subcategory" INTEGER NOT NULL,
        "Quantity" INTEGER NOT NULL,
        "Unit" TEXT NOT NULL,
        FOREIGN KEY("Subcategory") REFERENCES "SubCategory"("SubCatID")
    );

    CREATE TABLE IF NOT EXISTS "SuppliesInventory"(
        "Sup_Inv_ID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "Supplies_ID" INTEGER NOT NULL,
        "Type" TEXT NOT NULL CHECK(Type IN ('In','Out')),
        "Quantity" INTEGER NOT NULL CHECK(Quantity > 0),
        "Price" REAL NOT NULL CHECK(Price >= 0),
        "Date" TEXT NOT NULL,
        "ReferencesNotes" TEXT,
        FOREIGN KEY("Supplies_ID") REFERENCES "Supplies"("SuppliesID")
    );

    CREATE TABLE IF NOT EXISTS "Expenses"(
        "ExpID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "ExpSubCategory" INTEGER NOT NULL,
        "ExpAmount" REAL NOT NULL,
        "ExpDate" TEXT NOT NULL,
        "ExpReference" TEXT,
        "ExpDescription" TEXT,
        FOREIGN KEY("ExpSubCategory") REFERENCES "SubCategory"("SubCatID")
    );

    CREATE TABLE IF NOT EXISTS "UniformExpenses"(
        "U_ExpID" INTEGER PRIMARY KEY AUTOINCREMENT,
        "U_SubCategory" INTEGER NOT NULL,
        "U_ExpAmount" REAL NOT NULL,
        "U_ExpDate" TEXT NOT NULL,
        "U_ExpReference" TEXT,
        "U_ExpDescription" TEXT NOT NULL,
        FOREIGN KEY("U_SubCategory") REFERENCES "SubCategory"("SubCatID")
    );
    """

@contextmanager
def get_db_connection():
    """Context manager for database connections that handles opening and closing automatically"""
    conn = None
    try:
        # Ensure database is initialized
        initialize_database()
        
        conn = sqlite3.connect("Database.db")
        conn.row_factory = sqlite3.Row  # Optional: for named column access
        c = conn.cursor()
        
        # Enable foreign keys for every connection
        c.execute("PRAGMA foreign_keys = ON")
        
        yield conn, c
    except sqlite3.Error as e:
        QMessageBox.warning(None, "Database Error", f"Database connection error: {str(e)}")
        raise
    finally:
        if conn:
            conn.close()

### Dashboard ###
class Dashboard(QWidget):
    def __init__(self):
        super().__init__()
        self.finance_service = FinanceService()
        self.init_ui()
        self.load_terms()

    def init_ui(self):
        # Main Layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        # Scroll Area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        content_widget = QWidget()
        self.content_layout = QVBoxLayout(content_widget)
        self.content_layout.setSpacing(20)
        scroll.setWidget(content_widget)
        
        # --- Header Section ---
        header_layout = QHBoxLayout()
        
        # Title
        title_label = QLabel("School Finance Dashboard")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #2D3748;")
        
        # Term Selector
        self.term_selector = QComboBox()
        self.term_selector.setFixedWidth(200)
        self.term_selector.setStyleSheet("""
            QComboBox {
                padding: 8px;
                border: 1px solid #E2E8F0;
                border-radius: 5px;
                font-size: 14px;
            }
        """)
        self.term_selector.currentIndexChanged.connect(self.update_dashboard)
        
        # Current Date
        date_label = QLabel(datetime.now().strftime("%B %d, %Y"))
        date_label.setStyleSheet("font-size: 16px; color: #718096;")
        
        # Report Button
        report_btn = QPushButton("Financial Report")
        report_btn.setStyleSheet("""
            QPushButton {
                background-color: #2B6CB0;
                color: white;
                padding: 8px 16px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2C5282;
            }
        """)
        report_btn.clicked.connect(self.open_financial_report)
        
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(report_btn) # Add button here
        header_layout.addWidget(QLabel("Term:"))
        header_layout.addWidget(self.term_selector)
        header_layout.addWidget(date_label)
        
        self.content_layout.addLayout(header_layout)

        # --- Summary Cards Section ---
        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(15)
        
        self.card_fees_collected = self.create_card("Fees Collected", "0.00", "#3182CE") # Blue
        self.card_outstanding = self.create_card("Outstanding Fees", "0.00", "#DD6B20") # Orange
        self.card_expenses = self.create_card("Total Expenses", "0.00", "#E53E3E") # Red
        self.card_supplies = self.create_card("Supplies Cost", "0.00", "#805AD5") # Purple
        
        cards_layout.addWidget(self.card_fees_collected)
        cards_layout.addWidget(self.card_outstanding)
        cards_layout.addWidget(self.card_expenses)
        cards_layout.addWidget(self.card_supplies)
        
        self.content_layout.addLayout(cards_layout)

        # --- Fees Overview Section ---
        fees_section_layout = QHBoxLayout()
        
        # Fees Chart
        fees_chart_frame = QFrame()
        fees_chart_frame.setStyleSheet("background-color: white; border-radius: 8px; border: 1px solid #E2E8F0;")
        fees_chart_layout = QVBoxLayout(fees_chart_frame)
        fees_chart_label = QLabel("Fees Collection Overview")
        fees_chart_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #2D3748; border: none;")
        fees_chart_layout.addWidget(fees_chart_label)
        
        self.fees_figure = Figure(figsize=(5, 4), dpi=100)
        self.fees_canvas = FigureCanvas(self.fees_figure)
        fees_chart_layout.addWidget(self.fees_canvas)
        
        # Outstanding Fees Table
        outstanding_frame = QFrame()
        outstanding_frame.setStyleSheet("background-color: white; border-radius: 8px; border: 1px solid #E2E8F0;")
        outstanding_layout = QVBoxLayout(outstanding_frame)
        outstanding_label = QLabel("Top Outstanding Classes")
        outstanding_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #2D3748; border: none;")
        outstanding_layout.addWidget(outstanding_label)
        
        self.outstanding_table = QTableWidget()
        self.outstanding_table.setColumnCount(3)
        self.outstanding_table.setHorizontalHeaderLabels(["Class", "Balance", "%"])
        self.outstanding_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.outstanding_table.verticalHeader().setVisible(False)
        self.outstanding_table.setStyleSheet("border: none;")
        outstanding_layout.addWidget(self.outstanding_table)
        
        fees_section_layout.addWidget(fees_chart_frame, 1)
        fees_section_layout.addWidget(outstanding_frame, 1)
        
        self.content_layout.addLayout(fees_section_layout)

        # --- Expenses & Supplies Section ---
        bottom_section_layout = QHBoxLayout()
        
        # Expenses Chart
        exp_chart_frame = QFrame()
        exp_chart_frame.setStyleSheet("background-color: white; border-radius: 8px; border: 1px solid #E2E8F0;")
        exp_chart_layout = QVBoxLayout(exp_chart_frame)
        exp_chart_label = QLabel("Expenses by Category")
        exp_chart_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #2D3748; border: none;")
        exp_chart_layout.addWidget(exp_chart_label)
        
        self.exp_figure = Figure(figsize=(5, 4), dpi=100)
        self.exp_canvas = FigureCanvas(self.exp_figure)
        exp_chart_layout.addWidget(self.exp_canvas)
        
        # Supplies Summary
        supplies_frame = QFrame()
        supplies_frame.setStyleSheet("background-color: white; border-radius: 8px; border: 1px solid #E2E8F0;")
        supplies_layout = QVBoxLayout(supplies_frame)
        supplies_label = QLabel("Supplies Summary (Top 5)")
        supplies_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #2D3748; border: none;")
        supplies_layout.addWidget(supplies_label)
        
        self.supplies_table = QTableWidget()
        self.supplies_table.setColumnCount(3)
        self.supplies_table.setHorizontalHeaderLabels(["Item", "Qty", "Cost"])
        self.supplies_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.supplies_table.verticalHeader().setVisible(False)
        self.supplies_table.setStyleSheet("border: none;")
        supplies_layout.addWidget(self.supplies_table)
        
        bottom_section_layout.addWidget(exp_chart_frame, 1)
        bottom_section_layout.addWidget(supplies_frame, 1)
        
        self.content_layout.addLayout(bottom_section_layout)
        
        main_layout.addWidget(scroll)

    def create_card(self, title, value, color):
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 8px;
                border-left: 5px solid {color};
                border-top: 1px solid #E2E8F0;
                border-right: 1px solid #E2E8F0;
                border-bottom: 1px solid #E2E8F0;
            }}
        """)
        layout = QVBoxLayout(card)
        
        title_lbl = QLabel(title)
        title_lbl.setStyleSheet("color: #718096; font-size: 14px; border: none;")
        
        self.value_lbl = QLabel(value)
        self.value_lbl.setStyleSheet("color: #2D3748; font-size: 24px; font-weight: bold; border: none;")
        # Store reference to value label to update it later
        card.value_label = self.value_lbl 
        
        layout.addWidget(title_lbl)
        layout.addWidget(self.value_lbl)
        return card

    def open_financial_report(self):
        dialog = FinancialReportDialog(self)
        dialog.exec_()

    def load_terms(self):
        self.term_selector.clear()
        try:
            terms = self.finance_service.get_all_terms()
            for term in terms:
                self.term_selector.addItem(f"{term['TermName']} ({term['Year']})", term['TermID'])
        except Exception as e:
            QMessageBox.warning(f"Error loading terms: {e}")

    def update_dashboard(self):
        term_id = self.term_selector.currentData()
        if not term_id:
            return

        try:
            # Get Term Dates
            term_dates = self.finance_service.get_term_dates(term_id)
            if not term_dates:
                return
            start_date = term_dates['StartDate']
            end_date = term_dates['EndDate']

            # --- 1. Fees Data ---
            fees_summary = self.finance_service.get_term_fees_summary(term_id)
            total_collected = fees_summary['collected']
            outstanding = fees_summary['outstanding']
            
            # Update Cards
            self.card_fees_collected.value_label.setText(f"{total_collected:,.2f}")
            self.card_outstanding.value_label.setText(f"{outstanding:,.2f}")

            # --- 2. Expenses Data ---
            total_expenses = self.finance_service.get_term_expenses(start_date, end_date)
            self.card_expenses.value_label.setText(f"{total_expenses:,.2f}")

            # --- 3. Supplies Data ---
            total_supplies = self.finance_service.get_term_supplies(start_date, end_date)
            self.card_supplies.value_label.setText(f"{total_supplies:,.2f}")

            # --- 5. Charts & Tables ---
            self.update_fees_chart(total_collected, outstanding)
            self.update_expenses_chart(start_date, end_date)
            self.update_outstanding_table(term_id)
            self.update_supplies_table(start_date, end_date)

        except Exception as e:
            QMessageBox.warning(f"Error updating dashboard: {e}")

    def update_fees_chart(self, collected, outstanding):
        self.fees_figure.clear()
        ax = self.fees_figure.add_subplot(111)
        
        if collected == 0 and outstanding == 0:
            labels = ['No Data']
            sizes = [1]
            colors = ['#E2E8F0']
        else:
            labels = ['Collected', 'Outstanding']
            sizes = [collected, outstanding]
            colors = ['#48BB78', '#E53E3E']
            
        ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
        ax.axis('equal')
        self.fees_canvas.draw()

    def update_expenses_chart(self, start_date, end_date):
        self.exp_figure.clear()
        ax = self.exp_figure.add_subplot(111)
        
        try:
            data = self.finance_service.get_expenses_by_category(start_date, end_date)
            
            if data:
                categories = [row[0] for row in data]
                amounts = [row[1] for row in data]
                ax.bar(categories, amounts, color='#4299E1')
                ax.set_ylabel('Amount ()')
                plt.setp(ax.get_xticklabels(), rotation=45, ha="right")
            else:
                ax.text(0.5, 0.5, 'No Expenses', ha='center', va='center')
                    
        except Exception as e:
            QMessageBox.warning(f"Chart Error: {e}")
            
        self.exp_figure.tight_layout()
        self.exp_canvas.draw()

    def update_outstanding_table(self, term_id):
        self.outstanding_table.setRowCount(0)
        try:
            rows = self.finance_service.get_outstanding_fees_by_class(term_id)
            for row_idx, row_data in enumerate(rows):
                self.outstanding_table.insertRow(row_idx)
                self.outstanding_table.setItem(row_idx, 0, QTableWidgetItem(row_data['class_name']))
                self.outstanding_table.setItem(row_idx, 1, QTableWidgetItem(f"{row_data['balance']:,.2f}"))
                self.outstanding_table.setItem(row_idx, 2, QTableWidgetItem(f"{row_data['percentage']:.1f}%"))
        except Exception as e:
            QMessageBox.warning(f"Table Error: {e}")

    def update_supplies_table(self, start_date, end_date):
        self.supplies_table.setRowCount(0)
        try:
            rows = self.finance_service.get_top_supplies(start_date, end_date)
            for row_idx, row in enumerate(rows):
                self.supplies_table.insertRow(row_idx)
                self.supplies_table.setItem(row_idx, 0, QTableWidgetItem(str(row[0])))
                self.supplies_table.setItem(row_idx, 1, QTableWidgetItem(str(row[1])))
                self.supplies_table.setItem(row_idx, 2, QTableWidgetItem(f"{row[2]:,.2f}"))
        except Exception as e:
            QMessageBox.warning(f"Supplies Error: {e}")

###End Dashboard ###

### Student Management Initial Start ###
class StudentManagement(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(15)

        # Tab widget to separate Students and Enrollments
        self.tabs = QTabWidget()
        self.students_tab = StudentsTab()
        self.enrollments_tab = EnrollmentsTab()
        self.tabs.addTab(self.students_tab, "Students Info")
        self.tabs.addTab(self.enrollments_tab, "Student Enrollments")
        layout.addWidget(self.tabs)

class StudentsTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Filters and add buttons
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search...")
        self.search_input.setMaximumWidth(200)
        self.search_input.textChanged.connect(self.load_data)

        self.filterbyStatus = QComboBox()
        self.filterbyStatus.addItem("All PayRate",None)
        self.filterbyStatus.addItem("Full Scholarship",0)
        self.filterbyStatus.addItem("20% Scholarship",80)
        self.filterbyStatus.addItem("No Scholarship",100)
        self.filterbyStatus.currentIndexChanged.connect(self.load_data)

        self.Filter_FeePayer = QComboBox()
        self.Filter_FeePayer.addItem("All Fee Payers")
        with get_db_connection() as (conn, c):
            c.execute("SELECT DISTINCT FeePayer From Students")
            for fp in c.fetchall():
                self.Filter_FeePayer.addItem(fp[0])
        self.Filter_FeePayer.currentIndexChanged.connect(self.load_data)

        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(self.filterbyStatus)
        filter_layout.addWidget(self.Filter_FeePayer)
        filter_layout.addStretch()
        add_student_btn = QPushButton("Add Student")
        add_student_btn.clicked.connect(self.show_add_student_dialog)
        Import_btn = QPushButton("Import")
        Import_btn.clicked.connect(show_import_wizard)
        filter_layout.addWidget(add_student_btn)
        filter_layout.addWidget(Import_btn)
        layout.addLayout(filter_layout)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(8)  # StudentID, Name, Fee Payer, Status, Parent Name, Contact, Contact2, Actions
        self.table.setHorizontalHeaderLabels(["StudentID", "Name", "Fee Payer", "Scholarship", "Parent Name", "Contact", "Contact2", "Actions"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        header = self.table.horizontalHeader()
        for col in range(7):
            header.setSectionResizeMode(col, QHeaderView.Stretch)
        header.setSectionResizeMode(7, QHeaderView.Fixed)
        header.resizeSection(7, 120)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        self.results_label = QLabel("0 results")
        self.results_label.setStyleSheet("font-size: 14px; color: #718096;")
        filter_layout.addWidget(self.results_label)

        self.load_data()

    def show_add_student_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Student")
        layout = QVBoxLayout()
        form_layout = QFormLayout()
        id_edit = QLineEdit()
        name_edit = QLineEdit()
        feepayer_edit = QLineEdit()
        status_combo = QComboBox()
        status_combo.addItem("Select PayRate",None)
        status_combo.addItem("Full Scholarship",0)
        status_combo.addItem("20% Scholarship",80)
        status_combo.addItem("No Scholarship",100)
        parent_name_edit = QLineEdit()
        contact_edit = QLineEdit()
        contact2_edit = QLineEdit()
        form_layout.addRow("Student ID:", id_edit)
        form_layout.addRow("Name:", name_edit)
        form_layout.addRow("Fee Payer:", feepayer_edit)
        form_layout.addRow("PayRate:", status_combo)
        form_layout.addRow("Parent Name:", parent_name_edit)
        form_layout.addRow("Parent Contact:", contact_edit)
        form_layout.addRow("Parent Contact 2:", contact2_edit)
        layout.addLayout(form_layout)
        buttons = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_student(dialog, id_edit, name_edit, feepayer_edit, status_combo, parent_name_edit, contact_edit, contact2_edit))
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        buttons.addWidget(save_btn)
        buttons.addWidget(close_btn)
        layout.addLayout(buttons)
        dialog.setLayout(layout)
        dialog.exec_()

    def save_student(self, dialog, id_edit, name_edit, feepayer_edit, status_combo, parent_name_edit, contact_edit, contact2_edit):
        student_id = id_edit.text().strip()
        name = name_edit.text().strip()
        feepayer = feepayer_edit.text().strip()
        status = status_combo.currentData()
        parent_name = parent_name_edit.text().strip()
        contact = contact_edit.text().strip()
        contact2 = contact2_edit.text().strip()
        if status_combo.currentIndex() == 0:
            QMessageBox.warning(dialog, "Error", "Please select a valid PayRate")
            return
        if not all([student_id, name, feepayer, parent_name, contact]):
            QMessageBox.warning(dialog, "Error", "Required fields are missing")
            return
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT COUNT(*) FROM Students WHERE StudentID = ?", (student_id,))
                if c.fetchone()[0] > 0:
                    QMessageBox.warning(dialog, "Error", "Duplicate Student ID")
                    return
                c.execute("""INSERT INTO Students (StudentID, Student_Name, FeePayer, PayRate, Parent_Name, Parent_Contact, Parent_Contact2)
                             VALUES (?, ?, ?, ?, ?, ?, ?)""", (student_id, name, feepayer, status, parent_name, contact, contact2))
                conn.commit()
                QMessageBox.information(dialog, "Success", "Student registered successfully")
                id_edit.clear()
                name_edit.clear()
                feepayer_edit.clear()
                parent_name_edit.clear()
                contact_edit.clear()
                contact2_edit.clear()
                status_combo.setCurrentIndex(0)
                self.load_data()
        except Exception as e:
            QMessageBox.warning(dialog, "Error", f"Failed to add student: {str(e)}")

    def load_data(self):
        try:
            with get_db_connection() as (conn, c):
                query = """SELECT StudentID, Student_Name, FeePayer, PayRate, Parent_Name, Parent_Contact, Parent_Contact2
                        FROM Students"""
                filters = []
                params = ()
                search_text = self.search_input.text().strip()
                if search_text:
                    filters.append("(StudentID LIKE ? OR Student_Name LIKE ? OR Parent_Name LIKE ? OR Parent_Contact LIKE ? OR Parent_Contact2 LIKE ?)")
                    params += (f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%")
                status_filter = self.filterbyStatus.currentData()
                if status_filter != None:
                    filters.append("PayRate = ?")
                    params += (status_filter,)
                fee_payer_filter = self.Filter_FeePayer.currentText()
                if fee_payer_filter != "All Fee Payers":
                    filters.append("FeePayer = ?")
                    params += (fee_payer_filter,)
                if filters:
                    query += " WHERE " + " AND ".join(filters)
                
                c.execute(query, params)
                rows = c.fetchall()
                
                # Set column headers including Actions column
                self.table.setColumnCount(8) 
                self.table.setHorizontalHeaderLabels([
                    "Student ID", "Student Name", "Fee Payer", "PayRate", 
                    "Parent Name", "Parent Contact", "Parent Contact2", "Actions"
                ])
                
                # Configure column resize modes with optimized widths
                header = self.table.horizontalHeader()
                
                # Set specific widths for columns
                header.setSectionResizeMode(0, QHeaderView.Fixed)  # Student ID
                header.resizeSection(0, 100)
                
                header.setSectionResizeMode(1, QHeaderView.Stretch)  # Student Name
                
                header.setSectionResizeMode(2, QHeaderView.Fixed)  # Fee Payer
                header.resizeSection(2, 120)
                
                header.setSectionResizeMode(3, QHeaderView.Fixed)  # PayRate
                header.resizeSection(3, 80)
                
                header.setSectionResizeMode(4, QHeaderView.Stretch)  # Parent Name
                
                header.setSectionResizeMode(5, QHeaderView.Fixed)  # Parent Contact
                header.resizeSection(5, 130)
                
                header.setSectionResizeMode(6, QHeaderView.Fixed)  # Parent Contact2
                header.resizeSection(6, 130)
                
                # Actions column - Fixed width to ensure combobox fits perfectly
                header.setSectionResizeMode(7, QHeaderView.Fixed)
                header.resizeSection(7, 160)  # Ensure enough space for combobox
                
                self.table.setRowCount(0)
                for row_data in rows:
                    row = self.table.rowCount()
                    self.table.insertRow(row)
                    self.table.setRowHeight(row, 36)
                    
                    # Add data columns
                    for col, data in enumerate(row_data):
                        item = QTableWidgetItem(str(data) if data else "0")
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        self.table.setItem(row, col, item)
                    
                    # Create action ComboBox
                    student_id = row_data[0]

                    action_combo = QComboBox()
                    action_combo.addItems(["Select Action", "Edit", "Delete"])
                    action_combo.setProperty("role", "actions")


                    # Connect directly with lambda
                    action_combo.currentIndexChanged.connect(
                        lambda index, sid=student_id, cb=action_combo: self.on_action_selected(index, sid, cb)
                    )
                    # Set the ComboBox in the cell
                    self.table.setCellWidget(row, 7, action_combo)
                
                # Update results count
                count = self.table.rowCount()
                self.results_label.setText(f"{count} results")
                
                # Ensure column widths are applied properly
                self.table.horizontalHeader().setStretchLastSection(False)
                
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")

    def on_action_selected(self, index, student_id, combo_box):
        """Handle action selection directly"""
        if index == 0:  # "Select Action"
            return
        
        action = combo_box.itemText(index)
        combo_box.setCurrentIndex(0)  # Reset to "Select Action"
        
        if action == "Edit":
            self.edit_student(student_id)
        elif action == "Delete":
            self.delete_student(student_id)

    def edit_student(self, student_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT * FROM Students WHERE StudentID = ?", (student_id,))
                student_data = c.fetchone()
            if student_data:
                dialog = QDialog(self)
                dialog.setWindowTitle("Edit Student")
                layout = QVBoxLayout()
                form_layout = QFormLayout()
                name_edit = QLineEdit(student_data[1])
                feepayer_edit = QLineEdit(student_data[2])
                status_combo = QComboBox()
                status_combo.addItem("Select PayRate",None)
                status_combo.addItem("Full Scholarship",0)
                status_combo.addItem("20% Scholarship",80)
                status_combo.addItem("No Scholarship",100)
                status_combo.setCurrentText(status_combo.itemText(status_combo.findData(student_data[3])))
                parent_name_edit = QLineEdit(student_data[4])
                contact_edit = QLineEdit(student_data[5])
                contact2_edit = QLineEdit(student_data[6] if student_data[6] else "")
                form_layout.addRow("Name:", name_edit)
                form_layout.addRow("Fee Payer:", feepayer_edit)
                form_layout.addRow("PayRate:", status_combo)
                form_layout.addRow("Parent Name:", parent_name_edit)
                form_layout.addRow("Parent Contact:", contact_edit)
                form_layout.addRow("Parent Contact 2:", contact2_edit)
                layout.addLayout(form_layout)
                buttons = QHBoxLayout()
                save_btn = QPushButton("Save")
                save_btn.clicked.connect(lambda: self.save_edited_student(dialog, student_id, name_edit, feepayer_edit, status_combo, parent_name_edit, contact_edit, contact2_edit))
                close_btn = QPushButton("Close")
                close_btn.clicked.connect(dialog.close)
                buttons.addWidget(save_btn)
                buttons.addWidget(close_btn)
                layout.addLayout(buttons)
                dialog.setLayout(layout)
                dialog.exec_()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load student: {str(e)}")

    def save_edited_student(self, dialog, student_id, name_edit, feepayer_edit, status_combo, parent_name_edit, contact_edit, contact2_edit):
        name = name_edit.text().strip()
        feepayer = feepayer_edit.text().strip()
        status = status_combo.currentData()
        parent_name = parent_name_edit.text().strip()
        contact = contact_edit.text().strip()
        contact2 = contact2_edit.text().strip()
        if not all([name, feepayer, status, parent_name, contact]):
            QMessageBox.warning(dialog, "Error", "Required fields are missing")
            return
        try:
            with get_db_connection() as (conn, c):
                c.execute("""UPDATE Students 
                             SET Student_Name = ?, FeePayer = ?, PayRate = ?, 
                                 Parent_Name = ?, Parent_Contact = ?, Parent_Contact2 = ?
                             WHERE StudentID = ?""", (name, feepayer, status, parent_name, contact, contact2, student_id))
                conn.commit()
                QMessageBox.information(dialog, "Success", "Student updated successfully")
                dialog.close()
                self.load_data()
        except Exception as e:
            QMessageBox.warning(dialog, "Error", f"Failed to update student: {str(e)}")

    def delete_student(self, student_id):
        reply = QMessageBox.question(self, 'Confirm Deletion', 
                                     f'Are you sure you want to delete student {student_id}? This will also delete all enrollments.',
                                     QMessageBox.Yes | QMessageBox.No, 
                                     QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    c.execute("DELETE FROM StudentTerm WHERE StudentID = ?", (student_id,))
                    c.execute("DELETE FROM Students WHERE StudentID = ?", (student_id,))
                    conn.commit()
                    QMessageBox.information(self, "Success", "Student deleted successfully")
                    self.load_data()
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to delete student: {str(e)}")

class EnrollmentsTab(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Filters and add buttons
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search...")
        self.search_input.setMaximumWidth(200)
        self.search_input.textChanged.connect(self.load_data)

        self.AcademicYear_filter = QComboBox()
        with get_db_connection() as (conn,c):
            c.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY YEAR DESC")
            for years in c.fetchall():
                self.AcademicYear_filter.addItem(years[1],years[0])
            if c.fetchall() is None:
                self.AcademicYear_filter.addItem("No years",None)
        self.AcademicYear_filter.currentIndexChanged.connect(self.load_data)



        self.class_filter = QComboBox()
        self.class_filter.addItem("All Classes", None)
        with get_db_connection() as (conn, c):
            c.execute("SELECT ClassID, ClassName FROM Class")
            for cl in c.fetchall():
                self.class_filter.addItem(cl[1], cl[0])
        self.class_filter.currentIndexChanged.connect(self.load_data)

        self.term_filter = QComboBox()
        self.term_filter.addItem("All Terms", None)
        with get_db_connection() as (conn, c):
            c.execute("SELECT TermID, TermName FROM Term")
            for t in c.fetchall():
                self.term_filter.addItem(t[1], t[0])
        self.term_filter.currentIndexChanged.connect(self.load_data)
        self.AcademicYear_filter.currentIndexChanged.connect(self.update_filters)

        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(self.AcademicYear_filter)
        filter_layout.addWidget(self.class_filter)
        filter_layout.addWidget(self.term_filter)
        filter_layout.addStretch()
        add_enrollment_btn = QPushButton("Add Enrollment")
        add_enrollment_btn.clicked.connect(self.show_add_enrollment_dialog)
        import_btn = QPushButton("Import")
        import_btn.clicked.connect(show_import_wizard)
        filter_layout.addWidget(add_enrollment_btn)
        filter_layout.addWidget(import_btn)
        layout.addLayout(filter_layout)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(6)  # StudentID, Name, Class, Term, Actions
        self.table.setHorizontalHeaderLabels(["StudentID", "Name", "Class", "Term", "PayRate", "Actions"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        header = self.table.horizontalHeader()
        for col in range(4):
            header.setSectionResizeMode(col, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.resizeSection(4, 120)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        self.results_label = QLabel("0 results")
        self.results_label.setStyleSheet("font-size: 14px; color: #718096;")
        filter_layout.addWidget(self.results_label)

        self.load_data()
    
    def update_filters(self):
        """Update class and term filters when academic year changes"""
        year_id = self.AcademicYear_filter.currentData()
        
        # Update TERM dropdown (your exact requirement)
        self.term_filter.clear()
        self.term_filter.addItem("All Terms", None)
        
        if year_id:
            with get_db_connection() as (conn, c):
                # Simple query: get terms for selected academic year
                c.execute("SELECT TermID, TermName FROM Term WHERE AcademicYear = ?", (year_id,))
                for term_id, term_name in c.fetchall():
                    self.term_filter.addItem(term_name, term_id)
        
        # Update CLASS dropdown (optional - if you want it too)
        self.class_filter.clear()
        self.class_filter.addItem("All Classes", None)
        
        if year_id:
            with get_db_connection() as (conn, c):
                # Get classes that have students in this academic year
                c.execute("""
                    SELECT DISTINCT c.ClassID, c.ClassName 
                    FROM Class c
                    WHERE EXISTS (
                        SELECT 1 FROM StudentTerm s
                        JOIN Term t ON s.TermID = t.TermID
                        WHERE s.ClassID = c.ClassID AND t.AcademicYear = ?
                    )
                    ORDER BY c.ClassName
                """, (year_id,))
                
                for class_id, class_name in c.fetchall():
                    self.class_filter.addItem(class_name, class_id)

    def show_add_enrollment_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Enrollment")
        layout = QVBoxLayout()
        form_layout = QFormLayout()
        student_id_edit = QLineEdit()
        class_combo = QComboBox()
        term_combo = QComboBox()
        self.populate_combos(class_combo, term_combo)
        form_layout.addRow("Student ID:", student_id_edit)
        form_layout.addRow("Class:", class_combo)
        form_layout.addRow("Term:", term_combo)
        layout.addLayout(form_layout)
        buttons = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_enrollment(dialog, student_id_edit, class_combo, term_combo))
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        buttons.addWidget(save_btn)
        buttons.addWidget(close_btn)
        layout.addLayout(buttons)
        dialog.setLayout(layout)
        dialog.exec_()

    def save_enrollment(self, dialog, student_id_edit, class_combo, term_combo):
        student_id = student_id_edit.text().strip()
        class_id = class_combo.currentData()
        term_id = term_combo.currentData()

        if not all([student_id, class_id, term_id]):
            QMessageBox.warning(dialog, "Error", "Required fields are missing")
            return
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT COUNT(*) FROM Students WHERE StudentID = ?", (student_id,))
                if c.fetchone()[0] == 0:
                    QMessageBox.warning(dialog, "Error", "Student ID does not exist")
                    return
                c.execute("SELECT PayRate FROM Students WHERE StudentID = ?",(student_id,))
                Payrate = c.fetchone()[0]
                c.execute("SELECT COUNT(*) FROM StudentTerm WHERE StudentID = ? AND ClassID = ? AND TermID = ?", (student_id, class_id, term_id))
                if c.fetchone()[0] > 0:
                    QMessageBox.warning(dialog, "Error", "Duplicate enrollment for this class and term")
                    return
                c.execute("""INSERT INTO StudentTerm (StudentID, ClassID, TermID, PayRate)
                             VALUES (?, ?, ?, ?)""", (student_id, class_id, term_id,Payrate))
                conn.commit()
                QMessageBox.information(dialog, "Success", "Enrollment added successfully")
                student_id_edit.clear()
                class_combo.setCurrentIndex(0)
                term_combo.setCurrentIndex(0)
                self.load_data()
        except Exception as e:
            QMessageBox.warning(dialog, "Error", f"Failed to add enrollment: {str(e)}")

    def load_data(self):
        try:
            with get_db_connection() as (conn, c):
                query = """SELECT st.StudentTermID, s.StudentID, s.Student_Name, 
                                cl.ClassName, t.TermName, s.FeePayer, st.PayRate
                        FROM StudentTerm st
                        LEFT JOIN Students s ON st.StudentID = s.StudentID
                        LEFT JOIN Class cl ON st.ClassID = cl.ClassID
                        LEFT JOIN Term t ON st.TermID = t.TermID """
                filters = []
                params = ()
                search_text = self.search_input.text().strip()
                academic_Year = self.AcademicYear_filter.currentData()
                if academic_Year:
                    filters.append("t.AcademicYear = ?")
                    params +=(academic_Year,)
                if search_text:
                    filters.append("(s.StudentID LIKE ? OR s.Student_Name LIKE ?)")
                    params += (f"%{search_text}%", f"%{search_text}%")
                class_id = self.class_filter.currentData()
                if class_id:
                    filters.append("cl.ClassID = ?")
                    params += (class_id,)
                term_id = self.term_filter.currentData()
                if term_id:
                    filters.append("t.TermID = ?")
                    params += (term_id,)
                if filters:
                    query += " WHERE " + " AND ".join(filters)
                c.execute(query, params)
                rows = c.fetchall()
                
                # Set column headers including Actions column
                self.table.setColumnCount(7)  # Increased to 7 columns for Fee Payer
                self.table.setHorizontalHeaderLabels(["Student ID", "Student Name", "Class", "Term", "Fee Payer", "PayRate", "Actions"])
                
                # Configure column resize modes
                header = self.table.horizontalHeader()
                for col in range(6):  # First 5 columns are data
                    header.setSectionResizeMode(col, QHeaderView.Stretch)
                header.setSectionResizeMode(6, QHeaderView.Fixed)  # Actions column fixed width
                header.resizeSection(6, 220)  # Wider width for action buttons (3 buttons now)
                
                self.table.setRowCount(0)
                for row_data in rows:
                    row = self.table.rowCount()
                    self.table.insertRow(row)
                    self.table.setRowHeight(row, 75)
                    
                    # Add data columns (skip StudentTermID at index 0)
                    for col, data in enumerate(row_data[1:7]):  # Columns 1-5: StudentID, Name, Class, Term, FeePayer
                        item = QTableWidgetItem(str(data) if data else "0")
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Make non-editable
                        self.table.setItem(row, col, item)
                    
                    student_term_id = row_data[0]  # StudentTermID for actions
                    student_id = row_data[1]       # StudentID
                    class_name = row_data[3]       # ClassName
                    term_name = row_data[4]        # TermName
                    
                    # Create action ComboBox with 3 actions
                    action_combo = QComboBox()
                    action_combo.addItems(["Select Action","Pay Fee", "Edit", "Delete"])
                    # Connect directly with lambda
                    action_combo.currentIndexChanged.connect(
                        lambda index, stid=student_term_id, sid=student_id, cname=class_name, tname=term_name, cb=action_combo: self.on_enrollment_action_selected(index, stid, sid, cname, tname, cb)
                    )

                    # Set the ComboBox in the cell
                    self.table.setCellWidget(row, 6, action_combo)
                    
                    # Set proper row height
                    self.table.setRowHeight(row, 36)
                
                count = self.table.rowCount()
                self.results_label.setText(f"{count} results")
                
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")    

    def  on_enrollment_action_selected(self, index, student_term_id, student_id, class_name, term_name, combo_box):
        """Handle enrollment action selection directly"""
        if index == 0:  # "Select Action"
            return
        
        action = combo_box.itemText(index)
        combo_box.setCurrentIndex(0)  # Reset to "Select Action"
        
        if action == "Pay Fee":
            self.pay_fee(student_term_id)
        elif action == "Edit":
            self.edit_enrollment(student_term_id)
        elif action == "Delete":
            self.delete_enrollment(student_term_id) 

    def edit_enrollment(self, student_term_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute("""SELECT st.StudentID, st.ClassID, st.TermID, s.Student_Name, cl.ClassName, t.TermName, st.PayRate
                             FROM StudentTerm st
                             LEFT JOIN Students s ON st.StudentID = s.StudentID
                             LEFT JOIN Class cl ON st.ClassID = cl.ClassID
                             LEFT JOIN Term t ON st.TermID = t.TermID
                             WHERE st.StudentTermID = ?""", (student_term_id,))
                data = c.fetchone()
            if data:
                dialog = QDialog(self)
                dialog.setWindowTitle("Edit Enrollment")
                layout = QVBoxLayout()
                form_layout = QFormLayout()
                student_id_label = QLabel(str(data[0]))
                class_combo = QComboBox()
                term_combo = QComboBox()
                payrate_combo = QComboBox()
                payrate_combo.addItem("Select PayRate",None)
                payrate_combo.addItem("Full Scholarship",0)
                payrate_combo.addItem("20% Scholarship",80)
                payrate_combo.addItem("No Scholarship",100)
                self.populate_combos(class_combo, term_combo)
                class_combo.setCurrentIndex(class_combo.findData(data[1]))
                term_combo.setCurrentIndex(term_combo.findData(data[2]))
                payrate_combo.setCurrentIndex(payrate_combo.findData(data[6]))
                form_layout.addRow("Student ID:", student_id_label)
                form_layout.addRow("Class:", class_combo)
                form_layout.addRow("Term:", term_combo)
                form_layout.addRow("PayRate:",payrate_combo )
                layout.addLayout(form_layout)
                buttons = QHBoxLayout()
                save_btn = QPushButton("Save")
                save_btn.clicked.connect(lambda: self.save_edited_enrollment(dialog, student_term_id, class_combo, term_combo,payrate_combo))
                close_btn = QPushButton("Close")
                close_btn.clicked.connect(dialog.close)
                buttons.addWidget(save_btn)
                buttons.addWidget(close_btn)
                layout.addLayout(buttons)
                dialog.setLayout(layout)
                dialog.exec_()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load enrollment: {str(e)}")

    def save_edited_enrollment(self, dialog, student_term_id, class_combo, term_combo,payrate_combo):
        class_id = class_combo.currentData()
        term_id = term_combo.currentData()
        payrate = payrate_combo.currentData()
        if not all([class_id, term_id, payrate]):
            QMessageBox.warning(dialog, "Error", "Required fields are missing")
            return
        try:
            with get_db_connection() as (conn, c):
                # Get student_id for duplicate check
                c.execute("SELECT StudentID FROM StudentTerm WHERE StudentTermID = ?", (student_term_id,))
                student_id = c.fetchone()[0]
                # Check for duplicate
                c.execute("SELECT COUNT(*) FROM StudentTerm WHERE StudentID = ? AND ClassID = ? AND TermID = ? AND StudentTermID != ?", (student_id, class_id, term_id, student_term_id))
                if c.fetchone()[0] > 0:
                    QMessageBox.warning(dialog, "Error", "Duplicate enrollment for this class and term")
                    return
                c.execute("""UPDATE StudentTerm 
                             SET ClassID = ?, TermID = ?, PayRate = ?
                             WHERE StudentTermID = ?""", (class_id, term_id, payrate,student_term_id))
                conn.commit()
                QMessageBox.information(dialog, "Success", "Enrollment updated successfully")
                dialog.close()
                self.load_data()
        except Exception as e:
            QMessageBox.warning(dialog, "Error", f"Failed to update enrollment: {str(e)}")

    def delete_enrollment(self, student_term_id):
        reply = QMessageBox.question(self, 'Confirm Deletion', 
                                     'Are you sure you want to delete this enrollment?',
                                     QMessageBox.Yes | QMessageBox.No, 
                                     QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    c.execute("DELETE FROM StudentTerm WHERE StudentTermID = ?", (student_term_id,))
                    conn.commit()
                    QMessageBox.information(self, "Success", "Enrollment deleted successfully")
                    self.load_data()
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to delete enrollment: {str(e)}")

    def populate_combos(self, class_combo, term_combo):
        
        with get_db_connection() as (conn, c):
            c.execute("SELECT ClassID, ClassName FROM Class")
            classes = c.fetchall()
            for cl in classes:
                class_combo.addItem(cl[1], cl[0])
            c.execute("SELECT TermID, TermName FROM Term")
            terms = c.fetchall()
            for t in terms:
                term_combo.addItem(t[1], t[0])
    

    def pay_fee(self, student_term_id):
        """Open fee payment dialog for a specific student term enrollment"""
        try:
            with get_db_connection() as (conn, c):
                # Get student term details including class and term
                c.execute("""
                    SELECT st.StudentTermID, st.ClassID, st.TermID, 
                        s.StudentID, s.Student_Name, c.ClassName, t.TermName
                    FROM StudentTerm st
                    JOIN Students s ON st.StudentID = s.StudentID
                    JOIN Class c ON st.ClassID = c.ClassID
                    JOIN Term t ON st.TermID = t.TermID
                    WHERE st.StudentTermID = ?
                """, (student_term_id,))
                enrollment_data = c.fetchone()
                
                if enrollment_data:
                    # Get the fee amount for this class and term
                    c.execute("""
                        SELECT FeesID, Amount 
                        FROM Fees 
                        WHERE Class = ? AND Term = ?
                    """, (enrollment_data[1], enrollment_data[2]))  # ClassID and TermID
                    
                    fee_data = c.fetchone()
                    if fee_data:
                        fees_id, amount_due = fee_data
                        self.show_payment_dialog(student_term_id, fees_id, amount_due, enrollment_data)
                    else:
                        QMessageBox.warning(self, "Error", "No fee structure found for this class and term combination.")
                else:
                    QMessageBox.warning(self, "Error", "Enrollment record not found.")
                    
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load fee information: {str(e)}")

    def show_payment_dialog(self, student_term_id, fees_id, amount_due, enrollment_data):
        """Show payment dialog with PayRate discount and validation"""
        try:
            with get_db_connection() as (conn, c):
                # Get the PayRate for this student term
                c.execute("SELECT PayRate FROM StudentTerm WHERE StudentTermID = ?", (student_term_id,))
                pay_rate_result = c.fetchone()
                pay_rate = pay_rate_result[0] if pay_rate_result else 100  # Default to 100% if not found
                
                # Calculate amount due after discount
                amount_after_discount = amount_due * (pay_rate / 100)
                
                # Get already paid amount for this student term
                c.execute("""
                    SELECT COALESCE(SUM(AmountPaid), 0) 
                    FROM FeesPayment 
                    WHERE StudentTermID = ?
                """, (student_term_id,))
                already_paid = c.fetchone()[0]
                
                # Calculate remaining amount to pay
                remaining_amount = max(0, amount_after_discount - already_paid)
                
                dialog = QDialog(self)
                dialog.setWindowTitle("Fee Payment")
                dialog.setFixedSize(550, 500)
                
                layout = QVBoxLayout(dialog)
                
                # Student information (read-only)
                info_group = QGroupBox("Student Information")
                info_layout = QFormLayout(info_group)
                
                student_info = QLabel(f"{enrollment_data[4]} ({enrollment_data[3]})")
                class_info = QLabel(enrollment_data[5])
                term_info = QLabel(enrollment_data[6])
                pay_rate_label = QLabel(f"{pay_rate}%")
                
                info_layout.addRow("Student:", student_info)
                info_layout.addRow("Class:", class_info)
                info_layout.addRow("Term:", term_info)
                info_layout.addRow("Pay Rate:", pay_rate_label)
                
                layout.addWidget(info_group)
                
                # Fee information
                fee_group = QGroupBox("Fee Details")
                fee_layout = QFormLayout(fee_group)
                
                original_amount_label = QLabel(f"{amount_due:,.2f}")
                discounted_amount_label = QLabel(f"{amount_after_discount:,.2f}")
                already_paid_label = QLabel(f"{already_paid:,.2f}")
                remaining_label = QLabel(f"{remaining_amount:,.2f}")
                
                # Style remaining amount based on value
                if remaining_amount > 0:
                    remaining_label.setStyleSheet("color: #E53E3E; font-weight: bold;")
                else:
                    remaining_label.setStyleSheet("color: #38A169; font-weight: bold;")
                
                fee_layout.addRow("Original Amount:", original_amount_label)
                fee_layout.addRow("After Discount:", discounted_amount_label)
                fee_layout.addRow("Already Paid:", already_paid_label)
                fee_layout.addRow("Remaining to Pay:", remaining_label)
                
                layout.addWidget(fee_group)
                
                # Payment information
                payment_group = QGroupBox("Payment Details")
                payment_layout = QFormLayout(payment_group)
                
                amount_paid_input = QLineEdit()
                amount_paid_input.setValidator(QDoubleValidator(0, remaining_amount, 2))
                amount_paid_input.setPlaceholderText(f"Max: {remaining_amount:,.2f}")
                
                # Real-time validation
                amount_paid_input.textChanged.connect(
                    lambda: self.validate_payment_amount(amount_paid_input, remaining_amount)
                )
                
                transaction_date = QDateEdit()
                transaction_date.setDate(QDate.currentDate())
                transaction_date.setCalendarPopup(True)
                transaction_date.setDisplayFormat("yyyy-MM-dd")
                
                reference_input = QLineEdit()
                reference_input.setPlaceholderText("Transaction reference number")
                
                payment_layout.addRow("Amount to Pay:", amount_paid_input)
                payment_layout.addRow("Transaction Date:", transaction_date)
                payment_layout.addRow("Reference Number:", reference_input)
                
                layout.addWidget(payment_group)
                
                # Buttons
                button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
                button_box.accepted.connect(lambda: self.process_payment(
                    dialog, student_term_id, fees_id, amount_paid_input.text(), 
                    transaction_date.date().toString("yyyy-MM-dd"), reference_input.text(),
                    remaining_amount
                ))
                button_box.rejected.connect(dialog.reject)
                
                layout.addWidget(button_box)
                
                # Disable OK button if no remaining amount
                if remaining_amount <= 0:
                    button_box.button(QDialogButtonBox.Ok).setEnabled(False)
                    button_box.button(QDialogButtonBox.Ok).setToolTip("No remaining amount to pay")
                
                dialog.exec_()
                
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load payment information: {str(e)}")

    def validate_payment_amount(self, amount_input, max_amount):
        """Validate that payment amount doesn't exceed remaining amount"""
        try:
            amount_text = amount_input.text().strip()
            if not amount_text:
                amount_input.setStyleSheet("")
                return
                
            amount = float(amount_text)
            
            if amount > max_amount:
                amount_input.setStyleSheet("color: red; border: 1px solid red;")
                amount_input.setToolTip(f"Amount cannot exceed {max_amount:,.2f}")
            elif amount <= 0:
                amount_input.setStyleSheet("color: red; border: 1px solid red;")
                amount_input.setToolTip("Amount must be greater than zero")
            else:
                amount_input.setStyleSheet("color: green;")
                amount_input.setToolTip("")
                
        except ValueError:
            amount_input.setStyleSheet("color: red; border: 1px solid red;")
            amount_input.setToolTip("Please enter a valid number")

    def process_payment(self, dialog, student_term_id, fees_id, amount_paid_str, 
                    transaction_date, reference_number, max_amount):
        """Process the fee payment with validation"""
        try:
            if not amount_paid_str.strip():
                QMessageBox.warning(dialog, "Error", "Please enter the amount to pay.")
                return
                
            amount_paid = float(amount_paid_str)
            
            if amount_paid <= 0:
                QMessageBox.warning(dialog, "Error", "Amount paid must be greater than zero.")
                return
                
            if amount_paid > max_amount:
                QMessageBox.warning(dialog, "Error", 
                                f"Amount paid ({amount_paid:,.2f}) cannot exceed remaining amount ({max_amount:,.2f}).")
                return
            
            # Get current datetime for DoneOn field
            done_on = QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss")
            
            with get_db_connection() as (conn, c):
                # Insert payment record
                c.execute("""
                    INSERT INTO FeesPayment 
                    (StudentTermID, FeesID, AmountPaid, TransactionDate, ReferenceNumber, DoneOn)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (student_term_id, fees_id, amount_paid, transaction_date, reference_number, done_on))
                
                conn.commit()
                
                # Check if fully paid
                c.execute("""
                    SELECT COALESCE(SUM(AmountPaid), 0) 
                    FROM FeesPayment 
                    WHERE StudentTermID = ?
                """, (student_term_id,))
                total_paid = c.fetchone()[0]
                
                c.execute("SELECT Amount FROM Fees WHERE FeesID = ?", (fees_id,))
                original_amount = c.fetchone()[0]
                
                c.execute("SELECT PayRate FROM StudentTerm WHERE StudentTermID = ?", (student_term_id,))
                pay_rate = c.fetchone()[0]
                
                amount_after_discount = original_amount * (pay_rate / 100)
                
                if total_paid >= amount_after_discount:
                    message = "Payment recorded successfully! ✅ Student has fully paid their fees."
                else:
                    remaining = amount_after_discount - total_paid
                    message = f"Payment recorded successfully! Remaining balance: {remaining:,.2f}"
                
                QMessageBox.information(dialog, "Success", message)
                dialog.accept()
                
        except ValueError:
            QMessageBox.warning(dialog, "Error", "Please enter a valid amount.")
        except sqlite3.Error as e:
            QMessageBox.critical(dialog, "Database Error", f"Failed to record payment: {str(e)}")

def show_import_wizard(self):
    wizard = ImportWizard()
    wizard.exec_()

class ImportWizard(QWizard):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Student Import Wizard")
        
        # Light theme stylesheet
        light_style = """
        QWizard {
            background-color: #f8f9fa;
            color: #212529;
        }
        QLabel {
            color: #495057;
        }
        QPushButton {
            background-color: #007bff;
            color: white;
            border: none;
            padding: 6px 12px;
            border-radius: 4px;
        }
        QPushButton:hover {
            background-color: #0056b3;
        }
        QComboBox, QLineEdit, QSpinBox {
            background-color: white;
            border: 1px solid #ced4da;
            padding: 4px;
            border-radius: 4px;
        }
        QRadioButton {
            color: #495057;
        }
        QTableView {
            background-color: white;
            border: 1px solid #dee2e6;
            gridline-color: #dee2e6;
        }
        QCheckBox {
            color: #495057;
        }
        QTableWidget {
            background-color: white;
            border: 1px solid #dee2e6;
        }
        """
        self.setStyleSheet(light_style)
        
        # Initialize variables
        self.mode = None
        self.excel_file = None
        self.term_id = None
        self.class_id = None
        self.academic_year_id = None
        self.data = None
        self.generate_ids = True  # Default to True
        self.starting_sequence = 1
        
        # Add pages
        self.intro_page_id = self.addPage(self.create_intro_page())
        self.file_page_id = self.addPage(self.create_file_page())
        self.config_page_id = self.addPage(self.create_config_page())
        self.id_generation_page_id = self.addPage(self.create_id_generation_page())
        self.preview_page_id = self.addPage(self.create_preview_page())
        self.finish_page_id = self.addPage(self.create_finish_page())
        
        # Connect page change signal
        self.currentIdChanged.connect(self.on_page_changed)

    def create_intro_page(self):
        page = QWizardPage()
        page.setTitle("Choose Import Mode")
        page.setSubTitle("Select how you want to import students")
        
        layout = QVBoxLayout()
        
        # Description label
        desc_label = QLabel("Choose the appropriate import method for your needs:")
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)
        
        # Mode selection
        self.mode_group = QButtonGroup()
        
        # Mode 1: New students with enrollment
        mode1_frame = QFrame()
        mode1_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        mode1_frame.setStyleSheet("""
            QFrame {
                background-color: #e8f4f8;
                border: 1px solid #b8daff;
                border-radius: 5px;
                padding: 10px;
                margin: 5px;
            }
        """)
        mode1_layout = QVBoxLayout(mode1_frame)
        self.radio_new_students = QRadioButton("Insert new students with initial term enrollment")
        self.radio_new_students.setChecked(True)
        mode1_desc = QLabel("Use this for importing completely new students.")
        mode1_desc.setWordWrap(True)
        mode1_desc.setStyleSheet("color: #666; font-size: 12px; margin-left: 20px;")
        mode1_layout.addWidget(self.radio_new_students)
        mode1_layout.addWidget(mode1_desc)
        
        # Mode 2: Enrollment only
        mode2_frame = QFrame()
        mode2_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        mode2_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 10px;
                margin: 5px;
            }
        """)
        mode2_layout = QVBoxLayout(mode2_frame)
        self.radio_enrollment_only = QRadioButton("Enrollment only (existing students)")
        mode2_desc = QLabel("Use this for enrolling existing students in a new term/class.")
        mode2_desc.setWordWrap(True)
        mode2_desc.setStyleSheet("color: #666; font-size: 12px; margin-left: 20px;")
        mode2_layout.addWidget(self.radio_enrollment_only)
        mode2_layout.addWidget(mode2_desc)
        
        self.mode_group.addButton(self.radio_new_students, 1)
        self.mode_group.addButton(self.radio_enrollment_only, 2)
        
        layout.addWidget(mode1_frame)
        layout.addWidget(mode2_frame)
        layout.addStretch()
        
        page.setLayout(layout)
        return page

    def create_file_page(self):
        page = QWizardPage()
        page.setTitle("Select Excel File")
        page.setSubTitle("Choose the Excel file containing student data")
        
        layout = QVBoxLayout()
        
        # File selection area
        file_frame = QFrame()
        file_frame.setFrameStyle(QFrame.Box | QFrame.Sunken)
        file_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 5px;
                padding: 15px;
            }
        """)
        file_layout = QVBoxLayout(file_frame)
        
        # File info
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet("font-style: italic; color: #6c757d;")
        
        # File format info
        format_label = QLabel("Expected Excel format:")
        format_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        
        format_info = QLabel("For new students: Student_Name, FeePayer, PayRate, Parent_Name, Parent_Contact\n"
                            "Note: StudentID column is optional (will be auto-generated if missing)")
        format_info.setStyleSheet("font-size: 12px; color: #666; background-color: #f8f9fa; padding: 5px; border-radius: 3px;")
        format_info.setWordWrap(True)
        
        # Browse button
        btn = QPushButton("📁 Browse for Excel File")
        btn.clicked.connect(self.browse_file)
        
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(format_label)
        file_layout.addWidget(format_info)
        file_layout.addWidget(btn)
        
        layout.addWidget(file_frame)
        layout.addStretch()
        
        page.setLayout(layout)
        return page

    def browse_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, 
            "Select Excel File", 
            "", 
            "Excel Files (*.xlsx *.xls);;All Files (*.*)"
        )
        if file:
            self.excel_file = file
            self.file_label.setText(f"Selected: {os.path.basename(file)}")
            self.file_label.setStyleSheet("color: #28a745; font-weight: bold;")

    def create_config_page(self):
        page = QWizardPage()
        page.setTitle("Configure Enrollment")
        page.setSubTitle("Select the academic year, term, and class for enrollment")
        
        layout = QVBoxLayout()
        
        # Configuration frame
        config_frame = QFrame()
        config_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        config_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 15px;
            }
        """)
        config_layout = QGridLayout(config_frame)
        config_layout.setSpacing(10)
        
        # Academic Year selection
        year_label = QLabel("Academic Year:")
        year_label.setStyleSheet("font-weight: bold;")
        self.academic_combo = QComboBox()
        self.academic_combo.setMinimumWidth(200)
        
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY Year DESC")
                years = c.fetchall()
                self.academic_combo.addItem("Select Academic Year", None)
                for yid, yname in years:
                    self.academic_combo.addItem(yname, yid)
        except sqlite3.Error as e:
            QMessageBox.warning(self, "Database Error", f"Failed to load academic years: {str(e)}")
        
        self.academic_combo.currentIndexChanged.connect(self.load_terms_and_classes)
        
        config_layout.addWidget(year_label, 0, 0)
        config_layout.addWidget(self.academic_combo, 0, 1)
        
        # Term selection
        term_label = QLabel("Term:")
        term_label.setStyleSheet("font-weight: bold;")
        self.term_combo = QComboBox()
        self.term_combo.setMinimumWidth(200)
        self.term_combo.addItem("Select Term", None)
        config_layout.addWidget(term_label, 1, 0)
        config_layout.addWidget(self.term_combo, 1, 1)
        
        # Class selection
        class_label = QLabel("Class:")
        class_label.setStyleSheet("font-weight: bold;")
        self.class_combo = QComboBox()
        self.class_combo.setMinimumWidth(200)
        self.class_combo.addItem("Select Class", None)
        config_layout.addWidget(class_label, 2, 0)
        config_layout.addWidget(self.class_combo, 2, 1)
        
        layout.addWidget(config_frame)
        layout.addStretch()
        
        page.setLayout(layout)
        return page

    def load_terms_and_classes(self):
        year_id = self.academic_combo.currentData()
        self.academic_year_id = year_id
        
        self.term_combo.clear()
        self.class_combo.clear()
        
        if year_id is None:
            self.term_combo.addItem("Select Term", None)
            self.class_combo.addItem("Select Class", None)
            return
        
        try:
            with get_db_connection() as (conn, c):
                # Load terms for selected year
                c.execute("SELECT TermID, TermName FROM Term WHERE AcademicYear = ? ORDER BY TermName", (year_id,))
                terms = c.fetchall()
                self.term_combo.addItem("Select Term", None)
                for tid, tname in terms:
                    self.term_combo.addItem(tname, tid)
                
                # Load classes for selected year
                c.execute("SELECT ClassID, ClassName FROM Class WHERE AcademicYear = ? ORDER BY ClassName", (year_id,))
                classes = c.fetchall()
                self.class_combo.addItem("Select Class", None)
                for cid, cname in classes:
                    self.class_combo.addItem(cname, cid)
                    
                # Update ID preview if we're on that page
                if self.currentId() == self.id_generation_page_id:
                    self.update_id_preview()
                    
        except sqlite3.Error as e:
            QMessageBox.warning(self, "Database Error", f"Failed to load terms/classes: {str(e)}")

    def create_id_generation_page(self):
        page = QWizardPage()
        page.setTitle("Student ID Generation")
        page.setSubTitle("Configure how Student IDs should be generated")
        
        layout = QVBoxLayout()
        
        # ID generation options
        options_frame = QFrame()
        options_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        options_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 15px;
            }
        """)
        options_layout = QVBoxLayout(options_frame)
        
        # Checkbox for ID generation
        self.generate_ids_checkbox = QCheckBox("Generate Student IDs automatically")
        self.generate_ids_checkbox.setChecked(True)
        self.generate_ids_checkbox.stateChanged.connect(self.on_generate_ids_changed)
        
        # Format explanation
        format_label = QLabel("Format: AcademicYear-0001 (e.g., 2024-0001)")
        format_label.setStyleSheet("color: #666; font-size: 12px; margin-left: 20px;")
        
        # Starting sequence
        seq_layout = QHBoxLayout()
        seq_label = QLabel("Starting sequence number:")
        self.start_sequence_spin = QSpinBox()
        self.start_sequence_spin.setRange(1, 9999)
        self.start_sequence_spin.setValue(1)
        self.start_sequence_spin.valueChanged.connect(self.update_id_preview)
        seq_layout.addWidget(seq_label)
        seq_layout.addWidget(self.start_sequence_spin)
        seq_layout.addStretch()
        
        # Last ID preview
        self.last_id_label = QLabel("Select an academic year to see ID preview")
        self.last_id_label.setStyleSheet("""
            QLabel {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 3px;
                padding: 10px;
                margin-top: 10px;
                font-family: monospace;
            }
        """)
        
        options_layout.addWidget(self.generate_ids_checkbox)
        options_layout.addWidget(format_label)
        options_layout.addLayout(seq_layout)
        options_layout.addWidget(self.last_id_label)
        
        layout.addWidget(options_frame)
        
        # Preview area for ID generation
        preview_label = QLabel("Preview of generated IDs:")
        preview_label.setStyleSheet("font-weight: bold; margin-top: 15px;")
        layout.addWidget(preview_label)
        
        self.id_preview_table = QTableWidget(5, 2)
        self.id_preview_table.setHorizontalHeaderLabels(["#", "Generated StudentID"])
        self.id_preview_table.horizontalHeader().setStretchLastSection(True)
        self.id_preview_table.setMaximumHeight(150)
        layout.addWidget(self.id_preview_table)
        
        layout.addStretch()
        
        page.setLayout(layout)
        return page

    def on_generate_ids_changed(self, state):
        self.generate_ids = (state == Qt.Checked)
        self.start_sequence_spin.setEnabled(self.generate_ids)
        self.update_id_preview()

    def generate_student_id(self, year_name, sequence_number):
        """Generate StudentID in format: Year-0001"""
        # Clean year name (take first part before any dash)
        if isinstance(year_name, str):
            year_clean = str(year_name).strip().split('-')[0]
        else:
            year_clean = str(year_name)
        
        # Format sequence number with leading zeros
        seq_formatted = f"{sequence_number:04d}"
        
        return f"{year_clean}-{seq_formatted}"

    def get_existing_sequence(self, year_id):
        """Get the highest existing sequence number for the given academic year"""
        try:
            with get_db_connection() as (conn, c):
                # Get the year name
                c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (year_id,))
                year_result = c.fetchone()
                if not year_result:
                    return 0
                
                year_name = year_result[0]
                
                # Get all existing StudentIDs for this year
                c.execute("""
                    SELECT StudentID FROM Students 
                    WHERE StudentID LIKE ? AND StudentID GLOB ? 
                    ORDER BY StudentID DESC
                """, (f"{year_name}-%", f"{year_name}-[0-9][0-9][0-9][0-9]"))
                
                existing_ids = [row[0] for row in c.fetchall()]
                
                if not existing_ids:
                    return 0
                
                # Find the highest sequence number
                max_seq = 0
                for student_id in existing_ids:
                    try:
                        # Extract sequence number from format "Year-0001"
                        parts = student_id.split('-')
                        if len(parts) == 2 and parts[1].isdigit():
                            seq = int(parts[1])
                            max_seq = max(max_seq, seq)
                    except (ValueError, IndexError):
                        continue
                
                return max_seq
                
        except sqlite3.Error as e:
            print(f"Error getting existing sequence: {str(e)}")
            return 0

    def update_id_preview(self):
        """Update the ID preview based on current settings"""
        self.id_preview_table.clearContents()
        
        if not self.generate_ids:
            self.last_id_label.setText("ID generation is disabled")
            self.id_preview_table.setRowCount(1)
            self.id_preview_table.setItem(0, 0, QTableWidgetItem("1"))
            self.id_preview_table.setItem(0, 1, QTableWidgetItem("Will use StudentID from Excel"))
            return
        
        if not self.academic_year_id:
            self.last_id_label.setText("Please select an academic year first")
            self.id_preview_table.setRowCount(1)
            self.id_preview_table.setItem(0, 0, QTableWidgetItem("1"))
            self.id_preview_table.setItem(0, 1, QTableWidgetItem("Select academic year to see preview"))
            return
        
        try:
            with get_db_connection() as (conn, c):
                # Get year name
                c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (self.academic_year_id,))
                year_result = c.fetchone()
                if not year_result:
                    return
                
                year_name = year_result[0]
                
                # Get existing highest sequence
                existing_max = self.get_existing_sequence(self.academic_year_id)
                
                # Get starting sequence (use higher of existing+1 or user specified)
                user_start = self.start_sequence_spin.value()
                actual_start = max(existing_max + 1, user_start)
                
                # Update preview
                self.id_preview_table.setRowCount(5)
                
                for i in range(5):
                    sequence_num = actual_start + i
                    student_id = self.generate_student_id(year_name, sequence_num)
                    
                    # Add to table
                    self.id_preview_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
                    self.id_preview_table.setItem(i, 1, QTableWidgetItem(student_id))
                
                # Update label
                if existing_max > 0:
                    self.last_id_label.setText(f"Last existing ID: {year_name}-{existing_max:04d}\n"
                                              f"Starting from: {year_name}-{actual_start:04d}")
                else:
                    self.last_id_label.setText(f"No existing IDs found for {year_name}\n"
                                              f"Starting from: {year_name}-{actual_start:04d}")
        
        except sqlite3.Error as e:
            self.last_id_label.setText(f"Error: {str(e)}")

    def create_preview_page(self):
        page = QWizardPage()
        page.setTitle("Preview Data")
        page.setSubTitle("Review the data before importing")
        
        layout = QVBoxLayout()
        
        # Data preview table
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.horizontalHeader().setStretchLastSection(True)
        
        # Summary label
        self.summary_label = QLabel("No data loaded")
        self.summary_label.setStyleSheet("font-weight: bold; color: #6c757d;")
        
        layout.addWidget(self.summary_label)
        layout.addWidget(self.table_view)
        
        page.setLayout(layout)
        return page

    def create_finish_page(self):
        page = QWizardPage()
        page.setTitle("Ready to Import")
        page.setSubTitle("Click Finish to import the data")
        
        layout = QVBoxLayout()
        
        # Summary frame
        summary_frame = QFrame()
        summary_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        summary_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 2px solid #28a745;
                border-radius: 5px;
                padding: 15px;
            }
        """)
        summary_layout = QVBoxLayout(summary_frame)
        
        self.import_summary_label = QLabel("Import summary will appear here")
        self.import_summary_label.setWordWrap(True)
        self.import_summary_label.setStyleSheet("font-size: 14px;")
        
        summary_layout.addWidget(self.import_summary_label)
        layout.addWidget(summary_frame)
        
        # Warning/Info label
        self.final_warning_label = QLabel("⚠️ Review the preview carefully. This action cannot be undone.")
        self.final_warning_label.setStyleSheet("color: #856404; background-color: #fff3cd; padding: 10px; border-radius: 3px;")
        self.final_warning_label.setWordWrap(True)
        layout.addWidget(self.final_warning_label)
        
        layout.addStretch()
        
        page.setLayout(layout)
        return page

    def on_page_changed(self, page_id):
        """Handle page changes to update content"""
        if page_id == self.preview_page_id:
            self.load_data()
        elif page_id == self.finish_page_id:
            self.update_finish_summary()
        elif page_id == self.id_generation_page_id:
            self.update_id_preview()

    def load_data(self):
        if not self.excel_file:
                QMessageBox.warning(self, "Error", "No Excel file selected.")
                return
        try:
            # Load the Excel file with proper handling
            self.data = pd.read_excel(self.excel_file)
            
            # Fix column names - replace spaces with underscores
            self.data.columns = [col.strip().replace(' ', '_') for col in self.data.columns]
            
            # Handle "Parent Contact2" -> "Parent_Contact2"
            if 'Parent_Contact2' in self.data.columns:
                # Already correct
                pass
            elif 'Parent_Contact_2' in self.data.columns:
                self.data.rename(columns={'Parent_Contact_2': 'Parent_Contact2'}, inplace=True)
            
            # Check mode
            self.mode = self.mode_group.checkedId()
            
            # For new students mode
            if self.mode == 1:
                # Define expected columns (StudentID is optional)
                expected_cols = ['Student_Name', 'FeePayer', 'PayRate', 'Parent_Name', 'Parent_Contact']
                
                # Check for missing required columns
                missing_cols = []
                for col in expected_cols:
                    if col not in self.data.columns:
                        missing_cols.append(col)
                
                if missing_cols:
                    raise ValueError(f"Missing required columns: {', '.join(missing_cols)}\n"
                                    f"Found columns: {', '.join(self.data.columns)}")
                
                # === FIX 1: Convert columns to string BEFORE using .str accessor ===
                # First, ensure all expected string columns exist and are object type
                for col in expected_cols:
                    if col in self.data.columns:
                        # Convert to string, handling NaN values
                        self.data[col] = self.data[col].astype(str)
                
                # Now we can safely use .str methods
                if 'FeePayer' in self.data.columns:
                    # Convert to string first (already done above)
                    # Then strip whitespace
                    self.data['FeePayer'] = self.data['FeePayer'].str.strip()
                    # Replace empty FeePayer with 'parents' as default
                    self.data.loc[self.data['FeePayer'] == 'nan', 'FeePayer'] = 'parents'
                    self.data.loc[self.data['FeePayer'] == '', 'FeePayer'] = 'parents'
                
                # Handle PayRate - allow 0 and convert properly
                if 'PayRate' in self.data.columns:
                    # First convert to numeric, coerce errors to NaN
                    self.data['PayRate'] = pd.to_numeric(self.data['PayRate'], errors='coerce')
                    # Fill NaN with 0
                    self.data['PayRate'] = self.data['PayRate'].fillna(0)
                    # Convert to float
                    self.data['PayRate'] = self.data['PayRate'].astype(float)
                
                # === FIX 2: Clean up string columns SAFELY ===
                string_columns = self.data.select_dtypes(include=['object']).columns
                for col in string_columns:
                    # Ensure it's string type
                    self.data[col] = self.data[col].astype(str)
                    # Now we can use .str methods
                    self.data[col] = self.data[col].str.strip()
                    # Replace 'nan' string with empty string
                    self.data.loc[self.data[col] == 'nan', col] = ''
                
                # Generate preview StudentIDs if needed
                if self.generate_ids and self.academic_year_id:
                    with get_db_connection() as (conn, c):
                        c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (self.academic_year_id,))
                        year_result = c.fetchone()
                        if year_result:
                            year_name = year_result[0]
                            
                            # Get starting sequence
                            existing_max = self.get_existing_sequence(self.academic_year_id)
                            user_start = self.start_sequence_spin.value()
                            actual_start = max(existing_max + 1, user_start)
                            
                            # Generate IDs for preview
                            generated_ids = []
                            for i in range(len(self.data)):
                                student_id = f"{year_name}-{(actual_start + i):04d}"
                                generated_ids.append(student_id)
                            
                            # Add to DataFrame for display
                            self.data['StudentID'] = generated_ids
                            # Move StudentID to first column
                            cols = ['StudentID'] + [col for col in self.data.columns if col != 'StudentID']
                            self.data = self.data[cols]
                
                # If not generating IDs and StudentID column exists, clean it
                elif 'StudentID' in self.data.columns:
                    # Ensure it's string type first
                    self.data['StudentID'] = self.data['StudentID'].astype(str)
                    # Then use .str methods
                    self.data['StudentID'] = self.data['StudentID'].str.strip()
                    # Replace 'nan' with empty string
                    self.data.loc[self.data['StudentID'] == 'nan', 'StudentID'] = ''
                    # Move StudentID to first column
                    cols = ['StudentID'] + [col for col in self.data.columns if col != 'StudentID']
                    self.data = self.data[cols]
            
            elif self.mode == 2:  # Enrollment only
                if 'StudentID' not in self.data.columns:
                    raise ValueError("Missing 'StudentID' column for enrollment only.")
                
                # === FIX 3: Clean StudentID column safely ===
                # Convert to string first
                self.data['StudentID'] = self.data['StudentID'].astype(str)
                # Then strip
                self.data['StudentID'] = self.data['StudentID'].str.strip()
                # Replace 'nan' with empty
                self.data.loc[self.data['StudentID'] == 'nan', 'StudentID'] = ''
            
            else:
                raise ValueError("No mode selected.")
            
            # Display in table
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(self.data.columns)
            
            # Color code rows based on validation
            for idx, row in self.data.iterrows():
                items = []
                for col in self.data.columns:
                    value = row.get(col, '')
                    
                    # Handle NaN values
                    if pd.isna(value):
                        value = ''
                    
                    item = QStandardItem(str(value))
                    
                    # Highlight issues
                    if self.mode == 1:
                        if col == 'Student_Name' and not str(value).strip():
                            item.setBackground(QColor(255, 243, 205))  # Warning yellow
                        elif col == 'Parent_Contact' and not str(value).strip():
                            item.setBackground(QColor(255, 243, 205))
                        elif col == 'FeePayer' and not str(value).strip():
                            item.setBackground(QColor(255, 243, 205))
                        elif col == 'PayRate':
                            try:
                                pay_rate = float(value) if value else 0
                                if pay_rate < 0:
                                    item.setBackground(QColor(255, 220, 220))  # Light red for negative
                            except ValueError:
                                item.setBackground(QColor(255, 220, 220))
                    
                    items.append(item)
                
                model.appendRow(items)
            
            self.table_view.setModel(model)
            
            # Update summary
            row_count = len(self.data)
            issues = []
            
            if self.mode == 1:
                # Count missing required fields - use direct string comparison
                missing_names = (self.data['Student_Name'].astype(str).str.strip() == '').sum()
                missing_contacts = (self.data['Parent_Contact'].astype(str).str.strip() == '').sum()
                
                if missing_names > 0:
                    issues.append(f"{missing_names} missing student names")
                if missing_contacts > 0:
                    issues.append(f"{missing_contacts} missing parent contacts")
            
            summary = f"Previewing {row_count} records"
            if issues:
                summary += f" - Issues: {', '.join(issues)}"
            
            self.summary_label.setText(summary)
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")
            import traceback
            traceback.print_exc()  # Print full error to console
            self.data = None
            self.summary_label.setText("Error loading data")

    def update_finish_summary(self):
        """Update the summary on the finish page"""
        if self.data is None or self.data.empty:
            self.import_summary_label.setText("No data to import. Please go back and load valid data.")
            return
        
        mode_text = "New Students with Enrollment" if self.mode == 1 else "Enrollment Only"
        year_text = self.academic_combo.currentText() if self.academic_combo.currentText() else "Not selected"
        term_text = self.term_combo.currentText() if self.term_combo.currentText() else "Not selected"
        class_text = self.class_combo.currentText() if self.class_combo.currentText() else "Not selected"
        
        summary = f"""
        <b>Import Summary:</b><br><br>
        <b>Mode:</b> {mode_text}<br>
        <b>Records to import:</b> {len(self.data)}<br>
        <b>Academic Year:</b> {year_text}<br>
        <b>Term:</b> {term_text}<br>
        <b>Class:</b> {class_text}<br><br>
        """
        
        if self.mode == 1 and self.generate_ids:
            summary += f"<b>Student IDs:</b> Will be generated automatically ({year_text.split('-')[0]}-####)<br>"
        elif self.mode == 1:
            summary += "<b>Student IDs:</b> Will be taken from Excel file (if present)<br>"
        
        self.import_summary_label.setText(summary)

    def accept(self):
        """Handle the import when Finish is clicked"""
        if self.data is None or self.data.empty:
            QMessageBox.warning(self, "Error", "No data to import.")
            return
        
        self.term_id = self.term_combo.currentData()
        self.class_id = self.class_combo.currentData()
        
        if not self.term_id or not self.class_id:
            QMessageBox.warning(self, "Error", "Term and Class must be selected.")
            return
        
        # Confirm import
        reply = QMessageBox.question(
            self,
            "Confirm Import",
            f"Are you sure you want to import {len(self.data)} record(s)?\n\n"
            f"This action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        success_count = 0
        error_count = 0
        error_messages = []
        
        try:
            with get_db_connection() as (conn, c):
                # Get academic year name for ID generation
                year_name = None
                if self.mode == 1:
                    c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (self.academic_year_id,))
                    year_result = c.fetchone()
                    if year_result:
                        year_name = year_result[0]
                
                # Get starting sequence for ID generation
                starting_sequence = self.start_sequence_spin.value()
                if self.mode == 1 and year_name:
                    # Get highest existing sequence
                    existing_max = self.get_existing_sequence(self.academic_year_id)
                    actual_start = max(existing_max + 1, starting_sequence)
                else:
                    actual_start = starting_sequence
                
                for index, row in self.data.iterrows():
                    excel_row = index + 2  # Excel row number (header is row 1)
                    try:
                        if self.mode == 1:  # New students with enrollment
                            # Generate or get StudentID
                            if self.generate_ids and year_name:
                                # Generate new ID
                                student_id = f"{year_name}-{actual_start:04d}"
                                actual_start += 1
                            else:
                                # Use ID from Excel or generate if missing
                                student_id = str(row.get('StudentID', '')).strip()
                                if not student_id and year_name:
                                    # Fallback generation
                                    student_id = f"{year_name}-{actual_start:04d}"
                                    actual_start += 1
                            
                            if not student_id:
                                raise ValueError("Could not generate StudentID")
                            
                            # Check if StudentID already exists
                            c.execute("SELECT StudentID FROM Students WHERE StudentID = ?", (student_id,))
                            if c.fetchone():
                                raise ValueError(f"StudentID '{student_id}' already exists")
                            
                            # Get and validate required fields
                            student_name = str(row.get('Student_Name', '')).strip()
                            fee_payer = str(row.get('FeePayer', '')).strip()
                            pay_rate = row.get('PayRate', 0)
                            parent_name = str(row.get('Parent_Name', '')).strip()
                            parent_contact = str(row.get('Parent_Contact', '')).strip()
                            parent_contact2 = str(row.get('Parent_Contact2', '')).strip()
                            
                            # Validate required fields
                            if not student_name:
                                raise ValueError("Student Name is required")
                            if not parent_name:
                                raise ValueError("Parent Name is required")
                            if not parent_contact:
                                raise ValueError("Parent Contact is required")
                            
                            # Handle FeePayer - if empty, use default
                            if not fee_payer:
                                fee_payer = "parents"
                            
                            # Handle PayRate - convert to float, allow 0
                            try:
                                pay_rate_float = float(pay_rate) if pay_rate else 0.0
                                if pay_rate_float < 0:
                                    raise ValueError("PayRate cannot be negative")
                            except ValueError:
                                raise ValueError("PayRate must be a valid number")
                            
                            # Insert into Students table
                            c.execute("""
                                INSERT INTO Students (StudentID, Student_Name, FeePayer, PayRate, 
                                                     Parent_Name, Parent_Contact, Parent_Contact2)
                                VALUES (?, ?, ?, ?, ?, ?, ?)
                            """, (
                                student_id,
                                student_name,
                                fee_payer,
                                pay_rate_float,
                                parent_name,
                                parent_contact,
                                parent_contact2 if parent_contact2 else None
                            ))
                            
                            # Insert into StudentTerm
                            c.execute("""
                                INSERT INTO StudentTerm (StudentID, ClassID, TermID, PayRate)
                                VALUES (?, ?, ?, ?)
                            """, (student_id, self.class_id, self.term_id, pay_rate_float))
                            
                            success_count += 1
                            
                        elif self.mode == 2:  # Enrollment only
                            student_id = str(row.get('StudentID', '')).strip()
                            if not student_id:
                                raise ValueError("StudentID is required")
                            
                            # Check if StudentID exists
                            c.execute("SELECT StudentID, Student_Name FROM Students WHERE StudentID = ?", (student_id,))
                            student_data = c.fetchone()
                            if not student_data:
                                raise ValueError(f"StudentID '{student_id}' does not exist")
                            
                            # Check if already enrolled
                            c.execute("""
                                SELECT StudentTermID FROM StudentTerm 
                                WHERE StudentID = ? AND ClassID = ? AND TermID = ?
                            """, (student_id, self.class_id, self.term_id))
                            if c.fetchone():
                                raise ValueError(f"Already enrolled in this term/class")
                            
                            # Insert enrollment
                            c.execute("""
                                INSERT INTO StudentTerm (StudentID, ClassID, TermID)
                                VALUES (?, ?, ?)
                            """, (student_id, self.class_id, self.term_id))
                            
                            success_count += 1
                        
                    except Exception as e:
                        error_count += 1
                        error_messages.append(f"Row {excel_row}: {str(e)}")
                
                # Commit if there were successful operations
                if success_count > 0:
                    conn.commit()
                else:
                    conn.rollback()
                    
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to import data: {str(e)}")
            return
        
        # Show results
        if error_count > 0:
            error_details = "\n".join(error_messages[:10])
            if error_count > 10:
                error_details += f"\n... and {error_count - 10} more errors"
            
            result_text = f"""
            Import completed with errors:
            
            ✅ Successfully imported: {success_count} record(s)
            ❌ Failed to import: {error_count} record(s)
            
            First few errors:
            {error_details}
            """
            
            QMessageBox.warning(self, "Import Completed with Errors", result_text)
        else:
            QMessageBox.information(
                self, 
                "Import Successful", 
                f"✅ Successfully imported {success_count} record(s)"
            )
        
        super().accept()
### Student Management End ###

class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Pinnacle-SF")
        self.setGeometry(100, 100, 1200, 800)
        
        # Main widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # Sidebar
        self.sidebar = QWidget()
        self.sidebar.setFixedWidth(250)
        self.sidebar_layout = QVBoxLayout(self.sidebar)
        self.sidebar_layout.setContentsMargins(10, 10, 10, 10)
        self.sidebar_layout.setSpacing(10)
        
        # Toggle button
        self.toggle_btn = QPushButton(
            QIcon(resource_path("Icons/menu-icon.png")), ""
        )
        self.toggle_btn.setCursor(Qt.SizeHorCursor)
        self.toggle_btn.setIconSize(QSize(24, 24))
        self.toggle_btn.clicked.connect(self.toggle_sidebar)
        self.sidebar_layout.addWidget(self.toggle_btn)

        # Sidebar menu items
        self.menu_items = [
            ("Dashboard", "Icons/dashboard-icon.png"),
            ("Students", "Icons/student-icon.png"),
            ("Fees", "Icons/fees-icon.png"),
            ("Uniforms", "Icons/uniform-icon.png"),
            ("Expenditure", "Icons/expenses-icon.png"),
            ("Classes", "Icons/classes-icon.png"),
            ("Academic Calendar", "Icons/Calendar.png"),
            ("Categorization", "Icons/category_64px.png"),
            ("Reports", "Icons/report-icon.png"),
        ]

        self.menu_buttons = []

        for name, icon_path in self.menu_items:
            btn = QPushButton(
                QIcon(resource_path(icon_path)), name
            )
            btn.setFixedHeight(40)
            btn.setIconSize(QSize(24, 24))
            btn.original_text = name
            btn.icon_path = icon_path
            btn.setCheckable(True)
            btn.setAutoExclusive(True)
            self.menu_buttons.append(btn)
            self.sidebar_layout.addWidget(btn)

        self.sidebar_layout.addStretch()

        # Bottom menu items
        all_menu_items = [
            ("About", "Icons/about-icon.png")
        ]

        for name2, icon_path2 in all_menu_items:
            buttons = QPushButton(
                QIcon(resource_path(icon_path2)), name2
            )
            buttons.setFixedHeight(40)
            buttons.setIconSize(QSize(24, 24))
            buttons.original_text = name2
            buttons.icon_path = icon_path2
            buttons.setCheckable(True)
            buttons.setAutoExclusive(True)
            self.menu_buttons.append(buttons)
            self.sidebar_layout.addWidget(buttons)


        # Main content
        self.content = QWidget()
        self.content_layout = QVBoxLayout(self.content)
        self.content_layout.setContentsMargins(20, 20, 20, 20)
        self.content_layout.setSpacing(10)
        
        # Header
        self.header = QWidget()
        self.header_layout = QHBoxLayout(self.header)
        self.header_layout.setContentsMargins(0, 0, 0, 10)
        
        self.page_title = QLabel("Dashboard")
        self.header_layout.addWidget(self.page_title)
        self.header_layout.addStretch()
        
        self.school_name = QLabel("Pinnacle-SF")
        self.header_layout.addWidget(self.school_name)
        
        self.content_layout.addWidget(self.header)
        
        # Separator line
        self.separator = QWidget()
        self.separator.setFixedHeight(2)
        self.content_layout.addWidget(self.separator)
        
        # Create a container widget for the stacked widget
        self.body_container = QWidget()
        self.body_layout = QVBoxLayout(self.body_container)
        self.body_layout.setContentsMargins(0, 0, 0, 0)
        self.body_layout.setSpacing(0)
        
        # Stacked widget for pages
        self.body = QStackedWidget()
        
        # Initialize pages
        self.AcademicYearTermPage = AcademicYearTerm()
        self.DashboardPage = Dashboard()
        self.StudentPage = StudentManagement()
        self.ClassPage = Class()
        self.FeesPage = Fees()
        self.UniformsPage = Uniforms()
        self.ExpensesPage = Expenditure()
        self.ReportsPage = Reports()
        self.CategoryPage = CategoryManagement()
        self.AboutPage = About()
        
        # Add pages to stacked widget
        self.body.addWidget(self.DashboardPage)
        self.body.addWidget(self.StudentPage)
        self.body.addWidget(self.FeesPage)
        self.body.addWidget(self.UniformsPage)
        self.body.addWidget(self.ExpensesPage)
        self.body.addWidget(self.ClassPage)
        self.body.addWidget(self.AcademicYearTermPage)
        self.body.addWidget(self.ReportsPage)
        self.body.addWidget(self.CategoryPage)
        self.body.addWidget(self.AboutPage)
        
        # Add stacked widget to container layout
        self.body_layout.addWidget(self.body)
        
        # Scroll area for body
        self.MainBodyHolder = QScrollArea()
        self.MainBodyHolder.setWidgetResizable(True)
        self.MainBodyHolder.setWidget(self.body_container)
        self.content_layout.addWidget(self.MainBodyHolder)
        
        # Add sidebar and content to main layout
        self.main_layout.addWidget(self.sidebar)
        self.main_layout.addWidget(self.content)
        
        # Sidebar state
        self.sidebar_expanded = True
        self.apply_styles()
        
        # Connect sidebar buttons to stacked pages
        page_map = {
            "Dashboard": (self.DashboardPage, 0),
            "Students": (self.StudentPage, 1),
            "Fees": (self.FeesPage, 2),
            "Uniforms": (self.UniformsPage, 3),
            "Expenditure": (self.ExpensesPage, 4),
            "Classes": (self.ClassPage, 5),
            "Academic Calendar": (self.AcademicYearTermPage, 6),
            "Reports": (self.ReportsPage, 7),
            "Categorization": (self.CategoryPage, 8),
            "About": (self.AboutPage, 9),
        }
        
        # Set Dashboard as active initially
        self.set_active_button("Dashboard")

        def make_nav_handler(page_widget, page_name, button):
            def handler():
                self.body.setCurrentWidget(page_widget)
                self.page_title.setText(page_name)
                # Set the clicked button as active
                self.set_active_button(page_name)
            return handler

        for btn in self.menu_buttons:
            page_data = page_map.get(btn.original_text)
            if page_data:
                page_widget, _ = page_data
                btn.clicked.connect(make_nav_handler(page_widget, btn.original_text, btn))

    def set_active_button(self, button_name):
        """Set the active sidebar button and style it"""
        for btn in self.menu_buttons:
            if btn.original_text == button_name:
                btn.setChecked(True)
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #0A7B83;
                        color: white;
                        border: none;
                        text-align: left;
                        border-radius: 6px;
                        font-weight: 500;
                    }
                    QPushButton:hover {
                        background-color: #086972;
                    }
                    QPushButton:checked {
                        background-color: #0A7B83;
                        color: white;
                        font-weight: 600;
                    }
                """)
            else:
                btn.setChecked(False)
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: transparent;
                        color: #343A40;
                        border: none;
                        text-align: left;
                        border-radius: 6px;
                        font-weight: 400;
                    }
                    QPushButton:hover {
                        background-color: #E9ECEF;
                        color: #1A365D;
                    }
                    QPushButton:checked {
                        background-color: #0A7B83;
                        color: white;
                    }
                """)

    def apply_styles(self):
        """Apply styling to the main window"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F8F9FA;
            }
            
            QWidget#sidebar {
                background-color: white;
                border-right: 1px solid #E9ECEF;
            }
            
            
            
            QLabel#page_title {
                font-size: 22px;
                font-weight: 600;
                color: #1A365D;
            }
            
            QLabel#school_name {
                font-size: 18px;
                color: #6C757D;
                font-weight: 500;
            }
            
            QWidget#separator {
                background-color: #E9ECEF;
            }
            
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            
            QScrollArea > QWidget > QWidget {
                background-color: transparent;
            }
            
            QScrollBar:vertical {
                background-color: #F8F9FA;
                width: 12px;
                border-radius: 6px;
            }
            
            QScrollBar::handle:vertical {
                background-color: #CED4DA;
                border-radius: 6px;
                min-height: 20px;
            }
            
            QScrollBar::handle:vertical:hover {
                background-color: #ADB5BD;
            }
        """)
        
        # Set object names for styling
        self.sidebar.setObjectName("sidebar")
        self.page_title.setObjectName("page_title")
        self.school_name.setObjectName("school_name")
        self.separator.setObjectName("separator")

    def toggle_sidebar(self):
        """Toggle sidebar visibility"""
        if self.sidebar_expanded:
            self.sidebar.setFixedWidth(60)
            # Hide text, show only icons
            for btn in self.menu_buttons:
                btn.setText("")
                btn.setIconSize(QSize(20, 20))
            self.sidebar_expanded = False
        else:
            self.sidebar.setFixedWidth(250)
            # Show text with icons
            for btn in self.menu_buttons:
                btn.setText(btn.original_text)
                btn.setIconSize(QSize(24, 24))
            self.sidebar_expanded = True
        
class AcademicYearTerm(QWidget):
    def __init__(self):
        super().__init__()
        self.current_year_id = None
        self.init_ui()
        self.load_academic_years()
        
    def init_ui(self):
        # Main Layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Split View Container
        split_container = QFrame()
        split_container.setObjectName("splitContainer")
        split_layout = QHBoxLayout(split_container)
        split_layout.setSpacing(20)
        split_layout.setContentsMargins(0, 0, 0, 0)

        # ========== LEFT PANEL: Academic Years (40%) ==========
        years_panel = QFrame()
        years_panel.setObjectName("yearsPanel")
        years_panel.setMinimumWidth(300)
        years_layout = QVBoxLayout(years_panel)
        years_layout.setContentsMargins(0, 0, 0, 0)
        years_layout.setSpacing(10)

        # Years Header
        years_header = QHBoxLayout()
        years_title = QLabel("Academic Years")
        years_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #2C3E50; margin-bottom: 5px;")
        years_header.addWidget(years_title)
        years_header.addStretch()
        
        # Years Search
        self.years_search = QLineEdit()
        self.years_search.setPlaceholderText(" Search years...")
        self.years_search.setFixedWidth(180)
        self.years_search.textChanged.connect(self.load_academic_years)
        years_header.addWidget(self.years_search)
        years_layout.addLayout(years_header)

        # Add Year Button
        add_year_btn = QPushButton("Add Year")
        add_year_btn.clicked.connect(self.addYear)
        years_header.addWidget(add_year_btn)

        # Years List Widget
        self.years_list = QListWidget()
        self.years_list.setObjectName("yearsList")
        self.years_list.setSelectionMode(QListWidget.SingleSelection)
        self.years_list.setStyleSheet("""
            QListWidget {
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 6px;
                padding: 5px;
            }
            QListWidget::item {
                padding: 12px 15px;
                border-bottom: 1px solid #f0f0f0;
                border-radius: 4px;
                margin: 2px 0;
            }
            QListWidget::item:selected {
                background-color: #3498db;
                color: white;
                border: 1px solid #2980b9;
            }
            QListWidget::item:hover {
                background-color: #f8f9fa;
            }
        """)
        self.years_list.itemClicked.connect(self.on_year_selected)
        years_layout.addWidget(self.years_list)

        # Years Actions
        years_actions = QHBoxLayout()
        years_actions.setSpacing(10)
        
        edit_year_btn = QPushButton("Edit")
        edit_year_btn.setObjectName("infoButton")
        edit_year_btn.clicked.connect(self.edit_selected_year)
        edit_year_btn.setFixedHeight(32)
        years_actions.addWidget(edit_year_btn)
        
        delete_year_btn = QPushButton("Delete")
        delete_year_btn.setObjectName("dangerButton")
        delete_year_btn.clicked.connect(self.delete_selected_year)
        delete_year_btn.setFixedHeight(32)
        years_actions.addWidget(delete_year_btn)
        
        years_actions.addStretch()
        years_layout.addLayout(years_actions)

        split_layout.addWidget(years_panel, 4)  # 40% width

        # ========== RIGHT PANEL: Terms (60%) ==========
        terms_panel = QFrame()
        terms_panel.setObjectName("termsPanel")
        terms_layout = QVBoxLayout(terms_panel)
        terms_layout.setContentsMargins(0, 0, 0, 0)
        terms_layout.setSpacing(10)

        # Terms Header
        terms_header = QHBoxLayout()
        self.terms_title = QLabel("Terms")
        self.terms_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #2C3E50; margin-bottom: 5px;")
        terms_header.addWidget(self.terms_title)
        terms_header.addStretch()
        
        # Add Term Button
        add_term_btn = QPushButton("Add Term")
        add_term_btn.clicked.connect(self.addTerm)
        add_term_btn.setFixedHeight(32)
        terms_header.addWidget(add_term_btn)
        terms_layout.addLayout(terms_header)

        # Terms Table
        self.terms_table = QTableWidget()
        self.terms_table.setColumnCount(5)
        self.terms_table.setHorizontalHeaderLabels(["Term", "Start Date", "End Date", "Duration", "Actions"])
        self.terms_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.terms_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.terms_table.setColumnWidth(3, 100)  # Duration column
        self.terms_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.terms_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.terms_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.terms_table.setAlternatingRowColors(True)
        self.terms_table.verticalHeader().setVisible(False)
        terms_layout.addWidget(self.terms_table)

        # Terms Statistics
        stats_frame = QFrame()
        stats_frame.setMaximumHeight(60)
        stats_layout = QHBoxLayout(stats_frame)
        stats_layout.setContentsMargins(0, 0, 0, 0)
        stats_layout.setSpacing(20)
        
        self.total_terms_label = QLabel("Total Terms: 0")
        self.total_terms_label.setStyleSheet("font-size: 13px; color: #7f8c8d; padding: 8px;")
        stats_layout.addWidget(self.total_terms_label)
        
        self.current_term_label = QLabel("Current Term: None")
        self.current_term_label.setStyleSheet("font-size: 13px; color: #7f8c8d; padding: 8px;")
        stats_layout.addWidget(self.current_term_label)
        
        stats_layout.addStretch()
        terms_layout.addWidget(stats_frame)

        split_layout.addWidget(terms_panel, 6)  # 60% width

        main_layout.addWidget(split_container)

        # Calendar Summary
        summary_frame = QFrame()
        summary_frame.setObjectName("summaryFrame")
        summary_frame.setMaximumHeight(80)
        summary_layout = QHBoxLayout(summary_frame)
        summary_layout.setContentsMargins(15, 10, 15, 10)
        
        # Today's date
        today = QDate.currentDate().toString("dddd, MMMM d, yyyy")
        date_label = QLabel(f"Today: {today}")
        date_label.setStyleSheet("font-size: 13px; color: #2C3E50;")
        summary_layout.addWidget(date_label)
        
        summary_layout.addStretch()
        
        # Active year info
        self.active_year_label = QLabel("Active Year: None")
        self.active_year_label.setStyleSheet("font-size: 13px; color: #27ae60; font-weight: bold;")
        summary_layout.addWidget(self.active_year_label)
        
        main_layout.addWidget(summary_frame)

    def load_academic_years(self):
        """Load academic years into the list widget"""
        try:
            self.years_list.clear()
            search_text = self.years_search.text().strip()
            
            with get_db_connection() as (conn, c):
                if search_text:
                    c.execute("SELECT AcademicYearID, Year FROM AcademicYear WHERE Year LIKE ? ORDER BY Year DESC", 
                            (f"%{search_text}%",))
                else:
                    c.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY Year DESC")
                
                years = c.fetchall()
                
                for year_id, year_name in years:
                    item = QListWidgetItem(year_name)
                    item.setData(Qt.UserRole, year_id)
                    
                    # Check if this year has terms
                    c.execute("SELECT COUNT(*) FROM Term WHERE AcademicYear = ?", (year_id,))
                    term_count = c.fetchone()[0]
                    
                    # Add term count as tooltip
                    item.setToolTip(f"{term_count} term{'s' if term_count != 1 else ''}")
                    
                    # Style based on whether it's the current year
                    if year_name.startswith(str(QDate.currentDate().year())):
                        item.setBackground(QColor("#e8f5e9"))
                        item.setForeground(QColor("#2e7d32"))
                    
                    self.years_list.addItem(item)
                
                # Select first item if available
                if years and not self.current_year_id:
                    self.years_list.setCurrentRow(0)
                    self.on_year_selected(self.years_list.item(0))
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load academic years: {str(e)}")

    def on_year_selected(self, item):
        """Handle year selection"""
        if not item:
            return
            
        self.current_year_id = item.data(Qt.UserRole)
        year_name = item.text()
        
        # Update terms title
        self.terms_title.setText(f" Terms - {year_name}")
        
        # Update active year label
        self.active_year_label.setText(f"Active Year: {year_name}")
        
        # Load terms for selected year
        self.load_terms_for_year()

    def load_terms_for_year(self):
        """Load terms for the selected academic year"""
        if not self.current_year_id:
            self.terms_table.setRowCount(0)
            self.total_terms_label.setText("Total Terms: 0")
            self.current_term_label.setText("Current Term: None")
            return
            
        try:
            with get_db_connection() as (conn, c):
                # Get year name
                c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (self.current_year_id,))
                year_result = c.fetchone()
                if not year_result:
                    return
                year_name = year_result[0]
                
                # Get terms for this year
                c.execute("""
                    SELECT TermID, TermName, StartDate, EndDate 
                    FROM Term 
                    WHERE AcademicYear = ? 
                    ORDER BY StartDate
                """, (self.current_year_id,))
                
                terms = c.fetchall()
                
                self.terms_table.setRowCount(len(terms))
                current_date = QDate.currentDate().toString("yyyy-MM-dd")
                has_current_term = False
                
                for i, (term_id, term_name, start_date, end_date) in enumerate(terms):
                    self.terms_table.setRowHeight(i, 40)
                    # Term Name
                    name_item = QTableWidgetItem(term_name)
                    name_item.setData(Qt.UserRole, term_id)
                    self.terms_table.setItem(i, 0, name_item)
                    
                    # Start Date
                    start_item = QTableWidgetItem(start_date)
                    self.terms_table.setItem(i, 1, start_item)
                    
                    # End Date
                    end_item = QTableWidgetItem(end_date)
                    self.terms_table.setItem(i, 2, end_item)
                    
                    # Calculate and display duration
                    start = QDate.fromString(start_date, "yyyy-MM-dd")
                    end = QDate.fromString(end_date, "yyyy-MM-dd")
                    days = start.daysTo(end)
                    duration_item = QTableWidgetItem(f"{days} days")
                    self.terms_table.setItem(i, 3, duration_item)
                    
                    # Check if current date is within term
                    if start_date <= current_date <= end_date:
                        has_current_term = True
                        for col in range(4):
                            item = self.terms_table.item(i, col)
                            if item:
                                item.setBackground(QColor("#e8f5e9"))
                                item.setForeground(QColor("#2e7d32"))
                    
                    # Add action comboBox
                    action_combo = QComboBox()
                    action_combo.addItems(["Select Actions", "Edit", "Delete"])
                    action_combo.currentIndexChanged.connect(
                        lambda index, tid=term_id, cb=action_combo: 
                        self.on_term_action_selected(index, tid, cb)
                    )
                    self.terms_table.setCellWidget(i, 4, action_combo)
                
                # Update statistics
                self.total_terms_label.setText(f"Total Terms: {len(terms)}")
                if has_current_term:
                    self.current_term_label.setText("Current Term: In Session")
                else:
                    self.current_term_label.setText("Current Term: Not in Session")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load terms: {str(e)}")

    def on_term_action_selected(self, index, term_id, combo_box):
        """Handle term action selection"""
        if index == 0:  # "Actions" placeholder
            return
        
        action = combo_box.currentText()
        combo_box.setCurrentIndex(0)  # Reset
        
        if "Edit" in action:
            self.editTerm_by_id(term_id)
        elif "Delete" in action:
            self.deleteTerm_by_id(term_id)

    def edit_selected_year(self):
        """Edit the currently selected year"""
        current_item = self.years_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "No Selection", "Please select an academic year to edit.")
            return
            
        year_id = current_item.data(Qt.UserRole)
        year_name = current_item.text()
        
        # Call your existing editYear method with the selected year
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Academic Year")
        dialog.setFixedSize(350, 200)
        dialog.setStyleSheet("""
            QDialog {
                background-color: #FFFFFF;
                border-radius: 10px;
            }
            QLabel {
                font-size: 16px;
                color: #2D3748;
            }
            QLineEdit {
                background-color: #F7FAFC;
                border: 1px solid #E0E3E7;
                padding: 10px;
                border-radius: 5px;
                font-size: 15px;
            }
            QPushButton {
                background-color: #4299E1;
                color: white;
                border: none;
                padding: 8px 18px;
                border-radius: 5px;
                font-size: 15px;
            }
            QPushButton#cancel {
                background-color: #EDF2F7;
                color: #2D3748;
                border: 1px solid #E0E3E7;
            }
            QPushButton:hover {
                background-color: #3182CE;
            }
            QPushButton#cancel:hover {
                background-color: #E2E8F0;
            }
            QLabel#error {
                color: #E53E3E;
                font-size: 13px;
            }
        """)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)

        label = QLabel("Edit Academic Year")
        layout.addWidget(label)

        year_input = QLineEdit(year_name)
        year_input.setPlaceholderText("e.g. 2024-2025")
        layout.addWidget(year_input)

        error_label = QLabel("")
        error_label.setObjectName("error")
        layout.addWidget(error_label)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("cancel")
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        def save_edit():
            new_year_name = year_input.text().strip()
            if not new_year_name:
                error_label.setText("Year name cannot be empty.")
                return
            
            try:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT 1 FROM AcademicYear WHERE Year = ? AND AcademicYearID != ?", 
                            (new_year_name, year_id))
                    if c.fetchone():
                        error_label.setText("This academic year already exists.")
                        return
                    
                    c.execute("UPDATE AcademicYear SET Year = ? WHERE AcademicYearID = ?", 
                            (new_year_name, year_id))
                    conn.commit()
                    dialog.accept()
                    QMessageBox.information(self, "Success", f"Academic Year updated successfully.")
                    self.load_academic_years()
                    
            except Exception as e:
                error_label.setText(f"Error updating year: {str(e)}")

        save_btn.clicked.connect(save_edit)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def delete_selected_year(self):
        """Delete the currently selected year"""
        current_item = self.years_list.currentItem()
        if not current_item:
            QMessageBox.warning(self, "No Selection", "Please select an academic year to delete.")
            return
            
        year_id = current_item.data(Qt.UserRole)
        year_name = current_item.text()
        
        # Call your existing deleteYear logic
        try:
            with get_db_connection() as (conn, c):
                # Check for child relations
                c.execute("SELECT 1 FROM Term WHERE AcademicYear = ?", (year_id,))
                if c.fetchone():
                    QMessageBox.warning(self, "Cannot Delete", 
                                      "This year has related terms and cannot be deleted.")
                    return
                
                # Confirmation
                reply = QMessageBox.question(
                    self, 'Confirm Delete',
                    f"Are you sure you want to delete the academic year '{year_name}'?\n\n"
                    "This action cannot be undone.",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
                
                if reply == QMessageBox.Yes:
                    c.execute("DELETE FROM AcademicYear WHERE AcademicYearID = ?", (year_id,))
                    conn.commit()
                    QMessageBox.information(self, "Deleted", f"Academic Year '{year_name}' deleted.")
                    self.current_year_id = None
                    self.load_academic_years()
                    self.load_terms_for_year()  # Clear terms display
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to delete year: {str(e)}")

    def addYear(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Academic Year")
        dialog.setFixedSize(350, 180)
        dialog.setStyleSheet("""
            QDialog {
            background-color: #FFFFFF;
            border-radius: 10px;
            }
            QLabel {
            font-size: 16px;
            color: #2D3748;
            }
            QLineEdit {
            background-color: #F7FAFC;
            border: 1px solid #E0E3E7;
            padding: 10px;
            border-radius: 5px;
            font-size: 15px;
            }
            QPushButton {
            background-color: #4299E1;
            color: white;
            border: none;
            padding: 8px 18px;
            border-radius: 5px;
            font-size: 15px;
            }
            QPushButton#cancel {
            background-color: #EDF2F7;
            color: #2D3748;
            border: 1px solid #E0E3E7;
            }
            QPushButton:hover {
            background-color: #3182CE;
            }
            QPushButton#cancel:hover {
            background-color: #E2E8F0;
            }
            QLabel#error {
            color: #E53E3E;
            font-size: 13px;
            }
        """)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)

        label = QLabel("Academic Year Name")
        layout.addWidget(label)

        year_input = QLineEdit()
        year_input.setPlaceholderText("e.g. 2024-2025")
        layout.addWidget(year_input)

        error_label = QLabel("")
        error_label.setObjectName("error")
        layout.addWidget(error_label)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        save_btn = QPushButton("Save")
        # save_btn.clicked.connect(lambda :self.save_year(dialog,year_input))
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("cancel")
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        def save_year():
            year_name = year_input.text().strip()
            if not year_name:
                error_label.setText("Year name cannot be empty.")
                return

            try:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT 1 FROM AcademicYear WHERE Year = ?", (year_name,))
                    if c.fetchone():
                        error_label.setText("This academic year already exists.")
                        return
                    
                    c.execute("INSERT INTO AcademicYear (Year) VALUES (?)", (year_name,))
                    conn.commit()
                    dialog.accept()
                    QMessageBox.information(self, "Success", f"Academic Year '{year_name}' added successfully.")
                    self.load_academic_years()
                    
            except sqlite3.Error as e:
                error_label.setText(f"Error saving year: {str(e)}")

        save_btn.clicked.connect(save_year)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()
                
    def addTerm(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Term")
        dialog.setFixedSize(400, 500)
        dialog.setStyleSheet("""
            QDialog {
                background-color: #FFFFFF;
                border-radius: 10px;
            }
            QLabel {
                font-size: 16px;
                color: #2D3748;
            }
            QComboBox, QDateEdit {
                background-color: #F7FAFC;
                border: 1px solid #E0E3E7;
                padding: 5px;
                border-radius: 5px;
                font-size: 15px;
            }
            QPushButton {
                background-color: #4299E1;
                color: white;
                border: none;
                padding: 8px 18px;
                border-radius: 5px;
                font-size: 15px;
            }
            QPushButton#cancel {
                background-color: #EDF2F7;
                color: #2D3748;
                border: 1px solid #E0E3E7;
            }
            QPushButton:hover {
                background-color: #3182CE;
            }
            QPushButton#cancel:hover {
                background-color: #E2E8F0;
            }
            QLabel#error {
                color: #E53E3E;
                font-size: 13px;
            }
        """)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)

        # Academic Year
        year_label = QLabel("Academic Year")
        layout.addWidget(year_label)
        year_combo = QComboBox()
        year_combo.addItem("Select Year", None)
        
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear")
                years = c.fetchall()
                for year in years:
                    year_combo.addItem(year[1], year[0])
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load academic years: {str(e)}")
            
        layout.addWidget(year_combo)

        # Term
        term_label = QLabel("Term")
        layout.addWidget(term_label)
        term_combo = QComboBox()
        term_combo.addItem("Select Term", None)
        term_combo.addItem("Term 1", "Term 1")
        term_combo.addItem("Term 2", "Term 2")
        term_combo.addItem("Term 3", "Term 3")
        layout.addWidget(term_combo)

        # Start Date
        start_label = QLabel("Starting Date")
        layout.addWidget(start_label)
        start_date = QDateEdit()
        start_date.setCalendarPopup(True)
        start_date.setDisplayFormat("yyyy-MM-dd")
        layout.addWidget(start_date)

        # End Date
        end_label = QLabel("Ending Date")
        layout.addWidget(end_label)
        end_date = QDateEdit()
        end_date.setCalendarPopup(True)
        end_date.setDisplayFormat("yyyy-MM-dd")
        layout.addWidget(end_date)

        # Error label
        error_label = QLabel("")
        error_label.setObjectName("error")
        layout.addWidget(error_label)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("cancel")
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        def save_term():
            year_id = year_combo.currentData()
            term = term_combo.currentData()
            start = start_date.date().toString("yyyy-MM-dd")
            end = end_date.date().toString("yyyy-MM-dd")

            if not year_id or not term or not start or not end:
                error_label.setText("All fields are required.")
                return

            if start > end:
                error_label.setText("Start date must be before end date.")
                return

            try:
                with get_db_connection() as (conn, c):
                    # Check for duplicate term in the same year
                    c.execute("SELECT 1 FROM Term WHERE TermName=? AND AcademicYear=?", (term, year_id))
                    if c.fetchone():
                        error_label.setText("This term already exists for the selected year.")
                        return

                    c.execute(
                        "INSERT INTO Term (TermName, AcademicYear, StartDate, EndDate) VALUES (?, ?, ?, ?)",
                        (term, year_id, start, end)
                    )
                    conn.commit()
                    dialog.accept()
                    QMessageBox.information(self, "Success", f"{term} added successfully.")
                    self.load_terms_for_year()
                    
            except sqlite3.Error as e:
                error_label.setText(f"Error saving term: {str(e)}")

        save_btn.clicked.connect(save_term)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()
    
    def editTerm_by_id(self, term_id):
        """Edit term using term_id directly"""
        try:
            with get_db_connection() as (conn, c):
                # Get term data using term_id
                c.execute("""
                    SELECT AcademicYear.Year, Term.TermName, Term.StartDate, Term.EndDate, 
                        Term.TermID, AcademicYear.AcademicYearID
                    FROM Term
                    JOIN AcademicYear ON Term.AcademicYear = AcademicYear.AcademicYearID
                    WHERE Term.TermID=?
                """, (term_id,))
                result = c.fetchone()
                
            if not result:
                QMessageBox.warning(self, "Error", "Could not find the selected term in the database.")
                return
                
            year, term, start, end, term_id, year_id = result

            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Term")
            dialog.setFixedSize(400, 500)
            dialog.setStyleSheet("""
                QDialog {
                    background-color: #FFFFFF;
                    border-radius: 10px;
                }
                QLabel {
                    font-size: 16px;
                    color: #2D3748;
                }
                QComboBox, QDateEdit {
                    background-color: #F7FAFC;
                    border: 1px solid #E0E3E7;
                    padding: 5px;
                    border-radius: 5px;
                    font-size: 15px;
                }
                QPushButton {
                    background-color: #4299E1;
                    color: white;
                    border: none;
                    padding: 8px 18px;
                    border-radius: 5px;
                    font-size: 15px;
                }
                QPushButton#cancel {
                    background-color: #EDF2F7;
                    color: #2D3748;
                    border: 1px solid #E0E3E7;
                }
                QPushButton:hover {
                    background-color: #3182CE;
                }
                QPushButton#cancel:hover {
                    background-color: #E2E8F0;
                }
                QLabel#error {
                    color: #E53E3E;
                    font-size: 13px;
                }
            """)

            layout = QVBoxLayout(dialog)
            layout.setContentsMargins(20, 20, 20, 20)

            # Academic Year
            year_label = QLabel("Academic Year")
            layout.addWidget(year_label)
            year_combo = QComboBox()
            
            with get_db_connection() as (conn, c):
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear")
                years = c.fetchall()
                for y in years:
                    year_combo.addItem(y[1], y[0])
            
            # Set current year
            idx = [y[0] for y in years].index(year_id) if year_id in [y[0] for y in years] else 0
            year_combo.setCurrentIndex(idx)
            layout.addWidget(year_combo)

            # Term
            term_label = QLabel("Term")
            layout.addWidget(term_label)
            term_combo = QComboBox()
            term_combo.addItem("Term 1", "Term 1")
            term_combo.addItem("Term 2", "Term 2")
            term_combo.addItem("Term 3", "Term 3")
            # Set current term
            idx = ["Term 1", "Term 2", "Term 3"].index(term) if term in ["Term 1", "Term 2", "Term 3"] else 0
            term_combo.setCurrentIndex(idx)
            layout.addWidget(term_combo)

            # Start Date
            start_label = QLabel("Starting Date")
            layout.addWidget(start_label)
            start_date = QDateEdit()
            start_date.setCalendarPopup(True)
            start_date.setDisplayFormat("yyyy-MM-dd")
            start_date.setDate(QDate.fromString(start, "yyyy-MM-dd"))
            layout.addWidget(start_date)

            # End Date
            end_label = QLabel("Ending Date")
            layout.addWidget(end_label)
            end_date = QDateEdit()
            end_date.setCalendarPopup(True)
            end_date.setDisplayFormat("yyyy-MM-dd")
            end_date.setDate(QDate.fromString(end, "yyyy-MM-dd"))
            layout.addWidget(end_date)

            # Error label
            error_label = QLabel("")
            error_label.setObjectName("error")
            layout.addWidget(error_label)

            # Buttons
            btn_layout = QHBoxLayout()
            btn_layout.addStretch()
            save_btn = QPushButton("Save")
            cancel_btn = QPushButton("Cancel")
            cancel_btn.setObjectName("cancel")
            btn_layout.addWidget(save_btn)
            btn_layout.addWidget(cancel_btn)
            layout.addLayout(btn_layout)

            def save_edit():
                new_year_id = year_combo.currentData()
                new_term = term_combo.currentData()
                new_start = start_date.date().toString("yyyy-MM-dd")
                new_end = end_date.date().toString("yyyy-MM-dd")

                if not new_year_id or not new_term or not new_start or not new_end:
                    error_label.setText("All fields are required.")
                    return

                if new_start > new_end:
                    error_label.setText("Start date must be before end date.")
                    return

                try:
                    with get_db_connection() as (conn, c):
                        # Check for duplicate term in the same year (excluding current term)
                        c.execute("""
                            SELECT 1 FROM Term 
                            WHERE TermName=? AND AcademicYear=? AND TermID<>?
                        """, (new_term, new_year_id, term_id))
                        if c.fetchone():
                            error_label.setText("This term already exists for the selected year.")
                            return

                        c.execute("""
                            UPDATE Term SET TermName=?, AcademicYear=?, StartDate=?, EndDate=?
                            WHERE TermID=?
                        """, (new_term, new_year_id, new_start, new_end, term_id))
                        conn.commit()
                        dialog.accept()
                        QMessageBox.information(self, "Success", f"{new_term} updated successfully.")
                        self.load_terms_for_year()
                        
                except sqlite3.Error as e:
                    error_label.setText(f"Error updating term: {str(e)}")

            save_btn.clicked.connect(save_edit)
            cancel_btn.clicked.connect(dialog.reject)
            dialog.exec_()
            
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load term data: {str(e)}")
    
    def deleteTerm_by_id(self, term_id):
        """Delete term using term_id directly"""
        try:
            with get_db_connection() as (conn, c):
                # Get term details for display in confirmation message
                c.execute("""
                    SELECT AcademicYear.Year, Term.TermName, Term.StartDate, Term.EndDate
                    FROM Term
                    JOIN AcademicYear ON Term.AcademicYear = AcademicYear.AcademicYearID
                    WHERE Term.TermID=?
                """, (term_id,))
                result = c.fetchone()
                
            if not result:
                QMessageBox.warning(self, "Error", "Could not find the selected term in the database.")
                return
                
            year, term, start, end = result

            # Confirmation message
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("Delete Term")
            msg.setText(
                f"Are you sure you want to delete the term '{term}' for academic year '{year}'?\n\n"
                "This action cannot be undone.\n\n"
                "Note:\n"
                "- If this term is linked to any fees, deletion is not allowed.\n"
                "- If students are assigned to this term, their term assignment will be set to 'None'."
            )
            msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            ret = msg.exec_()
            
            if ret != QMessageBox.Yes:
                return

            try:
                with get_db_connection() as (conn, c):
                    # Check for Fees relation
                    c.execute("SELECT 1 FROM Fees WHERE Term = ?", (term_id,))
                    if c.fetchone():
                        QMessageBox.critical(self, "Restricted", 
                            "This term cannot be deleted because it is linked to one or more fees records.")
                        return

                    # Set Student.TermID to NULL for students assigned to this term
                    c.execute("UPDATE StudentTerm SET TermID = NULL WHERE TermID = ?", (term_id,))
                    
                    # Delete the term
                    c.execute("DELETE FROM Term WHERE TermID = ?", (term_id,))
                    conn.commit()
                    
                    QMessageBox.information(self, "Deleted", 
                        f"Term '{term}' for academic year '{year}' has been deleted.\n\n"
                        "Any students previously assigned to this term now have no term assigned.")
                    
                    # Refresh the table
                    self.load_terms_for_year()
                    
            except sqlite3.Error as e:
                QMessageBox.critical(self, "Error", f"An error occurred while deleting the term: {str(e)}")
                
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load term data: {str(e)}")    


class Class(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        # Header row
        add_class_btn = QPushButton("Add Class")
        add_class_btn.clicked.connect(self.addClass)
        export_btn = QPushButton("Export")
        export_btn.clicked.connect(export_table_to_file(self,))
        # Filter row
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(20, 20, 20, 5)
        self.Search_Input = QLineEdit()
        self.Search_Input.setPlaceholderText("Search...")
        self.Search_Input.setClearButtonEnabled(True)
        self.Search_Input.setMaximumWidth(200)
        self.year_filter = QComboBox()
        self.year_filter.addItem("All Years", None)
        with get_db_connection() as (conn, c):
            c.execute("SELECT AcademicYearID, Year FROM AcademicYear")
            years = c.fetchall()
        for year in years:
            self.year_filter.addItem(year[1], year[0])
        self.year_filter.setCurrentIndex(0)
        self.year_filter.setMaximumWidth(150)
        filter_layout.addWidget(self.Search_Input)
        filter_layout.addWidget(self.year_filter)
        filter_layout.addStretch()
        filter_layout.addWidget(add_class_btn)
        filter_layout.addWidget(export_btn)
        self.results_label = QLabel("0 results")
        self.results_label.setStyleSheet("font-size: 14px; color: #718096;")
        filter_layout.addWidget(self.results_label)
        layout.addLayout(filter_layout)

        self.Search_Input.textChanged.connect(self.Load_Data)
        self.year_filter.currentIndexChanged.connect(self.Load_Data)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["Class Name", "Academic Year", "Action"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        
        layout.addWidget(self.table)

        self.Load_Data()
        self.update_results_count()

    def Load_Data(self):
        self.table.setRowCount(0)
        selected_year_id = self.year_filter.currentData()
        query = """
            SELECT Class.ClassID, Class.ClassName, AcademicYear.Year
            FROM Class
            JOIN AcademicYear ON Class.AcademicYear = AcademicYear.AcademicYearID ORDER BY AcademicYear.Year DESC
        """
        params = ()
        search_text = self.Search_Input.text().strip()
        filters = []
        if selected_year_id:
            filters.append("Class.AcademicYear = ?")
            params += (selected_year_id,)
        if search_text:
            filters.append("(Class.ClassName LIKE ?)")
            params += (f"%{search_text}%",)
        if filters:
            query += " WHERE " + " AND ".join(filters)
        with get_db_connection() as (conn, c):
            c.execute(query, params)
            results = c.fetchall()
        for row_data in results:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setRowHeight(row, 50)
            # ClassName, Year
            self.table.setItem(row, 0, QTableWidgetItem(str(row_data[1])))
            self.table.setItem(row, 1, QTableWidgetItem(str(row_data[2])))
            
            # Create action ComboBox
            class_id = row_data[0]
            action_combo = QComboBox()
            action_combo.addItems(["Select Action","Edit","Delete"])
            action_combo.currentIndexChanged.connect(lambda index, r=row, cid=class_id: self.handle_class_action(index, r, cid))
            
            self.table.setCellWidget(row, 2, action_combo)
        # self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.update_results_count()
    
    def handle_class_action(self, index, row, class_id):
        action = self.table.cellWidget(row, 2).itemData(index)
        if index == 0:  # "Select Action"
            return

        action = self.table.cellWidget(row, 2).itemText(index)
        self.table.cellWidget(row, 2).setCurrentIndex(0)  # Reset to "Select Action"

        if action == "Edit":
            self.editClass(class_id)
        elif action == "Delete":
            self.deleteClass(class_id)    

    def update_results_count(self):
        count = self.table.rowCount()
        self.results_label.setText(f"{count} results")

    def addClass(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Class")
        dialog.setFixedSize(350, 250)
        dialog.setStyleSheet("""
            QDialog {
            background-color: #FFFFFF;
            border-radius: 10px;
            }
            QLabel {
            font-size: 16px;
            color: #2D3748;
            }
            QComboBox, QLineEdit {
            background-color: #F7FAFC;
            border: 1px solid #E0E3E7;
            padding: 10px;
            border-radius: 5px;
            font-size: 15px;
            }
            QPushButton {
            background-color: #4299E1;
            color: white;
            border: none;
            padding: 8px 18px;
            border-radius: 5px;
            font-size: 15px;
            }
            QPushButton#cancel {
            background-color: #EDF2F7;
            color: #2D3748;
            border: 1px solid #E0E3E7;
            }
            QPushButton:hover {
            background-color: #3182CE;
            }
            QPushButton#cancel:hover {
            background-color: #E2E8F0;
            }
            QLabel#error {
            color: #E53E3E;
            font-size: 13px;
            }
        """)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)

        # Academic Year
        year_label = QLabel("Academic Year")
        layout.addWidget(year_label)
        year_combo = QComboBox()
        year_combo.addItem("Select Year", None)
        with get_db_connection() as (conn, c): 
            c.execute("SELECT AcademicYearID, Year FROM AcademicYear")
            years = c.fetchall()
        for year in years:
            year_combo.addItem(year[1], year[0])
        layout.addWidget(year_combo)

        # Class Name
        class_label = QLabel("Class Name")
        layout.addWidget(class_label)
        class_input = QLineEdit()
        class_input.setPlaceholderText("e.g. Form 1A")
        layout.addWidget(class_input)

        error_label = QLabel("")
        error_label.setObjectName("error")
        layout.addWidget(error_label)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("cancel")
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        def save_class():
            year_id = year_combo.currentData()
            class_name = class_input.text().strip()
            if not year_id or not class_name:
                error_label.setText("Year and Class Name are required.")
                return
            with get_db_connection() as (conn, c):
                c.execute("SELECT 1 FROM Class WHERE ClassName=? AND AcademicYear=?", (class_name, year_id))
                if c.fetchone():
                    error_label.setText("This class already exists for the selected year.")
                    return
                try:
                    c.execute(
                        "INSERT INTO Class (ClassName, AcademicYear) VALUES (?, ?)",
                        (class_name, year_id)
                    )
                    conn.commit()
                    dialog.accept()
                    QMessageBox.information(self, "Success", f"Class '{class_name}' added successfully.")
                    self.Load_Data()
                except Exception as e:
                    error_label.setText("Error saving class.")

        save_btn.clicked.connect(save_class)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def editClass(self, class_id):
        try:
            with get_db_connection() as (conn, c):
                # Get current class details
                c.execute("SELECT ClassName, AcademicYear FROM Class WHERE ClassID=?", (class_id,))
                result = c.fetchone()
                
                if not result:
                    QMessageBox.warning(self, "Error", "Could not find the selected class in the database.")
                    return
                
                class_name, current_year_id = result
                
                # Check if class has students enrolled
                c.execute("SELECT COUNT(*) FROM StudentTerm WHERE ClassID = ?", (class_id,))
                student_count = c.fetchone()[0]
                has_students = student_count > 0
                
                # Get all academic years for dropdown
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY Year DESC")
                years = c.fetchall()
                
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load class data: {str(e)}")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Class")
        dialog.setFixedSize(400, 300)
        dialog.setStyleSheet("""
            QDialog {
                background-color: #FFFFFF;
                border-radius: 10px;
            }
            QLabel {
                font-size: 14px;
                color: #2D3748;
            }
            QComboBox, QLineEdit {
                background-color: #F7FAFC;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton {
                background-color: #4299E1;
                color: white;
                border: none;
                padding: 8px 18px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton#cancel {
                background-color: #EDF2F7;
                color: #2D3748;
                border: 1px solid #E0E3E7;
            }
            QPushButton:hover {
                background-color: #3182CE;
            }
            QPushButton#cancel:hover {
                background-color: #E2E8F0;
            }
            QLabel#error {
                color: #E53E3E;
                font-size: 13px;
            }
            QLabel#warning {
                color: #D69E2E;
                font-size: 12px;
            }
        """)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(8)

        # Class Name (ALWAYS editable)
        class_label = QLabel("Class Name")
        layout.addWidget(class_label)
        
        class_input = QLineEdit(class_name)
        layout.addWidget(class_input)

        # Academic Year 
        year_label = QLabel("Academic Year")
        layout.addWidget(year_label)
        
        year_combo = QComboBox()
        
        # Populate academic years
        for year_id, year_name in years:
            year_combo.addItem(year_name, year_id)
        
        # Set current year
        for i in range(year_combo.count()):
            if year_combo.itemData(i) == current_year_id:
                year_combo.setCurrentIndex(i)
                break
        
        # Disable year combo if class has students
        if has_students:
            year_combo.setEnabled(False)
            year_combo.setStyleSheet("""
                QComboBox {
                    background-color: #EDF2F7;
                    color: #718096;
                }
            """)
            
            # Add warning message
            warning_label = QLabel(f"⚠️ Academic year cannot be changed ({student_count} student(s) enrolled)")
            warning_label.setObjectName("warning")
            layout.addWidget(warning_label)
        
        layout.addWidget(year_combo)

        error_label = QLabel("")
        error_label.setObjectName("error")
        layout.addWidget(error_label)

        # Button layout
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        save_btn = QPushButton("Save Changes")
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("cancel")
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        def save_edit():
            # Get values from form
            new_class_name = class_input.text().strip()
            new_year_id = year_combo.currentData()
            
            # Clear previous errors
            error_label.setText("")
            
            # Validate inputs
            if not new_class_name:
                error_label.setText("Class name is required.")
                return
            
            if not new_year_id:
                error_label.setText("Academic year is required.")
                return
            
            # Check year change restriction
            if has_students and new_year_id != current_year_id:
                error_label.setText("Cannot change academic year because students are enrolled in this class.")
                return
            
            # Check for duplicate class name in the SAME academic year
            try:
                with get_db_connection() as (conn, c):
                    # Query: Check if another class with same name exists in the same year
                    c.execute("""
                        SELECT COUNT(*) FROM Class 
                        WHERE ClassName = ? 
                        AND AcademicYear = ? 
                        AND ClassID != ?
                    """, (new_class_name, new_year_id, class_id))
                    
                    duplicate_count = c.fetchone()[0]
                    
                    if duplicate_count > 0:
                        error_label.setText(f"A class named '{new_class_name}' already exists in this academic year.")
                        return
                    
                    # Update the class
                    c.execute("""
                        UPDATE Class 
                        SET ClassName = ?, AcademicYear = ? 
                        WHERE ClassID = ?
                    """, (new_class_name, new_year_id, class_id))
                    
                    conn.commit()
                    
                    # Success
                    dialog.accept()
                    QMessageBox.information(self, "Success", "Class updated successfully.")
                    self.Load_Data()
                    
            except sqlite3.Error as e:
                error_label.setText(f"Database error: {str(e)}")
            except Exception as e:
                error_label.setText(f"Error: {str(e)}")

        save_btn.clicked.connect(save_edit)
        cancel_btn.clicked.connect(dialog.reject)
        
        dialog.exec_() 

    def deleteClass(self, class_id):
        # Check for child relations (e.g., students assigned to this class)
        with get_db_connection() as (conn, c): 
            c.execute("SELECT 1 FROM StudentTerm WHERE ClassID = ?", (class_id,))
            if c.fetchone():
                QMessageBox.critical(self, "Restricted", "This class cannot be deleted because it is linked to one or more students.")
                return

            # Get class name for message
            c.execute("SELECT ClassName FROM Class WHERE ClassID = ?", (class_id,))
            row = c.fetchone()
            class_name = row[0] if row else "the class"

        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("Delete Class")
        msg.setText(
            f"Are you sure you want to delete '{class_name}'?\n\n"
            "This action cannot be undone.\n\n"
            "Note:\n"
            "- If this class is linked to any students, deletion is not allowed."
        )
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        ret = msg.exec_()
        if ret != QMessageBox.Yes:
            return

        try:
            with get_db_connection() as (conn, c):
                c.execute("DELETE FROM Class WHERE ClassID = ?", (class_id,))
                conn.commit()
                QMessageBox.information(self, "Deleted", f"Class '{class_name}' has been deleted successfully.")
                self.Load_Data()
        except Exception as e:
            QMessageBox.critical(self, "Error", "An error occurred while deleting the class.")

#### Fees Management Pages #####
class Fees(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(15)
        Tab = QTabWidget()
        self.Feesmanagement = FeesManagementPage()
        Tab.addTab(self.Feesmanagement,"Fees")
        self.FeesPayment = FeesPaymentPage()
        Tab.addTab(self.FeesPayment,"Payment")
        self.PaymentHistory = PaymentHistoryPage()
        Tab.addTab(self.PaymentHistory,"History")
        layout.addWidget(Tab)

class FeesManagementPage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        # Filters and add buttons
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search...")
        self.search_input.setMaximumWidth(200)
        self.search_input.textChanged.connect(self.load_data)
        self.AcademicYear_filter = QComboBox()
        with get_db_connection() as (conn,c):
            c.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY YEAR DESC")
            for years in c.fetchall():
                self.AcademicYear_filter.addItem(years[1],years[0])
            if c.fetchall() is None:
                self.AcademicYear_filter.addItem("No years",None)
        self.AcademicYear_filter.currentIndexChanged.connect(self.load_data)
        self.term_filter = QComboBox()
        self.term_filter.addItem("All Terms", None)
        with get_db_connection() as (conn, c):
            c.execute("SELECT TermID, TermName FROM Term")
            for t in c.fetchall():
                self.term_filter.addItem(t[1], t[0])
        self.term_filter.currentIndexChanged.connect(self.load_data)
        self.AcademicYear_filter.currentIndexChanged.connect(lambda: self.update_term_combo())

        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(self.AcademicYear_filter)
        filter_layout.addWidget(self.term_filter)
        filter_layout.addStretch()
        add_btn = QPushButton("Add Fee Structure")
        add_btn.clicked.connect(self.show_add_fee_dialog)
        # import_btn = QPushButton("Import Fees")
        # import_btn.clicked.connect(self.show_import_wizard)
        filter_layout.addWidget(add_btn)
        # filter_layout.addWidget(import_btn)
        layout.addLayout(filter_layout)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(5)  
        self.table.setHorizontalHeaderLabels(["Year", "Term","Class","Amount", "Actions"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # Load data (placeholder)
        self.load_data()
 

    def update_term_combo(self):
        year_id = self.AcademicYear_filter.currentData()
        self.term_filter.clear()
        self.term_filter.addItem("All Terms", None)
        
        if year_id:
            with get_db_connection() as (conn, c):
                c.execute("SELECT TermID, TermName FROM Term WHERE AcademicYear = ?", (year_id,))
                for tid, tname in c.fetchall():
                    self.term_filter.addItem(tname, tid)

    def load_data(self):
        try:
            with get_db_connection() as (conn, c):
                self.table.setRowCount(0)
                
                # Query to join all four tables and get the required information
                query = """
                    SELECT 
                        f.FeesID,
                        ay.Year AS AcademicYear,
                        t.TermName AS Term,
                        c.ClassName AS Class,
                        f.Amount
                    FROM Fees f
                    JOIN Class c ON f.Class = c.ClassID
                    JOIN Term t ON f.Term = t.TermID
                    JOIN AcademicYear ay ON c.AcademicYear = ay.AcademicYearID
                """
                
                # Build filters if needed (you can add filter functionality later)
                filters = []
                params = []
                
                # Example: Add search filter if needed
                search_text = self.search_input.text().strip() if hasattr(self, 'search_input') else ""
                if search_text:
                    filters.append("(ay.Year LIKE ? OR t.TermName LIKE ? OR c.ClassName LIKE ? OR f.Amount LIKE ?)")
                    params.extend([f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%"])
                # academic year filter
                selected_year_id = self.AcademicYear_filter.currentData()
                if selected_year_id:
                    filters.append("ay.AcademicYearID = ?")
                    params.append(selected_year_id)
                # term filter
                selected_term_id = self.term_filter.currentData()
                if selected_term_id:
                    filters.append("t.TermID = ?")
                    params.append(selected_term_id)
                if filters:
                    query += " WHERE " + " AND ".join(filters)
                    query += " ORDER BY ay.Year DESC, t.TermName, c.ClassName"
                
                # Execute query
                c.execute(query, params)
                rows = c.fetchall()
                
                # Set table columns
                self.table.setColumnCount(5)  # Year, Term, Class, Amount, Actions
                self.table.setHorizontalHeaderLabels(["Academic Year", "Term", "Class", "Amount", "Actions"])
                
                # Populate table
                for row_data in rows:
                    row = self.table.rowCount()
                    self.table.insertRow(row)
                    
                    # Add data columns
                    for col, data in enumerate(row_data[1:5]):  # Skip FeesID (index 0)
                        item = QTableWidgetItem(str(data))
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        self.table.setItem(row, col, item)
                    
                    # Add action buttons - pass FeesID for edit/delete operations
                    fees_id = row_data[0]  # First column is FeesID
                    self.add_action_buttons(row, fees_id)
                    self.table.setRowHeight(row, 40)
                # Update results count
                count = self.table.rowCount()
                if hasattr(self, 'results_label'):
                    self.results_label.setText(f"{count} results")
                
                # Ensure proper column resizing
                self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
                self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Fixed)  # Actions column fixed width
                self.table.setColumnWidth(4, 120)  # Width for action buttons
                
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Error loading fees data: {str(e)}")

    def add_action_buttons(self, row, fees_id):
        """Add action ComboBox to table row for fees"""
        action_combo = QComboBox()
        action_combo.addItems(["Select Action","Edit","Delete"])
        action_combo.currentIndexChanged.connect(lambda index, fid=fees_id: self.handle_fee_action(index, fid, row))
        
        # Set the widget in the Actions column (column 4)
        self.table.setCellWidget(row, 4, action_combo)

    def handle_fee_action(self, index, fees_id, row):
        if index == 0:  # "Select Action"
            return

        action = self.table.cellWidget(row, 4).itemText(index)
        self.table.cellWidget(row, 4).setCurrentIndex(0)  # Reset to "Select Action"

        if action == "Edit":
            self.edit_fee(fees_id)
        elif action == "Delete":
            self.delete_fee(fees_id)    

    def edit_fee(self, fees_id):
        """Edit fee record"""
        try:
            with get_db_connection() as (conn, c):
                # Get current fee data
                c.execute("""
                    SELECT f.FeesID, f.Class, f.Term, f.Amount, 
                        c.ClassName, t.TermName, ay.Year, ay.AcademicYearID
                    FROM Fees f
                    JOIN Class c ON f.Class = c.ClassID
                    JOIN Term t ON f.Term = t.TermID
                    JOIN AcademicYear ay ON c.AcademicYear = ay.AcademicYearID
                    WHERE f.FeesID = ?
                """, (fees_id,))
                fee_data = c.fetchone()
            
            if not fee_data:
                QMessageBox.warning(self, "Error", "Fee record not found")
                return
                
            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Fee")
            dialog.setModal(True)
            dialog.setMinimumWidth(500)
            
            layout = QVBoxLayout(dialog)
            form_layout = QFormLayout()
            
            # Academic Year (read-only since it's determined by Class)
            year_label = QLabel(fee_data[6])  # Year from AcademicYear
            form_layout.addRow("Academic Year:", year_label)
            
            # Class selection
            class_combo = QComboBox()
            class_combo.addItem("Select Class", None)
            
            # Term selection
            term_combo = QComboBox()
            term_combo.addItem("Select Term", None)
            
            # Amount input
            amount_edit = QLineEdit(str(fee_data[3]))
            amount_edit.setValidator(QDoubleValidator(0, 1000000, 2))  # Allow decimal values
            
            # Load classes and terms
            with get_db_connection() as (conn, c):
                # Load classes
                c.execute(f"SELECT ClassID, ClassName FROM Class WHERE AcademicYear ={fee_data[7]} ORDER BY ClassName")
                classes = c.fetchall()
                for class_id, class_name in classes:
                    class_combo.addItem(class_name, class_id)
                
                # Load terms
                c.execute(f"SELECT TermID, TermName FROM Term WHERE AcademicYear ={fee_data[7]} ORDER BY TermName")
                terms = c.fetchall()
                for term_id, term_name in terms:
                    term_combo.addItem(term_name, term_id)
            
            # Set current values
            current_class_index = class_combo.findData(fee_data[1])
            current_term_index = term_combo.findData(fee_data[2])
            
            if current_class_index >= 0:
                class_combo.setCurrentIndex(current_class_index)
            if current_term_index >= 0:
                term_combo.setCurrentIndex(current_term_index)
            
            form_layout.addRow("Class:", class_combo)
            form_layout.addRow("Term:", term_combo)
            form_layout.addRow("Amount:", amount_edit)
            
            layout.addLayout(form_layout)
            
            button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            button_box.accepted.connect(dialog.accept)
            button_box.rejected.connect(dialog.reject)
            layout.addWidget(button_box)
            
            if dialog.exec_() == QDialog.Accepted:
                # Validation (similar to save_fee)
                class_id = class_combo.currentData()
                term_id = term_combo.currentData()
                amount_str = amount_edit.text().strip()
                
                # Check required fields
                if not all([class_id, term_id, amount_str]):
                    QMessageBox.warning(dialog, "Error", "All fields are required")
                    return
                
                # Validate amount is a valid number
                try:
                    amount = float(amount_str)
                    if amount <= 0:
                        QMessageBox.warning(dialog, "Error", "Amount must be greater than 0")
                        return
                except ValueError:
                    QMessageBox.warning(dialog, "Error", "Amount must be a valid number")
                    return
                
                # Check for duplicate fee structure (excluding current record)
                try:
                    with get_db_connection() as (conn, c):
                        c.execute("""
                            SELECT COUNT(*) FROM Fees 
                            WHERE Class = ? AND Term = ? AND FeesID != ?
                        """, (str(class_id), str(term_id), fees_id))
                        
                        if c.fetchone()[0] > 0:
                            QMessageBox.warning(dialog, "Error", 
                                            "Duplicate fee structure for this class and term")
                            return
                        
                        # Update the fee record
                        c.execute("""
                            UPDATE Fees 
                            SET Class = ?, Term = ?, Amount = ?
                            WHERE FeesID = ?
                        """, (
                            str(class_id),
                            str(term_id),
                            amount,
                            fees_id
                        ))
                        conn.commit()
                        QMessageBox.information(dialog, "Success", "Fee updated successfully")
                        self.load_data()
                        
                except sqlite3.Error as e:
                    QMessageBox.critical(dialog, "Database Error", f"Failed to update fee: {str(e)}")
                except Exception as e:
                    QMessageBox.critical(dialog, "Error", f"An error occurred: {str(e)}")
                        
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load fee data: {str(e)}")
    
    def delete_fee(self, fees_id):
        """Delete fee record with confirmation"""
        reply = QMessageBox.question(
            self, 
            'Confirm Deletion', 
            'Are you sure you want to delete this fee record?',
            QMessageBox.Yes | QMessageBox.No, 
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    c.execute("DELETE FROM Fees WHERE FeesID = ?", (fees_id,))
                    conn.commit()
                    QMessageBox.information(self, "Success", "Fee record deleted successfully")
                    self.load_data()
            except sqlite3.Error as e:
                QMessageBox.critical(self, "Error", f"Failed to delete fee: {str(e)}")
            
    def populate_combos(self, academic_year_combo, class_combo, term_combo):
        """Populate academic year combo and set up signals to update class/term combos"""
        try:
            with get_db_connection() as (conn, c):
                # Clear existing items
                academic_year_combo.clear()
                class_combo.clear()
                term_combo.clear()
                
                # Add default items
                academic_year_combo.addItem("Select Academic Year", None)
                class_combo.addItem("Select Class", None)
                term_combo.addItem("Select Term", None)
                
                # Populate academic years
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY Year")
                years = c.fetchall()
                for year_id, year_name in years:
                    academic_year_combo.addItem(year_name, year_id)
                
                # Connect signal to update class and term combos when year changes
                academic_year_combo.currentIndexChanged.connect(
                    lambda: self.update_class_term_combos(academic_year_combo, class_combo, term_combo)
                )
                
        except sqlite3.Error as e:
            QMessageBox.warning(f"Error populating academic years: {e}")

    def update_class_term_combos(self, academic_year_combo, class_combo, term_combo):
        """Update class and term combos based on selected academic year"""
        selected_year_id = academic_year_combo.currentData()
        
        # Clear existing items (except the first "Select" item)
        class_combo.clear()
        term_combo.clear()
        class_combo.addItem("Select Class", None)
        term_combo.addItem("Select Term", None)
        
        if not selected_year_id:
            return
        
        try:
            with get_db_connection() as (conn, c):
                # Populate classes for selected academic year
                c.execute("SELECT ClassID, ClassName FROM Class WHERE AcademicYear = ? ORDER BY ClassName", 
                        (selected_year_id,))
                classes = c.fetchall()
                for class_id, class_name in classes:
                    class_combo.addItem(class_name, class_id)
                
                # Populate terms for selected academic year
                c.execute("SELECT TermID, TermName FROM Term WHERE AcademicYear = ? ORDER BY TermName", 
                        (selected_year_id,))
                terms = c.fetchall()
                for term_id, term_name in terms:
                    term_combo.addItem(term_name, term_id)
                    
        except sqlite3.Error as e:
            QMessageBox.warning(f"Error updating class/term combos: {e}")

    def show_add_fee_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Fee Structure")
        layout = QVBoxLayout()
        form_layout = QFormLayout()
        year_combo = QComboBox()
        class_combo = QComboBox()
        term_combo = QComboBox()
        amount_edit = QLineEdit()
        self.populate_combos(year_combo,class_combo, term_combo)  # Assume this populates with names and IDs as data
        form_layout.addRow("Academic Year:", year_combo)
        form_layout.addRow("Class:", class_combo)
        form_layout.addRow("Term:", term_combo)
        form_layout.addRow("Amount:", amount_edit)
        layout.addLayout(form_layout)
        buttons = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_fee(dialog, class_combo, term_combo, amount_edit))
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        buttons.addWidget(save_btn)
        buttons.addWidget(close_btn)
        layout.addLayout(buttons)
        dialog.setLayout(layout)
        dialog.exec_()

    def save_fee(self, dialog, class_combo, term_combo, amount_edit):
        class_id = class_combo.currentData()
        term_id = term_combo.currentData()
        amount_str = amount_edit.text().strip()
        if not all([class_id, term_id, amount_str]):
            QMessageBox.warning(dialog, "Error", "All fields are required")
            return
        try:
            amount = float(amount_str)
        except ValueError:
            QMessageBox.warning(dialog, "Error", "Amount must be a valid number")
            return
        with get_db_connection() as (conn, c):
            try:
                c.execute("SELECT COUNT(*) FROM Fees WHERE Class = ? AND Term = ?", (str(class_id), str(term_id)))
                if c.fetchone()[0] > 0:
                    QMessageBox.warning(dialog, "Error", "Duplicate fee structure for this class and term")
                    return
                c.execute("""INSERT INTO Fees (Class, Term, Amount)
                            VALUES (?, ?, ?)""", (str(class_id), str(term_id), amount))
                conn.commit()
                QMessageBox.information(dialog, "Success", "Fee structure added successfully")
                amount_edit.clear()
                class_combo.setCurrentIndex(0)
                term_combo.setCurrentIndex(0)
                self.load_data()  # Refresh table if needed
            except Exception as e:
                conn.rollback()
                QMessageBox.warning(dialog, "Error", f"Failed to add fee structure: {str(e)}")
            finally:
                conn.close()

class FeesPaymentPage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setSpacing(15)


        # Filters
        filter_layout = QHBoxLayout()
        
        # Search
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search students...")
        self.search_input.setMaximumWidth(200)
        self.search_input.textChanged.connect(self.load_data)
        
        # Year filter
        self.year_combo = QComboBox()
        self.year_combo.addItem("All Years", None)
        self.year_combo.currentIndexChanged.connect(self.load_data)
        
        # Term filter
        self.term_combo = QComboBox()
        self.term_combo.addItem("All Terms", None)
        self.term_combo.currentIndexChanged.connect(self.load_data)
        
        # Class filter
        self.class_combo = QComboBox()
        self.class_combo.addItem("All Classes", None)
        self.class_combo.currentIndexChanged.connect(self.load_data)
        
        # Load filter options
        self.load_filter_options()
        
        filter_layout.addWidget(QLabel("Search:"))
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(QLabel("Year:"))
        filter_layout.addWidget(self.year_combo)
        filter_layout.addWidget(QLabel("Term:"))
        filter_layout.addWidget(self.term_combo)
        filter_layout.addWidget(QLabel("Class:"))
        filter_layout.addWidget(self.class_combo)
        filter_layout.addStretch()
        
        layout.addLayout(filter_layout)

        # Action buttons
        action_layout = QHBoxLayout()
        
        generate_btn = QPushButton("Generate Template")
        generate_btn.clicked.connect(self.generate_payment_template_dialog)
        import_btn = QPushButton("Import Payments")
        import_btn.clicked.connect(self.import_payments_dialog)
        action_layout.addWidget(generate_btn)
        action_layout.addWidget(import_btn)
        action_layout.addStretch()
        layout.addLayout(action_layout)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "Name", "Term", "Class", "Fee", 
            "Amount Paid", "UnPaid", "Status", "Payment History"
        ])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)        
        # Load initial data
        self.load_data()

    def load_filter_options(self):
        """Load filter options from database"""
        try:
            with get_db_connection() as (conn, c):
                # Load years
                c.execute("SELECT DISTINCT ay.Year FROM AcademicYear ay ORDER BY ay.Year DESC")
                years = c.fetchall()
                for year in years:
                    self.year_combo.addItem(year[0], year[0])
                
                # Load terms
                c.execute("SELECT DISTINCT t.TermName FROM Term t ORDER BY t.TermName")
                terms = c.fetchall()
                for term in terms:
                    self.term_combo.addItem(term[0], term[0])
                
                # Load classes
                c.execute("SELECT DISTINCT c.ClassName FROM Class c ORDER BY c.ClassName")
                classes = c.fetchall()
                for class_ in classes:
                    self.class_combo.addItem(class_[0], class_[0])
                    
        except Exception as e:
            QMessageBox.warning(f"Error loading filter options: {e}")

    def load_data(self):
        """Load payment data with filters"""
        try:
            with get_db_connection() as (conn, c):
                query = """
                SELECT 
                    s.Student_Name,
                    t.TermName || ' (' || ay.Year || ')' as TermWithYear,
                    c.ClassName,
                    ROUND(f.Amount * (st.PayRate / 100.0), 2) as FeeAfterDiscount,
                    COALESCE(SUM(fp.AmountPaid), 0) as AmountPaid,
                    ROUND(MAX(f.Amount * (st.PayRate / 100.0)) - COALESCE(SUM(fp.AmountPaid), 0), 2) as UnpaidAmount,
                    CASE 
                        WHEN COALESCE(SUM(fp.AmountPaid), 0) >= (f.Amount * (st.PayRate / 100.0)) THEN 'Paid'
                        WHEN COALESCE(SUM(fp.AmountPaid), 0) > 0 THEN 'Partial'
                        ELSE 'Unpaid'
                    END as PaymentStatus,
                    st.StudentTermID,
                    f.FeesID
                FROM StudentTerm st
                JOIN Students s ON st.StudentID = s.StudentID
                JOIN Class c ON st.ClassID = c.ClassID
                JOIN Term t ON st.TermID = t.TermID
                JOIN AcademicYear ay ON c.AcademicYear = ay.AcademicYearID
                JOIN Fees f ON f.Class = c.ClassID AND f.Term = t.TermID
                LEFT JOIN FeesPayment fp ON st.StudentTermID = fp.StudentTermID
                """
                
                filters = []
                params = []
                
                # Apply filters
                search_text = self.search_input.text().strip()
                if search_text:
                    filters.append("(s.Student_Name LIKE ? OR s.StudentID LIKE ?)")
                    params.extend([f"%{search_text}%", f"%{search_text}%"])
                
                selected_year = self.year_combo.currentData()
                if selected_year:
                    filters.append("ay.Year = ?")
                    params.append(selected_year)
                
                selected_term = self.term_combo.currentData()
                if selected_term:
                    filters.append("t.TermName = ?")
                    params.append(selected_term)
                
                selected_class = self.class_combo.currentData()
                if selected_class:
                    filters.append("c.ClassName = ?")
                    params.append(selected_class)
                
                if filters:
                    query += " WHERE " + " AND ".join(filters)
                
                query += """
                GROUP BY st.StudentTermID
                ORDER BY s.Student_Name, ay.Year, t.TermName
                """
                
                c.execute(query, params)
                rows = c.fetchall()
                
                self.table.setRowCount(0)
                for row_data in rows:
                    row = self.table.rowCount()
                    self.table.insertRow(row)
                    
                    # Add data columns
                    for col, data in enumerate(row_data[:7]):  # First 7 columns are data
                        item = QTableWidgetItem(str(data))
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        
                        # Color code status
                        if col == 6:  # Status column
                            if data == 'Paid':
                                item.setBackground(QColor(200, 255, 200))
                            elif data == 'Partial':
                                item.setBackground(QColor(255, 255, 200))
                            else:
                                item.setBackground(QColor(255, 200, 200))
                        
                        self.table.setItem(row, col, item)
                    
                    # Add Payment History button
                    student_term_id = row_data[7]  # StudentTermID
                    fees_id = row_data[8]  # FeesID
                    
                    history_btn = QPushButton("View History")
                    history_btn.setStyleSheet("""
                        QPushButton {
                            background-color: #4299E1;
                            color: white;
                            border: none;
                            padding: 5px;
                            border-radius: 3px;
                            font-size: 11px;
                        }
                        QPushButton:hover {
                            background-color: #3182CE;
                        }
                    """)
                    history_btn.clicked.connect(
                        lambda checked, stid=student_term_id, fid=fees_id: 
                        self.show_payment_history(stid, fid)
                    )
                    
                    self.table.setCellWidget(row, 7, history_btn)
                    self.table.setRowHeight(row,40)
                    

                
        except Exception as e:
            QMessageBox.warning(f"Error loading data: {e}")

    def show_payment_history(self, student_term_id, fees_id):
        """Show payment history dialog for a student"""
        try:
            with get_db_connection() as (conn, c):
                # Get student info
                c.execute("""
                    SELECT s.Student_Name, c.ClassName, t.TermName, ay.Year
                    FROM StudentTerm st
                    JOIN Students s ON st.StudentID = s.StudentID
                    JOIN Class c ON st.ClassID = c.ClassID
                    JOIN Term t ON st.TermID = t.TermID
                    JOIN AcademicYear ay ON c.AcademicYear = ay.AcademicYearID
                    WHERE st.StudentTermID = ?
                """, (student_term_id,))
                student_info = c.fetchone()
                
                # Get payment history
                c.execute("""
                    SELECT 
                        AmountPaid, 
                        TransactionDate, 
                        ReferenceNumber, 
                        DoneOn
                    FROM FeesPayment 
                    WHERE StudentTermID = ? AND FeesID = ?
                    ORDER BY TransactionDate DESC
                """, (student_term_id, fees_id))
                payments = c.fetchall()
                
                # Create dialog
                dialog = QDialog(self)
                dialog.setWindowTitle(f"Payment History - {student_info[0]}")
                dialog.setMinimumSize(600, 400)
                
                layout = QVBoxLayout()
                
                # Student info
                info_text = f"""
                <h3>{student_info[0]}</h3>
                <p>Class: {student_info[1]}</p>
                <p>Term: {student_info[2]} ({student_info[3]})</p>
                """
                info_label = QLabel(info_text)
                info_label.setTextFormat(Qt.RichText)
                layout.addWidget(info_label)
                
                # Payments table
                if payments:
                    table = QTableWidget()
                    table.setColumnCount(4)
                    table.setHorizontalHeaderLabels(["Amount Paid", "Transaction Date", "Reference", "Recorded On"])
                    table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
                    
                    table.setRowCount(len(payments))
                    for row_idx, payment in enumerate(payments):
                        for col_idx, data in enumerate(payment):
                            item = QTableWidgetItem(str(data))
                            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                            table.setItem(row_idx, col_idx, item)
                    
                    layout.addWidget(QLabel("Payment History:"))
                    layout.addWidget(table)
                else:
                    layout.addWidget(QLabel("No payment history found."))
                
                # Close button
                close_btn = QPushButton("Close")
                close_btn.clicked.connect(dialog.accept)
                layout.addWidget(close_btn)
                
                dialog.setLayout(layout)
                dialog.exec_()
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load payment history: {str(e)}")

        
    def get_academic_years(self):
        """Fetch list of academic years"""
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY Year DESC")
                return c.fetchall()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to fetch academic years: {str(e)}")
            return []

    def get_classes_for_year(self, academic_year_id):
        """Fetch classes for a specific academic year"""
        try:
            with get_db_connection() as (conn, c):
                c.execute("""
                    SELECT ClassID, ClassName 
                    FROM Class 
                    WHERE AcademicYear = ? 
                    ORDER BY ClassName
                """, (academic_year_id,))
                return c.fetchall()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to fetch classes: {str(e)}")
            return []

    def get_terms_for_year(self, academic_year_id):
        """Fetch terms for a specific academic year"""
        try:
            with get_db_connection() as (conn, c):
                c.execute("""
                    SELECT TermID, TermName 
                    FROM Term 
                    WHERE AcademicYear = ? 
                    ORDER BY TermName
                """, (academic_year_id,))
                return c.fetchall()
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to fetch terms: {str(e)}")
            return []

    # New method to show dialog for selecting criteria and generating template
    def generate_payment_template_dialog(self):
        """Show dialog to select academic year, class, term and generate spreadsheet"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Generate Payment Template")
        dialog.setFixedSize(400, 300)
        
        layout = QVBoxLayout(dialog)
        
        # Academic Year selector
        year_group = QGroupBox("Select Academic Year")
        year_layout = QFormLayout(year_group)
        year_combo = QComboBox()
        years = self.get_academic_years()
        if not years:
            QMessageBox.warning(self, "No Data", "No academic years found.")
            return
        for year_id, year_name in years:
            year_combo.addItem(year_name, year_id)
        year_layout.addRow("Academic Year:", year_combo)
        layout.addWidget(year_group)
        
        # Class selector (updates on year change)
        class_group = QGroupBox("Select Class")
        class_layout = QFormLayout(class_group)
        class_combo = QComboBox()
        class_layout.addRow("Class:", class_combo)
        layout.addWidget(class_group)
        
        # Term selector (updates on year change)
        term_group = QGroupBox("Select Term")
        term_layout = QFormLayout(term_group)
        term_combo = QComboBox()
        term_layout.addRow("Term:", term_combo)
        layout.addWidget(term_group)
        
        # Update classes and terms when year changes
        def update_selections():
            year_id = year_combo.currentData()
            # Update classes
            class_combo.clear()
            classes = self.get_classes_for_year(year_id)
            for class_id, class_name in classes:
                class_combo.addItem(class_name, class_id)
            # Update terms
            term_combo.clear()
            terms = self.get_terms_for_year(year_id)
            for term_id, term_name in terms:
                term_combo.addItem(term_name, term_id)
        
        year_combo.currentIndexChanged.connect(update_selections)
        update_selections()  # Initial load
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(lambda: self.generate_payment_template(
            year_combo.currentData(), class_combo.currentData(), term_combo.currentData(), dialog
        ))
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        dialog.exec_()

    def generate_payment_template(self, academic_year_id, class_id, term_id, dialog):
        """Generate XLSX template for bulk payments"""
        if not all([academic_year_id, class_id, term_id]):
            QMessageBox.warning(self, "Error", "Please select academic year, class, and term.")
            return
        
        try:
            with get_db_connection() as (conn, c):
                # Fetch fee for this class and term
                c.execute("""
                    SELECT FeesID, Amount 
                    FROM Fees 
                    WHERE Class = ? AND Term = ?
                """, (class_id, term_id))
                fee_data = c.fetchone()
                if not fee_data:
                    QMessageBox.warning(self, "Error", "No fee structure found for this class and term.")
                    return
                fees_id, original_amount = fee_data
                
                # Fetch all student enrollments for this class and term
                # Assuming StudentTerm has enrollments per class/term (which imply year via links)
                c.execute("""
                    SELECT st.StudentTermID, st.StudentID, s.Student_Name, st.PayRate,
                           c.ClassName, t.TermName, y.Year
                    FROM StudentTerm st
                    JOIN Students s ON st.StudentID = s.StudentID
                    JOIN Class c ON st.ClassID = c.ClassID
                    JOIN Term t ON st.TermID = t.TermID
                    JOIN AcademicYear y ON c.AcademicYear = y.AcademicYearID
                    WHERE st.ClassID = ? AND st.TermID = ? AND c.AcademicYear = ?
                """, (class_id, term_id, academic_year_id))
                enrollments = c.fetchall()
                
                if not enrollments:
                    QMessageBox.warning(self, "No Data", "No students enrolled in this class and term.")
                    return
                
                # Create workbook
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Bulk Payments"
                
                # Headers
                headers = [
                    "StudentTermID", "FeesID", "StudentID", "Student Name", "Class", "Term", "Academic Year",
                    "Pay Rate (%)", "Original Amount", "Discounted Amount", "Already Paid",
                    "Remaining", "Amount to Pay", "Transaction Date", "Reference Number"
                ]
                ws.append(headers)
                
                # Style headers
                header_font = Font(bold=True, color="FFFFFF")
                header_fill = PatternFill(start_color="4A5568", end_color="4A5568", fill_type="solid")
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                for cell in ws[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center")
                    cell.border = thin_border
                
                # Hide sensitive columns (StudentTermID, FeesID) by setting width to 0
                ws.column_dimensions['A'].width = 0  # StudentTermID
                ws.column_dimensions['B'].width = 0  # FeesID
                
                # Populate rows
                today_str = datetime.now().strftime("%Y-%m-%d")
                for row, enrollment in enumerate(enrollments, start=2):
                    student_term_id, student_id, student_name, pay_rate, class_name, term_name, year_name = enrollment
                    
                    # Calculate amounts (similar to dialog)
                    discounted_amount = original_amount * (pay_rate / 100)
                    
                    c.execute("""
                        SELECT COALESCE(SUM(AmountPaid), 0) 
                        FROM FeesPayment 
                        WHERE StudentTermID = ?
                    """, (student_term_id,))
                    already_paid = c.fetchone()[0]
                    
                    remaining = max(0, discounted_amount - already_paid)
                    
                    # Append row data
                    ws.append([
                        student_term_id, fees_id, student_id, student_name, class_name, term_name, year_name,
                        pay_rate, original_amount, discounted_amount, already_paid, remaining,
                        None, today_str, None  # Editable: Amount to Pay, Date, Ref
                    ])
                    
                    # Style read-only cells (gray background)
                    read_only_fill = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid")
                    for col in range(1, 13):  # A to L (up to Remaining)
                        ws.cell(row=row, column=col).fill = read_only_fill
                        ws.cell(row=row, column=col).border = thin_border
                    
                    # Editable cells (white, with borders)
                    for col in range(13, 16):  # M to O
                        ws.cell(row=row, column=col).border = thin_border
                
                # Auto-adjust visible column widths
                for col in ws.columns:
                    max_length = 0
                    column = col[0].column_letter
                    if column in ['A', 'B']: continue  # Hidden
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = (max_length + 2)
                    ws.column_dimensions[column].width = adjusted_width
                
                # Save file
                file_path = QFileDialog.getSaveFileName(self, "Save Template", "", "Excel Files (*.xlsx)")[0]
                if file_path:
                    wb.save(file_path)
                    QMessageBox.information(self, "Success", f"Template generated at {file_path}.\nEdit the 'Amount to Pay', 'Transaction Date', and 'Reference Number' columns, then import.")
                    dialog.accept()
        
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to generate template: {str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Unexpected error: {str(e)}")

    # New method to import the edited spreadsheet
    def import_payments_dialog(self):
        """Show dialog to select file and import payments"""
        file_path = QFileDialog.getOpenFileName(self, "Select Payment File", "", "Excel Files (*.xlsx)")[0]
        if not file_path:
            return
        
        self.import_payments(file_path)

    def import_payments(self, file_path):
        """Import and process payments from XLSX with enhanced real-time validations"""
        try:
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            
            # Expected headers (validate)
            expected_headers = [
                "StudentTermID", "FeesID", "StudentID", "Student Name", "Class", "Term", "Academic Year",
                "Pay Rate (%)", "Original Amount", "Discounted Amount", "Already Paid",
                "Remaining", "Amount to Pay", "Transaction Date", "Reference Number"
            ]
            actual_headers = [cell.value for cell in ws[1]]
            if actual_headers != expected_headers:
                QMessageBox.warning(self, "Error", "Invalid file format. Headers do not match expected template.")
                return
            
            # Process rows
            success_count = 0
            skipped_count = 0
            errors = []
            done_on = QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss")
            
            with get_db_connection() as (conn, c):
                for row in range(2, ws.max_row + 1):
                    try:
                        # Read row (assume minimal tampering, but re-validate from DB)
                        student_term_id = int(ws.cell(row=row, column=1).value)
                        fees_id = int(ws.cell(row=row, column=2).value)
                        amount_to_pay_str = ws.cell(row=row, column=13).value
                        transaction_date = ws.cell(row=row, column=14).value
                        reference_number = str(ws.cell(row=row, column=15).value or "")
                        
                        if amount_to_pay_str is None:
                            errors.append(f"Row {row}: No amount to pay specified.")
                            continue
                        
                        amount_paid = float(amount_to_pay_str)
                        
                        # Basic validations
                        if amount_paid < 0:
                            errors.append(f"Row {row}: Amount must not be less than zero.")
                            continue
                        if not transaction_date:
                            errors.append(f"Row {row}: Transaction date required.")
                            continue
                        if reference_number is None:
                            errors.append(f"Row {row}: Reference number required.")
                            continue
                        
                        # Real-time DB fetch and validations
                        # 1. Verify StudentTermID and FeesID exist and match
                        c.execute("""
                            SELECT st.PayRate, f.Amount
                            FROM StudentTerm st
                            JOIN Fees f ON f.FeesID = ?
                            WHERE st.StudentTermID = ?
                        """, (fees_id, student_term_id))
                        data = c.fetchone()
                        if not data:
                            errors.append(f"Row {row}: Invalid StudentTermID or FeesID.")
                            continue
                        pay_rate, original_amount = data
                        
                        # 2. Calculate discounted amount
                        discounted_amount = original_amount * (pay_rate / 100)
                        
                        # 3. Fetch current already paid
                        c.execute("""
                            SELECT COALESCE(SUM(AmountPaid), 0) 
                            FROM FeesPayment 
                            WHERE StudentTermID = ?
                        """, (student_term_id,))
                        already_paid = c.fetchone()[0]
                        
                        # 4. Calculate current remaining
                        current_remaining = max(0, discounted_amount - already_paid)
                        
                        # 5. Validate amount <= current_remaining (prevents overpayment)
                        if amount_paid > current_remaining:
                            errors.append(f"Row {row}: Amount ({amount_paid:,.2f}) exceeds current remaining ({current_remaining:,.2f}).")
                            continue
                        
                        # 6. Check for potential duplicate: See if exact same payment exists
                        # (same amount, date, reference for this StudentTermID)
                        c.execute("""
                            SELECT COUNT(*) 
                            FROM FeesPayment 
                            WHERE StudentTermID = ? 
                            AND AmountPaid = ? 
                            AND TransactionDate = ? 
                            AND ReferenceNumber = ?
                        """, (student_term_id, amount_paid, transaction_date, reference_number))
                        duplicate_count = c.fetchone()[0]
                        if duplicate_count > 0:
                            skipped_count += 1
                            errors.append(f"Row {row}: Skipped duplicate payment.")
                            continue
                        c.execute("SELECT COUNT(*) FROM FeesPayment WHERE ReferenceNumber = ?", (reference_number,))
                        ref_count = c.fetchone()[0]
                        if ref_count > 0:
                            errors.append(f"Row {row}: Reference number '{reference_number}' already used in another payment.")
                            continue
                        # If all good, insert payment
                        c.execute("""
                            INSERT INTO FeesPayment 
                            (StudentTermID, FeesID, AmountPaid, TransactionDate, ReferenceNumber, DoneOn)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (student_term_id, fees_id, amount_paid, transaction_date, reference_number, done_on))
                        
                        success_count += 1
                    
                    except ValueError:
                        errors.append(f"Row {row}: Invalid amount or data format.")
                    except sqlite3.Error as e:
                        errors.append(f"Row {row}: Database error - {str(e)}")
                
                conn.commit()
            
            # Show results
            message = f"Imported {success_count} payments successfully.\nSkipped {skipped_count} potential duplicates."
            if errors:
                message += "\nErrors/Warnings:\n" + "\n".join(errors[:7])  # Limit to first 10 for brevity
                if len(errors) > 10:
                    message += f"\n... and {len(errors) - 10} more errors."
            QMessageBox.information(self, "Import Result", message)
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to import file: {str(e)}")

class PaymentHistoryPage(QWidget):
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.load_data()
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Header
        header_layout = QHBoxLayout()
        title = QLabel("Payment History - Raw Data")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        
        # Export button
        export_btn = QPushButton("Export to Excel")
        export_btn.clicked.connect(self.export_to_excel)
        export_btn.setStyleSheet("""
            QPushButton {
                background-color: #4299E1;
                color: white;
                border: none;
                padding: 8px 15px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3182CE;
            }
        """)
        header_layout.addWidget(export_btn)
        
        # Delete selected button
        delete_selected_btn = QPushButton("Delete Selected")
        delete_selected_btn.clicked.connect(self.delete_selected_payments)
        delete_selected_btn.setStyleSheet("""
            QPushButton {
                background-color: #E53E3E;
                color: white;
                border: none;
                padding: 8px 15px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #C53030;
            }
        """)
        header_layout.addWidget(delete_selected_btn)
        self.table = QTableWidget()
        # Summary label
        self.summary_label = QLabel()
        self.summary_label.setStyleSheet("font-weight: bold; color: #2D3748;")

        layout.addLayout(header_layout)

        # Filters
        filter_layout = QHBoxLayout()
        
        # Search
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search student, reference...")
        self.search_input.setMaximumWidth(250)
        self.search_input.textChanged.connect(self.load_data)
        
        # Date range
        self.date_from = QDateEdit()
        self.date_from.setDate(QDate.currentDate().addMonths(-1))
        self.date_from.setCalendarPopup(True)
        self.date_from.setDisplayFormat("yyyy-MM-dd")
        self.date_from.dateChanged.connect(self.load_data)
        
        self.date_to = QDateEdit()
        self.date_to.setDate(QDate.currentDate())
        self.date_to.setCalendarPopup(True)
        self.date_to.setDisplayFormat("yyyy-MM-dd")
        self.date_to.dateChanged.connect(self.load_data)
        
        # Academic Year filter
        self.year_combo = QComboBox()
        self.year_combo.currentIndexChanged.connect(self.update_class_term_filters)
        self.year_combo.currentIndexChanged.connect(self.load_data)
        
        # Class filter
        self.class_combo = QComboBox()
        self.class_combo.addItem("All Classes", None)
        self.class_combo.currentIndexChanged.connect(self.load_data)
        
        # Term filter
        self.term_combo = QComboBox()
        self.term_combo.addItem("All Terms", None)
        self.term_combo.currentIndexChanged.connect(self.load_data)
        
        # Load filter options
        self.load_filter_options()
        
        filter_layout.addWidget(QLabel("Search:"))
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(QLabel("From:"))
        filter_layout.addWidget(self.date_from)
        filter_layout.addWidget(QLabel("To:"))
        filter_layout.addWidget(self.date_to)
        filter_layout.addWidget(QLabel("Year:"))
        filter_layout.addWidget(self.year_combo)
        filter_layout.addWidget(QLabel("Class:"))
        filter_layout.addWidget(self.class_combo)
        filter_layout.addWidget(QLabel("Term:"))
        filter_layout.addWidget(self.term_combo)
        
        layout.addLayout(filter_layout)

        # Table for raw payment data with Actions column
        self.table.setColumnCount(11)  # Added Actions column
        self.table.setHorizontalHeaderLabels([
            "Payment ID", "Student", "Student ID", "Class", "Term", 
            "Year", "Amount Paid", "Transaction Date", "Reference", "Recorded On", "Actions"
        ])
        
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.MultiSelection)  # Allow multiple selection
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        # Set column widths
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # Payment ID
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Student Name
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Student ID
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Class
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Term
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Year
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)  # Amount Paid
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)  # Transaction Date
        header.setSectionResizeMode(8, QHeaderView.ResizeToContents)  # Reference
        header.setSectionResizeMode(9, QHeaderView.ResizeToContents)  # Recorded On
        header.setSectionResizeMode(10, QHeaderView.ResizeToContents)  # Actions
        
        layout.addWidget(self.table)
        
        
        layout.addWidget(self.summary_label)

    def load_filter_options(self):
        """Load initial filter options"""
        try:
            with get_db_connection() as (conn, c):
                # Load academic years
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY Year DESC")
                for year_id, year_name in c.fetchall():
                    self.year_combo.addItem(year_name, year_id)
                
                # Initially load all classes and terms (when "All Years" is selected)
                self.update_class_term_filters()
                
        except Exception as e:
            QMessageBox.warning(f"Error loading filter options: {e}")

    def update_class_term_filters(self):
        """Update class and term dropdowns based on selected academic year"""
        selected_year = self.year_combo.currentData()
        
        # Update CLASS dropdown
        self.class_combo.clear()
        self.class_combo.addItem("All Classes", None)
        
        if selected_year:  # Only populate if a specific year is selected
            try:
                with get_db_connection() as (conn, c):
                    # Get classes for this academic year
                    c.execute("""
                        SELECT DISTINCT ClassID, ClassName 
                        FROM Class
                        WHERE AcademicYear = ?
                        ORDER BY ClassName
                    """, (selected_year,))
                    
                    for class_id, class_name in c.fetchall():
                        self.class_combo.addItem(class_name, class_id)
            except Exception as e:
                QMessageBox.warning(f"Error loading classes: {e}")
        
        # Update TERM dropdown  
        self.term_combo.clear()
        self.term_combo.addItem("All Terms", None)
        
        if selected_year:  # Only populate if a specific year is selected
            try:
                with get_db_connection() as (conn, c):
                    # Get terms for this academic year
                    c.execute("""
                        SELECT TermID, TermName 
                        FROM Term 
                        WHERE AcademicYear = ?
                        ORDER BY TermName
                    """, (selected_year,))
                    
                    for term_id, term_name in c.fetchall():
                        self.term_combo.addItem(term_name, term_id)
            except Exception as e:
                QMessageBox.warning(f"Error loading terms: {e}")

    def load_data(self):
        """Load raw payment history data with filters"""
        try:
            with get_db_connection() as (conn, c):
                query = """
                SELECT 
                    fp.FeesPaymentID,
                    s.Student_Name,
                    s.StudentID,
                    c.ClassName,
                    t.TermName,
                    ay.Year as AcademicYear,
                    fp.AmountPaid,
                    fp.TransactionDate,
                    fp.ReferenceNumber,
                    fp.DoneOn,
                    st.StudentTermID,
                    f.FeesID
                FROM FeesPayment fp
                JOIN StudentTerm st ON fp.StudentTermID = st.StudentTermID
                JOIN Students s ON st.StudentID = s.StudentID
                JOIN Class c ON st.ClassID = c.ClassID
                JOIN Term t ON st.TermID = t.TermID
                JOIN AcademicYear ay ON c.AcademicYear = ay.AcademicYearID
                JOIN Fees f ON f.Class= c.ClassID AND f.Term= t.TermID
                WHERE 1=1
                """
                
                filters = []
                params = []
                
                # Apply search filter
                search_text = self.search_input.text().strip()
                if search_text:
                    filters.append("""
                        (s.Student_Name LIKE ? OR s.StudentID LIKE ? 
                        OR fp.ReferenceNumber LIKE ? OR fp.FeesPaymentID LIKE ?)
                    """)
                    params.extend([f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%"])
                
                # Apply date filters
                date_from = self.date_from.date().toString("yyyy-MM-dd")
                date_to = self.date_to.date().toString("yyyy-MM-dd")
                filters.append("fp.TransactionDate BETWEEN ? AND ?")
                params.extend([date_from, date_to])
                
                # Apply academic year filter
                selected_year = self.year_combo.currentData()
                if selected_year:
                    filters.append("ay.AcademicYearID = ?")
                    params.append(selected_year)
                
                # Apply class filter
                selected_class = self.class_combo.currentData()
                if selected_class:
                    filters.append("c.ClassName = ?")
                    params.append(selected_class)
                
                # Apply term filter
                selected_term = self.term_combo.currentData()
                if selected_term:
                    filters.append("t.TermName = ?")
                    params.append(selected_term)
                
                if filters:
                    query += " AND " + " AND ".join(filters)
                
                query += " ORDER BY fp.TransactionDate DESC, fp.DoneOn DESC"
                
                c.execute(query, params)
                rows = c.fetchall()
                
                # Update table
                self.table.setRowCount(len(rows))
                total_amount = 0
                
                for row_idx, row_data in enumerate(rows):
                    total_amount += float(row_data[6] or 0)  # AmountPaid
                    
                    for col_idx in range(10):  # Only display first 10 columns (not StudentTermID and FeesID)
                        data = row_data[col_idx]
                        item = QTableWidgetItem(str(data) if data is not None else "")
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        
                        # Color code based on amount
                        if col_idx == 6:  # AmountPaid column
                            item.setText(f"{float(data or 0):,.2f}")
                            if float(data or 0) > 0:
                                item.setForeground(QColor(0, 100, 0))  # Dark green
                        
                        self.table.setItem(row_idx, col_idx, item)
                    
                    # Add Action ComboBox instead of buttons
                    self.add_action_combo(row_idx, row_data[0], row_data[10], row_data[11])
                    self.table.setRowHeight(row_idx, 40)    
                
                # Update summary
                self.summary_label.setText(
                    f"Total Records: {len(rows)} | "
                    f"Total Amount Paid: {total_amount:,.2f} | "
                    f"Date Range: {date_from} to {date_to}"
                )
                
        except Exception as e:
            QMessageBox.warning(f"Error loading payment history: {e}")
            QMessageBox.critical(self, "Error", f"Failed to load payment history: {str(e)}")

    def add_action_combo(self, row, payment_id, student_term_id, fees_id):
        """Add comboBox for actions instead of buttons"""
        action_combo = QComboBox()
        action_combo.addItems(["Select Action", "Edit", "Delete"])
        
        # Connect the comboBox using lambda to capture current values
        action_combo.currentIndexChanged.connect(
            lambda index, pid=payment_id, stid=student_term_id, fid=fees_id, cb=action_combo: 
            self.handle_payment_action(index, pid, stid, fid, cb)
        )
        
        self.table.setCellWidget(row, 10, action_combo)  # Actions column is column 10

    def handle_payment_action(self, index, payment_id, student_term_id, fees_id, combo_box):
        """Handle payment action selection from comboBox"""
        if index == 0:  # "Select Action"
            return
        
        action = combo_box.itemText(index)
        combo_box.setCurrentIndex(0)  # Reset to "Select Action"
        
        if action == "Edit":
            self.edit_payment(payment_id, student_term_id, fees_id)
        elif action == "Delete":
            self.delete_payment(payment_id)
    
    def edit_payment(self, payment_id, student_term_id, fees_id):
        """Open edit dialog for payment record"""
        try:
            with get_db_connection() as (conn, c):
                # Fetch payment details
                c.execute("""
                    SELECT 
                        fp.FeesPaymentID,
                        s.Student_Name,
                        c.ClassName,
                        t.TermName,
                        t.TermID,
                        f.Amount,
                        fp.AmountPaid,
                        fp.TransactionDate,
                        fp.ReferenceNumber,
                        st.StudentTermID
                    FROM FeesPayment fp
                    JOIN StudentTerm st ON fp.StudentTermID = st.StudentTermID
                    JOIN Students s ON st.StudentID = s.StudentID
                    JOIN Class c ON st.ClassID = c.ClassID
                    JOIN Term t ON st.TermID = t.TermID
                    JOIN Fees f ON f.FeesID = ?
                    WHERE fp.FeesPaymentID = ?
                """, (fees_id, payment_id))
                
                payment_data = c.fetchone()
                
                if payment_data:
                    dialog = QDialog(self)
                    dialog.setWindowTitle(f"Edit Payment - {payment_data[1]}")
                    dialog.setFixedSize(500, 350)
                    dialog.setStyleSheet("""
                        QDialog {
                            background-color: #F7FAFC;
                        }
                        QLabel {
                            color: #2D3748;
                            font-weight: 500;
                        }
                        QLineEdit, QDateEdit {
                            background-color: white;
                            border: 1px solid #CBD5E0;
                            border-radius: 4px;
                            padding: 8px;
                            font-size: 11px;
                        }
                        QLineEdit:focus, QDateEdit:focus {
                            border: 2px solid #4299E1;
                        }
                    """)
                    
                    layout = QVBoxLayout()
                    
                    # Student info (read-only)
                    info_layout = QHBoxLayout()
                    info_layout.addWidget(QLabel(f"<b>Student:</b> {payment_data[1]}"))
                    info_layout.addWidget(QLabel(f"<b>Class:</b> {payment_data[2]}"))
                    layout.addLayout(info_layout)
                    
                    # Form layout
                    form = QFormLayout()
                    form.setSpacing(10)
                    
                    # Term (read-only)
                    term_label = QLabel(payment_data[3])
                    form.addRow(QLabel("Term:"), term_label)
                    
                    # Original Fee (read-only)
                    fee_label = QLabel(f"{payment_data[5]:,.2f}")
                    form.addRow(QLabel("Original Fee:"), fee_label)
                    
                    # Amount Paid
                    amount_input = QLineEdit()
                    amount_input.setText(str(payment_data[6]))
                    amount_input.setValidator(QDoubleValidator(0, 99999999, 2))
                    form.addRow(QLabel("Amount Paid:"), amount_input)
                    
                    # Reference Number
                    ref_input = QLineEdit()
                    ref_input.setText(str(payment_data[8] or ""))
                    form.addRow(QLabel("Reference Number:"), ref_input)
                    
                    # Transaction Date
                    date_input = QDateEdit()
                    date_input.setDate(QDate.fromString(str(payment_data[7]), "yyyy-MM-dd"))
                    date_input.setCalendarPopup(True)
                    date_input.setDisplayFormat("yyyy-MM-dd")
                    form.addRow(QLabel("Transaction Date:"), date_input)
                    
                    layout.addLayout(form)
                    
                    # Buttons
                    button_layout = QHBoxLayout()
                    
                    save_btn = QPushButton("Save")
                    save_btn.setFixedSize(100, 35)
                    save_btn.setStyleSheet("""
                        QPushButton {
                            background-color: #48BB78;
                            color: white;
                            border: none;
                            border-radius: 4px;
                            font-weight: bold;
                        }
                        QPushButton:hover {
                            background-color: #38A169;
                        }
                    """)
                    save_btn.clicked.connect(
                        lambda: self.save_edited_payment(
                            dialog, payment_id, amount_input, ref_input, date_input
                        )
                    )
                    
                    cancel_btn = QPushButton("Cancel")
                    cancel_btn.setFixedSize(100, 35)
                    cancel_btn.setStyleSheet("""
                        QPushButton {
                            background-color: #A0AEC0;
                            color: white;
                            border: none;
                            border-radius: 4px;
                            font-weight: bold;
                        }
                        QPushButton:hover {
                            background-color: #718096;
                        }
                    """)
                    cancel_btn.clicked.connect(dialog.reject)
                    
                    button_layout.addStretch()
                    button_layout.addWidget(save_btn)
                    button_layout.addWidget(cancel_btn)
                    
                    layout.addLayout(button_layout)
                    dialog.setLayout(layout)
                    dialog.exec_()
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load payment for editing: {str(e)}")

    def save_edited_payment(self, dialog, payment_id, amount_input, ref_input, date_input):
        """Save edited payment details to database"""
        try:
            amount = float(amount_input.text())
            reference = ref_input.text().strip()
            transaction_date = date_input.date().toString("yyyy-MM-dd")
            
            # Validate inputs
            if amount < 0:
                QMessageBox.warning(self, "Invalid Input", "Amount paid cannot be negative.")
                return
            
            with get_db_connection() as (conn, c):
                c.execute("""
                    UPDATE FeesPayment
                    SET AmountPaid = ?, ReferenceNumber = ?, TransactionDate = ?
                    WHERE FeesPaymentID = ?
                """, (amount, reference, transaction_date, payment_id))
                
                conn.commit()
                
                QMessageBox.information(self, "Success", "Payment record updated successfully!")
                dialog.accept()
                self.load_data()  # Reload table
                
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid amount (numeric value).")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to update payment: {str(e)}")

    def delete_payment(self, payment_id):
        """Delete a single payment record with confirmation"""
        reply = QMessageBox.question(
            self, 
            "Confirm Delete", 
            "Are you sure you want to delete this payment record?\nThis action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    c.execute("DELETE FROM FeesPayment WHERE FeesPaymentID = ?", (payment_id,))
                    conn.commit()
                    
                    QMessageBox.information(self, "Success", "Payment record deleted successfully!")
                    self.load_data()  # Reload table
                    
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete payment: {str(e)}")

    def delete_selected_payments(self):
        """Delete multiple selected payment records with confirmation"""
        selected_rows = self.table.selectionModel().selectedRows()
        
        if not selected_rows:
            QMessageBox.warning(self, "No Selection", "Please select at least one payment record to delete.")
            return
        
        # Get payment IDs from selected rows
        payment_ids = []
        for row in selected_rows:
            payment_id = self.table.item(row.row(), 0).text()
            payment_ids.append(payment_id)
        
        reply = QMessageBox.question(
            self,
            "Confirm Delete Multiple Records",
            f"Are you sure you want to delete {len(payment_ids)} payment record(s)?\nThis action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    for payment_id in payment_ids:
                        c.execute("DELETE FROM FeesPayment WHERE FeesPaymentID = ?", (payment_id,))
                    
                    conn.commit()
                    
                    QMessageBox.information(
                        self, 
                        "Success", 
                        f"{len(payment_ids)} payment record(s) deleted successfully!"
                    )
                    self.load_data()  # Reload table
                    
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete payment records: {str(e)}")

    def export_to_excel(self):
        """Export payment history to Excel"""
        try:
            if self.table.rowCount() == 0:
                QMessageBox.warning(self, "No Data", "No payment history to export.")
                return
            
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save Payment History", 
                f"Payment_History_{QDate.currentDate().toString('yyyyMMdd')}.xlsx",
                "Excel Files (*.xlsx)"
            )
            
            if file_path:
                # Collect data from table
                data = []
                headers = []
                
                # Get headers (exclude Actions column)
                for col in range(self.table.columnCount() - 1):  # Exclude Actions column
                    headers.append(self.table.horizontalHeaderItem(col).text())
                
                # Get data
                for row in range(self.table.rowCount()):
                    row_data = []
                    for col in range(self.table.columnCount() - 1):  # Exclude Actions column
                        item = self.table.item(row, col)
                        row_data.append(item.text() if item else "")
                    data.append(row_data)
                
                # Create DataFrame and save
                df = pd.DataFrame(data, columns=headers)
                df.to_excel(file_path, index=False)
                
                QMessageBox.information(self, "Success", 
                                      f"Payment history exported successfully!\n"
                                      f"Saved to: {file_path}")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export: {str(e)}")

    def show_context_menu(self, position):
        """Show context menu for table actions"""
        menu = QMenu()
        
        view_details = menu.addAction("View Payment Details")
        copy_row = menu.addAction("Copy Row Data")
        menu.addSeparator()
        refresh = menu.addAction("Refresh Data")
        
        action = menu.exec_(self.table.viewport().mapToGlobal(position))
        
        if action == view_details:
            self.view_payment_details()
        elif action == copy_row:
            self.copy_row_data()
        elif action == refresh:
            self.load_data()

    def view_payment_details(self):
        """View details of selected payment"""
        selected_row = self.table.currentRow()
        if selected_row >= 0:
            payment_id = self.table.item(selected_row, 0).text()
            student_name = self.table.item(selected_row, 1).text()
            
            # Show detailed dialog
            self.show_payment_detail_dialog(payment_id, student_name)

    def copy_row_data(self):
        """Copy selected row data to clipboard"""
        selected_row = self.table.currentRow()
        if selected_row >= 0:
            row_data = []
            for col in range(self.table.columnCount() - 1):  # Exclude Actions column
                item = self.table.item(selected_row, col)
                row_data.append(item.text() if item else "")
            
            clipboard = QApplication.clipboard()
            clipboard.setText("\t".join(row_data))
            
            QMessageBox.information(self, "Copied", "Row data copied to clipboard.")

    def show_payment_detail_dialog(self, payment_id, student_name):
        """Show detailed payment information"""
        try:
            with get_db_connection() as (conn, c):
                c.execute("""
                SELECT 
                    fp.*,
                    s.Student_Name,
                    c.ClassName,
                    t.TermName,
                    ay.Year,
                    f.Amount as OriginalFee,
                    st.PayRate
                FROM FeesPayment fp
                JOIN StudentTerm st ON fp.StudentTermID = st.StudentTermID
                JOIN Students s ON st.StudentID = s.StudentID
                JOIN Class c ON st.ClassID = c.ClassID
                JOIN Term t ON st.TermID = t.TermID
                JOIN AcademicYear ay ON c.AcademicYear = ay.AcademicYearID
                JOIN Fees f ON f.Class= c.ClassID AND f.Term= t.TermID
                WHERE fp.FeesPaymentID = ?
                """, (payment_id,))
                
                payment_data = c.fetchone()
                
                if payment_data:
                    dialog = QDialog(self)
                    dialog.setWindowTitle(f"Payment Details - {student_name}")
                    dialog.setFixedSize(500, 400)
                    
                    layout = QVBoxLayout()
                    
                    # Payment details in form layout
                    form = QFormLayout()
                    
                    details = [
                        ("Payment ID", payment_data[0]),
                        ("Student", payment_data[11]),
                        ("Class", payment_data[12]),
                        ("Term", f"{payment_data[13]} ({payment_data[14]})"),
                        ("Original Fee", f"{payment_data[15]:,.2f}"),
                        ("Pay Rate", f"{payment_data[16]}%"),
                        ("Amount Paid", f"{payment_data[3]:,.2f}"),
                        ("Transaction Date", payment_data[4]),
                        ("Reference", payment_data[5] or "N/A"),
                        ("Recorded On", payment_data[6])
                    ]
                    
                    for label, value in details:
                        value_label = QLabel(str(value))
                        form.addRow(QLabel(f"{label}:"), value_label)
                    
                    layout.addLayout(form)
                    
                    # Close button
                    close_btn = QPushButton("Close")
                    close_btn.clicked.connect(dialog.accept)
                    layout.addWidget(close_btn)
                    
                    dialog.setLayout(layout)
                    dialog.exec_()
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load payment details: {str(e)}")

#### End Of Fees Management ######
#### Start Of Uniforms #####
class Uniforms(QTabWidget):
    def __init__(self):
        super().__init__()
        self.Dashboard = UniformDashBoard()
        self.Inventory = InventoryManagement()
        self.Report = ReportGenerator()
        self.addTab(self.Dashboard,"Dashboard & Insights")
        self.addTab(self.Inventory,"Inventory")
        self.addTab(self.Report,"Reports")

class UniformDialog(QDialog):
    def __init__(self, parent=None, uniform_id=None):
        super().__init__(parent)
        self.uniform_id = uniform_id
        self.setWindowTitle("Add New Uniform" if uniform_id is None else "Edit Uniform")
        self.setModal(True)
        
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()
        
        self.name_edit = QLineEdit()
        self.price_edit = QLineEdit()
        
        form_layout.addRow("Uniform Name:", self.name_edit)
        form_layout.addRow("Price:", self.price_edit)
        
        layout.addLayout(form_layout)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        # If editing, load existing data
        if uniform_id is not None:
            self.load_uniform_data()
    
    def load_uniform_data(self):
        """Load uniform data for editing"""
        try:
            with get_db_connection() as (conn, c):
                
                c.execute("SELECT * FROM Uniform WHERE UniformId = ?", (self.uniform_id,))
                uniform = c.fetchone()
                
                if uniform:
                    self.name_edit.setText(uniform['UniformName'])
                    self.price_edit.setText(str(uniform['UniformPrice']))
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load uniform data: {str(e)}")
    
    def validate_and_save(self):
        """Validate input and save uniform data"""
        name = self.name_edit.text().strip()
        price_str = self.price_edit.text().strip()
        
        # Validation
        if not all([name, price_str]):
            QMessageBox.warning(self, "Error", "All fields are required")
            return
        
        try:
            price = float(price_str)
            if price <= 0:
                QMessageBox.warning(self, "Error", "Price must be positive")
                return
        except ValueError:
            QMessageBox.warning(self, "Error", "Invalid price")
            return
        
        # Save to database
        try:
            with get_db_connection() as (conn,c):
                
                # Check for duplicate name (excluding current uniform if editing)
                if self.uniform_id is None:
                    c.execute("SELECT COUNT(*) FROM Uniform WHERE UniformName = ?", (name,))
                    if c.fetchone()[0] > 0:
                        QMessageBox.warning(self, "Error", "A uniform with this name already exists")
                        return
                else:
                    c.execute("SELECT COUNT(*) FROM Uniform WHERE UniformName = ? AND UniformId != ?", 
                             (name, self.uniform_id))
                    if c.fetchone()[0] > 0:
                        QMessageBox.warning(self, "Error", "Another uniform with this name already exists")
                        return
                
                if self.uniform_id is None:
                    # Insert new record
                    c.execute("INSERT INTO Uniform (UniformName, UniformPrice, Quantity) VALUES (?, ?, 0)", 
                             (name, price))
                else:
                    # Update existing record
                    c.execute("UPDATE Uniform SET UniformName = ?, UniformPrice = ? WHERE UniformId = ?", 
                             (name, price, self.uniform_id))
                
                conn.commit()
                self.accept()
                
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to save uniform: {str(e)}")

class UniformDashBoard(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.load_data()

    def init_ui(self):
        # Main Layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        # Scroll Area for better responsiveness
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setSpacing(20)
        scroll.setWidget(content_widget)

        # --- Header Section ---
        header_layout = QHBoxLayout()
        title_label = QLabel("Uniform Dashboard & Insights")
        title_label.setObjectName("pageTitle")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        
        refresh_btn = QPushButton("Refresh")
        refresh_btn.setObjectName("infoButton")
        refresh_btn.clicked.connect(self.load_data)
        header_layout.addWidget(refresh_btn)
        
        content_layout.addLayout(header_layout)

        # --- Summary Cards Section ---
        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(15)
        
        self.card_total_uniforms = self.create_summary_card("Total Uniforms", "0", "#3182CE")
        self.card_total_quantity = self.create_summary_card("Total Quantity", "0", "#38A169")
        self.card_low_stock = self.create_summary_card("Low Stock Items", "0", "#DD6B20")
        self.card_recent_entries = self.create_summary_card("Recent Entries", "0", "#805AD5")
        
        cards_layout.addWidget(self.card_total_uniforms)
        cards_layout.addWidget(self.card_total_quantity)
        cards_layout.addWidget(self.card_low_stock)
        cards_layout.addWidget(self.card_recent_entries)
        
        content_layout.addLayout(cards_layout)
        
         # --- Uniform List Section (Full Width CRUD) ---
        uniform_frame = QFrame()
        uniform_frame.setObjectName("summaryCard")
        uniform_frame.setMinimumHeight(300)
        uniform_layout = QVBoxLayout(uniform_frame)
        
        uni_header = QHBoxLayout()
        uni_title = QLabel("Uniform Catalog")
        uni_title.setObjectName("sectionTitle")
        uni_header.addWidget(uni_title)
        uni_header.addStretch()
        
        # Search
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search uniforms...")
        self.search_input.setFixedWidth(250)
        self.search_input.textChanged.connect(self.load_uniform_list)
        uni_header.addWidget(self.search_input)
        
        # Add button in header
        add_uniform_btn = QPushButton("+ Add New Uniform")
        add_uniform_btn.setObjectName("successButton")
        add_uniform_btn.clicked.connect(self.show_add_uniform_dialog)
        uni_header.addWidget(add_uniform_btn)
        
        uniform_layout.addLayout(uni_header)
        
        self.uniform_table = QTableWidget()
        self.uniform_table.setColumnCount(5)
        self.uniform_table.setHorizontalHeaderLabels(["ID", "Name", "Price", "Quantity", "Actions"])
        self.uniform_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.uniform_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.uniform_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.uniform_table.setShowGrid(True)
        self.uniform_table.verticalHeader().setVisible(False)
        uniform_layout.addWidget(self.uniform_table)

        content_layout.addWidget(uniform_frame)

        # --- Charts Section (Insights) ---
        charts_frame = QFrame()
        charts_frame.setObjectName("summaryCard")
        charts_layout = QHBoxLayout(charts_frame)
        
        self.in_canvas = self.create_pie_canvas("Stock In")
        self.out_canvas = self.create_pie_canvas("Stock Out")
        
        charts_layout.addWidget(self.in_canvas)
        charts_layout.addWidget(self.out_canvas)
        
        content_layout.addWidget(charts_frame)

        # --- Recent Inventory Section ---
        inventory_frame = QFrame()
        inventory_frame.setObjectName("summaryCard")
        inventory_frame.setMinimumHeight(300)
        inventory_layout = QVBoxLayout(inventory_frame)
        
        inv_header = QHBoxLayout()
        inv_title = QLabel("Recent Inventory Entries")
        inv_title.setObjectName("sectionTitle")
        inv_header.addWidget(inv_title)
        inv_header.addStretch()
        inventory_layout.addLayout(inv_header)

        self.inventory_table = QTableWidget()
        self.inventory_table.setColumnCount(6)
        self.inventory_table.setHorizontalHeaderLabels(["ID", "Uniform", "Type", "Quantity", "Date", "Reference"])
        self.inventory_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.inventory_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.inventory_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.inventory_table.setShowGrid(True)
        self.inventory_table.verticalHeader().setVisible(False)
        self.inventory_table.setMaximumHeight(250)
        inventory_layout.addWidget(self.inventory_table)
        
        content_layout.addWidget(inventory_frame)

       
        
        main_layout.addWidget(scroll)

    def create_pie_canvas(self, title):
        fig = Figure(figsize=(5, 4), dpi=100)
        canvas = FigureCanvas(fig)
        canvas.setMinimumHeight(250)
        return canvas

    def create_summary_card(self, title, value, color):
        """Create a summary card similar to Dashboard pattern"""
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 8px;
                border-left: 5px solid {color};
                border-top: 1px solid #E2E8F0;
                border-right: 1px solid #E2E8F0;
                border-bottom: 1px solid #E2E8F0;
            }}
        """)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(15, 15, 15, 15)
        
        title_lbl = QLabel(title)
        title_lbl.setStyleSheet("color: #718096; font-size: 14px; border: none;")
        
        value_lbl = QLabel(value)
        value_lbl.setStyleSheet("color: #2D3748; font-size: 28px; font-weight: bold; border: none;")
        
        layout.addWidget(title_lbl)
        layout.addWidget(value_lbl)
        layout.addStretch()
        
        # Store reference to update later
        card.value_label = value_lbl
        
        return card

    def update_pie_chart(self, canvas, title, data, color):
        fig = canvas.figure
        fig.clear()
        
        if not data:
            # Create a single slice with "No Data" label
            sizes = [1]
            labels = ['No Data']
            colors = ['#E0E0E0']  # Light gray for no data
            explode = [0]  # No explosion
            
            ax = fig.add_subplot(111)
            wedges, texts, autotexts = ax.pie(sizes, explode=explode, labels=labels, colors=colors,
                                            autopct='', shadow=False, startangle=90)
            
            ax.set_title(title, fontsize=12, fontweight='bold', pad=10, color='#2D3748')
            ax.axis('equal')
            
        else:
            # Normal pie chart with data
            labels = [row[0] for row in data]
            sizes = [row[1] for row in data]
            
            # Create a color palette based on the chart type
            if color == '#2E8B57':  # Stock In - Green shades
                colors = ['#2E8B57', '#3CB371', '#66CDAA', '#98FB98', '#90EE90']
            else:  # Stock Out - Red shades
                colors = ['#CD5C5C', '#DC143C', '#FF6347', '#FF7F50', '#F08080']
            
            # Use available colors cyclically if we have more items than colors
            if len(sizes) > len(colors):
                colors = colors * (len(sizes) // len(colors) + 1)
            colors = colors[:len(sizes)]
            
            ax = fig.add_subplot(111)
            wedges, texts, autotexts = ax.pie(sizes, labels=labels, colors=colors,
                                            autopct='%1.1f%%', shadow=True, startangle=90)
            
            # Style the text
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
                autotext.set_fontsize(9)
            
            for text in texts:
                text.set_fontsize(9)
            
            ax.set_title(title, fontsize=12, fontweight='bold', pad=10, color='#2D3748')
            ax.axis('equal')
        
        fig.tight_layout()
        canvas.draw()

    def load_data(self):
        self.load_charts_and_inventory()
        self.load_uniform_list()

    def load_charts_and_inventory(self):
        # Dynamic date ranges: Current Month -> 3 Months -> 6 Months -> 1 Year
        date_ranges = [
            (1, "Current Month"),
            (3, "Last 3 Months"),
            (6, "Last 6 Months"),
            (12, "Last Year")
        ]
        
        try:
            with get_db_connection() as (conn, c):
                
                stock_in_data = []
                stock_out_data = []
                found_range_label = "Current Month"
                
                # Loop through date ranges to find data
                for months, label in date_ranges:
                    # Calculate date threshold
                    date_threshold = QDate.currentDate().addMonths(-months).toString("yyyy-MM-dd")
                    
                    # Get stock in
                    c.execute("""
                        SELECT u.UniformName, SUM(ui.Quantity) as TotalQuantity
                        FROM UniformInventory ui
                        JOIN Uniform u ON ui.UniformID = u.UniformId
                        WHERE ui.Type = 'In' AND ui.Date >= ?
                        GROUP BY u.UniformId
                        ORDER BY TotalQuantity DESC
                    """, (date_threshold,))
                    stock_in_data = c.fetchall()
                    
                    # Get stock out
                    c.execute("""
                        SELECT u.UniformName, SUM(ui.Quantity) as TotalQuantity
                        FROM UniformInventory ui
                        JOIN Uniform u ON ui.UniformID = u.UniformId
                        WHERE ui.Type = 'Out' AND ui.Date >= ?
                        GROUP BY u.UniformId
                        ORDER BY TotalQuantity DESC
                    """, (date_threshold,))
                    stock_out_data = c.fetchall()
                    
                    if stock_in_data or stock_out_data:
                        found_range_label = label
                        break # Stop if we found data
                
                # Update charts
                self.update_pie_chart(self.in_canvas, f"Stock In ({found_range_label})", 
                                    stock_in_data, '#2E8B57')
                
                self.update_pie_chart(self.out_canvas, f"Stock Out ({found_range_label})", 
                                    stock_out_data, '#CD5C5C')
                
                # Recent inventory entries (Always show last 15 regardless of date range for better context)
                c.execute("""
                    SELECT ui.UniformInventoryID, u.UniformName, ui.Type, ui.Quantity, ui.Date, ui.ReferencesNotes
                    FROM UniformInventory ui
                    LEFT JOIN Uniform u ON ui.UniformID = u.UniformId
                    ORDER BY ui.Date DESC, ui.UniformInventoryID DESC
                    LIMIT 15
                """)
                inventory_rows = c.fetchall()
                self.inventory_table.setRowCount(len(inventory_rows))
                for i, row in enumerate(inventory_rows):
                    for j in range(len(row)):
                        self.inventory_table.setItem(i, j, QTableWidgetItem(str(row[j])))
                
                # --- Update Summary Cards ---
                # Total Uniforms
                c.execute("SELECT COUNT(*) FROM Uniform")
                total_uniforms = c.fetchone()[0]
                self.card_total_uniforms.value_label.setText(str(total_uniforms))
                
                # Total Quantity
                c.execute("SELECT SUM(Quantity) FROM Uniform")
                total_qty = c.fetchone()[0] or 0
                self.card_total_quantity.value_label.setText(str(total_qty))
                
                # Low Stock Items (quantity <= 10)
                c.execute("SELECT COUNT(*) FROM Uniform WHERE Quantity <= 10")
                low_stock = c.fetchone()[0]
                self.card_low_stock.value_label.setText(str(low_stock))
                
                # Recent Entries (entries from last 30 days)
                c.execute("""
                    SELECT COUNT(*) FROM UniformInventory 
                    WHERE Date >= datetime('now', '-30 days')
                """)
                recent_entries = c.fetchone()[0]
                self.card_recent_entries.value_label.setText(str(recent_entries))
                    
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load dashboard data: {str(e)}")

    def load_uniform_list(self):
        search_text = self.search_input.text().strip()
        try:
            with get_db_connection() as (conn, c):
                # Uniform list with search filter
                if search_text:
                    c.execute("""
                        SELECT UniformId, UniformName, UniformPrice, Quantity 
                        FROM Uniform 
                        WHERE UniformName LIKE ?
                        ORDER BY UniformName
                    """, (f"%{search_text}%",))
                else:
                    c.execute("SELECT UniformId, UniformName, UniformPrice, Quantity FROM Uniform ORDER BY UniformName")
                
                uniform_rows = c.fetchall()
                self.uniform_table.setRowCount(len(uniform_rows))
                for i, row in enumerate(uniform_rows):
                    for j in range(4):  # ID, Name, Price, Quantity
                        self.uniform_table.setItem(i, j, QTableWidgetItem(str(row[j])))
                    
                    # Add action buttons
                    self.add_uniform_action_buttons(i, row[0])
        except Exception as e:
            QMessageBox.warning(f"Error loading uniform list: {e}")

    def generate_colors(self, n, base_color):
        """Generate a color palette based on a base color"""
        # For simplicity, using matplotlib's default color cycle
        import matplotlib.pyplot as plt
        return plt.cm.Set3(range(n))

    def add_uniform_action_buttons(self, row, uniform_id):
        # Create action ComboBox
        action_combo = QComboBox()
        action_combo.addItems(["Select Action", "Edit", "Delete"])
        
        # Direct lambda approach
        action_combo.currentIndexChanged.connect(
            lambda index, uid=uniform_id, cb=action_combo: 
            self.on_uniform_action_selected(index, uid, cb)
        )
        
        self.uniform_table.setCellWidget(row, 4, action_combo)

    def on_uniform_action_selected(self, index, uniform_id, combo_box):
        """Handle uniform action selection"""
        if index == 0:  # "Select Action"
            return
        
        action = combo_box.itemText(index)
        combo_box.setCurrentIndex(0)  # Reset to "Select Action"
        
        if action == "Edit":
            self.edit_uniform(uniform_id)
        elif action == "Delete":
            self.delete_uniform(uniform_id)
   
    def show_add_uniform_dialog(self):
        """Show dialog to add new uniform"""
        dialog = UniformDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.load_uniform_list()

    def edit_uniform(self, uniform_id):
        """Edit existing uniform"""
        dialog = UniformDialog(self, uniform_id)
        if dialog.exec_() == QDialog.Accepted:
            self.load_uniform_list()

    def delete_uniform(self, uniform_id):
        """Delete uniform after checking for inventory constraints"""
        try:
            with get_db_connection() as (conn, c):
                
                # Check if uniform has inventory entries
                c.execute("SELECT COUNT(*) FROM UniformInventory WHERE UniformID = ?", (uniform_id,))
                inventory_count = c.fetchone()[0]
                
                if inventory_count > 0:
                    QMessageBox.warning(self, "Cannot Delete", 
                                      f"This uniform has {inventory_count} inventory entries and cannot be deleted.\n\n"
                                      "Please delete or reassign the inventory entries first.")
                    return
                
                # Confirm deletion
                reply = QMessageBox.question(
                    self, 'Confirm Delete', 
                    f'Are you sure you want to delete uniform #{uniform_id}?', 
                    QMessageBox.Yes | QMessageBox.No
                )
                
                if reply == QMessageBox.Yes:
                    c.execute("DELETE FROM Uniform WHERE UniformId = ?", (uniform_id,))
                    conn.commit()
                    QMessageBox.information(self, "Success", "Uniform deleted successfully")
                    self.load_uniform_list()
                    
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to delete uniform: {str(e)}")

class InventoryManagement(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # First row: Filters
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search...")
        self.search_input.setMaximumWidth(200)
        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDate(QDate.currentDate().addMonths(-1))  # Default to last month
        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDate(QDate.currentDate())
        self.type_combo = QComboBox()
        self.type_combo.addItems(["All Types", "In", "Out"])
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(QLabel("From Date:"))
        filter_layout.addWidget(self.from_date)
        filter_layout.addWidget(QLabel("To Date:"))
        filter_layout.addWidget(self.to_date)
        filter_layout.addWidget(QLabel("Type:"))
        filter_layout.addWidget(self.type_combo)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        # Second row: Results, metrics, buttons
        metrics_layout = QHBoxLayout()
        self.results_label = QLabel("0 results retrieved")
        self.results_label.setStyleSheet("font-size: 14px; color: #718096;")
        self.high_in_label = QLabel("High Qty In: 0")
        self.high_in_label.setStyleSheet("font-size: 14px; color: #00C000;")
        self.high_out_label = QLabel("High Qty Out: 0")
        self.high_out_label.setStyleSheet("font-size: 14px; color: #C00000;")
        print_btn = QPushButton("Print")
        print_btn.clicked.connect(self.print_table)
        new_inv_btn = QPushButton("New Inventory")
        new_inv_btn.clicked.connect(self.show_add_inventory_dialog)
        metrics_layout.addWidget(self.results_label)
        metrics_layout.addWidget(self.high_in_label)
        metrics_layout.addWidget(self.high_out_label)
        metrics_layout.addStretch()
        metrics_layout.addWidget(new_inv_btn)
        layout.addLayout(metrics_layout)

        # Third row: Table
        self.table = QTableWidget()
        self.table.setColumnCount(8)  # +1 for Actions
        self.table.setHorizontalHeaderLabels(["ID", "Uniform", "Type", "Quantity", "Price", "Date", "References", "Actions"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # Connect filters to load data
        self.search_input.textChanged.connect(self.load_data)
        self.from_date.dateChanged.connect(self.load_data)
        self.to_date.dateChanged.connect(self.load_data)
        self.type_combo.currentIndexChanged.connect(self.load_data)

        self.load_data()  # Initial load

    def load_data(self):
        search_text = self.search_input.text().strip()
        from_date_str = self.from_date.date().toString("yyyy-MM-dd")
        to_date_str = self.to_date.date().toString("yyyy-MM-dd")
        inv_type = self.type_combo.currentText()
        if inv_type == "All Types":
            inv_type = None
        
        try:
            with get_db_connection() as (conn, c):
                query = """
                    SELECT ui.UniformInventoryID, u.UniformName, ui.Type, ui.Quantity, ui.Price, ui.Date, ui.ReferencesNotes
                    FROM UniformInventory ui
                    LEFT JOIN Uniform u ON ui.UniformID = u.UniformId
                    WHERE 1=1
                """
                params = []
                
                if search_text:
                    query += " AND (u.UniformName LIKE ? OR ui.ReferencesNotes LIKE ?)"
                    params.extend([f"%{search_text}%", f"%{search_text}%"])
                
                if from_date_str:
                    query += " AND ui.Date >= ?"
                    params.append(from_date_str)
                
                if to_date_str:
                    query += " AND ui.Date <= ?"
                    params.append(to_date_str)
                
                if inv_type:
                    query += " AND ui.Type = ?"
                    params.append(inv_type)
                
                query += " ORDER BY ui.Date DESC"
                
                c.execute(query, params)
                rows = c.fetchall()
                
                self.table.setRowCount(len(rows))
                high_in = 0
                high_out = 0
                for i, row in enumerate(rows):
                    for j, val in enumerate(row):
                        self.table.setItem(i, j, QTableWidgetItem(str(val)))
                    # Calculate high qty
                    qty = row[3]
                    if row[2] == "In" and qty > high_in:
                        high_in = qty
                    elif row[2] == "Out" and qty > high_out:
                        high_out = qty
                    # Add actions COMBOBOX instead of buttons
                    self.add_action_combo(i, row[0])
                    self.table.setRowHeight(i, 40)
                
                self.results_label.setText(f"{len(rows)} results retrieved")
                self.high_in_label.setText(f"High Qty In: {high_in}")
                self.high_out_label.setText(f"High Qty Out: {high_out}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")

    def add_action_combo(self, row, inv_id):
        """Add comboBox for actions instead of buttons"""
        action_combo = QComboBox()
        action_combo.addItems(["Select Action", "Edit", "Delete"])
        
        # Style the comboBox to match your existing design
        action_combo.setFixedWidth(100)
        action_combo.setStyleSheet("""
            QComboBox {
                padding: 3px;
                border: 1px solid #d1d5db;
                border-radius: 4px;
                background-color: white;
            }
            QComboBox:hover {
                border-color: #9ca3af;
            }
            QComboBox::drop-down {
                border: none;
            }
        """)
        
        # Connect the comboBox using lambda to capture inv_id
        action_combo.currentIndexChanged.connect(
            lambda index, iid=inv_id, cb=action_combo: 
            self.handle_inventory_action(index, iid, cb)
        )
        
        self.table.setCellWidget(row, 7, action_combo)

    def handle_inventory_action(self, index, inv_id, combo_box):
        """Handle inventory action selection from comboBox"""
        if index == 0:  # "Select Action"
            return
        
        action = combo_box.itemText(index)
        combo_box.setCurrentIndex(0)  # Reset to "Select Action"
        
        if action == "Edit":
            self.update_inventory(inv_id)  # Your existing edit method
        elif action == "Delete":
            self.delete_inventory(inv_id)  # Your existing delete method
    
    def print_table(self):
        # Placeholder for print functionality
        QMessageBox.information(self, "Print", "Printing table... (placeholder)")

    def show_add_inventory_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add New Inventory")
        dialog.setMinimumWidth(400)
        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        
        # Uniform selection
        uniform_combo = QComboBox()
        uniform_combo.addItem("Select Uniform", None)
        
        # Type selection
        type_combo = QComboBox()
        type_combo.addItems(["In", "Out"])
        
        # Quantity input
        quantity_edit = QLineEdit()
        quantity_edit.setValidator(QIntValidator(1, 1000000))
        
        # Price input - will be dynamic based on type
        price_edit = QLineEdit()
        price_edit.setValidator(QDoubleValidator(0, 1000000, 2))
        
        # Date selection
        date_edit = QDateEdit()
        date_edit.setCalendarPopup(True)
        date_edit.setDate(QDate.currentDate())
        date_edit.setDisplayFormat("yyyy-MM-dd")
        
        # Reference input
        ref_edit = QLineEdit()
        ref_edit.setPlaceholderText("Optional reference notes")
        
        form_layout.addRow("Uniform:", uniform_combo)
        form_layout.addRow("Type:", type_combo)
        form_layout.addRow("Quantity:", quantity_edit)
        form_layout.addRow("Price:", price_edit)
        form_layout.addRow("Date:", date_edit)
        form_layout.addRow("References:", ref_edit)
        
        layout.addLayout(form_layout)
        
        # Buttons
        buttons = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_inventory(dialog, uniform_combo, type_combo, quantity_edit, price_edit, date_edit, ref_edit))
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        buttons.addWidget(save_btn)
        buttons.addWidget(close_btn)
        layout.addLayout(buttons)
        
        dialog.setLayout(layout)
        
        # Populate uniform combo
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT UniformId, UniformName FROM Uniform ORDER BY UniformName")
                for uniform in c.fetchall():
                    uniform_combo.addItem(uniform[1], uniform[0])
        except Exception as e:
            QMessageBox.warning(dialog, "Error", f"Failed to load uniforms: {str(e)}")
        
        # Connect signals for dynamic behavior
        def update_dynamic_fields():
            self.update_quantity_placeholder(uniform_combo, type_combo, quantity_edit)
            self.update_price_field(uniform_combo, type_combo, price_edit)
        
        uniform_combo.currentIndexChanged.connect(update_dynamic_fields)
        type_combo.currentIndexChanged.connect(update_dynamic_fields)
        
        # Set initial state
        update_dynamic_fields()
        
        dialog.exec_()

    def update_quantity_placeholder(self, uniform_combo, type_combo, quantity_edit):
        """Update quantity placeholder based on uniform and type selection"""
        uniform_id = uniform_combo.currentData()
        inv_type = type_combo.currentText()
        
        if uniform_id and inv_type == "Out":
            try:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT Quantity FROM Uniform WHERE UniformId = ?", (uniform_id,))
                    result = c.fetchone()
                    if result:
                        available_qty = result[0]
                        quantity_edit.setPlaceholderText(f"Max: {available_qty} (available stock)")
                        # Set tooltip with more information
                        quantity_edit.setToolTip(f"Current available quantity: {available_qty}. You cannot issue more than available stock.")
                        return
            except Exception as e:
                QMessageBox.warning(f"Error fetching available quantity: {e}")
        
        # Default placeholder for "In" or no uniform selected
        quantity_edit.setPlaceholderText("Enter quantity")
        quantity_edit.setToolTip("Enter the quantity for this inventory transaction")

    def update_price_field(self, uniform_combo, type_combo, price_edit):
        """Update price field based on uniform and type selection"""
        uniform_id = uniform_combo.currentData()
        inv_type = type_combo.currentText()
        
        if uniform_id and inv_type == "Out":
            try:
                with get_db_connection() as (conn, c):
                    # Get the current price from Uniform table
                    c.execute("SELECT UniformPrice FROM Uniform WHERE UniformId = ?", (uniform_id,))
                    result = c.fetchone()
                    if result and result[0] is not None:
                        current_price = result[0]
                        price_edit.setText(f"{current_price:.2f}")
                        price_edit.setReadOnly(True)
                        price_edit.setStyleSheet("background-color: #f0f0f0; color: #666;")
                        price_edit.setToolTip("Price is automatically set from uniform's current price (read-only for Out transactions)")
                        return
            except Exception as e:
                QMessageBox.warning(f"Error fetching uniform price: {e}")
        
        # For "In" transactions or if price not available
        price_edit.clear()
        price_edit.setReadOnly(False)
        price_edit.setStyleSheet("")
        price_edit.setPlaceholderText("Enter price per unit")
        price_edit.setToolTip("Enter the price per unit for this inventory transaction")
    
    def save_inventory(self, dialog, uniform_combo, type_combo, quantity_edit, price_edit, date_edit, ref_edit):
        """Save new inventory transaction"""
        uniform_id = uniform_combo.currentData()
        inv_type = type_combo.currentText()
        quantity_str = quantity_edit.text().strip()
        price_str = price_edit.text().strip()
        date_str = date_edit.date().toString("yyyy-MM-dd")
        references = ref_edit.text().strip()
        
        # Validation
        if not uniform_id:
            QMessageBox.warning(dialog, "Error", "Please select a uniform")
            return
            
        if not quantity_str:
            QMessageBox.warning(dialog, "Error", "Please enter quantity")
            return
            
        if not price_str and inv_type == "In":  # Price is required for "In" transactions
            QMessageBox.warning(dialog, "Error", "Please enter price")
            return
            
        try:
            quantity = int(quantity_str)
            if quantity <= 0:
                QMessageBox.warning(dialog, "Error", "Quantity must be positive")
                return
                
            price = float(price_str) if price_str else 0.0
            if price <= 0:
                QMessageBox.warning(dialog, "Error", "Price cannot be negative")
                return
                
        except ValueError:
            QMessageBox.warning(dialog, "Error", "Invalid quantity or price format")
            return
        
        # Additional validation for "Out" transactions
        if inv_type == "Out":
            try:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT Quantity FROM Uniform WHERE UniformId = ?", (uniform_id,))
                    result = c.fetchone()
                    if result and quantity > result[0]:
                        QMessageBox.warning(dialog, "Error", 
                                        f"Cannot issue {quantity} units. Only {result[0]} units available in stock.")
                        return
            except Exception as e:
                QMessageBox.warning(dialog, "Error", f"Failed to check available stock: {str(e)}")
                return
        
        try:
            with get_db_connection() as (conn, c):
                # Insert inventory transaction
                c.execute('''INSERT INTO UniformInventory 
                            (UniformID, Type, Quantity, Price, Date, ReferencesNotes)
                            VALUES (?, ?, ?, ?, ?, ?)''',
                        (uniform_id, inv_type, quantity, price, date_str, references))
                
                # Update uniform quantity
                if inv_type == "In":
                    c.execute("UPDATE Uniform SET Quantity = Quantity + ?, Price = ? WHERE UniformId = ?",
                            (quantity, price, uniform_id))
                else:
                    c.execute("UPDATE Uniform SET Quantity = Quantity - ? WHERE UniformId = ?",
                            (quantity, uniform_id))
                
                conn.commit()
                QMessageBox.information(dialog, "Success", "Inventory transaction added successfully")
                dialog.accept()
                self.load_data()  # Refresh the main table
                
        except Exception as e:
            QMessageBox.warning(dialog, "Error", f"Failed to save inventory: {str(e)}")
    
    def update_inventory(self, inv_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute('''SELECT ui.UniformInventoryID, u.UniformName, ui.Type, ui.Quantity, ui.Price, ui.Date, ui.ReferencesNotes, u.UniformId
                             FROM UniformInventory ui
                             LEFT JOIN Uniform u ON ui.UniformID = u.UniformId
                             WHERE ui.UniformInventoryID = ?''', (inv_id,))
                data = c.fetchone()
            if data:
                dialog = QDialog(self)
                dialog.setWindowTitle("Update Inventory")
                layout = QVBoxLayout()
                form_layout = QFormLayout()
                uniform_combo = QComboBox()
                try:
                    with get_db_connection() as (conn, c):
                        c.execute("SELECT UniformId, UniformName FROM Uniform")
                        for uniform in c.fetchall():
                            uniform_combo.addItem(uniform[1], uniform[0])
                        uniform_combo.setCurrentIndex(uniform_combo.findData(data[7]))
                except Exception as e:
                    QMessageBox.warning(dialog, "Error", f"Failed to load uniforms: {str(e)}")
                    return
                type_combo = QComboBox()
                type_combo.addItems(["In", "Out"])
                type_combo.setCurrentText(data[2])
                quantity_edit = QLineEdit(str(data[3]))
                price_edit = QLineEdit(str(data[4]))
                date_edit = QDateEdit(QDate.fromString(data[5], "yyyy-MM-dd"))
                date_edit.setCalendarPopup(True)
                ref_edit = QLineEdit(data[6])
                form_layout.addRow("Uniform:", uniform_combo)
                form_layout.addRow("Type:", type_combo)
                form_layout.addRow("Quantity:", quantity_edit)
                form_layout.addRow("Price:", price_edit)
                form_layout.addRow("Date:", date_edit)
                form_layout.addRow("References:", ref_edit)
                layout.addLayout(form_layout)
                buttons = QHBoxLayout()
                save_btn = QPushButton("Save")
                save_btn.clicked.connect(lambda: self.save_updated_inventory(dialog, inv_id, uniform_combo, type_combo, quantity_edit, price_edit, date_edit, ref_edit))
                close_btn = QPushButton("Close")
                close_btn.clicked.connect(dialog.close)
                buttons.addWidget(save_btn)
                buttons.addWidget(close_btn)
                layout.addLayout(buttons)
                dialog.setLayout(layout)
                
                # Connect type change for validation
                type_combo.currentIndexChanged.connect(lambda: self.update_quantity_placeholder(uniform_combo, type_combo, quantity_edit))
                
                dialog.exec_()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load inventory: {str(e)}")

    def save_updated_inventory(self, dialog, inv_id, uniform_combo, type_combo, quantity_edit, price_edit, date_edit, ref_edit):
        uniform_id = uniform_combo.currentData()
        inv_type = type_combo.currentText()
        quantity_str = quantity_edit.text().strip()
        price_str = price_edit.text().strip()
        date_str = date_edit.date().toString("yyyy-MM-dd")
        references = ref_edit.text().strip()
        
        if not all([uniform_id, inv_type, quantity_str, price_str, date_str, references]):
            QMessageBox.warning(dialog, "Error", "All fields are required")
            return
        
        try:
            quantity = int(quantity_str)
            price = float(price_str)
            if quantity <= 0 or price < 0:
                QMessageBox.warning(dialog, "Error", "Quantity must be positive and Price non-negative")
                return
        except ValueError:
            QMessageBox.warning(dialog, "Error", "Invalid quantity or price")
            return
        
        try:
            with get_db_connection() as (conn, c):
                # Get original data for quantity adjustment
                c.execute('''SELECT UniformID, Type, Quantity, Date FROM UniformInventory WHERE UniformInventoryID = ?''', (inv_id,))
                original = c.fetchone()
                if not original:
                    QMessageBox.warning(dialog, "Error", "Inventory entry not found")
                    return
                    
                original_uniform_id = original[0]
                original_type = original[1]
                original_qty = original[2]
                original_date = original[3]
                
                # CRITICAL VALIDATION: Check if this "In" inventory has corresponding "Out" transactions
                if original_type == 'In':
                    # Check if there are any "Out" transactions for this uniform that might have used this inventory
                    c.execute('''
                        SELECT SUM(Quantity) FROM UniformInventory 
                        WHERE UniformID = ? AND Type = 'Out' AND Date >= ? AND UniformInventoryID != ?
                    ''', (original_uniform_id, original_date, inv_id))
                    out_quantity_after = c.fetchone()[0] or 0
                    
                    # If there are out transactions after this inventory date, we need to be cautious
                    if out_quantity_after > 0:
                        # Check if changing this would cause negative inventory
                        c.execute('''
                            SELECT SUM(CASE WHEN Type = 'In' THEN Quantity ELSE -Quantity END) 
                            FROM UniformInventory 
                            WHERE UniformID = ? AND Date < ?
                        ''', (original_uniform_id, original_date))
                        inventory_before = c.fetchone()[0] or 0
                        
                        # Calculate what the new inventory would be after our changes
                        if inv_type == 'In':
                            new_inventory_after = inventory_before + quantity - out_quantity_after
                        else:
                            new_inventory_after = inventory_before - quantity - out_quantity_after
                        
                        if new_inventory_after < 0:
                            QMessageBox.warning(dialog, "Error", 
                                "Cannot modify this inventory entry. The change would cause negative inventory "
                                "due to existing sales transactions that may have used this stock.")
                            return
                
                # VALIDATION: Prevent changing from "In" to "Out" if it would cause negative inventory
                if original_type == 'In' and inv_type == 'Out':
                    c.execute("SELECT Quantity FROM Uniform WHERE UniformId = ?", (uniform_id,))
                    current_stock = c.fetchone()[0] or 0
                    if quantity > current_stock + original_qty:  # We're removing original "In" and adding new "Out"
                        QMessageBox.warning(dialog, "Error",
                            "Changing from 'In' to 'Out' would cause negative inventory. "
                            f"Available: {current_stock}, Trying to remove: {quantity}")
                        return
                
                # VALIDATION: Prevent changing uniform if there are transactions for the original uniform
                if uniform_id != original_uniform_id:
                    c.execute('''
                        SELECT COUNT(*) FROM UniformInventory 
                        WHERE UniformID = ? AND Date >= ? AND UniformInventoryID != ?
                    ''', (original_uniform_id, original_date, inv_id))
                    later_transactions = c.fetchone()[0] > 0
                    
                    if later_transactions:
                        QMessageBox.warning(dialog, "Error",
                            "Cannot change uniform type. There are later inventory transactions "
                            "for the original uniform that depend on inventory consistency.")
                        return
                
                # Adjust original quantity reverse
                if original_type == 'In':
                    c.execute("UPDATE Uniform SET Quantity = Quantity - ? WHERE UniformId = ?", (original_qty, original_uniform_id))
                else:
                    c.execute("UPDATE Uniform SET Quantity = Quantity + ? WHERE UniformId = ?", (original_qty, original_uniform_id))
                
                # Validate new 'Out' if applicable
                if inv_type == 'Out':
                    c.execute("SELECT Quantity FROM Uniform WHERE UniformId = ?", (uniform_id,))
                    available = c.fetchone()[0] or 0
                    if quantity > available:
                        # Revert adjustment
                        if original_type == 'In':
                            c.execute("UPDATE Uniform SET Quantity = Quantity + ? WHERE UniformId = ?", (original_qty, original_uniform_id))
                        else:
                            c.execute("UPDATE Uniform SET Quantity = Quantity - ? WHERE UniformId = ?", (original_qty, original_uniform_id))
                        QMessageBox.warning(dialog, "Error", f"Quantity exceeds available stock. Available: {available}")
                        return
                
                # Update inventory
                c.execute('''UPDATE UniformInventory 
                            SET UniformID = ?, Type = ?, Quantity = ?, Price = ?, Date = ?, ReferencesNotes = ?
                            WHERE UniformInventoryID = ?''', (uniform_id, inv_type, quantity, price, date_str, references, inv_id))
                
                # Apply new quantity adjustment
                if inv_type == 'In':
                    c.execute("UPDATE Uniform SET Quantity = Quantity + ? WHERE UniformId = ?", (quantity, uniform_id))
                else:
                    c.execute("UPDATE Uniform SET Quantity = Quantity - ? WHERE UniformId = ?", (quantity, uniform_id))
                
                conn.commit()
                QMessageBox.information(dialog, "Success", "Inventory updated successfully")
                dialog.close()
                self.load_data()
                
        except Exception as e:
            # Try to revert if possible
            try:
                if 'conn' in locals() and 'c' in locals():
                    if original_type == 'In':
                        c.execute("UPDATE Uniform SET Quantity = Quantity + ? WHERE UniformId = ?", (original_qty, original_uniform_id))
                    else:
                        c.execute("UPDATE Uniform SET Quantity = Quantity - ? WHERE UniformId = ?", (original_qty, original_uniform_id))
                    conn.commit()
            except:
                pass
            QMessageBox.warning(dialog, "Error", f"Failed to update inventory: {str(e)}")

    def delete_inventory(self, inv_id):
        reply = QMessageBox.question(self, 'Confirm Delete', 
                                'Are you sure you want to delete this inventory record?\n\n'
                                'WARNING: This action cannot be undone and may affect inventory accuracy.',
                                QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    # Get original for reverse adjustment
                    c.execute('''SELECT UniformID, Type, Quantity, Date FROM UniformInventory WHERE UniformInventoryID = ?''', (inv_id,))
                    original = c.fetchone()
                    
                    if not original:
                        QMessageBox.warning(self, "Error", "Inventory entry not found")
                        return
                        
                    uniform_id = original[0]
                    inv_type = original[1]
                    quantity = original[2]
                    inventory_date = original[3]
                    
                    # CRITICAL VALIDATION: Check if deleting would cause inventory inconsistencies
                    if inv_type == 'In':
                        # Check if there are "Out" transactions after this date that might have used this stock
                        c.execute('''
                            SELECT SUM(Quantity) FROM UniformInventory 
                            WHERE UniformID = ? AND Type = 'Out' AND Date >= ?
                        ''', (uniform_id, inventory_date))
                        out_quantity_after = c.fetchone()[0] or 0
                        
                        if out_quantity_after > 0:
                            # Check current inventory before this transaction
                            c.execute('''
                                SELECT SUM(CASE WHEN Type = 'In' THEN Quantity ELSE -Quantity END) 
                                FROM UniformInventory 
                                WHERE UniformID = ? AND Date < ?
                            ''', (uniform_id, inventory_date))
                            inventory_before = c.fetchone()[0] or 0
                            
                            # Check if deletion would cause negative inventory
                            if inventory_before - out_quantity_after < 0:
                                QMessageBox.warning(self, "Error",
                                    "Cannot delete this inventory entry. It would cause negative inventory "
                                    "because there are sales transactions that may have used this stock.\n\n"
                                    f"Inventory before: {inventory_before}, Sales after: {out_quantity_after}")
                                return
                    
                    # VALIDATION: Check for dependent transactions after this date
                    c.execute('''
                        SELECT COUNT(*) FROM UniformInventory 
                        WHERE UniformID = ? AND Date >= ? AND UniformInventoryID != ?
                    ''', (uniform_id, inventory_date, inv_id))
                    later_transactions = c.fetchone()[0] > 0
                    
                    if later_transactions:
                        reply = QMessageBox.question(self, 'Warning',
                            "There are inventory transactions after this date. Deleting may cause inventory inaccuracies.\n\n"
                            "Are you sure you want to proceed?",
                            QMessageBox.Yes | QMessageBox.No)
                        
                        if reply != QMessageBox.Yes:
                            return
                    
                    # Reverse quantity update
                    if inv_type == 'In':
                        c.execute("UPDATE Uniform SET Quantity = Quantity - ? WHERE UniformId = ?", (quantity, uniform_id))
                    else:
                        c.execute("UPDATE Uniform SET Quantity = Quantity + ? WHERE UniformId = ?", (quantity, uniform_id))
                    
                    # For audit purposes, you might want to create an archive table first:
                    # CREATE TABLE UniformInventoryArchive AS SELECT * FROM UniformInventory WHERE 1=0;
                    # Then insert the record being deleted
                    try:
                        c.execute('''
                            INSERT INTO UniformInventoryArchive 
                            SELECT *, datetime('now'), 'Deleted by user' FROM UniformInventory 
                            WHERE UniformInventoryID = ?
                        ''', (inv_id,))
                    except:
                        # Archive table doesn't exist, proceed with regular delete
                        pass
                    
                    c.execute("DELETE FROM UniformInventory WHERE UniformInventoryID = ?", (inv_id,))
                    conn.commit()
                    
                    QMessageBox.information(self, "Success", 
                        "Inventory deleted successfully. Inventory quantities have been adjusted.")
                    self.load_data()
                    
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to delete inventory: {str(e)}")

class ReportGenerator(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.db = sqlite3.connect('Database.db')
        self.setup_ui()
        self.generate_report()

    def setup_ui(self):
        # Main Layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(20)

        # --- Header Section ---
        header_layout = QHBoxLayout()
        
        title_label = QLabel("Uniforms Report")
        title_label.setObjectName("pageTitle")
        header_layout.addWidget(title_label)
        
        header_layout.addStretch()
        
        # Date Range
        date_container = QFrame()
        date_container.setObjectName("cardFrame") # Simple frame for controls
        date_container.setStyleSheet(".QFrame { background-color: white; border: 1px solid #E2E8F0; border-radius: 5px; }")
        date_layout = QHBoxLayout(date_container)
        date_layout.setContentsMargins(10, 5, 10, 5)
        
        self.from_date = QDateEdit(QDate.currentDate().addMonths(-1))
        self.from_date.setCalendarPopup(True)
        self.from_date.setDisplayFormat("yyyy-MM-dd")
        
        self.to_date = QDateEdit(QDate.currentDate())
        self.to_date.setCalendarPopup(True)
        self.to_date.setDisplayFormat("yyyy-MM-dd")
        
        generate_btn = QPushButton("Generate")
        generate_btn.setObjectName("infoButton")
        generate_btn.clicked.connect(self.generate_report)
        
        date_layout.addWidget(QLabel("From:"))
        date_layout.addWidget(self.from_date)
        date_layout.addWidget(QLabel("To:"))
        date_layout.addWidget(self.to_date)
        date_layout.addWidget(generate_btn)
        
        header_layout.addWidget(date_container)
        
        # Export Buttons
        pdf_btn = QPushButton("PDF")
        pdf_btn.setObjectName("dangerButton") # Red for PDF
        pdf_btn.clicked.connect(self.export_pdf)
        
        excel_btn = QPushButton("Excel")
        excel_btn.setObjectName("successButton") # Green for Excel
        excel_btn.clicked.connect(self.export_excel)
        
        header_layout.addWidget(pdf_btn)
        header_layout.addWidget(excel_btn)
        
        main_layout.addLayout(header_layout)

        # --- Summary Cards Section ---
        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(15)
        
        self.card_stock_value = self.create_summary_card("Current Stock Value", "0.00", "#3182CE")
        self.card_stock_in = self.create_summary_card("Stock In (Period)", "0.00", "#38A169")
        self.card_stock_out = self.create_summary_card("Stock Out (Period)", "0.00", "#E53E3E")
        self.card_expenses = self.create_summary_card("Expenses (Period)", "0.00", "#DD6B20")
        
        cards_layout.addWidget(self.card_stock_value)
        cards_layout.addWidget(self.card_stock_in)
        cards_layout.addWidget(self.card_stock_out)
        cards_layout.addWidget(self.card_expenses)
        
        main_layout.addLayout(cards_layout)

        # --- Charts Section ---
        charts_layout = QHBoxLayout()
        charts_layout.setSpacing(20)
        
        # Stock Chart Frame
        stock_chart_frame = QFrame()
        stock_chart_frame.setObjectName("summaryCard")
        stock_chart_layout = QVBoxLayout(stock_chart_frame)
        stock_chart_layout.addWidget(QLabel("Stock Movement (In vs Out)"))
        self.stock_chart_view = QChartView()
        self.stock_chart_view.setMinimumHeight(300)
        self.stock_chart_view.setRenderHint(QPainter.Antialiasing)
        stock_chart_layout.addWidget(self.stock_chart_view)
        
        # Expense Chart Frame
        exp_chart_frame = QFrame()
        exp_chart_frame.setObjectName("summaryCard")
        exp_chart_layout = QVBoxLayout(exp_chart_frame)
        exp_chart_layout.addWidget(QLabel("Expenses by Category"))
        self.expense_chart_view = QChartView()
        self.expense_chart_view.setMinimumHeight(300)
        self.expense_chart_view.setRenderHint(QPainter.Antialiasing)
        exp_chart_layout.addWidget(self.expense_chart_view)
        
        charts_layout.addWidget(stock_chart_frame)
        charts_layout.addWidget(exp_chart_frame)
        
        main_layout.addLayout(charts_layout)

        # --- Detailed Tables Section (Tabs) ---
        tabs = QTabWidget()
        tabs.setMinimumHeight(300)
        
        # Procurement Tab
        self.procurement_table = QTableWidget()
        self.procurement_table.setColumnCount(5)
        self.procurement_table.setHorizontalHeaderLabels(["Supplier/Ref", "Item", "Qty", "Unit Price", "Total"])
        self.procurement_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.procurement_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.procurement_table.setAlternatingRowColors(True)
        tabs.addTab(self.procurement_table, "Procurement List (Stock In)")
        
        # Expenses Tab
        self.expense_table = QTableWidget()
        self.expense_table.setColumnCount(2)
        self.expense_table.setHorizontalHeaderLabels(["Category", "Total Amount"])
        self.expense_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.expense_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.expense_table.setAlternatingRowColors(True)
        tabs.addTab(self.expense_table, "Expense List")
        
        main_layout.addWidget(tabs)

    def create_summary_card(self, title, value, color):
        card = QFrame()
        card.setObjectName("summaryCard")
        card.setStyleSheet(f"border-left: 5px solid {color};")
        layout = QVBoxLayout(card)
        
        title_lbl = QLabel(title)
        title_lbl.setStyleSheet("color: #718096; font-size: 12px; font-weight: bold;")
        
        value_lbl = QLabel(value)
        value_lbl.setStyleSheet("color: #2D3748; font-size: 20px; font-weight: bold;")
        card.value_label = value_lbl # Store ref
        
        layout.addWidget(title_lbl)
        layout.addWidget(value_lbl)
        return card

    def generate_report(self):
        from_date = self.from_date.date().toString("yyyy-MM-dd")
        to_date = self.to_date.date().toString("yyyy-MM-dd")

        try:
            with get_db_connection() as (conn, cur):
                # 1. Current Stock Value (All time)
                cur.execute("SELECT SUM(Quantity), SUM(Quantity * UniformPrice) FROM Uniform")
                current_qty, current_value = cur.fetchone()
                self.card_stock_value.value_label.setText(f"{current_value or 0:,.2f}")

                # 2. Stock In (Period)
                cur.execute("""
                    SELECT SUM(Quantity), SUM(Quantity * Price) 
                    FROM UniformInventory 
                    WHERE Type = 'In' AND Date BETWEEN ? AND ?
                """, (from_date, to_date))
                in_qty, in_value = cur.fetchone()
                self.card_stock_in.value_label.setText(f"{in_value or 0:,.2f}")

                # 3. Stock Out (Period)
                cur.execute("""
                    SELECT SUM(Quantity), SUM(Quantity * Price) 
                    FROM UniformInventory 
                    WHERE Type = 'Out' AND Date BETWEEN ? AND ?
                """, (from_date, to_date))
                out_qty, out_value = cur.fetchone()
                self.card_stock_out.value_label.setText(f"{out_value or 0:,.2f}")
                
                # 4. Expenses (Period) - FROM UniformExpenses
                cur.execute("""
                    SELECT SUM(U_ExpAmount)
                    FROM UniformExpenses
                    WHERE U_ExpDate BETWEEN ? AND ?
                """, (from_date, to_date))
                exp_total = cur.fetchone()[0]
                self.card_expenses.value_label.setText(f"{exp_total or 0:,.2f}")

                # 5. Stock Chart
                self.update_stock_chart(in_value or 0, out_value or 0)

                # 6. Procurement Table
                cur.execute("""
                    SELECT ReferencesNotes, UniformName, ui.Quantity, Price, (ui.Quantity * Price) 
                    FROM UniformInventory ui
                    JOIN Uniform u ON ui.UniformID = u.UniformId 
                    WHERE Type = 'In' AND Date BETWEEN ? AND ?
                    ORDER BY Date DESC
                """, (from_date, to_date))
                proc_rows = cur.fetchall()
                self.procurement_table.setRowCount(len(proc_rows))
                for i, row in enumerate(proc_rows):
                    for j, val in enumerate(row):
                        self.procurement_table.setItem(i, j, QTableWidgetItem(str(val)))

                # 7. Expense Table & Chart - FROM UniformExpenses
                cur.execute("""
                    SELECT sc.SubCategoryName, SUM(ue.U_ExpAmount)
                    FROM UniformExpenses ue
                    JOIN SubCategory sc ON ue.U_SubCategory = sc.SubCatID
                    WHERE ue.U_ExpDate BETWEEN ? AND ?
                    GROUP BY sc.SubCategoryName
                    ORDER BY SUM(ue.U_ExpAmount) DESC
                """, (from_date, to_date))
                exp_rows = cur.fetchall()
                
                self.expense_table.setRowCount(len(exp_rows))
                for i, (cat, total) in enumerate(exp_rows):
                    self.expense_table.setItem(i, 0, QTableWidgetItem(cat))
                    self.expense_table.setItem(i, 1, QTableWidgetItem(f"{total:,.2f}"))
                
                self.update_expense_chart(exp_rows)

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to generate report: {str(e)}")

    def update_stock_chart(self, in_val, out_val):
        set0 = QBarSet("Value")
        set0.append([in_val, out_val])
        
        series = QBarSeries()
        series.append(set0)
        
        chart = QChart()
        chart.addSeries(series)
        chart.setTitle("Stock Value Movement")
        chart.setAnimationOptions(QChart.SeriesAnimations)
        
        categories = ["Stock In", "Stock Out"]
        axis_x = QBarCategoryAxis()
        axis_x.append(categories)
        chart.addAxis(axis_x, Qt.AlignBottom)
        series.attachAxis(axis_x)
        
        axis_y = QValueAxis()
        chart.addAxis(axis_y, Qt.AlignLeft)
        series.attachAxis(axis_y)
        
        chart.legend().setVisible(False)
        self.stock_chart_view.setChart(chart)

    def update_expense_chart(self, data):
        series = QPieSeries()
        for cat, total in data:
            series.append(cat, total)
        
        if not data:
             series.append("No Expenses", 1)
        
        chart = QChart()
        chart.addSeries(series)
        chart.setTitle("Expenses by Category")
        chart.setAnimationOptions(QChart.SeriesAnimations)
        chart.legend().setAlignment(Qt.AlignRight)
        
        self.expense_chart_view.setChart(chart)

    def export_pdf(self):
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import letter
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import inch
        
        file_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "", "PDF Files (*.pdf)")
        if not file_path:
            return

        try:
            from_date = self.from_date.date().toString("yyyy-MM-dd")
            to_date = self.to_date.date().toString("yyyy-MM-dd")
            
            doc = SimpleDocTemplate(file_path, pagesize=letter, topMargin=0.5*inch, bottomMargin=0.5*inch)
            styles = getSampleStyleSheet()
            story = []
            
            # Custom Styles
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=24,
                textColor=colors.HexColor('#1A365D'),
                spaceAfter=6,
                alignment=1
            )
            
            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading2'],
                fontSize=14,
                textColor=colors.HexColor('#2C5282'),
                spaceAfter=12,
                spaceBefore=12,
                textTransform='uppercase'
            )
            
            normal_style = styles['Normal']
            normal_style.fontSize = 11
            
            # --- HEADER ---
            story.append(Paragraph("Uniform Stock Report", title_style))
            story.append(Paragraph(f"<b>Period:</b> {from_date} to {to_date}", normal_style))
            story.append(Spacer(1, 0.3*inch))
            
            # --- SECTION 1: SUMMARY ---
            story.append(Paragraph("1. Summary Overview", heading_style))
            
            summary_data = [
                ['Metric', 'Value'],
                ['Current Stock Value', self.card_stock_value.value_label.text()],
                ['Stock In (Period)', self.card_stock_in.value_label.text()],
                ['Stock Out (Period)', self.card_stock_out.value_label.text()],
                ['Total Expenses (Period)', self.card_expenses.value_label.text()],
            ]
            
            summary_table = Table(summary_data, colWidths=[3*inch, 2*inch])
            summary_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2C5282')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 11),
                ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CBD5E0')),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7FAFC')]),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ]))
            story.append(summary_table)
            story.append(Spacer(1, 0.3*inch))
            
            # --- SECTION 2: CURRENT STOCK ---
            story.append(Paragraph("2. Current Stock Details by Item", heading_style))
            
            stock_data = [['Uniform Item', 'Qty', 'Unit Price', 'Stock Value']]
            total_stock = 0
            
            try:
                with get_db_connection() as (conn, cur):
                    cur.execute("""
                        SELECT UniformName, Quantity, UniformPrice, (Quantity * UniformPrice)
                        FROM Uniform
                        ORDER BY UniformName
                    """)
                    for name, qty, price, value in cur.fetchall():
                        qty = qty or 0
                        price = price or 0
                        value = value or 0
                        total_stock += value
                        stock_data.append([name, str(qty), f"{price:,.2f}", f"{value:,.2f}"])
                    
                    stock_data.append(['TOTAL', '', '', f"{total_stock:,.2f}"])
            except Exception as e:
                stock_data.append(['Error', str(e), '', ''])
            
            stock_table = Table(stock_data, colWidths=[2.5*inch, 0.8*inch, 1.2*inch, 1.5*inch])
            stock_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2C5282')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CBD5E0')),
                ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#F7FAFC')]),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#EDF2F7')),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ]))
            story.append(stock_table)
            story.append(Spacer(1, 0.3*inch))
            
            # --- SECTION 3: INVENTORY MOVEMENT ---
            story.append(Paragraph("3. Inventory Movement (Purchases & Sales)", heading_style))
            
            inv_data = [['Reference', 'Item', 'Type', 'Qty', 'Unit Price', 'Total Value', 'Date']]
            
            try:
                with get_db_connection() as (conn, cur):
                    cur.execute(f"""
                        SELECT ReferencesNotes, UniformName, Type, ui.Quantity, Price, (ui.Quantity * Price), Date
                        FROM UniformInventory ui
                        JOIN Uniform u ON ui.UniformID = u.UniformId 
                        WHERE Date BETWEEN '{from_date}' AND '{to_date}'
                        ORDER BY Date DESC
                    """)
                    for ref, name, itype, qty, price, value, date in cur.fetchall():
                        inv_data.append([
                            ref if ref else '-',
                            name,
                            itype,
                            str(qty),
                            f"{price:,.2f}",
                            f"{value:,.2f}",
                            date
                        ])
            except Exception as e:
                inv_data.append(['Error', str(e), '', '', '', '', ''])
            
            inv_table = Table(inv_data, colWidths=[0.9*inch, 1.5*inch, 0.7*inch, 0.6*inch, 0.9*inch, 1.0*inch, 0.9*inch])
            inv_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2C5282')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (3, 0), (-1, -1), 'RIGHT'),
                ('ALIGN', (0, 0), (2, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CBD5E0')),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F7FAFC')]),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ]))
            story.append(inv_table)
            story.append(Spacer(1, 0.3*inch))
            
            # --- SECTION 4: EXPENSES ---
            story.append(Paragraph("4. Expenses Details", heading_style))
            
            exp_data = [['Category', 'Amount', '% of Total']]
            total_expenses = 0
            
            try:
                with get_db_connection() as (conn, cur):
                    cur.execute(f"""
                        SELECT sc.SubCategoryName, SUM(ue.U_ExpAmount)
                        FROM UniformExpenses ue
                        JOIN SubCategory sc ON ue.U_SubCategory = sc.SubCatID
                        WHERE ue.U_ExpDate BETWEEN '{from_date}' AND '{to_date}'
                        GROUP BY sc.SubCategoryName
                        ORDER BY SUM(ue.U_ExpAmount) DESC
                    """)
                    expense_list = cur.fetchall()
                    total_expenses = sum(amount for _, amount in expense_list) if expense_list else 0
                    
                    for category, amount in expense_list:
                        percentage = (amount / total_expenses * 100) if total_expenses > 0 else 0
                        exp_data.append([category, f"{amount:,.2f}", f"{percentage:.1f}%"])
                    
                    exp_data.append(['TOTAL EXPENSES', f"{total_expenses:,.2f}", '100.0%'])
            except Exception as e:
                exp_data.append(['Error', str(e), ''])
            
            exp_table = Table(exp_data, colWidths=[3*inch, 1.5*inch, 1.5*inch])
            exp_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2C5282')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 11),
                ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CBD5E0')),
                ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#F7FAFC')]),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#FFE5B4')),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ]))
            story.append(exp_table)
            story.append(Spacer(1, 0.5*inch))
            
            # --- FOOTER ---
            footer_style = ParagraphStyle(
                'Footer',
                parent=styles['Normal'],
                fontSize=9,
                textColor=colors.HexColor('#718096'),
                alignment=1
            )
            story.append(Paragraph(f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", footer_style))
            story.append(Paragraph("School Finance MIS — System-managed data only", footer_style))
            
            doc.build(story)
            QMessageBox.information(self, "Success", "PDF Report exported successfully!")
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to export PDF: {str(e)}")

    def export_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        try:
            from_date = self.from_date.date().toString("yyyy-MM-dd")
            to_date = self.to_date.date().toString("yyyy-MM-dd")
            
            wb = openpyxl.Workbook()
            
            # Remove default sheet and create our sheets
            wb.remove(wb.active)
            
            # Sheet 1: Summary
            ws_summary = wb.create_sheet("Summary", 0)
            ws_summary = self._create_summary_sheet(ws_summary, from_date, to_date)
            
            # Sheet 2: Current Stock
            ws_stock = wb.create_sheet("Current Stock", 1)
            ws_stock = self._create_stock_sheet(ws_stock)
            
            # Sheet 3: Inventory Movement
            ws_inventory = wb.create_sheet("Inventory Movement", 2)
            ws_inventory = self._create_inventory_sheet(ws_inventory, from_date, to_date)
            
            # Sheet 4: Expenses
            ws_expenses = wb.create_sheet("Expenses", 3)
            ws_expenses = self._create_expenses_sheet(ws_expenses, from_date, to_date)
            
            wb.save(file_path)
            QMessageBox.information(self, "Success", "Excel Report exported successfully!\n\nFile: " + file_path)
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to export Excel: {str(e)}")

    def _create_summary_sheet(self, ws, from_date, to_date):
        """Create Summary sheet with overview"""
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        # Styles
        title_font = Font(size=18, bold=True, color="FFFFFF")
        title_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
        summary_fill = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                    top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Title
        ws.merge_cells('A1:D1')
        title_cell = ws['A1']
        title_cell.value = "Uniform Stock Report - Summary"
        title_cell.font = title_font
        title_cell.fill = title_fill
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 30
        
        # Period
        ws.merge_cells('A2:D2')
        period_cell = ws['A2']
        period_cell.value = f"Report Period: {from_date} to {to_date}"
        period_cell.font = Font(size=11, color="718096")
        period_cell.alignment = Alignment(horizontal='center')
        
        # Summary Section
        ws.merge_cells('A4:D4')
        summary_header = ws['A4']
        summary_header.value = "SUMMARY OVERVIEW"
        summary_header.font = header_font
        summary_header.fill = header_fill
        
        # Summary Data
        row = 6
        summary_items = [
            ("Current Stock Value", self.card_stock_value.value_label.text()),
            ("Total Stock In (Period)", self.card_stock_in.value_label.text()),
            ("Total Stock Out (Period)", self.card_stock_out.value_label.text()),
            ("Total Expenses (Period)", self.card_expenses.value_label.text()),
        ]
        
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 25
        
        for label, value in summary_items:
            ws[f'A{row}'] = label
            ws[f'A{row}'].font = Font(bold=True, size=11)
            ws[f'B{row}'] = value
            ws[f'B{row}'].font = Font(size=11)
            ws[f'B{row}'].fill = summary_fill
            ws[f'B{row}'].alignment = Alignment(horizontal='right')
            ws[f'A{row}'].border = border
            ws[f'B{row}'].border = border
            row += 1
        
        # Generated date
        ws.merge_cells('A13:D13')
        footer = ws['A13']
        footer.value = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        footer.font = Font(size=9, italic=True, color="A0AEC0")
        footer.alignment = Alignment(horizontal='center')
        
        return ws

    def _create_stock_sheet(self, ws):
        """Create Current Stock Details sheet"""
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
        total_fill = PatternFill(start_color="E8EFFB", end_color="E8EFFB", fill_type="solid")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                    top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Title
        ws.merge_cells('A1:D1')
        title = ws['A1']
        title.value = "Current Stock Details by Item"
        title.font = Font(size=14, bold=True, color="1A365D")
        title.alignment = Alignment(horizontal='center')
        ws.row_dimensions[1].height = 25
        
        # Headers
        headers = ["Uniform Item", "Current Qty", "Unit Price", "Stock Value"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 18
        ws.row_dimensions[3].height = 20
        
        # Data
        try:
            with get_db_connection() as (conn, cur):
                cur.execute("""
                    SELECT UniformName, Quantity, UniformPrice, (Quantity * UniformPrice)
                    FROM Uniform
                    ORDER BY UniformName
                """)
                items = cur.fetchall()
                
                row = 4
                total_value = 0
                
                for name, qty, price, value in items:
                    ws.cell(row, 1).value = name
                    ws.cell(row, 2).value = qty if qty else 0
                    ws.cell(row, 3).value = price if price else 0
                    ws.cell(row, 4).value = value if value else 0
                    
                    ws.cell(row, 2).alignment = Alignment(horizontal='center')
                    ws.cell(row, 3).alignment = Alignment(horizontal='right')
                    ws.cell(row, 4).alignment = Alignment(horizontal='right')
                    ws.cell(row, 3).number_format = '#,##0.00'
                    ws.cell(row, 4).number_format = '#,##0.00'
                    
                    for col in range(1, 5):
                        ws.cell(row, col).border = border
                    
                    total_value += value if value else 0
                    row += 1
                
                # Total Row
                ws.cell(row, 1).value = "TOTAL"
                ws.cell(row, 1).font = Font(bold=True, size=11)
                ws.cell(row, 4).value = total_value
                ws.cell(row, 4).font = Font(bold=True, size=11)
                ws.cell(row, 4).number_format = '#,##0.00'
                ws.cell(row, 4).alignment = Alignment(horizontal='right')
                
                for col in range(1, 5):
                    ws.cell(row, col).fill = total_fill
                    ws.cell(row, col).border = border
        
        except Exception as e:
            ws['A4'] = f"Error loading data: {str(e)}"
        
        return ws

    def _create_inventory_sheet(self, ws, from_date, to_date):
        """Create Inventory Movement sheet (Purchases & Sales)"""
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        header_font = Font(bold=True, color="FFFFFF", size=10)
        header_fill = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
        in_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
        out_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                    top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Title
        ws.merge_cells('A1:G1')
        title = ws['A1']
        title.value = "Inventory Movement (Purchases & Sales)"
        title.font = Font(size=14, bold=True, color="1A365D")
        title.alignment = Alignment(horizontal='center')
        ws.row_dimensions[1].height = 25
        
        # Period
        ws.merge_cells('A2:G2')
        period = ws['A2']
        period.value = f"Period: {from_date} to {to_date}"
        period.font = Font(size=10, italic=True)
        period.alignment = Alignment(horizontal='center')
        
        # Headers
        headers = ["Reference", "Uniform Item", "Type", "Quantity", "Unit Price", "Total Value", "Date"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 10
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 12
        ws.column_dimensions['F'].width = 15
        ws.column_dimensions['G'].width = 15
        ws.row_dimensions[4].height = 20
        
        # Data
        try:
            with get_db_connection() as (conn, cur):
                cur.execute(f"""
                    SELECT ReferencesNotes, UniformName, Type, ui.Quantity, Price, (ui.Quantity * Price), Date
                    FROM UniformInventory ui
                    JOIN Uniform u ON ui.UniformID = u.UniformId 
                    WHERE Date BETWEEN '{from_date}' AND '{to_date}'
                    ORDER BY Date DESC
                """)
                items = cur.fetchall()
                
                row = 5
                total_in = 0
                total_out = 0
                
                for ref, name, itype, qty, price, value, date in items:
                    ws.cell(row, 1).value = ref if ref else "-"
                    ws.cell(row, 2).value = name
                    ws.cell(row, 3).value = itype
                    ws.cell(row, 4).value = qty
                    ws.cell(row, 5).value = price if price else 0
                    ws.cell(row, 6).value = value if value else 0
                    ws.cell(row, 7).value = date
                    
                    # Apply color based on type
                    row_fill = in_fill if itype == "In" else out_fill
                    
                    for col in range(1, 8):
                        ws.cell(row, col).border = border
                        ws.cell(row, col).fill = row_fill
                    
                    ws.cell(row, 3).alignment = Alignment(horizontal='center', vertical='center')
                    ws.cell(row, 4).alignment = Alignment(horizontal='center')
                    ws.cell(row, 5).alignment = Alignment(horizontal='right')
                    ws.cell(row, 5).number_format = '#,##0.00'
                    ws.cell(row, 6).alignment = Alignment(horizontal='right')
                    ws.cell(row, 6).number_format = '#,##0.00'
                    ws.cell(row, 7).alignment = Alignment(horizontal='center')
                    
                    if itype == "In":
                        total_in += value if value else 0
                    else:
                        total_out += value if value else 0
                    
                    row += 1
                
                # Summary row
                ws.cell(row, 2).value = "TOTALS"
                ws.cell(row, 2).font = Font(bold=True)
                ws.cell(row, 6).value = total_in + total_out
                ws.cell(row, 6).font = Font(bold=True)
                ws.cell(row, 6).number_format = '#,##0.00'
                
                for col in range(1, 8):
                    ws.cell(row, col).fill = PatternFill(start_color="E8EFFB", end_color="E8EFFB", fill_type="solid")
                    ws.cell(row, col).border = border
        
        except Exception as e:
            ws['A5'] = f"Error loading data: {str(e)}"
        
        return ws

    def _create_expenses_sheet(self, ws, from_date, to_date):
        """Create Expenses Details sheet"""
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="2C5282", end_color="2C5282", fill_type="solid")
        total_fill = PatternFill(start_color="FFE5B4", end_color="FFE5B4", fill_type="solid")
        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                    top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Title
        ws.merge_cells('A1:C1')
        title = ws['A1']
        title.value = "Expenses Details"
        title.font = Font(size=14, bold=True, color="1A365D")
        title.alignment = Alignment(horizontal='center')
        ws.row_dimensions[1].height = 25
        
        # Period
        ws.merge_cells('A2:C2')
        period = ws['A2']
        period.value = f"Period: {from_date} to {to_date}"
        period.font = Font(size=10, italic=True)
        period.alignment = Alignment(horizontal='center')
        
        # Headers
        headers = ["Category", "Amount", "Percentage"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border
        
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 15
        ws.row_dimensions[4].height = 20
        
        # Data
        try:
            with get_db_connection() as (conn, cur):
                cur.execute(f"""
                    SELECT sc.SubCategoryName, SUM(ue.U_ExpAmount)
                    FROM UniformExpenses ue
                    JOIN SubCategory sc ON ue.U_SubCategory = sc.SubCatID
                    WHERE ue.U_ExpDate BETWEEN '{from_date}' AND '{to_date}'
                    GROUP BY sc.SubCategoryName
                    ORDER BY SUM(ue.U_ExpAmount) DESC
                """)
                items = cur.fetchall()
                
                row = 5
                total_expenses = sum(amount for _, amount in items) if items else 0
                
                for category, amount in items:
                    percentage = (amount / total_expenses * 100) if total_expenses > 0 else 0
                    
                    ws.cell(row, 1).value = category
                    ws.cell(row, 2).value = amount if amount else 0
                    ws.cell(row, 3).value = percentage if percentage else 0
                    
                    ws.cell(row, 2).alignment = Alignment(horizontal='right')
                    ws.cell(row, 2).number_format = '#,##0.00'
                    ws.cell(row, 3).alignment = Alignment(horizontal='right')
                    ws.cell(row, 3).number_format = '0.0"%"'
                    
                    for col in range(1, 4):
                        ws.cell(row, col).border = border
                    
                    row += 1
                
                # Total Row
                ws.cell(row, 1).value = "TOTAL EXPENSES"
                ws.cell(row, 1).font = Font(bold=True, size=11)
                ws.cell(row, 2).value = total_expenses
                ws.cell(row, 2).font = Font(bold=True, size=11)
                ws.cell(row, 2).number_format = '#,##0.00'
                ws.cell(row, 2).alignment = Alignment(horizontal='right')
                ws.cell(row, 3).value = 100.0
                ws.cell(row, 3).font = Font(bold=True, size=11)
                ws.cell(row, 3).number_format = '0.0"%"'
                ws.cell(row, 3).alignment = Alignment(horizontal='right')
                
                for col in range(1, 4):
                    ws.cell(row, col).fill = total_fill
                    ws.cell(row, col).border = border
        
        except Exception as e:
            ws['A5'] = f"Error loading data: {str(e)}"
        
        return ws


#### start of expenditure #####
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
        QMessageBox.warning(f"Stock recalculation error: {e}")
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
        QMessageBox.warning(f"Validation error: {e}")
        return False, 0

class Expenditure(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Expenditure Management")
        layout = QVBoxLayout()
        self.tabs = QTabWidget()
        self.tab1 = Supplies()
        self.tab2 = SuppliesInventory()
        self.tab3 = Expense()
        self.tabs.addTab(self.tab1, "Supplies")
        self.tabs.addTab(self.tab2, "Supplies Inventory")
        self.tabs.addTab(self.tab3, "Expenses")
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
                    self.table.setRowHeight(row_idx, 40)
                    # Data columns
                    for col_idx, data in enumerate(row_data):
                        self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(data) if data is not None else ""))
                    
                    # Actions - ComboBox
                    supply_id = row_data[0]
                    supply_name = row_data[1]
                    action_combo = QComboBox()
                    action_combo.addItems(["Select Action", "Manage", "Edit", "Delete"])
                    
                    # Direct lambda with handler
                    action_combo.currentIndexChanged.connect(
                        lambda idx, sid=supply_id, sname=supply_name, cb=action_combo: 
                        self.on_supply_action_selected(idx, sid, sname, cb)
                    )
                    self.table.setCellWidget(row_idx, 5, action_combo)
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load data: {e}")

    def on_supply_action_selected(self, index, supply_id, supply_name, combo_box):
        """Handle supply action selection"""
        if index == 0:
            return
        
        action = combo_box.itemText(index)
        combo_box.setCurrentIndex(0)  # Reset
        
        if action == "Manage":
            self.manage_inventory(supply_id, supply_name)
        elif action == "Edit":
            self.edit_supply(supply_id)
        elif action == "Delete":
            self.delete_supply(supply_id)

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
                    QMessageBox.warning(e)

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
                    QMessageBox.warning(e)

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
                    self.table.setRowHeight(row_idx, 40)
                    for col_idx, data in enumerate(row_data):
                        self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(data)))
                    
                    # Actions
                    action_widget = QWidget()
                    action_layout = QHBoxLayout(action_widget)
                    action_layout.setContentsMargins(5, 2, 5, 2)
                    action_combo = QComboBox()
                    action_combo.addItems(["Select Action", "Edit", "Delete"])
                    action_combo.currentIndexChanged.connect(lambda idx, inv_id=row_data[0], supply_id=row_data[1], cb=action_combo: 
                                                     self.on_inventory_action_selected(idx, inv_id, supply_id, cb))
                    action_layout.addWidget(action_combo)
                    self.table.setCellWidget(row_idx, 7, action_widget)
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load inventory: {e}")
    
    def on_inventory_action_selected(self, index, inv_id, supply_id, combo_box):
        """Handle inventory action selection"""
        if index == 0:
            return
        action = combo_box.itemText(index)
        combo_box.setCurrentIndex(0)  # Reset
        if action == "Edit":
            self.edit_inventory(inv_id)
        elif action == "Delete":
            self.delete_inventory(inv_id, supply_id)

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
        self.Heading.setStyleSheet("font-size: 18px; font-weight: bold; color: #2D3748;")
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
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["ID", "Category", "Subcategory", "Amount", "Date", "Description", "Actions"])
        self.table.hideColumn(0)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Fixed)
        self.table.setColumnWidth(6, 180)
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
                    self.table.setRowHeight(idx, 40)
                    # row: ExpID, CatName, SubCategoryName, ExpAmount, ExpDate, ExpDescription
                    self.table.setItem(idx, 0, QTableWidgetItem(str(row[0] if row[0] is not None else "")))
                    self.table.setItem(idx, 1, QTableWidgetItem(str(row[1] if row[1] else "")))
                    self.table.setItem(idx, 2, QTableWidgetItem(str(row[2] if row[2] else "")))
                    self.table.setItem(idx, 3, QTableWidgetItem(str(row[3] if row[3] is not None else "")))
                    self.table.setItem(idx, 4, QTableWidgetItem(str(row[4] if row[4] else "")))
                    self.table.setItem(idx, 5, QTableWidgetItem(str(row[5] if row[5] else "")))

                    # Actions - ComboBox
                    row_id = row[0]
                    action_combo = QComboBox()
                    action_combo.addItems(["Select Action","Edit","Delete"])
                    action_combo.currentIndexChanged.connect(lambda index, rid=row_id: self.handle_action_change(index, rid, action_combo))
                    self.table.setCellWidget(idx, 6, action_combo)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load expenses: {e}")
   
    def handle_action_change(self, index, exp_id, combo):
        if index == 0:
            return
        if index == 1:
            self.edit_expense(exp_id)
        elif index == 2:
            self.delete_expense(exp_id)

    def add_expense(self):
        dialog = QDialog(self)
        dialog.setStyleSheet(DialogStyles())
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
            QMessageBox.warning(f"Error loading categories: {e}")

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
                    QMessageBox.warning(e)

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
            dialog.setStyleSheet(DialogStyles())
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

### end of expenditure ####
### Categories management
class CategoryManagement(QWidget):
    def __init__(self):
        super().__init__()
        self.current_department_id = None
        self.current_category_id = None
        self.init_ui()
        # Delay loading departments until UI is fully initialized
        QTimer.singleShot(100, self.load_departments)
        
    def init_ui(self):
        # Main Layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # ========== HEADER SECTION ==========
        header_frame = QFrame()
        header_frame.setMaximumHeight(80)
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(10)

        # ========== DEPARTMENT QUICK ACCESS BAR ==========
        self.dept_bar_frame = QFrame()
        self.dept_bar_frame.setMinimumHeight(60)
        dept_bar_layout = QHBoxLayout(self.dept_bar_frame)
        dept_bar_layout.setContentsMargins(0, 5, 0, 5)
        dept_bar_layout.setSpacing(15)

        # Department label
        dept_label = QLabel("Department:")
        dept_label.setStyleSheet("font-weight: bold; color: #4A5568; font-size: 14px;")
        dept_bar_layout.addWidget(dept_label)

        # Departments as horizontal tabs/buttons (using a widget container)
        self.dept_buttons_container = QWidget()
        self.dept_buttons_layout = QHBoxLayout(self.dept_buttons_container)
        self.dept_buttons_layout.setSpacing(8)
        self.dept_buttons_layout.setContentsMargins(0, 0, 0, 0)
        dept_bar_layout.addWidget(self.dept_buttons_container)

        # Department Action Buttons (Edit/Delete)
        dept_action_layout = QHBoxLayout()
        dept_action_layout.setSpacing(5)

        self.edit_dept_btn = QPushButton("Edit")
        self.edit_dept_btn.clicked.connect(self.edit_department)
        self.edit_dept_btn.setFixedHeight(30)
        self.edit_dept_btn.setFixedWidth(60)
        self.edit_dept_btn.setEnabled(False)
        dept_action_layout.addWidget(self.edit_dept_btn)

        self.delete_dept_btn = QPushButton("Delete")
        self.delete_dept_btn.clicked.connect(self.delete_department)
        self.delete_dept_btn.setFixedHeight(30)
        self.delete_dept_btn.setFixedWidth(70)
        self.delete_dept_btn.setEnabled(False)
        dept_action_layout.addWidget(self.delete_dept_btn)

        dept_bar_layout.addLayout(dept_action_layout)
        dept_bar_layout.addStretch()

        # Add Department Button
        add_dept_btn = QPushButton("+ Add Department")
        add_dept_btn.clicked.connect(self.add_department)
        add_dept_btn.setFixedHeight(30)
        dept_bar_layout.addWidget(add_dept_btn)

        header_layout.addWidget(self.dept_bar_frame)
        main_layout.addWidget(header_frame)

        # ========== MAIN CONTENT AREA ==========
        content_container = QFrame()
        content_container.setObjectName("contentContainer")
        content_layout = QHBoxLayout(content_container)
        content_layout.setSpacing(20)
        content_layout.setContentsMargins(0, 0, 0, 0)

        # ========== LEFT PANEL: Categories (45%) ==========
        category_panel = QFrame()
        category_panel.setObjectName("categoryPanel")
        category_layout = QVBoxLayout(category_panel)
        category_layout.setContentsMargins(0, 0, 0, 0)
        category_layout.setSpacing(10)

        # Categories Header
        category_header = QFrame()
        category_header.setMaximumHeight(50)
        category_header_layout = QHBoxLayout(category_header)
        category_header_layout.setContentsMargins(0, 0, 0, 0)
        
        self.category_title = QLabel("CATEGORIES")
        self.category_title.setStyleSheet("""
            QLabel {
                font-size: 15px;
                font-weight: bold;
                color: #2C3E50;
                padding: 10px 0;
            }
        """)
        category_header_layout.addWidget(self.category_title)
        category_header_layout.addStretch()
        
        # Add Category Button
        add_category_btn = QPushButton("+ Add Category")
        add_category_btn.clicked.connect(self.add_category)
        add_category_btn.setFixedHeight(30)
        category_header_layout.addWidget(add_category_btn)
        
        category_layout.addWidget(category_header)

        # Category Filter Bar
        filter_frame = QFrame()
        filter_frame.setMaximumHeight(40)
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(0, 0, 0, 0)
        
        filter_label = QLabel("Filter by Type:")
        filter_label.setStyleSheet("color: #718096; font-size: 13px;")
        filter_layout.addWidget(filter_label)
        
        self.type_filter = QComboBox()
        self.type_filter.addItems(["All Types", "Supplies", "Expenses"])
        self.type_filter.setFixedHeight(28)
        self.type_filter.currentIndexChanged.connect(self.load_categories)
        filter_layout.addWidget(self.type_filter)
        
        filter_layout.addStretch()
        category_layout.addWidget(filter_frame)

        # Categories Table
        self.category_table = QTableWidget()
        self.category_table.setColumnCount(4)
        self.category_table.setHorizontalHeaderLabels(["Type", "Category", "Subcats", "Actions"])
        self.category_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.category_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.category_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.category_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.category_table.setAlternatingRowColors(True)
        self.category_table.itemClicked.connect(self.on_category_selected)
        self.category_table.verticalHeader().setVisible(False)
        category_layout.addWidget(self.category_table)

        content_layout.addWidget(category_panel, 4)

        # ========== RIGHT PANEL: Subcategories (55%) ==========
        subcategory_panel = QFrame()
        subcategory_panel.setObjectName("subcategoryPanel")
        subcategory_layout = QVBoxLayout(subcategory_panel)
        subcategory_layout.setContentsMargins(0, 0, 0, 0)
        subcategory_layout.setSpacing(10)

        # Subcategories Header
        subcategory_header = QFrame()
        subcategory_header.setMaximumHeight(50)
        subcategory_header_layout = QHBoxLayout(subcategory_header)
        subcategory_header_layout.setContentsMargins(0, 0, 0, 0)
        
        self.subcategory_title = QLabel("SUBCATEGORIES")
        self.subcategory_title.setStyleSheet("""
            QLabel {
                font-size: 15px;
                font-weight: bold;
                color: #2C3E50;
                padding: 10px 0;
            }
        """)
        subcategory_header_layout.addWidget(self.subcategory_title)
        subcategory_header_layout.addStretch()
        
        # Add Subcategory Button
        add_subcat_btn = QPushButton("+ Add Subcategory")
        add_subcat_btn.clicked.connect(self.add_subcategory)
        add_subcat_btn.setFixedHeight(30)
        subcategory_header_layout.addWidget(add_subcat_btn)
        
        subcategory_layout.addWidget(subcategory_header)

        # Subcategories Info Bar
        info_frame = QFrame()
        info_frame.setMaximumHeight(40)
        info_layout = QHBoxLayout(info_frame)
        info_layout.setContentsMargins(0, 0, 0, 0)
        
        self.subcat_info_label = QLabel("Select a category to view subcategories")
        self.subcat_info_label.setStyleSheet("color: #718096; font-size: 13px; font-style: italic;")
        info_layout.addWidget(self.subcat_info_label)
        
        info_layout.addStretch()
        
        self.usage_label = QLabel("Usage: 0 items")
        self.usage_label.setStyleSheet("""
            QLabel {
                background-color: #F7FAFC;
                color: #4A5568;
                padding: 4px 10px;
                border-radius: 4px;
                border: 1px solid #E2E8F0;
                font-size: 12px;
            }
        """)
        info_layout.addWidget(self.usage_label)
        
        subcategory_layout.addWidget(info_frame)

        # Subcategories Table
        self.subcategory_table = QTableWidget()
        self.subcategory_table.setColumnCount(3)
        self.subcategory_table.setHorizontalHeaderLabels(["Subcategory", "Usage", "Actions"])
        self.subcategory_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.subcategory_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.subcategory_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.subcategory_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.subcategory_table.setAlternatingRowColors(True)
        self.subcategory_table.verticalHeader().setVisible(False)
        subcategory_layout.addWidget(self.subcategory_table)

        content_layout.addWidget(subcategory_panel, 5)

        main_layout.addWidget(content_container)

        # ========== BREADCRUMB NAVIGATION ==========
        breadcrumb_frame = QFrame()
        breadcrumb_frame.setMaximumHeight(50)
        breadcrumb_layout = QHBoxLayout(breadcrumb_frame)
        breadcrumb_layout.setContentsMargins(15, 10, 15, 10)
        
        self.breadcrumb_label = QLabel("Select a department to begin")
        self.breadcrumb_label.setStyleSheet("font-size: 13px; color: #4A5568;")
        breadcrumb_layout.addWidget(self.breadcrumb_label)
        
        breadcrumb_layout.addStretch()
        
        # Navigation indicators
        indicator_layout = QHBoxLayout()
        indicator_layout.setSpacing(15)
        
        self.dept_indicator = QLabel("Department")
        self.dept_indicator.setStyleSheet("color: #718096; font-size: 12px;")
        indicator_layout.addWidget(self.dept_indicator)
        
        self.cat_indicator = QLabel("Category")
        self.cat_indicator.setStyleSheet("color: #718096; font-size: 12px;")
        indicator_layout.addWidget(self.cat_indicator)
        
        self.subcat_indicator = QLabel("Subcategory")
        self.subcat_indicator.setStyleSheet("color: #718096; font-size: 12px;")
        indicator_layout.addWidget(self.subcat_indicator)
        
        breadcrumb_layout.addLayout(indicator_layout)
        
        main_layout.addWidget(breadcrumb_frame)

        # Apply styles
        self.apply_styles()

    def apply_styles(self):
        """Apply all stylesheets"""
        self.setStyleSheet("""
            QFrame#contentContainer {
                background-color: transparent;
            }
            QPushButton {
                border-radius: 4px;
                font-weight: 500;
                border: 1px solid #CBD5E0;
                background-color: #F7FAFC;
                color: #4A5568;
                padding: 0 12px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #EDF2F7;
                border-color: #A0AEC0;
            }
            QPushButton:checked {
                background-color: #4299E1;
                color: white;
                border: none;
            }
            QPushButton:checked:hover {
                background-color: #3182CE;
            }
            
        """)

    def load_departments(self):
        """Load departments as horizontal buttons"""
        try:
            # Clear existing buttons
            self.clear_layout(self.dept_buttons_layout)
            
            with get_db_connection() as (conn, c):
                c.execute("SELECT DeptID, DeptName FROM Department ORDER BY DeptName")
                departments = c.fetchall()
                
                for dept_id, dept_name in departments:
                    dept_btn = QPushButton(dept_name)
                    dept_btn.setCheckable(True)
                    dept_btn.setFixedHeight(32)
                    dept_btn.setProperty("dept_id", dept_id)
                    
                    # Set initial checked state
                    if dept_id == self.current_department_id:
                        dept_btn.setChecked(True)
                    
                    dept_btn.clicked.connect(lambda checked, d_id=dept_id: self.on_department_selected(d_id))
                    self.dept_buttons_layout.addWidget(dept_btn)
                
                
                # Auto-select first department if none selected
                if not self.current_department_id and departments:
                    first_dept_id = departments[0][0]
                    self.on_department_selected(first_dept_id)
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load departments: {str(e)}")

    def clear_layout(self, layout):
        """Safely clear all widgets from a layout"""
        if layout:
            while layout.count():
                item = layout.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()

    def on_department_selected(self, dept_id):
        """Handle department selection from button"""
        self.current_department_id = dept_id
        
        # Enable edit/delete buttons
        self.edit_dept_btn.setEnabled(True)
        self.delete_dept_btn.setEnabled(True)
        
        # Update button states
        for i in range(self.dept_buttons_layout.count()):
            widget = self.dept_buttons_layout.itemAt(i).widget()
            if widget and isinstance(widget, QPushButton):
                widget.setChecked(widget.property("dept_id") == dept_id)
        
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT DeptName FROM Department WHERE DeptID = ?", (dept_id,))
                result = c.fetchone()
                if result:
                    dept_name = result[0]
                    
                    # Update UI
                    self.category_title.setText(f"CATEGORIES - {dept_name.upper()}")
                    self.breadcrumb_label.setText(f"Department: {dept_name}")
                    self.dept_indicator.setStyleSheet("color: #4299E1; font-weight: bold; font-size: 12px;")
                    self.cat_indicator.setStyleSheet("color: #718096; font-size: 12px;")
                    self.subcat_indicator.setStyleSheet("color: #718096; font-size: 12px;")
        
        except Exception as e:
            QMessageBox.warning(f"Error getting department name: {e}")
        
        # Load categories and clear subcategories
        self.load_categories()
        self.clear_subcategories()

    def on_category_selected(self, item):
        """Handle category selection"""
        if not item:
            return
            
        row = self.category_table.currentRow()
        if row == -1:
            return
            
        cat_id = self.category_table.item(row, 0).data(Qt.UserRole)
        cat_name = self.category_table.item(row, 1).text()
        
        self.current_category_id = cat_id
        
        # Get department name
        dept_name = ""
        for i in range(self.dept_buttons_layout.count()):
            widget = self.dept_buttons_layout.itemAt(i).widget()
            if widget and isinstance(widget, QPushButton) and widget.isChecked():
                dept_name = widget.text()
                break
        
        # Update UI
        self.subcategory_title.setText(f"SUBCATEGORIES - {cat_name.upper()}")
        self.breadcrumb_label.setText(f"Department: {dept_name} → Category: {cat_name}")
        self.dept_indicator.setStyleSheet("color: #4299E1; font-weight: bold; font-size: 12px;")
        self.cat_indicator.setStyleSheet("color: #4299E1; font-weight: bold; font-size: 12px;")
        self.subcat_indicator.setStyleSheet("color: #718096; font-size: 12px;")
        self.subcat_info_label.setText(f"Subcategories for {cat_name}")
        
        # Load subcategories
        self.load_subcategories()

    def clear_subcategories(self):
        """Clear subcategories display"""
        self.subcategory_table.setRowCount(0)
        self.subcategory_title.setText("SUBCATEGORIES")
        self.subcat_info_label.setText("Select a category to view subcategories")
        self.usage_label.setText("Usage: 0 items")
        
        # Reset indicators if no department selected
        if not self.current_department_id:
            self.breadcrumb_label.setText("Select a department to begin")
            self.dept_indicator.setStyleSheet("color: #718096; font-size: 12px;")
            self.cat_indicator.setStyleSheet("color: #718096; font-size: 12px;")
            self.subcat_indicator.setStyleSheet("color: #718096; font-size: 12px;")
            self.edit_dept_btn.setEnabled(False)
            self.delete_dept_btn.setEnabled(False)

    def load_categories(self):
        """Load categories for the selected department"""
        if not self.current_department_id:
            self.category_table.setRowCount(0)
            self.category_count_label.setText("Categories: 0")
            return
            
        try:
            type_filter = self.type_filter.currentText()
            
            with get_db_connection() as (conn, c):
                # Build query
                if type_filter == "All Types":
                    query = "SELECT CatID, Type, CatName FROM Category WHERE Department = ? ORDER BY Type, CatName"
                    params = (self.current_department_id,)
                else:
                    query = "SELECT CatID, Type, CatName FROM Category WHERE Department = ? AND Type = ? ORDER BY CatName"
                    params = (self.current_department_id, type_filter)
                
                c.execute(query, params)
                categories = c.fetchall()
                
                self.category_table.setRowCount(len(categories))
                for i, (cat_id, cat_type, cat_name) in enumerate(categories):
                    self.category_table.setRowHeight(i, 40)
                    # Type
                    type_item = QTableWidgetItem(cat_type)
                    type_item.setData(Qt.UserRole, cat_id)
                    self.category_table.setItem(i, 0, type_item)
                    
                    # Category Name
                    name_item = QTableWidgetItem(cat_name)
                    self.category_table.setItem(i, 1, name_item)
                    
                    # Subcategories count
                    c.execute("SELECT COUNT(*) FROM SubCategory WHERE CategoryID = ?", (cat_id,))
                    subcat_count = c.fetchone()[0]
                    count_item = QTableWidgetItem(str(subcat_count))
                    count_item.setTextAlignment(Qt.AlignCenter)
                    self.category_table.setItem(i, 2, count_item)
                    
                    # Actions
                    action_combo = QComboBox()
                    action_combo.addItems(["Actions", "Edit", "Delete"])
                    action_combo.setFixedHeight(28)
                    action_combo.currentIndexChanged.connect(
                        lambda index, cid=cat_id, cb=action_combo: 
                        self.on_category_action(index, cid, cb)
                    )
                    self.category_table.setCellWidget(i, 3, action_combo)
                
               
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load categories: {str(e)}")

    def load_subcategories(self):
        """Load subcategories for the selected category"""
        if not self.current_category_id:
            self.subcategory_table.setRowCount(0)
            self.subcategory_count_label.setText("Subcategories: 0")
            self.usage_label.setText("Usage: 0 items")
            return
            
        try:
            with get_db_connection() as (conn, c):
                # Get subcategories
                c.execute("SELECT SubCatID, SubCategoryName FROM SubCategory WHERE CategoryID = ? ORDER BY SubCategoryName", 
                         (self.current_category_id,))
                subcategories = c.fetchall()
                
                self.subcategory_table.setRowCount(len(subcategories))
                total_usage = 0
                
                for i, (subcat_id, subcat_name) in enumerate(subcategories):
                    self.subcategory_table.setRowHeight(i, 40)
                    # Subcategory Name
                    name_item = QTableWidgetItem(subcat_name)
                    name_item.setData(Qt.UserRole, subcat_id)
                    self.subcategory_table.setItem(i, 0, name_item)
                    
                    # Usage count
                    c.execute("SELECT COUNT(*) FROM Supplies WHERE Subcategory = ?", (subcat_id,))
                    usage_count = c.fetchone()[0]
                    total_usage += usage_count
                    
                    usage_item = QTableWidgetItem(str(usage_count))
                    usage_item.setTextAlignment(Qt.AlignCenter)
                    self.subcategory_table.setItem(i, 1, usage_item)
                    
                    # Actions
                    action_combo = QComboBox()
                    action_combo.addItems(["Actions", "Edit", "Delete"])
                    action_combo.setFixedHeight(28)
                    action_combo.currentIndexChanged.connect(
                        lambda index, sid=subcat_id, cb=action_combo: 
                        self.on_subcategory_action(index, sid, cb)
                    )
                    self.subcategory_table.setCellWidget(i, 2, action_combo)
                
                # Update counts
                
                self.usage_label.setText(f"Usage: {total_usage} items")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load subcategories: {str(e)}")

    def on_category_action(self, index, cat_id, combo_box):
        """Handle category action selection"""
        if index == 0:
            return
        
        action = combo_box.currentText()
        combo_box.setCurrentIndex(0)
        
        if action == "Edit":
            self.edit_category(cat_id)
        elif action == "Delete":
            self.delete_category(cat_id)

    def on_subcategory_action(self, index, subcat_id, combo_box):
        """Handle subcategory action selection"""
        if index == 0:
            return
        
        action = combo_box.currentText()
        combo_box.setCurrentIndex(0)
        
        if action == "Edit":
            self.edit_subcategory(subcat_id)
        elif action == "Delete":
            self.delete_subcategory(subcat_id)

    # ========== CRUD METHODS ==========
    
    def add_department(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Department")
        dialog.setFixedSize(350, 180)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        
        name_label = QLabel("Department Name:")
        layout.addWidget(name_label)
        
        name_input = QLineEdit()
        name_input.setPlaceholderText("e.g., Administration, Academics, Sports")
        layout.addWidget(name_input)
        
        error_label = QLabel("")
        error_label.setStyleSheet("color: #e74c3c;")
        layout.addWidget(error_label)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        def save_department():
            dept_name = name_input.text().strip()
            if not dept_name:
                error_label.setText("Department name is required")
                return
            
            try:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT 1 FROM Department WHERE DeptName = ?", (dept_name,))
                    if c.fetchone():
                        error_label.setText("Department already exists")
                        return
                    
                    c.execute("INSERT INTO Department (DeptName) VALUES (?)", (dept_name,))
                    conn.commit()
                    dialog.accept()
                    self.load_departments()
                    QMessageBox.information(self, "Success", f"Department '{dept_name}' added")
                    
            except Exception as e:
                error_label.setText(f"Error: {str(e)}")
        
        save_btn.clicked.connect(save_department)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def edit_department(self):
        if not self.current_department_id:
            QMessageBox.warning(self, "No Selection", "Please select a department to edit")
            return
        
        # Find department name
        dept_name = ""
        for i in range(self.dept_buttons_layout.count()):
            widget = self.dept_buttons_layout.itemAt(i).widget()
            if widget and isinstance(widget, QPushButton) and widget.isChecked():
                dept_name = widget.text()
                break
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Department")
        dialog.setFixedSize(350, 180)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        
        name_label = QLabel("Department Name:")
        layout.addWidget(name_label)
        
        name_input = QLineEdit(dept_name)
        layout.addWidget(name_input)
        
        error_label = QLabel("")
        error_label.setStyleSheet("color: #e74c3c;")
        layout.addWidget(error_label)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        def save_edit():
            new_name = name_input.text().strip()
            if not new_name:
                error_label.setText("Department name is required")
                return
            
            try:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT 1 FROM Department WHERE DeptName = ? AND DeptID != ?", 
                            (new_name, self.current_department_id))
                    if c.fetchone():
                        error_label.setText("Department name already exists")
                        return
                    
                    c.execute("UPDATE Department SET DeptName = ? WHERE DeptID = ?", 
                            (new_name, self.current_department_id))
                    conn.commit()
                    dialog.accept()
                    self.load_departments()
                    QMessageBox.information(self, "Success", "Department updated")
                    
            except Exception as e:
                error_label.setText(f"Error: {str(e)}")
        
        save_btn.clicked.connect(save_edit)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def delete_department(self):
        if not self.current_department_id:
            QMessageBox.warning(self, "No Selection", "Please select a department to delete")
            return
        
        # Get department name
        dept_name = ""
        for i in range(self.dept_buttons_layout.count()):
            widget = self.dept_buttons_layout.itemAt(i).widget()
            if widget and isinstance(widget, QPushButton) and widget.isChecked():
                dept_name = widget.text()
                break
        
        # Check dependencies
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT COUNT(*) FROM Category WHERE Department = ?", (self.current_department_id,))
                category_count = c.fetchone()[0]
                
                if category_count > 0:
                    QMessageBox.warning(self, "Cannot Delete", 
                                      f"Cannot delete '{dept_name}' because it has {category_count} categories")
                    return
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to check dependencies: {str(e)}")
            return
        
        # Confirmation
        reply = QMessageBox.question(self, "Confirm Delete",
                                   f"Delete department '{dept_name}'?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    c.execute("DELETE FROM Department WHERE DeptID = ?", (self.current_department_id,))
                    conn.commit()
                    
                    # Reset selection and reload
                    self.current_department_id = None
                    self.current_category_id = None
                    self.load_departments()
                    self.clear_subcategories()
                    
                    QMessageBox.information(self, "Deleted", f"Department '{dept_name}' deleted")
                    
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete department: {str(e)}")

    def add_category(self):
        if not self.current_department_id:
            QMessageBox.warning(self, "No Department", "Please select a department first")
            return
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Category")
        dialog.setFixedSize(400, 250)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Type
        type_label = QLabel("Category Type:")
        layout.addWidget(type_label)
        
        type_combo = QComboBox()
        type_combo.addItems(["Supplies", "Expenses"])
        layout.addWidget(type_combo)
        
        # Name
        name_label = QLabel("Category Name:")
        layout.addWidget(name_label)
        
        name_input = QLineEdit()
        layout.addWidget(name_input)
        
        error_label = QLabel("")
        error_label.setStyleSheet("color: #e74c3c;")
        layout.addWidget(error_label)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        def save_category():
            cat_type = type_combo.currentText()
            cat_name = name_input.text().strip()
            
            if not cat_name:
                error_label.setText("Category name is required")
                return
            
            try:
                with get_db_connection() as (conn, c):
                    c.execute("""
                        SELECT 1 FROM Category 
                        WHERE Type = ? AND Department = ? AND CatName = ?
                    """, (cat_type, self.current_department_id, cat_name))
                    
                    if c.fetchone():
                        error_label.setText("Category already exists")
                        return
                    
                    c.execute("""
                        INSERT INTO Category (Type, Department, CatName) 
                        VALUES (?, ?, ?)
                    """, (cat_type, self.current_department_id, cat_name))
                    
                    conn.commit()
                    dialog.accept()
                    self.load_categories()
                    QMessageBox.information(self, "Success", f"Category '{cat_name}' added")
                    
            except Exception as e:
                error_label.setText(f"Error: {str(e)}")
        
        save_btn.clicked.connect(save_category)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def edit_category(self, cat_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT Type, CatName FROM Category WHERE CatID = ?", (cat_id,))
                result = c.fetchone()
                if not result:
                    return
                cat_type, cat_name = result
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load category: {str(e)}")
            return
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Category")
        dialog.setFixedSize(400, 250)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Type (display only)
        type_label = QLabel("Category Type:")
        layout.addWidget(type_label)
        
        type_display = QLabel(cat_type)
        layout.addWidget(type_display)
        
        # Name
        name_label = QLabel("Category Name:")
        layout.addWidget(name_label)
        
        name_input = QLineEdit(cat_name)
        layout.addWidget(name_input)
        
        error_label = QLabel("")
        error_label.setStyleSheet("color: #e74c3c;")
        layout.addWidget(error_label)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        def save_edit():
            new_name = name_input.text().strip()
            if not new_name:
                error_label.setText("Category name is required")
                return
            
            try:
                with get_db_connection() as (conn, c):
                    c.execute("""
                        SELECT 1 FROM Category 
                        WHERE Type = ? AND Department = ? AND CatName = ? AND CatID != ?
                    """, (cat_type, self.current_department_id, new_name, cat_id))
                    
                    if c.fetchone():
                        error_label.setText("Category name already exists")
                        return
                    
                    c.execute("UPDATE Category SET CatName = ? WHERE CatID = ?", (new_name, cat_id))
                    conn.commit()
                    dialog.accept()
                    self.load_categories()
                    QMessageBox.information(self, "Success", "Category updated")
                    
            except Exception as e:
                error_label.setText(f"Error: {str(e)}")
        
        save_btn.clicked.connect(save_edit)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def delete_category(self, cat_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT CatName FROM Category WHERE CatID = ?", (cat_id,))
                result = c.fetchone()
                if not result:
                    return
                cat_name = result[0]
                
                c.execute("SELECT COUNT(*) FROM SubCategory WHERE CategoryID = ?", (cat_id,))
                subcat_count = c.fetchone()[0]
                
                if subcat_count > 0:
                    QMessageBox.warning(self, "Cannot Delete", 
                                      f"Cannot delete '{cat_name}' because it has {subcat_count} subcategories")
                    return
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to check dependencies: {str(e)}")
            return
        
        reply = QMessageBox.question(self, "Confirm Delete",
                                   f"Delete category '{cat_name}'?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    c.execute("DELETE FROM Category WHERE CatID = ?", (cat_id,))
                    conn.commit()
                    self.load_categories()
                    self.clear_subcategories()
                    QMessageBox.information(self, "Deleted", f"Category '{cat_name}' deleted")
                    
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete category: {str(e)}")

    def add_subcategory(self):
        if not self.current_category_id:
            QMessageBox.warning(self, "No Category", "Please select a category first")
            return
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Subcategory")
        dialog.setFixedSize(350, 180)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        
        name_label = QLabel("Subcategory Name:")
        layout.addWidget(name_label)
        
        name_input = QLineEdit()
        layout.addWidget(name_input)
        
        error_label = QLabel("")
        error_label.setStyleSheet("color: #e74c3c;")
        layout.addWidget(error_label)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        def save_subcategory():
            subcat_name = name_input.text().strip()
            if not subcat_name:
                error_label.setText("Subcategory name is required")
                return
            
            try:
                with get_db_connection() as (conn, c):
                    c.execute("""
                        SELECT 1 FROM SubCategory 
                        WHERE CategoryID = ? AND SubCategoryName = ?
                    """, (self.current_category_id, subcat_name))
                    
                    if c.fetchone():
                        error_label.setText("Subcategory already exists")
                        return
                    
                    c.execute("""
                        INSERT INTO SubCategory (CategoryID, SubCategoryName) 
                        VALUES (?, ?)
                    """, (self.current_category_id, subcat_name))
                    
                    conn.commit()
                    dialog.accept()
                    self.load_subcategories()
                    QMessageBox.information(self, "Success", f"Subcategory '{subcat_name}' added")
                    
            except Exception as e:
                error_label.setText(f"Error: {str(e)}")
        
        save_btn.clicked.connect(save_subcategory)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def edit_subcategory(self, subcat_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT SubCategoryName FROM SubCategory WHERE SubCatID = ?", (subcat_id,))
                result = c.fetchone()
                if not result:
                    return
                subcat_name = result[0]
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load subcategory: {str(e)}")
            return
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Subcategory")
        dialog.setFixedSize(350, 180)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        
        name_label = QLabel("Subcategory Name:")
        layout.addWidget(name_label)
        
        name_input = QLineEdit(subcat_name)
        layout.addWidget(name_input)
        
        error_label = QLabel("")
        error_label.setStyleSheet("color: #e74c3c;")
        layout.addWidget(error_label)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        save_btn = QPushButton("Save")
        cancel_btn = QPushButton("Cancel")
        
        btn_layout.addWidget(save_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
        
        def save_edit():
            new_name = name_input.text().strip()
            if not new_name:
                error_label.setText("Subcategory name is required")
                return
            
            try:
                with get_db_connection() as (conn, c):
                    c.execute("""
                        SELECT 1 FROM SubCategory 
                        WHERE CategoryID = ? AND SubCategoryName = ? AND SubCatID != ?
                    """, (self.current_category_id, new_name, subcat_id))
                    
                    if c.fetchone():
                        error_label.setText("Subcategory name already exists")
                        return
                    
                    c.execute("UPDATE SubCategory SET SubCategoryName = ? WHERE SubCatID = ?", 
                            (new_name, subcat_id))
                    conn.commit()
                    dialog.accept()
                    self.load_subcategories()
                    QMessageBox.information(self, "Success", "Subcategory updated")
                    
            except Exception as e:
                error_label.setText(f"Error: {str(e)}")
        
        save_btn.clicked.connect(save_edit)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def delete_subcategory(self, subcat_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT SubCategoryName FROM SubCategory WHERE SubCatID = ?", (subcat_id,))
                result = c.fetchone()
                if not result:
                    return
                subcat_name = result[0]
                
                c.execute("SELECT COUNT(*) FROM Supplies WHERE Subcategory = ?", (subcat_id,))
                usage_count = c.fetchone()[0]
                
                if usage_count > 0:
                    QMessageBox.warning(self, "Cannot Delete", 
                                      f"Cannot delete '{subcat_name}' because it is used in {usage_count} items")
                    return
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to check dependencies: {str(e)}")
            return
        
        reply = QMessageBox.question(self, "Confirm Delete",
                                   f"Delete subcategory '{subcat_name}'?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    c.execute("DELETE FROM SubCategory WHERE SubCatID = ?", (subcat_id,))
                    conn.commit()
                    self.load_subcategories()
                    QMessageBox.information(self, "Deleted", f"Subcategory '{subcat_name}' deleted")
                    
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to delete subcategory: {str(e)}")

### end of categories management ###
### Report###
class Reports(QWidget):
    def __init__(self):
        super().__init__()
        self.finance_service = FinanceService()
        self.report_generator = ReportGeneratorEngine()
        self.current_term_data = {}
        self.init_ui()
        self.load_terms()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        # Header
        header_lbl = QLabel("Financial Reports")
        header_lbl.setStyleSheet("font-size: 24px; font-weight: bold; color: #2D3748;")
        layout.addWidget(header_lbl)

        # Controls Container
        controls_frame = QFrame()
        controls_frame.setStyleSheet("background-color: white; border-radius: 8px; border: 1px solid #E2E8F0;")
        controls_layout = QHBoxLayout(controls_frame)
        controls_layout.setContentsMargins(20, 20, 20, 20)
        
        # Term Selection
        term_lbl = QLabel("Select Term:")
        term_lbl.setStyleSheet("font-size: 14px; font-weight: bold; color: #4A5568;")
        self.term_selector = QComboBox()
        self.term_selector.setFixedWidth(200)
        self.term_selector.setStyleSheet("""
            QComboBox {
                padding: 8px;
                border: 1px solid #E2E8F0;
                border-radius: 5px;
                font-size: 14px;
            }
        """)
        self.term_selector.currentIndexChanged.connect(self.load_term_data)
        
        controls_layout.addWidget(term_lbl)
        controls_layout.addWidget(self.term_selector)
        controls_layout.addStretch()
        
        layout.addWidget(controls_frame)

        # Summary Cards Area
        self.cards_layout = QHBoxLayout()
        self.lbl_fees = self.create_summary_card("Fees Collected", "0.00", "#3182CE")
        self.lbl_expenses = self.create_summary_card("Total Expenses", "0.00", "#E53E3E")
        self.lbl_supplies = self.create_summary_card("Supplies Cost", "0.00", "#805AD5")
        
        self.cards_layout.addWidget(self.lbl_fees)
        self.cards_layout.addWidget(self.lbl_expenses)
        self.cards_layout.addWidget(self.lbl_supplies)
        layout.addLayout(self.cards_layout)

        # Actions Area
        actions_frame = QFrame()
        actions_frame.setStyleSheet("background-color: white; border-radius: 8px; border: 1px solid #E2E8F0;")
        actions_layout = QVBoxLayout(actions_frame)
        actions_layout.setContentsMargins(20, 20, 20, 20)
        
        actions_title = QLabel("Export Options")
        actions_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #2D3748; margin-bottom: 10px;")
        actions_layout.addWidget(actions_title)
        
        btn_layout = QHBoxLayout()
        
        btn_pdf = button("Generate PDF Report","secondary")
        btn_pdf.setCursor(Qt.PointingHandCursor)
        btn_pdf.clicked.connect(self.export_pdf)
        
        btn_excel = button("Generate Excel Report","Success","Lg")
        btn_excel.setCursor(Qt.PointingHandCursor)
        btn_excel.clicked.connect(self.export_excel)
        
        btn_layout.addWidget(btn_pdf)
        btn_layout.addWidget(btn_excel)
        btn_layout.addStretch()
        
        actions_layout.addLayout(btn_layout)
        layout.addWidget(actions_frame)
        
        # Disclaimer
        disclaimer_lbl = QLabel("Note: Reports reflect only system-recorded data. External income (grants, donations) is excluded.")
        disclaimer_lbl.setStyleSheet("color: #718096; font-style: italic; margin-top: 20px;")
        layout.addWidget(disclaimer_lbl)
        
        layout.addStretch()

    def create_summary_card(self, title, value, color):
        frame = QFrame()
        frame.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 8px;
                border-left: 5px solid {color};
                border-top: 1px solid #E2E8F0;
                border-right: 1px solid #E2E8F0;
                border-bottom: 1px solid #E2E8F0;
                min-height: 100px;
            }}
        """)
        l = QVBoxLayout(frame)
        t = QLabel(title)
        t.setStyleSheet("color: #718096; font-size: 14px; border: none;")
        v = QLabel(value)
        v.setStyleSheet("color: #2D3748; font-size: 24px; font-weight: bold; border: none;")
        v.setAlignment(Qt.AlignLeft)
        l.addWidget(t)
        l.addWidget(v)
        frame.value_label = v
        return frame

    def load_terms(self):
        self.term_selector.clear()
        try:
            terms = self.finance_service.get_all_terms()
            for term in terms:
                self.term_selector.addItem(f"{term['TermName']} - {term['Year']}", term['TermID'])
        except Exception as e:
            QMessageBox.warning(f"Error loading terms: {e}")

    def load_term_data(self):
        term_id = self.term_selector.currentData()
        if not term_id:
            return

        try:
            term_dates = self.finance_service.get_term_dates(term_id)
            if not term_dates:
                return
            start_date = term_dates['StartDate']
            end_date = term_dates['EndDate']

            fees_summary = self.finance_service.get_term_fees_summary(term_id)
            total_expenses = self.finance_service.get_term_expenses(start_date, end_date)
            total_supplies = self.finance_service.get_term_supplies(start_date, end_date)

            self.lbl_fees.value_label.setText(f"{fees_summary['collected']:,.2f}")
            self.lbl_expenses.value_label.setText(f"{total_expenses:,.2f}")
            self.lbl_supplies.value_label.setText(f"{total_supplies:,.2f}")

            # Store basic info for export
            self.current_term_data = {
                'term_name': self.term_selector.currentText(),
                'start_date': start_date,
                'end_date': end_date,
                'fees_summary': fees_summary,
                'expenses_total': total_expenses,
                'supplies_total': total_supplies,
                'term_id': term_id
            }

        except Exception as e:
            QMessageBox.warning(f"Error loading data: {e}")

    def fetch_detailed_data(self):
        """Helper to fetch all detailed data needed for reports"""
        if not self.current_term_data:
            return None
            
        term_id = self.current_term_data['term_id']
        start_date = self.current_term_data['start_date']
        end_date = self.current_term_data['end_date']
        
        return {
            'fees_breakdown': self.finance_service.get_term_fees_breakdown_by_class(term_id),
            'supplies_breakdown': self.finance_service.get_supplies_by_category(start_date, end_date),
            'supplies_details': self.finance_service.get_supplies_details(start_date, end_date),
            'expenses_breakdown': self.finance_service.get_expenses_by_category_with_percentage(start_date, end_date),
            'expenses_details': self.finance_service.get_expenses_details(start_date, end_date)
        }

    def export_pdf(self):
        if not self.current_term_data:
            return
        
        filename, _ = QFileDialog.getSaveFileName(self, "Save PDF Report", f"Financial_Report_{self.current_term_data['term_name']}.pdf", "PDF Files (*.pdf)")
        if filename:
            try:
                details = self.fetch_detailed_data()
                self.report_generator.generate_pdf(
                    filename,
                    self.current_term_data['term_name'],
                    self.current_term_data['start_date'],
                    self.current_term_data['end_date'],
                    self.current_term_data['fees_summary'],
                    details['fees_breakdown'],
                    self.current_term_data['supplies_total'],
                    details['supplies_breakdown'],
                    details['supplies_details'],
                    self.current_term_data['expenses_total'],
                    details['expenses_breakdown'],
                    details['expenses_details']
                )
                QMessageBox.information(self, "Success", "PDF Report generated successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to generate PDF: {str(e)}")

    def export_excel(self):
        if not self.current_term_data:
            return

        filename, _ = QFileDialog.getSaveFileName(self, "Save Excel Report", f"Financial_Report_{self.current_term_data['term_name']}.xlsx", "Excel Files (*.xlsx)")
        if filename:
            try:
                details = self.fetch_detailed_data()
                self.report_generator.generate_excel(
                    filename,
                    self.current_term_data['term_name'],
                    self.current_term_data['start_date'],
                    self.current_term_data['end_date'],
                    self.current_term_data['fees_summary'],
                    details['fees_breakdown'],
                    self.current_term_data['supplies_total'],
                    details['supplies_breakdown'],
                    details['supplies_details'],
                    self.current_term_data['expenses_total'],
                    details['expenses_breakdown'],
                    details['expenses_details']
                )
                QMessageBox.information(self, "Success", "Excel Report generated successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to generate Excel: {str(e)}")

class FinancialReportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Term Financial Summary Report")
        self.setFixedSize(600, 500)
        self.finance_service = FinanceService()
        self.report_generator = ReportGeneratorEngine()
        self.current_term_data = {}
        self.init_ui()
        self.load_terms()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        # Header
        header_lbl = QLabel("Generate Financial Report")
        header_lbl.setStyleSheet("font-size: 20px; font-weight: bold; color: #2D3748;")
        layout.addWidget(header_lbl)

        # Term Selection
        term_layout = QHBoxLayout()
        term_lbl = QLabel("Select Term:")
        self.term_selector = QComboBox()
        self.term_selector.currentIndexChanged.connect(self.load_term_data)
        term_layout.addWidget(term_lbl)
        term_layout.addWidget(self.term_selector)
        layout.addLayout(term_layout)

        # Disclaimer
        disclaimer_frame = QFrame()
        disclaimer_frame.setStyleSheet("""
            QFrame {
                background-color: #FFF5F5;
                border: 1px solid #FC8181;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        disclaimer_layout = QVBoxLayout(disclaimer_frame)
        disclaimer_lbl = QLabel("<b>IMPORTANT:</b> This report does not include donations, grants, rentals, or other external income sources. It reflects only the data recorded in this system.")
        disclaimer_lbl.setWordWrap(True)
        disclaimer_lbl.setStyleSheet("color: #C53030; font-size: 12px;")
        disclaimer_layout.addWidget(disclaimer_lbl)
        layout.addWidget(disclaimer_frame)

        # Summary Cards
        self.cards_layout = QHBoxLayout()
        self.lbl_fees = self.create_summary_label("Fees Collected", "0.00")
        self.lbl_expenses = self.create_summary_label("Total Expenses", "0.00")
        self.lbl_supplies = self.create_summary_label("Supplies Cost", "0.00")
        
        self.cards_layout.addWidget(self.lbl_fees)
        self.cards_layout.addWidget(self.lbl_expenses)
        self.cards_layout.addWidget(self.lbl_supplies)
        layout.addLayout(self.cards_layout)

        # Export Buttons
        btn_layout = QHBoxLayout()
        btn_pdf = QPushButton("Export PDF")
        btn_pdf.setStyleSheet("background-color: #E53E3E; color: white; padding: 10px; font-weight: bold;")
        btn_pdf.clicked.connect(self.export_pdf)
        
        btn_excel = QPushButton("Export Excel")
        btn_excel.setStyleSheet("background-color: #38A169; color: white; padding: 10px; font-weight: bold;")
        btn_excel.clicked.connect(self.export_excel)
        
        btn_layout.addWidget(btn_pdf)
        btn_layout.addWidget(btn_excel)
        layout.addLayout(btn_layout)
        
        layout.addStretch()

    def create_summary_label(self, title, value):
        frame = QFrame()
        frame.setStyleSheet("background-color: #F7FAFC; border-radius: 5px; border: 1px solid #E2E8F0;")
        l = QVBoxLayout(frame)
        t = QLabel(title)
        t.setStyleSheet("color: #718096; font-size: 12px;")
        v = QLabel(value)
        v.setStyleSheet("color: #2D3748; font-size: 18px; font-weight: bold;")
        v.setAlignment(Qt.AlignCenter)
        l.addWidget(t, alignment=Qt.AlignCenter)
        l.addWidget(v, alignment=Qt.AlignCenter)
        frame.value_label = v
        return frame

    def load_terms(self):
        self.term_selector.clear()
        try:
            terms = self.finance_service.get_all_terms()
            for term in terms:
                self.term_selector.addItem(f"{term['TermName']} ({term['Year']})", term['TermID'])
        except Exception as e:
            QMessageBox.warning(f"Error loading terms: {e}")

    def load_term_data(self):
        term_id = self.term_selector.currentData()
        if not term_id:
            return

        try:
            term_dates = self.finance_service.get_term_dates(term_id)
            if not term_dates:
                return
            start_date = term_dates['StartDate']
            end_date = term_dates['EndDate']

            fees_summary = self.finance_service.get_term_fees_summary(term_id)
            total_expenses = self.finance_service.get_term_expenses(start_date, end_date)
            total_supplies = self.finance_service.get_term_supplies(start_date, end_date)

            self.lbl_fees.value_label.setText(f"{fees_summary['collected']:,.2f}")
            self.lbl_expenses.value_label.setText(f"{total_expenses:,.2f}")
            self.lbl_supplies.value_label.setText(f"{total_supplies:,.2f}")

            # Store for export
            self.current_term_data = {
                'term_name': self.term_selector.currentText(),
                'start_date': start_date,
                'end_date': end_date,
                'fees': fees_summary,
                'expenses': total_expenses,
                'supplies': total_supplies,
                'term_id': term_id
            }

        except Exception as e:
            QMessageBox.warning(f"Error loading data: {e}")

    def export_pdf(self):
        if not self.current_term_data:
            return
        
        filename, _ = QFileDialog.getSaveFileName(self, "Save PDF Report", f"Financial_Report_{self.current_term_data['term_name']}.pdf", "PDF Files (*.pdf)")
        if filename:
            try:
                self.report_generator.generate_pdf(
                    filename,
                    self.current_term_data['term_name'],
                    self.current_term_data['start_date'],
                    self.current_term_data['end_date'],
                    self.current_term_data['fees'],
                    self.current_term_data['expenses'],
                    self.current_term_data['supplies']
                )
                QMessageBox.information(self, "Success", "PDF Report generated successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to generate PDF: {str(e)}")

    def export_excel(self):
        if not self.current_term_data:
            return

        filename, _ = QFileDialog.getSaveFileName(self, "Save Excel Report", f"Financial_Report_{self.current_term_data['term_name']}.xlsx", "Excel Files (*.xlsx)")
        if filename:
            try:
                # Fetch breakdown data for Excel
                fees_breakdown = self.finance_service.get_outstanding_fees_by_class(self.current_term_data['term_id'], limit=100)
                expenses_breakdown = self.finance_service.get_expenses_by_category(self.current_term_data['start_date'], self.current_term_data['end_date'], limit=100)

                self.report_generator.generate_excel(
                    filename,
                    self.current_term_data['term_name'],
                    self.current_term_data['start_date'],
                    self.current_term_data['end_date'],
                    self.current_term_data['fees'],
                    self.current_term_data['expenses'],
                    self.current_term_data['supplies'],
                    fees_breakdown,
                    expenses_breakdown
                )
                QMessageBox.information(self, "Success", "Excel Report generated successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to generate Excel: {str(e)}")

### End of Report###
class About(QWidget):
    def __init__(self):
        super().__init__()

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(24, 24, 24, 24)

        # ---- Scroll Area (important for long content) ----
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.NoFrame)

        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setSpacing(14)

        # ---- Title ----
        title = QLabel("About Pinnacle SF")
        title.setAlignment(Qt.AlignLeft)
        title.setProperty("role", "title")  # styled via QSS if needed
        content_layout.addWidget(title)

        # ---- Section helper ----
        def section_title(text):
            lbl = QLabel(text)
            lbl.setProperty("role", "section")
            return lbl

        def paragraph(text):
            lbl = QLabel(text)
            lbl.setWordWrap(True)
            lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            return lbl

        # ---- System Overview ----
        content_layout.addWidget(section_title("System Overview"))
        content_layout.addWidget(paragraph(
            "Pinnacle SF is a school finance management system designed to improve "
            "clarity, transparency, and accountability in the handling of school "
            "financial data. It focuses on monitoring how funds are collected, used, "
            "and reported within clearly defined boundaries."
        ))

        # ---- Purpose & Motivation ----
        content_layout.addWidget(section_title("Purpose & Motivation"))
        content_layout.addWidget(paragraph(
            "The system was developed to reduce uncertainty and mistrust surrounding "
            "the management of school funds. By clearly recording financial "
            "transactions, Pinnacle SF helps ensure that financial activities are "
            "transparent and verifiable."
        ))

        # ---- Scope of Management ----
        content_layout.addWidget(section_title("Scope of Management"))
        content_layout.addWidget(paragraph(
            "Pinnacle SF manages a specific portion of school financial operations, "
            "including school fees, supplies, and operational expenses."
        ))
        content_layout.addWidget(paragraph(
            "Important: This system does not represent the complete financial position "
            "of a school. It reflects only the data managed within the system."
        ))

        # ---- Fees ----
        content_layout.addWidget(section_title("Fees Management"))
        content_layout.addWidget(paragraph(
            "The system provides clear term-based insights into expected fees, collected "
            "fees, and outstanding balances. Percentages are used to help finance "
            "officers quickly understand collection performance."
        ))

        # ---- Supplies & Expenses ----
        content_layout.addWidget(section_title("Supplies & Expenses Tracking"))
        content_layout.addWidget(paragraph(
            "Supplies and expenses are recorded using categories. This structure makes "
            "it easier to understand how funds were used, analyze spending patterns, "
            "and support internal reviews."
        ))

        # ---- Reporting ----
        content_layout.addWidget(section_title("Reporting & Decision Support"))
        content_layout.addWidget(paragraph(
            "All reports are term-based and designed to be easy to read and interpret. "
            "The system emphasizes summarized insights supported by categorized data. "
            "Reports can be exported for use in PDF or Excel formats."
        ))

        # ---- Target Users ----
        content_layout.addWidget(section_title("Target Users"))
        content_layout.addWidget(paragraph(
            "Pinnacle SF is primarily designed for finance officers responsible for "
            "managing and reviewing school financial records."
        ))

        # ---- Limitations ----
        content_layout.addWidget(section_title("Limitations & Transparency"))
        content_layout.addWidget(paragraph(
            "Non-fees income is not currently included in the system. This limitation "
            "is intentional and clearly defined to prevent misinterpretation of "
            "financial reports. Future improvements may be considered, but are not "
            "guaranteed."
        ))

        # ---- System Info ----
        content_layout.addWidget(section_title("System Information"))
        content_layout.addWidget(paragraph(
            "System Name: Pinnacle SF\n"
            "Developer: Shema Nathan"
        ))

        content_layout.addStretch()

        scroll.setWidget(content)
        main_layout.addWidget(scroll)

def button(text, button_type="primary", size="md", icon_path=None, icon_toolTip = None, icon_size=24 ):
    """
    Create a Bootstrap-inspired button with optional icons.

    :param text: The text to display on the button.
    :param button_type: The type of the button (primary, success, danger, etc.).
    :param size: The size of the button (sm, md, lg).
    :param icon_path: The file path to the icon (optional).
    :param icon_size: The size of the icon (default: 24px).
    :return: QPushButton styled as a Bootstrap button.
    """
    button = QPushButton(text)
    button.setToolTip(icon_toolTip)
    # Button Styles Dictionary
    styles = {
        "primary": "background-color: #007bff; color: white; border: 1px solid #007bff;",
        "secondary": "background-color: #6c757d; color: white; border: 1px solid #6c757d;",
        "success": "background-color: #28a745; color: white; border: 1px solid #28a745;",
        "danger": "background-color: #dc3545; color: white; border: 1px solid #dc3545;",
        "warning": "background-color: #ffc107; color: black; border: 1px solid #ffc107;",
        "info": "background-color: #17a2b8; color: white; border: 1px solid #17a2b8;",
        "light": "background-color: #f8f9fa; color: black; border: 1px solid #f8f9fa;",
        "dark": "background-color: #343a40; color: white; border: 1px solid #343a40;",
    }

    # Button Sizes Dictionary
    sizes = {
        "sm": "padding: 4px 10px; font-size: 12px;",
        "md": "padding: 6px 12px; font-size: 14px;",
        "lg": "padding: 10px 16px; font-size: 18px;",
    }

    # Apply Styles
    button.setStyleSheet(f"""
        QPushButton {{
            border-radius: 5px;
            font-family: 'Poppins', sans-serif;
            font-weight: bold;
            {styles.get(button_type, styles['primary'])} /* Default to primary if type is invalid */
            {sizes.get(size, sizes['md'])} /* Default to medium if size is invalid */
        }}
        QPushButton:hover {{
            opacity: 0.9; /* Slight fade effect on hover */
        }}
    """)

    # Optional: Set Icon
    if icon_path:
        button.setIcon(QIcon(icon_path))
        button.setIconSize(QSize(icon_size, icon_size))

    # Optional: Set Cursor to Hand Pointer
    button.setCursor(Qt.PointingHandCursor)

    return button

def DialogStyles():
    dialog = """
    QDialog {
        background-color: #F9F9F9;
        border-radius: 12px;
        font-family: 'Segoe UI', 'Lato', sans-serif;
        font-size: 14px;
        color: #2E2E2E;
    }

    QLabel {
        font-weight: 500;
        margin-bottom: 4px;
    }

    QLineEdit, QComboBox, QDateEdit, QSpinBox {
        background-color: #FFFFFF;
        border: 1px solid #CCCCCC;
        border-radius: 6px;
        padding: 6px 10px;
        font-size: 14px;
        color: #2E2E2E;
    }

    QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QSpinBox:focus {
        border: 1.5px solid #007BFF;
        background-color: #FFFFFF;
    }

    QPushButton {
        background-color: #007BFF;
        color: white;
        padding: 8px 16px;
        border: none;
        border-radius: 6px;
        font-weight: 600;
    }

    QPushButton:hover {
        background-color: #0056b3;
    }

    QPushButton:disabled {
        background-color: #cccccc;
        color: #666666;
    }

    QCheckBox {
        spacing: 6px;
    }

    QCheckBox::indicator {
        width: 16px;
        height: 16px;
    }

    QCheckBox::indicator:checked {
        background-color: #007BFF;
        border: 1px solid #007BFF;
    }

    QGroupBox {
        border: 1px solid #E0E0E0;
        border-radius: 8px;
        margin-top: 10px;
        padding: 10px;
    }

    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 3px 0 3px;
        font-weight: bold;
    }
    """
    return dialog

def FormStyles():
    form_css = """
       *{
        background-color: white;
        }
        QLineEdit, QSpinBox {
        padding: 6px 10px;
        border: 1px solid #ced4da;
        border-radius: 3px;
        height: 30px;
        font: 12pt;
        } 
        QDateEdit {
        border: 1px solid #ced4da; /* Light gray border */
        border-radius: 8px; /* Rounded corners */
        background-color: #ffffff; /* White background */
        padding: 5px 10px; /* Padding inside the input */
        color: #495057; /* Text color */
        font-size: 14px; /* Font size */
        height: 30px; /* Input height */
        }
    
        QDateEdit:hover {
            border: 1px solid #80bdff; /* Blue border on hover */
            background-color: #f8f9fa; /* Light gray background */
        }
    
        QDateEdit:focus {
            border: 1px solid #007bff; /* Bootstrap blue for focus */
            outline: none; /* Remove outline */
        }
    
        QDateEdit::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 20px; /* Dropdown width */
            border-left: 1px solid #ced4da; /* Divider for dropdown */
        }
    
        QDateEdit::down-arrow {
            image: url('icons/calendar.png'); /* Replace with your icon path */
            width: 14px;
            height: 14px;
        }
    
        QDateEdit::down-arrow:hover {
            image: url('icons/calendar_hover_16px.png'); /* Hover icon */
        }
    
        QDateEdit::down-arrow:pressed {
            image: url('icons/calendar_pressed_16px.png'); /* Pressed icon */
        }
    
        QCalendarWidget {
            border: 1px solid #dee2e6; /* Border for calendar popup */
            border-radius: 8px;
            background-color: #ffffff; /* White background */
            color: #212529; /* Text color */
            font-size: 14px;
        }
    
        QCalendarWidget QToolButton {
            color: #495057; /* Month and Year text color */
            background-color: #f8f9fa; /* Light gray button background */
            border: none; /* No border */
            border-radius: 5px;
            padding: 5px;
            font-weight: bold;
        }
    
        QCalendarWidget QToolButton:hover {
            background-color: #e2e6ea; /* Hover background color */
        }
    
        QCalendarWidget QToolButton:pressed {
            background-color: #d4d8dc; /* Pressed background color */
        }
    
        QCalendarWidget QTableView {
            border: none;
            color: #495057; /* Day numbers color */
            gridline-color: #dee2e6; /* Gridline color */
        }
    
        QCalendarWidget QTableView:selected {
            background-color: #007bff; /* Selected day background */
            color: #ffffff; /* Selected day text */
        }
        QComboBox {
        border: 1px solid #ced4da;
        border-radius: 4px;
        padding: 8px 12px;
        background-color: #f8f9fa;
        font-size: 16px;
        color: #212529;
        }

        QComboBox:hover {
            border-color: #86b7fe;
        }

        QComboBox:focus {
            border-color: #0d6efd;
            outline: 2px solid #0d6efd;
            outline-offset: 1px;
        }

        QComboBox QAbstractItemView {
            border: 1px solid #ced4da;
            background-color: #fff;
            selection-background-color: #0d6efd;
            selection-color: white;
            padding: 6px 12px;
            font-size: 16px;
        }

        QComboBox::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 25px;
            border-left: 1px solid #ced4da;
            background-color: #e9ecef;
        }

        QFrame{
        border-radius: 5px;
        }
        QPushButton {
        border-radius: 5px;
        font-family: 'Poppins', sans-serif;
        font-weight: bold;
        background-color: #007bff;
        color: white; 
        border: 1px solid #007bff;
        padding: 6px 12px;
        font-size: 14px;
        text-align: center;
        }

        QPushButton:hover {
            background-color: #0056b3;
            border: 1px solid #0056b3;
        }

        QPushButton:pressed {
            background-color: #004080;
            border: 1px solid #004080;
        }

        QPushButton:focus {
            outline: 2px solid #80bdff;
        }
    """
    return form_css

class FinanceService:
    def get_term_dates(self, term_id):
        with get_db_connection() as (conn, c):
            c.execute("SELECT StartDate, EndDate FROM Term WHERE TermID = ?", (term_id,))
            return c.fetchone()

    def get_term_fees_summary(self, term_id):
        """
        Returns a dictionary with:
        - expected: Total fees expected for the term
        - collected: Total fees collected for the term
        - outstanding: expected - collected
        """
        with get_db_connection() as (conn, c):
            # Total Fees Expected
            c.execute("""
                SELECT SUM(
                    (f.Amount * (st.PayRate / 100.0))
                ) as TotalExpected
                FROM StudentTerm st
                JOIN Fees f ON st.ClassID = f.Class AND st.TermID = f.Term
                WHERE st.TermID = ?
            """, (term_id,))
            total_expected = c.fetchone()['TotalExpected'] or 0.0

            # Total Fees Collected
            c.execute("""
                SELECT SUM(AmountPaid) 
                FROM FeesPayment fp
                JOIN StudentTerm st ON fp.StudentTermID = st.StudentTermID
                WHERE st.TermID = ?
            """, (term_id,))
            total_collected = c.fetchone()[0] or 0.0
            
            outstanding = max(0, total_expected - total_collected)
            
            return {
                "expected": total_expected,
                "collected": total_collected,
                "outstanding": outstanding
            }

    def get_term_expenses(self, start_date, end_date):
        """Returns total expenses between start and end date."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT SUM(ExpAmount) 
                FROM Expenses 
                WHERE ExpDate BETWEEN ? AND ?
            """, (start_date, end_date))
            return c.fetchone()[0] or 0.0

    def get_term_supplies(self, start_date, end_date):
        """Returns total supplies cost (Type='In') between start and end date."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT SUM(Price) 
                FROM SuppliesInventory 
                WHERE Type = 'In' AND Date BETWEEN ? AND ?
            """, (start_date, end_date))
            return c.fetchone()[0] or 0.0

    def get_expenses_by_category(self, start_date, end_date, limit=5):
        """Returns list of (CategoryName, Amount) tuples."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT c.CatName, SUM(e.ExpAmount)
                FROM Expenses e
                JOIN SubCategory sc ON e.ExpSubCategory = sc.SubCatID
                JOIN Category c ON sc.CategoryID = c.CatID
                WHERE e.ExpDate BETWEEN ? AND ?
                GROUP BY c.CatName
                ORDER BY SUM(e.ExpAmount) DESC
                LIMIT ?
            """, (start_date, end_date, limit))
            data = c.fetchall()
            return [(row[0], row[1]) for row in data]

    def get_expenses_by_category_with_percentage(self, start_date, end_date):
        """Returns list of (CategoryName, Amount, Percentage) tuples."""
        with get_db_connection() as (conn, c):
            # First get total for percentage calculation
            c.execute("""
                SELECT SUM(ExpAmount) 
                FROM Expenses 
                WHERE ExpDate BETWEEN ? AND ?
            """, (start_date, end_date))
            total_expenses = c.fetchone()[0] or 0.0

            if total_expenses == 0:
                return []

            c.execute("""
                SELECT c.CatName, SUM(e.ExpAmount) as Amount
                FROM Expenses e
                JOIN SubCategory sc ON e.ExpSubCategory = sc.SubCatID
                JOIN Category c ON sc.CategoryID = c.CatID
                WHERE e.ExpDate BETWEEN ? AND ?
                GROUP BY c.CatName
                ORDER BY Amount DESC
            """, (start_date, end_date))
            data = c.fetchall()
            return [(row['CatName'], row['Amount'], (row['Amount']/total_expenses)*100) for row in data]

    def get_outstanding_fees_by_class(self, term_id, limit=10):
        """Returns list of dictionaries with class_name, expected, collected, balance, percentage."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT cl.ClassName, 
                       SUM(f.Amount * (st.PayRate / 100.0)) as Expected,
                       (SELECT SUM(AmountPaid) 
                        FROM FeesPayment fp 
                        WHERE fp.StudentTermID IN (
                            SELECT StudentTermID FROM StudentTerm WHERE ClassID = cl.ClassID AND TermID = ?
                        )
                       ) as Collected
                FROM StudentTerm st
                JOIN Class cl ON st.ClassID = cl.ClassID
                JOIN Fees f ON st.ClassID = f.Class AND st.TermID = f.Term
                WHERE st.TermID = ?
                GROUP BY cl.ClassName
                ORDER BY (Expected - COALESCE(Collected, 0)) DESC
                LIMIT ?
            """, (term_id, term_id, limit))
            
            rows = c.fetchall()
            result = []
            for row in rows:
                expected = row['Expected'] or 0
                collected = row['Collected'] or 0
                balance = expected - collected
                if balance > 0:
                    percent = (balance / expected * 100) if expected > 0 else 0
                    result.append({
                        "class_name": row['ClassName'],
                        "expected": expected,
                        "collected": collected,
                        "balance": balance,
                        "percentage": percent
                    })
            return result

    def get_term_fees_breakdown_by_class(self, term_id):
        """Returns list of dictionaries with class_name, expected, collected, balance, percentage."""
        return self.get_outstanding_fees_by_class(term_id, limit=-1)

    def get_top_supplies(self, start_date, end_date, limit=5):
        """Returns list of (Name, Quantity, Cost) tuples."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT s.Name, SUM(si.Quantity), SUM(si.Price)
                FROM SuppliesInventory si
                JOIN Supplies s ON si.Supplies_ID = s.SuppliesID
                WHERE si.Type = 'In' AND si.Date BETWEEN ? AND ?
                GROUP BY s.Name
                ORDER BY SUM(si.Price) DESC
                LIMIT ?
            """, (start_date, end_date, limit))
            
            data = c.fetchall()
            return [(row[0], row[1], row[2]) for row in data]

    def get_supplies_by_category(self, start_date, end_date):
        """Returns list of (CategoryName, Amount, Percentage) tuples."""
        with get_db_connection() as (conn, c):
            # First get total for percentage calculation
            c.execute("""
                SELECT SUM(si.Price)
                FROM SuppliesInventory si
                WHERE si.Type = 'In' AND si.Date BETWEEN ? AND ?
            """, (start_date, end_date))
            total_supplies = c.fetchone()[0] or 0.0

            if total_supplies == 0:
                return []

            c.execute("""
                SELECT c.CatName, SUM(si.Price) as Amount
                FROM SuppliesInventory si
                JOIN Supplies s ON si.Supplies_ID = s.SuppliesID
                JOIN SubCategory sc ON s.Subcategory = sc.SubCatID
                JOIN Category c ON sc.CategoryID = c.CatID
                WHERE si.Type = 'In' AND si.Date BETWEEN ? AND ?
                GROUP BY c.CatName
                ORDER BY Amount DESC
            """, (start_date, end_date))
            
            data = c.fetchall()
            return [(row['CatName'], row['Amount'], (row['Amount']/total_supplies)*100) for row in data]

    def get_supplies_details(self, start_date, end_date):
        """Returns list of dictionaries for detailed supplies report."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT si.Date, c.CatName, s.Name, si.Quantity, si.Price
                FROM SuppliesInventory si
                JOIN Supplies s ON si.Supplies_ID = s.SuppliesID
                JOIN SubCategory sc ON s.Subcategory = sc.SubCatID
                JOIN Category c ON sc.CategoryID = c.CatID
                WHERE si.Type = 'In' AND si.Date BETWEEN ? AND ?
                ORDER BY si.Date DESC
            """, (start_date, end_date))
            
            rows = c.fetchall()
            return [{
                'date': row['Date'],
                'category': row['CatName'],
                'item': row['Name'],
                'quantity': row['Quantity'],
                'cost': row['Price']
            } for row in rows]

    def get_expenses_details(self, start_date, end_date):
        """Returns list of dictionaries for detailed expenses report."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT e.ExpDate, c.CatName, e.ExpDescription, e.ExpAmount
                FROM Expenses e
                JOIN SubCategory sc ON e.ExpSubCategory = sc.SubCatID
                JOIN Category c ON sc.CategoryID = c.CatID
                WHERE e.ExpDate BETWEEN ? AND ?
                ORDER BY e.ExpDate DESC
            """, (start_date, end_date))
            
            rows = c.fetchall()
            return [{
                'date': row['ExpDate'],
                'category': row['CatName'],
                'description': row['ExpDescription'],
                'amount': row['ExpAmount']
            } for row in rows]

    def get_all_terms(self):
        """Returns list of all terms ordered by ID desc."""
        with get_db_connection() as (conn, c):
            c.execute("""SELECT TermID, TermName,Year FROM Term JOIN
                        AcademicYear WHERE 
                        Term.AcademicYear= AcademicYear.AcademicYearID ORDER BY TermID DESC """)
            return c.fetchall()

class ReportGeneratorEngine:
    def __init__(self, school_name="Pinnacle School"):
        self.school_name = school_name

    def generate_pdf(self, filepath, term_name, start_date, end_date, fees_summary, fees_breakdown, supplies_total, supplies_breakdown, supplies_details, expenses_total, expenses_breakdown, expenses_details):
        doc = SimpleDocTemplate(filepath, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        # Styles
        title_style = styles['Title']
        heading_style = styles['Heading2']
        normal_style = styles['Normal']
        
        # --- Title Page / Header ---
        story.append(Paragraph(f"{self.school_name} - Financial Report", title_style))
        story.append(Spacer(1, 12))
        story.append(Paragraph(f"<b>Term:</b> {term_name}", normal_style))
        story.append(Paragraph(f"<b>Period:</b> {start_date} to {end_date}", normal_style))
        story.append(Spacer(1, 24))

        # --- 1. FEES SECTION ---
        story.append(Paragraph("1. Fees Overview", heading_style))
        
        # Explanation
        collected_pct = (fees_summary['collected'] / fees_summary['expected'] * 100) if fees_summary['expected'] > 0 else 0
        outstanding_pct = (fees_summary['outstanding'] / fees_summary['expected'] * 100) if fees_summary['expected'] > 0 else 0
        
        explanation = f"Out of the expected fees of <b>{fees_summary['expected']:,.2f}</b> for this term, " \
                      f"<b>{fees_summary['collected']:,.2f} ({collected_pct:.1f}%)</b> has been collected, " \
                      f"leaving <b>{fees_summary['outstanding']:,.2f} ({outstanding_pct:.1f}%)</b> outstanding."
        story.append(Paragraph(explanation, normal_style))
        story.append(Spacer(1, 12))

        # Fees Breakdown Table
        if fees_breakdown:
            data = [['Class', 'Expected', 'Collected', 'Outstanding', '% Out']]
            for item in fees_breakdown:
                data.append([
                    item['class_name'],
                    f"{item['expected']:,.2f}", 
                    f"{item['collected']:,.2f}",
                    f"{item['balance']:,.2f}",
                    f"{item['percentage']:.1f}%"
                ])
            
            t = Table(data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1.5*inch, 1*inch])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            story.append(t)
        story.append(Spacer(1, 24))

        # --- 2. SUPPLIES SECTION ---
        story.append(Paragraph("2. Supplies Spending", heading_style))
        
        # Summary & Explanation
        story.append(Paragraph(f"<b>Total Spent on Supplies:</b> {supplies_total:,.2f}", normal_style))
        
        if supplies_breakdown:
            top_cat = supplies_breakdown[0]
            explanation = f"Most supply spending this term was on <b>{top_cat[0]} ({top_cat[2]:.1f}%)</b>."
            if len(supplies_breakdown) > 1:
                second_cat = supplies_breakdown[1]
                explanation += f" Followed by <b>{second_cat[0]} ({second_cat[2]:.1f}%)</b>."
            story.append(Paragraph(explanation, normal_style))
        story.append(Spacer(1, 12))

        # Supplies Breakdown Table
        if supplies_breakdown:
            data = [['Category', 'Amount', '% of Total']]
            for item in supplies_breakdown:
                data.append([
                    item[0],
                    f"{item[1]:,.2f}",
                    f"{item[2]:.1f}%"
                ])
            
            t = Table(data, colWidths=[3*inch, 2*inch, 1.5*inch])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            story.append(t)
        story.append(Spacer(1, 24))

        # --- 3. EXPENSES SECTION ---
        story.append(Paragraph("3. Operational Expenses", heading_style))
        
        # Summary & Explanation
        story.append(Paragraph(f"<b>Total Expenses:</b> {expenses_total:,.2f}", normal_style))
        
        if expenses_breakdown:
            top_cat = expenses_breakdown[0]
            explanation = f"Operational expenses were mainly driven by <b>{top_cat[0]}</b>, which accounted for <b>{top_cat[2]:.1f}%</b> of total expenses."
            story.append(Paragraph(explanation, normal_style))
        story.append(Spacer(1, 12))

        # Expenses Breakdown Table
        if expenses_breakdown:
            data = [['Category', 'Amount', '% of Total']]
            for item in expenses_breakdown:
                data.append([
                    item[0],
                    f"{item[1]:,.2f}",
                    f"{item[2]:.1f}%"
                ])
            
            t = Table(data, colWidths=[3*inch, 2*inch, 1.5*inch])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            story.append(t)
        
        # Disclaimer Footer
        story.append(Spacer(1, 36))
        disclaimer_style = ParagraphStyle('Disclaimer', parent=styles['Normal'], textColor=colors.red, fontSize=9)
        story.append(Paragraph("Note: This report reflects only system-recorded data.", disclaimer_style))

        doc.build(story)

    def generate_excel(self, filepath, term_name, start_date, end_date, fees_summary, fees_breakdown, supplies_total, supplies_breakdown, supplies_details, expenses_total, expenses_breakdown, expenses_details):
        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            workbook = writer.book
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})
            
            # --- 1. Fees Summary Sheet ---
            fees_data = {
                'Metric': ['Expected Fees', 'Collected Fees', 'Outstanding Fees', 'Collection Rate'],
                'Value': [
                    fees_summary['expected'],
                    fees_summary['collected'],
                    fees_summary['outstanding'],
                    f"{(fees_summary['collected']/fees_summary['expected']*100) if fees_summary['expected'] > 0 else 0:.1f}%"
                ]
            }
            pd.DataFrame(fees_data).to_excel(writer, sheet_name='Fees Summary', index=False)
            
            if fees_breakdown:
                df_fees = pd.DataFrame(fees_breakdown)
                # Ensure columns exist before selecting
                cols = ['class_name', 'expected', 'collected', 'balance', 'percentage']
                # Map to nice names
                df_fees = df_fees[cols].rename(columns={
                    'class_name': 'Class',
                    'expected': 'Expected',
                    'collected': 'Collected',
                    'balance': 'Outstanding',
                    'percentage': '% Outstanding'
                })
                df_fees.to_excel(writer, sheet_name='Fees Summary', startrow=6, index=False)
                
            worksheet = writer.sheets['Fees Summary']
            worksheet.set_column('B:B', 20, money_fmt)

            # --- 2. Supplies Sheet ---
            if supplies_breakdown:
                df_sup_cat = pd.DataFrame(supplies_breakdown, columns=['Category', 'Amount', 'Percentage'])
                df_sup_cat.to_excel(writer, sheet_name='Supplies', startrow=0, index=False)
            
            if supplies_details:
                df_sup_det = pd.DataFrame(supplies_details)
                df_sup_det.to_excel(writer, sheet_name='Supplies', startrow=len(supplies_breakdown) + 5, index=False)
                
            worksheet = writer.sheets['Supplies']
            worksheet.set_column('B:B', 15, money_fmt)

            # --- 3. Expenses Sheet ---
            if expenses_breakdown:
                df_exp_cat = pd.DataFrame(expenses_breakdown, columns=['Category', 'Amount', 'Percentage'])
                df_exp_cat.to_excel(writer, sheet_name='Expenses', startrow=0, index=False)
            
            if expenses_details:
                df_exp_det = pd.DataFrame(expenses_details)
                df_exp_det.to_excel(writer, sheet_name='Expenses', startrow=len(expenses_breakdown) + 5, index=False)
                
            worksheet = writer.sheets['Expenses']
            worksheet.set_column('B:B', 15, money_fmt)

def export_table_to_file(table_widget, parent=None, title="Exported Table", id_column_pattern=None):
    """
    Export the given QTableWidget to PDF or Excel, with a file dialog for user choice.
    
    New behavior:
    - If ID column exists (matches pattern): Remove Actions column (old behavior)
    - If NO ID column exists: Replace Actions column with IDs
    
    :param table_widget: QTableWidget to export
    :param parent: parent widget for dialogs
    :param title: Title for PDF export
    :param id_column_pattern: Pattern to identify ID columns (e.g., "*ID", "StudentID", "FeesID")
                             Default: Looks for columns ending with "ID", "No", "Code", "Number"
    """
    # Ask user for export type and file path
    options = QFileDialog.Options()
    file_types = "PDF Files (*.pdf);;Excel Files (*.xlsx)"
    file_path, selected_filter = QFileDialog.getSaveFileName(
        parent, "Export Table", "", file_types, options=options
    )
    if not file_path:
        return

    # Determine export type
    if selected_filter.startswith("PDF"):
        if not file_path.lower().endswith(".pdf"):
            file_path += ".pdf"
        try:
            export_table_to_pdf(table_widget, file_path, title=title, id_column_pattern=id_column_pattern)
            QMessageBox.information(parent, "Export Successful", f"Table exported to {file_path}")
        except Exception as e:
            QMessageBox.warning(parent, "Export Failed", f"PDF export failed:\n{e}")
    elif selected_filter.startswith("Excel"):
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"
        try:
            export_table_to_excel(table_widget, file_path, id_column_pattern=id_column_pattern)
            QMessageBox.information(parent, "Export Successful", f"Table exported to {file_path}")
        except Exception as e:
            QMessageBox.warning(parent, "Export Failed", f"Excel export failed:\n{e}")


def export_table_to_excel(table_widget, file_path, id_column_pattern=None):
    """
    Export QTableWidget to Excel file.
    
    Logic:
    1. Detect ID columns based on pattern
    2. If ID column exists: Remove Actions column
    3. If NO ID column: Replace Actions column with row IDs/numbers
    """
    wb = Workbook()
    ws = wb.active
    
    # Get all headers
    headers = []
    for col in range(table_widget.columnCount()):
        header = table_widget.horizontalHeaderItem(col)
        headers.append(header.text() if header else f"Column {col+1}")
    
    # Detect ID columns
    id_columns = []
    if id_column_pattern:
        # Use provided pattern
        for idx, header in enumerate(headers):
            if id_column_pattern == "*ID" and header.endswith("ID"):
                id_columns.append(idx)
            elif header == id_column_pattern:
                id_columns.append(idx)
            elif id_column_pattern in header:
                id_columns.append(idx)
    else:
        # Auto-detect ID columns (ends with ID, No, Code, Number)
        id_patterns = ["ID$", "No$", "Code$", "Number$"]
        for idx, header in enumerate(headers):
            for pattern in id_patterns:
                if re.search(pattern, header, re.IGNORECASE):
                    id_columns.append(idx)
                    break
    
    # Check if Actions column exists (last column with widgets)
    has_actions = False
    actions_col = table_widget.columnCount() - 1
    if actions_col >= 0:
        # Check if any row has widget in last column (typical for action buttons)
        for row in range(table_widget.rowCount()):
            if table_widget.cellWidget(row, actions_col):
                has_actions = True
                break
    
    # Determine export strategy
    if id_columns:
        # ID column exists → Remove Actions column (old behavior)
        if has_actions:
            # Exclude the Actions column
            export_cols = list(range(table_widget.columnCount() - 1))
            export_headers = [headers[i] for i in export_cols]
        else:
            # No Actions column, export all
            export_cols = list(range(table_widget.columnCount()))
            export_headers = headers
    else:
        # NO ID column exists → Replace Actions with IDs
        if has_actions:
            # Replace Actions header with "ID"
            export_cols = list(range(table_widget.columnCount()))
            export_headers = headers.copy()
            export_headers[-1] = "ID"  # Replace Actions header with "ID"
        else:
            # No Actions column, add ID column
            export_cols = list(range(table_widget.columnCount()))
            export_headers = headers
            export_headers.append("ID")  # Add ID column
    
    # Write headers
    ws.append(export_headers)
    
    # Write data rows
    for row in range(table_widget.rowCount()):
        row_data = []
        
        # Add regular columns
        for col in export_cols:
            if col < table_widget.columnCount():  # Skip if col is beyond original columns
                item = table_widget.item(row, col)
                if item:
                    row_data.append(item.text())
                else:
                    cell_widget = table_widget.cellWidget(row, col)
                    if cell_widget:
                        # For action widgets, try to extract ID if available
                        if not id_columns and col == actions_col and has_actions:
                            # Replace Actions with row number (or extract ID from button if possible)
                            # Try to find ID from other columns first
                            found_id = None
                            # Look for any column that might contain ID-like data
                            for check_col in range(table_widget.columnCount()):
                                check_item = table_widget.item(row, check_col)
                                if check_item:
                                    text = check_item.text()
                                    # Check if text looks like an ID
                                    if (re.match(r'^\d+$', text) or  # All numbers
                                        re.match(r'^[A-Z]+-\d+$', text) or  # ABC-123 pattern
                                        re.match(r'^\d{4}-\d{4}$', text)):  # 2024-0001 pattern
                                        found_id = text
                                        break
                            
                            if found_id:
                                row_data.append(found_id)
                            else:
                                # Use row number as ID
                                row_data.append(f"Row-{row+1}")
                        else:
                            row_data.append("")
                    else:
                        row_data.append("")
            else:
                # This is for added ID column when no Actions existed
                # Use row number as ID
                row_data.append(f"Row-{row+1}")
        
        # If we need to add ID column (no Actions, no ID columns)
        if not has_actions and not id_columns:
            # Add row number as ID
            row_data.append(f"Row-{row+1}")
        
        ws.append(row_data)
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(file_path)


def export_table_to_pdf(table_widget, file_path, title="Exported Table", id_column_pattern=None):
    """
    Export QTableWidget to PDF file.
    
    Logic:
    1. Detect ID columns based on pattern
    2. If ID column exists: Remove Actions column
    3. If NO ID column: Replace Actions column with row IDs/numbers
    """
    # Get all headers
    headers = []
    for col in range(table_widget.columnCount()):
        header = table_widget.horizontalHeaderItem(col)
        headers.append(header.text() if header else f"Column {col+1}")
    
    # Detect ID columns
    id_columns = []
    if id_column_pattern:
        # Use provided pattern
        for idx, header in enumerate(headers):
            if id_column_pattern == "*ID" and header.endswith("ID"):
                id_columns.append(idx)
            elif header == id_column_pattern:
                id_columns.append(idx)
            elif id_column_pattern in header:
                id_columns.append(idx)
    else:
        # Auto-detect ID columns (ends with ID, No, Code, Number)
        id_patterns = ["ID$", "No$", "Code$", "Number$"]
        for idx, header in enumerate(headers):
            for pattern in id_patterns:
                if re.search(pattern, header, re.IGNORECASE):
                    id_columns.append(idx)
                    break
    
    # Check if Actions column exists (last column with widgets)
    has_actions = False
    actions_col = table_widget.columnCount() - 1
    if actions_col >= 0:
        for row in range(table_widget.rowCount()):
            if table_widget.cellWidget(row, actions_col):
                has_actions = True
                break
    
    # Prepare data for PDF
    data = []
    
    # Determine which columns to include and headers
    if id_columns:
        # ID column exists → Remove Actions column
        if has_actions:
            col_count = table_widget.columnCount() - 1  # Exclude Actions
            pdf_headers = [headers[i] for i in range(col_count)]
        else:
            col_count = table_widget.columnCount()
            pdf_headers = headers
    else:
        # NO ID column → Replace Actions with ID
        if has_actions:
            col_count = table_widget.columnCount()
            pdf_headers = headers.copy()
            pdf_headers[-1] = "ID"  # Replace Actions header
        else:
            col_count = table_widget.columnCount()
            pdf_headers = headers + ["ID"]  # Add ID column
    
    # Add headers to data
    data.append(pdf_headers)
    
    # Add data rows
    for row in range(table_widget.rowCount()):
        row_data = []
        
        # Add data from visible columns
        for col in range(col_count):
            if col < table_widget.columnCount():  # Regular column
                if col < table_widget.columnCount() - 1 or not has_actions:
                    # Regular cell
                    item = table_widget.item(row, col)
                    if item:
                        row_data.append(item.text())
                    else:
                        cell_widget = table_widget.cellWidget(row, col)
                        if cell_widget and not id_columns and col == actions_col and has_actions:
                            # This is Actions column with no ID columns
                            # Try to find ID from row data
                            found_id = None
                            for check_col in range(table_widget.columnCount()):
                                check_item = table_widget.item(row, check_col)
                                if check_item:
                                    text = check_item.text()
                                    if (re.match(r'^\d+$', text) or
                                        re.match(r'^[A-Z]+-\d+$', text) or
                                        re.match(r'^\d{4}-\d{4}$', text)):
                                        found_id = text
                                        break
                            
                            if found_id:
                                row_data.append(found_id)
                            else:
                                row_data.append(f"Row-{row+1}")
                        else:
                            row_data.append("")
                else:
                    # Actions column when no ID columns exist
                    if not id_columns and has_actions:
                        # Try to extract ID
                        found_id = None
                        for check_col in range(table_widget.columnCount() - 1):  # Exclude Actions column
                            check_item = table_widget.item(row, check_col)
                            if check_item:
                                text = check_item.text()
                                if (re.match(r'^\d+$', text) or
                                    re.match(r'^[A-Z]+-\d+$', text) or
                                    re.match(r'^\d{4}-\d{4}$', text)):
                                    found_id = text
                                    break
                        
                        if found_id:
                            row_data.append(found_id)
                        else:
                            row_data.append(f"Row-{row+1}")
                    else:
                        row_data.append("")
            else:
                # Added ID column (when no Actions existed)
                row_data.append(f"Row-{row+1}")
        
        # If we need to add ID column (no Actions, no ID columns)
        if not has_actions and not id_columns:
            row_data.append(f"Row-{row+1}")
        
        data.append(row_data)
    
    # Create PDF
    pdf = SimpleDocTemplate(file_path, pagesize=landscape(A4))
    
    # Create table
    table = Table(data, repeatRows=1)
    
    # Table style
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#007bff")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.whitesmoke]),
    ])
    
    # Add zebra striping
    for i in range(1, len(data)):
        if i % 2 == 0:
            bg_color = colors.white
        else:
            bg_color = colors.whitesmoke
        style.add('BACKGROUND', (0, i), (-1, i), bg_color)
    
    table.setStyle(style)
    
    # Title paragraph
    styles = getSampleStyleSheet()
    title_paragraph = Paragraph(f"<b>{title}</b>", styles['Title'])
    spacer = Spacer(1, 0.2 * inch)
    
    # Footer callback
    def footer(canvas, doc):
        canvas.saveState()
        footer_text = f"Printed by Pinnacle Ledger    {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(doc.pagesize[0] - 30, 20, footer_text)
        canvas.restoreState()
    
    elems = [title_paragraph, spacer, table]
    pdf.build(elems, onFirstPage=footer, onLaterPages=footer)


# Helper function to check if a table has ID columns
def has_id_columns(table_widget, id_column_pattern=None):
    """
    Check if table has ID columns.
    
    :return: Tuple (has_id, id_columns_list, has_actions)
    """
    headers = []
    for col in range(table_widget.columnCount()):
        header = table_widget.horizontalHeaderItem(col)
        headers.append(header.text() if header else "")
    
    # Detect ID columns
    id_columns = []
    if id_column_pattern:
        for idx, header in enumerate(headers):
            if id_column_pattern == "*ID" and header.endswith("ID"):
                id_columns.append(idx)
            elif header == id_column_pattern:
                id_columns.append(idx)
            elif id_column_pattern in header:
                id_columns.append(idx)
    else:
        id_patterns = ["ID$", "No$", "Code$", "Number$"]
        for idx, header in enumerate(headers):
            for pattern in id_patterns:
                if re.search(pattern, header, re.IGNORECASE):
                    id_columns.append(idx)
                    break
    
    # Check for Actions column
    has_actions = False
    actions_col = table_widget.columnCount() - 1
    if actions_col >= 0:
        for row in range(table_widget.rowCount()):
            if table_widget.cellWidget(row, actions_col):
                has_actions = True
                break
    
    return len(id_columns) > 0, id_columns, has_actions


if __name__ == "__main__":
    app = QApplication(sys.argv)
    initialize_database()
    # add window icon
    app.setWindowIcon(QIcon(resource_path("icons/app-icon.ico")))
    # Load global stylesheet
    try:
        with open(resource_path("theme.qss"), "r") as f:
            stylesheet = f.read()
            app.setStyleSheet(stylesheet)
    except FileNotFoundError:
        QMessageBox.warning("Warning: theme.qss not found. Running without global stylesheet.")

    window = Window()
    window.show()
    sys.exit(app.exec_())
    