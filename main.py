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


# Enhanced connection function with context manager
@contextmanager
def get_db_connection():
    """Context manager for database connections that handles opening and closing automatically"""
    conn = None
    try:
        conn = sqlite3.connect("Database.db")
        conn.row_factory = sqlite3.Row  # Optional: for named column access
        c = conn.cursor()
        yield conn, c
    except sqlite3.Error as e:
        print(f"Database connection error: {e}")
        raise
    finally:
        if conn:
            conn.close()
            
# Alternative: Simple function that always returns a fresh connection
def get_db_connection_simple():
    """Simple function that always returns a fresh connection"""
    try:
        conn = sqlite3.connect("Database.db")
        c = conn.cursor()
        return conn, c
    except sqlite3.Error as e:
        print(f"Database connection error: {e}")
        raise
### Student Management Initial Start ###
class StudentManagement(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(15)

        # Header for the entire page
        header_layout = QHBoxLayout()
        title = QLabel("Students Management")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        layout.addLayout(header_layout)

        # Tab widget to separate Students and Enrollments
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #E0E3E7;
                border-radius: 5px;
            }
            QTabBar::tab {
                background: #F7FAFC;
                border: 1px solid #E0E3E7;
                border-bottom: none;
                padding: 10px 20px;
                min-width: 100px;
            }
            QTabBar::tab:selected {
                background: #FFFFFF;
                font-weight: bold;
            }
        """)
        self.students_tab = StudentsTab()
        self.enrollments_tab = EnrollmentsTab()
        self.tabs.addTab(self.students_tab, "Students Info")
        self.tabs.addTab(self.enrollments_tab, "Student Enrollments")
        layout.addWidget(self.tabs)

class StudentsTab(QWidget):
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
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QCheckBox {
                spacing: 5px;
                font-size: 14px;
                color: #2D3748;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border-radius: 3px;
                border: 1px solid #E0E3E7;
            }
            QCheckBox::indicator:checked {
                background-color: #4299E1;
                border: 1px solid #4299E1;
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
        self.filterbyStatus.addItem("25% Scholarship",75)
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
        status_combo.addItem("25% Scholarship",75)
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
        if not all([student_id, name, feepayer, status, parent_name, contact]):
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
                self.table.setColumnCount(8)  # 7 data columns + 1 actions column
                self.table.setHorizontalHeaderLabels([
                    "Student ID", "Student Name", "Fee Payer", "PayRate", 
                    "Parent Name", "Parent Contact", "Parent Contact2", "Actions"
                ])
                
                # Configure column resize modes
                header = self.table.horizontalHeader()
                for col in range(7):  # First 7 columns are data
                    header.setSectionResizeMode(col, QHeaderView.Stretch)
                header.setSectionResizeMode(7, QHeaderView.Fixed)  # Actions column fixed width
                header.resizeSection(7, 150)  # Wider width for action buttons
                
                self.table.setRowCount(0)
                for row_data in rows:
                    row = self.table.rowCount()
                    self.table.insertRow(row)
                    
                    # Add data columns
                    for col, data in enumerate(row_data):
                        item = QTableWidgetItem(str(data) if data else "")
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Make non-editable
                        self.table.setItem(row, col, item)
                    
                    # Create actions widget with proper sizing
                    actions_widget = QWidget()
                    actions_widget.setFixedWidth(140)  # Fixed width for the widget
                    actions_layout = QHBoxLayout(actions_widget)
                    actions_layout.setContentsMargins(5, 2, 5, 2)  # Reduced margins
                    actions_layout.setSpacing(8)  # Slightly more spacing
                    
                    # Edit button with fixed size
                    edit_btn = QPushButton("Edit")
                    edit_btn.setFixedSize(60, 28)  # Larger size for better visibility
                    edit_btn.setStyleSheet("""
                        QPushButton {
                            background-color: #48BB78;
                            color: white;
                            border: none;
                            border-radius: 4px;
                            font-size: 11px;
                            font-weight: bold;
                        }
                        QPushButton:hover {
                            background-color: #38A169;
                        }
                        QPushButton:pressed {
                            background-color: #2F855A;
                        }
                    """)
                    edit_btn.clicked.connect(lambda checked, sid=row_data[0]: self.edit_student(sid))
                    
                    # Delete button with fixed size
                    delete_btn = QPushButton("Delete")
                    delete_btn.setFixedSize(60, 28)  # Larger size for better visibility
                    delete_btn.setStyleSheet("""
                        QPushButton {
                            background-color: #F56565;
                            color: white;
                            border: none;
                            border-radius: 4px;
                            font-size: 11px;
                            font-weight: bold;
                        }
                        QPushButton:hover {
                            background-color: #E53E3E;
                        }
                        QPushButton:pressed {
                            background-color: #C53030;
                        }
                    """)
                    delete_btn.clicked.connect(lambda checked, sid=row_data[0]: self.delete_student(sid))
                    
                    actions_layout.addWidget(edit_btn)
                    actions_layout.addWidget(delete_btn)
                    actions_layout.addStretch()
                    
                    # Set the widget in the Actions column (column 7)
                    self.table.setCellWidget(row, 7, actions_widget)
                    
                    # Set row height to accommodate buttons
                    self.table.setRowHeight(row, 40)
                
                count = self.table.rowCount()
                self.results_label.setText(f"{count} results")
                
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")

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
                status_combo.addItem("25% Scholarship",75)
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
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
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
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Filters and add buttons
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search...")
        self.search_input.setMaximumWidth(200)
        self.search_input.textChanged.connect(self.load_data)

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

        filter_layout.addWidget(self.search_input)
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
                        LEFT JOIN Term t ON st.TermID = t.TermID"""
                filters = []
                params = ()
                search_text = self.search_input.text().strip()
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
                        item = QTableWidgetItem(str(data) if data else "")
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)  # Make non-editable
                        self.table.setItem(row, col, item)
                    
                    student_term_id = row_data[0]  # StudentTermID for actions
                    student_id = row_data[1]       # StudentID
                    class_name = row_data[3]       # ClassName
                    term_name = row_data[4]        # TermName
                    
                    # Create actions widget with proper sizing (now includes Pay Fee button)
                    actions_widget = QWidget()
                    actions_widget.setFixedWidth(210)  # Increased width for 3 buttons
                    actions_layout = QHBoxLayout(actions_widget)
                    actions_layout.setContentsMargins(5, 2, 5, 2)
                    actions_layout.setSpacing(5)
                    
                    # Pay Fee button
                    pay_btn = QPushButton("Pay Fee")
                    pay_btn.setFixedSize(60, 28)
                    pay_btn.setStyleSheet("""
                        QPushButton {
                            background-color: #4299E1;
                            color: white;
                            border: none;
                            border-radius: 4px;
                            font-size: 11px;
                            font-weight: bold;
                        }
                        QPushButton:hover {
                            background-color: #3182CE;
                        }
                        QPushButton:pressed {
                            background-color: #2B6CB0;
                        }
                    """)
                    pay_btn.clicked.connect(lambda checked, stid=student_term_id: self.pay_fee(stid))
                    
                    # Edit button
                    edit_btn = QPushButton("Edit")
                    edit_btn.setFixedSize(60, 28)
                    edit_btn.setStyleSheet("""
                        QPushButton {
                            background-color: #48BB78;
                            color: white;
                            border: none;
                            border-radius: 4px;
                            font-size: 11px;
                            font-weight: bold;
                        }
                        QPushButton:hover {
                            background-color: #38A169;
                        }
                        QPushButton:pressed {
                            background-color: #2F855A;
                        }
                    """)
                    edit_btn.clicked.connect(lambda checked, tid=student_term_id: self.edit_enrollment(tid))
                    
                    # Delete button
                    delete_btn = QPushButton("Delete")
                    delete_btn.setFixedSize(60, 28)
                    delete_btn.setStyleSheet("""
                        QPushButton {
                            background-color: #F56565;
                            color: white;
                            border: none;
                            border-radius: 4px;
                            font-size: 11px;
                            font-weight: bold;
                        }
                        QPushButton:hover {
                            background-color: #E53E3E;
                        }
                        QPushButton:pressed {
                            background-color: #C53030;
                        }
                    """)
                    delete_btn.clicked.connect(lambda checked, tid=student_term_id: self.delete_enrollment(tid))
                    
                    actions_layout.addWidget(pay_btn)
                    actions_layout.addWidget(edit_btn)
                    actions_layout.addWidget(delete_btn)
                    actions_layout.addStretch()
                    
                    # Set the widget in the Actions column (column 5)
                    self.table.setCellWidget(row, 5, actions_widget)
                    
                    # Set row height to accommodate buttons
                    self.table.setRowHeight(row, 40)
                
                count = self.table.rowCount()
                self.results_label.setText(f"{count} results")
                
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")    
    
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
                payrate_combo.addItem("25% Scholarship",75)
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
    
    def add_action_buttons(self, row, student_term_id, student_id, class_name, term_name):
        """Add action buttons including Pay Fee button"""
        actions_widget = QWidget()
        actions_widget.setFixedWidth(220)  # Increased width for additional button
        actions_layout = QHBoxLayout(actions_widget)
        actions_layout.setContentsMargins(5, 2, 5, 2)
        actions_layout.setSpacing(5)
        
        # Pay Fee button
        pay_btn = QPushButton("Pay Fee")
        pay_btn.setFixedSize(60, 28)
        pay_btn.setStyleSheet("""
            QPushButton {
                background-color: #4299E1;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3182CE;
            }
            QPushButton:pressed {
                background-color: #2B6CB0;
            }
        """)
        pay_btn.clicked.connect(lambda checked, stid=student_term_id: self.pay_fee(stid))
        
        # Edit button
        edit_btn = QPushButton("Edit")
        edit_btn.setFixedSize(60, 28)
        edit_btn.setStyleSheet("""
            QPushButton {
                background-color: #48BB78;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #38A169;
            }
        """)
        edit_btn.clicked.connect(lambda checked, stid=student_term_id: self.edit_enrollment(stid))
        
        # Delete button
        delete_btn = QPushButton("Delete")
        delete_btn.setFixedSize(60, 28)
        delete_btn.setStyleSheet("""
            QPushButton {
                background-color: #F56565;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #E53E3E;
            }
        """)
        delete_btn.clicked.connect(lambda checked, stid=student_term_id: self.delete_enrollment(stid))
        
        actions_layout.addWidget(pay_btn)
        actions_layout.addWidget(edit_btn)
        actions_layout.addWidget(delete_btn)
        actions_layout.addStretch()
        
        self.table.setCellWidget(row, 5, actions_widget)  # Adjust column index as needed

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
        
        # Light theme stylesheet for a modern, web-like light tone
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
        QComboBox {
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
        """
        self.setStyleSheet(light_style)
        
        # Initialize variables
        self.mode = None
        self.excel_file = None
        self.term_id = None
        self.class_id = None
        self.data = None
        
        # Add pages
        self.intro_page_id = self.addPage(self.create_intro_page())
        self.file_page_id = self.addPage(self.create_file_page())
        self.config_page_id = self.addPage(self.create_config_page())
        self.preview_page_id = self.addPage(self.create_preview_page())
        self.finish_page_id = self.addPage(self.create_finish_page())

    def create_intro_page(self):
        page = QWizardPage()
        page.setTitle("Choose Import Mode")
        layout = QVBoxLayout()
        label = QLabel("Select the import method:")
        layout.addWidget(label)
        self.mode_group = QButtonGroup()
        radio1 = QRadioButton("Insert new students with initial term enrollment")
        radio2 = QRadioButton("Enrollment only")
        self.mode_group.addButton(radio1, 1)
        self.mode_group.addButton(radio2, 2)
        layout.addWidget(radio1)
        layout.addWidget(radio2)
        page.setLayout(layout)
        return page

    def create_file_page(self):
        page = QWizardPage()
        page.setTitle("Select Excel File")
        layout = QVBoxLayout()
        self.file_label = QLabel("No file selected")
        layout.addWidget(self.file_label)
        btn = QPushButton("Browse")
        btn.clicked.connect(self.browse_file)
        layout.addWidget(btn)
        page.setLayout(layout)
        return page

    def browse_file(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.excel_file = file
            self.file_label.setText(file)

    def create_config_page(self):
        page = QWizardPage()
        page.setTitle("Configure Enrollment")
        layout = QVBoxLayout()
        
        # Academic Year selection
        year_label = QLabel("Select Academic Year:")
        self.academic_combo = QComboBox()
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear")
                years = c.fetchall()
                for yid, yname in years:
                    self.academic_combo.addItem(yname if yname else f"Year {yid}", yid)
        except sqlite3.Error as e:
            QMessageBox.warning(self, "Database Error", f"Failed to load academic years: {str(e)}")
        
        self.academic_combo.currentIndexChanged.connect(self.load_terms_and_classes)
        
        layout.addWidget(year_label)
        layout.addWidget(self.academic_combo)
        
        # Term selection
        term_label = QLabel("Select Term:")
        self.term_combo = QComboBox()
        layout.addWidget(term_label)
        layout.addWidget(self.term_combo)
        
        # Class selection
        class_label = QLabel("Select Class:")
        self.class_combo = QComboBox()
        layout.addWidget(class_label)
        layout.addWidget(self.class_combo)
        
        page.setLayout(layout)
        
        # Preselect first year if available
        if self.academic_combo.count() > 0:
            self.academic_combo.setCurrentIndex(0)
        
        return page

    def load_terms_and_classes(self):
        year_id = self.academic_combo.currentData()
        if year_id is None:
            return
        
        self.term_combo.clear()
        self.class_combo.clear()
        
        try:
            with get_db_connection() as (conn, c):
                # Load terms for selected year
                c.execute("SELECT TermID, TermName FROM Term WHERE AcademicYear = ?", (year_id,))
                terms = c.fetchall()
                for tid, tname in terms:
                    self.term_combo.addItem(tname if tname else f"Term {tid}", tid)
                
                # Load classes for selected year
                c.execute("SELECT ClassID, ClassName FROM Class WHERE AcademicYear = ?", (year_id,))
                classes = c.fetchall()
                for cid, cname in classes:
                    self.class_combo.addItem(cname if cname else f"Class {cid}", cid)
        except sqlite3.Error as e:
            QMessageBox.warning(self, "Database Error", f"Failed to load terms/classes: {str(e)}")
    
    def create_preview_page(self):
        page = QWizardPage()
        page.setTitle("Preview Data")
        layout = QVBoxLayout()
        self.table_view = QTableView()
        layout.addWidget(self.table_view)
        page.setLayout(layout)
        return page

    def create_finish_page(self):
        page = QWizardPage()
        page.setTitle("Finish Import")
        layout = QVBoxLayout()
        label = QLabel("Review the preview and click Finish to import the data.")
        layout.addWidget(label)
        page.setLayout(layout)
        return page

    def initializePage(self, page_id):
        super().initializePage(page_id)
        if page_id == self.preview_page_id:
            self.load_data()

    def load_data(self):
        if not self.excel_file:
            QMessageBox.warning(self, "Error", "No Excel file selected.")
            return
        
        try:
            self.data = pd.read_excel(self.excel_file)
            self.mode = self.mode_group.checkedId()
            
            if self.mode == 1:  # New students
                required_cols = ['StudentID', 'Student_Name', 'FeePayer', 'PayRate', 'Parent_Name', 'Parent_Contact']
                if not all(col in self.data.columns for col in required_cols):
                    raise ValueError("Missing required columns for new students: " + ", ".join(required_cols))
            elif self.mode == 2:  # Enrollment only
                if 'StudentID' not in self.data.columns:
                    raise ValueError("Missing 'StudentID' column for enrollment only.")
            else:
                raise ValueError("No mode selected.")
            
            # Display in table
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(self.data.columns)
            for _, row in self.data.iterrows():
                items = [QStandardItem(str(row.get(col, ''))) for col in self.data.columns]
                model.appendRow(items)
            self.table_view.setModel(model)
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")
            self.data = None

    def accept(self):
        if self.data is None or self.data.empty:
            QMessageBox.warning(self, "Error", "No data to import.")
            return
        
        self.term_id = self.term_combo.currentData()
        self.class_id = self.class_combo.currentData()
        
        if not self.term_id or not self.class_id:
            QMessageBox.warning(self, "Error", "Term and Class must be selected.")
            return
        
        success_count = 0
        error_count = 0
        error_messages = []
        
        try:
            with get_db_connection() as (conn, c):
                for index, row in self.data.iterrows():
                    try:
                        sid = str(row.get('StudentID', '')).strip()
                        if not sid:
                            raise ValueError("Invalid or missing StudentID.")
                        
                        if self.mode == 1:  # New students with enrollment
                            # Validation: Check required fields not null/empty
                            required_cols_new = ['StudentID', 'Student_Name', 'FeePayer', 'PayRate', 'Parent_Name', 'Parent_Contact']
                            if any(pd.isnull(row.get(col)) or not str(row.get(col)).strip() for col in required_cols_new):
                                raise ValueError("Missing or empty required field.")
                            
                            # Check if StudentID already exists
                            c.execute("SELECT StudentID FROM Students WHERE StudentID = ?", (sid,))
                            if c.fetchone():
                                raise ValueError("StudentID already exists.")
                            
                            # Insert into Students
                            values = (
                                sid,
                                str(row['Student_Name']).strip(),
                                str(row['FeePayer']).strip(),
                                str(row['PayRate']).strip(),
                                str(row['Parent_Name']).strip(),
                                str(row['Parent_Contact']).strip(),
                                str(row.get('Parent_Contact2', '')).strip() or None
                            )
                            c.execute("""
                                INSERT INTO Students (StudentID, Student_Name, FeePayer, PayRate, Parent_Name, Parent_Contact, Parent_Contact2)
                                VALUES (?, ?, ?, ?, ?, ?, ?)
                            """, values)
                            
                            # Insert into StudentTerm
                            c.execute("""
                                INSERT INTO StudentTerm (StudentID, ClassID, TermID)
                                VALUES (?, ?, ?)
                            """, (sid, self.class_id, self.term_id))
                        
                        elif self.mode == 2:  # Enrollment only
                            # Check if StudentID exists
                            c.execute("SELECT StudentID FROM Students WHERE StudentID = ?", (sid,))
                            if not c.fetchone():
                                raise ValueError("StudentID does not exist.")
                            
                            # Check if already enrolled in this term/class
                            c.execute("""
                                SELECT StudentTermID FROM StudentTerm 
                                WHERE StudentID = ? AND ClassID = ? AND TermID = ?
                            """, (sid, self.class_id, self.term_id))
                            if c.fetchone():
                                raise ValueError("Already enrolled in this term/class.")
                            
                            # Insert into StudentTerm
                            c.execute("""
                                INSERT INTO StudentTerm (StudentID, ClassID, TermID)
                                VALUES (?, ?, ?)
                            """, (sid, self.class_id, self.term_id))
                        
                        success_count += 1
                        
                    except Exception as e:
                        error_count += 1
                        error_messages.append(f"Row {index + 2}: {str(e)}")  # +2 because Excel rows start at 1 + header
                
                # Commit only if there were successful operations
                if success_count > 0:
                    conn.commit()
                else:
                    conn.rollback()
                    
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to connect to database: {str(e)}")
            return
        
        # Show results
        if error_messages and error_count > 0:
            error_details = "\n".join(error_messages[:10])  # Show first 10 errors
            if error_count > 10:
                error_details += f"\n... and {error_count - 10} more errors"
            
            QMessageBox.warning(self, "Import Completed with Errors", 
                               f"Successfully imported {success_count} records.\n"
                               f"Failed to import {error_count} records.\n\n"
                               f"Errors:\n{error_details}")
        else:
            QMessageBox.information(self, "Import Complete", 
                                   f"Successfully imported {success_count} records.")
        
        super().accept()

### Student Management End ###

class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("School Finance MIS")
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
        self.toggle_btn = QPushButton(QIcon("Icons/menu-icon.png"), "Menu")
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
            ("Reports", "Icons/report-icon.png")
        ]
        
        self.menu_buttons = []
        for name, icon_path in self.menu_items:
            btn = QPushButton(QIcon(icon_path), name)
            btn.setFixedHeight(40)
            btn.setIconSize(QSize(24, 24))  # Enlarge the icon
            btn.original_text = name  # Store the original text
            btn.icon_path = icon_path  # Store the icon path for later use
            self.menu_buttons.append(btn)
            self.sidebar_layout.addWidget(btn)
        self.sidebar_layout.addStretch()
        # Combine both menu lists for easier management
        all_menu_items =[
            ("Account", "Icons/account-icon.png"),
            ("About", "Icons/about-icon.png")
        ]
        for name2,icon_path2 in all_menu_items:
            buttons = QPushButton(QIcon(icon_path2), name2)
            buttons.setFixedHeight(40)
            buttons.setIconSize(QSize(24, 24))
            buttons.original_text = name2  # Store the original text
            buttons.icon_path = icon_path2
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
        
        self.school_name = QLabel("Springfield High School")
        self.header_layout.addWidget(self.school_name)
        
        self.content_layout.addWidget(self.header)
        
        # Separator line
        self.separator = QWidget()
        self.separator.setFixedHeight(2)
        self.content_layout.addWidget(self.separator)
        
        # Scroll area for body
        self.MainBodyHolder = QScrollArea()
        self.MainBodyHolder.setWidgetResizable(True)
        self.body = QStackedWidget()
        self.AcademicYearTermPage = AcademicYearTerm()
        self.DashboardPage = Dashboard()
        self.StudentPage = StudentManagement()
        self.ClassPage = Class()
        self.FeesPage = Fees()
        self.UniformsPage = Uniforms()
        self.ExpensesPage = Expenditure()
        self.ReportsPage = Reports()
        self.AboutPage = About()
        self.AccountPage = Account()
        self.body.addWidget(self.DashboardPage)
        self.body.addWidget(self.StudentPage)
        self.body.addWidget(self.FeesPage)
        self.body.addWidget(self.UniformsPage)
        self.body.addWidget(self.ExpensesPage)
        self.body.addWidget(self.ClassPage)
        self.body.addWidget(self.AcademicYearTermPage)
        self.body.addWidget(self.ReportsPage)
        self.body.addWidget(self.AboutPage)
        self.body.addWidget(self.AccountPage)
        self.body_layout = QVBoxLayout(self.body)
        self.body_layout.addStretch()
        self.MainBodyHolder.setWidget(self.body)
        self.content_layout.addWidget(self.MainBodyHolder)
        
        # Add sidebar and content to main layout
        self.main_layout.addWidget(self.sidebar)
        self.main_layout.addWidget(self.content)
        
        # Sidebar state
        self.sidebar_expanded = True
        self.apply_styles()
            # Connect sidebar buttons to stacked pages
        page_map = {
            "Dashboard": self.DashboardPage,
            "Students": self.StudentPage,
            "Fees": self.FeesPage,
            "Uniforms": self.UniformsPage,
            "Expenditure": self.ExpensesPage,
            "Classes": self.ClassPage,
            "Academic Calendar": self.AcademicYearTermPage,
            "Reports": self.ReportsPage,
            "About": self.AboutPage,
            "Account": self.AccountPage
        }

        def make_nav_handler(page_widget, page_name):
            def handler():
                self.body.setCurrentWidget(page_widget)
                self.page_title.setText(page_name)
            return handler

        for btn in self.menu_buttons:
            page_widget = page_map.get(btn.original_text)
            if page_widget:
                btn.clicked.connect(make_nav_handler(page_widget, btn.original_text))

    def toggle_sidebar(self):
        self.sidebar_expanded = not self.sidebar_expanded
        new_width = 60 if not self.sidebar_expanded else 250
        self.sidebar.setFixedWidth(new_width)

        for btn in self.menu_buttons:
            if self.sidebar_expanded:
                btn.setText(btn.original_text)  # Restore original text
            else:
                btn.setText("")  # Hide text
            btn.setIcon(QIcon(btn.icon_path))  # Always use the stored icon path
            btn.setIconSize(QSize(24, 24))  # Ensure icon stays enlarged

    def apply_styles(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F5F7FA;
            }
            QWidget#sidebar {
                background-color: #FFFFFF;
                border-right: 1px solid #E0E3E7;
            }
            QPushButton {
                background-color: transparent;
                border: none;
                text-align: left;
                padding: 10px;
                font-size: 16px;
                color: #2D3748;
            }
            QPushButton:hover {
                background-color: #EDF2F7;
                border-radius: 5px;
            }
            QPushButton#toggle_btn {
                background-color: transparent;
                border-radius: 20px;
            }
            QPushButton#toggle_btn:hover {
                background-color: #EDF2F7;
                border-radius: 5px;
            }
            QLabel#page_title {
                font-size: 24px;
                font-weight: bold;
                color: #2D3748;
            }
            QLabel#school_name {
                font-size: 18px;
                color: #4A5568;
            }
            QWidget#separator {
                background-color: #E0E3E7;
            }
            QScrollArea {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                border-radius: 5px;
            }
            QScrollArea QWidget {
                background-color: transparent;
            }
        """)
        # Set object names for specific styling
        self.sidebar.setObjectName("sidebar")
        self.toggle_btn.setObjectName("toggle_btn")
        self.page_title.setObjectName("page_title")
        self.school_name.setObjectName("school_name")
        self.separator.setObjectName("separator")
    
class AcademicYearTerm(QWidget):
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
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QTableWidget {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                border-radius: 5px;
                gridline-color: #E0E3E7;
                selection-background-color: #0c3157;
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
            QTableWidget QPushButton {
                background-color: transparent;
                color: #4299E1;
                padding: 5px 10px;
                border: 1px solid #4299E1;
                border-radius: 3px;
                font-size: 12px;
            }
            QTableWidget QPushButton:hover {
                background-color: #EDF2F7;
            }
            QTableWidget QPushButton#delete {
                color: #E53E3E;
                border: 1px solid #E53E3E;
            }
            QTableWidget QPushButton#delete:hover {
                background-color: #FED7D7;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(15)

        # Header row
        header_layout = QHBoxLayout()
        title = QLabel("Academic Year/Term")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        # Action ComboBox for Academic Year
        self.year_action_combo = QComboBox()
        self.year_action_combo.setStyleSheet("""
            QComboBox {
            background-color: #4299E1;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 5px;
            font-size: 14px;
            }
            QComboBox::drop-down {
            border: none;
            }
            QComboBox QAbstractItemView {
            background-color: #FFFFFF;
            border: 1px solid #E0E3E7;
            selection-background-color: #EDF2F7;
            color: #2D3748;
            font-size: 14px;
            }
        """)
        self.year_action_combo.addItem("Year ▼", None)
        self.year_action_combo.addItem("Add Year", "add")
        self.year_action_combo.addItem("Edit Year", "edit")
        self.year_action_combo.addItem("Delete Year", "delete")
        self.year_action_combo.activated.connect(self.handle_year_action)

        add_term_btn = QPushButton("Add Term")
        add_term_btn.clicked.connect(self.addTerm)
        header_layout.addWidget(self.year_action_combo)
        header_layout.addWidget(add_term_btn)
        layout.addLayout(header_layout)

        # Filter row
        filter_layout = QHBoxLayout()
        self.Search_Input = QLineEdit()
        self.Search_Input.setPlaceholderText("Search...")
        self.Search_Input.setClearButtonEnabled(True)
        self.Search_Input.setMaximumWidth(200)
        
        self.year_filter = QComboBox()
        self.year_filter.addItem("All Years")
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT * FROM AcademicYear")
                years = c.fetchall()
                for year in years:
                    self.year_filter.addItem(year[1], year[0])
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load academic years: {str(e)}")
        
        self.year_filter.setCurrentIndex(0)
        self.year_filter.setMaximumWidth(150)
        filter_layout.addWidget(self.Search_Input)
        filter_layout.addWidget(self.year_filter)
        filter_layout.addStretch()
        
        self.results_label = QLabel("0 results")
        self.results_label.setStyleSheet("font-size: 14px; color: #718096;")
        filter_layout.addWidget(self.results_label)
        layout.addLayout(filter_layout)

        # Connect search input and year filter to reload data
        self.Search_Input.textChanged.connect(self.Load_Data)
        self.year_filter.currentIndexChanged.connect(self.Load_Data)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Academic Year", "Term", "Starting Date", "Ending Date", "Action"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        
        layout.addWidget(self.table)

        # Add Load data
        self.Load_Data()
        self.update_results_count()
    
        
    def Load_Data(self):
        try:
            with get_db_connection() as (conn, c):
                self.table.setRowCount(0)
                # Get selected year filter
                selected_year_id = self.year_filter.currentData()
                
                query = """
                    SELECT AcademicYear.Year, Term.TermName, Term.StartDate, Term.EndDate
                    FROM Term
                    JOIN AcademicYear ON Term.AcademicYear = AcademicYear.AcademicYearID
                """
                
                # Add search filter
                search_text = self.Search_Input.text().strip()
                params = ()
                if search_text:
                    query += " WHERE (AcademicYear.Year LIKE ? OR Term.TermName LIKE ?)"
                    params += (f"%{search_text}%", f"%{search_text}%")
                
                # Add year filter if selected
                if selected_year_id and selected_year_id != "All Years":
                    if "WHERE" in query:
                        query += " AND AcademicYear.AcademicYearID = ?"
                    else:
                        query += " WHERE AcademicYear.AcademicYearID = ?"
                    params += (selected_year_id,)
                
                c.execute(query, params)
                results = c.fetchall()

                # Check if year exists and has no terms
                if selected_year_id and selected_year_id != "All Years":
                    c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (selected_year_id,))
                    year_row = c.fetchone()
                    if year_row and not results:
                        self.table.setRowCount(0)
                        self.update_results_count()
                        msg_box = QMessageBox(self)
                        msg_box.setIcon(QMessageBox.Information)
                        msg_box.setWindowTitle("No Terms")
                        msg_box.setText(f"The year '{year_row[0]}' does not have any terms.")
                        msg_box.setStandardButtons(QMessageBox.Ok)
                        msg_box.setModal(False)
                        msg_box.show()
                        return

                for row_data in results:
                    row = self.table.rowCount()
                    self.table.insertRow(row)
                    self.table.setRowHeight(row, 50)
                    for col, data in enumerate(row_data):
                        self.table.setItem(row, col, QTableWidgetItem(str(data)))
                    
                    action_widget = QWidget()
                    action_layout = QHBoxLayout(action_widget)
                    action_layout.setContentsMargins(5, 0, 5, 0)
                    action_layout.setSpacing(5)
                    
                    edit_btn = QPushButton("Edit")
                    edit_btn.clicked.connect(lambda _, r=row: self.editTerm(r))
                    delete_btn = QPushButton("Delete")
                    delete_btn.clicked.connect(lambda _, r=row: self.deleteTerm(r))
                    delete_btn.setObjectName("delete")
                    
                    action_layout.addWidget(edit_btn)
                    action_layout.addWidget(delete_btn)
                    action_layout.addStretch()
                    self.table.setCellWidget(row, 4, action_widget)

                # Ensure header stretches after data reload
                self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
                self.update_results_count()
                
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load data: {str(e)}")

    def update_results_count(self):
        count = self.table.rowCount()
        self.results_label.setText(f"{count} results")
    
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
                    self.Load_Data()
                    
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
                    self.Load_Data()
                    
            except sqlite3.Error as e:
                error_label.setText(f"Error saving term: {str(e)}")

        save_btn.clicked.connect(save_term)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()
    
    def editTerm(self, row):
        # Get current term data from the table
        year = self.table.item(row, 0).text()
        term = self.table.item(row, 1).text()
        start = self.table.item(row, 2).text()
        end = self.table.item(row, 3).text()
        
        try:
            with get_db_connection() as (conn, c):
                # Get the term's database ID
                c.execute("""
                    SELECT Term.TermID, AcademicYear.AcademicYearID
                    FROM Term
                    JOIN AcademicYear ON Term.AcademicYear = AcademicYear.AcademicYearID
                    WHERE AcademicYear.Year=? AND Term.TermName=? AND Term.StartDate=? AND Term.EndDate=?
                """, (year, term, start, end))
                result = c.fetchone()
                
            if not result:
                QMessageBox.warning(self, "Error", "Could not find the selected term in the database.")
                return
                
            term_id, year_id = result

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
                        self.Load_Data()
                        
                except sqlite3.Error as e:
                    error_label.setText(f"Error updating term: {str(e)}")

            save_btn.clicked.connect(save_edit)
            cancel_btn.clicked.connect(dialog.reject)
            dialog.exec_()
            
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load term data: {str(e)}")
    
    def deleteTerm(self, row):
        # Get current term data from the table
        year = self.table.item(row, 0).text()
        term = self.table.item(row, 1).text()
        start = self.table.item(row, 2).text()
        end = self.table.item(row, 3).text()
        
        try:
            with get_db_connection() as (conn, c):
                # Get the term's database ID
                c.execute("""
                    SELECT Term.TermID
                    FROM Term
                    JOIN AcademicYear ON Term.AcademicYear = AcademicYear.AcademicYearID
                    WHERE AcademicYear.Year=? AND Term.TermName=? AND Term.StartDate=? AND Term.EndDate=?
                """, (year, term, start, end))
                result = c.fetchone()
                
            if not result:
                QMessageBox.warning(self, "Error", "Could not find the selected term in the database.")
                return
                
            term_id = result[0]

            # Precaution message
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
                    c.execute("SELECT 1 FROM Fees WHERE TermID = ?", (term_id,))
                    if c.fetchone():
                        QMessageBox.critical(self, "Restricted", "This term cannot be deleted because it is linked to one or more fees records.")
                        return

                    # Set Student.TermID to NULL for students assigned to this term
                    c.execute("UPDATE Student SET TermID = NULL WHERE TermID = ?", (term_id,))
                    # Now delete the term
                    c.execute("DELETE FROM Term WHERE TermID = ?", (term_id,))
                    conn.commit()
                    QMessageBox.information(self, "Deleted", f"Term '{term}' for academic year '{year}' has been deleted.\n\n"
                                                             "Any students previously assigned to this term now have no term assigned.")
                    self.Load_Data()
                    
            except sqlite3.Error as e:
                QMessageBox.critical(self, "Error", f"An error occurred while deleting the term: {str(e)}")
                
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load term data: {str(e)}")
    
    def editYear(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Academic Year")
        dialog.setFixedSize(350, 220)
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

        label = QLabel("Select Academic Year to Edit")
        layout.addWidget(label)

        year_combo = QComboBox()
        
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear")
                years = c.fetchall()
                year_combo.addItem("Select Year", None)
                for y in years:
                    year_combo.addItem(y[1], y[0])
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load academic years: {str(e)}")
            
        layout.addWidget(year_combo)

        new_label = QLabel("New Year Name")
        layout.addWidget(new_label)

        year_input = QLineEdit()
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
            year_id = year_combo.currentData()
            new_year_name = year_input.text().strip()
            if not year_id:
                error_label.setText("Please select a year to edit.")
                return
            if not new_year_name:
                error_label.setText("Year name cannot be empty.")
                return
            
            try:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT 1 FROM AcademicYear WHERE Year = ? AND AcademicYearID != ?", (new_year_name, year_id))
                    if c.fetchone():
                        error_label.setText("This academic year already exists.")
                        return
                    
                    c.execute("UPDATE AcademicYear SET Year = ? WHERE AcademicYearID = ?", (new_year_name, year_id))
                    conn.commit()
                    dialog.accept()
                    QMessageBox.information(self, "Success", f"Academic Year updated successfully.")
                    self.Load_Data()
                    
            except sqlite3.Error as e:
                error_label.setText(f"Error updating year: {str(e)}")

        save_btn.clicked.connect(save_edit)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def deleteYear(self):
        # Dialog for deleting academic year
        dialog = QDialog(self)
        dialog.setWindowTitle("Delete Academic Year")
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
            QComboBox {
                background-color: #F7FAFC;
                border: 1px solid #E0E3E7;
                padding: 10px;
                border-radius: 5px;
                font-size: 15px;
            }
            QPushButton {
                background-color: #E53E3E;
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
                background-color: #C53030;
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

        label = QLabel("Select Academic Year to Delete")
        layout.addWidget(label)

        year_combo = QComboBox()
        
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear")
                years = c.fetchall()
                year_combo.addItem("Select Year", None)
                for y in years:
                    year_combo.addItem(y[1], y[0])
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load academic years: {str(e)}")
            
        layout.addWidget(year_combo)

        error_label = QLabel("")
        error_label.setObjectName("error")
        layout.addWidget(error_label)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        delete_btn = QPushButton("Delete")
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("cancel")
        btn_layout.addWidget(delete_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)

        def do_delete():
            year_id = year_combo.currentData()
            if not year_id:
                error_label.setText("Please select a year to delete.")
                return
            
            try:
                with get_db_connection() as (conn, c):
                    # Check for child relations in Term
                    c.execute("SELECT 1 FROM Term WHERE AcademicYear = ?", (year_id,))
                    if c.fetchone():
                        error_label.setText("Cannot delete: This year has related terms.")
                        return
                    
                    # Check for child relations in Class
                    c.execute("SELECT 1 FROM Class WHERE AcademicYear = ?", (year_id,))
                    if c.fetchone():
                        error_label.setText("Cannot delete: This year has related classes.")
                        return

                    # Get year name for message
                    year_name = year_combo.currentText()
                    msg = QMessageBox(self)
                    msg.setIcon(QMessageBox.Warning)
                    msg.setWindowTitle("Delete Academic Year")
                    msg.setText(
                        f"Are you sure you want to delete the academic year '{year_name}'?\n\n"
                        "This action cannot be undone.\n\n"
                        "Note:\n"
                        "- If this year is linked to any terms or classes, deletion is not allowed."
                    )
                    msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                    ret = msg.exec_()
                    if ret != QMessageBox.Yes:
                        return

                    c.execute("DELETE FROM AcademicYear WHERE AcademicYearID = ?", (year_id,))
                    conn.commit()
                    dialog.accept()
                    QMessageBox.information(self, "Deleted", f"Academic Year '{year_name}' has been deleted successfully.")
                    self.Load_Data()
                    
            except sqlite3.Error as e:
                error_label.setText(f"Error deleting year: {str(e)}")

        delete_btn.clicked.connect(do_delete)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def handle_year_action(self, index):
        action = self.year_action_combo.itemData(index)
        if action == "add":
            self.addYear()
        elif action == "edit":
            self.editYear()
        elif action == "delete":
            self.deleteYear()
        self.year_action_combo.setCurrentIndex(0)

class Dashboard(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addStretch()
        label = QLabel("Dashboard Page")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2D3748;")
        layout.addWidget(label)
        layout.addStretch()

class Class(QWidget):
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
            QLineEdit {
            background-color: #FFFFFF;
            border: 1px solid #E0E3E7;
            padding: 8px;
            border-radius: 5px;
            font-size: 14px;
            }
            QComboBox {
            background-color: #FFFFFF;
            border: 1px solid #E0E3E7;
            padding: 8px;
            border-radius: 5px;
            font-size: 14px;
            }
            QTableWidget {
            background-color: #FFFFFF;
            border: 1px solid #E0E3E7;
            border-radius: 5px;
            gridline-color: #E0E3E7;
            selection-background-color: #0c3157;
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
            QTableWidget QPushButton {
            background-color: transparent;
            color: #4299E1;
            padding: 5px 10px;
            border: 1px solid #4299E1;
            border-radius: 3px;
            font-size: 12px;
            }
            QTableWidget QPushButton:hover {
            background-color: #EDF2F7;
            }
            QTableWidget QPushButton#delete {
            color: #E53E3E;
            border: 1px solid #E53E3E;
            }
            QTableWidget QPushButton#delete:hover {
            background-color: #FED7D7;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(15)

        # Header row
        header_layout = QHBoxLayout()
        title = QLabel("Class Management")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        add_class_btn = QPushButton("Add Class")
        add_class_btn.clicked.connect(self.addClass)
        header_layout.addWidget(add_class_btn)
        layout.addLayout(header_layout)

        # Filter row
        filter_layout = QHBoxLayout()
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
            JOIN AcademicYear ON Class.AcademicYear = AcademicYear.AcademicYearID
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
            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(5, 0, 5, 0)
            action_layout.setSpacing(5)
            edit_btn = QPushButton("Edit")
            edit_btn.clicked.connect(lambda _, cid=row_data[0]: self.editClass(cid))
            delete_btn = QPushButton("Delete")
            delete_btn.clicked.connect(lambda _, cid=row_data[0]: self.deleteClass(cid))
            delete_btn.setObjectName("delete")
            action_layout.addWidget(edit_btn)
            action_layout.addWidget(delete_btn)
            action_layout.addStretch()
            self.table.setCellWidget(row, 2, action_widget)
        # self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.update_results_count()

    def update_results_count(self):
        count = self.table.rowCount()
        self.results_label.setText(f"{count} results")

    def addClass(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Class")
        dialog.setFixedSize(350, 220)
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
        with get_db_connection() as (conn, c):
            c.execute("SELECT ClassName, AcademicYear FROM Class WHERE ClassID=?", (class_id,))
            result = c.fetchone()
        if not result:
            QMessageBox.warning(self, "Error", "Could not find the selected class in the database.")
            return
        class_name, year_id = result

        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Class")
        dialog.setFixedSize(350, 220)
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
        with get_db_connection() as (conn, c):
            c.execute("SELECT AcademicYearID, Year FROM AcademicYear")
            years = c.fetchall()
        for y in years:
            year_combo.addItem(y[1], y[0])
        idx = [y[0] for y in years].index(year_id) if year_id in [y[0] for y in years] else 0
        year_combo.setCurrentIndex(idx)
        layout.addWidget(year_combo)

        # Class Name
        class_label = QLabel("Class Name")
        layout.addWidget(class_label)
        class_input = QLineEdit(class_name)
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

        def save_edit():
            new_year_id = year_combo.currentData()
            new_class_name = class_input.text().strip()
            if not new_year_id or not new_class_name:
                error_label.setText("Year and Class Name are required.")
                return
            with get_db_connection() as (conn, c):
                c.execute(
                    "SELECT 1 FROM Class WHERE ClassName=? AND AcademicYear=? AND ClassID<>?",
                    (new_class_name, new_year_id, class_id)
                )
                if c.fetchone():
                    error_label.setText("This class already exists for the selected year.")
                    return
                try:
                    c.execute(
                        "UPDATE Class SET ClassName=?, AcademicYear=? WHERE ClassID=?",
                        (new_class_name, new_year_id, class_id)
                    )
                    conn.commit()
                    dialog.accept()
                    QMessageBox.information(self, "Success", f"Class updated successfully.")
                    self.Load_Data()
                except Exception as e:
                    error_label.setText("Error updating class.")

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
            QPushButton#nav_btn {
                background-color: transparent;
                color: #4299E1;
                padding: 5px 10px;
                font-size: 16px;
            }
            QPushButton#nav_btn:selected {
                background-color: #EDF2F7;
                font-weight: bold;
            }
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QCheckBox {
                spacing: 5px;
                font-size: 14px;
                color: #2D3748;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                border-radius: 3px;
                border: 1px solid #E0E3E7;
            }
            QCheckBox::indicator:checked {
                background-color: #4299E1;
                border: 1px solid #4299E1;
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
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(15)

        # Navigation links
        nav_layout = QHBoxLayout()
        nav_layout.setSpacing(20)
        self.management_btn = QPushButton("Fees Management")
        self.management_btn.setObjectName("nav_btn")
        self.management_btn.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        self.payment_btn = QPushButton("Fees Payment Insight")
        self.payment_btn.setObjectName("nav_btn")
        self.payment_btn.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        self.rawFeePayments = QPushButton("Payments History")
        self.rawFeePayments.setObjectName("nav_btn")
        self.rawFeePayments.clicked.connect(lambda: self.stack.setCurrentIndex(2))
        nav_layout.addWidget(self.management_btn)
        nav_layout.addWidget(self.payment_btn)
        nav_layout.addWidget(self.rawFeePayments)
        nav_layout.addStretch()
        layout.addLayout(nav_layout)

        # Stacked widget for sub-pages
        self.stack = QStackedWidget()
        self.management_page = FeesManagementPage()
        self.payment_page = FeesPaymentPage()
        self.PaymentHistoryPage = PaymentHistoryPage()
        self.stack.addWidget(self.management_page)
        self.stack.addWidget(self.payment_page)
        self.stack.addWidget(self.PaymentHistoryPage)
        layout.addWidget(self.stack)

        # Initial selection
        self.management_btn.setStyleSheet("font-weight: bold; background-color: #EDF2F7;")
        self.stack.setCurrentIndex(0)

class FeesManagementPage(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Header
        header_layout = QHBoxLayout()
        title = QLabel("Fees Management")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        layout.addLayout(header_layout)

        # Filters and add buttons
        filter_layout = QHBoxLayout()
        search_input = QLineEdit()
        search_input.setPlaceholderText("Search...")
        search_input.setMaximumWidth(200)
        # Add more filters as needed, e.g., class_filter, term_filter
        filter_layout.addWidget(search_input)
        filter_layout.addStretch()
        add_btn = QPushButton("Add Fee Structure")
        add_btn.clicked.connect(self.show_add_fee_dialog)
        import_btn = QPushButton("Import Fees")
        # import_btn.clicked.connect(self.show_import_wizard)
        filter_layout.addWidget(add_btn)
        filter_layout.addWidget(import_btn)
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
                search_text = self.Search_Input.text().strip() if hasattr(self, 'Search_Input') else ""
                if search_text:
                    filters.append("(ay.Year LIKE ? OR t.TermName LIKE ? OR c.ClassName LIKE ? OR f.Amount LIKE ?)")
                    params.extend([f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%"])
                
                # Apply filters if any
                if filters:
                    query += " WHERE " + " AND ".join(filters)
                
                # Add ordering
                query += " ORDER BY ay.Year, t.TermName, c.ClassName"
                
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
        """Add edit and delete buttons to table row for fees"""
        actions_widget = QWidget()
        actions_layout = QHBoxLayout(actions_widget)
        actions_layout.setContentsMargins(5, 5, 5, 5)
        actions_layout.setSpacing(5)
        
        # Edit button
        edit_btn = QPushButton("Edit")
        edit_btn.setFixedSize(50, 25)
        edit_btn.setStyleSheet("""
            QPushButton {
                background-color: #48BB78;
                color: white;
                border: none;
                border-radius: 3px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #38A169;
            }
            QPushButton:pressed {
                background-color: #2F855A;
            }
        """)
        edit_btn.clicked.connect(lambda: self.edit_fee(fees_id))
        
        # Delete button
        delete_btn = QPushButton("Delete")
        delete_btn.setFixedSize(50, 25)
        delete_btn.setStyleSheet("""
            QPushButton {
                background-color: #F56565;
                color: white;
                border: none;
                border-radius: 3px;
                font-size: 11px;
            }
            QPushButton:hover {
                background-color: #E53E3E;
            }
            QPushButton:pressed {
                background-color: #C53030;
            }
        """)
        delete_btn.clicked.connect(lambda: self.delete_fee(fees_id))
        
        actions_layout.addWidget(edit_btn)
        actions_layout.addWidget(delete_btn)
        actions_layout.addStretch()
        
        # Set the widget in the Actions column (column 4)
        self.table.setCellWidget(row, 4, actions_widget)

    def edit_fee(self, fees_id):
        """Edit fee record"""
        try:
            with get_db_connection() as (conn, c):
                # Get current fee data
                c.execute("""
                    SELECT f.FeesID, f.Class, f.Term, f.Amount, 
                        c.ClassName, t.TermName, ay.Year,ay.AcademicYearID
                    FROM Fees f
                    JOIN Class c ON f.Class = c.ClassID
                    JOIN Term t ON f.Term = t.TermID
                    JOIN AcademicYear ay ON c.AcademicYear = ay.AcademicYearID
                    WHERE f.FeesID = ?
                """, (fees_id,))
                fee_data = c.fetchone()
            
            if fee_data:
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
                    c.execute(f"SELECT ClassID, ClassName FROM Class  WHERE AcademicYear ={fee_data[7]} ORDER BY ClassName")
                    classes = c.fetchall()
                    for class_id, class_name in classes:
                        class_combo.addItem(class_name, class_id)
                    
                    # Load terms
                    c.execute(f"SELECT TermID, TermName FROM Term  WHERE AcademicYear ={fee_data[7]} ORDER BY TermName")
                    terms = c.fetchall()
                    for term_id, term_name in terms:
                        term_combo.addItem(term_name, term_id)
                
                # Set current values
                class_combo.setCurrentIndex(class_combo.findData(fee_data[1]))
                term_combo.setCurrentIndex(term_combo.findData(fee_data[2]))
                
                form_layout.addRow("Class:", class_combo)
                form_layout.addRow("Term:", term_combo)
                form_layout.addRow("Amount:", amount_edit)
                
                layout.addLayout(form_layout)
                
                button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
                button_box.accepted.connect(dialog.accept)
                button_box.rejected.connect(dialog.reject)
                layout.addWidget(button_box)
                
                if dialog.exec_() == QDialog.Accepted:
                    try:
                        with get_db_connection() as (conn, c):
                            c.execute("""
                                UPDATE Fees 
                                SET Class = ?, Term = ?, Amount = ?
                                WHERE FeesID = ?
                            """, (
                                class_combo.currentData(),
                                term_combo.currentData(),
                                float(amount_edit.text()),
                                fees_id
                            ))
                            conn.commit()
                            QMessageBox.information(self, "Success", "Fee updated successfully")
                            self.load_data()
                    except Exception as e:
                        QMessageBox.critical(self, "Error", f"Failed to update fee: {str(e)}")
                        
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
            print(f"Error populating academic years: {e}")

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
            print(f"Error updating class/term combos: {e}")

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

        # Header
        header_layout = QHBoxLayout()
        title = QLabel("Fees Payment")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        layout.addLayout(header_layout)

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
            print(f"Error loading filter options: {e}")

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
                
        except Exception as e:
            print(f"Error loading data: {e}")

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
                        if amount_paid <= 0:
                            errors.append(f"Row {row}: Amount must be greater than zero.")
                            continue
                        if not transaction_date:
                            errors.append(f"Row {row}: Transaction date required.")
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
        header_layout.addWidget(export_btn)
        
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
        
        # Student filter
        self.student_combo = QComboBox()
        self.student_combo.addItem("All Students", None)
        self.student_combo.currentIndexChanged.connect(self.load_data)
        
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
        filter_layout.addWidget(QLabel("Student:"))
        filter_layout.addWidget(self.student_combo)
        filter_layout.addWidget(QLabel("Class:"))
        filter_layout.addWidget(self.class_combo)
        filter_layout.addWidget(QLabel("Term:"))
        filter_layout.addWidget(self.term_combo)
        
        layout.addLayout(filter_layout)

        # Table for raw payment data
        self.table = QTableWidget()
        self.table.setColumnCount(10)
        self.table.setHorizontalHeaderLabels([
            "Payment ID", "Student", "Student ID", "Class", "Term", 
            "Year", "Amount Paid", "Transaction Date", "Reference", "Recorded On"
        ])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSortingEnabled(True)
        
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
        
        layout.addWidget(self.table)
        
        # Summary label
        self.summary_label = QLabel()
        self.summary_label.setStyleSheet("font-weight: bold; color: #2D3748;")
        layout.addWidget(self.summary_label)

    def load_filter_options(self):
        """Load filter options from database"""
        try:
            with get_db_connection() as (conn, c):
                # Load students
                c.execute("SELECT StudentID, Student_Name FROM Students ORDER BY Student_Name")
                students = c.fetchall()
                for student_id, student_name in students:
                    self.student_combo.addItem(f"{student_name} ({student_id})", student_id)
                
                # Load classes
                c.execute("SELECT DISTINCT ClassName FROM Class ORDER BY ClassName")
                classes = c.fetchall()
                for class_name in classes:
                    self.class_combo.addItem(class_name, class_name)
                
                # Load terms
                c.execute("SELECT DISTINCT TermName FROM Term ORDER BY TermName")
                terms = c.fetchall()
                for term_name in terms:
                    self.term_combo.addItem(term_name, term_name)
                    
        except Exception as e:
            print(f"Error loading filter options: {e}")

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
                    fp.DoneOn
                FROM FeesPayment fp
                JOIN StudentTerm st ON fp.StudentTermID = st.StudentTermID
                JOIN Students s ON st.StudentID = s.StudentID
                JOIN Class c ON st.ClassID = c.ClassID
                JOIN Term t ON st.TermID = t.TermID
                JOIN AcademicYear ay ON c.AcademicYear = ay.AcademicYearID
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
                
                # Apply student filter
                selected_student = self.student_combo.currentData()
                if selected_student:
                    filters.append("s.StudentID = ?")
                    params.append(selected_student)
                
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
                    
                    for col_idx, data in enumerate(row_data):
                        item = QTableWidgetItem(str(data) if data is not None else "")
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        
                        # Color code based on amount
                        if col_idx == 6:  # AmountPaid column
                            item.setText(f"{float(data or 0):,.2f}")
                            if float(data or 0) > 0:
                                item.setForeground(QColor(0, 100, 0))  # Dark green
                        
                        self.table.setItem(row_idx, col_idx, item)
                
                # Update summary
                self.summary_label.setText(
                    f"Total Records: {len(rows)} | "
                    f"Total Amount Paid: {total_amount:,.2f} | "
                    f"Date Range: {date_from} to {date_to}"
                )
                
        except Exception as e:
            print(f"Error loading payment history: {e}")
            QMessageBox.critical(self, "Error", f"Failed to load payment history: {str(e)}")

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
                
                # Get headers
                for col in range(self.table.columnCount()):
                    headers.append(self.table.horizontalHeaderItem(col).text())
                
                # Get data
                for row in range(self.table.rowCount()):
                    row_data = []
                    for col in range(self.table.columnCount()):
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
            for col in range(self.table.columnCount()):
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
                JOIN Fees f ON f.Class = c.ClassID AND f.Term = t.TermID
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
        self.Procurement = ProcurementManagement()
        self.Expense = ExpenseManagement()
        self.addTab(self.Dashboard,"Dashboard & Insights")
        self.addTab(self.Inventory,"Inventory")
        self.addTab(self.Procurement,"Procurement")
        self.addTab(self.Expense,"Expense")

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
        self.in_canvas = self.create_pie_canvas("Stock In Distribution (Current Month)")
        self.out_canvas = self.create_pie_canvas("Stock Out Distribution (Current Month)")
        charts_layout.addWidget(self.in_canvas)
        charts_layout.addWidget(self.out_canvas)
        left_layout.addLayout(charts_layout)

        # Last 5 inventory table
        inventory_title = QLabel("Recent Inventory Entries")
        inventory_title.setStyleSheet("font-size: 16px; font-weight: bold;")
        left_layout.addWidget(inventory_title)
        self.inventory_table = QTableWidget()
        self.inventory_table.setColumnCount(6)
        self.inventory_table.setHorizontalHeaderLabels(["ID", "Uniform", "Type", "Quantity", "Date", "Reference"])
        self.inventory_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.inventory_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.inventory_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.inventory_table.setShowGrid(True)
        self.inventory_table.verticalHeader().setVisible(False)
        self.inventory_table.setMaximumHeight(200)
        left_layout.addWidget(self.inventory_table)

        # Right column: Uniform list and add button
        right_layout = QVBoxLayout()
        right_layout.setSpacing(15)

        uniform_title = QLabel("Uniform List")
        uniform_title.setStyleSheet("font-size: 16px; font-weight: bold;")
        right_layout.addWidget(uniform_title)
        
        # Add filter for uniform list
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search uniforms...")
        self.search_input.textChanged.connect(self.load_data)
        filter_layout.addWidget(self.search_input)
        right_layout.addLayout(filter_layout)
        
        self.uniform_table = QTableWidget()
        self.uniform_table.setColumnCount(5)
        self.uniform_table.setHorizontalHeaderLabels(["ID", "Name", "Price", "Quantity", "Actions"])
        self.uniform_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.uniform_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.uniform_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.uniform_table.setShowGrid(True)
        self.uniform_table.verticalHeader().setVisible(False)
        right_layout.addWidget(self.uniform_table)

        add_uniform_btn = QPushButton("Add New Uniform")
        add_uniform_btn.clicked.connect(self.show_add_uniform_dialog)
        right_layout.addWidget(add_uniform_btn, alignment=Qt.AlignBottom)

        main_layout.addLayout(left_layout, 2)  # 2/3 of space
        main_layout.addLayout(right_layout, 1)  # 1/3 of space

        self.load_data()

    def create_pie_canvas(self, title):
        fig = Figure(figsize=(5, 4), dpi=100)
        canvas = FigureCanvas(fig)
        canvas.setMinimumHeight(300)
        return canvas

    def load_data(self):
        current_month = datetime.now().strftime('%Y-%m')
        search_text = self.search_input.text().strip()
        
        try:
            with get_db_connection() as (conn, c):
                
                # Get stock in distribution by uniform type for current month
                c.execute("""
                    SELECT u.UniformName, SUM(ui.Quantity) as TotalQuantity
                    FROM UniformInventory ui
                    JOIN Uniform u ON ui.UniformID = u.UniformId
                    WHERE ui.Type = 'In' AND strftime('%Y-%m', ui.Date) = ?
                    GROUP BY u.UniformId
                    ORDER BY TotalQuantity DESC
                """, (current_month,))
                stock_in_data = c.fetchall()
                
                # Get stock out distribution by uniform type for current month
                c.execute("""
                    SELECT u.UniformName, SUM(ui.Quantity) as TotalQuantity
                    FROM UniformInventory ui
                    JOIN Uniform u ON ui.UniformID = u.UniformId
                    WHERE ui.Type = 'Out' AND strftime('%Y-%m', ui.Date) = ?
                    GROUP BY u.UniformId
                    ORDER BY TotalQuantity DESC
                """, (current_month,))
                stock_out_data = c.fetchall()
                
                # Update in pie chart
                self.update_pie_chart(self.in_canvas, "Stock In Distribution (Current Month)", 
                                     stock_in_data, '#2E8B57')
                
                # Update out pie chart
                self.update_pie_chart(self.out_canvas, "Stock Out Distribution (Current Month)", 
                                     stock_out_data, '#CD5C5C')
                
                # Recent inventory entries (last 10 instead of 5 for better overview)
                c.execute("""
                    SELECT ui.UniformInventoryID, u.UniformName, ui.Type, ui.Quantity, ui.Date, ui.ReferencesNotes
                    FROM UniformInventory ui
                    LEFT JOIN Uniform u ON ui.UniformID = u.UniformId
                    ORDER BY ui.Date DESC, ui.UniformInventoryID DESC
                    LIMIT 10
                """)
                inventory_rows = c.fetchall()
                self.inventory_table.setRowCount(len(inventory_rows))
                for i, row in enumerate(inventory_rows):
                    for j in range(len(row)):
                        self.inventory_table.setItem(i, j, QTableWidgetItem(str(row[j])))
                
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
            QMessageBox.warning(self, "Error", f"Failed to load dashboard data: {str(e)}")

    def update_pie_chart(self, canvas, title, data, color_scheme):
        """Update pie chart with distribution data"""
        fig = canvas.figure
        fig.clear()
        
        if not data:
            ax = fig.add_subplot(111)
            ax.text(0.5, 0.5, 'No data available', ha='center', va='center', transform=ax.transAxes)
            ax.set_title(title)
            canvas.draw()
            return
        
        labels = [row[0] for row in data]
        quantities = [row[1] for row in data]
        total = sum(quantities)
        percentages = [q/total*100 for q in quantities]
        
        ax = fig.add_subplot(111)
        wedges, texts, autotexts = ax.pie(quantities, labels=labels, autopct='%1.1f%%',
                                         startangle=90, colors=self.generate_colors(len(labels), color_scheme))
        
        # Improve label appearance
        for text in texts:
            text.set_fontsize(8)
        for autotext in autotexts:
            autotext.set_fontsize(8)
            autotext.set_color('white')
            autotext.set_fontweight('bold')
        
        ax.set_title(title, fontsize=10, fontweight='bold')
        ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle
        canvas.draw()

    def generate_colors(self, n, base_color):
        """Generate a color palette based on a base color"""
        # For simplicity, using matplotlib's default color cycle
        import matplotlib.pyplot as plt
        return plt.cm.Set3(range(n))

    def add_uniform_action_buttons(self, row, uniform_id):
        actions_widget = QWidget()
        actions_layout = QHBoxLayout(actions_widget)
        actions_layout.setContentsMargins(5, 2, 5, 2)
        actions_layout.setSpacing(3)

        edit_btn = QPushButton("Edit")
        edit_btn.setFixedWidth(50)
        edit_btn.setStyleSheet("font-size: 11px; padding: 3px;")
        edit_btn.clicked.connect(lambda: self.edit_uniform(uniform_id))
        
        delete_btn = QPushButton("Delete")
        delete_btn.setFixedWidth(50)
        delete_btn.setStyleSheet("font-size: 11px; padding: 3px;")
        delete_btn.clicked.connect(lambda: self.delete_uniform(uniform_id))

        actions_layout.addWidget(edit_btn)
        actions_layout.addWidget(delete_btn)
        actions_layout.addStretch()

        self.uniform_table.setCellWidget(row, 4, actions_widget)

    def show_add_uniform_dialog(self):
        """Show dialog to add new uniform"""
        dialog = UniformDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.load_data()

    def edit_uniform(self, uniform_id):
        """Edit existing uniform"""
        dialog = UniformDialog(self, uniform_id)
        if dialog.exec_() == QDialog.Accepted:
            self.load_data()

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
                    self.load_data()
                    
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to delete uniform: {str(e)}")

class InventoryManagement(QWidget):
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
            QLineEdit, QDateEdit {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
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
        metrics_layout.addWidget(print_btn)
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
                    # Add actions
                    self.add_action_buttons(i, row[0])
                
                self.results_label.setText(f"{len(rows)} results retrieved")
                self.high_in_label.setText(f"High Qty In: {high_in}")
                self.high_out_label.setText(f"High Qty Out: {high_out}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")

    def add_action_buttons(self, row, inv_id):
        actions_widget = QWidget()
        actions_layout = QHBoxLayout(actions_widget)
        actions_layout.setContentsMargins(2, 2, 2, 2)
        actions_layout.setSpacing(2)

        edit_btn = QPushButton("Edit")
        edit_btn.setFixedWidth(60)
        edit_btn.clicked.connect(lambda: self.update_inventory(inv_id))
        delete_btn = QPushButton("Delete")
        delete_btn.setFixedWidth(60)
        delete_btn.clicked.connect(lambda: self.delete_inventory(inv_id))

        actions_layout.addWidget(edit_btn)
        actions_layout.addWidget(delete_btn)
        actions_layout.addStretch()

        self.table.setCellWidget(row, 7, actions_widget)

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
                print(f"Error fetching available quantity: {e}")
        
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
                print(f"Error fetching uniform price: {e}")
        
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

class ProcurementDialog(QDialog):
    def __init__(self, parent=None, procurement_id=None):
        super().__init__(parent)
        self.procurement_id = procurement_id
        self.setWindowTitle("Add New Procurement" if procurement_id is None else "Edit Procurement")
        self.setModal(True)
        
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()
        
        self.name_edit = QLineEdit()
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.unit_edit = QLineEdit()
        self.price_edit = QLineEdit()
        self.ref_edit = QLineEdit()
        
        form_layout.addRow("Procurement Name:", self.name_edit)
        form_layout.addRow("Date:", self.date_edit)
        form_layout.addRow("Unit:", self.unit_edit)
        form_layout.addRow("Price:", self.price_edit)
        form_layout.addRow("Reference:", self.ref_edit)
        
        layout.addLayout(form_layout)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        # If editing, load existing data
        if procurement_id is not None:
            self.load_procurement_data()
    
    def load_procurement_data(self):
        """Load procurement data for editing"""
        with get_db_connection() as (conn, c):
            c.execute("SELECT * FROM Procurement WHERE PRO_ID = ?", (self.procurement_id,))
            procurement = c.fetchone()
            
            if procurement:
                self.name_edit.setText(procurement['PRO_NAME'])
                self.date_edit.setDate(QDate.fromString(procurement['PRO_DATE'], "yyyy-MM-dd"))
                self.unit_edit.setText(procurement['PRO_UNIT'])
                self.price_edit.setText(str(procurement['PRO_PRICE']))
                self.ref_edit.setText(procurement['PRO_REFERENCE'])
    
    def validate_and_save(self):
        """Validate input and save procurement data"""
        name = self.name_edit.text().strip()
        date = self.date_edit.date().toString("yyyy-MM-dd")
        unit = self.unit_edit.text().strip()
        price_str = self.price_edit.text().strip()
        ref = self.ref_edit.text().strip()
        
        # Validation
        if not all([name, date, unit, price_str, ref]):
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
        with get_db_connection() as (conn, c):
            if self.procurement_id is None:
                # Insert new record
                c.execute('''
                    INSERT INTO Procurement (PRO_NAME, PRO_DATE, PRO_UNIT, PRO_PRICE, PRO_REFERENCE)
                    VALUES (?, ?, ?, ?, ?)
                ''', (name, date, unit, price, ref))
            else:
                # Update existing record
                c.execute('''
                    UPDATE Procurement 
                    SET PRO_NAME=?, PRO_DATE=?, PRO_UNIT=?, PRO_PRICE=?, PRO_REFERENCE=?
                    WHERE PRO_ID=?
                ''', (name, date, unit, price, ref, self.procurement_id))
            
            conn.commit()
        
        self.accept()

class ProcurementManagement(QWidget):
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
            QLineEdit, QDateEdit {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
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
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Header
        header_layout = QHBoxLayout()
        title = QLabel("Procurement Management")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        layout.addLayout(header_layout)

        # First row: Filters (Search, From Date, To Date)
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by name or reference...")
        self.search_input.setMaximumWidth(300)
        self.search_input.textChanged.connect(self.load_data)

        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDate(QDate.currentDate().addMonths(-1))
        self.from_date.dateChanged.connect(self.load_data)

        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDate(QDate.currentDate())
        self.to_date.dateChanged.connect(self.load_data)

        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(QLabel("From Date:"))
        filter_layout.addWidget(self.from_date)
        filter_layout.addWidget(QLabel("To Date:"))
        filter_layout.addWidget(self.to_date)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        # Second row: Results, Total Amount, Print, New Procurement buttons
        metrics_layout = QHBoxLayout()
        self.results_label = QLabel("0 results retrieved")
        self.results_label.setStyleSheet("font-size: 14px; color: #718096;")

        self.total_amount_label = QLabel("Total Amount: 0.00")
        self.total_amount_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #2D3748;")

        print_btn = QPushButton("Print")
        print_btn.clicked.connect(self.print_table)

        new_procurement_btn = QPushButton("New Procurement")
        new_procurement_btn.clicked.connect(self.show_add_procurement_dialog)

        metrics_layout.addWidget(self.results_label)
        metrics_layout.addWidget(self.total_amount_label)
        metrics_layout.addStretch()
        metrics_layout.addWidget(print_btn)
        metrics_layout.addWidget(new_procurement_btn)
        layout.addLayout(metrics_layout)

        # Third row: Table
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["ID", "Name", "Date", "Unit", "Price", "Reference", "Actions"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        header = self.table.horizontalHeader()
        for col in range(6):
            header.setSectionResizeMode(col, QHeaderView.Stretch)
        header.setSectionResizeMode(6, QHeaderView.Fixed)
        header.resizeSection(6, 150)  # Fixed width for actions
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # Initialize database and load data
        self.load_data()

    def load_data(self):
        """Load procurement data from database with filtering"""
        search_text = self.search_input.text().strip()
        from_date = self.from_date.date().toString("yyyy-MM-dd")
        to_date = self.to_date.date().toString("yyyy-MM-dd")
        
        # Build query with filters
        query = """
            SELECT * FROM Procurement 
            WHERE PRO_DATE BETWEEN ? AND ?
        """
        params = [from_date, to_date]
        
        if search_text:
            query += " AND (PRO_NAME LIKE ? OR PRO_REFERENCE LIKE ?)"
            params.extend([f"%{search_text}%", f"%{search_text}%"])
        
        query += " ORDER BY PRO_DATE DESC"
        
        with get_db_connection() as (conn, c):
            c.execute(query, params)
            data = c.fetchall()
        
        # Update table
        self.table.setRowCount(len(data))
        total_amount = 0.0
        
        for i, row in enumerate(data):
            self.table.setItem(i, 0, QTableWidgetItem(str(row['PRO_ID'])))
            self.table.setItem(i, 1, QTableWidgetItem(row['PRO_NAME']))
            self.table.setItem(i, 2, QTableWidgetItem(row['PRO_DATE']))
            self.table.setItem(i, 3, QTableWidgetItem(row['PRO_UNIT']))
            self.table.setItem(i, 4, QTableWidgetItem(f"{row['PRO_PRICE']:.2f}"))
            self.table.setItem(i, 5, QTableWidgetItem(row['PRO_REFERENCE']))
            
            # Calculate total amount for all items
            total_amount += row['PRO_PRICE']
            
            # Add action buttons
            self.add_action_buttons(i, row['PRO_ID'])
        
        # Update metrics
        self.total_amount_label.setText(f"Total Amount: {total_amount:.2f}")
        self.results_label.setText(f"{len(data)} results retrieved")

    def add_action_buttons(self, row, pro_id):
        actions_widget = QWidget()
        actions_widget.setMinimumWidth(140)
        actions_layout = QHBoxLayout(actions_widget)
        actions_layout.setContentsMargins(2, 2, 2, 2)
        actions_layout.setSpacing(2)

        edit_btn = QPushButton("Edit")
        edit_btn.setFixedWidth(60)
        edit_btn.clicked.connect(lambda: self.edit_procurement(pro_id))
        delete_btn = QPushButton("Delete")
        delete_btn.setFixedWidth(60)
        delete_btn.clicked.connect(lambda: self.delete_procurement(pro_id))

        actions_layout.addWidget(edit_btn)
        actions_layout.addWidget(delete_btn)
        actions_layout.addStretch()

        self.table.setCellWidget(row, 6, actions_widget)

    def print_table(self):
        """Print functionality placeholder"""
        QMessageBox.information(self, "Print", "Printing procurement table... (placeholder)")

    def show_add_procurement_dialog(self):
        """Show dialog to add new procurement"""
        dialog = ProcurementDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.load_data()

    def edit_procurement(self, pro_id):
        """Edit existing procurement"""
        dialog = ProcurementDialog(self, pro_id)
        if dialog.exec_() == QDialog.Accepted:
            self.load_data()

    def delete_procurement(self, pro_id):
        """Delete procurement after confirmation"""
        reply = QMessageBox.question(
            self, 
            'Confirm Delete', 
            f'Are you sure you want to delete procurement #{pro_id}?', 
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            with get_db_connection() as (conn, c):
                c.execute("DELETE FROM Procurement WHERE PRO_ID = ?", (pro_id,))
                conn.commit()
            
            QMessageBox.information(self, "Success", "Procurement deleted successfully")
            self.load_data()

class ExpenseDialog(QDialog):
    def __init__(self, parent=None, expense_id=None):
        super().__init__(parent)
        self.expense_id = expense_id
        self.setWindowTitle("Add New Expense" if expense_id is None else "Edit Expense")
        self.setModal(True)
        
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()
        
        self.category_combo = QComboBox()
        self.category_combo.addItems(["Salaries", "Rent", "Utilities", "Office Supplies", 
                                     "Marketing", "Travel", "Training", "Equipment", "Other"])
        
        self.amount_edit = QLineEdit()
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.desc_edit = QLineEdit()
        
        form_layout.addRow("Category:", self.category_combo)
        form_layout.addRow("Amount:", self.amount_edit)
        form_layout.addRow("Date:", self.date_edit)
        form_layout.addRow("Description:", self.desc_edit)
        
        layout.addLayout(form_layout)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_save)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        # If editing, load existing data
        if expense_id is not None:
            self.load_expense_data()
    
    def load_expense_data(self):
        """Load expense data for editing"""
        with get_db_connection() as (conn, c):
            c.execute("SELECT * FROM Expenses WHERE ExpID = ?", (self.expense_id,))
            expense = c.fetchone()
            
            if expense:
                # Set category
                index = self.category_combo.findText(expense['ExpCategory'])
                if index >= 0:
                    self.category_combo.setCurrentIndex(index)
                
                self.amount_edit.setText(str(expense['ExpAmount']))
                self.date_edit.setDate(QDate.fromString(expense['ExpDate'], "yyyy-MM-dd"))
                self.desc_edit.setText(expense['ExpDescription'] if expense['ExpDescription'] else "")
    
    def validate_and_save(self):
        """Validate input and save expense data"""
        category = self.category_combo.currentText()
        amount_str = self.amount_edit.text().strip()
        date = self.date_edit.date().toString("yyyy-MM-dd")
        desc = self.desc_edit.text().strip()
        
        # Validation
        if not all([category, amount_str, date]):
            QMessageBox.warning(self, "Error", "Category, Amount, and Date are required")
            return
        
        try:
            amount = float(amount_str)
            if amount <= 0:
                QMessageBox.warning(self, "Error", "Amount must be positive")
                return
        except ValueError:
            QMessageBox.warning(self, "Error", "Invalid amount")
            return
        
        # Save to database
        with get_db_connection() as (conn, c):
            if self.expense_id is None:
                # Insert new record
                c.execute('''
                    INSERT INTO Expenses (ExpCategory, ExpAmount, ExpDate, ExpDescription)
                    VALUES (?, ?, ?, ?)
                ''', (category, amount, date, desc if desc else None))
            else:
                # Update existing record
                c.execute('''
                    UPDATE Expenses 
                    SET ExpCategory=?, ExpAmount=?, ExpDate=?, ExpDescription=?
                    WHERE ExpID=?
                ''', (category, amount, date, desc if desc else None, self.expense_id))
            
            conn.commit()
        
        self.accept()

class ExpenseManagement(QWidget):
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
            QLineEdit, QDateEdit {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
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
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Header
        header_layout = QHBoxLayout()
        title = QLabel("Expense Management")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        layout.addLayout(header_layout)

        # First row: Filters (Search, From Date, To Date, Category ComboBox)
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by description...")
        self.search_input.setMaximumWidth(250)
        self.search_input.textChanged.connect(self.load_data)

        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDate(QDate.currentDate().addMonths(-1))
        self.from_date.dateChanged.connect(self.load_data)

        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDate(QDate.currentDate())
        self.to_date.dateChanged.connect(self.load_data)

        self.category_combo = QComboBox()
        self.category_combo.currentIndexChanged.connect(self.load_data)

        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(QLabel("From Date:"))
        filter_layout.addWidget(self.from_date)
        filter_layout.addWidget(QLabel("To Date:"))
        filter_layout.addWidget(self.to_date)
        filter_layout.addWidget(QLabel("Category:"))
        filter_layout.addWidget(self.category_combo)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        # Second row: Results, Total Amount, Print, Add New Expense buttons
        metrics_layout = QHBoxLayout()
        self.results_label = QLabel("0 results retrieved")
        self.results_label.setStyleSheet("font-size: 14px; color: #718096;")

        self.total_amount_label = QLabel("Total Amount: 0.00")
        self.total_amount_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #2D3748;")

        print_btn = QPushButton("Print")
        print_btn.clicked.connect(self.print_table)

        add_exp_btn = QPushButton("Add New Expense")
        add_exp_btn.clicked.connect(self.show_add_expense_dialog)

        metrics_layout.addWidget(self.results_label)
        metrics_layout.addWidget(self.total_amount_label)
        metrics_layout.addStretch()
        metrics_layout.addWidget(print_btn)
        metrics_layout.addWidget(add_exp_btn)
        layout.addLayout(metrics_layout)

        # Third row: Table
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["ID", "Category", "Amount", "Date", "Description", "Actions"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        header = self.table.horizontalHeader()
        for col in range(5):
            header.setSectionResizeMode(col, QHeaderView.Stretch)
        header.setSectionResizeMode(5, QHeaderView.Fixed)
        header.resizeSection(5, 150)  # Fixed width for actions
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # Initialize database and load data
        self.load_categories()
        self.load_data()

    def load_categories(self):
        """Load distinct categories from database for the filter dropdown"""
        with get_db_connection() as (conn, c):
            c.execute("SELECT DISTINCT ExpCategory FROM Expenses ORDER BY ExpCategory")
            categories = [row['ExpCategory'] for row in c.fetchall()]
        
        self.category_combo.clear()
        self.category_combo.addItem("All Categories")
        self.category_combo.addItems(categories)

    def load_data(self):
        """Load expense data from database with filtering"""
        search_text = self.search_input.text().strip()
        from_date = self.from_date.date().toString("yyyy-MM-dd")
        to_date = self.to_date.date().toString("yyyy-MM-dd")
        selected_category = self.category_combo.currentText()
        
        # Build query with filters
        query = """
            SELECT * FROM Expenses 
            WHERE ExpDate BETWEEN ? AND ?
        """
        params = [from_date, to_date]
        
        if search_text:
            query += " AND (ExpDescription LIKE ? OR ExpCategory LIKE ?)"
            params.extend([f"%{search_text}%", f"%{search_text}%"])
        
        if selected_category != "All Categories":
            query += " AND ExpCategory = ?"
            params.append(selected_category)
        
        query += " ORDER BY ExpDate DESC"
        
        with get_db_connection() as (conn, c):
            c.execute(query, params)
            data = c.fetchall()
        
        # Update table
        self.table.setRowCount(len(data))
        total_amount = 0.0
        
        for i, row in enumerate(data):
            self.table.setItem(i, 0, QTableWidgetItem(str(row['ExpID'])))
            self.table.setItem(i, 1, QTableWidgetItem(row['ExpCategory']))
            self.table.setItem(i, 2, QTableWidgetItem(f"{row['ExpAmount']:.2f}"))
            self.table.setItem(i, 3, QTableWidgetItem(row['ExpDate']))
            self.table.setItem(i, 4, QTableWidgetItem(row['ExpDescription'] if row['ExpDescription'] else ""))
            
            # Calculate total amount for all expenses
            total_amount += row['ExpAmount']
            
            # Add action buttons
            self.add_action_buttons(i, row['ExpID'])
        
        # Update metrics
        self.total_amount_label.setText(f"Total Amount: {total_amount:.2f}")
        self.results_label.setText(f"{len(data)} results retrieved")

    def add_action_buttons(self, row, exp_id):
        actions_widget = QWidget()
        actions_widget.setMinimumWidth(140)
        actions_layout = QHBoxLayout(actions_widget)
        actions_layout.setContentsMargins(2, 2, 2, 2)
        actions_layout.setSpacing(2)

        edit_btn = QPushButton("Edit")
        edit_btn.setFixedWidth(60)
        edit_btn.clicked.connect(lambda: self.edit_expense(exp_id))
        delete_btn = QPushButton("Delete")
        delete_btn.setFixedWidth(60)
        delete_btn.clicked.connect(lambda: self.delete_expense(exp_id))

        actions_layout.addWidget(edit_btn)
        actions_layout.addWidget(delete_btn)
        actions_layout.addStretch()

        self.table.setCellWidget(row, 5, actions_widget)

    def print_table(self):
        """Print functionality placeholder"""
        QMessageBox.information(self, "Print", "Printing expense table... (placeholder)")

    def show_add_expense_dialog(self):
        """Show dialog to add new expense"""
        dialog = ExpenseDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            self.load_categories()  # Refresh categories in case a new one was added
            self.load_data()

    def edit_expense(self, exp_id):
        """Edit existing expense"""
        dialog = ExpenseDialog(self, exp_id)
        if dialog.exec_() == QDialog.Accepted:
            self.load_categories()  # Refresh categories in case a new one was added
            self.load_data()

    def delete_expense(self, exp_id):
        """Delete expense after confirmation"""
        reply = QMessageBox.question(
            self, 
            'Confirm Delete', 
            f'Are you sure you want to delete expense #{exp_id}?', 
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            with get_db_connection() as (conn, c):
                c.execute("DELETE FROM Expenses WHERE ExpID = ?", (exp_id,))
                conn.commit()
            
            QMessageBox.information(self, "Success", "Expense deleted successfully")
            self.load_categories()  # Refresh categories in case we deleted the last item of a category
            self.load_data()

#### This is End of the Uniforms  ####
#### start of expenditure #####

class Expenditure(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Expenditure Management System")
        self.setGeometry(100, 100, 1000, 700)
       
        
        layout = QVBoxLayout(self)
        
        # Main tabs for different sections
        self.main_tabs = QTabWidget()
        
        # Categories Tab
        self.expenditure_tab = ExpenditureManagement()
        self.main_tabs.addTab(self.expenditure_tab, "Expenditure")
        
        # Placeholder for future tabs
        self.assets_tab = AssetsManagement()
        self.main_tabs.addTab(self.assets_tab, "Assets")
        
        self.categories_tab = CategoriesTab()
        self.main_tabs.addTab(self.categories_tab, "Categories")
        
        layout.addWidget(self.main_tabs)

class CategoryDialog(QDialog):
    def __init__(self, parent=None, category_data=None, is_subcategory=False, categories=None):
        super().__init__(parent)
        self.category_data = category_data
        self.is_subcategory = is_subcategory
        self.categories = categories or []
        self.setup_ui()
        
        
    def setup_ui(self):
        if self.is_subcategory:
            self.setWindowTitle("Edit Subcategory" if self.category_data else "Add Subcategory")
        else:
            self.setWindowTitle("Edit Category" if self.category_data else "Add Category")
            
        self.setModal(True)
        self.resize(400, 200)
        
        layout = QFormLayout(self)
        
        if self.is_subcategory:
            self.category_combo = QComboBox()
            for cat_id, cat_name in self.categories:
                self.category_combo.addItem(cat_name, cat_id)
            layout.addRow("Category:", self.category_combo)
            
        self.name_edit = QLineEdit()
        layout.addRow("Name:" if self.is_subcategory else "Category Name:", self.name_edit)
        
        if not self.is_subcategory:
            self.type_combo = QComboBox()
            self.type_combo.addItems(["Expense", "Income", "Transfer"])
            layout.addRow("Type:", self.type_combo)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)
        
        if self.category_data:
            self.populate_data()
    
    def populate_data(self):
        if self.is_subcategory:
            self.name_edit.setText(self.category_data['ExpenditureSubCategoryName'])
            index = self.category_combo.findData(self.category_data['ExpenditureCategoryID'])
            if index >= 0:
                self.category_combo.setCurrentIndex(index)
        else:
            self.name_edit.setText(self.category_data['ExpCatName'])
            index = self.type_combo.findText(self.category_data['ExpCatType'])
            if index >= 0:
                self.type_combo.setCurrentIndex(index)
    
    def validate_and_accept(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Validation Error", "Name is required!")
            return
        
        self.accept()
    
    def get_data(self):
        if self.is_subcategory:
            return {
                'name': self.name_edit.text().strip(),
                'category_id': self.category_combo.currentData()
            }
        else:
            return {
                'name': self.name_edit.text().strip(),
                'type': self.type_combo.currentText()
            }

class CategoriesTab(QWidget):
    def __init__(self):
        super().__init__()
        self.setup_ui()
        self.load_categories()
        self.load_subcategories()
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
            QLineEdit, QDateEdit, QTextEdit {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
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
        
        
    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        
        # Title
        title = QLabel("Categories Management")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignLeft)
        main_layout.addWidget(title)
        
        # Filters row
        filters_layout = QHBoxLayout()
        
        # Left side filters
        left_filters = QHBoxLayout()
        
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Search...")
        self.search_edit.textChanged.connect(self.filter_table)
        left_filters.addWidget(QLabel("Search:"))
        left_filters.addWidget(self.search_edit)
        
        self.type_combo = QComboBox()
        self.type_combo.addItem("All Types", "")
        self.type_combo.addItems(["Expense", "Income", "Transfer"])
        self.type_combo.currentTextChanged.connect(self.filter_table)
        left_filters.addWidget(QLabel("Type:"))
        left_filters.addWidget(self.type_combo)
        
        self.category_combo = QComboBox()
        self.category_combo.addItem("All Categories", "")
        self.category_combo.currentTextChanged.connect(self.filter_table)
        left_filters.addWidget(QLabel("Category:"))
        left_filters.addWidget(self.category_combo)
        
        filters_layout.addLayout(left_filters)
        filters_layout.addStretch()
        
        # Right side buttons
        right_buttons = QHBoxLayout()
        
        self.new_category_btn = QPushButton("New Category")
        self.new_category_btn.clicked.connect(self.add_category)
        right_buttons.addWidget(self.new_category_btn)
        
        self.new_subcat_btn = QPushButton("New Subcategory")
        self.new_subcat_btn.clicked.connect(self.add_subcategory)
        right_buttons.addWidget(self.new_subcat_btn)
        
        self.print_btn = QPushButton("Print")
        self.print_btn.clicked.connect(self.print_data)
        right_buttons.addWidget(self.print_btn)
        
        filters_layout.addLayout(right_buttons)
        main_layout.addLayout(filters_layout)
        
        # Create tabs for Categories and Subcategories
        self.tabs = QTabWidget()
        
        # Categories Tab
        self.categories_tab = QWidget()
        self.setup_categories_tab()
        self.tabs.addTab(self.categories_tab, "Categories")
        
        # Subcategories Tab
        self.subcategories_tab = QWidget()
        self.setup_subcategories_tab()
        self.tabs.addTab(self.subcategories_tab, "Subcategories")
        
        main_layout.addWidget(self.tabs)
    
    def setup_categories_tab(self):
        layout = QVBoxLayout(self.categories_tab)
        
        # Categories table
        self.categories_table = QTableWidget()
        self.categories_table.setColumnCount(4)
        self.categories_table.setHorizontalHeaderLabels(["ID", "Name", "Type", "Actions"])
        self.categories_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.categories_table)
    
    def setup_subcategories_tab(self):
        layout = QVBoxLayout(self.subcategories_tab)
        
        # Subcategories table
        self.subcategories_table = QTableWidget()
        self.subcategories_table.setColumnCount(4)
        self.subcategories_table.setHorizontalHeaderLabels(["ID", "Name", "Category", "Type", "Actions"])
        self.subcategories_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.subcategories_table)
    
    def get_db_connection(self):
        """Context manager for database connections"""
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
    
    def load_categories(self):
        try:
            conn_gen = self.get_db_connection()
            conn, c = next(conn_gen)
            c.execute("SELECT ExpCatID, ExpCatName, ExpCatType FROM ExpenditureCategory")
            self.categories = [dict(row) for row in c.fetchall()]
            
            # Update category filter combo
            self.category_combo.clear()
            self.category_combo.addItem("All Categories", "")
            for cat in self.categories:
                self.category_combo.addItem(cat['ExpCatName'], cat['ExpCatID'])
                
            self.populate_categories_table()
            
        except Exception as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load categories: {str(e)}")
    
    def load_subcategories(self):
        try:
            conn_gen = self.get_db_connection()
            conn, c = next(conn_gen)
            query = """
                SELECT esc.ExpeSubCatID, esc.ExpenditureSubCategoryName, 
                       ec.ExpCatID, ec.ExpCatName, ec.ExpCatType
                FROM ExpenditureSubCategory esc
                JOIN ExpenditureCategory ec ON esc.ExpenditureCategoryID = ec.ExpCatID
            """
            c.execute(query)
            self.subcategories = [dict(row) for row in c.fetchall()]
            self.populate_subcategories_table()
            
        except Exception as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load subcategories: {str(e)}")
    
    def populate_categories_table(self):
        self.categories_table.setRowCount(0)
        
        for category in self.categories:
            row_position = self.categories_table.rowCount()
            self.categories_table.insertRow(row_position)
            
            self.categories_table.setItem(row_position, 0, QTableWidgetItem(str(category['ExpCatID'])))
            self.categories_table.setItem(row_position, 1, QTableWidgetItem(category['ExpCatName']))
            self.categories_table.setItem(row_position, 2, QTableWidgetItem(category['ExpCatType']))
            
            # Actions column
            actions_widget = QWidget()
            actions_layout = QHBoxLayout(actions_widget)
            actions_layout.setContentsMargins(5, 2, 5, 2)
            
            edit_btn = QPushButton("Edit")
            edit_btn.clicked.connect(lambda checked, data=category: self.edit_category(data))
            edit_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; border: none; padding: 5px; }")
            
            delete_btn = QPushButton("Delete")
            delete_btn.clicked.connect(lambda checked, data=category: self.delete_category(data))
            delete_btn.setStyleSheet("QPushButton { background-color: #f44336; color: white; border: none; padding: 5px; }")
            
            actions_layout.addWidget(edit_btn)
            actions_layout.addWidget(delete_btn)
            actions_layout.addStretch()
            
            self.categories_table.setCellWidget(row_position, 3, actions_widget)
    
    def populate_subcategories_table(self):
        self.subcategories_table.setRowCount(0)
        
        for subcat in self.subcategories:
            row_position = self.subcategories_table.rowCount()
            self.subcategories_table.insertRow(row_position)
            
            self.subcategories_table.setItem(row_position, 0, QTableWidgetItem(str(subcat['ExpeSubCatID'])))
            self.subcategories_table.setItem(row_position, 1, QTableWidgetItem(subcat['ExpenditureSubCategoryName']))
            self.subcategories_table.setItem(row_position, 2, QTableWidgetItem(subcat['ExpCatName']))
            self.subcategories_table.setItem(row_position, 3, QTableWidgetItem(subcat['ExpCatType']))
            
            # Actions column
            actions_widget = QWidget()
            actions_layout = QHBoxLayout(actions_widget)
            actions_layout.setContentsMargins(5, 2, 5, 2)
            
            edit_btn = QPushButton("Edit")
            edit_btn.clicked.connect(lambda checked, data=subcat: self.edit_subcategory(data))
            edit_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; border: none; padding: 5px; }")
            
            delete_btn = QPushButton("Delete")
            delete_btn.clicked.connect(lambda checked, data=subcat: self.delete_subcategory(data))
            delete_btn.setStyleSheet("QPushButton { background-color: #f44336; color: white; border: none; padding: 5px; }")
            
            actions_layout.addWidget(edit_btn)
            actions_layout.addWidget(delete_btn)
            actions_layout.addStretch()
            
            self.subcategories_table.setCellWidget(row_position, 4, actions_widget)
    
    def filter_table(self):
        # This now filters the subcategories table
        search_text = self.search_edit.text().lower()
        type_filter = self.type_combo.currentText()
        category_filter = self.category_combo.currentText()
        
        current_tab = self.tabs.currentIndex()
        
        if current_tab == 0:  # Categories tab
            table = self.categories_table
            for row in range(table.rowCount()):
                should_show = True
                
                if search_text:
                    name_item = table.item(row, 1)
                    if name_item and search_text not in name_item.text().lower():
                        should_show = False
                
                if should_show and type_filter and type_filter != "All Types":
                    type_item = table.item(row, 2)
                    if type_item and type_item.text() != type_filter:
                        should_show = False
                
                table.setRowHidden(row, not should_show)
                
        else:  # Subcategories tab
            table = self.subcategories_table
            for row in range(table.rowCount()):
                should_show = True
                
                if search_text:
                    name_item = table.item(row, 1)
                    if name_item and search_text not in name_item.text().lower():
                        should_show = False
                
                if should_show and type_filter and type_filter != "All Types":
                    type_item = table.item(row, 3)
                    if type_item and type_item.text() != type_filter:
                        should_show = False
                
                if should_show and category_filter and category_filter != "All Categories":
                    category_item = table.item(row, 2)
                    if category_item and category_item.text() != category_filter:
                        should_show = False
                
                table.setRowHidden(row, not should_show)
    
    def add_category(self):
        dialog = CategoryDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            try:
                conn_gen = self.get_db_connection()
                conn, c = next(conn_gen)
                c.execute("INSERT INTO ExpenditureCategory (ExpCatName, ExpCatType) VALUES (?, ?)",
                         (data['name'], data['type']))
                conn.commit()
                self.load_categories()
                QMessageBox.information(self, "Success", "Category added successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Database Error", f"Failed to add category: {str(e)}")
    
    def edit_category(self, data):
        dialog = CategoryDialog(self, category_data=data)
        if dialog.exec_() == QDialog.Accepted:
            new_data = dialog.get_data()
            try:
                conn_gen = self.get_db_connection()
                conn, c = next(conn_gen)
                c.execute("UPDATE ExpenditureCategory SET ExpCatName = ?, ExpCatType = ? WHERE ExpCatID = ?",
                         (new_data['name'], new_data['type'], data['ExpCatID']))
                conn.commit()
                self.load_categories()
                self.load_subcategories()  # Reload subcategories in case category type changed
                QMessageBox.information(self, "Success", "Category updated successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Database Error", f"Failed to update category: {str(e)}")
    
    def delete_category(self, data):
        reply = QMessageBox.question(self, "Confirm Delete", 
                                   f"Are you sure you want to delete category '{data['ExpCatName']}'?",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                # Check if category has subcategories
                conn_gen = self.get_db_connection()
                conn, c = next(conn_gen)
                c.execute("SELECT COUNT(*) FROM ExpenditureSubCategory WHERE ExpenditureCategoryID = ?", (data['ExpCatID'],))
                subcat_count = c.fetchone()[0]
                
                if subcat_count > 0:
                    QMessageBox.warning(self, "Cannot Delete", 
                                      "This category has subcategories. Please delete all subcategories first.")
                    return
                
                c.execute("DELETE FROM ExpenditureCategory WHERE ExpCatID = ?", (data['ExpCatID'],))
                conn.commit()
                self.load_categories()
                QMessageBox.information(self, "Success", "Category deleted successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Database Error", f"Failed to delete category: {str(e)}")
    
    def add_subcategory(self):
        if not self.categories:
            QMessageBox.warning(self, "No Categories", "Please create a category first!")
            return
            
        dialog = CategoryDialog(self, is_subcategory=True, categories=[(cat['ExpCatID'], cat['ExpCatName']) for cat in self.categories])
        if dialog.exec_() == QDialog.Accepted:
            data = dialog.get_data()
            try:
                conn_gen = self.get_db_connection()
                conn, c = next(conn_gen)
                c.execute("INSERT INTO ExpenditureSubCategory (ExpenditureCategoryID, ExpenditureSubCategoryName) VALUES (?, ?)",
                         (data['category_id'], data['name']))
                conn.commit()
                self.load_subcategories()
                QMessageBox.information(self, "Success", "Subcategory added successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Database Error", f"Failed to add subcategory: {str(e)}")
    
    def edit_subcategory(self, data):
        dialog = CategoryDialog(self, category_data=data, is_subcategory=True, 
                               categories=[(cat['ExpCatID'], cat['ExpCatName']) for cat in self.categories])
        if dialog.exec_() == QDialog.Accepted:
            new_data = dialog.get_data()
            try:
                conn_gen = self.get_db_connection()
                conn, c = next(conn_gen)
                c.execute("UPDATE ExpenditureSubCategory SET ExpenditureSubCategoryName = ?, ExpenditureCategoryID = ? WHERE ExpeSubCatID = ?",
                         (new_data['name'], new_data['category_id'], data['ExpeSubCatID']))
                conn.commit()
                self.load_subcategories()
                QMessageBox.information(self, "Success", "Subcategory updated successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Database Error", f"Failed to update subcategory: {str(e)}")
    
    def delete_subcategory(self, data):
        reply = QMessageBox.question(self, "Confirm Delete", 
                                   f"Are you sure you want to delete '{data['ExpenditureSubCategoryName']}'?",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            try:
                # Check if subcategory is used in expenditures
                conn_gen = self.get_db_connection()
                conn, c = next(conn_gen)
                c.execute("SELECT COUNT(*) FROM Expenditure WHERE ExpeSubID = ?", (data['ExpeSubCatID'],))
                count = c.fetchone()[0]
                
                if count > 0:
                    QMessageBox.warning(self, "Cannot Delete", 
                                      "This subcategory is used in existing expenditures and cannot be deleted.")
                    return
                
                c.execute("DELETE FROM ExpenditureSubCategory WHERE ExpeSubCatID = ?", (data['ExpeSubCatID'],))
                conn.commit()
                self.load_subcategories()
                QMessageBox.information(self, "Success", "Subcategory deleted successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Database Error", f"Failed to delete subcategory: {str(e)}")
    
    def print_data(self):
        current_tab = self.tabs.currentIndex()
        if current_tab == 0:
            QMessageBox.information(self, "Print", "Printing categories data...")
        else:
            QMessageBox.information(self, "Print", "Printing subcategories data...")

class ExpenditureManagement(QWidget):
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
            QLineEdit, QDateEdit, QTextEdit {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
            }
            QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
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
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Header (unchanged)
        header_layout = QHBoxLayout()
        title = QLabel("Expenditure Management")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        add_btn = QPushButton("Add New Expenditure")
        add_btn.clicked.connect(self.show_add_dialog)
        header_layout.addWidget(add_btn)
        layout.addLayout(header_layout)

        # FILTERS - REMOVED TYPE FILTER FROM MAIN PAGE
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search...")
        self.search_input.setMaximumWidth(200)
        self.search_input.textChanged.connect(self.load_data)

        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDate(QDate.currentDate().addMonths(-1))
        self.from_date.dateChanged.connect(self.load_data)

        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDate(QDate.currentDate())
        self.to_date.dateChanged.connect(self.load_data)

        # Category and Subcategory filters (for ALL types on main page)
        self.category_combo = QComboBox()
        self.category_combo.addItem("All Categories")
        self.load_categories()  # Load ALL categories for filtering
        self.category_combo.currentIndexChanged.connect(self.on_category_changed)

        self.subcategory_combo = QComboBox()
        self.subcategory_combo.addItem("All Subcategories")
        self.subcategory_combo.currentIndexChanged.connect(self.load_data)
    
        self.payment_combo = QComboBox()
        self.payment_combo.addItem("All Payment Methods")
        self.payment_combo.addItems(["Cash", "Bank Transfer", "Credit Card", "Other"])
        self.payment_combo.currentIndexChanged.connect(self.load_data)

        filter_layout.addWidget(QLabel("Search:"))
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(QLabel("From:"))
        filter_layout.addWidget(self.from_date)
        filter_layout.addWidget(QLabel("To:"))
        filter_layout.addWidget(self.to_date)
        filter_layout.addWidget(QLabel("Category:"))
        filter_layout.addWidget(self.category_combo)
        filter_layout.addWidget(QLabel("Subcategory:"))
        filter_layout.addWidget(self.subcategory_combo)
        filter_layout.addWidget(QLabel("Payment:"))
        filter_layout.addWidget(self.payment_combo)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        # Metrics and Table (unchanged)
        metrics_layout = QHBoxLayout()
        self.results_label = QLabel("0 results")
        self.results_label.setStyleSheet("font-size: 14px; color: #718096;")
        self.total_label = QLabel("Total Amount: 0.00")
        self.total_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        metrics_layout.addWidget(self.results_label)
        metrics_layout.addWidget(self.total_label)
        metrics_layout.addStretch()
        layout.addLayout(metrics_layout)

        # Table - SIMPLIFIED for Expenditure data only
        self.table = QTableWidget()
        self.table.setColumnCount(8)  # Removed Asset Name column
        self.table.setHorizontalHeaderLabels(["ID", "Category", "Subcategory", "Amount", "Date", "Description", "Payment Method", "Actions"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        header = self.table.horizontalHeader()
        for col in range(6):  # Adjusted columns
            header.setSectionResizeMode(col, QHeaderView.Stretch)
        header.setSectionResizeMode(6, QHeaderView.Fixed)  # Actions column
        header.resizeSection(6, 150)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # Load initial data
        self.load_data()

    def on_category_changed(self):
        """Handle category selection change - for FILTERING on main page"""
        if self.category_combo.currentIndex() == 0:  # "All Categories"
            cat_id = None
        else:
            cat_id = self.category_combo.currentData()
        
        self.load_subcategories(cat_id)
        self.load_data()

    def load_categories(self):
        """Load ALL categories for FILTERING on main page"""
        self.category_combo.clear()
        self.category_combo.addItem("All Categories")
        
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT ExpCatID, ExpCatName FROM ExpenditureCategory")
                for row in c.fetchall():
                    self.category_combo.addItem(row[1], row[0])
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load categories: {str(e)}")

    def load_subcategories(self, cat_id=None):
        """Load subcategories for FILTERING on main page"""
        self.subcategory_combo.clear()
        self.subcategory_combo.addItem("All Subcategories")
        
        if cat_id is None:
            return
        
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT ExpeSubCatID, ExpenditureSubCategoryName FROM ExpenditureSubCategory WHERE ExpenditureCategoryID = ?", (cat_id,))
                for row in c.fetchall():
                    self.subcategory_combo.addItem(row[1], row[0])
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load subcategories: {str(e)}")
   
    def load_data(self):
        """Load expenditure data with filters - SIMPLIFIED without type filtering"""
        search_text = self.search_input.text().strip()
        from_date_str = self.from_date.date().toString("yyyy-MM-dd")
        to_date_str = self.to_date.date().toString("yyyy-MM-dd")
        
        # Get filters safely
        category_id = None
        if self.category_combo.currentIndex() > 0:
            category_id = self.category_combo.currentData()
        
        subcategory_id = None
        if self.subcategory_combo.currentIndex() > 0:
            subcategory_id = self.subcategory_combo.currentData()
        
        payment_method = self.payment_combo.currentText()
        if payment_method == "All Payment Methods":
            payment_method = None
        
        try:
            with get_db_connection() as (conn, c):
                query = """
                    SELECT e.ExpenseID, ec.ExpCatName, esc.ExpenditureSubCategoryName, e.ExpenseAmount,
                           e.ExpenseDate, e.ExpenseDesc, e.PaymentMethod
                    FROM Expenditure e
                    JOIN ExpenditureSubCategory esc ON e.ExpeSubID = esc.ExpeSubCatID
                    JOIN ExpenditureCategory ec ON esc.ExpenditureCategoryID = ec.ExpCatID
                    WHERE 1=1
                """
                params = []
                
                if search_text:
                    query += " AND (ec.ExpCatName LIKE ? OR esc.ExpenditureSubCategoryName LIKE ? OR e.ExpenseDesc LIKE ?)"
                    params.extend([f"%{search_text}%", f"%{search_text}%", f"%{search_text}%"])
                
                if from_date_str:
                    query += " AND e.ExpenseDate >= ?"
                    params.append(from_date_str)
                
                if to_date_str:
                    query += " AND e.ExpenseDate <= ?"
                    params.append(to_date_str)
                
                if category_id:
                    query += " AND ec.ExpCatID = ?"
                    params.append(category_id)
                
                if subcategory_id:
                    query += " AND esc.ExpeSubCatID = ?"
                    params.append(subcategory_id)
                
                if payment_method:
                    query += " AND e.PaymentMethod = ?"
                    params.append(payment_method)
                
                query += " ORDER BY e.ExpenseDate DESC"
                
                c.execute(query, params)
                rows = c.fetchall()
                
                self.table.setRowCount(len(rows))
                total_amount = 0
                for i, row in enumerate(rows):
                    for j, val in enumerate(row):
                        self.table.setItem(i, j, QTableWidgetItem(str(val) if val else ""))
                    total_amount += row[3] or 0
                    self.add_action_buttons(i, row[0])
                
                self.results_label.setText(f"{len(rows)} results")
                self.total_label.setText(f"Total Amount: {total_amount:.2f}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")

    def add_action_buttons(self, row, exp_id):
        """Add action buttons to table row"""
        actions_widget = QWidget()
        actions_layout = QHBoxLayout(actions_widget)
        actions_layout.setContentsMargins(5, 2, 5, 2)
        actions_layout.setSpacing(5)

        edit_btn = QPushButton("Edit")
        edit_btn.setFixedWidth(60)
        edit_btn.setStyleSheet("background-color: #4CAF50; color: white;")
        edit_btn.clicked.connect(lambda: self.edit_expenditure(exp_id))
        
        delete_btn = QPushButton("Delete")
        delete_btn.setFixedWidth(60)
        delete_btn.setStyleSheet("background-color: #f44336; color: white;")
        delete_btn.clicked.connect(lambda: self.delete_expenditure(exp_id))

        actions_layout.addWidget(edit_btn)
        actions_layout.addWidget(delete_btn)
        actions_layout.addStretch()

        self.table.setCellWidget(row, 7, actions_widget)  # Updated column index

    def show_add_dialog(self):
        """Show dialog to add new expenditure - WITH TYPE FILTERING FOR CATEGORIES"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Add New Expenditure")
        dialog.setMinimumWidth(500)
        layout = QVBoxLayout(dialog)
        
        form_layout = QFormLayout()
        
        # TYPE SELECTION - Filters categories by ExpCatType
        type_combo = QComboBox()
        type_combo.addItems(["Expenses", "Assets"])
        
        category_combo = QComboBox()
        category_combo.addItem("Select Category")
        
        subcategory_combo = QComboBox()
        subcategory_combo.addItem("Select Subcategory")
        
        # When type changes, filter categories by type
        type_combo.currentTextChanged.connect(
            lambda text: self.load_categories_by_type(
                category_combo, 
                "Expense" if text == "Expenses" else "Asset"
            )
        )
        
        # When category changes, load its subcategories
        category_combo.currentIndexChanged.connect(
            lambda: self.load_subcategories_dialog(
                subcategory_combo, 
                category_combo.currentData() if category_combo.currentIndex() > 0 else None
            )
        )
        
        # Load initial categories for default type
        self.load_categories_by_type(category_combo, "Expense")
        
        amount_edit = QLineEdit()
        amount_edit.setValidator(QDoubleValidator(0.01, 9999999.99, 2))
        
        date_edit = QDateEdit()
        date_edit.setCalendarPopup(True)
        date_edit.setDate(QDate.currentDate())
        
        desc_edit = QTextEdit()
        desc_edit.setMaximumHeight(80)
        desc_edit.setPlaceholderText("Enter description...")
        
        payment_combo = QComboBox()
        payment_combo.addItems(["Cash", "Bank Transfer", "Credit Card", "Other"])
        
        form_layout.addRow("Type:", type_combo)
        form_layout.addRow("Category:", category_combo)
        form_layout.addRow("Subcategory:", subcategory_combo)
        form_layout.addRow("Amount:", amount_edit)
        form_layout.addRow("Date:", date_edit)
        form_layout.addRow("Description:", desc_edit)
        form_layout.addRow("Payment Method:", payment_combo)
        
        layout.addLayout(form_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_expenditure(
            dialog, subcategory_combo, amount_edit, date_edit, desc_edit, payment_combo
        ))
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        
        dialog.exec_()

    def load_categories_by_type(self, combo, category_type):
        """Load categories filtered by type for DIALOG"""
        try:
            combo.clear()
            combo.addItem("Select Category")
            with get_db_connection() as (conn, c):
                c.execute("SELECT ExpCatID, ExpCatName FROM ExpenditureCategory WHERE ExpCatType = ?", (category_type,))
                for row in c.fetchall():
                    combo.addItem(row[1], row[0])
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load categories: {str(e)}")

    def load_subcategories_dialog(self, combo, cat_id):
        """Load subcategories for DIALOG"""
        combo.clear()
        combo.addItem("Select Subcategory")
        
        if not cat_id:
            return
            
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT ExpeSubCatID, ExpenditureSubCategoryName FROM ExpenditureSubCategory WHERE ExpenditureCategoryID = ?", (cat_id,))
                for row in c.fetchall():
                    combo.addItem(row[1], row[0])
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load subcategories: {str(e)}")

    def save_expenditure(self, dialog, subcategory_combo, amount_edit, date_edit, desc_edit, payment_combo):
        """Save new expenditure to database"""
        # Validate inputs
        if subcategory_combo.currentIndex() == 0:
            QMessageBox.warning(dialog, "Error", "Please select a subcategory")
            return
            
        sub_id = subcategory_combo.currentData()
        amount_str = amount_edit.text().strip()
        date_str = date_edit.date().toString("yyyy-MM-dd")
        desc = desc_edit.toPlainText().strip()
        payment = payment_combo.currentText()
        
        if not amount_str:
            QMessageBox.warning(dialog, "Error", "Please enter an amount")
            return
            
        try:
            amount = float(amount_str)
            if amount <= 0:
                QMessageBox.warning(dialog, "Error", "Amount must be positive")
                return
        except ValueError:
            QMessageBox.warning(dialog, "Error", "Please enter a valid amount")
            return
        
        # Save to database
        try:
            with get_db_connection() as (conn, c):
                c.execute("""INSERT INTO Expenditure (ExpeSubID, ExpenseAmount, ExpenseDate, ExpenseDesc, PaymentMethod)
                             VALUES (?, ?, ?, ?, ?)""", 
                         (sub_id, amount, date_str, desc, payment))
                conn.commit()
                
                QMessageBox.information(dialog, "Success", "Expenditure added successfully")
                dialog.accept()
                self.load_data()
                
        except Exception as e:
            QMessageBox.warning(dialog, "Error", f"Failed to save expenditure: {str(e)}")

    def edit_expenditure(self, exp_id):
        try:
            with get_db_connection() as (conn, c):
                c.execute("""
                    SELECT e.ExpenseID, esc.ExpenditureCategoryID, e.ExpeSubID, e.ExpenseAmount, 
                        e.ExpenseDate, e.ExpenseDesc, e.PaymentMethod, ec.ExpCatType
                    FROM Expenditure e
                    JOIN ExpenditureSubCategory esc ON e.ExpeSubID = esc.ExpeSubCatID
                    JOIN ExpenditureCategory ec ON esc.ExpenditureCategoryID = ec.ExpCatID
                    WHERE e.ExpenseID = ?
                """, (exp_id,))
                data = c.fetchone()
                
            if not data:
                QMessageBox.warning(self, "Error", "Expenditure not found")
                return
                
            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Expenditure")
            dialog.setMinimumWidth(500)
            layout = QVBoxLayout(dialog)
            
            form_layout = QFormLayout()
            
            # Type selection (based on category type from database)
            type_combo = QComboBox()
            type_combo.addItems(["Expenses", "Assets"])
            
            # Set the type based on category type from database
            category_type = "Expenses" if data[7] == "Expense" else "Assets"
            type_combo.setCurrentText(category_type)
            type_combo.setEnabled(False)  # Cannot change type when editing
            
            category_combo = QComboBox()
            category_combo.addItem("Select Category")
            
            subcategory_combo = QComboBox()
            subcategory_combo.addItem("Select Subcategory")
            
            # Load categories for the specific type
            self.load_categories_by_type(category_combo, data[7])
            
            # Set current category and load its subcategories
            category_combo.setCurrentIndex(category_combo.findData(data[1]))
            self.load_subcategories_dialog(subcategory_combo, data[1])
            subcategory_combo.setCurrentIndex(subcategory_combo.findData(data[2]))
        
            # Connect signals
            type_combo.currentTextChanged.connect(
                lambda text: self.load_categories_by_type(
                    category_combo, 
                    "Expense" if text == "Expenses" else "Asset"
                )
            )
            
            category_combo.currentIndexChanged.connect(
                lambda: self.load_subcategories_dialog(
                    subcategory_combo, 
                    category_combo.currentData() if category_combo.currentIndex() > 0 else None
                )
            )
            
            amount_edit = QLineEdit(str(data[3]))
            amount_edit.setValidator(QDoubleValidator(0.01, 9999999.99, 2))
            
            date_edit = QDateEdit(QDate.fromString(data[4], "yyyy-MM-dd"))
            date_edit.setCalendarPopup(True)
            date_edit.setEnabled(False)  # Date cannot be changed as per requirements
            
            desc_edit = QTextEdit(data[5] or "")
            desc_edit.setMaximumHeight(80)
            
            payment_combo = QComboBox()
            payment_combo.addItems(["Cash", "Bank Transfer", "Credit Card", "Other"])
            payment_combo.setCurrentText(data[6] or "Cash")
            
            form_layout.addRow("Type:", type_combo)
            form_layout.addRow("Category:", category_combo)
            form_layout.addRow("Subcategory:", subcategory_combo)
            form_layout.addRow("Amount:", amount_edit)
            form_layout.addRow("Date:", date_edit)
            form_layout.addRow("Description:", desc_edit)
            form_layout.addRow("Payment Method:", payment_combo)
            
            layout.addLayout(form_layout)
            
            # Buttons
            button_layout = QHBoxLayout()
            save_btn = QPushButton("Save")
            save_btn.clicked.connect(lambda: self.save_edited_expenditure(
                dialog, exp_id, subcategory_combo, amount_edit, desc_edit, payment_combo
            ))
            cancel_btn = QPushButton("Cancel")
            cancel_btn.clicked.connect(dialog.reject)
            
            button_layout.addStretch()
            button_layout.addWidget(save_btn)
            button_layout.addWidget(cancel_btn)
            layout.addLayout(button_layout)
            
            dialog.exec_()
        
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load expenditure: {str(e)}")

    def save_edited_expenditure(self, dialog, exp_id, subcategory_combo, amount_edit, desc_edit, payment_combo):
        """Save edited expenditure"""
        # Validate inputs
        if subcategory_combo.currentIndex() == 0:
            QMessageBox.warning(dialog, "Error", "Please select a subcategory")
            return
            
        sub_id = subcategory_combo.currentData()
        amount_str = amount_edit.text().strip()
        desc = desc_edit.toPlainText().strip()
        payment = payment_combo.currentText()
        
        if not amount_str:
            QMessageBox.warning(dialog, "Error", "Please enter an amount")
            return
            
        try:
            amount = float(amount_str)
            if amount <= 0:
                QMessageBox.warning(dialog, "Error", "Amount must be positive")
                return
        except ValueError:
            QMessageBox.warning(dialog, "Error", "Please enter a valid amount")
            return
        
        # Save to database
        try:
            with get_db_connection() as (conn, c):
                c.execute("""UPDATE Expenditure SET ExpeSubID=?, ExpenseAmount=?, ExpenseDesc=?, PaymentMethod=?
                             WHERE ExpenseID=?""", 
                         (sub_id, amount, desc, payment, exp_id))
                conn.commit()
                
                QMessageBox.information(dialog, "Success", "Expenditure updated successfully")
                dialog.accept()
                self.load_data()
                
        except Exception as e:
            QMessageBox.warning(dialog, "Error", f"Failed to update expenditure: {str(e)}")

    def delete_expenditure(self, exp_id):
        """Delete expenditure/asset with confirmation"""
        reply = QMessageBox.question(
            self, 'Confirm Delete', 
            f'Are you sure you want to delete record #{exp_id}?', 
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    # First delete asset if exists, then expenditure
                    c.execute("DELETE FROM Assets WHERE ExpenseID = ?", (exp_id,))
                    c.execute("DELETE FROM Expenditure WHERE ExpenseID = ?", (exp_id,))
                    conn.commit()
                    
                    QMessageBox.information(self, "Success", "Record deleted successfully")
                    self.load_data()
                    
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to delete record: {str(e)}")

class AssetsManagement(QWidget):
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
            QLineEdit, QDateEdit, QTextEdit, QComboBox {
                background-color: #FFFFFF;
                border: 1px solid #E0E3E7;
                padding: 8px;
                border-radius: 5px;
                font-size: 14px;
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
            QGroupBox {
                font-weight: bold;
                margin-top: 10px;
                padding-top: 10px;
                border: 1px solid #E0E3E7;
                border-radius: 5px;
            }
        """)
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # Header
        header_layout = QHBoxLayout()
        title = QLabel("Assets Management")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        
        add_btn = QPushButton("Add New Asset")
        add_btn.clicked.connect(self.show_add_dialog)
        header_layout.addWidget(add_btn)
        
        refresh_btn = QPushButton("Refresh")
        refresh_btn.clicked.connect(self.load_data)
        header_layout.addWidget(refresh_btn)
        
        layout.addLayout(header_layout)

        # Filters
        filter_layout = QHBoxLayout()
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search by asset name...")
        self.search_input.setMaximumWidth(250)
        self.search_input.textChanged.connect(self.load_data)

        self.status_combo = QComboBox()
        self.status_combo.addItem("All Status")
        self.status_combo.addItems(["Active", "Inactive", "Sold", "Disposed"])
        self.status_combo.currentIndexChanged.connect(self.load_data)

        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDate(QDate.currentDate().addYears(-1))
        self.from_date.dateChanged.connect(self.load_data)

        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDate(QDate.currentDate())
        self.to_date.dateChanged.connect(self.load_data)

        self.category_combo = QComboBox()
        self.category_combo.addItem("All Categories")
        self.load_asset_categories()
        self.category_combo.currentIndexChanged.connect(self.load_data)

        filter_layout.addWidget(QLabel("Search:"))
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(QLabel("Status:"))
        filter_layout.addWidget(self.status_combo)
        filter_layout.addWidget(QLabel("From:"))
        filter_layout.addWidget(self.from_date)
        filter_layout.addWidget(QLabel("To:"))
        filter_layout.addWidget(self.to_date)
        filter_layout.addWidget(QLabel("Category:"))
        filter_layout.addWidget(self.category_combo)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        # Metrics
        metrics_layout = QHBoxLayout()
        self.results_label = QLabel("0 assets")
        self.results_label.setStyleSheet("font-size: 14px; color: #718096;")
        
        self.total_value_label = QLabel("Total Value: 0.00")
        self.total_value_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #2D3748;")
        
        self.active_assets_label = QLabel("Active Assets: 0")
        self.active_assets_label.setStyleSheet("font-size: 14px; color: #38A169;")
        
        metrics_layout.addWidget(self.results_label)
        metrics_layout.addWidget(self.total_value_label)
        metrics_layout.addWidget(self.active_assets_label)
        metrics_layout.addStretch()
        layout.addLayout(metrics_layout)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            "ID", "Asset Name", "Category", "Subcategory", "Purchase Amount", 
            "Current Value", "Status", "Last Updated", "Actions"
        ])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        header = self.table.horizontalHeader()
        
        # Set resize modes
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # ID
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Asset Name
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Category
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Subcategory
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Purchase Amount
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Current Value
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)  # Status
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)  # Last Updated
        header.setSectionResizeMode(8, QHeaderView.Fixed)  # Actions
        header.resizeSection(8, 120)
        
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        layout.addWidget(self.table)

        # Load initial data
        self.load_data()

    def load_asset_categories(self):
        """Load only asset categories"""
        try:
            self.category_combo.clear()
            self.category_combo.addItem("All Categories")
            with get_db_connection() as (conn, c):
                c.execute("SELECT ExpCatID, ExpCatName FROM ExpenditureCategory WHERE ExpCatType = 'Assets'")
                for row in c.fetchall():
                    self.category_combo.addItem(row[1], row[0])
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load categories: {str(e)}")

    def load_data(self):
        """Load assets data with filters"""
        search_text = self.search_input.text().strip()
        status_filter = self.status_combo.currentText()
        from_date_str = self.from_date.date().toString("yyyy-MM-dd")
        to_date_str = self.to_date.date().toString("yyyy-MM-dd")
        category_id = self.category_combo.currentData() if self.category_combo.currentIndex() > 0 else None

        try:
            with get_db_connection() as (conn, c):
                query = """
                    SELECT 
                        a.AssetID, a.AssetName, ec.ExpCatName, esc.ExpenditureSubCategoryName,
                        e.ExpenseAmount, a.AssetValue, a.Status, a.updatedON, e.ExpenseDate
                    FROM Assets a
                    JOIN Expenditure e ON a.ExpenseID = e.ExpenseID
                    JOIN ExpenditureSubCategory esc ON e.ExpeSubID = esc.ExpeSubCatID
                    JOIN ExpenditureCategory ec ON esc.ExpenditureCategoryID = ec.ExpCatID
                    WHERE 1=1
                """
                params = []
                
                if search_text:
                    query += " AND (a.AssetName LIKE ? OR a.Status LIKE ?)"
                    params.extend([f"%{search_text}%", f"%{search_text}%"])
                
                if status_filter != "All Status":
                    query += " AND a.Status = ?"
                    params.append(status_filter)
                
                if from_date_str:
                    query += " AND a.updatedON >= ?"
                    params.append(from_date_str)
                
                if to_date_str:
                    query += " AND a.updatedON <= ?"
                    params.append(to_date_str)
                
                if category_id:
                    query += " AND ec.ExpCatID = ?"
                    params.append(category_id)
                
                query += " ORDER BY a.updatedON DESC"
                
                c.execute(query, params)
                rows = c.fetchall()
                
                self.table.setRowCount(len(rows))
                total_value = 0
                active_count = 0
                
                for i, row in enumerate(rows):
                    for j, val in enumerate(row[:8]):  # Exclude ExpenseDate from display
                        item = QTableWidgetItem(str(val) if val else "")
                        
                        # Color coding for status
                        if j == 6:  # Status column
                            if val == "Active":
                                item.setForeground(Qt.darkGreen)
                            elif val == "Inactive":
                                item.setForeground(Qt.darkYellow)
                            elif val in ["Sold", "Disposed"]:
                                item.setForeground(Qt.darkRed)
                        
                        # Format currency columns
                        if j in [4, 5]:  # Amount columns
                            item.setText(f"{float(val):.2f}" if val else "0.00")
                        
                        self.table.setItem(i, j, item)
                    
                    # Calculate metrics
                    asset_value = row[5] or 0
                    total_value += asset_value
                    if row[6] == "Active":
                        active_count += 1
                    
                    self.add_action_buttons(i, row[0])
                
                # Update metrics
                self.results_label.setText(f"{len(rows)} assets")
                self.total_value_label.setText(f"Total Value: {total_value:.2f}")
                self.active_assets_label.setText(f"Active Assets: {active_count}")
                
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load assets: {str(e)}")

    def add_action_buttons(self, row, asset_id):
        """Add action buttons to table row"""
        actions_widget = QWidget()
        actions_layout = QHBoxLayout(actions_widget)
        actions_layout.setContentsMargins(5, 2, 5, 2)
        actions_layout.setSpacing(3)

        edit_btn = QPushButton("Edit")
        edit_btn.setFixedWidth(50)
        edit_btn.setStyleSheet("background-color: #4CAF50; color: white; font-size: 12px;")
        edit_btn.clicked.connect(lambda: self.edit_asset(asset_id))
        
        delete_btn = QPushButton("Delete")
        delete_btn.setFixedWidth(50)
        delete_btn.setStyleSheet("background-color: #f44336; color: white; font-size: 12px;")
        delete_btn.clicked.connect(lambda: self.delete_asset(asset_id))

        actions_layout.addWidget(edit_btn)
        actions_layout.addWidget(delete_btn)
        actions_layout.addStretch()

        self.table.setCellWidget(row, 8, actions_widget)

    def show_add_dialog(self):
        """Show dialog to add new asset"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Add New Asset")
        dialog.setMinimumWidth(600)
        layout = QVBoxLayout(dialog)
        
        form_layout = QFormLayout()
        
        # Asset basic info
        asset_name_edit = QLineEdit()
        asset_name_edit.setPlaceholderText("Enter asset name...")
        
        category_combo = QComboBox()
        category_combo.addItem("Select Category")
        self.load_asset_categories_dialog(category_combo)
        
        subcategory_combo = QComboBox()
        subcategory_combo.addItem("Select Subcategory")
        
        # Connect category change to load subcategories
        category_combo.currentIndexChanged.connect(
            lambda: self.load_subcategories_dialog(
                subcategory_combo, 
                category_combo.currentData() if category_combo.currentIndex() > 0 else None
            )
        )
        
        purchase_amount_edit = QLineEdit()
        purchase_amount_edit.setValidator(QDoubleValidator(0.01, 9999999.99, 2))
        purchase_amount_edit.setPlaceholderText("Purchase amount...")
        
        current_value_edit = QLineEdit()
        current_value_edit.setValidator(QDoubleValidator(0.01, 9999999.99, 2))
        current_value_edit.setPlaceholderText("Current value...")
        
        status_combo = QComboBox()
        status_combo.addItems(["Active", "Inactive", "Sold", "Disposed"])
        
        purchase_date_edit = QDateEdit()
        purchase_date_edit.setCalendarPopup(True)
        purchase_date_edit.setDate(QDate.currentDate())
        
        description_edit = QTextEdit()
        description_edit.setMaximumHeight(80)
        description_edit.setPlaceholderText("Asset description...")
        
        form_layout.addRow("Asset Name*:", asset_name_edit)
        form_layout.addRow("Category*:", category_combo)
        form_layout.addRow("Subcategory*:", subcategory_combo)
        form_layout.addRow("Purchase Amount*:", purchase_amount_edit)
        form_layout.addRow("Current Value*:", current_value_edit)
        form_layout.addRow("Status*:", status_combo)
        form_layout.addRow("Purchase Date*:", purchase_date_edit)
        form_layout.addRow("Description:", description_edit)
        
        layout.addLayout(form_layout)
        
        # Buttons
        button_layout = QHBoxLayout()
        save_btn = QPushButton("Save Asset")
        save_btn.clicked.connect(lambda: self.save_asset(
            dialog, asset_name_edit, category_combo, subcategory_combo, purchase_amount_edit,
            current_value_edit, status_combo, purchase_date_edit, description_edit
        ))
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(dialog.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        
        dialog.exec_()

    def load_asset_categories_dialog(self, combo):
        """Load asset categories for dialog"""
        try:
            combo.clear()
            combo.addItem("Select Category")
            with get_db_connection() as (conn, c):
                c.execute("SELECT ExpCatID, ExpCatName FROM ExpenditureCategory WHERE ExpCatType = 'Assets'")
                for row in c.fetchall():
                    combo.addItem(row[1], row[0])
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load categories: {str(e)}")

    def load_subcategories_dialog(self, combo, cat_id):
        """Load subcategories for dialog"""
        combo.clear()
        combo.addItem("Select Subcategory")
        
        if not cat_id:
            return
            
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT ExpeSubCatID, ExpenditureSubCategoryName FROM ExpenditureSubCategory WHERE ExpenditureCategoryID = ?", (cat_id,))
                for row in c.fetchall():
                    combo.addItem(row[1], row[0])
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load subcategories: {str(e)}")

    def save_asset(self, dialog, asset_name_edit, category_combo, subcategory_combo, purchase_amount_edit,
                  current_value_edit, status_combo, purchase_date_edit, description_edit):
        """Save new asset to database"""
        # Validate inputs
        if not all([
            asset_name_edit.text().strip(),
            category_combo.currentIndex() > 0,
            subcategory_combo.currentIndex() > 0,
            purchase_amount_edit.text().strip(),
            current_value_edit.text().strip()
        ]):
            QMessageBox.warning(dialog, "Error", "Please fill all required fields (*)")
            return
        
        asset_name = asset_name_edit.text().strip()
        subcategory_id = subcategory_combo.currentData()
        status = status_combo.currentText()
        purchase_date = purchase_date_edit.date().toString("yyyy-MM-dd")
        description = description_edit.toPlainText().strip()
        
        try:
            purchase_amount = float(purchase_amount_edit.text())
            current_value = float(current_value_edit.text())
            
            if purchase_amount <= 0 or current_value <= 0:
                QMessageBox.warning(dialog, "Error", "Amounts must be positive")
                return
                
        except ValueError:
            QMessageBox.warning(dialog, "Error", "Please enter valid amounts")
            return
        
        try:
            with get_db_connection() as (conn, c):
                # First create expenditure record
                c.execute("""INSERT INTO Expenditure (ExpeSubID, ExpenseAmount, ExpenseDate, ExpenseDesc, PaymentMethod)
                             VALUES (?, ?, ?, ?, ?)""", 
                         (subcategory_id, purchase_amount, purchase_date, description, "Asset Purchase"))
                expense_id = c.lastrowid
                
                # Then create asset record
                updated_on = QDate.currentDate().toString("yyyy-MM-dd")
                c.execute("""INSERT INTO Assets (ExpenseID, AssetName, AssetValue, Status, updatedON)
                             VALUES (?, ?, ?, ?, ?)""", 
                         (expense_id, asset_name, current_value, status, updated_on))
                
                conn.commit()
                
                QMessageBox.information(dialog, "Success", "Asset added successfully")
                dialog.accept()
                self.load_data()
                
        except Exception as e:
            QMessageBox.warning(dialog, "Error", f"Failed to save asset: {str(e)}")

    def edit_asset(self, asset_id):
        """Edit existing asset"""
        try:
            with get_db_connection() as (conn, c):
                c.execute("""
                    SELECT 
                        a.AssetID, a.AssetName, esc.ExpenditureCategoryID, e.ExpeSubID,
                        e.ExpenseAmount, a.AssetValue, a.Status, a.updatedON, e.ExpenseDate,
                        e.ExpenseDesc
                    FROM Assets a
                    JOIN Expenditure e ON a.ExpenseID = e.ExpenseID
                    JOIN ExpenditureSubCategory esc ON e.ExpeSubID = esc.ExpeSubCatID
                    WHERE a.AssetID = ?
                """, (asset_id,))
                data = c.fetchone()
                
            if not data:
                QMessageBox.warning(self, "Error", "Asset not found")
                return
                
            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Asset")
            dialog.setMinimumWidth(600)
            layout = QVBoxLayout(dialog)
            
            form_layout = QFormLayout()
            
            # Asset basic info
            asset_name_edit = QLineEdit(data[1])
            
            category_combo = QComboBox()
            category_combo.addItem("Select Category")
            self.load_asset_categories_dialog(category_combo)
            category_combo.setCurrentIndex(category_combo.findData(data[2]))
            
            subcategory_combo = QComboBox()
            subcategory_combo.addItem("Select Subcategory")
            self.load_subcategories_dialog(subcategory_combo, data[2])
            subcategory_combo.setCurrentIndex(subcategory_combo.findData(data[3]))
            
            category_combo.currentIndexChanged.connect(
                lambda: self.load_subcategories_dialog(
                    subcategory_combo, 
                    category_combo.currentData() if category_combo.currentIndex() > 0 else None
                )
            )
            
            purchase_amount_edit = QLineEdit(str(data[4]))
            purchase_amount_edit.setValidator(QDoubleValidator(0.01, 9999999.99, 2))
            
            current_value_edit = QLineEdit(str(data[5]))
            current_value_edit.setValidator(QDoubleValidator(0.01, 9999999.99, 2))
            
            status_combo = QComboBox()
            status_combo.addItems(["Active", "Inactive", "Sold", "Disposed"])
            status_combo.setCurrentText(data[6])
            
            purchase_date_edit = QDateEdit(QDate.fromString(data[8], "yyyy-MM-dd"))
            purchase_date_edit.setCalendarPopup(True)
            purchase_date_edit.setEnabled(False)  # Purchase date cannot be changed
            
            description_edit = QTextEdit(data[9] or "")
            description_edit.setMaximumHeight(80)
            
            form_layout.addRow("Asset Name*:", asset_name_edit)
            form_layout.addRow("Category*:", category_combo)
            form_layout.addRow("Subcategory*:", subcategory_combo)
            form_layout.addRow("Purchase Amount*:", purchase_amount_edit)
            form_layout.addRow("Current Value*:", current_value_edit)
            form_layout.addRow("Status*:", status_combo)
            form_layout.addRow("Purchase Date:", purchase_date_edit)
            form_layout.addRow("Description:", description_edit)
            
            layout.addLayout(form_layout)
            
            # Buttons
            button_layout = QHBoxLayout()
            save_btn = QPushButton("Save Changes")
            save_btn.clicked.connect(lambda: self.save_edited_asset(
                dialog, asset_id, asset_name_edit, subcategory_combo, purchase_amount_edit,
                current_value_edit, status_combo, description_edit
            ))
            cancel_btn = QPushButton("Cancel")
            cancel_btn.clicked.connect(dialog.reject)
            
            button_layout.addStretch()
            button_layout.addWidget(save_btn)
            button_layout.addWidget(cancel_btn)
            layout.addLayout(button_layout)
            
            dialog.exec_()
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load asset: {str(e)}")

    def save_edited_asset(self, dialog, asset_id, asset_name_edit, subcategory_combo, purchase_amount_edit,
                         current_value_edit, status_combo, description_edit):
        """Save edited asset"""
        # Validate inputs
        if not all([
            asset_name_edit.text().strip(),
            subcategory_combo.currentIndex() > 0,
            purchase_amount_edit.text().strip(),
            current_value_edit.text().strip()
        ]):
            QMessageBox.warning(dialog, "Error", "Please fill all required fields (*)")
            return
        
        asset_name = asset_name_edit.text().strip()
        subcategory_id = subcategory_combo.currentData()
        status = status_combo.currentText()
        description = description_edit.toPlainText().strip()
        
        try:
            purchase_amount = float(purchase_amount_edit.text())
            current_value = float(current_value_edit.text())
            
            if purchase_amount <= 0 or current_value <= 0:
                QMessageBox.warning(dialog, "Error", "Amounts must be positive")
                return
                
        except ValueError:
            QMessageBox.warning(dialog, "Error", "Please enter valid amounts")
            return
        
        try:
            with get_db_connection() as (conn, c):
                # Get the expense ID for this asset
                c.execute("SELECT ExpenseID FROM Assets WHERE AssetID = ?", (asset_id,))
                expense_id = c.fetchone()[0]
                
                # Update expenditure record
                c.execute("""UPDATE Expenditure SET ExpeSubID=?, ExpenseAmount=?, ExpenseDesc=?
                             WHERE ExpenseID=?""", 
                         (subcategory_id, purchase_amount, description, expense_id))
                
                # Update asset record
                updated_on = QDate.currentDate().toString("yyyy-MM-dd")
                c.execute("""UPDATE Assets SET AssetName=?, AssetValue=?, Status=?, updatedON=?
                             WHERE AssetID=?""", 
                         (asset_name, current_value, status, updated_on, asset_id))
                
                conn.commit()
                
                QMessageBox.information(dialog, "Success", "Asset updated successfully")
                dialog.accept()
                self.load_data()
                
        except Exception as e:
            QMessageBox.warning(dialog, "Error", f"Failed to update asset: {str(e)}")

    def delete_asset(self, asset_id):
        """Delete asset with confirmation"""
        reply = QMessageBox.question(
            self, 'Confirm Delete', 
            f'Are you sure you want to delete this asset?\nThis will also delete the associated expenditure record.', 
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                with get_db_connection() as (conn, c):
                    # Get expense ID first
                    c.execute("SELECT ExpenseID FROM Assets WHERE AssetID = ?", (asset_id,))
                    result = c.fetchone()
                    if result:
                        expense_id = result[0]
                        # Delete asset record
                        c.execute("DELETE FROM Assets WHERE AssetID = ?", (asset_id,))
                        # Delete associated expenditure record
                        c.execute("DELETE FROM Expenditure WHERE ExpenseID = ?", (expense_id,))
                        conn.commit()
                    
                    QMessageBox.information(self, "Success", "Asset deleted successfully")
                    self.load_data()
                    
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to delete asset: {str(e)}")

### end of expenditure ####

class Reports(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addStretch()
        label = QLabel("Reports Page")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2D3748;")
        layout.addWidget(label)
        layout.addStretch()                    

class About(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addStretch()
        label = QLabel("About Page")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2D3748;")
        layout.addWidget(label)
        layout.addStretch()                    

class Account(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addStretch()
        label = QLabel("Account Page")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2D3748;")
        layout.addWidget(label)
        layout.addStretch()                    

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    print("Application opened.")
    window.show()
    sys.exit(app.exec_())
    