import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sqlite3
import pandas as pd
from xlsxwriter import *
from contextlib import contextmanager
from datetime import datetime

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
    wizard = ImportWizard(self)
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
            ("Expenses", "Icons/expenses-icon.png"),
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
        self.ExpensesPage = Expenses()
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
            "Expenses": self.ExpensesPage,
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
        self.payment_btn = QPushButton("Fees Payment")
        self.payment_btn.setObjectName("nav_btn")
        self.payment_btn.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        nav_layout.addWidget(self.management_btn)
        nav_layout.addWidget(self.payment_btn)
        nav_layout.addStretch()
        layout.addLayout(nav_layout)

        # Stacked widget for sub-pages
        self.stack = QStackedWidget()
        self.management_page = FeesManagementPage()
        self.payment_page = FeesPaymentPage()
        self.stack.addWidget(self.management_page)
        self.stack.addWidget(self.payment_page)
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

        # Action buttons
        action_layout = QHBoxLayout()
        generate_btn = QPushButton("Generate Template")
        generate_btn.clicked.connect(self.generate_fee_payment_template)
        import_btn = QPushButton("Import Payments")
        import_btn.clicked.connect(self.import_fee_payments)
        action_layout.addWidget(generate_btn)
        action_layout.addWidget(import_btn)
        action_layout.addStretch()
        layout.addLayout(action_layout)

        # Table - Updated columns
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "StudentTermID", "Name", "FeesID", "AmountDue", 
            "AmountPaid", "TransactionDate", "Reference"
        ])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.table)
        
        # Load initial data
        self.load_data()

    def generate_fee_payment_template(self):
        """Open wizard to generate fee payment template"""
        wizard = FeePaymentWizard(self)
        wizard.exec_()

    def import_fee_payments(self):
        """Simple import function"""
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Select Fee Payment Template", "", "Excel Files (*.xlsx *.xls)"
            )
            
            if file_path:
                # Use the simple import function I provided earlier
                self.simple_import_payments(file_path)
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to import payments: {str(e)}")

    def load_data(self):
            """Load payment data for display"""
            # Implement based on your needs
            pass

class FeePaymentWizard(QWizard):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Generate Fee Payment Template")
        self.setMinimumSize(800, 600)
        
        self.selected_data = None
        self.filtered_students = []
        
        # Add pages
        self.addPage(self.create_selection_page())
        self.addPage(self.create_preview_page())
        self.addPage(self.create_generation_page())
        
    def create_selection_page(self):
        page = QWizardPage()
        page.setTitle("Step 1: Select Filters")
        page.setSubTitle("Choose Academic Year, Class, and Term to include ALL students")
        
        layout = QVBoxLayout()
        
        # Academic Year selection
        year_group = QGroupBox("Academic Year")
        year_layout = QVBoxLayout()
        self.year_combo = QComboBox()
        self.year_combo.currentIndexChanged.connect(self.on_year_changed)
        year_layout.addWidget(self.year_combo)
        year_group.setLayout(year_layout)
        layout.addWidget(year_group)
        
        # Class and Term selection
        filter_group = QGroupBox("Class and Term")
        filter_layout = QGridLayout()
        
        filter_layout.addWidget(QLabel("Class:"), 0, 0)
        self.class_combo = QComboBox()
        filter_layout.addWidget(self.class_combo, 0, 1)
        
        filter_layout.addWidget(QLabel("Term:"), 1, 0)
        self.term_combo = QComboBox()
        filter_layout.addWidget(self.term_combo, 1, 1)
        
        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)
        
        # Information label
        info_label = QLabel("This will include ALL students in the selected class and term, regardless of payment status.")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(info_label)
        
        # Load initial data
        self.load_academic_years()
        
        # Connect signals
        self.class_combo.currentIndexChanged.connect(self.load_students)
        self.term_combo.currentIndexChanged.connect(self.load_students)
        
        page.setLayout(layout)
        return page
    
    def create_preview_page(self):
        page = QWizardPage()
        page.setTitle("Step 2: Preview Students")
        page.setSubTitle("Review ALL students who will be included in the template")
        
        layout = QVBoxLayout()
        
        self.preview_label = QLabel("Select filters in Step 1 to see students")
        self.preview_label.setWordWrap(True)
        
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(8)  # Added PaymentStatus column
        self.preview_table.setHorizontalHeaderLabels([
            "StudentTermID", "Name", "FeesID", "AmountDue", "AmountPaid", 
            "PaymentStatus", "TransactionDate", "Reference"
        ])
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        layout.addWidget(self.preview_label)
        layout.addWidget(self.preview_table)
        
        page.setLayout(layout)
        return page
    
    def create_generation_page(self):
        page = QWizardPage()
        page.setTitle("Step 3: Generate Template")
        page.setSubTitle("Ready to generate the Excel template with ALL students")
        
        layout = QVBoxLayout()
        
        info_label = QLabel(
            "Click 'Generate' to create the Excel template with ALL students in the selected class and term. "
            "The template will include payment status information and columns for payment details."
        )
        info_label.setWordWrap(True)
        
        self.generate_btn = QPushButton("Generate Excel Template")
        self.generate_btn.clicked.connect(self.generate_template)
        
        layout.addWidget(info_label)
        layout.addWidget(self.generate_btn)
        layout.addStretch()
        
        page.setLayout(layout)
        return page
    
    def load_academic_years(self):
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY Year DESC")
                years = c.fetchall()
                
                self.year_combo.addItem("Select Academic Year", None)
                for year_id, year_name in years:
                    self.year_combo.addItem(year_name, year_id)
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load academic years: {str(e)}")
    
    def on_year_changed(self):
        year_id = self.year_combo.currentData()
        if not year_id:
            return
            
        try:
            with get_db_connection() as (conn, c):
                # Load classes for selected year
                self.class_combo.clear()
                c.execute("SELECT ClassID, ClassName FROM Class WHERE AcademicYear = ? ORDER BY ClassName", (year_id,))
                classes = c.fetchall()
                self.class_combo.addItem("Select Class", None)
                for class_id, class_name in classes:
                    self.class_combo.addItem(class_name, class_id)
                
                # Load terms for selected year
                self.term_combo.clear()
                c.execute("SELECT TermID, TermName FROM Term WHERE AcademicYear = ? ORDER BY TermName", (year_id,))
                terms = c.fetchall()
                self.term_combo.addItem("Select Term", None)
                for term_id, term_name in terms:
                    self.term_combo.addItem(term_name, term_id)
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load classes and terms: {str(e)}")
    
    def load_students(self):
        class_id = self.class_combo.currentData()
        term_id = self.term_combo.currentData()
        
        if not class_id or not term_id:
            return
            
        try:
            with get_db_connection() as (conn, c):
                query = """
                    SELECT 
                        st.StudentTermID,
                        s.Student_Name as Name,
                        f.FeesID,
                        (f.Amount * (st.PayRate / 100)) as AmountDue,
                        COALESCE(SUM(fp.AmountPaid), 0) as AmountPaid,
                        CASE 
                            WHEN COALESCE(SUM(fp.AmountPaid), 0) >= (f.Amount * (st.PayRate / 100)) THEN 'Fully Paid'
                            WHEN COALESCE(SUM(fp.AmountPaid), 0) > 0 THEN 'Partially Paid'
                            ELSE 'Not Paid'
                        END as PaymentStatus,
                        '' as TransactionDate,  -- Placeholder for user input
                        '' as Reference         -- Placeholder for user input
                    FROM StudentTerm st
                    JOIN Students s ON st.StudentID = s.StudentID
                    JOIN Fees f ON f.Class = st.ClassID AND f.Term = st.TermID
                    LEFT JOIN FeesPayment fp ON st.StudentTermID = fp.StudentTermID
                    WHERE st.ClassID = ? AND st.TermID = ?
                    GROUP BY st.StudentTermID
                    ORDER BY s.Student_Name
                """
                
                c.execute(query, (class_id, term_id))
                self.filtered_students = c.fetchall()
                
                # Update preview table
                self.preview_table.setRowCount(0)
                for row_data in self.filtered_students:
                    row = self.preview_table.rowCount()
                    self.preview_table.insertRow(row)
                    
                    for col, data in enumerate(row_data):
                        item = QTableWidgetItem(str(data))
                        # Make only TransactionDate and Reference editable
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable if col not in [6, 7] else item.flags())
                        
                        # Color code payment status
                        if col == 5:  # PaymentStatus column
                            if data == 'Fully Paid':
                                item.setBackground(QColor(200, 255, 200))  # Light green
                            elif data == 'Partially Paid':
                                item.setBackground(QColor(255, 255, 200))  # Light yellow
                            else:
                                item.setBackground(QColor(255, 200, 200))  # Light red
                        
                        self.preview_table.setItem(row, col, item)
                
                count = len(self.filtered_students)
                class_name = self.class_combo.currentText()
                term_name = self.term_combo.currentText()
                
                self.preview_label.setText(
                    f"Found {count} students in {class_name} for {term_name}\n"
                    f"Payment status is color-coded: Green = Fully Paid, Yellow = Partially Paid, Red = Not Paid\n"
                    f"TransactionDate and Reference columns are for your input in Excel."
                )
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load students: {str(e)}")
    
    def generate_template(self):
        if not self.filtered_students:
            QMessageBox.warning(self, "No Data", "No students found to generate template.")
            return
            
        try:
            # Convert to DataFrame
            data = []
            for student in self.filtered_students:
                data.append({
                    'StudentTermID': student[0],
                    'Name': student[1],
                    'FeesID': student[2],
                    'AmountDue': student[3],
                    'AmountPaid': student[4],
                    'PaymentStatus': student[5],
                    'TransactionDate': student[6],  # Empty for user input
                    'Reference': student[7]         # Empty for user input
                })
            
            df = pd.DataFrame(data)
            
            # Save to Excel
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save Fee Payment Template", 
                f"Fee_Payment_Template_{self.class_combo.currentText()}_{self.term_combo.currentText()}.xlsx",
                "Excel Files (*.xlsx)"
            )
            
            if file_path:
                with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name='Fee Payments', index=False)
                    
                    workbook = writer.book
                    worksheet = writer.sheets['Fee Payments']
                    
                    # Format headers
                    header_format = workbook.add_format({
                        'bold': True,
                        'bg_color': '#3498db',
                        'font_color': 'white',
                        'border': 1
                    })
                    
                    # Format input columns
                    input_format = workbook.add_format({
                        'bg_color': '#FFF2CC',
                        'border': 1
                    })
                    
                    # Format read-only columns
                    readonly_format = workbook.add_format({
                        'bg_color': '#F0F0F0',
                        'border': 1,
                        'locked': True
                    })
                    
                    # Format payment status columns
                    paid_format = workbook.add_format({'bg_color': '#C6EFCE', 'border': 1, 'locked': True})
                    partial_format = workbook.add_format({'bg_color': '#FFEB9C', 'border': 1, 'locked': True})
                    unpaid_format = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1, 'locked': True})
                    
                    # Apply formatting
                    for col_num, value in enumerate(df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                        
                        # Set column widths
                        if value == 'Name':
                            worksheet.set_column(col_num, col_num, 25)
                        elif value in ['TransactionDate', 'Reference']:
                            worksheet.set_column(col_num, col_num, 20)
                        else:
                            worksheet.set_column(col_num, col_num, 15)
                        
                        # Format cells
                        if value in ['TransactionDate', 'Reference']:
                            worksheet.set_column(col_num, col_num, 20, input_format)
                        elif value == 'PaymentStatus':
                            # Apply conditional formatting for payment status
                            for row in range(1, len(df) + 1):
                                status = df.iloc[row-1]['PaymentStatus']
                                if status == 'Fully Paid':
                                    worksheet.write(row, col_num, status, paid_format)
                                elif status == 'Partially Paid':
                                    worksheet.write(row, col_num, status, partial_format)
                                else:
                                    worksheet.write(row, col_num, status, unpaid_format)
                        else:
                            worksheet.set_column(col_num, col_num, 15, readonly_format)
                    
                    # Add data validation for AmountPaid (if needed)
                    # Add date validation for TransactionDate
                    date_col = df.columns.get_loc('TransactionDate')
                    date_validation = {
                        'validate': 'date',
                        'criteria': 'between',
                        'minimum': datetime(2000, 1, 1),
                        'maximum': datetime(2100, 12, 31),
                        'input_title': 'Enter date',
                        'input_message': 'Date must be in YYYY-MM-DD format',
                        'error_title': 'Invalid Date',
                        'error_message': 'Please enter a valid date in YYYY-MM-DD format'
                    }
                    worksheet.data_validation(1, date_col, len(df), date_col, date_validation)
                
                # Show success message
                QMessageBox.information(self, "Success", 
                                      f"Template generated successfully!\n"
                                      f"Saved to: {file_path}\n\n"
                                      f"Students included: {len(self.filtered_students)}\n"
                                      f"Class: {self.class_combo.currentText()}\n"
                                      f"Term: {self.term_combo.currentText()}\n\n"
                                      f"Template includes ALL students with payment status color-coding.")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate template: {str(e)}")
#### End Of Fees Management ######

class Uniforms(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addStretch()
        label = QLabel("Uniforms Page")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2D3748;")
        layout.addWidget(label)
        layout.addStretch()

class Expenses(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addStretch()
        label = QLabel("Expenses Page")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2D3748;")
        layout.addWidget(label)
        layout.addStretch()                    

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
    