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
from dashboard import Dashboard
from reports_page import Reports


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
                background-color: #F7FAFC;
            }
            QWidget#sidebar {
                background-color: #FFFFFF;
                border-right: 1px solid #E0E3E7;
            }
            QWidget#centralWidget {
                background-color: #F7FAFC;
            }
            QWidget#contentWidget {
                background-color: #F7FAFC;
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
                background: transparent;
                border: none;
            }
            QLabel#school_name {
                font-size: 18px;
                color: #4A5568;
                background: transparent;
                border: none;
            }
            QWidget#separator {
                background-color: #E0E3E7;
            }
            QScrollArea {
                background-color: #F7FAFC;
                border: none;
                border-radius: 5px;
            }
            QScrollArea > QWidget {
                background-color: #F7FAFC;
            }
        """)
        # Set object names for specific styling
        self.central_widget.setObjectName("centralWidget")
        self.content.setObjectName("contentWidget")
        self.sidebar.setObjectName("sidebar")
        self.toggle_btn.setObjectName("toggle_btn")
        self.page_title.setObjectName("page_title")
        self.school_name.setObjectName("school_name")
        self.separator.setObjectName("separator")
        # Ensure body widget has proper background
        self.body.setStyleSheet("QStackedWidget { background-color: #F7FAFC; }")
    
class AcademicYearTerm(QWidget):
    def __init__(self):
        super().__init__()
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
        
        self.year_action_combo.addItem("Year", None)
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
                    SELECT AcademicYear.Year, Term.TermName, Term.StartDate, Term.EndDate, Term.TermID
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
                    
                    # Display only first 4 columns (hide TermID)
                    for col, data in enumerate(row_data[:4]):
                        self.table.setItem(row, col, QTableWidgetItem(str(data)))
                    
                    # Actions - ComboBox
                    action_combo = QComboBox()
                    action_combo.addItems(["Select Action", "Edit", "Delete"])
                    
                    # Store term_id (5th column)
                    term_id = row_data[4]
                    
                    # Connect with lambda using term_id
                    action_combo.currentIndexChanged.connect(
                        lambda index, tid=term_id, cb=action_combo: 
                        self.on_term_action_selected_by_id(index, tid, cb)
                    )
                    
                    self.table.setCellWidget(row, 4, action_combo)
                # Ensure header stretches after data reload
                self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
                self.update_results_count()
                
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to load data: {str(e)}")
    
    def on_term_action_selected_by_id(self, index, term_id, combo_box):
        """Handle term action selection using term_id"""
        if index == 0:
            return
        
        action = combo_box.itemText(index)
        combo_box.setCurrentIndex(0)
        
        
        if action == "Edit":
            self.editTerm_by_id(term_id)
        elif action == "Delete":
            self.deleteTerm_by_id(term_id)

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
                        self.Load_Data()
                        
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
                    c.execute("SELECT 1 FROM Fees WHERE TermID = ?", (term_id,))
                    if c.fetchone():
                        QMessageBox.critical(self, "Restricted", 
                            "This term cannot be deleted because it is linked to one or more fees records.")
                        return

                    # Set Student.TermID to NULL for students assigned to this term
                    c.execute("UPDATE Student SET TermID = NULL WHERE TermID = ?", (term_id,))
                    
                    # Delete the term
                    c.execute("DELETE FROM Term WHERE TermID = ?", (term_id,))
                    conn.commit()
                    
                    QMessageBox.information(self, "Deleted", 
                        f"Term '{term}' for academic year '{year}' has been deleted.\n\n"
                        "Any students previously assigned to this term now have no term assigned.")
                    
                    # Refresh the table
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

class Class(QWidget):
    def __init__(self):
        super().__init__()
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
                    self.table.setRowHeight(row,40)
                    

                
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
            print(f"Error loading filter options: {e}")

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
                print(f"Error loading classes: {e}")
        
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
                print(f"Error loading terms: {e}")

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
            print(f"Error loading payment history: {e}")
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
        uniform_layout = QVBoxLayout(uniform_frame)
        
        uni_header = QHBoxLayout()
        uni_title = QLabel("Uniform Catalog (CRUD)")
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
            print(f"Error loading uniform list: {e}")

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
                    self.table.setRowHeight(i, 40)
                
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

### end of expenditure ####


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
    
    # Load global stylesheet
    try:
        with open("theme.qss", "r") as f:
            stylesheet = f.read()
            app.setStyleSheet(stylesheet)
    except FileNotFoundError:
        print("Warning: theme.qss not found. Running without global stylesheet.")
    
    window = Window()
    print("Application opened.")
    window.show()
    sys.exit(app.exec_())
    