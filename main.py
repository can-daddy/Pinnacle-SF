import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sqlite3
import pandas as pd

# Connect to the SQLite database
def get_db_connection():
    conn = sqlite3.connect("Database.db")
    return conn, conn.cursor()
conn, c = get_db_connection()

 # Assuming get_db_connection() is defined elsewhere, e.g.:
    # def get_db_connection():
    #     global conn, c
    #     conn = sqlite3.connect('school.db')
    #     c = conn.cursor()

class Student(QWidget):
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
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(15)

        # Header row
        header_layout = QHBoxLayout()
        title = QLabel("Students")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        header_layout.addWidget(title)
        header_layout.addStretch()
        add_student_btn = QPushButton("Add Student")
        add_student_btn.clicked.connect(self.show_add_student_dialog)
        import_btn = QPushButton("Import Excel")
        import_btn.clicked.connect(self.show_import_excel_wizard)
        header_layout.addWidget(add_student_btn)
        header_layout.addWidget(import_btn)
        layout.addLayout(header_layout)

        # Filter row
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search...")
        self.search_input.setMaximumWidth(200)
        self.search_input.textChanged.connect(self.load_data)
        
        get_db_connection()  # Open connection for populating combos
        
        self.class_filter = QComboBox()
        self.class_filter.addItem("All Classes", None)
        c.execute("SELECT ClassID, ClassName FROM Class")
        for cl in c.fetchall():
            self.class_filter.addItem(cl[1], cl[0])
        self.class_filter.currentIndexChanged.connect(self.load_data)
        
        self.term_filter = QComboBox()
        self.term_filter.addItem("All Terms", None)
        c.execute("SELECT TermID, TermName FROM Term")
        for t in c.fetchall():
            self.term_filter.addItem(t[1], t[0])
        self.term_filter.currentIndexChanged.connect(self.load_data)
        
        self.filterbyStatus = QComboBox()
        self.filterbyStatus.addItems(["All Status", "Full Scholarship", "75% Scholarship", "No Scholarship"])
        self.filterbyStatus.currentIndexChanged.connect(self.load_data)
        
        self.StudentOnlycheck = QCheckBox("Student Only")
        self.StudentOnlycheck.stateChanged.connect(self.load_data)
        
        self.Filter_FeePayer = QComboBox()
        self.Filter_FeePayer.addItem("All Fee Payers")
        c.execute("SELECT DISTINCT Feepayer From Students")
        for fp in c.fetchall():
            self.Filter_FeePayer.addItem(fp[0])
        self.Filter_FeePayer.currentIndexChanged.connect(self.load_data)
        
        #conn.close()  # Close connection after populating combos
        
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(self.class_filter)
        filter_layout.addWidget(self.term_filter)
        filter_layout.addWidget(self.filterbyStatus)
        filter_layout.addWidget(self.StudentOnlycheck)
        filter_layout.addWidget(self.Filter_FeePayer)
        filter_layout.addStretch()
        self.results_label = QLabel("0 results")
        self.results_label.setStyleSheet("font-size: 14px; color: #718096;")
        filter_layout.addWidget(self.results_label)
        layout.addLayout(filter_layout)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(10)  # Added columns for actions
        self.table.setHorizontalHeaderLabels(["StudentID", "Name", "Fee Payer", "Status", "Parent Name", "Contact", "Contact2", "Class", "Term", "Actions"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        
        layout.addWidget(self.table)
        self.load_data()

    def populate_combos(self, class_combo, term_combo):
        get_db_connection()
        c = conn.cursor()
        c.execute("SELECT ClassID, ClassName FROM Class")
        classes = c.fetchall()
        for cl in classes:
            class_combo.addItem(cl[1], cl[0])
        c.execute("SELECT TermID, TermName FROM Term")
        terms = c.fetchall()
        for t in terms:
            term_combo.addItem(t[1], t[0])

    def load_data(self):
        get_db_connection()
        c = conn.cursor()
        
        # Determine which table to query based on Student Only checkbox
        if self.StudentOnlycheck.isChecked():
            # Query only Students table
            query = """SELECT s.StudentID, s.Student_Name, s.FeePayer, s.Status, 
                            s.Parent_Name, s.Parent_Contact, s.Parent_Contact2,
                            'N/A' as ClassName, 'N/A' as TermName
                    FROM Students s"""
        else:
            # Query joined tables
            query = """SELECT s.StudentID, s.Student_Name, s.FeePayer, s.Status, 
                            s.Parent_Name, s.Parent_Contact, s.Parent_Contact2,
                            cl.ClassName, t.TermName
                    FROM Students s
                    LEFT JOIN StudentTerm st ON s.StudentID = st.StudentID
                    LEFT JOIN Class cl ON st.ClassID = cl.ClassID
                    LEFT JOIN Term t ON st.TermID = t.TermID"""
        
        filters = []
        params = ()
        
        # Search filter
        search_text = self.search_input.text().strip()
        if search_text:
            filters.append("(s.StudentID LIKE ? OR s.Student_Name LIKE ? OR s.Parent_Name LIKE ? OR s.Parent_Contact LIKE ? OR s.Parent_Contact2 LIKE ?)")
            params += (f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%", f"%{search_text}%")
        
        # Class filter (only if not Student Only mode)
        if not self.StudentOnlycheck.isChecked():
            class_id = self.class_filter.currentData()
            if class_id:
                filters.append("cl.ClassID = ?")
                params += (class_id,)
        
        # Term filter (only if not Student Only mode)
        if not self.StudentOnlycheck.isChecked():
            term_id = self.term_filter.currentData()
            if term_id:
                filters.append("t.TermID = ?")
                params += (term_id,)
        
        # Status filter
        status_filter = self.filterbyStatus.currentText()
        if status_filter != "All Status":
            filters.append("s.Status = ?")
            params += (status_filter,)
        
        # Fee Payer filter
        fee_payer_filter = self.Filter_FeePayer.currentText()
        if fee_payer_filter != "All Fee Payers":
            filters.append("s.FeePayer = ?")
            params += (fee_payer_filter,)
        
        # Apply filters if any
        if filters:
            query += " WHERE " + " AND ".join(filters)
        
        # Execute query
        c.execute(query, params)
        rows = c.fetchall()
        
        # Populate table
        self.table.setRowCount(0)
        for row_data in rows:
            row = self.table.rowCount()
            self.table.insertRow(row)
            
            # Add data columns
            for col, data in enumerate(row_data):
                self.table.setItem(row, col, QTableWidgetItem(str(data) if data else ""))
            
            # Add action buttons
            actions_widget = QWidget()
            actions_layout = QHBoxLayout(actions_widget)
            actions_layout.setContentsMargins(5, 5, 5, 5)
            actions_layout.setSpacing(5)
            
            edit_btn = QPushButton("Edit")
            edit_btn.setStyleSheet("""
                QPushButton {
                    background-color: #48BB78;
                    color: white;
                    border: none;
                    padding: 5px;
                    border-radius: 3px;
                    font-size: 12px;
                }
                QPushButton:hover {
                    background-color: #38A169;
                }
            """)
            edit_btn.clicked.connect(lambda checked, student_id=row_data[0]: self.edit_student(student_id))
            
            delete_btn = QPushButton("Delete")
            delete_btn.setStyleSheet("""
                QPushButton {
                    background-color: #F56565;
                    color: white;
                    border: none;
                    padding: 5px;
                    border-radius: 3px;
                    font-size: 12px;
                }
                QPushButton:hover {
                    background-color: #E53E3E;
                }
            """)
            delete_btn.clicked.connect(lambda checked, student_id=row_data[0]: self.delete_student(student_id))
            
            actions_layout.addWidget(edit_btn)
            actions_layout.addWidget(delete_btn)
            actions_layout.addStretch()
            
            self.table.setCellWidget(row, 9, actions_widget)
        
        # Update results count
        count = self.table.rowCount()
        self.results_label.setText(f"{count} results")
        
        # #conn.close()  # Close connection after loading data

    def edit_student(self, student_id):
        # Check if we're in Student Only mode
        if self.StudentOnlycheck.isChecked():
            # Edit only student information
            self.edit_student_info(student_id)
        else:
            # Edit both student and student term information
            self.edit_student_and_term(student_id)

    def delete_student(self, student_id):
        # Confirm deletion
        reply = QMessageBox.question(self, 'Confirm Deletion', 
                                    f'Are you sure you want to delete student {student_id}?',
                                    QMessageBox.Yes | QMessageBox.No, 
                                    QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            get_db_connection()
            c = conn.cursor()
            
            try:
                # Check if we're in Student Only mode
                if self.StudentOnlycheck.isChecked():
                    # Delete only from Students table
                    c.execute("DELETE FROM Students WHERE StudentID = ?", (student_id,))
                else:
                    # Delete from both StudentTerm and Students tables
                    c.execute("DELETE FROM StudentTerm WHERE StudentID = ?", (student_id,))
                    c.execute("DELETE FROM Students WHERE StudentID = ?", (student_id,))
                
                conn.commit()
                QMessageBox.information(self, "Success", "Student deleted successfully")
                self.load_data()  # Refresh the table
                
            except Exception as e:
                conn.rollback()
                QMessageBox.critical(self, "Error", f"Failed to delete student: {str(e)}")
            finally:
                conn.close()

    def edit_student_info(self, student_id):
        get_db_connection()
        c = conn.cursor()
        c.execute("SELECT * FROM Students WHERE StudentID = ?", (student_id,))
        student_data = c.fetchone()
        #conn.close()
        
        if student_data:
            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Student")
            dialog.setModal(True)
            dialog.setMinimumWidth(500)
            
            layout = QVBoxLayout(dialog)
            
            form_layout = QFormLayout()
            
            name_edit = QLineEdit(student_data[1])
            fee_payer_combo = QComboBox()
            fee_payer_combo.addItems(["School", "Parents", "Sponsor"])
            fee_payer_combo.setCurrentText(student_data[2])
            status_combo = QComboBox()
            status_combo.addItems(["Full Scholarship", "75% Scholarship", "No Scholarship"])
            status_combo.setCurrentText(student_data[3])
            parent_name_edit = QLineEdit(student_data[4])
            parent_contact_edit = QLineEdit(student_data[5])
            parent_contact2_edit = QLineEdit(student_data[6] if student_data[6] else "")
            
            form_layout.addRow("Student Name:", name_edit)
            form_layout.addRow("Fee Payer:", fee_payer_combo)
            form_layout.addRow("Status:", status_combo)
            form_layout.addRow("Parent Name:", parent_name_edit)
            form_layout.addRow("Parent Contact:", parent_contact_edit)
            form_layout.addRow("Parent Contact 2:", parent_contact2_edit)
            
            layout.addLayout(form_layout)
            
            button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            button_box.accepted.connect(dialog.accept)
            button_box.rejected.connect(dialog.reject)
            layout.addWidget(button_box)
            
            if dialog.exec_() == QDialog.Accepted:
                get_db_connection()
                c = conn.cursor()
                try:
                    c.execute("""UPDATE Students 
                                SET Student_Name=?, FeePayer=?, Status=?, 
                                    Parent_Name=?, Parent_Contact=?, Parent_Contact2=?
                                WHERE StudentID=?""",
                            (name_edit.text(), fee_payer_combo.currentText(), status_combo.currentText(),
                            parent_name_edit.text(), parent_contact_edit.text(), parent_contact2_edit.text(),
                            student_id))
                    conn.commit()
                    QMessageBox.information(self, "Success", "Student updated successfully")
                    self.load_data()
                except Exception as e:
                    conn.rollback()
                    QMessageBox.critical(self, "Error", f"Failed to update student: {str(e)}")
                finally:
                    conn.close()

    def edit_student_and_term(self, student_id):
        get_db_connection()
        c = conn.cursor()
        
        c.execute("SELECT * FROM Students WHERE StudentID = ?", (student_id,))
        student_data = c.fetchone()
        
        c.execute("""SELECT st.ClassID, st.TermID, cl.ClassName, t.TermName 
                    FROM StudentTerm st
                    LEFT JOIN Class cl ON st.ClassID = cl.ClassID
                    LEFT JOIN Term t ON st.TermID = t.TermID
                    WHERE st.StudentID = ?""", (student_id,))
        term_data = c.fetchone()
        
        #conn.close()
        
        if student_data:
            dialog = QDialog(self)
            dialog.setWindowTitle("Edit Student and Term")
            dialog.setModal(True)
            dialog.setMinimumWidth(500)
            
            layout = QVBoxLayout(dialog)
            
            student_group = QGroupBox("Student Information")
            student_layout = QFormLayout(student_group)
            
            name_edit = QLineEdit(student_data[1])
            fee_payer_combo = QComboBox()
            fee_payer_combo.addItems(["School", "Parents", "Sponsor"])
            fee_payer_combo.setCurrentText(student_data[2])
            status_combo = QComboBox()
            status_combo.addItems(["Full Scholarship", "75% Scholarship", "No Scholarship"])
            status_combo.setCurrentText(student_data[3])
            parent_name_edit = QLineEdit(student_data[4])
            parent_contact_edit = QLineEdit(student_data[5])
            parent_contact2_edit = QLineEdit(student_data[6] if student_data[6] else "")
            
            student_layout.addRow("Student Name:", name_edit)
            student_layout.addRow("Fee Payer:", fee_payer_combo)
            student_layout.addRow("Status:", status_combo)
            student_layout.addRow("Parent Name:", parent_name_edit)
            student_layout.addRow("Parent Contact:", parent_contact_edit)
            student_layout.addRow("Parent Contact 2:", parent_contact2_edit)
            
            layout.addWidget(student_group)
            
            term_group = QGroupBox("Term Information")
            term_layout = QFormLayout(term_group)
            
            class_combo = QComboBox()
            class_combo.addItem("Select Class", None)
            get_db_connection()
            c.execute("SELECT ClassID, ClassName FROM Class")
            for cl in c.fetchall():
                class_combo.addItem(cl[1], cl[0])
            c.execute("SELECT TermID, TermName FROM Term")
            term_combo = QComboBox()
            term_combo.addItem("Select Term", None)
            for t in c.fetchall():
                term_combo.addItem(t[1], t[0])
            #conn.close()
            
            if term_data:
                class_combo.setCurrentIndex(class_combo.findData(term_data[0]))
                term_combo.setCurrentIndex(term_combo.findData(term_data[1]))
            
            term_layout.addRow("Class:", class_combo)
            term_layout.addRow("Term:", term_combo)
            
            layout.addWidget(term_group)
            
            button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
            button_box.accepted.connect(dialog.accept)
            button_box.rejected.connect(dialog.reject)
            layout.addWidget(button_box)
            
            if dialog.exec_() == QDialog.Accepted:
                get_db_connection()
                c = conn.cursor()
                try:
                    c.execute("""UPDATE Students 
                                SET Student_Name=?, FeePayer=?, Status=?, 
                                    Parent_Name=?, Parent_Contact=?, Parent_Contact2=?
                                WHERE StudentID=?""",
                            (name_edit.text(), fee_payer_combo.currentText(), status_combo.currentText(),
                            parent_name_edit.text(), parent_contact_edit.text(), parent_contact2_edit.text(),
                            student_id))
                    
                    class_id = class_combo.currentData()
                    term_id = term_combo.currentData()
                    
                    if class_id and term_id:
                        c.execute("SELECT COUNT(*) FROM StudentTerm WHERE StudentID = ?", (student_id,))
                        if c.fetchone()[0] > 0:
                            c.execute("UPDATE StudentTerm SET ClassID = ?, TermID = ? WHERE StudentID = ?",
                                      (class_id, term_id, student_id))
                        else:
                            c.execute("INSERT INTO StudentTerm (StudentID, ClassID, TermID) VALUES (?, ?, ?)",
                                      (student_id, class_id, term_id))
                    
                    conn.commit()
                    QMessageBox.information(self, "Success", "Student and term updated successfully")
                    self.load_data()
                except Exception as e:
                    conn.rollback()
                    QMessageBox.critical(self, "Error", f"Failed to update student: {str(e)}")
                finally:
                    conn.close()

    def show_add_student_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Add Student")
        layout = QVBoxLayout()
        form_layout = QFormLayout()
        id_edit = QLineEdit()
        name_edit = QLineEdit()
        feepayer_edit = QLineEdit()
        status_combo = QComboBox()
        status_combo.addItems(["Full Scholarship", "75% Scholarship", "No Scholarship"])
        parent_name_edit = QLineEdit()
        contact_edit = QLineEdit()
        contact2_edit = QLineEdit()
        class_combo = QComboBox()
        term_combo = QComboBox()
        self.populate_combos(class_combo, term_combo)
        form_layout.addRow("Student ID:", id_edit)
        form_layout.addRow("Name:", name_edit)
        form_layout.addRow("Fee Payer:", feepayer_edit)
        form_layout.addRow("Status:", status_combo)
        form_layout.addRow("Parent Name:", parent_name_edit)
        form_layout.addRow("Parent Contact:", contact_edit)
        form_layout.addRow("Parent Contact 2:", contact2_edit)
        form_layout.addRow("Class:", class_combo)
        form_layout.addRow("Term:", term_combo)
        layout.addLayout(form_layout)
        buttons = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(lambda: self.save_student(dialog, id_edit, name_edit, feepayer_edit, status_combo, parent_name_edit, contact_edit, contact2_edit, class_combo, term_combo))
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(dialog.close)
        buttons.addWidget(save_btn)
        buttons.addWidget(close_btn)
        layout.addLayout(buttons)
        dialog.setLayout(layout)
        dialog.exec_()

    def save_student(self, dialog, id_edit, name_edit, feepayer_edit, status_combo, parent_name_edit, contact_edit, contact2_edit, class_combo, term_combo):
        student_id = id_edit.text().strip()
        name = name_edit.text().strip()
        feepayer = feepayer_edit.text().strip()
        status = status_combo.currentText()
        parent_name = parent_name_edit.text().strip()
        contact = contact_edit.text().strip()
        contact2 = contact2_edit.text().strip()
        class_id = class_combo.currentData()
        term_id = term_combo.currentData()
        if not all([student_id, name, feepayer, status, parent_name, contact, class_id, term_id]):
            QMessageBox.warning(dialog, "Error", "Required fields are missing")
            return
        get_db_connection()
        c = conn.cursor()
        try:
            c.execute("SELECT COUNT(*) FROM Students WHERE StudentID = ?", (student_id,))
            if c.fetchone()[0] > 0:
                QMessageBox.warning(dialog, "Error", "Duplicate Student ID")
                return
            c.execute("SELECT COUNT(*) FROM StudentTerm WHERE StudentID = ? AND ClassID = ? AND TermID = ?", (student_id, class_id, term_id))
            if c.fetchone()[0] > 0:
                QMessageBox.warning(dialog, "Error", f"Student is already enrolled in {class_combo.currentText()} for {term_combo.currentText()}")
                return
            c.execute("""INSERT INTO Students (StudentID, Student_Name, FeePayer, Status, Parent_Name, Parent_Contact, Parent_Contact2)
                         VALUES (?, ?, ?, ?, ?, ?, ?)""", (student_id, name, feepayer, status, parent_name, contact, contact2))
            c.execute("""INSERT INTO StudentTerm (StudentID, ClassID, TermID)
                         VALUES (?, ?, ?)""", (student_id, class_id, term_id))
            conn.commit()
            QMessageBox.information(dialog, "Success", "Student added successfully")
            id_edit.clear()
            name_edit.clear()
            feepayer_edit.clear()
            parent_name_edit.clear()
            contact_edit.clear()
            contact2_edit.clear()
            class_combo.setCurrentIndex(0)
            term_combo.setCurrentIndex(0)
            status_combo.setCurrentIndex(0)
            self.load_data()
        except Exception as e:
            conn.rollback()
            QMessageBox.warning(dialog, "Error", f"Failed to add student: {str(e)}")
        finally:
            conn.close()

    def show_import_excel_wizard(self):
        wizard = QWizard(self)
        wizard.setWindowTitle("Import Students from Excel")
        page1 = self.create_select_class_term_page()
        page2 = self.create_select_file_page(wizard)
        wizard.addPage(page1)
        wizard.addPage(page2)
        wizard.setOption(QWizard.NoBackButtonOnStartPage, True)
        if wizard.exec_():
            self.load_data()

    def create_select_class_term_page(self):
        page = QWizardPage()
        page.setTitle("Select Import Mode, Class, and Term")
        layout = QVBoxLayout()
        mode_combo = QComboBox()
        mode_combo.addItems(["Register and Enroll New Students", "Enroll Existing Students Only"])
        class_combo = QComboBox()
        class_combo.addItem("Select Class", None)
        term_combo = QComboBox()
        term_combo.addItem("Select Term", None)
        self.populate_combos(class_combo, term_combo)
        layout.addWidget(QLabel("Import Mode:"))
        layout.addWidget(mode_combo)
        layout.addWidget(QLabel("Select Class:"))
        layout.addWidget(class_combo)
        layout.addWidget(QLabel("Select Term:"))
        layout.addWidget(term_combo)
        page.setLayout(layout)
        page.registerField("mode*", mode_combo)
        page.registerField("class_id*", class_combo, "currentData")
        page.registerField("term_id*", term_combo, "currentData")
        def is_complete():
            return (mode_combo.currentIndex() >= 0 and
                    class_combo.currentData() is not None and
                    term_combo.currentData() is not None)
        page.isComplete = is_complete
        mode_combo.currentIndexChanged.connect(page.completeChanged.emit)
        class_combo.currentIndexChanged.connect(page.completeChanged.emit)
        term_combo.currentIndexChanged.connect(page.completeChanged.emit)
        return page

    def create_select_file_page(self, wizard):
        page = QWizardPage()
        page.setTitle("Select Excel File")
        layout = QVBoxLayout()
        file_label = QLabel("No file selected")
        select_btn = QPushButton("Select Excel File")
        layout.addWidget(QLabel("Choose an Excel file to import:"))
        layout.addWidget(select_btn)
        layout.addWidget(file_label)
        layout.addStretch()
        page.setLayout(layout)
        page.file_path = None
        def select_file():
            file_path, _ = QFileDialog.getOpenFileName(page, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
            if file_path:
                page.file_path = file_path
                file_label.setText(file_path.split('/')[-1])
                page.completeChanged.emit()
        select_btn.clicked.connect(select_file)
        def is_complete():
            return page.file_path is not None
        def validate_page():
            mode = wizard.field("mode")
            class_id = wizard.field("class_id")
            term_id = wizard.field("term_id")
            try:
                df = pd.read_excel(page.file_path)
                get_db_connection()
                c = conn.cursor()
                if mode == "Register and Enroll New Students":
                    required_cols = ["StudentID", "Student_Name", "FeePayer", "Status", "Parent_Name", "Parent_Contact"]
                    if not all(col in df.columns for col in required_cols):
                        QMessageBox.warning(page, "Error", "Excel missing required columns: StudentID, Student_Name, FeePayer, Status, Parent_Name, Parent_Contact")
                        return False
                    duplicates = []
                    duplicate_terms = []
                    inserted = 0
                    for _, row in df.iterrows():
                        student_id = str(row["StudentID"])
                        c.execute("SELECT COUNT(*) FROM Students WHERE StudentID = ?", (student_id,))
                        if c.fetchone()[0] > 0:
                            duplicates.append(student_id)
                            continue
                        c.execute("SELECT COUNT(*) FROM StudentTerm WHERE StudentID = ? AND ClassID = ? AND TermID = ?", (student_id, class_id, term_id))
                        if c.fetchone()[0] > 0:
                            duplicate_terms.append(student_id)
                            continue
                        parent_contact2 = str(row.get("Parent_Contact2", ""))
                        c.execute("""INSERT INTO Students (StudentID, Student_Name, FeePayer, Status, Parent_Name, Parent_Contact, Parent_Contact2)
                                     VALUES (?, ?, ?, ?, ?, ?, ?)""", (student_id, str(row["Student_Name"]), str(row["FeePayer"]), str(row["Status"]), str(row["Parent_Name"]), str(row["Parent_Contact"]), parent_contact2))
                        c.execute("""INSERT INTO StudentTerm (StudentID, ClassID, TermID)
                                     VALUES (?, ?, ?)""", (student_id, class_id, term_id))
                        inserted += 1
                    conn.commit()
                    msg = f"Inserted {inserted} new students and enrollments."
                    if duplicates:
                        msg += f"\nDuplicate StudentIDs skipped: {', '.join(duplicates)}"
                    if duplicate_terms:
                        msg += f"\nDuplicate Term-Class enrollments skipped: {', '.join(duplicate_terms)}"
                    QMessageBox.information(page, "Import Complete", msg)
                else:  # Enroll Existing Students Only
                    if "StudentID" not in df.columns:
                        QMessageBox.warning(page, "Error", "Excel must contain StudentID column")
                        return False
                    invalid_ids = []
                    duplicate_terms = []
                    inserted = 0
                    for _, row in df.iterrows():
                        student_id = str(row["StudentID"])
                        c.execute("SELECT COUNT(*) FROM Students WHERE StudentID = ?", (student_id,))
                        if c.fetchone()[0] == 0:
                            invalid_ids.append(student_id)
                            continue
                        c.execute("SELECT COUNT(*) FROM StudentTerm WHERE StudentID = ? AND ClassID = ? AND TermID = ?", (student_id, class_id, term_id))
                        if c.fetchone()[0] > 0:
                            duplicate_terms.append(student_id)
                            continue
                        c.execute("""INSERT INTO StudentTerm (StudentID, ClassID, TermID)
                                     VALUES (?, ?, ?)""", (student_id, class_id, term_id))
                        inserted += 1
                    conn.commit()
                    msg = f"Inserted {inserted} StudentTerm enrollments."
                    if invalid_ids:
                        msg += f"\nInvalid StudentIDs (not in Students table): {', '.join(invalid_ids)}"
                    if duplicate_terms:
                        msg += f"\nDuplicate Term-Class enrollments skipped: {', '.join(duplicate_terms)}"
                    QMessageBox.information(page, "Import Complete", msg)
                return True
            except Exception as e:
                if 'conn' in locals():
                    conn.rollback()
                QMessageBox.warning(page, "Error", f"Import failed: {str(e)}")
                return False
            finally:
                if 'conn' in locals():
                    conn.close()
        page.isComplete = is_complete
        page.validatePage = validate_page
        return page
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
        self.StudentPage = Student()
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
        year_filter = QComboBox()
        year_filter.addItem("All Years")
        c.execute("SELECT * FROM AcademicYear")
        years = c.fetchall()
        for year in years:
            year_filter.addItem(year[1], year[0])
        year_filter.setCurrentIndex(0)
        
        year_filter.setMaximumWidth(150)
        filter_layout.addWidget(self.Search_Input)
        filter_layout.addWidget(year_filter)
        filter_layout.addStretch()
        results_label = QLabel("0 results")
        results_label.setStyleSheet("font-size: 14px; color: #718096;")
        filter_layout.addWidget(results_label)
        layout.addLayout(filter_layout)

        # Connect search input and year filter to reload data
        self.Search_Input.textChanged.connect(self.Load_Data)
        year_filter.currentIndexChanged.connect(self.Load_Data)

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
        self.table.setRowCount(0)
        # Get selected year filter
        year_filter = self.layout().itemAt(1).layout().itemAt(1).widget()
        selected_year_id = year_filter.currentData()
        query = """
            SELECT AcademicYear.Year, Term.TermName, Term.StartDate, Term.EndDate
            FROM Term
            JOIN AcademicYear ON Term.AcademicYear = AcademicYear.AcademicYearID
        """
        # Add search filter
        search_text = self.Search_Input.text().strip()
        params = ()
        if search_text:
            if "WHERE" not in query:
                query += " WHERE"
            else:
                query += " AND"
            query += " (AcademicYear.Year LIKE ? OR Term.TermName LIKE ?)"
            params += (f"%{search_text}%", f"%{search_text}%")
        c.execute(query, params)
        results = c.fetchall()

        if selected_year_id and selected_year_id != "All Years":
            # Check if year exists and has no terms
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

        # self.table.resizeColumnsToContents()
        # Ensure header stretches after data reload
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.update_results_count()

        # Connect filter to reload data
        year_filter.currentIndexChanged.connect(self.Load_Data)

    def update_results_count(self):
        count = self.table.rowCount()
        self.layout().itemAt(1).layout().itemAt(3).widget().setText(f"{count} results")
    
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

            
            c.execute("SELECT 1 FROM AcademicYear WHERE Year = ?", (year_name,))
            if c.fetchone():
                error_label.setText("This academic year already exists.")
                return
            try:
                c.execute("INSERT INTO AcademicYear (Year) VALUES (?)", (year_name,))
                conn.commit()
                dialog.accept()
                # Optionally, refresh the table or notify user
                QMessageBox.information(self, "Success", f"Academic Year '{year_name}' added successfully.")
            except Exception as e:
                error_label.setText("Error saving year.")

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
        c.execute("SELECT AcademicYearID, Year FROM AcademicYear")
        years = c.fetchall()
        for year in years:
            year_combo.addItem(year[1], year[0])
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

            # Check for duplicate term in the same year
            c.execute("SELECT 1 FROM Term WHERE TermName=? AND AcademicYear=?",
                    (term, year_id))
            if c.fetchone():
                error_label.setText("This term already exists for the selected year.")
                return

            try:
                c.execute(
                    "INSERT INTO Term (TermName, AcademicYear, StartDate, EndDate) VALUES (?, ?, ?, ?)",
                    (term, year_id, start, end)
                )
                conn.commit()
                dialog.accept()
                QMessageBox.information(self, "Success", f"{term} added successfully.")
                # Optionally, refresh the table or notify user
            except Exception as e:
                error_label.setText("Error saving term.")

        save_btn.clicked.connect(save_term)
        cancel_btn.clicked.connect(dialog.reject)

        dialog.exec_()
    
    def editTerm(self, row):
        # Get current term data from the table
        year = self.table.item(row, 0).text()
        term = self.table.item(row, 1).text()
        start = self.table.item(row, 2).text()
        end = self.table.item(row, 3).text()

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

            # Check for duplicate term in the same year (excluding current term)
            c.execute("""
                SELECT 1 FROM Term 
                WHERE TermName=? AND AcademicYear=? AND TermID<>?
            """, (new_term, new_year_id, term_id))
            if c.fetchone():
                error_label.setText("This term already exists for the selected year.")
                return

            try:
                c.execute("""
                    UPDATE Term SET TermName=?, AcademicYear=?, StartDate=?, EndDate=?
                    WHERE TermID=?
                """, (new_term, new_year_id, new_start, new_end, term_id))
                conn.commit()
                dialog.accept()
                QMessageBox.information(self, "Success", f"{new_term} updated successfully.")
                self.Load_Data()
            except Exception as e:
                error_label.setText("Error updating term.")

        save_btn.clicked.connect(save_edit)
        cancel_btn.clicked.connect(dialog.reject)

        dialog.exec_()
    
    def deleteTerm(self, row):
        # Get current term data from the table
        year = self.table.item(row, 0).text()
        term = self.table.item(row, 1).text()
        start = self.table.item(row, 2).text()
        end = self.table.item(row, 3).text()

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

        # Check for Fees relation
        c.execute("SELECT 1 FROM Fees WHERE TermID = ?", (term_id,))
        if c.fetchone():
            QMessageBox.critical(self, "Restricted", "This term cannot be deleted because it is linked to one or more fees records.")
            return

        # Set Student.TermID to NULL for students assigned to this term
        c.execute("UPDATE Student SET TermID = NULL WHERE TermID = ?", (term_id,))
        # Now delete the term
        try:
            c.execute("DELETE FROM Term WHERE TermID = ?", (term_id,))
            conn.commit()
            QMessageBox.information(self, "Deleted", f"Term '{term}' for academic year '{year}' has been deleted.\n\n"
                                                     "Any students previously assigned to this term now have no term assigned.")
            self.Load_Data()
        except Exception as e:
            QMessageBox.critical(self, "Error", "An error occurred while deleting the term.")
    
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
        c.execute("SELECT AcademicYearID, Year FROM AcademicYear")
        years = c.fetchall()
        year_combo.addItem("Select Year", None)
        for y in years:
            year_combo.addItem(y[1], y[0])
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
            c.execute("SELECT 1 FROM AcademicYear WHERE Year = ? AND AcademicYearID != ?", (new_year_name, year_id))
            if c.fetchone():
                error_label.setText("This academic year already exists.")
                return
            try:
                c.execute("UPDATE AcademicYear SET Year = ? WHERE AcademicYearID = ?", (new_year_name, year_id))
                conn.commit()
                dialog.accept()
                QMessageBox.information(self, "Success", f"Academic Year updated successfully.")
                self.Load_Data()
            except Exception as e:
                error_label.setText("Error updating year.")

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
        c.execute("SELECT AcademicYearID, Year FROM AcademicYear")
        years = c.fetchall()
        year_combo.addItem("Select Year", None)
        for y in years:
            year_combo.addItem(y[1], y[0])
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

            try:
                c.execute("DELETE FROM AcademicYear WHERE AcademicYearID = ?", (year_id,))
                conn.commit()
                dialog.accept()
                QMessageBox.information(self, "Deleted", f"Academic Year '{year_name}' has been deleted successfully.")
                self.Load_Data()
            except Exception as e:
                error_label.setText("Error deleting year.")

        delete_btn.clicked.connect(do_delete)
        cancel_btn.clicked.connect(dialog.reject)
        dialog.exec_()

    def handle_year_action(self, index):
        action = self.year_action_combo.itemData(index)
        if action == "add":
            self.addYear()
            self.Load_Data()
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
            c.execute("DELETE FROM Class WHERE ClassID = ?", (class_id,))
            conn.commit()
            QMessageBox.information(self, "Deleted", f"Class '{class_name}' has been deleted successfully.")
            self.Load_Data()
        except Exception as e:
            QMessageBox.critical(self, "Error", "An error occurred while deleting the class.")

class Fees(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addStretch()
        label = QLabel("Fees Page")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2D3748;")
        layout.addWidget(label)
        layout.addStretch()

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
    window.show()
    sys.exit(app.exec_())