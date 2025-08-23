import sys
import sqlite3
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QScrollArea, QLineEdit, QComboBox, QTableWidget, QTableWidgetItem, QDialog, QFormLayout, QMessageBox, QWizard, QWizardPage, QFileDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import pandas as pd

class StudentWidget(QWidget):
    def __init__(self, db_path):
        super().__init__()
        self.db_path = db_path
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
        search_input = QLineEdit()
        search_input.setPlaceholderText("Search...")
        search_input.setMaximumWidth(200)
        class_filter = QComboBox()
        class_filter.addItem("All Classes")
        term_filter = QComboBox()
        term_filter.addItem("All Terms")
        filter_layout.addWidget(search_input)
        filter_layout.addWidget(class_filter)
        filter_layout.addWidget(term_filter)
        filter_layout.addStretch()
        self.results_label = QLabel("0 results")
        self.results_label.setStyleSheet("font-size: 14px; color: #718096;")
        filter_layout.addWidget(self.results_label)
        layout.addLayout(filter_layout)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels(["StudentID", "Name", "FeePayer", "Status", "Parent Name", "Contact", "Contact2", "Class", "Term"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)
        self.load_data()

    def load_data(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("""SELECT s.StudentID, s.Student_Name, s.FeePayer, s.Status, s.Parent_Name, s.Parent_Contact, s.Parent_Contact2,
                     cl.ClassName, t.Term
                     FROM Students s
                     LEFT JOIN StudentTerm st ON s.StudentID = st.StudentID
                     LEFT JOIN Class cl ON st.ClassID = cl.ClassID
                     LEFT JOIN Term t ON st.TermID = t.TermID""")
        rows = c.fetchall()
        self.table.setRowCount(0)
        for row_data in rows:
            row = self.table.rowCount()
            self.table.insertRow(row)
            for col, data in enumerate(row_data):
                self.table.setItem(row, col, QTableWidgetItem(str(data) if data else ""))
        self.table.resizeColumnsToContents()
        count = self.table.rowCount()
        self.results_label.setText(f"{count} results")
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
        status_combo.addItems(["Active", "Inactive"])
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

    def populate_combos(self, class_combo, term_combo):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT ClassID, ClassName FROM Class")
        classes = c.fetchall()
        for cl in classes:
            class_combo.addItem(cl[1], cl[0])
        c.execute("SELECT TermID, Term FROM Term")
        terms = c.fetchall()
        for t in terms:
            term_combo.addItem(t[1], t[0])
        conn.close()

    def save_student(self, dialog, id_edit, name_edit, feepayer_edit, status_combo, parent_name_edit, contact_edit, contact2_edit, class_combo, term_combo):
        student_id_str = id_edit.text().strip()
        if not student_id_str.isdigit():
            QMessageBox.warning(dialog, "Error", "Student ID must be an integer")
            return
        student_id = int(student_id_str)
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
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM Students WHERE StudentID = ?", (student_id,))
        if c.fetchone()[0] > 0:
            QMessageBox.warning(dialog, "Error", "Duplicate Student ID")
            conn.close()
            return
        try:
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
            QMessageBox.warning(dialog, "Error", str(e))
        finally:
            conn.close()

    def show_import_excel_wizard(self):
        wizard = QWizard(self)
        wizard.setWindowTitle("Import Students from Excel")
        page1 = self.create_select_class_term_page()
        page2 = self.create_select_file_page(wizard)
        wizard.addPage(page1)
        wizard.addPage(page2)
        if wizard.exec_():
            self.load_data()

    def create_select_class_term_page(self):
        page = QWizardPage()
        layout = QVBoxLayout()
        mode_combo = QComboBox()
        mode_combo.addItems(["Add New Students", "Add StudentTerm Only"])
        class_combo = QComboBox()
        term_combo = QComboBox()
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
        return page

    def create_select_file_page(self, wizard):
        page = QWizardPage()
        layout = QVBoxLayout()
        file_label = QLabel("No file selected")
        select_btn = QPushButton("Select Excel File")
        layout.addWidget(select_btn)
        layout.addWidget(file_label)
        page.setLayout(layout)
        page.file_path = None
        def select_file():
            file_path, _ = QFileDialog.getOpenFileName(page, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
            if file_path:
                page.file_path = file_path
                file_label.setText(file_path)
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
                conn = sqlite3.connect(self.db_path)
                c = conn.cursor()
                if mode == "Add New Students":
                    required_cols = ["StudentID", "Student_Name", "FeePayer", "Status", "Parent_Name", "Parent_Contact"]
                    if not all(col in df.columns for col in required_cols):
                        QMessageBox.warning(page, "Error", "Excel missing required columns: StudentID, Student_Name, FeePayer, Status, Parent_Name, Parent_Contact")
                        return False
                    duplicates = []
                    inserted = 0
                    for _, row in df.iterrows():
                        student_id = int(row["StudentID"])
                        c.execute("SELECT COUNT(*) FROM Students WHERE StudentID = ?", (student_id,))
                        if c.fetchone()[0] > 0:
                            duplicates.append(student_id)
                            continue
                        parent_contact2 = row.get("Parent_Contact2", "")
                        c.execute("""INSERT INTO Students (StudentID, Student_Name, FeePayer, Status, Parent_Name, Parent_Contact, Parent_Contact2)
                                     VALUES (?, ?, ?, ?, ?, ?, ?)""", (student_id, row["Student_Name"], row["FeePayer"], row["Status"], row["Parent_Name"], row["Parent_Contact"], parent_contact2))
                        c.execute("""INSERT INTO StudentTerm (StudentID, ClassID, TermID)
                                     VALUES (?, ?, ?)""", (student_id, class_id, term_id))
                        inserted += 1
                    conn.commit()
                    msg = f"Inserted {inserted} students."
                    if duplicates:
                        msg += f"\nDuplicates skipped: {', '.join(map(str, duplicates))}"
                    QMessageBox.information(page, "Import Complete", msg)
                else:  # Add StudentTerm Only
                    if "StudentID" not in df.columns:
                        QMessageBox.warning(page, "Error", "Excel must contain StudentID column")
                        return False
                    invalid_ids = []
                    inserted = 0
                    for _, row in df.iterrows():
                        student_id = int(row["StudentID"])
                        c.execute("SELECT COUNT(*) FROM Students WHERE StudentID = ?", (student_id,))
                        if c.fetchone()[0] == 0:
                            invalid_ids.append(student_id)
                            continue
                        c.execute("SELECT COUNT(*) FROM StudentTerm WHERE StudentID = ? AND ClassID = ? AND TermID = ?", (student_id, class_id, term_id))
                        if c.fetchone()[0] > 0:
                            continue  # Skip duplicate StudentTerm entries
                        c.execute("""INSERT INTO StudentTerm (StudentID, ClassID, TermID)
                                     VALUES (?, ?, ?)""", (student_id, class_id, term_id))
                        inserted += 1
                    conn.commit()
                    msg = f"Inserted {inserted} StudentTerm entries."
                    if invalid_ids:
                        msg += f"\nInvalid StudentIDs (not in Students table): {', '.join(map(str, invalid_ids))}"
                    QMessageBox.information(page, "Import Complete", msg)
                return True
            except Exception as e:
                if 'conn' in locals():
                    conn.rollback()
                QMessageBox.warning(page, "Error", str(e))
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
        self.db_path = "school.db"
        self.init_db()
        
        # Main widget and layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # Sidebar
        self.sidebar = QWidget()
        self.sidebar.setFixedWidth(200)
        self.sidebar_layout = QVBoxLayout(self.sidebar)
        self.sidebar_layout.setContentsMargins(10, 10, 10, 10)
        self.sidebar_layout.setSpacing(10)
        
        # Toggle button
        self.toggle_btn = QPushButton(QIcon("menu-icon.png"), "")
        self.toggle_btn.setFixedSize(40, 40)
        self.toggle_btn.clicked.connect(self.toggle_sidebar)
        self.sidebar_layout.addWidget(self.toggle_btn)
        
        # Sidebar menu items
        self.menu_items = [
            ("Fees", "fees-icon.png"),
            ("Expenses", "expenses-icon.png"),
            ("Uniforms", "uniform-icon.png"),
            ("Reports", "reports-icon.png")
        ]
        
        self.menu_buttons = []
        for name, icon in self.menu_items:
            btn = QPushButton(QIcon(icon), name)
            btn.setFixedHeight(40)
            self.menu_buttons.append(btn)
            self.sidebar_layout.addWidget(btn)
        
        self.sidebar_layout.addStretch()
        
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
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.body = QWidget()
        self.body_layout = QVBoxLayout(self.body)
        self.body_layout.setContentsMargins(0, 0, 0, 0)
        
        # Add StudentWidget to body
        self.student_widget = StudentWidget(self.db_path)
        self.body_layout.addWidget(self.student_widget)
        
        self.scroll_area.setWidget(self.body)
        self.content_layout.addWidget(self.scroll_area)
        
        # Add sidebar and content to main layout
        self.main_layout.addWidget(self.sidebar)
        self.main_layout.addWidget(self.content)
        
        # Sidebar state
        self.sidebar_expanded = True
        self.apply_styles()
        
    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("""CREATE TABLE IF NOT EXISTS Students (
            StudentID INTEGER NOT NULL PRIMARY KEY,
            Student_Name TEXT NOT NULL,
            FeePayer TEXT NOT NULL,
            Status TEXT NOT NULL,
            Parent_Name TEXT NOT NULL,
            Parent_Contact TEXT NOT NULL,
            Parent_Contact2 TEXT
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS StudentTerm (
            StudentTermID INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
            StudentID INTEGER NOT NULL,
            ClassID INTEGER NOT NULL,
            TermID INTEGER NOT NULL,
            FOREIGN KEY(StudentID) REFERENCES Students(StudentID) ON DELETE CASCADE,
            FOREIGN KEY(ClassID) REFERENCES Class(ClassID) ON DELETE RESTRICT,
            FOREIGN KEY(TermID) REFERENCES Term(TermID) ON DELETE SET NULL
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS Class (
            ClassID INTEGER PRIMARY KEY AUTOINCREMENT,
            ClassName TEXT NOT NULL
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS AcademicYear (
            AcademicYearID INTEGER PRIMARY KEY AUTOINCREMENT,
            AcademicYear TEXT NOT NULL
        )""")
        c.execute("""CREATE TABLE IF NOT EXISTS Term (
            TermID INTEGER PRIMARY KEY AUTOINCREMENT,
            Term TEXT NOT NULL,
            StartingDate DATE,
            EndingDate DATE,
            AcademicYearID INTEGER,
            FOREIGN KEY(AcademicYearID) REFERENCES AcademicYear(AcademicYearID)
        )""")
        c.execute("SELECT COUNT(*) FROM Class")
        if c.fetchone()[0] == 0:
            classes = ["Class 1A", "Class 2B", "Class 3C"]
            for cl in classes:
                c.execute("INSERT INTO Class (ClassName) VALUES (?)", (cl,))
        c.execute("SELECT COUNT(*) FROM AcademicYear")
        if c.fetchone()[0] == 0:
            years = [
                ("2023-2024", [("Term 1", "2023-09-01", "2023-12-31"), ("Term 2", "2024-01-01", "2024-04-30")]),
                ("2024-2025", [("Term 1", "2024-09-01", "2024-12-31"), ("Term 2", "2025-01-01", "2025-04-30")]),
                ("2025-2026", [("Term 1", "2025-09-01", "2025-12-31"), ("Term 2", "2026-01-01", "2026-04-30")])
            ]
            for year, terms in years:
                c.execute("INSERT INTO AcademicYear (AcademicYear) VALUES (?)", (year,))
                year_id = c.lastrowid
                for t in terms:
                    c.execute("INSERT INTO Term (Term, StartingDate, EndingDate, AcademicYearID) VALUES (?, ?, ?, ?)", (t[0], t[1], t[2], year_id))
        conn.commit()
        conn.close()

    def toggle_sidebar(self):
        self.sidebar_expanded = not self.sidebar_expanded
        new_width = 60 if not self.sidebar_expanded else 200
        self.sidebar.setFixedWidth(new_width)
        
        for btn in self.menu_buttons:
            if self.sidebar_expanded:
                btn.setText(self.menu_items[self.menu_buttons.index(btn)][0])
            else:
                btn.setText("")
    
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
                background-color: #E2E8F0;
                border-radius: 20px;
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
                background-color: transparent;
                border: none;
            }
            QScrollArea QWidget {
                background-color: transparent;
            }
        """)
        self.sidebar.setObjectName("sidebar")
        self.toggle_btn.setObjectName("toggle_btn")
        self.page_title.setObjectName("page_title")
        self.school_name.setObjectName("school_name")
        self.separator.setObjectName("separator")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())