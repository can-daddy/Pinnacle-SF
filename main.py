import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sqlite3

# Connect to the SQLite database
conn = sqlite3.connect("Database.db")
c = conn.cursor()

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
        add_year_btn = QPushButton("Add Year")
        add_year_btn.clicked.connect(self.addYear)
        add_term_btn = QPushButton("Add Term")
        header_layout.addWidget(add_year_btn)
        header_layout.addWidget(add_term_btn)
        layout.addLayout(header_layout)

        # Filter row
        filter_layout = QHBoxLayout()
        search_input = QLineEdit()
        search_input.setPlaceholderText("Search...")
        search_input.setMaximumWidth(200)
        year_filter = QComboBox()
        year_filter.addItem("All Years")
        c.exexute("SELECT * FROM AcademicYear")
        year_filter.setMaximumWidth(150)
        filter_layout.addWidget(search_input)
        filter_layout.addWidget(year_filter)
        filter_layout.addStretch()
        results_label = QLabel("0 results")
        results_label.setStyleSheet("font-size: 14px; color: #718096;")
        filter_layout.addWidget(results_label)
        layout.addLayout(filter_layout)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Academic Year", "Term", "Starting Date", "Ending Date", "Action"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)

        # Add sample data
        self.add_sample_data()
        self.update_results_count()
   
    def add_sample_data(self):
        sample_data = [
            ("2023-2024", "Term 1", "2023-09-01", "2023-12-31"),
            ("2023-2024", "Term 2", "2024-01-01", "2024-04-30"),
            ("2024-2025", "Term 1", "2024-09-01", "2024-12-31"),
            ("2024-2025", "Term 2", "2025-01-01", "2025-04-30"),
            ("2025-2026", "Term 1", "2025-09-01", "2025-12-31"),
        ]
        for row_data in sample_data:
            row = self.table.rowCount()
            self.table.insertRow(row)
            self.table.setRowHeight(row, 50)

            for col, data in enumerate(row_data):
                self.table.setItem(row, col, QTableWidgetItem(data))
            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(5, 0, 5, 0)
            action_layout.setSpacing(5)
            edit_btn = QPushButton("Edit")
            delete_btn = QPushButton("Delete")
            delete_btn.setObjectName("delete")
            action_layout.addWidget(edit_btn)
            action_layout.addWidget(delete_btn)
            action_layout.addStretch()
            self.table.setCellWidget(row, 4, action_widget)

        self.table.resizeColumnsToContents()

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
        # layout.setSpacing(15)

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
                conn.close()
                return
            try:
                c.execute("INSERT INTO AcademicYear (Year) VALUES (?)", (year_name,))
                conn.commit()
                conn.close()
                dialog.accept()
                # Optionally, refresh the table or notify user
                QMessageBox.information(self, "Success", f"Academic Year '{year_name}' added successfully.")
            except Exception as e:
                error_label.setText("Error saving year.")
                conn.close()

        save_btn.clicked.connect(save_year)
        cancel_btn.clicked.connect(dialog.reject)

        dialog.exec_()

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

class Student(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addStretch()
        label = QLabel("Student Page")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2D3748;")
        layout.addWidget(label)
        layout.addStretch()

class Class(QWidget):
    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.addStretch()
        label = QLabel("Class Page")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 28px; font-weight: bold; color: #2D3748;")
        layout.addWidget(label)
        layout.addStretch()

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