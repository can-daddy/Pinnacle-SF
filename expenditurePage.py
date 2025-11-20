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



class Expenditure(QTabWidget):
        def __init__(self):
         super().__init__()
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

class Supplies(QWidget):
     def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)
        self.header = QHBoxLayout()
        self.Heading = QLabel("Supplies")
        self.searchbar = QLineEdit()
        self.searchbar.setPlaceholderText("Search Supplies...")
        self.addbtn = QPushButton("Add Supply")
        self.addbtn.clicked.connect(self.add_supply)
        self.header.addWidget(self.Heading)
        self.header.addStretch()
        self.header.addWidget(self.searchbar)
        self.header.addWidget(self.addbtn)
        layout.addLayout(self.header)
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Name","Subcategory","Quantity","Unit","Actions"])
        layout.addWidget(self.table)

     def add_supply(self):
         dialog = QDialog()
         dialog.setWindowTitle("Add New Supply")
         dialogLayout = QVBoxLayout()
         FormLayout = QFormLayout()
         nameinput = QLineEdit()
         selectcategory = QComboBox()
         selectsubcategory = QComboBox()
         UnitInput = QLineEdit()
         savebtn = QPushButton("Save")
         FormLayout.addRow("Name:",nameinput)
         FormLayout.addRow("Category:",selectcategory)
         FormLayout.addRow("Subcategory:",selectsubcategory)
         FormLayout.addRow("Unit:",UnitInput)
         self.heading = QLabel("Add New Supply")
         dialogLayout.addWidget(self.heading)
         dialogLayout.addLayout(FormLayout)
         dialogLayout.addWidget(savebtn)
         dialog.setLayout(dialogLayout)
         dialog.exec_()

class SuppliesInventory(QWidget):
     def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)
        self.header = QHBoxLayout()
        self.Heading = QLabel("Supplies Inventory")
        self.searchbar = QLineEdit()
        self.searchbar.setPlaceholderText("Search Supplies Inventory...")
        self.header.addWidget(self.Heading)
        self.header.addStretch()
        self.header.addWidget(self.searchbar)
        layout.addLayout(self.header)
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["Name","Type","Quantity","Price","Date Added","Reference","Actions"])
        layout.addWidget(self.table)

class Expense(QWidget):
     def __init__(self):
        super().__init__()
        layout = QVBoxLayout()
        self.setLayout(layout)
        self.header = QHBoxLayout()
        self.Heading = QLabel("Expenses")
        self.searchbar = QLineEdit()
        self.searchbar.setPlaceholderText("Search Expenses...")
        self.addbtn = QPushButton("Add Expense")
        self.addbtn.clicked.connect(self.add_expense)
        self.header.addWidget(self.Heading)
        self.header.addStretch()
        self.header.addWidget(self.searchbar)
        self.header.addWidget(self.addbtn)
        layout.addLayout(self.header)
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["Category","Subcategory","Amount","Date","Reference","Actions"])
        layout.addWidget(self.table)

     def add_expense(self):
         dialog = QDialog()
         dialog.setWindowTitle("Add New Expense")
         dialogLayout = QVBoxLayout()
         FormLayout = QFormLayout()
         selectcategory = QComboBox()
         selectSubcategory = QComboBox()
         amountinput = QLineEdit()
         dateinput = QDateEdit()
         referenceinput = QLineEdit()
         savebtn = QPushButton("Save")
         FormLayout.addRow("Select Category:",selectcategory)
         FormLayout.addRow("Select Subcategory:",selectSubcategory)
         FormLayout.addRow("Amount:",amountinput)
         FormLayout.addRow("Date:",dateinput)
         FormLayout.addRow("Reference:",referenceinput)
         self.heading = QLabel("Add New Expense")
         dialogLayout.addWidget(self.heading)
         dialogLayout.addLayout(FormLayout)
         dialogLayout.addWidget(savebtn)
         dialog.setLayout(dialogLayout)
         dialog.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Expenditure()
    window.setWindowTitle("Expenditure Management")
    window.resize(800,600)
    window.show()
    sys.exit(app.exec_())