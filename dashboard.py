import sys
import sqlite3
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
                             QFrame, QGridLayout, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QScrollArea, QSizePolicy, QPushButton)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QColor
from contextlib import contextmanager

import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from finance_service import FinanceService
from financial_report_dialog import FinancialReportDialog

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
                self.term_selector.addItem(term['TermName'], term['TermID'])
        except Exception as e:
            print(f"Error loading terms: {e}")

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
            print(f"Error updating dashboard: {e}")

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
            print(f"Chart Error: {e}")
            
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
            print(f"Table Error: {e}")

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
            print(f"Supplies Error: {e}")
