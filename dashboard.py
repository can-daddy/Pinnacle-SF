import sys
import sqlite3
from datetime import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
                             QFrame, QGridLayout, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QScrollArea, QSizePolicy)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QColor
from contextlib import contextmanager

import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

# Database connection context manager (duplicated to avoid circular imports)
@contextmanager
def get_db_connection():
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

class Dashboard(QWidget):
    def __init__(self):
        super().__init__()
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
        
        # Net Balance (Highlighted)
        self.net_balance_label = QLabel("Net Balance: $0.00")
        self.net_balance_label.setStyleSheet("""
            background-color: #48BB78; 
            color: white; 
            font-size: 16px; 
            font-weight: bold; 
            padding: 8px 16px; 
            border-radius: 5px;
        """)
        
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(QLabel("Term:"))
        header_layout.addWidget(self.term_selector)
        header_layout.addWidget(date_label)
        header_layout.addWidget(self.net_balance_label)
        
        self.content_layout.addLayout(header_layout)

        # --- Summary Cards Section ---
        cards_layout = QHBoxLayout()
        cards_layout.setSpacing(15)
        
        self.card_fees_collected = self.create_card("Fees Collected", "$0.00", "#3182CE") # Blue
        self.card_outstanding = self.create_card("Outstanding Fees", "$0.00", "#DD6B20") # Orange
        self.card_expenses = self.create_card("Total Expenses", "$0.00", "#E53E3E") # Red
        self.card_supplies = self.create_card("Supplies Cost", "$0.00", "#805AD5") # Purple
        
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

    def load_terms(self):
        self.term_selector.clear()
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT TermID, TermName FROM Term ORDER BY TermID DESC")
                terms = c.fetchall()
                for term in terms:
                    self.term_selector.addItem(term['TermName'], term['TermID'])
        except Exception as e:
            print(f"Error loading terms: {e}")

    def update_dashboard(self):
        term_id = self.term_selector.currentData()
        if not term_id:
            return

        try:
            with get_db_connection() as (conn, c):
                # Get Term Dates
                c.execute("SELECT StartDate, EndDate FROM Term WHERE TermID = ?", (term_id,))
                term_dates = c.fetchone()
                if not term_dates:
                    return
                start_date = term_dates['StartDate']
                end_date = term_dates['EndDate']

                # --- 1. Fees Data ---
                # Total Fees Expected for this term
                # Logic: Sum of (Fee Amount for Class/Term) for all students enrolled in that class/term
                # Adjusted by PayRate (Scholarship)
                c.execute("""
                    SELECT SUM(
                        (f.Amount * (st.PayRate / 100.0))
                    ) as TotalExpected
                    FROM StudentTerm st
                    JOIN Fees f ON st.ClassID = f.Class AND st.TermID = f.Term
                    WHERE st.TermID = ?
                """, (term_id,))
                total_expected = c.fetchone()['TotalExpected'] or 0.0

                # Total Fees Collected for this term
                c.execute("""
                    SELECT SUM(AmountPaid) 
                    FROM FeesPayment fp
                    JOIN StudentTerm st ON fp.StudentTermID = st.StudentTermID
                    WHERE st.TermID = ?
                """, (term_id,))
                total_collected = c.fetchone()[0] or 0.0
                
                outstanding = max(0, total_expected - total_collected)
                
                # Update Cards
                self.card_fees_collected.value_label.setText(f"${total_collected:,.2f}")
                self.card_outstanding.value_label.setText(f"${outstanding:,.2f}")

                # --- 2. Expenses Data ---
                # Filter by Term Dates
                c.execute("""
                    SELECT SUM(ExpAmount) 
                    FROM Expenses 
                    WHERE ExpDate BETWEEN ? AND ?
                """, (start_date, end_date))
                total_expenses = c.fetchone()[0] or 0.0
                self.card_expenses.value_label.setText(f"${total_expenses:,.2f}")

                # --- 3. Supplies Data ---
                # Filter by Term Dates (SuppliesInventory 'In' type usually means cost, but 'Out' might be usage. 
                # Assuming 'In' is cost incurred to buy supplies. 'Out' is internal usage, not financial cost?)
                # Actually, usually 'In' means we bought it (Expense).
                c.execute("""
                    SELECT SUM(Price) 
                    FROM SuppliesInventory 
                    WHERE Type = 'In' AND Date BETWEEN ? AND ?
                """, (start_date, end_date))
                total_supplies = c.fetchone()[0] or 0.0
                self.card_supplies.value_label.setText(f"${total_supplies:,.2f}")

                # --- 4. Net Balance ---
                net_balance = total_collected - (total_expenses + total_supplies)
                self.net_balance_label.setText(f"Net Balance: ${net_balance:,.2f}")
                if net_balance >= 0:
                    self.net_balance_label.setStyleSheet("background-color: #48BB78; color: white; font-size: 16px; font-weight: bold; padding: 8px 16px; border-radius: 5px;")
                else:
                    self.net_balance_label.setStyleSheet("background-color: #E53E3E; color: white; font-size: 16px; font-weight: bold; padding: 8px 16px; border-radius: 5px;")

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
            with get_db_connection() as (conn, c):
                c.execute("""
                    SELECT c.CatName, SUM(e.ExpAmount)
                    FROM Expenses e
                    JOIN SubCategory sc ON e.ExpSubCategory = sc.SubCatID
                    JOIN Category c ON sc.CategoryID = c.CatID
                    WHERE e.ExpDate BETWEEN ? AND ?
                    GROUP BY c.CatName
                    ORDER BY SUM(e.ExpAmount) DESC
                    LIMIT 5
                """, (start_date, end_date))
                data = c.fetchall()
                
                if data:
                    categories = [row[0] for row in data]
                    amounts = [row[1] for row in data]
                    ax.bar(categories, amounts, color='#4299E1')
                    ax.set_ylabel('Amount ($)')
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
            with get_db_connection() as (conn, c):
                # Calculate outstanding per class
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
                    LIMIT 10
                """, (term_id, term_id))
                
                rows = c.fetchall()
                for row_idx, row in enumerate(rows):
                    class_name = row[0]
                    expected = row[1] or 0
                    collected = row[2] or 0
                    balance = expected - collected
                    
                    if balance > 0:
                        self.outstanding_table.insertRow(row_idx)
                        self.outstanding_table.setItem(row_idx, 0, QTableWidgetItem(class_name))
                        self.outstanding_table.setItem(row_idx, 1, QTableWidgetItem(f"${balance:,.2f}"))
                        
                        percent = (balance / expected * 100) if expected > 0 else 0
                        self.outstanding_table.setItem(row_idx, 2, QTableWidgetItem(f"{percent:.1f}%"))
        except Exception as e:
            print(f"Table Error: {e}")

    def update_supplies_table(self, start_date, end_date):
        self.supplies_table.setRowCount(0)
        try:
            with get_db_connection() as (conn, c):
                c.execute("""
                    SELECT s.Name, SUM(si.Quantity), SUM(si.Price)
                    FROM SuppliesInventory si
                    JOIN Supplies s ON si.Supplies_ID = s.SuppliesID
                    WHERE si.Type = 'In' AND si.Date BETWEEN ? AND ?
                    GROUP BY s.Name
                    ORDER BY SUM(si.Price) DESC
                    LIMIT 5
                """, (start_date, end_date))
                
                rows = c.fetchall()
                for row_idx, row in enumerate(rows):
                    self.supplies_table.insertRow(row_idx)
                    self.supplies_table.setItem(row_idx, 0, QTableWidgetItem(str(row[0])))
                    self.supplies_table.setItem(row_idx, 1, QTableWidgetItem(str(row[1])))
                    self.supplies_table.setItem(row_idx, 2, QTableWidgetItem(f"${row[2]:,.2f}"))
        except Exception as e:
            print(f"Supplies Error: {e}")
