"""
School Finance Management System - Report Generator Module
Handles comprehensive financial reporting with filtering, visualization, and export capabilities.
"""

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QComboBox, QPushButton, QTableWidget, QTableWidgetItem,
                             QTabWidget, QGroupBox, QGridLayout, QHeaderView,
                             QMessageBox, QFileDialog, QScrollArea, QFrame)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QColor,QPainter
from PyQt5.QtChart import QChart, QChartView, QPieSeries, QBarSeries, QBarSet, QBarCategoryAxis, QValueAxis
import sqlite3
from datetime import datetime
from typing import Dict, List, Tuple, Optional
import csv


class FinancialReportCard(QFrame):
    """Widget for displaying summary financial metrics"""
    
    def __init__(self, title: str, value: str, subtitle: str = "", parent=None):
        super().__init__(parent)
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet("""
            FinancialReportCard {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        
        layout = QVBoxLayout(self)
        
        # Title
        title_label = QLabel(title)
        title_font = QFont()
        title_font.setPointSize(9)
        title_font.setBold(False)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #6c757d;")
        
        # Value
        value_label = QLabel(value)
        value_font = QFont()
        value_font.setPointSize(20)
        value_font.setBold(True)
        value_label.setFont(value_font)
        value_label.setStyleSheet("color: #212529;")
        
        layout.addWidget(title_label)
        layout.addWidget(value_label)
        
        if subtitle:
            subtitle_label = QLabel(subtitle)
            subtitle_font = QFont()
            subtitle_font.setPointSize(8)
            subtitle_label.setFont(subtitle_font)
            subtitle_label.setStyleSheet("color: #868e96;")
            layout.addWidget(subtitle_label)
        
        layout.addStretch()


class ReportDataManager:
    """Handles all database queries for report generation"""
    
    def __init__(self, db_path: str):
        self.db_path = db_path
    
    def get_connection(self) -> sqlite3.Connection:
        """Create and return database connection"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn
    
    def get_academic_years(self) -> List[Tuple[int, str]]:
        """Fetch all academic years"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY Year DESC")
        years = cursor.fetchall()
        conn.close()
        return [(row[0], row[1]) for row in years]
    
    def get_terms_by_year(self, year_id: int) -> List[Tuple[int, str]]:
        """Fetch terms for a specific academic year"""
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT TermID, TermName 
            FROM Term 
            WHERE AcademicYear = ? 
            ORDER BY StartDate
        """, (year_id,))
        terms = cursor.fetchall()
        conn.close()
        return [(row[0], row[1]) for row in terms]
    
    def get_fees_completion_summary(self, term_id: int) -> Dict:
        """
        Calculate overall fees completion for a term
        Returns: {expected: float, collected: float, percentage: float}
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        # Get expected fees for the term
        cursor.execute("""
            SELECT SUM(f.Amount * st.PayRate / 100.0) as ExpectedTotal
            FROM StudentTerm st
            JOIN Fees f ON f.Class = st.ClassID AND f.Term = ?
            WHERE st.TermID = ?
        """, (term_id, term_id))
        
        expected_result = cursor.fetchone()
        expected_total = expected_result[0] if expected_result[0] else 0.0
        
        # Get collected fees for the term
        cursor.execute("""
            SELECT SUM(fp.AmountPaid) as CollectedTotal
            FROM FeesPayment fp
            JOIN StudentTerm st ON fp.StudentTermID = st.StudentTermID
            WHERE st.TermID = ?
        """, (term_id,))
        
        collected_result = cursor.fetchone()
        collected_total = collected_result[0] if collected_result[0] else 0.0
        
        conn.close()
        
        percentage = (collected_total / expected_total * 100) if expected_total > 0 else 0
        
        return {
            'expected': expected_total,
            'collected': collected_total,
            'outstanding': expected_total - collected_total,
            'percentage': percentage
        }
    
    def get_class_fees_breakdown(self, term_id: int) -> List[Dict]:
        """
        Get detailed fees breakdown by class
        Returns list of dicts with class-wise completion data
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT 
                    c.ClassName,
                    COUNT(DISTINCT st.StudentID) as StudentCount,
                    SUM(f.Amount * st.PayRate / 100.0) as ExpectedAmount,
                    COALESCE(SUM(fp.AmountPaid), 0) as CollectedAmount
                FROM Class c
                JOIN StudentTerm st ON st.ClassID = c.ClassID
                JOIN Fees f ON f.Class = c.ClassID AND f.Term = ?
                LEFT JOIN (
                    SELECT StudentTermID, SUM(AmountPaid) AS AmountPaid
                    FROM FeesPayment
                    GROUP BY StudentTermID
                ) fp ON fp.StudentTermID = st.StudentTermID
                WHERE st.TermID = ?
                GROUP BY c.ClassID, c.ClassName
                ORDER BY c.ClassName;
        """, (term_id, term_id))
        
        results = []
        for row in cursor.fetchall():
            expected = row[2] if row[2] else 0.0
            collected = row[3] if row[3] else 0.0
            percentage = (collected / expected * 100) if expected > 0 else 0
            
            results.append({
                'class_name': row[0],
                'student_count': row[1],
                'expected': expected,
                'collected': collected,
                'outstanding': expected - collected,
                'percentage': percentage
            })
        
        conn.close()
        return results
    
    def get_assets_list(self) -> List[Dict]:
        """Get all assets with their current status"""
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT 
                a.AssetID,
                a.AssetName,
                a.AssetValue,
                a.Status,
                a.updatedON,
                e.ExpenseDate as PurchaseDate,
                esc.ExpenditureSubCategoryName as Category
            FROM Assets a
            LEFT JOIN Expenditure e ON a.ExpenseID = e.ExpenseID
            LEFT JOIN ExpenditureSubCategory esc ON e.ExpeSubID = esc.ExpeSubCatID
            ORDER BY a.AssetID DESC
        """)
        
        assets = []
        for row in cursor.fetchall():
            assets.append({
                'asset_id': row[0],
                'name': row[1],
                'value': row[2],
                'status': row[3] if row[3] else 'Active',
                'updated_on': row[4],
                'purchase_date': row[5],
                'category': row[6]
            })
        
        conn.close()
        return assets
    
    def get_expenditure_report(self, year_id: int, expense_type: str) -> List[Dict]:
        """
        Get expenditure report grouped by category
        expense_type: 'Capital' for assets, 'Operating' for expenses
        """
        conn = self.get_connection()
        cursor = conn.cursor()
        
        cursor.execute("""
            SELECT 
                ec.ExpCatName as CategoryName,
                esc.ExpenditureSubCategoryName as SubCategoryName,
                COUNT(e.ExpenseID) as TransactionCount,
                SUM(e.ExpenseAmount) as TotalAmount
            FROM ExpenditureCategory ec
            JOIN ExpenditureSubCategory esc ON esc.ExpenditureCategoryID = ec.ExpCatID
            JOIN Expenditure e ON e.ExpeSubID = esc.ExpeSubCatID
            JOIN Term t ON CAST(strftime('%Y', e.ExpenseDate) AS INTEGER) 
                          >= CAST(substr(t.StartDate, 1, 4) AS INTEGER)
                          AND CAST(strftime('%Y', e.ExpenseDate) AS INTEGER) 
                          <= CAST(substr(t.EndDate, 1, 4) AS INTEGER)
            WHERE ec.ExpCatType = ? AND t.AcademicYear = ?
            GROUP BY ec.ExpCatID, ec.ExpCatName, esc.ExpeSubCatID, esc.ExpenditureSubCategoryName
            ORDER BY ec.ExpCatName, esc.ExpenditureSubCategoryName
        """, (expense_type, year_id))
        
        results = []
        for row in cursor.fetchall():
            results.append({
                'category': row[0],
                'subcategory': row[1],
                'count': row[2],
                'total': row[3]
            })
        
        conn.close()
        return results


class ReportGeneratorWidget(QWidget):
    """Main Report Generator Interface"""
    
    def __init__(self, db_path: str, parent=None):
        super().__init__(parent)
        self.db_path = db_path
        self.data_manager = ReportDataManager(db_path)
        self.current_year_id = None
        self.current_term_id = None
        
        self.init_ui()
        self.load_academic_years()
    
    def init_ui(self):
        """Initialize the user interface"""
        main_layout = QVBoxLayout(self)
        
        # Header Section
        header = QLabel("Financial Reports")
        header_font = QFont()
        header_font.setPointSize(16)
        header_font.setBold(True)
        header.setFont(header_font)
        main_layout.addWidget(header)
        
        # Filters Section
        filter_group = QGroupBox("Report Filters")
        filter_layout = QHBoxLayout()
        
        # Academic Year Filter
        filter_layout.addWidget(QLabel("Academic Year:"))
        self.year_combo = QComboBox()
        self.year_combo.currentIndexChanged.connect(self.on_year_changed)
        filter_layout.addWidget(self.year_combo)
        
        # Term Filter
        filter_layout.addWidget(QLabel("Term:"))
        self.term_combo = QComboBox()
        self.term_combo.currentIndexChanged.connect(self.on_term_changed)
        filter_layout.addWidget(self.term_combo)
        
        # Generate Button
        self.generate_btn = QPushButton("Generate Reports")
        self.generate_btn.clicked.connect(self.generate_reports)
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #007bff;
                color: white;
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)
        filter_layout.addWidget(self.generate_btn)
        
        # Export Button
        self.export_btn = QPushButton("Export All Reports")
        self.export_btn.clicked.connect(self.export_reports)
        self.export_btn.setEnabled(False)
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #218838;
            }
            QPushButton:disabled {
                background-color: #6c757d;
            }
        """)
        filter_layout.addWidget(self.export_btn)
        
        filter_layout.addStretch()
        filter_group.setLayout(filter_layout)
        main_layout.addWidget(filter_group)
        
        # Tabs for different reports
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Tab 1: Fees Collection Overview
        self.fees_overview_tab = self.create_fees_overview_tab()
        self.tabs.addTab(self.fees_overview_tab, "Fees Collection Overview")
        
        # Tab 2: Class-wise Breakdown
        self.class_breakdown_tab = self.create_class_breakdown_tab()
        self.tabs.addTab(self.class_breakdown_tab, "Class-wise Breakdown")
        
        # Tab 3: Assets Report
        self.assets_tab = self.create_assets_tab()
        self.tabs.addTab(self.assets_tab, "Current Assets")
        
        # Tab 4: Expenditure - Capital
        self.capital_exp_tab = self.create_expenditure_tab("Capital")
        self.tabs.addTab(self.capital_exp_tab, "Capital Expenditure")
        
        # Tab 5: Expenditure - Operating
        self.operating_exp_tab = self.create_expenditure_tab("Operating")
        self.tabs.addTab(self.operating_exp_tab, "Operating Expenditure")
    
    def create_fees_overview_tab(self) -> QWidget:
        """Create the fees collection overview tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Summary Cards Container
        cards_container = QWidget()
        cards_layout = QHBoxLayout(cards_container)
        
        self.expected_card = FinancialReportCard("Expected Fees", "RWF 0", "Total fees expected")
        self.collected_card = FinancialReportCard("Collected Fees", "RWF 0", "Total fees collected")
        self.outstanding_card = FinancialReportCard("Outstanding", "RWF 0", "Remaining balance")
        self.completion_card = FinancialReportCard("Completion Rate", "0%", "Collection progress")
        
        cards_layout.addWidget(self.expected_card)
        cards_layout.addWidget(self.collected_card)
        cards_layout.addWidget(self.outstanding_card)
        cards_layout.addWidget(self.completion_card)
        
        layout.addWidget(cards_container)
        
        # Chart
        self.fees_chart_view = QChartView()
        self.fees_chart_view.setRenderHint(QPainter.Antialiasing)
        self.fees_chart_view.setMinimumHeight(300)
        layout.addWidget(self.fees_chart_view)
        
        return widget
    
    def create_class_breakdown_tab(self) -> QWidget:
        """Create the class-wise breakdown tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Table for class breakdown
        self.class_table = QTableWidget()
        self.class_table.setColumnCount(6)
        self.class_table.setHorizontalHeaderLabels([
            "Class", "Students", "Expected (RWF)", "Collected (RWF)", 
            "Outstanding (RWF)", "Completion %"
        ])
        self.class_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.class_table.setAlternatingRowColors(True)
        self.class_table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        layout.addWidget(self.class_table)
        
        # Total row
        self.class_total_label = QLabel("Total: RWF 0 expected | RWF 0 collected | 0% completion")
        self.class_total_label.setStyleSheet("""
            QLabel {
                background-color: #e9ecef;
                padding: 10px;
                font-weight: bold;
                font-size: 12pt;
            }
        """)
        layout.addWidget(self.class_total_label)
        
        return widget
    
    def create_assets_tab(self) -> QWidget:
        """Create the assets report tab"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Summary cards for assets
        cards_container = QWidget()
        cards_layout = QHBoxLayout(cards_container)
        
        self.total_assets_card = FinancialReportCard("Total Assets", "0", "All registered assets")
        self.assets_value_card = FinancialReportCard("Total Value", "RWF 0", "Combined asset value")
        self.active_assets_card = FinancialReportCard("Active Assets", "0", "Currently in use")
        
        cards_layout.addWidget(self.total_assets_card)
        cards_layout.addWidget(self.assets_value_card)
        cards_layout.addWidget(self.active_assets_card)
        
        layout.addWidget(cards_container)
        
        # Assets table
        self.assets_table = QTableWidget()
        self.assets_table.setColumnCount(6)
        self.assets_table.setHorizontalHeaderLabels([
            "Asset ID", "Asset Name", "Category", "Value (RWF)", 
            "Status", "Purchase Date"
        ])
        self.assets_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.assets_table.setAlternatingRowColors(True)
        self.assets_table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        layout.addWidget(self.assets_table)
        
        return widget
    
    def create_expenditure_tab(self, exp_type: str) -> QWidget:
        """Create expenditure report tab (Capital or Operating)"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Summary card
        card = FinancialReportCard(
            f"Total {exp_type} Expenditure", 
            "RWF 0", 
            f"All {exp_type.lower()} expenses"
        )
        layout.addWidget(card)
        
        # Store reference to card
        if exp_type == "Capital":
            self.capital_total_card = card
        else:
            self.operating_total_card = card
        
        # Expenditure table
        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels([
            "Category", "Sub-Category", "Transactions", "Total Amount (RWF)"
        ])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.setAlternatingRowColors(True)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        layout.addWidget(table)
        
        # Store reference to table
        if exp_type == "Capital":
            self.capital_table = table
        else:
            self.operating_table = table
        
        return widget
    
    def load_academic_years(self):
        """Load academic years into combo box"""
        years = self.data_manager.get_academic_years()
        self.year_combo.clear()
        
        for year_id, year_name in years:
            self.year_combo.addItem(year_name, year_id)
    
    def on_year_changed(self):
        """Handle academic year selection change"""
        if self.year_combo.currentIndex() < 0:
            return
        
        self.current_year_id = self.year_combo.currentData()
        self.load_terms()
    
    def load_terms(self):
        """Load terms for selected academic year"""
        if not self.current_year_id:
            return
        
        terms = self.data_manager.get_terms_by_year(self.current_year_id)
        self.term_combo.clear()
        
        for term_id, term_name in terms:
            self.term_combo.addItem(term_name, term_id)
    
    def on_term_changed(self):
        """Handle term selection change"""
        if self.term_combo.currentIndex() < 0:
            return
        
        self.current_term_id = self.term_combo.currentData()
    
    def generate_reports(self):
        """Generate all reports based on selected filters"""
        if not self.current_year_id or not self.current_term_id:
            QMessageBox.warning(self, "Warning", "Please select both Academic Year and Term")
            return
        
        try:
            # Generate each report
            self.generate_fees_overview()
            self.generate_class_breakdown()
            self.generate_assets_report()
            self.generate_expenditure_reports()
            
            self.export_btn.setEnabled(True)
            QMessageBox.information(self, "Success", "Reports generated successfully!")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error generating reports: {str(e)}")
    
    def generate_fees_overview(self):
        """Generate fees collection overview"""
        summary = self.data_manager.get_fees_completion_summary(self.current_term_id)
        
        # Update summary cards
        self.expected_card.findChild(QLabel).setText(
            f"RWF {summary['expected']:,.0f}"
        )
        self.collected_card.findChild(QLabel).setText(
            f"RWF {summary['collected']:,.0f}"
        )
        self.outstanding_card.findChild(QLabel).setText(
            f"RWF {summary['outstanding']:,.0f}"
        )
        self.completion_card.findChild(QLabel).setText(
            f"{summary['percentage']:.1f}%"
        )
        
        # Update chart
        self.update_fees_chart(summary)
    
    def update_fees_chart(self, summary: Dict):
        """Update the fees collection pie chart"""
        series = QPieSeries()
        series.append("Collected", summary['collected'])
        series.append("Outstanding", summary['outstanding'])
        
        # Color slices
        slice_collected = series.slices()[0]
        slice_collected.setBrush(QColor("#28a745"))
        slice_collected.setLabelVisible(True)
        slice_collected.setLabel(f"Collected: RWF {summary['collected']:,.0f}")
        
        slice_outstanding = series.slices()[1]
        slice_outstanding.setBrush(QColor("#dc3545"))
        slice_outstanding.setLabelVisible(True)
        slice_outstanding.setLabel(f"Outstanding: RWF {summary['outstanding']:,.0f}")
        
        chart = QChart()
        chart.addSeries(series)
        chart.setTitle("Fees Collection Status")
        chart.legend().setVisible(True)
        chart.legend().setAlignment(Qt.AlignBottom)
        
        self.fees_chart_view.setChart(chart)
    
    def generate_class_breakdown(self):
        """Generate class-wise fees breakdown"""
        breakdown = self.data_manager.get_class_fees_breakdown(self.current_term_id)
        
        self.class_table.setRowCount(len(breakdown))
        
        total_expected = 0
        total_collected = 0
        total_outstanding = 0
        
        for row_idx, data in enumerate(breakdown):
            self.class_table.setItem(row_idx, 0, QTableWidgetItem(data['class_name']))
            self.class_table.setItem(row_idx, 1, QTableWidgetItem(str(data['student_count'])))
            self.class_table.setItem(row_idx, 2, QTableWidgetItem(f"{data['expected']:,.0f}"))
            self.class_table.setItem(row_idx, 3, QTableWidgetItem(f"{data['collected']:,.0f}"))
            self.class_table.setItem(row_idx, 4, QTableWidgetItem(f"{data['outstanding']:,.0f}"))
            self.class_table.setItem(row_idx, 5, QTableWidgetItem(f"{data['percentage']:.1f}%"))
            
            # Center align percentage
            self.class_table.item(row_idx, 5).setTextAlignment(Qt.AlignCenter)
            
            total_expected += data['expected']
            total_collected += data['collected']
            total_outstanding += data['outstanding']
        
        # Update totals
        overall_percentage = (total_collected / total_expected * 100) if total_expected > 0 else 0
        self.class_total_label.setText(
            f"Total: RWF {total_expected:,.0f} expected | "
            f"RWF {total_collected:,.0f} collected | "
            f"{overall_percentage:.1f}% completion"
        )
    
    def generate_assets_report(self):
        """Generate current assets report"""
        assets = self.data_manager.get_assets_list()
        
        self.assets_table.setRowCount(len(assets))
        
        total_value = 0
        active_count = 0
        
        for row_idx, asset in enumerate(assets):
            self.assets_table.setItem(row_idx, 0, QTableWidgetItem(str(asset['asset_id'])))
            self.assets_table.setItem(row_idx, 1, QTableWidgetItem(asset['name'] or 'N/A'))
            self.assets_table.setItem(row_idx, 2, QTableWidgetItem(asset['category'] or 'N/A'))
            self.assets_table.setItem(row_idx, 3, QTableWidgetItem(f"{asset['value']:,.0f}"))
            self.assets_table.setItem(row_idx, 4, QTableWidgetItem(asset['status']))
            self.assets_table.setItem(row_idx, 5, QTableWidgetItem(asset['purchase_date'] or 'N/A'))
            
            total_value += asset['value']
            if asset['status'].lower() == 'active':
                active_count += 1
        
        # Update summary cards
        self.total_assets_card.findChildren(QLabel)[1].setText(str(len(assets)))
        self.assets_value_card.findChildren(QLabel)[1].setText(f"RWF {total_value:,.0f}")
        self.active_assets_card.findChildren(QLabel)[1].setText(str(active_count))
    
    def generate_expenditure_reports(self):
        """Generate both capital and operating expenditure reports"""
        # Capital Expenditure
        capital_data = self.data_manager.get_expenditure_report(self.current_year_id, "Capital")
        self.populate_expenditure_table(self.capital_table, capital_data, self.capital_total_card)
        
        # Operating Expenditure
        operating_data = self.data_manager.get_expenditure_report(self.current_year_id, "Operating")
        self.populate_expenditure_table(self.operating_table, operating_data, self.operating_total_card)
    
    def populate_expenditure_table(self, table: QTableWidget, data: List[Dict], card: FinancialReportCard):
        """Populate expenditure table with data"""
        table.setRowCount(len(data))
        
        total_amount = 0
        
        for row_idx, exp in enumerate(data):
            table.setItem(row_idx, 0, QTableWidgetItem(exp['category']))
            table.setItem(row_idx, 1, QTableWidgetItem(exp['subcategory']))
            table.setItem(row_idx, 2, QTableWidgetItem(str(exp['count'])))
            table.setItem(row_idx, 3, QTableWidgetItem(f"{exp['total']:,.0f}"))
            
            table.item(row_idx, 2).setTextAlignment(Qt.AlignCenter)
            table.item(row_idx, 3).setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            
            total_amount += exp['total']
        
        # Update total card
        card.findChildren(QLabel)[1].setText(f"RWF {total_amount:,.0f}")
    
    def export_reports(self):
        """Export all reports to CSV files"""
        if not self.current_year_id or not self.current_term_id:
            QMessageBox.warning(self, "Warning", "Please generate reports first")
            return
        
        # Get save directory
        directory = QFileDialog.getExistingDirectory(self, "Select Export Directory")
        if not directory:
            return
        
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            year_name = self.year_combo.currentText()
            term_name = self.term_combo.currentText()
            
            # Export Fees Overview
            self.export_fees_overview(directory, timestamp, year_name, term_name)
            
            # Export Class Breakdown
            self.export_class_breakdown(directory, timestamp, year_name, term_name)
            
            # Export Assets
            self.export_assets(directory, timestamp)
            
            # Export Expenditures
            self.export_expenditure(directory, timestamp, year_name, "Capital")
            self.export_expenditure(directory, timestamp, year_name, "Operating")
            
            QMessageBox.information(
                self, 
                "Success", 
                f"All reports exported successfully to:\n{directory}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error exporting reports: {str(e)}")
    
    def export_fees_overview(self, directory: str, timestamp: str, year: str, term: str):
        """Export fees overview to CSV"""
        filename = f"{directory}/fees_overview_{year}_{term}_{timestamp}.csv"
        summary = self.data_manager.get_fees_completion_summary(self.current_term_id)
        
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['School Fees Collection Report'])
            writer.writerow(['Academic Year:', year])
            writer.writerow(['Term:', term])
            writer.writerow(['Generated:', datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
            writer.writerow([])
            writer.writerow(['Metric', 'Amount (RWF)'])
            writer.writerow(['Expected Fees', f"{summary['expected']:,.2f}"])
            writer.writerow(['Collected Fees', f"{summary['collected']:,.2f}"])
            writer.writerow(['Outstanding Fees', f"{summary['outstanding']:,.2f}"])
            writer.writerow(['Completion Percentage', f"{summary['percentage']:.2f}%"])
    
    def export_class_breakdown(self, directory: str, timestamp: str, year: str, term: str):
        """Export class breakdown to CSV"""
        filename = f"{directory}/class_breakdown_{year}_{term}_{timestamp}.csv"
        breakdown = self.data_manager.get_class_fees_breakdown(self.current_term_id)
        
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Class-wise Fees Breakdown Report'])
            writer.writerow(['Academic Year:', year])
            writer.writerow(['Term:', term])
            writer.writerow(['Generated:', datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
            writer.writerow([])
            writer.writerow(['Class', 'Students', 'Expected (RWF)', 'Collected (RWF)', 
                           'Outstanding (RWF)', 'Completion %'])
            
            total_expected = 0
            total_collected = 0
            total_outstanding = 0
            
            for data in breakdown:
                writer.writerow([
                    data['class_name'],
                    data['student_count'],
                    f"{data['expected']:,.2f}",
                    f"{data['collected']:,.2f}",
                    f"{data['outstanding']:,.2f}",
                    f"{data['percentage']:.2f}%"
                ])
                total_expected += data['expected']
                total_collected += data['collected']
                total_outstanding += data['outstanding']
            
            writer.writerow([])
            overall_percentage = (total_collected / total_expected * 100) if total_expected > 0 else 0
            writer.writerow(['TOTAL', '', f"{total_expected:,.2f}", f"{total_collected:,.2f}",
                           f"{total_outstanding:,.2f}", f"{overall_percentage:.2f}%"])
    
    def export_assets(self, directory: str, timestamp: str):
        """Export assets list to CSV"""
        filename = f"{directory}/assets_report_{timestamp}.csv"
        assets = self.data_manager.get_assets_list()
        
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Current Assets Report'])
            writer.writerow(['Generated:', datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
            writer.writerow([])
            writer.writerow(['Asset ID', 'Asset Name', 'Category', 'Value (RWF)', 
                           'Status', 'Purchase Date'])
            
            total_value = 0
            for asset in assets:
                writer.writerow([
                    asset['asset_id'],
                    asset['name'] or 'N/A',
                    asset['category'] or 'N/A',
                    f"{asset['value']:,.2f}",
                    asset['status'],
                    asset['purchase_date'] or 'N/A'
                ])
                total_value += asset['value']
            
            writer.writerow([])
            writer.writerow(['TOTAL ASSETS', len(assets)])
            writer.writerow(['TOTAL VALUE', f"{total_value:,.2f}"])
    
    def export_expenditure(self, directory: str, timestamp: str, year: str, exp_type: str):
        """Export expenditure report to CSV"""
        filename = f"{directory}/{exp_type.lower()}_expenditure_{year}_{timestamp}.csv"
        data = self.data_manager.get_expenditure_report(self.current_year_id, exp_type)
        
        with open(filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow([f'{exp_type} Expenditure Report'])
            writer.writerow(['Academic Year:', year])
            writer.writerow(['Generated:', datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
            writer.writerow([])
            writer.writerow(['Category', 'Sub-Category', 'Transactions', 'Total Amount (RWF)'])
            
            total_amount = 0
            for exp in data:
                writer.writerow([
                    exp['category'],
                    exp['subcategory'],
                    exp['count'],
                    f"{exp['total']:,.2f}"
                ])
                total_amount += exp['total']
            
            writer.writerow([])
            writer.writerow(['TOTAL', '', '', f"{total_amount:,.2f}"])


# Example usage and integration
if __name__ == "__main__":
    import sys
    from PyQt5.QtWidgets import QApplication, QMainWindow
    
    class MainWindow(QMainWindow):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("School Finance Management - Report Generator")
            self.setGeometry(100, 100, 1400, 900)
            
            # Replace 'school_finance.db' with your actual database path
            self.report_widget = ReportGeneratorWidget('database.db')
            self.setCentralWidget(self.report_widget)
    
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())