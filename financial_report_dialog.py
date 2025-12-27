from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
                             QPushButton, QFrame, QMessageBox, QFileDialog)
from PyQt5.QtCore import Qt
from finance_service import FinanceService
from report_generator import ReportGenerator
import os

class FinancialReportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Term Financial Summary Report")
        self.setFixedSize(600, 500)
        self.finance_service = FinanceService()
        self.report_generator = ReportGenerator()
        self.current_term_data = {}
        self.init_ui()
        self.load_terms()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        # Header
        header_lbl = QLabel("Generate Financial Report")
        header_lbl.setStyleSheet("font-size: 20px; font-weight: bold; color: #2D3748;")
        layout.addWidget(header_lbl)

        # Term Selection
        term_layout = QHBoxLayout()
        term_lbl = QLabel("Select Term:")
        self.term_selector = QComboBox()
        self.term_selector.currentIndexChanged.connect(self.load_term_data)
        term_layout.addWidget(term_lbl)
        term_layout.addWidget(self.term_selector)
        layout.addLayout(term_layout)

        # Disclaimer
        disclaimer_frame = QFrame()
        disclaimer_frame.setStyleSheet("""
            QFrame {
                background-color: #FFF5F5;
                border: 1px solid #FC8181;
                border-radius: 5px;
                padding: 10px;
            }
        """)
        disclaimer_layout = QVBoxLayout(disclaimer_frame)
        disclaimer_lbl = QLabel("<b>IMPORTANT:</b> This report does not include donations, grants, rentals, or other external income sources. It reflects only the data recorded in this system.")
        disclaimer_lbl.setWordWrap(True)
        disclaimer_lbl.setStyleSheet("color: #C53030; font-size: 12px;")
        disclaimer_layout.addWidget(disclaimer_lbl)
        layout.addWidget(disclaimer_frame)

        # Summary Cards
        self.cards_layout = QHBoxLayout()
        self.lbl_fees = self.create_summary_label("Fees Collected", "0.00")
        self.lbl_expenses = self.create_summary_label("Total Expenses", "0.00")
        self.lbl_supplies = self.create_summary_label("Supplies Cost", "0.00")
        
        self.cards_layout.addWidget(self.lbl_fees)
        self.cards_layout.addWidget(self.lbl_expenses)
        self.cards_layout.addWidget(self.lbl_supplies)
        layout.addLayout(self.cards_layout)

        # Export Buttons
        btn_layout = QHBoxLayout()
        btn_pdf = QPushButton("Export PDF")
        btn_pdf.setStyleSheet("background-color: #E53E3E; color: white; padding: 10px; font-weight: bold;")
        btn_pdf.clicked.connect(self.export_pdf)
        
        btn_excel = QPushButton("Export Excel")
        btn_excel.setStyleSheet("background-color: #38A169; color: white; padding: 10px; font-weight: bold;")
        btn_excel.clicked.connect(self.export_excel)
        
        btn_layout.addWidget(btn_pdf)
        btn_layout.addWidget(btn_excel)
        layout.addLayout(btn_layout)
        
        layout.addStretch()

    def create_summary_label(self, title, value):
        frame = QFrame()
        frame.setStyleSheet("background-color: #F7FAFC; border-radius: 5px; border: 1px solid #E2E8F0;")
        l = QVBoxLayout(frame)
        t = QLabel(title)
        t.setStyleSheet("color: #718096; font-size: 12px;")
        v = QLabel(value)
        v.setStyleSheet("color: #2D3748; font-size: 18px; font-weight: bold;")
        v.setAlignment(Qt.AlignCenter)
        l.addWidget(t, alignment=Qt.AlignCenter)
        l.addWidget(v, alignment=Qt.AlignCenter)
        frame.value_label = v
        return frame

    def load_terms(self):
        self.term_selector.clear()
        try:
            terms = self.finance_service.get_all_terms()
            for term in terms:
                self.term_selector.addItem(f"{term['TermName']} ({term['Year']})", term['TermID'])
        except Exception as e:
            print(f"Error loading terms: {e}")

    def load_term_data(self):
        term_id = self.term_selector.currentData()
        if not term_id:
            return

        try:
            term_dates = self.finance_service.get_term_dates(term_id)
            if not term_dates:
                return
            start_date = term_dates['StartDate']
            end_date = term_dates['EndDate']

            fees_summary = self.finance_service.get_term_fees_summary(term_id)
            total_expenses = self.finance_service.get_term_expenses(start_date, end_date)
            total_supplies = self.finance_service.get_term_supplies(start_date, end_date)

            self.lbl_fees.value_label.setText(f"{fees_summary['collected']:,.2f}")
            self.lbl_expenses.value_label.setText(f"{total_expenses:,.2f}")
            self.lbl_supplies.value_label.setText(f"{total_supplies:,.2f}")

            # Store for export
            self.current_term_data = {
                'term_name': self.term_selector.currentText(),
                'start_date': start_date,
                'end_date': end_date,
                'fees': fees_summary,
                'expenses': total_expenses,
                'supplies': total_supplies,
                'term_id': term_id
            }

        except Exception as e:
            print(f"Error loading data: {e}")

    def export_pdf(self):
        if not self.current_term_data:
            return
        
        filename, _ = QFileDialog.getSaveFileName(self, "Save PDF Report", f"Financial_Report_{self.current_term_data['term_name']}.pdf", "PDF Files (*.pdf)")
        if filename:
            try:
                self.report_generator.generate_pdf(
                    filename,
                    self.current_term_data['term_name'],
                    self.current_term_data['start_date'],
                    self.current_term_data['end_date'],
                    self.current_term_data['fees'],
                    self.current_term_data['expenses'],
                    self.current_term_data['supplies']
                )
                QMessageBox.information(self, "Success", "PDF Report generated successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to generate PDF: {str(e)}")

    def export_excel(self):
        if not self.current_term_data:
            return

        filename, _ = QFileDialog.getSaveFileName(self, "Save Excel Report", f"Financial_Report_{self.current_term_data['term_name']}.xlsx", "Excel Files (*.xlsx)")
        if filename:
            try:
                # Fetch breakdown data for Excel
                fees_breakdown = self.finance_service.get_outstanding_fees_by_class(self.current_term_data['term_id'], limit=100)
                expenses_breakdown = self.finance_service.get_expenses_by_category(self.current_term_data['start_date'], self.current_term_data['end_date'], limit=100)

                self.report_generator.generate_excel(
                    filename,
                    self.current_term_data['term_name'],
                    self.current_term_data['start_date'],
                    self.current_term_data['end_date'],
                    self.current_term_data['fees'],
                    self.current_term_data['expenses'],
                    self.current_term_data['supplies'],
                    fees_breakdown,
                    expenses_breakdown
                )
                QMessageBox.information(self, "Success", "Excel Report generated successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to generate Excel: {str(e)}")
