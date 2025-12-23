from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, 
                             QPushButton, QFrame, QMessageBox, QFileDialog, QScrollArea)
from PyQt5.QtCore import Qt
from finance_service import FinanceService
from report_generator import ReportGenerator

class Reports(QWidget):
    def __init__(self):
        super().__init__()
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
        header_lbl = QLabel("Financial Reports")
        header_lbl.setStyleSheet("font-size: 24px; font-weight: bold; color: #2D3748;")
        layout.addWidget(header_lbl)

        # Controls Container
        controls_frame = QFrame()
        controls_frame.setStyleSheet("background-color: white; border-radius: 8px; border: 1px solid #E2E8F0;")
        controls_layout = QHBoxLayout(controls_frame)
        controls_layout.setContentsMargins(20, 20, 20, 20)
        
        # Term Selection
        term_lbl = QLabel("Select Term:")
        term_lbl.setStyleSheet("font-size: 14px; font-weight: bold; color: #4A5568;")
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
        self.term_selector.currentIndexChanged.connect(self.load_term_data)
        
        controls_layout.addWidget(term_lbl)
        controls_layout.addWidget(self.term_selector)
        controls_layout.addStretch()
        
        layout.addWidget(controls_frame)

        # Summary Cards Area
        self.cards_layout = QHBoxLayout()
        self.lbl_fees = self.create_summary_card("Fees Collected", "0.00", "#3182CE")
        self.lbl_expenses = self.create_summary_card("Total Expenses", "0.00", "#E53E3E")
        self.lbl_supplies = self.create_summary_card("Supplies Cost", "0.00", "#805AD5")
        
        self.cards_layout.addWidget(self.lbl_fees)
        self.cards_layout.addWidget(self.lbl_expenses)
        self.cards_layout.addWidget(self.lbl_supplies)
        layout.addLayout(self.cards_layout)

        # Actions Area
        actions_frame = QFrame()
        actions_frame.setStyleSheet("background-color: white; border-radius: 8px; border: 1px solid #E2E8F0;")
        actions_layout = QVBoxLayout(actions_frame)
        actions_layout.setContentsMargins(20, 20, 20, 20)
        
        actions_title = QLabel("Export Options")
        actions_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #2D3748; margin-bottom: 10px;")
        actions_layout.addWidget(actions_title)
        
        btn_layout = QHBoxLayout()
        
        btn_pdf = QPushButton("Generate PDF Report")
        btn_pdf.setCursor(Qt.PointingHandCursor)
        btn_pdf.setObjectName("largeDangerButton")
        btn_pdf.clicked.connect(self.export_pdf)
        
        btn_excel = QPushButton("Generate Excel Report")
        btn_excel.setCursor(Qt.PointingHandCursor)
        btn_excel.setObjectName("largeSuccessButton")
        btn_excel.clicked.connect(self.export_excel)
        
        btn_layout.addWidget(btn_pdf)
        btn_layout.addWidget(btn_excel)
        btn_layout.addStretch()
        
        actions_layout.addLayout(btn_layout)
        layout.addWidget(actions_frame)
        
        # Disclaimer
        disclaimer_lbl = QLabel("Note: Reports reflect only system-recorded data. External income (grants, donations) is excluded.")
        disclaimer_lbl.setStyleSheet("color: #718096; font-style: italic; margin-top: 20px;")
        layout.addWidget(disclaimer_lbl)
        
        layout.addStretch()

    def create_summary_card(self, title, value, color):
        frame = QFrame()
        frame.setStyleSheet(f"""
            QFrame {{
                background-color: white;
                border-radius: 8px;
                border-left: 5px solid {color};
                border-top: 1px solid #E2E8F0;
                border-right: 1px solid #E2E8F0;
                border-bottom: 1px solid #E2E8F0;
                min-height: 100px;
            }}
        """)
        l = QVBoxLayout(frame)
        t = QLabel(title)
        t.setStyleSheet("color: #718096; font-size: 14px; border: none;")
        v = QLabel(value)
        v.setStyleSheet("color: #2D3748; font-size: 24px; font-weight: bold; border: none;")
        v.setAlignment(Qt.AlignLeft)
        l.addWidget(t)
        l.addWidget(v)
        frame.value_label = v
        return frame

    def load_terms(self):
        self.term_selector.clear()
        try:
            terms = self.finance_service.get_all_terms()
            for term in terms:
                self.term_selector.addItem(term['TermName'], term['TermID'])
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

            # Store basic info for export
            self.current_term_data = {
                'term_name': self.term_selector.currentText(),
                'start_date': start_date,
                'end_date': end_date,
                'fees_summary': fees_summary,
                'expenses_total': total_expenses,
                'supplies_total': total_supplies,
                'term_id': term_id
            }

        except Exception as e:
            print(f"Error loading data: {e}")

    def fetch_detailed_data(self):
        """Helper to fetch all detailed data needed for reports"""
        if not self.current_term_data:
            return None
            
        term_id = self.current_term_data['term_id']
        start_date = self.current_term_data['start_date']
        end_date = self.current_term_data['end_date']
        
        return {
            'fees_breakdown': self.finance_service.get_term_fees_breakdown_by_class(term_id),
            'supplies_breakdown': self.finance_service.get_supplies_by_category(start_date, end_date),
            'supplies_details': self.finance_service.get_supplies_details(start_date, end_date),
            'expenses_breakdown': self.finance_service.get_expenses_by_category_with_percentage(start_date, end_date),
            'expenses_details': self.finance_service.get_expenses_details(start_date, end_date)
        }

    def export_pdf(self):
        if not self.current_term_data:
            return
        
        filename, _ = QFileDialog.getSaveFileName(self, "Save PDF Report", f"Financial_Report_{self.current_term_data['term_name']}.pdf", "PDF Files (*.pdf)")
        if filename:
            try:
                details = self.fetch_detailed_data()
                self.report_generator.generate_pdf(
                    filename,
                    self.current_term_data['term_name'],
                    self.current_term_data['start_date'],
                    self.current_term_data['end_date'],
                    self.current_term_data['fees_summary'],
                    details['fees_breakdown'],
                    self.current_term_data['supplies_total'],
                    details['supplies_breakdown'],
                    details['supplies_details'],
                    self.current_term_data['expenses_total'],
                    details['expenses_breakdown'],
                    details['expenses_details']
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
                details = self.fetch_detailed_data()
                self.report_generator.generate_excel(
                    filename,
                    self.current_term_data['term_name'],
                    self.current_term_data['start_date'],
                    self.current_term_data['end_date'],
                    self.current_term_data['fees_summary'],
                    details['fees_breakdown'],
                    self.current_term_data['supplies_total'],
                    details['supplies_breakdown'],
                    details['supplies_details'],
                    self.current_term_data['expenses_total'],
                    details['expenses_breakdown'],
                    details['expenses_details']
                )
                QMessageBox.information(self, "Success", "Excel Report generated successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to generate Excel: {str(e)}")
