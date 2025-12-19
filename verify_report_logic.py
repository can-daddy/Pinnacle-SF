from finance_service import FinanceService
from report_generator import ReportGenerator
import os

def verify():
    print("Starting Verification...")
    service = FinanceService()
    generator = ReportGenerator()

    # 1. Get Terms
    terms = service.get_all_terms()
    if not terms:
        print("No terms found. Cannot verify.")
        return

    term_id = terms[0]['TermID']
    term_name = terms[0]['TermName']
    print(f"Verifying for Term: {term_name} (ID: {term_id})")

    # 2. Get Dates
    dates = service.get_term_dates(term_id)
    start_date = dates['StartDate']
    end_date = dates['EndDate']
    print(f"Date Range: {start_date} to {end_date}")

    # 3. Get Data
    fees = service.get_term_fees_summary(term_id)
    expenses = service.get_term_expenses(start_date, end_date)
    supplies = service.get_term_supplies(start_date, end_date)
    
    print(f"Fees: {fees}")
    print(f"Expenses: {expenses}")
    print(f"Supplies: {supplies}")

    # 4. Generate PDF
    pdf_path = "verify_report.pdf"
    try:
        generator.generate_pdf(pdf_path, term_name, start_date, end_date, fees, expenses, supplies)
        print(f"PDF Generated: {pdf_path}")
    except Exception as e:
        print(f"PDF Generation Failed: {e}")

    # 5. Generate Excel
    xlsx_path = "verify_report.xlsx"
    try:
        fees_breakdown = service.get_outstanding_fees_by_class(term_id)
        expenses_breakdown = service.get_expenses_by_category(start_date, end_date)
        generator.generate_excel(xlsx_path, term_name, start_date, end_date, fees, expenses, supplies, fees_breakdown, expenses_breakdown)
        print(f"Excel Generated: {xlsx_path}")
    except Exception as e:
        print(f"Excel Generation Failed: {e}")

if __name__ == "__main__":
    verify()
