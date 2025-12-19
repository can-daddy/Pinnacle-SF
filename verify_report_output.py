from finance_service import FinanceService
from report_generator import ReportGenerator
import os

def verify_report_output():
    fs = FinanceService()
    rg = ReportGenerator()
    
    # Get a valid term
    terms = fs.get_all_terms()
    if not terms:
        print("No terms found.")
        return
    
    term_id = terms[0]['TermID']
    term_name = terms[0]['TermName']
    print(f"Generating report for Term: {term_name} (ID: {term_id})")
    
    # Get dates
    dates = fs.get_term_dates(term_id)
    start_date = dates['StartDate']
    end_date = dates['EndDate']
    
    # Fetch Data
    fees_summary = fs.get_term_fees_summary(term_id)
    fees_breakdown = fs.get_term_fees_breakdown_by_class(term_id)
    
    supplies_total = fs.get_term_supplies(start_date, end_date)
    supplies_breakdown = fs.get_supplies_by_category(start_date, end_date)
    supplies_details = fs.get_supplies_details(start_date, end_date)
    
    expenses_total = fs.get_term_expenses(start_date, end_date)
    expenses_breakdown = fs.get_expenses_by_category_with_percentage(start_date, end_date)
    expenses_details = fs.get_expenses_details(start_date, end_date)
    
    # Generate PDF
    pdf_path = f"Verify_Report_{term_name}.pdf"
    print(f"Generating PDF: {pdf_path}")
    rg.generate_pdf(
        pdf_path,
        term_name,
        start_date,
        end_date,
        fees_summary,
        fees_breakdown,
        supplies_total,
        supplies_breakdown,
        supplies_details,
        expenses_total,
        expenses_breakdown,
        expenses_details
    )
    
    # Generate Excel
    excel_path = f"Verify_Report_{term_name}.xlsx"
    print(f"Generating Excel: {excel_path}")
    rg.generate_excel(
        excel_path,
        term_name,
        start_date,
        end_date,
        fees_summary,
        fees_breakdown,
        supplies_total,
        supplies_breakdown,
        supplies_details,
        expenses_total,
        expenses_breakdown,
        expenses_details
    )
    
    print("Verification Complete.")

if __name__ == "__main__":
    verify_report_output()
