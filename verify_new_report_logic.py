from finance_service import FinanceService

def verify_finance_service():
    fs = FinanceService()
    
    # Get a valid term ID (assuming 1 exists, or fetch one)
    terms = fs.get_all_terms()
    if not terms:
        print("No terms found.")
        return
    
    term_id = terms[0]['TermID']
    term_name = terms[0]['TermName']
    print(f"Testing with Term: {term_name} (ID: {term_id})")
    
    # Get dates for the term
    dates = fs.get_term_dates(term_id)
    start_date = dates['StartDate']
    end_date = dates['EndDate']
    print(f"Date Range: {start_date} to {end_date}")
    
    print("\n--- Testing get_term_fees_breakdown_by_class ---")
    fees_breakdown = fs.get_term_fees_breakdown_by_class(term_id)
    for item in fees_breakdown[:3]: # Print first 3
        print(item)
        
    print("\n--- Testing get_supplies_by_category ---")
    supplies_cat = fs.get_supplies_by_category(start_date, end_date)
    for item in supplies_cat:
        print(item)

    print("\n--- Testing get_expenses_by_category_with_percentage ---")
    expenses_cat = fs.get_expenses_by_category_with_percentage(start_date, end_date)
    for item in expenses_cat:
        print(item)
        
    print("\n--- Testing get_supplies_details ---")
    supplies_details = fs.get_supplies_details(start_date, end_date)
    for item in supplies_details[:3]:
        print(item)
        
    print("\n--- Testing get_expenses_details ---")
    expenses_details = fs.get_expenses_details(start_date, end_date)
    for item in expenses_details[:3]:
        print(item)

if __name__ == "__main__":
    verify_finance_service()
