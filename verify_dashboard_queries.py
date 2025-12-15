import sqlite3
from datetime import datetime

def get_db_connection():
    try:
        conn = sqlite3.connect("Database.db")
        conn.row_factory = sqlite3.Row
        return conn
    except sqlite3.Error as e:
        print(f"Database connection error: {e}")
        return None

def verify_queries():
    conn = get_db_connection()
    if not conn:
        return

    c = conn.cursor()
    print("--- Verifying Dashboard Queries ---")

    # 1. Load Terms
    print("\n1. Testing Load Terms:")
    try:
        c.execute("SELECT TermID, TermName, StartDate, EndDate FROM Term ORDER BY TermID DESC")
        terms = c.fetchall()
        print(f"Found {len(terms)} terms.")
        if terms:
            term = terms[0]
            term_id = term['TermID']
            start_date = term['StartDate']
            end_date = term['EndDate']
            print(f"Using Term: {term['TermName']} (ID: {term_id}, Start: {start_date}, End: {end_date})")
            
            # 2. Fees Data
            print("\n2. Testing Fees Data:")
            c.execute("""
                SELECT SUM(
                    (f.Amount * (st.PayRate / 100.0))
                ) as TotalExpected
                FROM StudentTerm st
                JOIN Fees f ON st.ClassID = f.Class AND st.TermID = f.Term
                WHERE st.TermID = ?
            """, (term_id,))
            total_expected = c.fetchone()['TotalExpected']
            print(f"Total Expected Fees: {total_expected}")

            c.execute("""
                SELECT SUM(AmountPaid) 
                FROM FeesPayment fp
                JOIN StudentTerm st ON fp.StudentTermID = st.StudentTermID
                WHERE st.TermID = ?
            """, (term_id,))
            total_collected = c.fetchone()[0]
            print(f"Total Collected Fees: {total_collected}")

            # 3. Expenses Data
            print("\n3. Testing Expenses Data:")
            c.execute("""
                SELECT SUM(ExpAmount) 
                FROM Expenses 
                WHERE ExpDate BETWEEN ? AND ?
            """, (start_date, end_date))
            total_expenses = c.fetchone()[0]
            print(f"Total Expenses: {total_expenses}")

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
            expenses_by_cat = c.fetchall()
            print(f"Expenses by Category: {[dict(row) for row in expenses_by_cat]}")

            # 4. Supplies Data
            print("\n4. Testing Supplies Data:")
            c.execute("""
                SELECT SUM(Price) 
                FROM SuppliesInventory 
                WHERE Type = 'In' AND Date BETWEEN ? AND ?
            """, (start_date, end_date))
            total_supplies = c.fetchone()[0]
            print(f"Total Supplies Cost: {total_supplies}")

            c.execute("""
                SELECT s.Name, SUM(si.Quantity), SUM(si.Price)
                FROM SuppliesInventory si
                JOIN Supplies s ON si.Supplies_ID = s.SuppliesID
                WHERE si.Type = 'In' AND si.Date BETWEEN ? AND ?
                GROUP BY s.Name
                ORDER BY SUM(si.Price) DESC
                LIMIT 5
            """, (start_date, end_date))
            supplies_summary = c.fetchall()
            print(f"Supplies Summary: {[dict(row) for row in supplies_summary]}")

            # 5. Outstanding Fees Table
            print("\n5. Testing Outstanding Fees Table:")
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
            outstanding_table = c.fetchall()
            print(f"Outstanding Fees Table: {[dict(row) for row in outstanding_table]}")

        else:
            print("No terms found. Cannot test term-specific queries.")

    except sqlite3.Error as e:
        print(f"SQL Error: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    verify_queries()
