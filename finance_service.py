import sqlite3
from contextlib import contextmanager

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

class FinanceService:
    def get_term_dates(self, term_id):
        with get_db_connection() as (conn, c):
            c.execute("SELECT StartDate, EndDate FROM Term WHERE TermID = ?", (term_id,))
            return c.fetchone()

    def get_term_fees_summary(self, term_id):
        """
        Returns a dictionary with:
        - expected: Total fees expected for the term
        - collected: Total fees collected for the term
        - outstanding: expected - collected
        """
        with get_db_connection() as (conn, c):
            # Total Fees Expected
            c.execute("""
                SELECT SUM(
                    (f.Amount * (st.PayRate / 100.0))
                ) as TotalExpected
                FROM StudentTerm st
                JOIN Fees f ON st.ClassID = f.Class AND st.TermID = f.Term
                WHERE st.TermID = ?
            """, (term_id,))
            total_expected = c.fetchone()['TotalExpected'] or 0.0

            # Total Fees Collected
            c.execute("""
                SELECT SUM(AmountPaid) 
                FROM FeesPayment fp
                JOIN StudentTerm st ON fp.StudentTermID = st.StudentTermID
                WHERE st.TermID = ?
            """, (term_id,))
            total_collected = c.fetchone()[0] or 0.0
            
            outstanding = max(0, total_expected - total_collected)
            
            return {
                "expected": total_expected,
                "collected": total_collected,
                "outstanding": outstanding
            }

    def get_term_expenses(self, start_date, end_date):
        """Returns total expenses between start and end date."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT SUM(ExpAmount) 
                FROM Expenses 
                WHERE ExpDate BETWEEN ? AND ?
            """, (start_date, end_date))
            return c.fetchone()[0] or 0.0

    def get_term_supplies(self, start_date, end_date):
        """Returns total supplies cost (Type='In') between start and end date."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT SUM(Price) 
                FROM SuppliesInventory 
                WHERE Type = 'In' AND Date BETWEEN ? AND ?
            """, (start_date, end_date))
            return c.fetchone()[0] or 0.0

    def get_expenses_by_category(self, start_date, end_date, limit=5):
        """Returns list of (CategoryName, Amount) tuples."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT c.CatName, SUM(e.ExpAmount)
                FROM Expenses e
                JOIN SubCategory sc ON e.ExpSubCategory = sc.SubCatID
                JOIN Category c ON sc.CategoryID = c.CatID
                WHERE e.ExpDate BETWEEN ? AND ?
                GROUP BY c.CatName
                ORDER BY SUM(e.ExpAmount) DESC
                LIMIT ?
            """, (start_date, end_date, limit))
            data = c.fetchall()
            return [(row[0], row[1]) for row in data]

    def get_expenses_by_category_with_percentage(self, start_date, end_date):
        """Returns list of (CategoryName, Amount, Percentage) tuples."""
        with get_db_connection() as (conn, c):
            # First get total for percentage calculation
            c.execute("""
                SELECT SUM(ExpAmount) 
                FROM Expenses 
                WHERE ExpDate BETWEEN ? AND ?
            """, (start_date, end_date))
            total_expenses = c.fetchone()[0] or 0.0

            if total_expenses == 0:
                return []

            c.execute("""
                SELECT c.CatName, SUM(e.ExpAmount) as Amount
                FROM Expenses e
                JOIN SubCategory sc ON e.ExpSubCategory = sc.SubCatID
                JOIN Category c ON sc.CategoryID = c.CatID
                WHERE e.ExpDate BETWEEN ? AND ?
                GROUP BY c.CatName
                ORDER BY Amount DESC
            """, (start_date, end_date))
            data = c.fetchall()
            return [(row['CatName'], row['Amount'], (row['Amount']/total_expenses)*100) for row in data]

    def get_outstanding_fees_by_class(self, term_id, limit=10):
        """Returns list of dictionaries with class_name, expected, collected, balance, percentage."""
        with get_db_connection() as (conn, c):
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
                LIMIT ?
            """, (term_id, term_id, limit))
            
            rows = c.fetchall()
            result = []
            for row in rows:
                expected = row['Expected'] or 0
                collected = row['Collected'] or 0
                balance = expected - collected
                if balance > 0:
                    percent = (balance / expected * 100) if expected > 0 else 0
                    result.append({
                        "class_name": row['ClassName'],
                        "expected": expected,
                        "collected": collected,
                        "balance": balance,
                        "percentage": percent
                    })
            return result

    def get_term_fees_breakdown_by_class(self, term_id):
        """Returns list of dictionaries with class_name, expected, collected, balance, percentage."""
        return self.get_outstanding_fees_by_class(term_id, limit=-1)

    def get_top_supplies(self, start_date, end_date, limit=5):
        """Returns list of (Name, Quantity, Cost) tuples."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT s.Name, SUM(si.Quantity), SUM(si.Price)
                FROM SuppliesInventory si
                JOIN Supplies s ON si.Supplies_ID = s.SuppliesID
                WHERE si.Type = 'In' AND si.Date BETWEEN ? AND ?
                GROUP BY s.Name
                ORDER BY SUM(si.Price) DESC
                LIMIT ?
            """, (start_date, end_date, limit))
            
            data = c.fetchall()
            return [(row[0], row[1], row[2]) for row in data]

    def get_supplies_by_category(self, start_date, end_date):
        """Returns list of (CategoryName, Amount, Percentage) tuples."""
        with get_db_connection() as (conn, c):
            # First get total for percentage calculation
            c.execute("""
                SELECT SUM(si.Price)
                FROM SuppliesInventory si
                WHERE si.Type = 'In' AND si.Date BETWEEN ? AND ?
            """, (start_date, end_date))
            total_supplies = c.fetchone()[0] or 0.0

            if total_supplies == 0:
                return []

            c.execute("""
                SELECT c.CatName, SUM(si.Price) as Amount
                FROM SuppliesInventory si
                JOIN Supplies s ON si.Supplies_ID = s.SuppliesID
                JOIN SubCategory sc ON s.Subcategory = sc.SubCatID
                JOIN Category c ON sc.CategoryID = c.CatID
                WHERE si.Type = 'In' AND si.Date BETWEEN ? AND ?
                GROUP BY c.CatName
                ORDER BY Amount DESC
            """, (start_date, end_date))
            
            data = c.fetchall()
            return [(row['CatName'], row['Amount'], (row['Amount']/total_supplies)*100) for row in data]

    def get_supplies_details(self, start_date, end_date):
        """Returns list of dictionaries for detailed supplies report."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT si.Date, c.CatName, s.Name, si.Quantity, si.Price
                FROM SuppliesInventory si
                JOIN Supplies s ON si.Supplies_ID = s.SuppliesID
                JOIN SubCategory sc ON s.Subcategory = sc.SubCatID
                JOIN Category c ON sc.CategoryID = c.CatID
                WHERE si.Type = 'In' AND si.Date BETWEEN ? AND ?
                ORDER BY si.Date DESC
            """, (start_date, end_date))
            
            rows = c.fetchall()
            return [{
                'date': row['Date'],
                'category': row['CatName'],
                'item': row['Name'],
                'quantity': row['Quantity'],
                'cost': row['Price']
            } for row in rows]

    def get_expenses_details(self, start_date, end_date):
        """Returns list of dictionaries for detailed expenses report."""
        with get_db_connection() as (conn, c):
            c.execute("""
                SELECT e.ExpDate, c.CatName, e.ExpDescription, e.ExpAmount
                FROM Expenses e
                JOIN SubCategory sc ON e.ExpSubCategory = sc.SubCatID
                JOIN Category c ON sc.CategoryID = c.CatID
                WHERE e.ExpDate BETWEEN ? AND ?
                ORDER BY e.ExpDate DESC
            """, (start_date, end_date))
            
            rows = c.fetchall()
            return [{
                'date': row['ExpDate'],
                'category': row['CatName'],
                'description': row['ExpDescription'],
                'amount': row['ExpAmount']
            } for row in rows]

    def get_all_terms(self):
        """Returns list of all terms ordered by ID desc."""
        with get_db_connection() as (conn, c):
            c.execute("SELECT TermID, TermName FROM Term ORDER BY TermID DESC")
            return c.fetchall()
