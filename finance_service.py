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
                        "balance": balance,
                        "percentage": percent
                    })
            return result

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

    def get_all_terms(self):
        """Returns list of all terms ordered by ID desc."""
        with get_db_connection() as (conn, c):
            c.execute("SELECT TermID, TermName FROM Term ORDER BY TermID DESC")
            return c.fetchall()
