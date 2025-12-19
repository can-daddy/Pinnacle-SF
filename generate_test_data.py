import sqlite3
from faker import Faker
import random
from datetime import datetime, timedelta

fake = Faker()

def get_db_connection():
    conn = sqlite3.connect("Database.db")
    conn.row_factory = sqlite3.Row
    return conn

def generate_data():
    conn = get_db_connection()
    c = conn.cursor()

    print("Generating test data...")

    # 1. Academic Year
    print("Creating Academic Year...")
    c.execute("INSERT OR IGNORE INTO AcademicYear (Year) VALUES (?)", ("2024-2025",))
    academic_year_id = c.lastrowid
    if not academic_year_id:
        c.execute("SELECT AcademicYearID FROM AcademicYear WHERE Year = ?", ("2024-2025",))
        academic_year_id = c.fetchone()[0]

    # 2. Terms
    print("Creating Terms...")
    terms = [
        ("Term 1", "2024-09-01", "2024-12-15"),
        ("Term 2", "2025-01-10", "2025-04-10"),
        ("Term 3", "2025-05-01", "2025-07-30")
    ]
    term_ids = []
    for name, start, end in terms:
        c.execute("SELECT TermID FROM Term WHERE TermName = ? AND AcademicYear = ?", (name, academic_year_id))
        existing = c.fetchone()
        if existing:
            term_ids.append(existing[0])
        else:
            c.execute("INSERT INTO Term (TermName, AcademicYear, StartDate, EndDate) VALUES (?, ?, ?, ?)",
                      (name, academic_year_id, start, end))
            term_ids.append(c.lastrowid)

    # 3. Classes
    print("Creating Classes...")
    class_names = ["Form 1", "Form 2", "Form 3", "Form 4"]
    class_ids = []
    for name in class_names:
        c.execute("SELECT ClassID FROM Class WHERE ClassName = ? AND AcademicYear = ?", (name, academic_year_id))
        existing = c.fetchone()
        if existing:
            class_ids.append(existing[0])
        else:
            c.execute("INSERT INTO Class (ClassName, AcademicYear) VALUES (?, ?)", (name, academic_year_id))
            class_ids.append(c.lastrowid)

    # 4. Students
    print("Creating Students...")
    student_ids = []
    for _ in range(50): # Generate 50 students
        student_id = fake.unique.random_number(digits=6)
        name = fake.name()
        fee_payer = "Parent"
        pay_rate = random.choice([100, 100, 100, 75, 0]) # Mostly full pay, some scholarship
        parent_name = fake.name()
        contact = fake.phone_number()
        
        c.execute("INSERT OR IGNORE INTO Students (StudentID, Student_Name, FeePayer, PayRate, Parent_Name, Parent_Contact) VALUES (?, ?, ?, ?, ?, ?)",
                  (student_id, name, fee_payer, pay_rate, parent_name, contact))
        student_ids.append(student_id)

    # 5. Fees Structure
    print("Setting up Fees...")
    for class_id in class_ids:
        for term_id in term_ids:
            amount = random.choice([15000, 20000, 25000])
            c.execute("SELECT FeesID FROM Fees WHERE Class = ? AND Term = ?", (class_id, term_id))
            if not c.fetchone():
                c.execute("INSERT INTO Fees (Class, Term, Amount) VALUES (?, ?, ?)", (class_id, term_id, amount))

    # 6. Enrollments & Payments
    print("Enrolling Students and Processing Payments...")
    for student_id in student_ids:
        # Enroll in a random class for all terms (simplified)
        class_id = random.choice(class_ids)
        
        for term_id in term_ids:
            # Check if already enrolled
            c.execute("SELECT StudentTermID FROM StudentTerm WHERE StudentID = ? AND ClassID = ? AND TermID = ?", 
                      (student_id, class_id, term_id))
            if c.fetchone():
                continue

            # Get PayRate
            c.execute("SELECT PayRate FROM Students WHERE StudentID = ?", (student_id,))
            pay_rate = c.fetchone()[0]

            c.execute("INSERT INTO StudentTerm (StudentID, ClassID, TermID, PayRate) VALUES (?, ?, ?, ?)",
                      (student_id, class_id, term_id, pay_rate))
            student_term_id = c.lastrowid

            # Make a payment (maybe)
            if random.random() > 0.2: # 80% chance of payment
                # Get Fee Amount
                c.execute("SELECT FeesID, Amount FROM Fees WHERE Class = ? AND Term = ?", (class_id, term_id))
                fee_row = c.fetchone()
                if fee_row:
                    base_fee = fee_row['Amount']
                    expected = base_fee * (pay_rate / 100.0)
                    
                    # Pay random amount up to expected
                    paid = round(random.uniform(0, expected), 2)
                    if paid > 0:
                        date = fake.date_between(start_date='-1y', end_date='today')
                        ref = fake.bothify(text='PAY-####-????')
                        c.execute("INSERT INTO FeesPayment (StudentTermID, FeesID, AmountPaid, TransactionDate, ReferenceNumber, DoneOn) VALUES (?, ?, ?, ?, ?, ?)",
                                  (student_term_id, fee_row['FeesID'], paid, date, ref, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

    # 7. Expenses
    print("Recording Expenses...")
    # Ensure Categories exist
    c.execute("INSERT OR IGNORE INTO Department (DeptName) VALUES (?)", ("Administration",))
    dept_id = c.lastrowid or c.execute("SELECT DeptID FROM Department WHERE DeptName = ?", ("Administration",)).fetchone()[0]
    
    c.execute("INSERT OR IGNORE INTO Category (Type, Department, CatName) VALUES (?, ?, ?)", ("Expenses", dept_id, "Utilities"))
    cat_id = c.lastrowid or c.execute("SELECT CatID FROM Category WHERE CatName = ?", ("Utilities",)).fetchone()[0]
    
    c.execute("INSERT OR IGNORE INTO SubCategory (CategoryID, SubCategoryName) VALUES (?, ?)", (cat_id, "Electricity"))
    sub_cat_id = c.lastrowid or c.execute("SELECT SubCatID FROM SubCategory WHERE SubCategoryName = ?", ("Electricity",)).fetchone()[0]

    for _ in range(20):
        amount = round(random.uniform(100, 5000), 2)
        date = fake.date_between(start_date='-1y', end_date='today')
        c.execute("INSERT INTO Expenses (ExpSubCategory, ExpAmount, ExpDate, ExpDescription) VALUES (?, ?, ?, ?)",
                  (sub_cat_id, amount, date, "Monthly Bill"))

    # 8. Supplies
    print("Adding Supplies...")
    c.execute("INSERT OR IGNORE INTO Category (Type, Department, CatName) VALUES (?, ?, ?)", ("Supplies", dept_id, "Office Supplies"))
    sup_cat_id = c.lastrowid or c.execute("SELECT CatID FROM Category WHERE CatName = ?", ("Office Supplies",)).fetchone()[0]
    
    c.execute("INSERT OR IGNORE INTO SubCategory (CategoryID, SubCategoryName) VALUES (?, ?)", (sup_cat_id, "Stationery"))
    sup_sub_cat_id = c.lastrowid or c.execute("SELECT SubCatID FROM SubCategory WHERE SubCategoryName = ?", ("Stationery",)).fetchone()[0]

    supplies_list = ["Pens", "Paper", "Markers", "Staplers"]
    for item in supplies_list:
        c.execute("SELECT SuppliesID FROM Supplies WHERE Name = ?", (item,))
        existing = c.fetchone()
        if not existing:
            c.execute("INSERT INTO Supplies (Name, Subcategory, Quantity, Unit) VALUES (?, ?, ?, ?)",
                      (item, sup_sub_cat_id, 0, "Box"))
            sup_id = c.lastrowid
        else:
            sup_id = existing[0]
        
        # Add Inventory (Purchase)
        for _ in range(3):
            qty = random.randint(5, 50)
            price = round(random.uniform(10, 100) * qty, 2)
            date = fake.date_between(start_date='-1y', end_date='today')
            c.execute("INSERT INTO SuppliesInventory (Supplies_ID, Type, Quantity, Price, Date) VALUES (?, ?, ?, ?, ?)",
                      (sup_id, 'In', qty, price, date))
            
            # Update stock
            c.execute("UPDATE Supplies SET Quantity = Quantity + ? WHERE SuppliesID = ?", (qty, sup_id))

    conn.commit()
    conn.close()
    print("Test data generation complete!")

if __name__ == "__main__":
    generate_data()
