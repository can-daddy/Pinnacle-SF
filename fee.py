class FeeTemplateWizard(QWizard):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Generate Fee Payment Template")
        self.setMinimumSize(600, 400)
        
        # Add pages
        self.addPage(self.create_selection_page())
        self.addPage(self.create_preview_page())
        
        self.selected_year = None
        self.selected_term = None
        self.selected_class = None
        self.enrollments_data = None
        
    def create_selection_page(self):
        page = QWizardPage()
        page.setTitle("Select Academic Year, Term and Class")
        
        layout = QVBoxLayout()
        
        # Academic Year selection
        year_layout = QHBoxLayout()
        year_label = QLabel("Academic Year:")
        self.year_combo = QComboBox()
        self.year_combo.currentIndexChanged.connect(self.on_year_changed)
        year_layout.addWidget(year_label)
        year_layout.addWidget(self.year_combo)
        layout.addLayout(year_layout)
        
        # Term selection
        term_layout = QHBoxLayout()
        term_label = QLabel("Term:")
        self.term_combo = QComboBox()
        term_layout.addWidget(term_label)
        term_layout.addWidget(self.term_combo)
        layout.addLayout(term_layout)
        
        # Class selection
        class_layout = QHBoxLayout()
        class_label = QLabel("Class:")
        self.class_combo = QComboBox()
        class_layout.addWidget(class_label)
        class_layout.addWidget(self.class_combo)
        layout.addLayout(class_layout)
        
        # Load initial data
        self.load_academic_years()
        
        page.setLayout(layout)
        return page
    
    def create_preview_page(self):
        page = QWizardPage()
        page.setTitle("Preview Enrollment Data")
        
        layout = QVBoxLayout()
        
        self.preview_label = QLabel("Select Year, Term and Class to see preview")
        self.preview_label.setWordWrap(True)
        
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(5)
        self.preview_table.setHorizontalHeaderLabels(["Student ID", "Name", "Class", "Term", "Amount Due"])
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        layout.addWidget(self.preview_label)
        layout.addWidget(self.preview_table)
        
        page.setLayout(layout)
        return page
    
    def load_academic_years(self):
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY Year")
                years = c.fetchall()
                
                self.year_combo.addItem("Select Academic Year", None)
                for year_id, year_name in years:
                    self.year_combo.addItem(year_name, year_id)
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load academic years: {str(e)}")
    
    def on_year_changed(self):
        year_id = self.year_combo.currentData()
        if not year_id:
            return
            
        # Load terms for selected year
        self.term_combo.clear()
        self.class_combo.clear()
        
        try:
            with get_db_connection() as (conn, c):
                # Load terms
                c.execute("SELECT TermID, TermName FROM Term WHERE AcademicYear = ? ORDER BY TermName", (year_id,))
                terms = c.fetchall()
                
                self.term_combo.addItem("Select Term", None)
                for term_id, term_name in terms:
                    self.term_combo.addItem(term_name, term_id)
                
                # Load classes
                c.execute("SELECT ClassID, ClassName FROM Class WHERE AcademicYear = ? ORDER BY ClassName", (year_id,))
                classes = c.fetchall()
                
                self.class_combo.addItem("Select Class", None)
                for class_id, class_name in classes:
                    self.class_combo.addItem(class_name, class_id)
                    
                # Connect signals
                self.term_combo.currentIndexChanged.connect(self.update_preview)
                self.class_combo.currentIndexChanged.connect(self.update_preview)
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load data: {str(e)}")
    
    def update_preview(self):
        year_id = self.year_combo.currentData()
        term_id = self.term_combo.currentData()
        class_id = self.class_combo.currentData()
        
        if not all([year_id, term_id, class_id]):
            return
            
        try:
            with get_db_connection() as (conn, c):
                query = """
                    SELECT st.StudentTermID, s.StudentID, s.Student_Name, 
                           c.ClassName, t.TermName, f.Amount, st.PayRate,
                           (f.Amount * (st.PayRate / 100)) as AmountDue,
                           COALESCE(SUM(fp.AmountPaid), 0) as AlreadyPaid,
                           (f.Amount * (st.PayRate / 100) - COALESCE(SUM(fp.AmountPaid), 0)) as RemainingAmount
                    FROM StudentTerm st
                    JOIN Students s ON st.StudentID = s.StudentID
                    JOIN Class c ON st.ClassID = c.ClassID
                    JOIN Term t ON st.TermID = t.TermID
                    JOIN Fees f ON f.Class = c.ClassID AND f.Term = t.TermID
                    LEFT JOIN FeesPayment fp ON st.StudentTermID = fp.StudentTermID
                    WHERE c.ClassID = ? AND t.TermID = ?
                    GROUP BY st.StudentTermID
                    HAVING RemainingAmount > 0
                """
                
                c.execute(query, (class_id, term_id))
                self.enrollments_data = c.fetchall()
                
                # Update preview
                self.preview_table.setRowCount(0)
                for row_data in self.enrollments_data:
                    row = self.preview_table.rowCount()
                    self.preview_table.insertRow(row)
                    
                    for col, data in enumerate(row_data[1:6]):  # StudentID, Name, Class, Term, AmountDue
                        item = QTableWidgetItem(str(data))
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        self.preview_table.setItem(row, col, item)
                
                count = len(self.enrollments_data)
                self.preview_label.setText(
                    f"Found {count} students with pending fees for {self.class_combo.currentText()} "
                    f"in {self.term_combo.currentText()} ({self.year_combo.currentText()})"
                )
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load preview: {str(e)}")
    
    def validateCurrentPage(self):
        if self.currentId() == 1:  # Preview page
            if not self.enrollments_data or len(self.enrollments_data) == 0:
                QMessageBox.warning(self, "No Data", "No students with pending fees found for the selected criteria.")
                return False
        return True
    
    def generate_template(self):
        """Generate the actual template file"""
        if not self.enrollments_data:
            return False
            
        try:
            # Create DataFrame for Excel (same as before)
            data = []
            for enrollment in self.enrollments_data:
                data.append({
                    'StudentTermID': enrollment[0],
                    'StudentID': enrollment[1],
                    'StudentName': enrollment[2],
                    'ClassName': enrollment[3],
                    'TermName': enrollment[4],
                    'OriginalAmount': enrollment[5],
                    'PayRate': enrollment[6],
                    'AmountDue': enrollment[7],
                    'AlreadyPaid': enrollment[8],
                    'RemainingAmount': enrollment[9],
                    'AmountToPay': '',
                    'TransactionDate': '',
                    'ReferenceNumber': ''
                })
            
            df = pd.DataFrame(data)
            
            # Save to Excel (same formatting as before)
            file_path, _ = QFileDialog.getSaveFileName(
                self, "Save Fee Payment Template", 
                f"Fee_Payments_{self.class_combo.currentText()}_{self.term_combo.currentText()}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                "Excel Files (*.xlsx)"
            )
            
            if file_path:
                # ... (same Excel writing code as before)
                return True
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate template: {str(e)}")
            return False
        
        return False

def generate_fee_payment_template(self):
    """Open wizard to generate fee payment template"""
    wizard = FeeTemplateWizard(self)
    if wizard.exec_() == QWizard.Accepted:
        # This will be called when wizard finishes successfully
        wizard.generate_template()

def import_bulk_payments(self):
    """Import bulk payments with optional filtering"""
    # First, let user select filters
    dialog = QDialog(self)
    dialog.setWindowTitle("Select Import Filters")
    dialog.setFixedSize(400, 300)
    
    layout = QVBoxLayout()
    
    # Year selection
    year_layout = QHBoxLayout()
    year_label = QLabel("Academic Year (optional):")
    year_combo = QComboBox()
    year_combo.addItem("All Years", None)
    # Load years...
    year_layout.addWidget(year_label)
    year_layout.addWidget(year_combo)
    
    # Term selection
    term_layout = QHBoxLayout()
    term_label = QLabel("Term (optional):")
    term_combo = QComboBox()
    term_combo.addItem("All Terms", None)
    term_layout.addWidget(term_label)
    term_layout.addWidget(term_combo)
    
    # Class selection
    class_layout = QHBoxLayout()
    class_label = QLabel("Class (optional):")
    class_combo = QComboBox()
    class_combo.addItem("All Classes", None)
    class_layout.addWidget(class_label)
    class_layout.addWidget(class_combo)
    
    layout.addLayout(year_layout)
    layout.addLayout(term_layout)
    layout.addLayout(class_layout)
    
    # Buttons
    button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
    button_box.accepted.connect(dialog.accept)
    button_box.rejected.connect(dialog.reject)
    layout.addWidget(button_box)
    
    dialog.setLayout(layout)
    
    if dialog.exec_() == QDialog.Accepted:
        selected_year = year_combo.currentData()
        selected_term = term_combo.currentData()
        selected_class = class_combo.currentData()
        
        # Now proceed with file selection and import
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Fee Payment File", "", "Excel Files (*.xlsx *.xls)"
        )
        
        if file_path:
            self.process_bulk_payments(file_path, selected_year, selected_term, selected_class)

def process_bulk_payments(self, file_path, filter_year=None, filter_term=None, filter_class=None):
    """Process bulk payments with optional filters"""
    try:
        df = pd.read_excel(file_path, sheet_name='Fee Payments')
        
        # Validate and process payments (similar to before)
        # Add filter validation in the processing loop
        
        for index, row in df.iterrows():
            # ... existing validation code ...
            
            # Apply filters if specified
            if filter_year or filter_term or filter_class:
                # Check if this payment matches the filters
                matches_filter = True
                
                # You'll need to join additional tables to check filters
                # This is a simplified example
                
            # Continue with payment processing...
            
    except Exception as e:
        QMessageBox.critical(self, "Error", f"Failed to import payments: {str(e)}")