class ImportWizard(QWizard):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Student Import Wizard")
        
        # Light theme stylesheet for a modern, web-like light tone
        light_style = """
        QWizard {
            background-color: #f8f9fa;
            color: #212529;
        }
        QLabel {
            color: #495057;
        }
        QPushButton {
            background-color: #007bff;
            color: white;
            border: none;
            padding: 6px 12px;
            border-radius: 4px;
        }
        QPushButton:hover {
            background-color: #0056b3;
        }
        QComboBox {
            background-color: white;
            border: 1px solid #ced4da;
            padding: 4px;
            border-radius: 4px;
        }
        QRadioButton {
            color: #495057;
        }
        QTableView {
            background-color: white;
            border: 1px solid #dee2e6;
            gridline-color: #dee2e6;
        }
        QCheckBox {
            color: #495057;
        }
        """
        self.setStyleSheet(light_style)
        
        # Initialize variables
        self.mode = None
        self.excel_file = None
        self.term_id = None
        self.class_id = None
        self.academic_year_id = None
        self.data = None
        self.generate_ids = False
        
        # Add pages
        self.intro_page_id = self.addPage(self.create_intro_page())
        self.file_page_id = self.addPage(self.create_file_page())
        self.config_page_id = self.addPage(self.create_config_page())
        self.id_generation_page_id = self.addPage(self.create_id_generation_page())
        self.preview_page_id = self.addPage(self.create_preview_page())
        self.finish_page_id = self.addPage(self.create_finish_page())
        
        # Connect page change signal
        self.currentIdChanged.connect(self.on_page_changed)

    def create_intro_page(self):
        page = QWizardPage()
        page.setTitle("Choose Import Mode")
        page.setSubTitle("Select how you want to import students")
        
        layout = QVBoxLayout()
        
        # Description label
        desc_label = QLabel("Choose the appropriate import method for your needs:")
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)
        
        # Mode selection
        self.mode_group = QButtonGroup()
        
        # Mode 1: New students with enrollment
        mode1_frame = QFrame()
        mode1_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        mode1_frame.setStyleSheet("""
            QFrame {
                background-color: #e8f4f8;
                border: 1px solid #b8daff;
                border-radius: 5px;
                padding: 10px;
                margin: 5px;
            }
        """)
        mode1_layout = QVBoxLayout(mode1_frame)
        radio1 = QRadioButton("Insert new students with initial term enrollment")
        radio1.setChecked(True)
        mode1_desc = QLabel("Use this for importing completely new students. The system will:\n"
                           "• Create new student records\n"
                           "• Generate Student IDs automatically\n"
                           "• Enroll students in the selected term and class")
        mode1_desc.setWordWrap(True)
        mode1_desc.setStyleSheet("color: #666; font-size: 12px; margin-left: 20px;")
        mode1_layout.addWidget(radio1)
        mode1_layout.addWidget(mode1_desc)
        
        # Mode 2: Enrollment only
        mode2_frame = QFrame()
        mode2_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        mode2_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 10px;
                margin: 5px;
            }
        """)
        mode2_layout = QVBoxLayout(mode2_frame)
        radio2 = QRadioButton("Enrollment only (existing students)")
        mode2_desc = QLabel("Use this for enrolling existing students in a new term/class.\n"
                           "• Students must already exist in the system\n"
                           "• You'll need their existing Student IDs\n"
                           "• Only creates enrollment records")
        mode2_desc.setWordWrap(True)
        mode2_desc.setStyleSheet("color: #666; font-size: 12px; margin-left: 20px;")
        mode2_layout.addWidget(radio2)
        mode2_layout.addWidget(mode2_desc)
        
        self.mode_group.addButton(radio1, 1)
        self.mode_group.addButton(radio2, 2)
        
        layout.addWidget(mode1_frame)
        layout.addWidget(mode2_frame)
        layout.addStretch()
        
        page.setLayout(layout)
        return page

    def create_file_page(self):
        page = QWizardPage()
        page.setTitle("Select Excel File")
        page.setSubTitle("Choose the Excel file containing student data")
        
        layout = QVBoxLayout()
        
        # File selection area
        file_frame = QFrame()
        file_frame.setFrameStyle(QFrame.Box | QFrame.Sunken)
        file_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 5px;
                padding: 15px;
            }
        """)
        file_layout = QVBoxLayout(file_frame)
        
        # File info
        self.file_label = QLabel("No file selected")
        self.file_label.setStyleSheet("font-style: italic; color: #6c757d;")
        
        # File format info
        format_label = QLabel("Expected Excel format:")
        format_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        
        format_info = QLabel("For new students: Student_Name, FeePayer, PayRate, Parent_Name, Parent_Contact\n"
                            "For enrollment only: StudentID")
        format_info.setStyleSheet("font-size: 12px; color: #666; background-color: #f8f9fa; padding: 5px; border-radius: 3px;")
        format_info.setWordWrap(True)
        
        # Browse button
        btn = QPushButton("📁 Browse for Excel File")
        btn.clicked.connect(self.browse_file)
        btn.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
        
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(format_label)
        file_layout.addWidget(format_info)
        file_layout.addWidget(btn)
        
        layout.addWidget(file_frame)
        layout.addStretch()
        
        page.setLayout(layout)
        return page

    def browse_file(self):
        file, _ = QFileDialog.getOpenFileName(
            self, 
            "Select Excel File", 
            "", 
            "Excel Files (*.xlsx *.xls);;All Files (*.*)"
        )
        if file:
            self.excel_file = file
            self.file_label.setText(f"Selected: {os.path.basename(file)}")
            self.file_label.setStyleSheet("color: #28a745; font-weight: bold;")
            
            # Try to load and validate the file
            try:
                # Quick validation
                df = pd.read_excel(file, nrows=5)  # Read only first 5 rows for validation
                if df.empty:
                    raise ValueError("Excel file is empty")
                
                self.validate_excel_columns(df.columns.tolist())
                
            except Exception as e:
                QMessageBox.warning(self, "File Error", f"Could not read Excel file:\n{str(e)}")
                self.excel_file = None
                self.file_label.setText("Invalid file - please select again")
                self.file_label.setStyleSheet("color: #dc3545; font-style: italic;")

    def validate_excel_columns(self, columns):
        """Validate that Excel file has required columns based on mode"""
        if not hasattr(self, 'mode_group') or not self.mode_group.checkedButton():
            return True
            
        mode = self.mode_group.checkedId()
        
        if mode == 1:  # New students
            required = ['Student_Name', 'FeePayer', 'PayRate', 'Parent_Name', 'Parent_Contact']
            missing = [col for col in required if col not in columns]
            if missing:
                QMessageBox.warning(
                    self, 
                    "Missing Columns",
                    f"For new student import, these columns are required but missing:\n"
                    f"{', '.join(missing)}\n\n"
                    f"Found columns: {', '.join(columns)}"
                )
                return False
        elif mode == 2:  # Enrollment only
            if 'StudentID' not in columns:
                QMessageBox.warning(
                    self,
                    "Missing Column",
                    f"For enrollment only, 'StudentID' column is required but not found.\n\n"
                    f"Found columns: {', '.join(columns)}"
                )
                return False
        return True

    def create_config_page(self):
        page = QWizardPage()
        page.setTitle("Configure Enrollment")
        page.setSubTitle("Select the academic year, term, and class for enrollment")
        
        layout = QVBoxLayout()
        
        # Configuration frame
        config_frame = QFrame()
        config_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        config_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 15px;
            }
        """)
        config_layout = QGridLayout(config_frame)
        config_layout.setSpacing(10)
        
        # Academic Year selection
        year_label = QLabel("Academic Year:")
        year_label.setStyleSheet("font-weight: bold;")
        self.academic_combo = QComboBox()
        self.academic_combo.setMinimumWidth(200)
        
        try:
            with get_db_connection() as (conn, c):
                c.execute("SELECT AcademicYearID, Year FROM AcademicYear ORDER BY Year DESC")
                years = c.fetchall()
                self.academic_combo.addItem("Select Academic Year", None)
                for yid, yname in years:
                    self.academic_combo.addItem(yname, yid)
        except sqlite3.Error as e:
            QMessageBox.warning(self, "Database Error", f"Failed to load academic years: {str(e)}")
        
        self.academic_combo.currentIndexChanged.connect(self.load_terms_and_classes)
        
        config_layout.addWidget(year_label, 0, 0)
        config_layout.addWidget(self.academic_combo, 0, 1)
        
        # Term selection
        term_label = QLabel("Term:")
        term_label.setStyleSheet("font-weight: bold;")
        self.term_combo = QComboBox()
        self.term_combo.setMinimumWidth(200)
        self.term_combo.addItem("Select Term", None)
        config_layout.addWidget(term_label, 1, 0)
        config_layout.addWidget(self.term_combo, 1, 1)
        
        # Class selection
        class_label = QLabel("Class:")
        class_label.setStyleSheet("font-weight: bold;")
        self.class_combo = QComboBox()
        self.class_combo.setMinimumWidth(200)
        self.class_combo.addItem("Select Class", None)
        config_layout.addWidget(class_label, 2, 0)
        config_layout.addWidget(self.class_combo, 2, 1)
        
        layout.addWidget(config_frame)
        layout.addStretch()
        
        page.setLayout(layout)
        return page

    def load_terms_and_classes(self):
        year_id = self.academic_combo.currentData()
        self.academic_year_id = year_id
        
        self.term_combo.clear()
        self.class_combo.clear()
        
        if year_id is None:
            self.term_combo.addItem("Select Term", None)
            self.class_combo.addItem("Select Class", None)
            return
        
        try:
            with get_db_connection() as (conn, c):
                # Load terms for selected year
                c.execute("SELECT TermID, TermName FROM Term WHERE AcademicYear = ? ORDER BY TermName", (year_id,))
                terms = c.fetchall()
                self.term_combo.addItem("Select Term", None)
                for tid, tname in terms:
                    self.term_combo.addItem(tname, tid)
                
                # Load classes for selected year
                c.execute("SELECT ClassID, ClassName FROM Class WHERE AcademicYear = ? ORDER BY ClassName", (year_id,))
                classes = c.fetchall()
                self.class_combo.addItem("Select Class", None)
                for cid, cname in classes:
                    self.class_combo.addItem(cname, cid)
        except sqlite3.Error as e:
            QMessageBox.warning(self, "Database Error", f"Failed to load terms/classes: {str(e)}")

    def create_id_generation_page(self):
        page = QWizardPage()
        page.setTitle("Student ID Generation")
        page.setSubTitle("Configure how Student IDs should be generated")
        
        layout = QVBoxLayout()
        
        # ID generation options
        options_frame = QFrame()
        options_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        options_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 5px;
                padding: 15px;
            }
        """)
        options_layout = QVBoxLayout(options_frame)
        
        # Checkbox for ID generation
        self.generate_ids_checkbox = QCheckBox("Generate Student IDs automatically")
        self.generate_ids_checkbox.setChecked(True)
        self.generate_ids_checkbox.stateChanged.connect(self.on_generate_ids_changed)
        
        # Format explanation
        format_label = QLabel("Format: AcademicYear-0001 (e.g., 2024-0001)")
        format_label.setStyleSheet("color: #666; font-size: 12px; margin-left: 20px;")
        
        # Last ID preview
        self.last_id_label = QLabel("Last generated ID will be displayed here")
        self.last_id_label.setStyleSheet("""
            QLabel {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 3px;
                padding: 10px;
                margin-top: 10px;
                font-family: monospace;
            }
        """)
        
        # Manual ID checkbox
        self.manual_id_checkbox = QCheckBox("Use StudentID column from Excel (if present)")
        self.manual_id_checkbox.setEnabled(False)
        
        options_layout.addWidget(self.generate_ids_checkbox)
        options_layout.addWidget(format_label)
        options_layout.addWidget(self.last_id_label)
        options_layout.addWidget(self.manual_id_checkbox)
        
        layout.addWidget(options_frame)
        
        # Preview area for ID generation
        preview_label = QLabel("Preview of first few generated IDs:")
        preview_label.setStyleSheet("font-weight: bold; margin-top: 15px;")
        layout.addWidget(preview_label)
        
        self.id_preview_table = QTableWidget(5, 2)
        self.id_preview_table.setHorizontalHeaderLabels(["Row", "Generated StudentID"])
        self.id_preview_table.horizontalHeader().setStretchLastSection(True)
        self.id_preview_table.setMaximumHeight(150)
        layout.addWidget(self.id_preview_table)
        
        layout.addStretch()
        
        page.setLayout(layout)
        return page

    def on_generate_ids_changed(self, state):
        self.generate_ids = (state == Qt.Checked)
        self.manual_id_checkbox.setEnabled(not self.generate_ids)
        
        # Update preview
        self.update_id_preview()

    def generate_student_id(self, year_name, sequence_number):
        """Generate StudentID in format: Year-0001"""
        # Clean year name (remove spaces, special chars)
        year_clean = str(year_name).strip().replace(' ', '').split('-')[0]
        
        # Format sequence number with leading zeros
        seq_formatted = f"{sequence_number:04d}"
        
        return f"{year_clean}-{seq_formatted}"

    def get_next_student_id(self, year_id):
        """Get the next available StudentID for the given academic year"""
        try:
            with get_db_connection() as (conn, c):
                # Get the year name
                c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (year_id,))
                year_result = c.fetchone()
                if not year_result:
                    return None
                
                year_name = year_result[0]
                
                # Get all existing StudentIDs for this year
                c.execute("""
                    SELECT StudentID FROM Students 
                    WHERE StudentID LIKE ?
                    ORDER BY StudentID DESC
                """, (f"{year_name}-%",))
                
                existing_ids = [row[0] for row in c.fetchall()]
                
                if not existing_ids:
                    return self.generate_student_id(year_name, 1)
                
                # Find the highest sequence number
                max_seq = 0
                for student_id in existing_ids:
                    try:
                        # Extract sequence number from format "Year-0001"
                        parts = student_id.split('-')
                        if len(parts) == 2 and parts[1].isdigit():
                            seq = int(parts[1])
                            max_seq = max(max_seq, seq)
                    except (ValueError, IndexError):
                        continue
                
                return self.generate_student_id(year_name, max_seq + 1)
                
        except sqlite3.Error as e:
            print(f"Error getting next StudentID: {str(e)}")
            return None

    def update_id_preview(self):
        """Update the ID preview based on current settings"""
        if not self.generate_ids or not self.academic_year_id:
            self.id_preview_table.clearContents()
            self.last_id_label.setText("ID generation requires academic year selection")
            return
        
        try:
            with get_db_connection() as (conn, c):
                # Get year name
                c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (self.academic_year_id,))
                year_result = c.fetchone()
                if not year_result:
                    return
                
                year_name = year_result[0]
                
                # Get last generated ID
                last_id = self.get_next_student_id(self.academic_year_id)
                if last_id:
                    self.last_id_label.setText(f"Next available ID: {last_id}")
                    
                    # Update preview table
                    self.id_preview_table.setRowCount(5)
                    for i in range(5):
                        # Extract sequence from last_id and calculate next ones
                        if '-' in last_id:
                            base, seq_str = last_id.split('-')
                            if seq_str.isdigit():
                                seq = int(seq_str) + i
                                preview_id = f"{base}-{seq:04d}"
                                
                                # Add to table
                                self.id_preview_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
                                self.id_preview_table.setItem(i, 1, QTableWidgetItem(preview_id))
        
        except sqlite3.Error as e:
            self.last_id_label.setText(f"Error: {str(e)}")

    def create_preview_page(self):
        page = QWizardPage()
        page.setTitle("Preview Data")
        page.setSubTitle("Review the data before importing")
        
        layout = QVBoxLayout()
        
        # Data preview table
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QTableView.SelectRows)
        self.table_view.horizontalHeader().setStretchLastSection(True)
        
        # Summary label
        self.summary_label = QLabel("No data loaded")
        self.summary_label.setStyleSheet("font-weight: bold; color: #6c757d;")
        
        layout.addWidget(self.summary_label)
        layout.addWidget(self.table_view)
        
        page.setLayout(layout)
        return page

    def create_finish_page(self):
        page = QWizardPage()
        page.setTitle("Ready to Import")
        page.setSubTitle("Click Finish to import the data")
        
        layout = QVBoxLayout()
        
        # Summary frame
        summary_frame = QFrame()
        summary_frame.setFrameStyle(QFrame.Box | QFrame.Raised)
        summary_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 2px solid #28a745;
                border-radius: 5px;
                padding: 15px;
            }
        """)
        summary_layout = QVBoxLayout(summary_frame)
        
        self.import_summary_label = QLabel("Import summary will appear here")
        self.import_summary_label.setWordWrap(True)
        self.import_summary_label.setStyleSheet("font-size: 14px;")
        
        summary_layout.addWidget(self.import_summary_label)
        layout.addWidget(summary_frame)
        
        # Warning/Info label
        self.final_warning_label = QLabel("⚠️ Review the preview carefully. This action cannot be undone.")
        self.final_warning_label.setStyleSheet("color: #856404; background-color: #fff3cd; padding: 10px; border-radius: 3px;")
        self.final_warning_label.setWordWrap(True)
        layout.addWidget(self.final_warning_label)
        
        layout.addStretch()
        
        page.setLayout(layout)
        return page

    def on_page_changed(self, page_id):
        """Handle page changes to update content"""
        if page_id == self.preview_page_id:
            self.load_data()
        elif page_id == self.finish_page_id:
            self.update_finish_summary()

    def load_data(self):
        if not self.excel_file:
            QMessageBox.warning(self, "Error", "No Excel file selected.")
            return
        
        try:
            # Load the Excel file
            self.data = pd.read_excel(self.excel_file)
            self.mode = self.mode_group.checkedId()
            
            # Validate based on mode
            if self.mode == 1:  # New students
                required_cols = ['Student_Name', 'FeePayer', 'PayRate', 'Parent_Name', 'Parent_Contact']
                missing_cols = [col for col in required_cols if col not in self.data.columns]
                
                if missing_cols:
                    raise ValueError(f"Missing required columns for new students: {', '.join(missing_cols)}")
                
                # If generating IDs, add StudentID column preview
                if self.generate_ids and self.academic_year_id:
                    with get_db_connection() as (conn, c):
                        c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (self.academic_year_id,))
                        year_result = c.fetchone()
                        if year_result:
                            year_name = year_result[0]
                            
                            # Get starting sequence
                            last_id = self.get_next_student_id(self.academic_year_id)
                            if last_id and '-' in last_id:
                                base, seq_str = last_id.split('-')
                                if seq_str.isdigit():
                                    start_seq = int(seq_str)
                                    
                                    # Generate IDs for preview
                                    generated_ids = []
                                    for i in range(len(self.data)):
                                        student_id = f"{base}-{(start_seq + i):04d}"
                                        generated_ids.append(student_id)
                                    
                                    # Add to DataFrame for display
                                    self.data['StudentID'] = generated_ids
            
            elif self.mode == 2:  # Enrollment only
                if 'StudentID' not in self.data.columns:
                    raise ValueError("Missing 'StudentID' column for enrollment only.")
            
            else:
                raise ValueError("No mode selected.")
            
            # Display in table
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(self.data.columns)
            
            # Color code rows based on validation
            for idx, row in self.data.iterrows():
                items = []
                for col in self.data.columns:
                    item = QStandardItem(str(row.get(col, '')))
                    
                    # Highlight missing required fields
                    if self.mode == 1 and col in ['Student_Name', 'Parent_Name', 'Parent_Contact']:
                        if pd.isnull(row.get(col)) or not str(row.get(col)).strip():
                            item.setBackground(QColor(255, 243, 205))  # Warning yellow
                    
                    items.append(item)
                
                model.appendRow(items)
            
            self.table_view.setModel(model)
            
            # Update summary
            row_count = len(self.data)
            self.summary_label.setText(f"Previewing {row_count} records. {row_count} rows will be imported.")
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")
            self.data = None
            self.summary_label.setText("Error loading data")

    def update_finish_summary(self):
        """Update the summary on the finish page"""
        if self.data is None or self.data.empty:
            self.import_summary_label.setText("No data to import. Please go back and load valid data.")
            return
        
        mode_text = "New Students with Enrollment" if self.mode == 1 else "Enrollment Only"
        year_text = self.academic_combo.currentText() if self.academic_combo.currentText() else "Not selected"
        term_text = self.term_combo.currentText() if self.term_combo.currentText() else "Not selected"
        class_text = self.class_combo.currentText() if self.class_combo.currentText() else "Not selected"
        
        summary = f"""
        <b>Import Summary:</b><br><br>
        <b>Mode:</b> {mode_text}<br>
        <b>Records to import:</b> {len(self.data)}<br>
        <b>Academic Year:</b> {year_text}<br>
        <b>Term:</b> {term_text}<br>
        <b>Class:</b> {class_text}<br><br>
        """
        
        if self.mode == 1 and self.generate_ids:
            summary += f"<b>Student IDs:</b> Will be generated automatically ({year_text.split('-')[0]}-####)<br>"
        elif self.mode == 1:
            summary += "<b>Student IDs:</b> Will be taken from Excel file<br>"
        
        self.import_summary_label.setText(summary)

    def accept(self):
        """Handle the import when Finish is clicked"""
        if self.data is None or self.data.empty:
            QMessageBox.warning(self, "Error", "No data to import.")
            return
        
        self.term_id = self.term_combo.currentData()
        self.class_id = self.class_combo.currentData()
        
        if not self.term_id or not self.class_id:
            QMessageBox.warning(self, "Error", "Term and Class must be selected.")
            return
        
        # Confirm import
        reply = QMessageBox.question(
            self,
            "Confirm Import",
            f"Are you sure you want to import {len(self.data)} record(s)?\n\n"
            f"This action cannot be undone.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        success_count = 0
        error_count = 0
        error_messages = []
        
        try:
            with get_db_connection() as (conn, c):
                # Get academic year name for ID generation
                year_name = None
                if self.generate_ids and self.academic_year_id:
                    c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (self.academic_year_id,))
                    year_result = c.fetchone()
                    if year_result:
                        year_name = year_result[0]
                
                # Get starting sequence for ID generation
                next_sequence = 1
                if self.generate_ids and year_name:
                    # Find the highest existing sequence for this year
                    c.execute("SELECT StudentID FROM Students WHERE StudentID LIKE ?", (f"{year_name}-%",))
                    existing_ids = [row[0] for row in c.fetchall()]
                    
                    max_seq = 0
                    for student_id in existing_ids:
                        try:
                            parts = student_id.split('-')
                            if len(parts) == 2 and parts[1].isdigit():
                                seq = int(parts[1])
                                max_seq = max(max_seq, seq)
                        except (ValueError, IndexError):
                            continue
                    
                    next_sequence = max_seq + 1
                
                for index, row in self.data.iterrows():
                    try:
                        if self.mode == 1:  # New students with enrollment
                            # Generate or get StudentID
                            if self.generate_ids and year_name:
                                # Generate new ID
                                student_id = f"{year_name}-{next_sequence:04d}"
                                next_sequence += 1
                            else:
                                # Use ID from Excel or generate if missing
                                student_id = str(row.get('StudentID', '')).strip()
                                if not student_id and year_name:
                                    # Fallback generation
                                    student_id = f"{year_name}-{next_sequence:04d}"
                                    next_sequence += 1
                            
                            if not student_id:
                                raise ValueError("Could not generate StudentID")
                            
                            # Check if StudentID already exists
                            c.execute("SELECT StudentID FROM Students WHERE StudentID = ?", (student_id,))
                            if c.fetchone():
                                raise ValueError(f"StudentID '{student_id}' already exists")
                            
                            # Validate required fields
                            required_fields = {
                                'Student_Name': str(row.get('Student_Name', '')).strip(),
                                'FeePayer': str(row.get('FeePayer', '')).strip(),
                                'PayRate': str(row.get('PayRate', '')).strip(),
                                'Parent_Name': str(row.get('Parent_Name', '')).strip(),
                                'Parent_Contact': str(row.get('Parent_Contact', '')).strip()
                            }
                            
                            for field_name, field_value in required_fields.items():
                                if not field_value:
                                    raise ValueError(f"Missing required field: {field_name}")
                            
                            # Insert into Students table
                            c.execute("""
                                INSERT INTO Students (StudentID, Student_Name, FeePayer, PayRate, 
                                                     Parent_Name, Parent_Contact, Parent_Contact2)
                                VALUES (?, ?, ?, ?, ?, ?, ?)
                            """, (
                                student_id,
                                required_fields['Student_Name'],
                                required_fields['FeePayer'],
                                required_fields['PayRate'],
                                required_fields['Parent_Name'],
                                required_fields['Parent_Contact'],
                                str(row.get('Parent_Contact2', '')).strip() or None
                            ))
                            
                            # Insert into StudentTerm
                            c.execute("""
                                INSERT INTO StudentTerm (StudentID, ClassID, TermID)
                                VALUES (?, ?, ?)
                            """, (student_id, self.class_id, self.term_id))
                        
                        elif self.mode == 2:  # Enrollment only
                            student_id = str(row.get('StudentID', '')).strip()
                            if not student_id:
                                raise ValueError("StudentID is required")
                            
                            # Check if StudentID exists
                            c.execute("SELECT StudentID FROM Students WHERE StudentID = ?", (student_id,))
                            if not c.fetchone():
                                raise ValueError(f"StudentID '{student_id}' does not exist")
                            
                            # Check if already enrolled
                            c.execute("""
                                SELECT StudentTermID FROM StudentTerm 
                                WHERE StudentID = ? AND ClassID = ? AND TermID = ?
                            """, (student_id, self.class_id, self.term_id))
                            if c.fetchone():
                                raise ValueError(f"Already enrolled in this term/class")
                            
                            # Insert enrollment
                            c.execute("""
                                INSERT INTO StudentTerm (StudentID, ClassID, TermID)
                                VALUES (?, ?, ?)
                            """, (student_id, self.class_id, self.term_id))
                        
                        success_count += 1
                        
                    except Exception as e:
                        error_count += 1
                        error_messages.append(f"Row {index + 2}: {str(e)}")
                
                # Commit if there were successful operations
                if success_count > 0:
                    conn.commit()
                else:
                    conn.rollback()
                    
        except sqlite3.Error as e:
            QMessageBox.critical(self, "Database Error", f"Failed to import data: {str(e)}")
            return
        
        # Show results
        if error_count > 0:
            error_details = "\n".join(error_messages[:10])
            if error_count > 10:
                error_details += f"\n... and {error_count - 10} more errors"
            
            result_text = f"""
            Import completed with errors:
            
            ✅ Successfully imported: {success_count} record(s)
            ❌ Failed to import: {error_count} record(s)
            
            First few errors:
            {error_details}
            """
            
            QMessageBox.warning(self, "Import Completed with Errors", result_text)
        else:
            QMessageBox.information(
                self, 
                "Import Successful", 
                f"✅ Successfully imported {success_count} record(s)"
            )
        
        super().accept()