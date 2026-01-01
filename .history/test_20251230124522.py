def load_data(self):
    # ... existing code ...
    
    try:
        # Load the Excel file with proper handling
        self.data = pd.read_excel(self.excel_file)
        
        # Fix column names - replace spaces with underscores
        self.data.columns = [col.strip().replace(' ', '_') for col in self.data.columns]
        
        # Handle "Parent Contact2" -> "Parent_Contact2"
        if 'Parent_Contact2' in self.data.columns:
            # Already correct
            pass
        elif 'Parent_Contact_2' in self.data.columns:
            self.data.rename(columns={'Parent_Contact_2': 'Parent_Contact2'}, inplace=True)
        
        # Check mode
        self.mode = self.mode_group.checkedId()
        
        # For new students mode
        if self.mode == 1:
            # Define expected columns (StudentID is optional)
            expected_cols = ['Student_Name', 'FeePayer', 'PayRate', 'Parent_Name', 'Parent_Contact']
            
            # Check for missing required columns
            missing_cols = []
            for col in expected_cols:
                if col not in self.data.columns:
                    missing_cols.append(col)
            
            if missing_cols:
                raise ValueError(f"Missing required columns: {', '.join(missing_cols)}\n"
                                f"Found columns: {', '.join(self.data.columns)}")
            
            # === FIX 1: Convert columns to string BEFORE using .str accessor ===
            # First, ensure all expected string columns exist and are object type
            for col in expected_cols:
                if col in self.data.columns:
                    # Convert to string, handling NaN values
                    self.data[col] = self.data[col].astype(str)
            
            # Now we can safely use .str methods
            if 'FeePayer' in self.data.columns:
                # Convert to string first (already done above)
                # Then strip whitespace
                self.data['FeePayer'] = self.data['FeePayer'].str.strip()
                # Replace empty FeePayer with 'parents' as default
                self.data.loc[self.data['FeePayer'] == 'nan', 'FeePayer'] = 'parents'
                self.data.loc[self.data['FeePayer'] == '', 'FeePayer'] = 'parents'
            
            # Handle PayRate - allow 0 and convert properly
            if 'PayRate' in self.data.columns:
                # First convert to numeric, coerce errors to NaN
                self.data['PayRate'] = pd.to_numeric(self.data['PayRate'], errors='coerce')
                # Fill NaN with 0
                self.data['PayRate'] = self.data['PayRate'].fillna(0)
                # Convert to float
                self.data['PayRate'] = self.data['PayRate'].astype(float)
            
            # === FIX 2: Clean up string columns SAFELY ===
            string_columns = self.data.select_dtypes(include=['object']).columns
            for col in string_columns:
                # Ensure it's string type
                self.data[col] = self.data[col].astype(str)
                # Now we can use .str methods
                self.data[col] = self.data[col].str.strip()
                # Replace 'nan' string with empty string
                self.data.loc[self.data[col] == 'nan', col] = ''
            
            # Generate preview StudentIDs if needed
            if self.generate_ids and self.academic_year_id:
                with get_db_connection() as (conn, c):
                    c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (self.academic_year_id,))
                    year_result = c.fetchone()
                    if year_result:
                        year_name = year_result[0]
                        
                        # Get starting sequence
                        existing_max = self.get_existing_sequence(self.academic_year_id)
                        user_start = self.start_sequence_spin.value()
                        actual_start = max(existing_max + 1, user_start)
                        
                        # Generate IDs for preview
                        generated_ids = []
                        for i in range(len(self.data)):
                            student_id = f"{year_name}-{(actual_start + i):04d}"
                            generated_ids.append(student_id)
                        
                        # Add to DataFrame for display
                        self.data['StudentID'] = generated_ids
                        # Move StudentID to first column
                        cols = ['StudentID'] + [col for col in self.data.columns if col != 'StudentID']
                        self.data = self.data[cols]
            
            # If not generating IDs and StudentID column exists, clean it
            elif 'StudentID' in self.data.columns:
                # Ensure it's string type first
                self.data['StudentID'] = self.data['StudentID'].astype(str)
                # Then use .str methods
                self.data['StudentID'] = self.data['StudentID'].str.strip()
                # Replace 'nan' with empty string
                self.data.loc[self.data['StudentID'] == 'nan', 'StudentID'] = ''
                # Move StudentID to first column
                cols = ['StudentID'] + [col for col in self.data.columns if col != 'StudentID']
                self.data = self.data[cols]
        
        elif self.mode == 2:  # Enrollment only
            if 'StudentID' not in self.data.columns:
                raise ValueError("Missing 'StudentID' column for enrollment only.")
            
            # === FIX 3: Clean StudentID column safely ===
            # Convert to string first
            self.data['StudentID'] = self.data['StudentID'].astype(str)
            # Then strip
            self.data['StudentID'] = self.data['StudentID'].str.strip()
            # Replace 'nan' with empty
            self.data.loc[self.data['StudentID'] == 'nan', 'StudentID'] = ''
        
        else:
            raise ValueError("No mode selected.")
        
        # Display in table
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(self.data.columns)
        
        # Color code rows based on validation
        for idx, row in self.data.iterrows():
            items = []
            for col in self.data.columns:
                value = row.get(col, '')
                
                # Handle NaN values
                if pd.isna(value):
                    value = ''
                
                item = QStandardItem(str(value))
                
                # Highlight issues
                if self.mode == 1:
                    if col == 'Student_Name' and not str(value).strip():
                        item.setBackground(QColor(255, 243, 205))  # Warning yellow
                    elif col == 'Parent_Contact' and not str(value).strip():
                        item.setBackground(QColor(255, 243, 205))
                    elif col == 'FeePayer' and not str(value).strip():
                        item.setBackground(QColor(255, 243, 205))
                    elif col == 'PayRate':
                        try:
                            pay_rate = float(value) if value else 0
                            if pay_rate < 0:
                                item.setBackground(QColor(255, 220, 220))  # Light red for negative
                        except ValueError:
                            item.setBackground(QColor(255, 220, 220))
                
                items.append(item)
            
            model.appendRow(items)
        
        self.table_view.setModel(model)
        
        # Update summary
        row_count = len(self.data)
        issues = []
        
        if self.mode == 1:
            # Count missing required fields - use direct string comparison
            missing_names = (self.data['Student_Name'].astype(str).str.strip() == '').sum()
            missing_contacts = (self.data['Parent_Contact'].astype(str).str.strip() == '').sum()
            
            if missing_names > 0:
                issues.append(f"{missing_names} missing student names")
            if missing_contacts > 0:
                issues.append(f"{missing_contacts} missing parent contacts")
        
        summary = f"Previewing {row_count} records"
        if issues:
            summary += f" - Issues: {', '.join(issues)}"
        
        self.summary_label.setText(summary)
        
    except Exception as e:
        QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")
        import traceback
        traceback.print_exc()  # Print full error to console
        self.data = None
        self.summary_label.setText("Error loading data")






    def load_data(self):
        if not self.excel_file:
            QMessageBox.warning(self, "Error", "No Excel file selected.")
            return
        
        try:
            # Load the Excel file with proper handling
            self.data = pd.read_excel(self.excel_file)
            
            # Fix column names - replace spaces with underscores
            self.data.columns = [col.strip().replace(' ', '_') for col in self.data.columns]
            
            # Handle "Parent Contact2" -> "Parent_Contact2"
            if 'Parent_Contact2' in self.data.columns:
                # Already correct
                pass
            elif 'Parent_Contact_2' in self.data.columns:
                self.data.rename(columns={'Parent_Contact_2': 'Parent_Contact2'}, inplace=True)
            
            # Check mode
            self.mode = self.mode_group.checkedId()
            
            # For new students mode
            if self.mode == 1:
                # Define expected columns (StudentID is optional)
                expected_cols = ['Student_Name', 'FeePayer', 'PayRate', 'Parent_Name', 'Parent_Contact']
                
                # Check for missing required columns
                missing_cols = []
                for col in expected_cols:
                    if col not in self.data.columns:
                        missing_cols.append(col)
                
                if missing_cols:
                    raise ValueError(f"Missing required columns: {', '.join(missing_cols)}\n"
                                    f"Found columns: {', '.join(self.data.columns)}")
                
                # Fill empty FeePayer values with a default
                if 'FeePayer' in self.data.columns:
                    self.data['FeePayer'] = self.data['FeePayer'].fillna('').astype(str).str.strip()
                    # Replace empty FeePayer with 'parents' as default
                    self.data.loc[self.data['FeePayer'] == '', 'FeePayer'] = 'parents'
                
                # Handle PayRate - allow 0 and convert to string
                if 'PayRate' in self.data.columns:
                    self.data['PayRate'] = self.data['PayRate'].fillna(0).astype(str).str.strip()
                    # Ensure numeric values
                    self.data['PayRate'] = pd.to_numeric(self.data['PayRate'], errors='coerce').fillna(0)
                
                # Clean up all string columns
                for col in self.data.select_dtypes(include=['object']).columns:
                    self.data[col] = self.data[col].fillna('').astype(str).str.strip()
                
                # Generate preview StudentIDs if needed
                if self.generate_ids and self.academic_year_id:
                    with get_db_connection() as (conn, c):
                        c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (self.academic_year_id,))
                        year_result = c.fetchone()
                        if year_result:
                            year_name = year_result[0]
                            
                            # Get starting sequence
                            existing_max = self.get_existing_sequence(self.academic_year_id)
                            user_start = self.start_sequence_spin.value()
                            actual_start = max(existing_max + 1, user_start)
                            
                            # Generate IDs for preview
                            generated_ids = []
                            for i in range(len(self.data)):
                                student_id = f"{year_name}-{(actual_start + i):04d}"
                                generated_ids.append(student_id)
                            
                            # Add to DataFrame for display
                            self.data['StudentID'] = generated_ids
                            # Move StudentID to first column
                            cols = ['StudentID'] + [col for col in self.data.columns if col != 'StudentID']
                            self.data = self.data[cols]
                
                # If not generating IDs and StudentID column exists, clean it
                elif 'StudentID' in self.data.columns:
                    # Replace NaN with empty string
                    self.data['StudentID'] = self.data['StudentID'].fillna('').astype(str).str.strip()
                    # Move StudentID to first column
                    cols = ['StudentID'] + [col for col in self.data.columns if col != 'StudentID']
                    self.data = self.data[cols]
            
            elif self.mode == 2:  # Enrollment only
                if 'StudentID' not in self.data.columns:
                    raise ValueError("Missing 'StudentID' column for enrollment only.")
                
                # Clean StudentID column
                self.data['StudentID'] = self.data['StudentID'].fillna('').astype(str).str.strip()
            
            else:
                raise ValueError("No mode selected.")
            
            # Display in table
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(self.data.columns)
            
            # Color code rows based on validation
            for idx, row in self.data.iterrows():
                items = []
                for col in self.data.columns:
                    value = row.get(col, '')
                    
                    # Handle NaN values
                    if pd.isna(value):
                        value = ''
                    
                    item = QStandardItem(str(value))
                    
                    # Highlight issues
                    if self.mode == 1:
                        if col == 'Student_Name' and not str(value).strip():
                            item.setBackground(QColor(255, 243, 205))  # Warning yellow
                        elif col == 'Parent_Contact' and not str(value).strip():
                            item.setBackground(QColor(255, 243, 205))
                        elif col == 'FeePayer' and not str(value).strip():
                            item.setBackground(QColor(255, 243, 205))
                        elif col == 'PayRate':
                            try:
                                pay_rate = float(value) if value else 0
                                if pay_rate < 0:
                                    item.setBackground(QColor(255, 220, 220))  # Light red for negative
                            except ValueError:
                                item.setBackground(QColor(255, 220, 220))
                    
                    items.append(item)
                
                model.appendRow(items)
            
            self.table_view.setModel(model)
            
            # Update summary
            row_count = len(self.data)
            issues = []
            
            if self.mode == 1:
                # Count missing required fields
                missing_names = self.data['Student_Name'].str.strip().eq('').sum()
                missing_contacts = self.data['Parent_Contact'].str.strip().eq('').sum()
                
                if missing_names > 0:
                    issues.append(f"{missing_names} missing student names")
                if missing_contacts > 0:
                    issues.append(f"{missing_contacts} missing parent contacts")
            
            summary = f"Previewing {row_count} records"
            if issues:
                summary += f" - Issues: {', '.join(issues)}"
            
            self.summary_label.setText(summary)
            
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")
            import traceback
            traceback.print_exc()  # Print full error to console
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
            summary += "<b>Student IDs:</b> Will be taken from Excel file (if present)<br>"
        
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
                if self.mode == 1:
                    c.execute("SELECT Year FROM AcademicYear WHERE AcademicYearID = ?", (self.academic_year_id,))
                    year_result = c.fetchone()
                    if year_result:
                        year_name = year_result[0]
                
                # Get starting sequence for ID generation
                starting_sequence = self.start_sequence_spin.value()
                if self.mode == 1 and year_name:
                    # Get highest existing sequence
                    existing_max = self.get_existing_sequence(self.academic_year_id)
                    actual_start = max(existing_max + 1, starting_sequence)
                else:
                    actual_start = starting_sequence
                
                for index, row in self.data.iterrows():
                    excel_row = index + 2  # Excel row number (header is row 1)
                    try:
                        if self.mode == 1:  # New students with enrollment
                            # Generate or get StudentID
                            if self.generate_ids and year_name:
                                # Generate new ID
                                student_id = f"{year_name}-{actual_start:04d}"
                                actual_start += 1
                            else:
                                # Use ID from Excel or generate if missing
                                student_id = str(row.get('StudentID', '')).strip()
                                if not student_id and year_name:
                                    # Fallback generation
                                    student_id = f"{year_name}-{actual_start:04d}"
                                    actual_start += 1
                            
                            if not student_id:
                                raise ValueError("Could not generate StudentID")
                            
                            # Check if StudentID already exists
                            c.execute("SELECT StudentID FROM Students WHERE StudentID = ?", (student_id,))
                            if c.fetchone():
                                raise ValueError(f"StudentID '{student_id}' already exists")
                            
                            # Get and validate required fields
                            student_name = str(row.get('Student_Name', '')).strip()
                            fee_payer = str(row.get('FeePayer', '')).strip()
                            pay_rate = row.get('PayRate', 0)
                            parent_name = str(row.get('Parent_Name', '')).strip()
                            parent_contact = str(row.get('Parent_Contact', '')).strip()
                            parent_contact2 = str(row.get('Parent_Contact2', '')).strip()
                            
                            # Validate required fields
                            if not student_name:
                                raise ValueError("Student Name is required")
                            if not parent_name:
                                raise ValueError("Parent Name is required")
                            if not parent_contact:
                                raise ValueError("Parent Contact is required")
                            
                            # Handle FeePayer - if empty, use default
                            if not fee_payer:
                                fee_payer = "parents"
                            
                            # Handle PayRate - convert to float, allow 0
                            try:
                                pay_rate_float = float(pay_rate) if pay_rate else 0.0
                                if pay_rate_float < 0:
                                    raise ValueError("PayRate cannot be negative")
                            except ValueError:
                                raise ValueError("PayRate must be a valid number")
                            
                            # Insert into Students table
                            c.execute("""
                                INSERT INTO Students (StudentID, Student_Name, FeePayer, PayRate, 
                                                     Parent_Name, Parent_Contact, Parent_Contact2)
                                VALUES (?, ?, ?, ?, ?, ?, ?)
                            """, (
                                student_id,
                                student_name,
                                fee_payer,
                                pay_rate_float,
                                parent_name,
                                parent_contact,
                                parent_contact2 if parent_contact2 else None
                            ))
                            
                            # Insert into StudentTerm
                            c.execute("""
                                INSERT INTO StudentTerm (StudentID, ClassID, TermID, PayRate)
                                VALUES (?, ?, ?, ?)
                            """, (student_id, self.class_id, self.term_id, pay_rate_float))
                            
                            success_count += 1
                            
                        elif self.mode == 2:  # Enrollment only
                            student_id = str(row.get('StudentID', '')).strip()
                            if not student_id:
                                raise ValueError("StudentID is required")
                            
                            # Check if StudentID exists
                            c.execute("SELECT StudentID, Student_Name FROM Students WHERE StudentID = ?", (student_id,))
                            student_data = c.fetchone()
                            if not student_data:
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
                        error_messages.append(f"Row {excel_row}: {str(e)}")
                
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