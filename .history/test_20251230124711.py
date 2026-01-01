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

