# Implementation Plan: Convert Action Buttons to ComboBox

## Overview
Replace all action buttons (Edit, Delete, Pay Fee) in tables with a single ComboBox dropdown that maintains the same functionality while improving UI consistency and cell fitting.

---

## Tables to be Updated

### 1. **StudentManagement > StudentsTab** ✅
- **Current Actions:** Edit (Green), Delete (Red)
- **New ComboBox Options:** [Select Action, Edit, Delete]
- **Cell Handling:** Single ComboBox in column 7
- **Code Location:** main.py, lines ~97-400

### 2. **StudentManagement > EnrollmentsTab** ✅
- **Current Actions:** Pay Fee (Blue), Edit (Green), Delete (Red)
- **New ComboBox Options:** [Select Action, Pay Fee, Edit, Delete]
- **Cell Handling:** Single ComboBox in column 6
- **Code Location:** main.py, lines ~400-900

### 3. **Class Page (Class class)** ✅
- **Current Actions:** Edit, Delete (typical pattern)
- **New ComboBox Options:** [Select Action, Edit, Delete]
- **Cell Handling:** Single ComboBox in Actions column
- **Code Location:** main.py, Class class definition

### 4. **Uniforms Page (Uniforms class)** ✅
- **Current Actions:** Edit, Delete (typical pattern)
- **New ComboBox Options:** [Select Action, Edit, Delete]
- **Cell Handling:** Single ComboBox in Actions column
- **Code Location:** main.py, Uniforms class definition

### 5. **Expenditure Page (Expenditure class)** ✅
- **Current Actions:** Edit, Delete (typical pattern)
- **New ComboBox Options:** [Select Action, Edit, Delete]
- **Cell Handling:** Single ComboBox in Actions column
- **Code Location:** main.py, Expenditure class definition

### 6. **Academic Calendar/Term (AcademicYearTerm class)** ✅
- **Current Actions:** Edit, Delete (typical pattern)
- **New ComboBox Options:** [Select Action, Edit, Delete]
- **Cell Handling:** Single ComboBox in Actions column
- **Code Location:** main.py, AcademicYearTerm class definition

### 7. **Fees Page (Fees class)** ✅
- **Exempt Tables:** Payment table in Payment tab (as requested)
- **Action Tables:** Student fees table, Class fees table
- **New ComboBox Options:** [Select Action, Edit, Delete]
- **Code Location:** main.py, Fees class definition

---

## Implementation Strategy

### A. Create Reusable Helper Function

Create a function to generate ComboBox actions for consistency:

```python
def create_action_combobox(row_id, actions_map):
    """
    Create a styled ComboBox for table actions
    
    Args:
        row_id: The unique identifier for the row (StudentID, etc.)
        actions_map: Dict mapping action names to callback functions
                    {'Edit': callback1, 'Delete': callback2, 'Pay Fee': callback3}
    
    Returns:
        QComboBox: Configured action ComboBox
    """
    combo = QComboBox()
    combo.addItem("Select Action")
    
    for action_name, callback in actions_map.items():
        combo.addItem(action_name)
    
    # Connect to handler
    combo.currentIndexChanged.connect(
        lambda idx: handle_action_selection(combo, row_id, actions_map)
    )
    
    return combo
```

### B. Create Action Handler Function

```python
def handle_action_selection(combo, row_id, actions_map):
    """Handle ComboBox action selection"""
    selected_action = combo.currentText()
    
    if selected_action == "Select Action":
        return
    
    # Get the callback for this action
    if selected_action in actions_map:
        actions_map[selected_action](row_id)
    
    # Reset ComboBox
    combo.setCurrentIndex(0)
```

### C. Cell Styling (from theme.qss)

```css
/* ComboBox in table cells */
QComboBox {
    background-color: #FFFFFF;
    color: #2D3748;
    border: 1px solid #E2E8F0;
    border-radius: 5px;
    padding: 4px 8px;
    min-height: 24px;
    font-size: 11px;
}

QComboBox:focus {
    border: 2px solid #1A365D;
}

QComboBox QAbstractItemView {
    background-color: #FFFFFF !important;
    color: #2D3748 !important;
    border: 1px solid #E2E8F0 !important;
    selection-background-color: #EBF8FF;
    selection-color: #1A365D;
}
```

---

## Code Changes Pattern

### Old Pattern (Buttons):
```python
actions_widget = QWidget()
actions_layout = QHBoxLayout(actions_widget)

edit_btn = QPushButton("Edit")
edit_btn.setObjectName("successButton")
edit_btn.clicked.connect(lambda: self.edit_student(row_id))

delete_btn = QPushButton("Delete")
delete_btn.setObjectName("dangerButton")
delete_btn.clicked.connect(lambda: self.delete_student(row_id))

actions_layout.addWidget(edit_btn)
actions_layout.addWidget(delete_btn)

self.table.setCellWidget(row, 7, actions_widget)
```

### New Pattern (ComboBox):
```python
combo = QComboBox()
combo.addItem("Select Action")
combo.addItem("Edit")
combo.addItem("Delete")

# Store row_id and callbacks
row_id = row_data[0]
actions_map = {
    'Edit': self.edit_student,
    'Delete': self.delete_student
}

combo.currentIndexChanged.connect(
    lambda idx: self.handle_action(combo, row_id, actions_map)
)

combo.setFixedHeight(28)
self.table.setCellWidget(row, 7, combo)
```

---

## Expected Benefits

✅ **UI Consistency:** Single standardized action widget across all tables
✅ **Space Efficiency:** ComboBox fits in one cell without overflow
✅ **Clean Design:** No visual clutter of multiple buttons
✅ **Mobile-Friendly:** Smaller footprint on smaller screens
✅ **Theme Integration:** Uses existing ComboBox styling from theme.qss
✅ **Maintainability:** Easier to add/remove actions

---

## Testing Checklist

- [ ] StudentsTab: Edit and Delete work
- [ ] EnrollmentsTab: Pay Fee, Edit, Delete all work
- [ ] Class: Edit and Delete work
- [ ] Uniforms: Edit and Delete work
- [ ] Expenditure: Edit and Delete work
- [ ] Academic Calendar: All actions work
- [ ] Fees: Action tables work, Payment table unchanged
- [ ] ComboBox styling matches theme
- [ ] Selecting "Select Action" doesn't trigger anything
- [ ] ComboBox resets after action
- [ ] All callbacks execute correctly

---

## Rollback Plan

If issues arise, the original button code is preserved in git history.
Simple revert: Replace ComboBox cell setup with button layout code.

---

## Priority Order

1. Helper functions (create once, use everywhere)
2. StudentManagement pages (highest impact, most used)
3. Class, Uniforms, Expenditure
4. Academic Calendar
5. Fees page (excluding Payment table)
6. Full testing
