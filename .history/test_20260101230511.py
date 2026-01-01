from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QTableWidget, QPushButton, QHBoxLayout, QWidget
from datetime import datetime
import rea


def export_table_to_file(table_widget, parent=None, title="Exported Table", id_column_pattern=None):
    """
    Export the given QTableWidget to PDF or Excel, with a file dialog for user choice.
    
    New behavior:
    - If ID column exists (matches pattern): Remove Actions column (old behavior)
    - If NO ID column exists: Replace Actions column with IDs
    
    :param table_widget: QTableWidget to export
    :param parent: parent widget for dialogs
    :param title: Title for PDF export
    :param id_column_pattern: Pattern to identify ID columns (e.g., "*ID", "StudentID", "FeesID")
                             Default: Looks for columns ending with "ID", "No", "Code", "Number"
    """
    # Ask user for export type and file path
    options = QFileDialog.Options()
    file_types = "PDF Files (*.pdf);;Excel Files (*.xlsx)"
    file_path, selected_filter = QFileDialog.getSaveFileName(
        parent, "Export Table", "", file_types, options=options
    )
    if not file_path:
        return

    # Determine export type
    if selected_filter.startswith("PDF"):
        if not file_path.lower().endswith(".pdf"):
            file_path += ".pdf"
        try:
            export_table_to_pdf(table_widget, file_path, title=title, id_column_pattern=id_column_pattern)
            QMessageBox.information(parent, "Export Successful", f"Table exported to {file_path}")
        except Exception as e:
            QMessageBox.warning(parent, "Export Failed", f"PDF export failed:\n{e}")
    elif selected_filter.startswith("Excel"):
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"
        try:
            export_table_to_excel(table_widget, file_path, id_column_pattern=id_column_pattern)
            QMessageBox.information(parent, "Export Successful", f"Table exported to {file_path}")
        except Exception as e:
            QMessageBox.warning(parent, "Export Failed", f"Excel export failed:\n{e}")


def export_table_to_excel(table_widget, file_path, id_column_pattern=None):
    """
    Export QTableWidget to Excel file.
    
    Logic:
    1. Detect ID columns based on pattern
    2. If ID column exists: Remove Actions column
    3. If NO ID column: Replace Actions column with row IDs/numbers
    """
    wb = Workbook()
    ws = wb.active
    
    # Get all headers
    headers = []
    for col in range(table_widget.columnCount()):
        header = table_widget.horizontalHeaderItem(col)
        headers.append(header.text() if header else f"Column {col+1}")
    
    # Detect ID columns
    id_columns = []
    if id_column_pattern:
        # Use provided pattern
        for idx, header in enumerate(headers):
            if id_column_pattern == "*ID" and header.endswith("ID"):
                id_columns.append(idx)
            elif header == id_column_pattern:
                id_columns.append(idx)
            elif id_column_pattern in header:
                id_columns.append(idx)
    else:
        # Auto-detect ID columns (ends with ID, No, Code, Number)
        id_patterns = ["ID$", "No$", "Code$", "Number$"]
        for idx, header in enumerate(headers):
            for pattern in id_patterns:
                if re.search(pattern, header, re.IGNORECASE):
                    id_columns.append(idx)
                    break
    
    # Check if Actions column exists (last column with widgets)
    has_actions = False
    actions_col = table_widget.columnCount() - 1
    if actions_col >= 0:
        # Check if any row has widget in last column (typical for action buttons)
        for row in range(table_widget.rowCount()):
            if table_widget.cellWidget(row, actions_col):
                has_actions = True
                break
    
    # Determine export strategy
    if id_columns:
        # ID column exists → Remove Actions column (old behavior)
        if has_actions:
            # Exclude the Actions column
            export_cols = list(range(table_widget.columnCount() - 1))
            export_headers = [headers[i] for i in export_cols]
        else:
            # No Actions column, export all
            export_cols = list(range(table_widget.columnCount()))
            export_headers = headers
    else:
        # NO ID column exists → Replace Actions with IDs
        if has_actions:
            # Replace Actions header with "ID"
            export_cols = list(range(table_widget.columnCount()))
            export_headers = headers.copy()
            export_headers[-1] = "ID"  # Replace Actions header with "ID"
        else:
            # No Actions column, add ID column
            export_cols = list(range(table_widget.columnCount()))
            export_headers = headers
            export_headers.append("ID")  # Add ID column
    
    # Write headers
    ws.append(export_headers)
    
    # Write data rows
    for row in range(table_widget.rowCount()):
        row_data = []
        
        # Add regular columns
        for col in export_cols:
            if col < table_widget.columnCount():  # Skip if col is beyond original columns
                item = table_widget.item(row, col)
                if item:
                    row_data.append(item.text())
                else:
                    cell_widget = table_widget.cellWidget(row, col)
                    if cell_widget:
                        # For action widgets, try to extract ID if available
                        if not id_columns and col == actions_col and has_actions:
                            # Replace Actions with row number (or extract ID from button if possible)
                            # Try to find ID from other columns first
                            found_id = None
                            # Look for any column that might contain ID-like data
                            for check_col in range(table_widget.columnCount()):
                                check_item = table_widget.item(row, check_col)
                                if check_item:
                                    text = check_item.text()
                                    # Check if text looks like an ID
                                    if (re.match(r'^\d+$', text) or  # All numbers
                                        re.match(r'^[A-Z]+-\d+$', text) or  # ABC-123 pattern
                                        re.match(r'^\d{4}-\d{4}$', text)):  # 2024-0001 pattern
                                        found_id = text
                                        break
                            
                            if found_id:
                                row_data.append(found_id)
                            else:
                                # Use row number as ID
                                row_data.append(f"Row-{row+1}")
                        else:
                            row_data.append("")
                    else:
                        row_data.append("")
            else:
                # This is for added ID column when no Actions existed
                # Use row number as ID
                row_data.append(f"Row-{row+1}")
        
        # If we need to add ID column (no Actions, no ID columns)
        if not has_actions and not id_columns:
            # Add row number as ID
            row_data.append(f"Row-{row+1}")
        
        ws.append(row_data)
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(file_path)


def export_table_to_pdf(table_widget, file_path, title="Exported Table", id_column_pattern=None):
    """
    Export QTableWidget to PDF file.
    
    Logic:
    1. Detect ID columns based on pattern
    2. If ID column exists: Remove Actions column
    3. If NO ID column: Replace Actions column with row IDs/numbers
    """
    # Get all headers
    headers = []
    for col in range(table_widget.columnCount()):
        header = table_widget.horizontalHeaderItem(col)
        headers.append(header.text() if header else f"Column {col+1}")
    
    # Detect ID columns
    id_columns = []
    if id_column_pattern:
        # Use provided pattern
        for idx, header in enumerate(headers):
            if id_column_pattern == "*ID" and header.endswith("ID"):
                id_columns.append(idx)
            elif header == id_column_pattern:
                id_columns.append(idx)
            elif id_column_pattern in header:
                id_columns.append(idx)
    else:
        # Auto-detect ID columns (ends with ID, No, Code, Number)
        id_patterns = ["ID$", "No$", "Code$", "Number$"]
        for idx, header in enumerate(headers):
            for pattern in id_patterns:
                if re.search(pattern, header, re.IGNORECASE):
                    id_columns.append(idx)
                    break
    
    # Check if Actions column exists (last column with widgets)
    has_actions = False
    actions_col = table_widget.columnCount() - 1
    if actions_col >= 0:
        for row in range(table_widget.rowCount()):
            if table_widget.cellWidget(row, actions_col):
                has_actions = True
                break
    
    # Prepare data for PDF
    data = []
    
    # Determine which columns to include and headers
    if id_columns:
        # ID column exists → Remove Actions column
        if has_actions:
            col_count = table_widget.columnCount() - 1  # Exclude Actions
            pdf_headers = [headers[i] for i in range(col_count)]
        else:
            col_count = table_widget.columnCount()
            pdf_headers = headers
    else:
        # NO ID column → Replace Actions with ID
        if has_actions:
            col_count = table_widget.columnCount()
            pdf_headers = headers.copy()
            pdf_headers[-1] = "ID"  # Replace Actions header
        else:
            col_count = table_widget.columnCount()
            pdf_headers = headers + ["ID"]  # Add ID column
    
    # Add headers to data
    data.append(pdf_headers)
    
    # Add data rows
    for row in range(table_widget.rowCount()):
        row_data = []
        
        # Add data from visible columns
        for col in range(col_count):
            if col < table_widget.columnCount():  # Regular column
                if col < table_widget.columnCount() - 1 or not has_actions:
                    # Regular cell
                    item = table_widget.item(row, col)
                    if item:
                        row_data.append(item.text())
                    else:
                        cell_widget = table_widget.cellWidget(row, col)
                        if cell_widget and not id_columns and col == actions_col and has_actions:
                            # This is Actions column with no ID columns
                            # Try to find ID from row data
                            found_id = None
                            for check_col in range(table_widget.columnCount()):
                                check_item = table_widget.item(row, check_col)
                                if check_item:
                                    text = check_item.text()
                                    if (re.match(r'^\d+$', text) or
                                        re.match(r'^[A-Z]+-\d+$', text) or
                                        re.match(r'^\d{4}-\d{4}$', text)):
                                        found_id = text
                                        break
                            
                            if found_id:
                                row_data.append(found_id)
                            else:
                                row_data.append(f"Row-{row+1}")
                        else:
                            row_data.append("")
                else:
                    # Actions column when no ID columns exist
                    if not id_columns and has_actions:
                        # Try to extract ID
                        found_id = None
                        for check_col in range(table_widget.columnCount() - 1):  # Exclude Actions column
                            check_item = table_widget.item(row, check_col)
                            if check_item:
                                text = check_item.text()
                                if (re.match(r'^\d+$', text) or
                                    re.match(r'^[A-Z]+-\d+$', text) or
                                    re.match(r'^\d{4}-\d{4}$', text)):
                                    found_id = text
                                    break
                        
                        if found_id:
                            row_data.append(found_id)
                        else:
                            row_data.append(f"Row-{row+1}")
                    else:
                        row_data.append("")
            else:
                # Added ID column (when no Actions existed)
                row_data.append(f"Row-{row+1}")
        
        # If we need to add ID column (no Actions, no ID columns)
        if not has_actions and not id_columns:
            row_data.append(f"Row-{row+1}")
        
        data.append(row_data)
    
    # Create PDF
    pdf = SimpleDocTemplate(file_path, pagesize=landscape(A4))
    
    # Create table
    table = Table(data, repeatRows=1)
    
    # Table style
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#007bff")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.whitesmoke]),
    ])
    
    # Add zebra striping
    for i in range(1, len(data)):
        if i % 2 == 0:
            bg_color = colors.white
        else:
            bg_color = colors.whitesmoke
        style.add('BACKGROUND', (0, i), (-1, i), bg_color)
    
    table.setStyle(style)
    
    # Title paragraph
    styles = getSampleStyleSheet()
    title_paragraph = Paragraph(f"<b>{title}</b>", styles['Title'])
    spacer = Spacer(1, 0.2 * inch)
    
    # Footer callback
    def footer(canvas, doc):
        canvas.saveState()
        footer_text = f"Printed by Pinnacle Ledger    {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(doc.pagesize[0] - 30, 20, footer_text)
        canvas.restoreState()
    
    elems = [title_paragraph, spacer, table]
    pdf.build(elems, onFirstPage=footer, onLaterPages=footer)


# Helper function to check if a table has ID columns
def has_id_columns(table_widget, id_column_pattern=None):
    """
    Check if table has ID columns.
    
    :return: Tuple (has_id, id_columns_list, has_actions)
    """
    headers = []
    for col in range(table_widget.columnCount()):
        header = table_widget.horizontalHeaderItem(col)
        headers.append(header.text() if header else "")
    
    # Detect ID columns
    id_columns = []
    if id_column_pattern:
        for idx, header in enumerate(headers):
            if id_column_pattern == "*ID" and header.endswith("ID"):
                id_columns.append(idx)
            elif header == id_column_pattern:
                id_columns.append(idx)
            elif id_column_pattern in header:
                id_columns.append(idx)
    else:
        id_patterns = ["ID$", "No$", "Code$", "Number$"]
        for idx, header in enumerate(headers):
            for pattern in id_patterns:
                if re.search(pattern, header, re.IGNORECASE):
                    id_columns.append(idx)
                    break
    
    # Check for Actions column
    has_actions = False
    actions_col = table_widget.columnCount() - 1
    if actions_col >= 0:
        for row in range(table_widget.rowCount()):
            if table_widget.cellWidget(row, actions_col):
                has_actions = True
                break
    
    return len(id_columns) > 0, id_columns, has_actions


# Usage examples
def example_usage():
    """
    Example usage of the export functions with different scenarios.
    """
    from PyQt5.QtWidgets import QApplication, QTableWidget, QTableWidgetItem
    import sys
    
    app = QApplication(sys.argv)
    
    # Create a sample table
    table = QTableWidget(5, 4)
    table.setHorizontalHeaderLabels(["Name", "Age", "StudentID", "Actions"])
    
    # Fill with data
    for i in range(5):
        table.setItem(i, 0, QTableWidgetItem(f"Student {i+1}"))
        table.setItem(i, 1, QTableWidgetItem(str(20 + i)))
        table.setItem(i, 2, QTableWidgetItem(f"STD-{1000 + i}"))
        # Actions column would typically have buttons
    
    # Example 1: Auto-detect (has StudentID column → removes Actions)
    # export_table_to_file(table, title="Students List")
    
    # Example 2: Specify ID pattern
    # export_table_to_file(table, title="Students", id_column_pattern="StudentID")
    
    # Example 3: Table without ID column (will replace Actions with IDs)
    table2 = QTableWidget(3, 3)
    table2.setHorizontalHeaderLabels(["Name", "Age", "Actions"])
    for i in range(3):
        table2.setItem(i, 0, QTableWidgetItem(f"Person {i+1}"))
        table2.setItem(i, 1, QTableWidgetItem(str(25 + i)))
    
    # This will replace Actions with Row-1, Row-2, etc.
    # export_table_to_file(table2, title="People without IDs")
    
    return app.exec_()


if __name__ == "__main__":
    example_usage()