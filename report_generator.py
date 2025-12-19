import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from datetime import datetime

class ReportGenerator:
    def __init__(self, school_name="Pinnacle School"):
        self.school_name = school_name

    def generate_pdf(self, filepath, term_name, start_date, end_date, fees_summary, fees_breakdown, supplies_total, supplies_breakdown, supplies_details, expenses_total, expenses_breakdown, expenses_details):
        doc = SimpleDocTemplate(filepath, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        # Styles
        title_style = styles['Title']
        heading_style = styles['Heading2']
        normal_style = styles['Normal']
        
        # --- Title Page / Header ---
        story.append(Paragraph(f"{self.school_name} - Financial Report", title_style))
        story.append(Spacer(1, 12))
        story.append(Paragraph(f"<b>Term:</b> {term_name}", normal_style))
        story.append(Paragraph(f"<b>Period:</b> {start_date} to {end_date}", normal_style))
        story.append(Spacer(1, 24))

        # --- 1. FEES SECTION ---
        story.append(Paragraph("1. Fees Overview", heading_style))
        
        # Explanation
        collected_pct = (fees_summary['collected'] / fees_summary['expected'] * 100) if fees_summary['expected'] > 0 else 0
        outstanding_pct = (fees_summary['outstanding'] / fees_summary['expected'] * 100) if fees_summary['expected'] > 0 else 0
        
        explanation = f"Out of the expected fees of <b>{fees_summary['expected']:,.2f}</b> for this term, " \
                      f"<b>{fees_summary['collected']:,.2f} ({collected_pct:.1f}%)</b> has been collected, " \
                      f"leaving <b>{fees_summary['outstanding']:,.2f} ({outstanding_pct:.1f}%)</b> outstanding."
        story.append(Paragraph(explanation, normal_style))
        story.append(Spacer(1, 12))

        # Fees Breakdown Table
        if fees_breakdown:
            data = [['Class', 'Expected', 'Collected', 'Outstanding', '% Out']]
            for item in fees_breakdown:
                data.append([
                    item['class_name'],
                    f"{item['expected']:,.2f}", 
                    f"{item['collected']:,.2f}",
                    f"{item['balance']:,.2f}",
                    f"{item['percentage']:.1f}%"
                ])
            
            t = Table(data, colWidths=[1.5*inch, 1.5*inch, 1.5*inch, 1.5*inch, 1*inch])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            story.append(t)
        story.append(Spacer(1, 24))

        # --- 2. SUPPLIES SECTION ---
        story.append(Paragraph("2. Supplies Spending", heading_style))
        
        # Summary & Explanation
        story.append(Paragraph(f"<b>Total Spent on Supplies:</b> {supplies_total:,.2f}", normal_style))
        
        if supplies_breakdown:
            top_cat = supplies_breakdown[0]
            explanation = f"Most supply spending this term was on <b>{top_cat[0]} ({top_cat[2]:.1f}%)</b>."
            if len(supplies_breakdown) > 1:
                second_cat = supplies_breakdown[1]
                explanation += f" Followed by <b>{second_cat[0]} ({second_cat[2]:.1f}%)</b>."
            story.append(Paragraph(explanation, normal_style))
        story.append(Spacer(1, 12))

        # Supplies Breakdown Table
        if supplies_breakdown:
            data = [['Category', 'Amount', '% of Total']]
            for item in supplies_breakdown:
                data.append([
                    item[0],
                    f"{item[1]:,.2f}",
                    f"{item[2]:.1f}%"
                ])
            
            t = Table(data, colWidths=[3*inch, 2*inch, 1.5*inch])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            story.append(t)
        story.append(Spacer(1, 24))

        # --- 3. EXPENSES SECTION ---
        story.append(Paragraph("3. Operational Expenses", heading_style))
        
        # Summary & Explanation
        story.append(Paragraph(f"<b>Total Expenses:</b> {expenses_total:,.2f}", normal_style))
        
        if expenses_breakdown:
            top_cat = expenses_breakdown[0]
            explanation = f"Operational expenses were mainly driven by <b>{top_cat[0]}</b>, which accounted for <b>{top_cat[2]:.1f}%</b> of total expenses."
            story.append(Paragraph(explanation, normal_style))
        story.append(Spacer(1, 12))

        # Expenses Breakdown Table
        if expenses_breakdown:
            data = [['Category', 'Amount', '% of Total']]
            for item in expenses_breakdown:
                data.append([
                    item[0],
                    f"{item[1]:,.2f}",
                    f"{item[2]:.1f}%"
                ])
            
            t = Table(data, colWidths=[3*inch, 2*inch, 1.5*inch])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            story.append(t)
        
        # Disclaimer Footer
        story.append(Spacer(1, 36))
        disclaimer_style = ParagraphStyle('Disclaimer', parent=styles['Normal'], textColor=colors.red, fontSize=9)
        story.append(Paragraph("Note: This report reflects only system-recorded data.", disclaimer_style))

        doc.build(story)

    def generate_excel(self, filepath, term_name, start_date, end_date, fees_summary, fees_breakdown, supplies_total, supplies_breakdown, supplies_details, expenses_total, expenses_breakdown, expenses_details):
        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            workbook = writer.book
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})
            
            # --- 1. Fees Summary Sheet ---
            fees_data = {
                'Metric': ['Expected Fees', 'Collected Fees', 'Outstanding Fees', 'Collection Rate'],
                'Value': [
                    fees_summary['expected'],
                    fees_summary['collected'],
                    fees_summary['outstanding'],
                    f"{(fees_summary['collected']/fees_summary['expected']*100) if fees_summary['expected'] > 0 else 0:.1f}%"
                ]
            }
            pd.DataFrame(fees_data).to_excel(writer, sheet_name='Fees Summary', index=False)
            
            if fees_breakdown:
                df_fees = pd.DataFrame(fees_breakdown)
                # Ensure columns exist before selecting
                cols = ['class_name', 'expected', 'collected', 'balance', 'percentage']
                # Map to nice names
                df_fees = df_fees[cols].rename(columns={
                    'class_name': 'Class',
                    'expected': 'Expected',
                    'collected': 'Collected',
                    'balance': 'Outstanding',
                    'percentage': '% Outstanding'
                })
                df_fees.to_excel(writer, sheet_name='Fees Summary', startrow=6, index=False)
                
            worksheet = writer.sheets['Fees Summary']
            worksheet.set_column('B:B', 20, money_fmt)

            # --- 2. Supplies Sheet ---
            if supplies_breakdown:
                df_sup_cat = pd.DataFrame(supplies_breakdown, columns=['Category', 'Amount', 'Percentage'])
                df_sup_cat.to_excel(writer, sheet_name='Supplies', startrow=0, index=False)
            
            if supplies_details:
                df_sup_det = pd.DataFrame(supplies_details)
                df_sup_det.to_excel(writer, sheet_name='Supplies', startrow=len(supplies_breakdown) + 5, index=False)
                
            worksheet = writer.sheets['Supplies']
            worksheet.set_column('B:B', 15, money_fmt)

            # --- 3. Expenses Sheet ---
            if expenses_breakdown:
                df_exp_cat = pd.DataFrame(expenses_breakdown, columns=['Category', 'Amount', 'Percentage'])
                df_exp_cat.to_excel(writer, sheet_name='Expenses', startrow=0, index=False)
            
            if expenses_details:
                df_exp_det = pd.DataFrame(expenses_details)
                df_exp_det.to_excel(writer, sheet_name='Expenses', startrow=len(expenses_breakdown) + 5, index=False)
                
            worksheet = writer.sheets['Expenses']
            worksheet.set_column('B:B', 15, money_fmt)
