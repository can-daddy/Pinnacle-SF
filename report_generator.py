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

    def generate_pdf(self, filepath, term_name, start_date, end_date, fees_data, expenses_total, supplies_total):
        doc = SimpleDocTemplate(filepath, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        # Title
        title_style = styles['Title']
        story.append(Paragraph(f"{self.school_name} - Financial Summary Report", title_style))
        story.append(Spacer(1, 12))

        # Term Info
        normal_style = styles['Normal']
        story.append(Paragraph(f"<b>Term:</b> {term_name}", normal_style))
        story.append(Paragraph(f"<b>Date Range:</b> {start_date} to {end_date}", normal_style))
        story.append(Spacer(1, 12))

        # Disclaimer
        disclaimer_style = ParagraphStyle(
            'Disclaimer',
            parent=styles['Normal'],
            textColor=colors.red,
            fontSize=10,
            leading=12
        )
        story.append(Paragraph("<b>IMPORTANT NOTE:</b> This report does not include donations, grants, rentals, or other external income sources. It reflects only the data recorded in this system.", disclaimer_style))
        story.append(Spacer(1, 24))

        # Summary Table
        data = [
            ['Category', 'Amount'],
            ['Fees Expected', f"{fees_data['expected']:,.2f}"],
            ['Fees Collected', f"{fees_data['collected']:,.2f}"],
            ['Outstanding Fees', f"{fees_data['outstanding']:,.2f}"],
            ['Total Expenses', f"{expenses_total:,.2f}"],
            ['Supplies Cost', f"{supplies_total:,.2f}"]
        ]

        t = Table(data, colWidths=[3*inch, 2*inch])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))
        story.append(t)
        story.append(Spacer(1, 24))

        # Explanation
        story.append(Paragraph("<b>Explanation:</b>", normal_style))
        story.append(Paragraph("This financial summary reflects only the data recorded in this system. Other school income sources are managed outside this report.", normal_style))
        
        doc.build(story)

    def generate_excel(self, filepath, term_name, start_date, end_date, fees_data, expenses_total, supplies_total, fees_breakdown, expenses_breakdown):
        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            # Summary Sheet
            summary_data = {
                'Metric': ['School Name', 'Term', 'Date Range', 'Fees Expected', 'Fees Collected', 'Outstanding Fees', 'Total Expenses', 'Supplies Cost', 'Note'],
                'Value': [
                    self.school_name,
                    term_name,
                    f"{start_date} to {end_date}",
                    fees_data['expected'],
                    fees_data['collected'],
                    fees_data['outstanding'],
                    expenses_total,
                    supplies_total,
                    "This report does not include donations, grants, rentals, or other external income sources."
                ]
            }
            df_summary = pd.DataFrame(summary_data)
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            # Formatting Summary
            workbook = writer.book
            worksheet = writer.sheets['Summary']
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})
            worksheet.set_column('B:B', 20, money_fmt)
            worksheet.set_column('A:A', 30)

            # Fees Breakdown Sheet
            if fees_breakdown:
                df_fees = pd.DataFrame(fees_breakdown)
                # Rename keys to nice headers if needed, assuming keys match
                df_fees.to_excel(writer, sheet_name='Fees Breakdown', index=False)
                worksheet_fees = writer.sheets['Fees Breakdown']
                worksheet_fees.set_column('B:D', 15, money_fmt)

            # Expenses Breakdown Sheet
            if expenses_breakdown:
                df_expenses = pd.DataFrame(expenses_breakdown, columns=['Category', 'Amount'])
                df_expenses.to_excel(writer, sheet_name='Expenses Breakdown', index=False)
                worksheet_exp = writer.sheets['Expenses Breakdown']
                worksheet_exp.set_column('B:B', 15, money_fmt)
