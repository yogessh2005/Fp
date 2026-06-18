import os
from datetime import datetime
from typing import Tuple
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from config import Config

class PDFGenerator:
    """Module to generate PDF reports from DataFrames"""
    
    @staticmethod
    def generate_report(df: pd.DataFrame, filename: str, company: str, title: str = "Data Report", summary_data: dict = None) -> Tuple[bool, str]:
        """
        Generate a nicely formatted PDF report from a pandas DataFrame.
        """
        try:
            # We use landscape orientation to accommodate more columns
            doc = SimpleDocTemplate(filename, pagesize=landscape(letter), 
                                    rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
            elements = []
            styles = getSampleStyleSheet()
            
            # 1. Company Logo
            logo_path = os.path.join(Config.BASE_DIR, "logo.png")
            if not os.path.exists(logo_path):
                logo_path = os.path.join(Config.BASE_DIR, "logo.gif")
                
            if os.path.exists(logo_path):
                try:
                    img = Image(logo_path, width=50, height=50)
                    img.hAlign = 'LEFT'
                    elements.append(img)
                    elements.append(Spacer(1, 10))
                except Exception as e:
                    print(f"Could not add logo to PDF: {e}")
                    
            # 2. Report Title
            # Using primary color from Config
            primary_color = Config.COLORS.get("primary", "#6B4F3B")
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=18,
                spaceAfter=15,
                textColor=colors.HexColor(primary_color)
            )
            elements.append(Paragraph(f"{company.upper()} - {title}", title_style))
            
            # 3. Summary Section
            if summary_data:
                summary_style = ParagraphStyle(
                    'Summary',
                    parent=styles['Normal'],
                    fontSize=11,
                    spaceAfter=8,
                    textColor=colors.HexColor(Config.COLORS.get("dark", "#4A3628"))
                )
                
                # Title for summary
                elements.append(Paragraph("<b>Report Summary:</b>", summary_style))
                
                # Add each summary item
                for key, val in summary_data.items():
                    elements.append(Paragraph(f"<b>{key}:</b> {val}", summary_style))
                    
                elements.append(Spacer(1, 15))
                
            # 4. Tables
            if df is not None and not df.empty:
                # Convert DataFrame to list of lists including header
                data = [df.columns.tolist()] + df.values.tolist()
                
                # Format numbers safely to string to prevent ReportLab errors
                for i in range(1, len(data)):
                    for j in range(len(data[i])):
                        val = data[i][j]
                        if isinstance(val, float):
                            data[i][j] = f"{val:.2f}"
                        elif val is None or pd.isna(val):
                            data[i][j] = ""
                        else:
                            data[i][j] = str(val)
                
                table = Table(data, repeatRows=1)
                
                # Pull table colors from configuration to match the Excel theme
                header_bg = colors.HexColor("#" + Config.EXCEL_HEADER_COLOR)
                header_fg = colors.HexColor("#" + Config.EXCEL_HEADER_TEXT_COLOR)
                row_fg = colors.HexColor("#" + Config.EXCEL_VALUE_COLOR)
                alt_row_bg = colors.HexColor("#" + Config.EXCEL_ALTERNATE_ROW_COLOR)
                
                style = TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), header_bg),
                    ('TEXTCOLOR', (0, 0), (-1, 0), header_fg),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                    ('TEXTCOLOR', (0, 1), (-1, -1), row_fg),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 9),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey), # Thin grid lines for readability
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ])
                
                # Apply alternate row colors (zebra striping)
                for i in range(1, len(data)):
                    if i % 2 == 0:
                        style.add('BACKGROUND', (0, i), (-1, i), alt_row_bg)
                        
                table.setStyle(style)
                elements.append(table)
            else:
                elements.append(Paragraph("<i>No data available for this report.</i>", styles['Normal']))
                
            # 5. Footer with timestamp
            def add_footer(canvas, doc):
                canvas.saveState()
                footer_text = f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Page {doc.page}"
                canvas.setFont('Helvetica', 8)
                canvas.setFillColor(colors.gray)
                
                # Draw the footer text aligned to the right
                page_width = landscape(letter)[0]
                canvas.drawRightString(page_width - 30, 15, footer_text)
                
                # Add a subtle line above the footer
                canvas.setStrokeColor(colors.lightgrey)
                canvas.line(30, 25, page_width - 30, 25)
                canvas.restoreState()

            # Build document with elements and inject footer callback for every page
            doc.build(elements, onFirstPage=add_footer, onLaterPages=add_footer)
            return True, f"Successfully created PDF: {filename}"
            
        except Exception as e:
            return False, f"Error generating PDF: {str(e)}"
