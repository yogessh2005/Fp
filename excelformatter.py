import tkinter as tk
from tkinter import ttk, messagebox, font, scrolledtext, filedialog
import pymssql
import pandas as pd
import os
import sys
from datetime import datetime, timedelta
import threading
import queue
from functools import partial
import json
import hashlib
import logging
from typing import Dict, Any, Optional, List, Tuple
import warnings
import calendar as cal
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import time
import schedule
import atexit
import re
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore", category=UserWarning)

from config import Config

# ==================== EXCEL FORMATTER ====================

class ExcelFormatter:
    """Apply clean formatting to Excel sheets - No gridlines, colored headers"""
   
    @staticmethod
    def apply_clean_formatting(worksheet, df):
        """Apply highly professional formatting with highlights and filters"""
        
        # Turn off default Excel gridlines for a cleaner look
        worksheet.sheet_view.showGridLines = False
       
        # 1. Header Style (Deep primary color with bold white text)
        header_font = Font(
            name='Segoe UI', size=11, bold=True, color='FFFFFF'
        )
        # Use a mustard/tan color to match the requested scorecard style
        header_fill = PatternFill(
            start_color='D28E3B', end_color='D28E3B', fill_type="solid"
        )
        header_alignment = Alignment(
            horizontal='center', vertical='center', wrap_text=False
        )
        
        # Thick bottom border for header
        header_border = Border(
            bottom=Side(border_style='medium', color='1A252F'),
            left=Side(border_style=None),
            right=Side(border_style=None),
            top=Side(border_style=None)
        )
       
        # Apply header formatting
        for col_idx, col_name in enumerate(df.columns, 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = header_border
       
        # 2. Value Style and Zebra Striping
        value_font = Font(name='Segoe UI', size=10, color='333333')
        value_alignment = Alignment(horizontal='center', vertical='center')
        
        # Very subtle alternate row color
        alternate_fill = PatternFill(start_color='F9F9F9', end_color='F9F9F9', fill_type="solid")
        
        # Keyword highlight styles
        success_fill = PatternFill(start_color='E8F5E9', end_color='E8F5E9', fill_type='solid')
        success_font = Font(name='Segoe UI', size=10, color='2E7D32', bold=True)
        
        pending_fill = PatternFill(start_color='FFF3E0', end_color='FFF3E0', fill_type='solid')
        pending_font = Font(name='Segoe UI', size=10, color='E65100', bold=True)
        
        danger_fill = PatternFill(start_color='FFEBEE', end_color='FFEBEE', fill_type='solid')
        danger_font = Font(name='Segoe UI', size=10, color='C62828', bold=True)
        # Score specific highlight styles (soft/pastel colors)
        score_green_fill = PatternFill(start_color='E8F5E9', end_color='E8F5E9', fill_type='solid')
        score_green_font = Font(name='Segoe UI', size=10, color='2E7D32', bold=True)
        
        score_orange_fill = PatternFill(start_color='FFF3E0', end_color='FFF3E0', fill_type='solid')
        score_orange_font = Font(name='Segoe UI', size=10, color='E65100', bold=True)
        
        score_red_fill = PatternFill(start_color='FFEBEE', end_color='FFEBEE', fill_type='solid')
        score_red_font = Font(name='Segoe UI', size=10, color='C62828', bold=True)
       
        # Apply value formatting and keyword highlighting
        for row_idx, row in enumerate(df.itertuples(index=False), 2):
            is_alternate = (row_idx - 2) % 2 == 1
            for col_idx in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.font = value_font
                cell.alignment = value_alignment
                
                # Check for keyword highlighting
                val_str = str(cell.value).strip().lower() if cell.value else ""
                
                # Check if it's a score column
                col_name_lower = str(df.columns[col_idx - 1]).lower()
                is_score_column = "score" in col_name_lower or "mark" in col_name_lower
                
                score_processed = False
                if is_score_column and cell.value is not None:
                    try:
                        score_val = float(cell.value)
                        if 85 <= score_val <= 100:
                            cell.fill = score_green_fill
                            cell.font = score_green_font
                        elif 70 <= score_val < 85:
                            cell.fill = score_orange_fill
                            cell.font = score_orange_font
                        elif score_val < 70:
                            cell.fill = score_red_fill
                            cell.font = score_red_font
                        score_processed = True
                    except ValueError:
                        pass
                
                if not score_processed:
                    if val_str in ['completed', 'success', 'active', 'delivered', 'done']:
                        cell.fill = success_fill
                        cell.font = success_font
                    elif val_str in ['pending', 'processing', 'in progress', 'waiting']:
                        cell.fill = pending_fill
                        cell.font = pending_font
                    elif val_str in ['failed', 'error', 'cancelled', 'inactive']:
                        cell.fill = danger_fill
                        cell.font = danger_font
                    elif is_alternate:
                        cell.fill = alternate_fill
                    
                # Standard no borders for data rows
                if row_idx > 1:
                    cell.border = Border(
                        left=Side(border_style=None),
                        right=Side(border_style=None),
                        top=Side(border_style=None),
                        bottom=Side(border_style='hair', color='E0E0E0') # Soft bottom line instead of no grid
                    )
       
        # 3. Auto-adjust column widths
        for col_idx, col_name in enumerate(df.columns, 1):
            max_length = len(str(col_name)) + 4
            for row_idx in range(2, min(len(df) + 2, 100)):
                cell_value = worksheet.cell(row=row_idx, column=col_idx).value
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)) + 2)
            adjusted_width = min(max_length, 60) # Cap at 60 so it's not too wide
            worksheet.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
            
        # 4. Pro Enhancements: Freeze Panes and Auto Filter
        worksheet.freeze_panes = 'A2'
        
        # Add auto-filter to the header row
        max_col_letter = get_column_letter(len(df.columns))
        worksheet.auto_filter.ref = f"A1:{max_col_letter}{len(df) + 1}"


