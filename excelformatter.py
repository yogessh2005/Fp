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
        """Apply clean formatting without gridlines"""
       
        # Set header style
        header_font = Font(
            name='Segoe UI',
            size=11,
            bold=True,
            color=Config.EXCEL_HEADER_TEXT_COLOR
        )
        header_fill = PatternFill(
            start_color=Config.EXCEL_HEADER_COLOR,
            end_color=Config.EXCEL_HEADER_COLOR,
            fill_type="solid"
        )
        header_alignment = Alignment(
            horizontal='center',
            vertical='center',
            wrap_text=False
        )
       
        # Apply header formatting
        for col_idx, col_name in enumerate(df.columns, 1):
            cell = worksheet.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
       
        # Set value style
        value_font = Font(
            name='Segoe UI',
            size=10,
            color=Config.EXCEL_VALUE_COLOR
        )
        value_alignment = Alignment(
            horizontal='center',
            vertical='center'
        )
       
        # Alternative row color for values
        alternate_fill = PatternFill(
            start_color=Config.EXCEL_ALTERNATE_ROW_COLOR,
            end_color=Config.EXCEL_ALTERNATE_ROW_COLOR,
            fill_type="solid"
        )
       
        # Apply value formatting (no borders)
        for row_idx, row in enumerate(df.itertuples(index=False), 2):
            is_alternate = (row_idx - 2) % 2 == 1
            for col_idx in range(1, len(df.columns) + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell.font = value_font
                cell.alignment = value_alignment
                if is_alternate:
                    cell.fill = alternate_fill
       
        # Auto-adjust column widths
        for col_idx, col_name in enumerate(df.columns, 1):
            max_length = len(str(col_name)) + 4
            for row_idx in range(2, min(len(df) + 2, 100)):
                cell_value = worksheet.cell(row=row_idx, column=col_idx).value
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)) + 2)
            adjusted_width = min(max_length, 80)
            worksheet.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
       
        # ⭐ Remove all borders (no gridlines)
        for row in worksheet.iter_rows():
            for cell in row:
                cell.border = Border(
                    left=Side(border_style=None),
                    right=Side(border_style=None),
                    top=Side(border_style=None),
                    bottom=Side(border_style=None)
                )


