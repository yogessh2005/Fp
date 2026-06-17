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

# ==================== CONFIGURATION ====================

class Config:
    """Application configuration management"""
   
    # Database Configuration
    SERVER = "fpdbm1.focusprism.in"
    UID = "clarity"
    PWD = "Cla$7676"
   
    DATABASES = {
        "SK": "SK",
        "Maximus": "MAXIMUS",
        "EVPL": "EVPL",
    }
   
    # ⭐ EMAIL CONFIGURATION - HARDCODED IN BACKGROUND
    EMAIL_CONFIG = {
        "smtp_server": "smtp.gmail.com",
        "smtp_port": 587,
        "sender_email": "71762333054@cit.edu.in",
        "sender_password": "zdtd myuo zbnn vqrz",
        "use_tls": True,
    }
   
    # User Management
    USERS = {
        "sk_admin": {
            "password": hashlib.sha256("sk123".encode()).hexdigest(),
            "company": "SK",
            "role": "admin",
            "full_name": "SK Administrator"
        },
        "evpl_admin": {
            "password": hashlib.sha256("evpl123".encode()).hexdigest(),
            "company": "EVPL",
            "role": "admin",
            "full_name": "EVPL Administrator"
        },
        "maxi_admin": {
            "password": hashlib.sha256("maxi123".encode()).hexdigest(),
            "company": "Maximus",
            "role": "admin",
            "full_name": "Maximus Administrator"
        },
        "guest": {
            "password": hashlib.sha256("guest123".encode()).hexdigest(),
            "company": "SK",
            "role": "viewer",
            "full_name": "Guest User"
        }
    }
   
    # Application Settings
    APP_NAME = "Enterprise Database Portal"
    APP_VERSION = "3.2.2"
    COMPANY_LOGO = "📊"
   
    # Date Settings
    DEFAULT_DATE = datetime.now().strftime("%d%m%Y")
    DEFAULT_DATE_DISPLAY = datetime.now().strftime("%d/%m/%Y")
   
    # File Paths
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    LOOKUP_FILE = os.path.join(BASE_DIR, "manager_incharge_lookupss.xlsx")
    CONFIG_FILE = os.path.join(BASE_DIR, "app_config.json")
    LOG_FILE = os.path.join(BASE_DIR, "app_log.txt")
    SCHEDULE_FILE = os.path.join(BASE_DIR, "schedule_config.json")
    HISTORY_FILE = os.path.join(BASE_DIR, "email_history.json")
    TEMP_DIR = os.path.join(BASE_DIR, "temp")
    GRN_TARGET_FILE = os.path.join(BASE_DIR, "grn_target.json")
   
    # ⭐ Excel Formatting Colors
    EXCEL_HEADER_COLOR = "6B4F3B"  # Dark Brown
    EXCEL_HEADER_TEXT_COLOR = "FFFFFF"  # White
    EXCEL_VALUE_COLOR = "4A3628"  # Very Dark Brown
    EXCEL_ALTERNATE_ROW_COLOR = "F5F3EF"  # Very light beige for alternate rows
    
    # Theme Colors
    COLORS = {
        "primary": "#6B4F3B",          # Dark Brown (Text, primary buttons)
        "secondary": "#A38A75",        # Lighter Brown
        "accent": "#8C6A50",           # Medium Brown
        "success": "#556B2F",          # Dark Olive Green
        "warning": "#D2691E",          # Chocolate / Burnt Orange
        "danger": "#8B0000",           # Dark Red
        "info": "#4682B4",             # Steel Blue (Muted)
        "light": "#EBE7E0",            # Beige Background
        "dark": "#4A3628",             # Very Dark Brown
        "white": "#FFFFFF",            # White for cards
        "black": "#2C1E16",            # Off-black / deep brown
        "gray": "#8C8C8C",             # Gray
        "border": "#D3CEC4",           # Subtle Beige-Gray Border
        "hover": "#8C6A50",            # Hover state
        "gradient_start": "#6B4F3B",
        "gradient_end": "#4A3628",
        "alternate_row": "#F5F3EF",    # Very light beige
        "weekend": "#F0EAE1",          
        "sunday": "#E8DFD5",           
        "email_highlight": "#EBE7E0",  
        "scheduled": "#F5F3EF",        
        "card_bg": "#FFFFFF",          
        "selected_row": "#DED9D1",     
        "selected_border": "#6B4F3B",  
    }
   
    # Company-specific colors
    COMPANY_COLORS = {
        "SK": "#27AE60",
        "EVPL": "#3498DB",
        "Maximus": "#9B59B6",
    }
   
    # Module configuration
    MODULES = [
        {"name": "Store Data", "icon": "🏪", "description": "Store-wise GRN and stock analysis"},
        {"name": "Overall Estimate", "icon": "📊", "description": "Overall performance metrics"},
        {"name": "Sessionwise Estimate", "icon": "⏰", "description": "Session-wise breakdown"},
        {"name": "Actual Sales", "icon": "💰", "description": "Sales and revenue analysis"},
        {"name": "Score", "icon": "🏆", "description": "Performance scoring system"}
    ]
   
    # ⭐ Scoring Rules - Updated with new PSERM logic
    SCORING_RULES = [
        ["📋 SCORING RULES (Max 100 Points)", "", ""],
        ["", "", ""],
        ["1️⃣ GRN (25)", "Condition", "Points"],
        ["Today GRN > Target", "25"],
        ["Today GRN == Target", "20 (Grace)"],
        ["Today GRN < Target", "0"],
        ["", "", ""],
        ["2️⃣ Closing Stock (25)", "PSERM Count", "Points"],
        ["> 35", "25"],
        ["30 - 35", "20"],
        ["25 - 29", "15"],
        ["20 - 24", "10"],
        ["< 20", "0"],
        ["", "", ""],
        ["3️⃣ Estimate Overall (15)", "Difference", "Points"],
        ["-5 to +5", "15"],
        ["-10 to -6 or +6 to +10", "10"],
        ["-15 to -11 or +11 to +15", "5"],
        ["Beyond ±15", "2"],
        ["", "", ""],
        ["4️⃣ Session Wise (15)", "Difference", "Points"],
        ["-5 to +5", "15"],
        ["-10 to -6 or +6 to +10", "10"],
        ["-15 to -11 or +11 to +15", "5"],
        ["Beyond ±15", "2"],
        ["", "", ""],
        ["5️⃣ Sales (20)", "Sales Value", "Points"],
        ["Sales > 0", "20"],
        ["Sales = 0", "0"],
        ["", "", ""],
        ["6️⃣ TOTAL SCORE (100)", "Formula", ""],
        ["Total =", "GRN + Closing Stock + Estimate Overall + Session Wise + Sales"],
        ["Maximum =", "100 Points"],
    ]


LOOKUP_SHEETS = {
    "SK": "SK",
    "EVPL": "evpl",
    "Maximus": "maximus",
}


