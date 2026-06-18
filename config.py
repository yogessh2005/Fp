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
        "primary": "#1A5276",          # Deep elegant blue
        "secondary": "#2980B9",        # Bright blue
        "accent": "#27AE60",           # Vibrant green accent
        "success": "#2ECC71",          # Emerald green
        "warning": "#F39C12",          # Bright orange
        "danger": "#E74C3C",           # Crimson red
        "info": "#3498DB",             # Light blue
        "light": "#F8F9FA",            # App background
        "dark": "#2C3E50",             # Dark slate for text
        "white": "#FFFFFF",            # White for cards
        "black": "#17202A",            # Almost black
        "gray": "#7F8C8D",             # Soft gray
        "border": "#E5E7E9",           # Soft border
        "hover": "#154360",            # Hover state (darker primary)
        "gradient_start": "#1A5276",
        "gradient_end": "#2980B9",
        "alternate_row": "#F4F6F7",    # Subtle zebra striping
        "weekend": "#EAEDED",          
        "sunday": "#E5E7E9",           
        "email_highlight": "#F8F9FA",  
        "scheduled": "#D6EAF8",        
        "card_bg": "#FFFFFF",          
        "selected_row": "#D6EAF8",     
        "selected_border": "#2980B9",  
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
    
    LOOKUP_SHEETS = {
        "SK": "SK",
        "EVPL": "evpl",
        "Maximus": "maximus",
    }
   
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

