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

# ==================== GRN TARGET MANAGER ====================

class GRNTargetManager:
    """Manage GRN target - stored in grn_target.json like LOOKUP_FILE"""
   
    @staticmethod
    def get_target():
        """Get the current GRN target from file"""
        try:
            if os.path.exists(Config.GRN_TARGET_FILE):
                with open(Config.GRN_TARGET_FILE, 'r') as f:
                    data = json.load(f)
                    return data.get('target', 4.0)
        except:
            pass
        return 4.0  # Default target if file doesn't exist
   
    @staticmethod
    def save_target(target: float):
        """Save the GRN target to file"""
        try:
            with open(Config.GRN_TARGET_FILE, 'w') as f:
                json.dump({'target': target}, f, indent=4)
            return True
        except Exception as e:
            return False
   
    @staticmethod
    def load_from_excel(filepath: str) -> Tuple[bool, str, float]:
        """Load GRN target from Excel file (background thread)"""
        try:
            df = pd.read_excel(filepath, header=None)
           
            # Get the first numeric value from the file
            for col in df.columns:
                for val in df[col]:
                    if pd.notna(val) and isinstance(val, (int, float)):
                        target = float(val)
                        GRNTargetManager.save_target(target)
                        return True, f"Target loaded: {target}", target
           
            return False, "No numeric value found in file", 0
        except Exception as e:
            return False, f"Error loading file: {str(e)}", 0


