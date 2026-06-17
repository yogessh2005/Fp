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

from applogger import AppLogger
from config import Config

# ==================== EMAIL HISTORY ====================

class EmailHistory:
    """Track email sending history"""
   
    def __init__(self, logger: AppLogger):
        self.logger = logger
        self.history_file = Config.HISTORY_FILE
        self.history = []
        self.load_history()
   
    def load_history(self):
        try:
            if os.path.exists(self.history_file):
                with open(self.history_file, 'r') as f:
                    self.history = json.load(f)
        except Exception as e:
            self.logger.warning(f"Could not load email history: {e}")
            self.history = []
   
    def save_history(self):
        try:
            with open(self.history_file, 'w') as f:
                json.dump(self.history, f, indent=4)
        except Exception as e:
            self.logger.warning(f"Could not save email history: {e}")
   
    def add_record(self, company: str, recipients: List[str],
                   selected_date: str, status: str, message: str = "", trigger_time: str = ""):
        record = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "company": company,
            "recipients": recipients,
            "recipient_count": len(recipients),
            "selected_date": selected_date,
            "status": status,
            "message": message,
            "trigger_time": trigger_time
        }
        self.history.insert(0, record)
        if len(self.history) > 500:
            self.history = self.history[:500]
        self.save_history()
   
    def get_history(self):
        return self.history


