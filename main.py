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
from grntargetmanager import GRNTargetManager
from loginpage import LoginPage

# ==================== APPLICATION ENTRY POINT ====================

if __name__ == "__main__":
    try:
        required_packages = ['pymssql', 'pandas', 'openpyxl', 'schedule']
        missing_packages = []
       
        for package in required_packages:
            try:
                __import__(package)
            except ImportError:
                missing_packages.append(package)
       
        if missing_packages:
            print(f"Missing required packages: {', '.join(missing_packages)}")
            print("Please install using: pip install " + ' '.join(missing_packages))
            sys.exit(1)
       
        print(f"Starting Email Scheduler Application...")
        print(f"Version: {Config.APP_VERSION}")
        print(f"Company: {list(Config.DATABASES.keys())}")
        print(f"Sender Email: {Config.EMAIL_CONFIG['sender_email']}")
        print(f"✅ GRN Target: {GRNTargetManager.get_target():.1f}")
        print(f"✅ GRN Scoring: > Target = 25, == Target = 20 (Grace), < Target = 0")
        print(f"✅ PSERM Scoring: > 35 = 25 marks, ≤ 35 = 0 marks")
        print(f"✅ Toggle Selection: Click PU to move to TOP, click again to restore")
        print(f"✅ Excel Formatting: No gridlines, colored headers & values")
        print(f"✅ Email Columns: Completely REMOVED from all views")
       
        app = LoginPage()
        app.mainloop()
       
    except Exception as e:
        print(f"Fatal error: {e}")
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")