import os
import re

def modularize_app():
    source_file = r"C:\Users\srini\Downloads\aaaa\bbbb.py"
    output_dir = r"C:\Users\srini\Downloads\app_code"

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    else:
        for f in os.listdir(output_dir):
            if f.endswith('.py') and f != "modularizer.py":
                try:
                    os.remove(os.path.join(output_dir, f))
                except: pass

    print(f"Reading source file: {source_file}")
    with open(source_file, 'r', encoding='utf-8') as f:
        content = f.read()

    # The standard imports that every file needs to function independently
    std_imports = """import tkinter as tk
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
"""

    # Split the file exactly by the structural headers
    sections = re.split(r'# ====================\s*(.*?)\s*====================\n', content)

    modules_to_create = []

    for i in range(1, len(sections), 2):
        header = sections[i].strip()
        code = sections[i+1]
        
        # Map the header to the correct filename
        if header == "CONFIGURATION": filename = "config.py"
        elif header == "GRN TARGET MANAGER": filename = "grntargetmanager.py"
        elif header == "LOGGING SETUP": filename = "applogger.py"
        elif header == "EXCEL FORMATTER": filename = "excelformatter.py"
        elif header == "EMAIL HISTORY": filename = "emailhistory.py"
        elif header == "EMAIL SCHEDULER": filename = "emailscheduler.py"
        elif header == "MODERN UI COMPONENTS": filename = "ui_components.py"
        elif header == "QUERY BUILDER": filename = "querybuilder.py"
        elif header == "DATABASE HANDLER": filename = "databasehandler.py"
        elif header == "EMAIL HISTORY DIALOG": filename = "emailhistorydialog.py"
        elif header == "TRIGGERS DIALOG": filename = "triggersdialog.py"
        elif header == "SCHEDULE EMAIL DIALOG": filename = "scheduleemaildialog.py"
        elif header == "MAIN APPLICATION": filename = "mainapp.py"
        elif header == "ADVANCED FILTER DIALOG": filename = "advancedfilterdialog.py"
        elif header == "LOGIN PAGE": filename = "loginpage.py"
        elif header == "APPLICATION ENTRY POINT": filename = "main.py"
        else: filename = header.lower().replace(" ", "_") + ".py"
        
        modules_to_create.append((filename, header, code))

    # A map of which class belongs to which generated module
    module_mapping = {
        "Config": "config",
        "GRNTargetManager": "grntargetmanager",
        "AppLogger": "applogger",
        "ExcelFormatter": "excelformatter",
        "EmailHistory": "emailhistory",
        "EmailScheduler": "emailscheduler",
        "ModernButton": "ui_components",
        "ModernCard": "ui_components",
        "ModernEntry": "ui_components",
        "LoadingOverlay": "ui_components",
        "EnhancedTable": "enhanced_table",
        "QueryBuilder": "querybuilder",
        "DatabaseHandler": "databasehandler",
        "EmailHistoryDialog": "emailhistorydialog",
        "TriggersDialog": "schedule_triggers_dialog",
        "ScheduleEmailDialog": "scheduleemaildialog",
        "MainApp": "mainapp",
        "AdvancedFilterDialog": "advancedfilterdialog",
        "LoginPage": "loginpage"
    }

    # Write each extracted section to its own file
    for filename, header, code in modules_to_create:
        filepath = os.path.join(output_dir, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            # 1. Add standard imports
            f.write(std_imports)
            f.write("\n")
            
            # 2. Intelligently add local imports only if the class is used in this code block
            local_imports = []
            for cls, mod_name in module_mapping.items():
                if filename != mod_name + ".py":
                    # Simple text match: if the class name appears in this code block, import it
                    # Using word boundaries to avoid partial matches
                    if re.search(r'\b' + cls + r'\b', code):
                        local_imports.append(f"from {mod_name} import {cls}")
                    
            if local_imports:
                f.write("\n".join(sorted(set(local_imports))))
                f.write("\n\n")
                
            # 3. Add the actual code
            f.write(f"# ==================== {header} ====================\n")
            f.write(code)
            print(f"Created: {filename}")

    # Create an __init__.py to tie the package together
    with open(os.path.join(output_dir, '__init__.py'), 'w', encoding='utf-8') as f:
        f.write("# Application Package\n")
        
    print("\n✅ Modularization complete! All files have been successfully split.")
    print(r"You can now run the app via: python C:\Users\srini\Downloads\app_code\main.py")

if __name__ == "__main__":
    modularize_app()
