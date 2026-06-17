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
from ui_components import ModernButton

# ==================== ADVANCED FILTER DIALOG ====================

class ColumnFilterDialog(tk.Toplevel):
    def __init__(self, parent, column_name, x, y, on_apply, on_clear, existing_filter=None):
        super().__init__(parent)
        self.column_name = column_name
        self.on_apply = on_apply
        self.on_clear = on_clear
        
        self.title("Custom Filter")
        self.configure(bg=Config.COLORS["white"])
        self.geometry(f"250x300+{x}+{y}")
        self.resizable(False, False)
        
        # Remove window decorations to make it look like a popup if desired,
        # but a normal Tool window is fine too.
        self.attributes('-toolwindow', True)
        
        self.operators = [
            "Is equal to",
            "Is not equal to",
            "Is greater than",
            "Is greater than or equal to",
            "Is less than",
            "Is less than or equal to",
            "Is null",
            "Is not null"
        ]
        
        self.setup_ui(existing_filter)
        
    def setup_ui(self, existing_filter):
        main_frame = tk.Frame(self, bg=Config.COLORS["white"], padx=15, pady=15)
        main_frame.pack(fill="both", expand=True)
        
        tk.Label(
            main_frame,
            text="Show items with value that:",
            font=("Segoe UI", 10),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["dark"],
            anchor="w"
        ).pack(fill="x", pady=(0, 10))
        
        # Condition 1
        self.op1_var = tk.StringVar(value=self.operators[0])
        self.op1_cb = ttk.Combobox(main_frame, textvariable=self.op1_var, values=self.operators, state="readonly", font=("Segoe UI", 9))
        self.op1_cb.pack(fill="x", pady=(0, 5))
        
        self.val1_entry = ttk.Spinbox(main_frame, font=("Segoe UI", 9), from_=-999999999, to=999999999)
        self.val1_entry.pack(fill="x", pady=(0, 15))
        
        # AND/OR logic
        self.logic_var = tk.StringVar(value="And")
        logic_frame = tk.Frame(main_frame, bg=Config.COLORS["white"])
        logic_frame.pack(fill="x", pady=(0, 15))
        self.logic_cb = ttk.Combobox(logic_frame, textvariable=self.logic_var, values=["And", "Or"], state="readonly", font=("Segoe UI", 9), width=8)
        self.logic_cb.pack(side="left")
        
        # Condition 2
        self.op2_var = tk.StringVar(value=self.operators[0])
        self.op2_cb = ttk.Combobox(main_frame, textvariable=self.op2_var, values=self.operators, state="readonly", font=("Segoe UI", 9))
        self.op2_cb.pack(fill="x", pady=(0, 5))
        
        self.val2_entry = ttk.Spinbox(main_frame, font=("Segoe UI", 9), from_=-999999999, to=999999999)
        self.val2_entry.pack(fill="x", pady=(0, 20))
        
        # Pre-fill existing if any
        if existing_filter:
            if "op1" in existing_filter: self.op1_var.set(existing_filter["op1"])
            if "val1" in existing_filter: self.val1_entry.insert(0, existing_filter["val1"])
            if "logic" in existing_filter: self.logic_var.set(existing_filter["logic"])
            if "op2" in existing_filter: self.op2_var.set(existing_filter["op2"])
            if "val2" in existing_filter: self.val2_entry.insert(0, existing_filter["val2"])
            
        def apply_filter():
            f = {}
            op1 = self.op1_var.get()
            val1 = self.val1_entry.get().strip()
            logic = self.logic_var.get()
            op2 = self.op2_var.get()
            val2 = self.val2_entry.get().strip()
            
            # Need to apply if value is present OR if the operator is Is null / Is not null
            valid_op1 = val1 or "null" in op1.lower()
            valid_op2 = val2 or "null" in op2.lower()
            
            if valid_op1 or valid_op2:
                if valid_op1:
                    f["op1"] = op1
                    f["val1"] = val1
                if valid_op2:
                    f["op2"] = op2
                    f["val2"] = val2
                f["logic"] = logic
                self.on_apply(self.column_name, f)
            else:
                self.on_clear(self.column_name)
            self.destroy()
            
        def clear_filter():
            self.on_clear(self.column_name)
            self.destroy()
            
        btn_frame = tk.Frame(main_frame, bg=Config.COLORS["white"])
        btn_frame.pack(fill="x", side="bottom")
        
        # Simple buttons to match standard styling
        filter_btn = tk.Button(btn_frame, text="Filter", bg=Config.COLORS["primary"], fg=Config.COLORS["white"], font=("Segoe UI", 9, "bold"), relief="flat", command=apply_filter, width=12)
        filter_btn.pack(side="left", padx=(0, 5))
        
        clear_btn = tk.Button(btn_frame, text="Clear", bg=Config.COLORS["white"], fg=Config.COLORS["dark"], font=("Segoe UI", 9), relief="solid", bd=1, command=clear_filter, width=12)
        clear_btn.pack(side="left")


