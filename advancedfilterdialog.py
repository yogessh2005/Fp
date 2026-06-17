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

class AdvancedFilterDialog(tk.Toplevel):
    def __init__(self, parent, columns, on_apply, on_clear):
        super().__init__(parent)
        self.columns = columns
        self.on_apply = on_apply
        self.on_clear = on_clear
        self.filters = {}
       
        self.title("Advanced Filters")
        self.configure(bg=Config.COLORS["white"])
        self.geometry("550x450")
       
        x = parent.winfo_x() + parent.winfo_width()//2 - 275
        y = parent.winfo_y() + parent.winfo_height()//2 - 225
        self.geometry(f"550x450+{x}+{y}")
       
        self.setup_ui()
   
    def setup_ui(self):
        header = tk.Frame(self, bg=Config.COLORS["primary"], height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
       
        tk.Label(
            header,
            text="🔍 Advanced Column Filters",
            font=("Segoe UI", 16, "bold"),
            bg=Config.COLORS["primary"],
            fg=Config.COLORS["white"]
        ).pack(pady=15)
       
        instr_frame = tk.Frame(self, bg=Config.COLORS["white"])
        instr_frame.pack(fill="x", padx=20, pady=10)
       
        tk.Label(
            instr_frame,
            text="Enter filter values for any column (case-insensitive partial match):",
            font=("Segoe UI", 10),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["secondary"]
        ).pack(anchor="w")
       
        filter_frame = tk.Frame(self, bg=Config.COLORS["white"])
        filter_frame.pack(fill="both", expand=True, padx=20, pady=10)
       
        canvas = tk.Canvas(filter_frame, bg=Config.COLORS["white"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(filter_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=Config.COLORS["white"])
       
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
       
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
       
        self.filter_entries = {}
       
        tk.Label(
            scrollable_frame,
            text="Column Name",
            font=("Segoe UI", 11, "bold"),
            bg=Config.COLORS["primary"],
            fg=Config.COLORS["white"],
            width=25,
            anchor="w"
        ).grid(row=0, column=0, padx=5, pady=5, sticky="w")
       
        tk.Label(
            scrollable_frame,
            text="Filter Value (contains)",
            font=("Segoe UI", 11, "bold"),
            bg=Config.COLORS["primary"],
            fg=Config.COLORS["white"],
            width=30,
            anchor="w"
        ).grid(row=0, column=1, padx=5, pady=5, sticky="w")
       
        separator = tk.Frame(scrollable_frame, height=2, bg=Config.COLORS["border"])
        separator.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
       
        for idx, col in enumerate(self.columns):
            col_label = tk.Label(
                scrollable_frame,
                text=col,
                font=("Segoe UI", 10),
                bg=Config.COLORS["white"],
                fg=Config.COLORS["dark"],
                anchor="w"
            )
            col_label.grid(row=idx+2, column=0, padx=10, pady=5, sticky="w")
           
            entry = tk.Entry(
                scrollable_frame,
                font=("Segoe UI", 10),
                bg=Config.COLORS["light"],
                relief="solid",
                bd=1,
                width=35
            )
            entry.grid(row=idx+2, column=1, padx=10, pady=5, sticky="ew")
            self.filter_entries[col] = entry
       
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
       
        button_frame = tk.Frame(self, bg=Config.COLORS["white"])
        button_frame.pack(fill="x", padx=20, pady=20)
       
        def apply_filters():
            filters = {}
            for col, entry in self.filter_entries.items():
                value = entry.get().strip()
                if value:
                    filters[col] = value
            self.on_apply(filters)
            self.destroy()
       
        def clear_all_filters():
            for entry in self.filter_entries.values():
                entry.delete(0, tk.END)
            self.on_clear()
            self.destroy()
       
        def cancel():
            self.destroy()
       
        apply_btn = ModernButton(button_frame, text="✓ Apply Filters", command=apply_filters, variant="success", width=15)
        apply_btn.pack(side="left", padx=5)
       
        clear_btn = ModernButton(button_frame, text="✗ Clear All", command=clear_all_filters, variant="warning", width=12)
        clear_btn.pack(side="left", padx=5)
       
        cancel_btn = ModernButton(button_frame, text="Cancel", command=cancel, variant="secondary", width=12)
        cancel_btn.pack(side="right", padx=5)


