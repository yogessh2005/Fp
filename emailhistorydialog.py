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
from emailscheduler import EmailScheduler
from ui_components import ModernButton

# ==================== EMAIL HISTORY DIALOG ====================

class EmailHistoryDialog(tk.Toplevel):
    def __init__(self, parent, email_scheduler: EmailScheduler):
        super().__init__(parent)
        self.parent = parent
        self.email_scheduler = email_scheduler
       
        self.setup_window()
        self.create_ui()
        self.load_history()
   
    def setup_window(self):
        self.title("📜 Email Sending History")
        self.configure(bg=Config.COLORS["white"])
        self.geometry("950x500")
       
        x = self.parent.winfo_x() + self.parent.winfo_width()//2 - 475
        y = self.parent.winfo_y() + self.parent.winfo_height()//2 - 250
        self.geometry(f"950x500+{x}+{y}")
       
        self.resizable(True, True)
        self.transient(self.parent)
        self.grab_set()
   
    def create_ui(self):
        header = tk.Frame(self, bg=Config.COLORS["primary"], height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
       
        tk.Label(
            header,
            text="📜 Email Sending History",
            font=("Segoe UI", 18, "bold"),
            bg=Config.COLORS["primary"],
            fg=Config.COLORS["white"]
        ).pack(pady=15)
       
        main_frame = tk.Frame(self, bg=Config.COLORS["white"])
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
       
        columns = ("Timestamp", "Company", "Recipients", "Count", "Date", "Trigger", "Status")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings", height=15)
       
        self.tree.heading("Timestamp", text="Timestamp")
        self.tree.heading("Company", text="Company")
        self.tree.heading("Recipients", text="Recipients")
        self.tree.heading("Count", text="Count")
        self.tree.heading("Date", text="Report Date")
        self.tree.heading("Trigger", text="Trigger Time")
        self.tree.heading("Status", text="Status")
       
        self.tree.column("Timestamp", width=150, anchor="center")
        self.tree.column("Company", width=80, anchor="center")
        self.tree.column("Recipients", width=280, anchor="w")
        self.tree.column("Count", width=50, anchor="center")
        self.tree.column("Date", width=100, anchor="center")
        self.tree.column("Trigger", width=80, anchor="center")
        self.tree.column("Status", width=100, anchor="center")
       
        vsb = ttk.Scrollbar(main_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(main_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
       
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
       
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
       
        button_frame = tk.Frame(self, bg=Config.COLORS["white"])
        button_frame.pack(fill="x", padx=15, pady=(0, 15))
       
        refresh_btn = ModernButton(button_frame, text="🔄 Refresh", command=self.load_history, variant="primary", width=12)
        refresh_btn.pack(side="left", padx=5)
       
        clear_btn = ModernButton(button_frame, text="🗑 Clear History", command=self.clear_history, variant="danger", width=12)
        clear_btn.pack(side="left", padx=5)
       
        close_btn = ModernButton(button_frame, text="✖ Close", command=self.destroy, variant="secondary", width=12)
        close_btn.pack(side="right", padx=5)
       
        self.info_label = tk.Label(button_frame, text="", font=("Segoe UI", 10), bg=Config.COLORS["white"], fg=Config.COLORS["secondary"])
        self.info_label.pack(side="left", padx=20)
   
    def load_history(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
       
        history = self.email_scheduler.history.get_history()
       
        if not history:
            self.tree.insert("", "end", values=("No history found", "", "", "", "", "", ""))
            self.info_label.config(text="No email history available")
            return
       
        for record in history:
            timestamp = record.get("timestamp", "")
            company = record.get("company", "")
            recipients = ", ".join(record.get("recipients", [])[:3])
            if len(record.get("recipients", [])) > 3:
                recipients += f" ... +{len(record.get('recipients', [])) - 3} more"
            count = record.get("recipient_count", 0)
            date = record.get("selected_date", "")
            trigger = record.get("trigger_time", "")
            status = record.get("status", "")
           
            if "SUCCESS" in status or "✅" in status:
                status = "✅ SUCCESS"
            else:
                status = "❌ FAILED"
           
            self.tree.insert("", "end", values=(timestamp, company, recipients, count, date, trigger, status))
       
        self.info_label.config(text=f"Total: {len(history)} records")
   
    def clear_history(self):
        if messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all email history?"):
            self.email_scheduler.history.history = []
            self.email_scheduler.history.save_history()
            self.load_history()
            messagebox.showinfo("History Cleared", "Email history has been cleared.")


