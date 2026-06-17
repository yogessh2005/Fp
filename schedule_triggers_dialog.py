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

# ==================== SCHEDULE TRIGGERS DIALOG ====================

class ScheduleTriggersDialog(tk.Toplevel):
    """Display all schedule triggers with Active/Inactive toggle"""
   
    def __init__(self, parent, email_scheduler: EmailScheduler):
        super().__init__(parent)
        self.parent = parent
        self.email_scheduler = email_scheduler
       
        self.setup_window()
        self.create_ui()
        self.load_triggers()
   
    def setup_window(self):
        self.title("⏰ Schedule Triggers - Active/Inactive Management")
        self.configure(bg=Config.COLORS["white"])
        self.geometry("950x550")
       
        x = self.parent.winfo_x() + self.parent.winfo_width()//2 - 475
        y = self.parent.winfo_y() + self.parent.winfo_height()//2 - 275
        self.geometry(f"950x550+{x}+{y}")
       
        self.resizable(True, True)
        self.transient(self.parent)
        self.grab_set()
   
    def create_ui(self):
        header = tk.Frame(self, bg=Config.COLORS["primary"], height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
       
        tk.Label(
            header,
            text="⏰ Schedule Triggers",
            font=("Segoe UI", 18, "bold"),
            bg=Config.COLORS["primary"],
            fg=Config.COLORS["white"]
        ).pack(pady=15)
       
        main_frame = tk.Frame(self, bg=Config.COLORS["white"])
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)
       
        summary_frame = tk.Frame(main_frame, bg=Config.COLORS["light"], relief="solid", bd=1)
        summary_frame.pack(fill="x", pady=(0, 10))
       
        self.summary_label = tk.Label(
            summary_frame,
            text="",
            font=("Segoe UI", 11),
            bg=Config.COLORS["light"],
            fg=Config.COLORS["dark"],
            padx=15,
            pady=10
        )
        self.summary_label.pack()
       
        instr_label = tk.Label(
            main_frame,
            text="💡 Click 'Deactivate' to turn OFF a trigger | Click 'Activate' to turn ON a trigger",
            font=("Segoe UI", 10),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["gray"]
        )
        instr_label.pack(anchor="w", pady=(0, 10))
       
        table_frame = tk.Frame(main_frame, bg=Config.COLORS["white"])
        table_frame.pack(fill="both", expand=True)
       
        columns = ("Status", "Time", "Company", "Recipients", "Created", "Last Sent", "Action")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=10)
       
        self.tree.heading("Status", text="Status")
        self.tree.heading("Time", text="Time")
        self.tree.heading("Company", text="Company")
        self.tree.heading("Recipients", text="Recipients")
        self.tree.heading("Created", text="Created")
        self.tree.heading("Last Sent", text="Last Sent")
        self.tree.heading("Action", text="Action")
       
        self.tree.column("Status", width=120, anchor="center")
        self.tree.column("Time", width=80, anchor="center")
        self.tree.column("Company", width=80, anchor="center")
        self.tree.column("Recipients", width=250, anchor="w")
        self.tree.column("Created", width=150, anchor="center")
        self.tree.column("Last Sent", width=150, anchor="center")
        self.tree.column("Action", width=100, anchor="center")
       
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
       
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
       
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
       
        self.tree.bind("<Button-1>", self.on_action_click)
        self.tree.bind("<Double-1>", self.on_trigger_double_click)
       
        button_frame = tk.Frame(self, bg=Config.COLORS["white"])
        button_frame.pack(fill="x", padx=15, pady=(0, 15))
       
        refresh_btn = ModernButton(button_frame, text="🔄 Refresh", command=self.load_triggers, variant="primary", width=12)
        refresh_btn.pack(side="left", padx=5)
       
        close_btn = ModernButton(button_frame, text="✖ Close", command=self.destroy, variant="secondary", width=12)
        close_btn.pack(side="right", padx=5)
   
    def get_trigger_id_from_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            item = self.tree.identify_row(event.y)
            if item and column == "#7":
                tags = self.tree.item(item, 'tags')
                if tags:
                    return tags[0]
        return None
   
    def on_action_click(self, event):
        trigger_id = self.get_trigger_id_from_click(event)
        if trigger_id:
            self.toggle_trigger(trigger_id)
   
    def on_trigger_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            tags = self.tree.item(item, 'tags')
            if tags:
                self.toggle_trigger(tags[0])
   
    def toggle_trigger(self, trigger_id):
        trigger = self.email_scheduler.get_trigger_by_id(trigger_id)
        if not trigger:
            return
       
        enabled = trigger.get('enabled', False)
        action = "deactivate" if enabled else "activate"
       
        recipient_names = trigger.get('recipient_names', [])
        recipient_display = ", ".join(recipient_names[:2])
        if len(recipient_names) > 2:
            recipient_display += f" ... +{len(recipient_names) - 2} more"
       
        confirm_msg = (
            f"Confirm {action.upper()}\n\n"
            f"⏰ Time: {trigger.get('time', '')}\n"
            f"🏢 Company: {trigger.get('company', '')}\n"
            f"👥 Recipients: {len(recipient_names)} people\n"
            f"📧 To: {recipient_display}\n"
            f"📊 Current Status: {'🟢 ACTIVE' if enabled else '🔴 INACTIVE'}\n\n"
            f"Are you sure you want to {action} this trigger?"
        )
       
        if messagebox.askyesno(f"Confirm {action.upper()}", confirm_msg):
            success, updated_trigger = self.email_scheduler.toggle_trigger(trigger_id)
            if success:
                self.load_triggers()
                if hasattr(self.parent, 'update_schedule_status'):
                    self.parent.update_schedule_status()
               
                new_status = "deactivated" if enabled else "activated"
                messagebox.showinfo(
                    "✅ Success",
                    f"Trigger {new_status} successfully!\n\n"
                    f"⏰ Time: {trigger.get('time', '')}\n"
                    f"🏢 Company: {trigger.get('company', '')}\n"
                    f"📊 New Status: {'🔴 INACTIVE' if enabled else '🟢 ACTIVE'}"
                )
            else:
                messagebox.showerror("❌ Error", "Failed to toggle trigger status")
   
    def load_triggers(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
       
        triggers = self.email_scheduler.get_all_triggers()
        active_count = len([t for t in triggers if t.get('enabled', False)])
       
        self.summary_label.config(
            text=f"📊 Total Triggers: {len(triggers)} | 🟢 Active: {active_count} | 🔴 Inactive: {len(triggers) - active_count}"
        )
       
        if not triggers:
            self.tree.insert("", "end", values=("No triggers configured", "", "", "", "", "", ""))
            return
       
        for trigger in triggers:
            enabled = trigger.get('enabled', False)
            time_str = trigger.get('time', '')
            company = trigger.get('company', '')
            recipients = ", ".join(trigger.get('recipient_names', [])[:3])
            if len(trigger.get('recipient_names', [])) > 3:
                recipients += f" ... +{len(trigger.get('recipient_names', [])) - 3} more"
            created = trigger.get('created_at', '')
            last_sent = trigger.get('last_sent', 'Never')
            trigger_id = trigger.get('id', '')
           
            if enabled:
                status = "🟢 ACTIVE"
                action_text = "🔴 Deactivate"
            else:
                status = "🔴 INACTIVE"
                action_text = "🟢 Activate"
           
            item_id = self.tree.insert("", "end", values=(status, time_str, company, recipients, created, last_sent, action_text))
            self.tree.item(item_id, tags=(trigger_id,))
           
            if enabled:
                self.tree.tag_configure('active', background='#E8F5E9')
            else:
                self.tree.tag_configure('inactive', background='#FFEBEE')


