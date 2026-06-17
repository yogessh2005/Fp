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
from emailhistorydialog import EmailHistoryDialog
from emailscheduler import EmailScheduler
from schedule_triggers_dialog import ScheduleTriggersDialog
from ui_components import ModernButton

# ==================== SCHEDULE EMAIL DIALOG ====================

class ScheduleEmailDialog(tk.Toplevel):
    def __init__(self, parent, company: str, get_data_func, email_scheduler: EmailScheduler):
        super().__init__(parent)
        self.parent = parent
        self.company = company
        self.get_data_func = get_data_func
        self.email_scheduler = email_scheduler
        self.recipients_list = []
        self.manager_dict = {}
       
        self.email_scheduler.set_data_fetch_function(get_data_func)
       
        self.setup_window()
        self.create_ui()
        self.load_schedule_status()
        self.load_managers()
   
    def setup_window(self):
        self.title(f"⏰ Schedule Email Reports - {self.company}")
        self.configure(bg=Config.COLORS["white"])
        self.geometry("950x850")
       
        x = self.parent.winfo_x() + self.parent.winfo_width()//2 - 475
        y = self.parent.winfo_y() + self.parent.winfo_height()//2 - 425
        self.geometry(f"950x850+{x}+{y}")
       
        self.resizable(False, False)
        self.transient(self.parent)
        self.grab_set()
   
    def create_ui(self):
        header = tk.Frame(self, bg=Config.COLORS["primary"], height=70)
        header.pack(fill="x")
        header.pack_propagate(False)
       
        tk.Label(
            header,
            text="📧 Schedule Email Reports",
            font=("Segoe UI", 20, "bold"),
            bg=Config.COLORS["primary"],
            fg=Config.COLORS["white"]
        ).pack(pady=20)
       
        main_container = tk.Frame(self, bg=Config.COLORS["white"])
        main_container.pack(fill="both", expand=True, padx=25, pady=15)
       
        canvas = tk.Canvas(main_container, bg=Config.COLORS["white"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=Config.COLORS["white"])
       
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
       
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
       
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
       
        sender_frame = tk.LabelFrame(
            scrollable_frame,
            text="📤 Sender Information",
            font=("Segoe UI", 12, "bold"),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["primary"],
            padx=10,
            pady=10
        )
        sender_frame.pack(fill="x", pady=(0, 15))
       
        sender_inner = tk.Frame(sender_frame, bg=Config.COLORS["white"])
        sender_inner.pack(fill="x", pady=5)
       
        tk.Label(
            sender_inner,
            text="📧 From:",
            font=("Segoe UI", 11, "bold"),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["secondary"]
        ).pack(side="left", padx=(0, 10))
       
        sender_email = self.email_scheduler.get_sender_email()
        sender_label = tk.Label(
            sender_inner,
            text=sender_email,
            font=("Segoe UI", 11),
            bg=Config.COLORS["light"],
            fg=Config.COLORS["dark"],
            relief="solid",
            bd=1,
            padx=15,
            pady=5
        )
        sender_label.pack(side="left")
       
        tk.Label(
            sender_inner,
            text="🔒",
            font=("Segoe UI", 9),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["gray"]
        ).pack(side="left", padx=(10, 0))
       
        add_frame = tk.LabelFrame(
            scrollable_frame,
            text="➕ Add Recipients",
            font=("Segoe UI", 12, "bold"),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["primary"],
            padx=10,
            pady=10
        )
        add_frame.pack(fill="x", pady=(0, 15))
       
        add_inner = tk.Frame(add_frame, bg=Config.COLORS["white"])
        add_inner.pack(fill="x")
       
        manual_frame = tk.Frame(add_inner, bg=Config.COLORS["white"])
        manual_frame.pack(fill="x", pady=(0, 10))
       
        tk.Label(
            manual_frame,
            text="Manual Entry:",
            font=("Segoe UI", 10, "bold"),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["secondary"]
        ).pack(side="left", padx=(0, 10))
       
        self.manual_name_var = tk.StringVar()
        self.manual_email_var = tk.StringVar()
       
        name_entry = tk.Entry(
            manual_frame,
            textvariable=self.manual_name_var,
            font=("Segoe UI", 10),
            width=20,
            bg=Config.COLORS["light"],
            relief="solid",
            bd=1
        )
        name_entry.pack(side="left", padx=(0, 5))
        self.set_placeholder(name_entry, "Name")
        self.manual_name_entry = name_entry
       
        email_entry = tk.Entry(
            manual_frame,
            textvariable=self.manual_email_var,
            font=("Segoe UI", 10),
            width=25,
            bg=Config.COLORS["light"],
            relief="solid",
            bd=1
        )
        email_entry.pack(side="left", padx=(0, 5))
        self.set_placeholder(email_entry, "email@example.com")
        self.manual_email_entry = email_entry
       
        add_btn = ModernButton(
            manual_frame,
            text="➕ Add",
            command=self.add_manual_recipient,
            variant="success",
            width=10
        )
        add_btn.pack(side="left", padx=5)
       
        manager_frame = tk.Frame(add_inner, bg=Config.COLORS["white"])
        manager_frame.pack(fill="x")
       
        tk.Label(
            manager_frame,
            text="Or Select from Managers:",
            font=("Segoe UI", 10, "bold"),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["secondary"]
        ).pack(side="left", padx=(0, 10))
       
        self.manager_var = tk.StringVar()
        self.manager_dropdown = ttk.Combobox(
            manager_frame,
            textvariable=self.manager_var,
            font=("Segoe UI", 10),
            width=35,
            state="readonly"
        )
        self.manager_dropdown.pack(side="left", padx=(0, 5))
       
        add_manager_btn = ModernButton(
            manager_frame,
            text="➕ Add Selected",
            command=self.add_manager_recipient,
            variant="primary",
            width=14
        )
        add_manager_btn.pack(side="left", padx=5)
       
        recipient_frame = tk.LabelFrame(
            scrollable_frame,
            text="📧 Recipient List",
            font=("Segoe UI", 12, "bold"),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["primary"],
            padx=10,
            pady=10
        )
        recipient_frame.pack(fill="both", expand=True, pady=(0, 15))
       
        control_frame = tk.Frame(recipient_frame, bg=Config.COLORS["white"])
        control_frame.pack(fill="x", pady=(0, 10))
       
        remove_selected_btn = ModernButton(
            control_frame,
            text="🗑 Remove Selected",
            command=self.remove_selected_recipients,
            variant="danger",
            width=15
        )
        remove_selected_btn.pack(side="left", padx=5)
       
        clear_all_btn = ModernButton(
            control_frame,
            text="✖ Clear All",
            command=self.clear_all_recipients,
            variant="warning",
            width=12
        )
        clear_all_btn.pack(side="left", padx=5)
       
        history_btn = ModernButton(
            control_frame,
            text="📜 History",
            command=self.show_history,
            variant="history",
            width=12
        )
        history_btn.pack(side="left", padx=5)
       
        triggers_btn = ModernButton(
            control_frame,
            text="⏰ Triggers",
            command=self.show_triggers,
            variant="info",
            width=12
        )
        triggers_btn.pack(side="left", padx=5)
       
        self.recipient_count_label = tk.Label(
            control_frame,
            text="0 recipients",
            font=("Segoe UI", 10, "bold"),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["accent"]
        )
        self.recipient_count_label.pack(side="right", padx=10)
       
        list_container = tk.Frame(recipient_frame, bg=Config.COLORS["white"])
        list_container.pack(fill="both", expand=True)
       
        r_canvas = tk.Canvas(list_container, bg=Config.COLORS["white"], highlightthickness=0)
        r_scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=r_canvas.yview)
        self.recipient_container = tk.Frame(r_canvas, bg=Config.COLORS["white"])
       
        self.recipient_container.bind(
            "<Configure>",
            lambda e: r_canvas.configure(scrollregion=r_canvas.bbox("all"))
        )
       
        r_canvas.create_window((0, 0), window=self.recipient_container, anchor="nw")
        r_canvas.configure(yscrollcommand=r_scrollbar.set)
       
        r_canvas.pack(side="left", fill="both", expand=True)
        r_scrollbar.pack(side="right", fill="y")
       
        trigger_frame = tk.LabelFrame(
            scrollable_frame,
            text="⏰ Add New Schedule Trigger",
            font=("Segoe UI", 12, "bold"),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["primary"],
            padx=10,
            pady=10
        )
        trigger_frame.pack(fill="x", pady=(0, 15))
       
        trigger_inner = tk.Frame(trigger_frame, bg=Config.COLORS["white"])
        trigger_inner.pack(fill="x")
       
        tk.Label(
            trigger_inner,
            text="Time:",
            font=("Segoe UI", 11),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["secondary"]
        ).pack(side="left", padx=(0, 10))
       
        self.time_var = tk.StringVar(value="17:00")
        time_entry = tk.Entry(
            trigger_inner,
            textvariable=self.time_var,
            font=("Segoe UI", 14),
            width=8,
            justify="center",
            bg=Config.COLORS["light"],
            relief="solid",
            bd=1
        )
        time_entry.pack(side="left", padx=(0, 10))
       
        tk.Label(
            trigger_inner,
            text="(24-hour format)",
            font=("Segoe UI", 9),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["gray"]
        ).pack(side="left", padx=(0, 20))
       
        add_trigger_btn = ModernButton(
            trigger_inner,
            text="➕ Add Trigger",
            command=self.add_trigger,
            variant="schedule",
            width=15
        )
        add_trigger_btn.pack(side="left", padx=5)
       
        self.trigger_summary_label = tk.Label(
            trigger_inner,
            text="",
            font=("Segoe UI", 10),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["success"]
        )
        self.trigger_summary_label.pack(side="left", padx=20)
       
        self.status_label = tk.Label(
            scrollable_frame,
            text="",
            font=("Segoe UI", 10),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["secondary"]
        )
        self.status_label.pack(pady=(5, 0))
       
        button_frame = tk.Frame(scrollable_frame, bg=Config.COLORS["white"])
        button_frame.pack(fill="x", pady=(10, 0))
       
        close_btn = ModernButton(
            button_frame,
            text="✖ Close",
            command=self.destroy,
            variant="secondary",
            width=12
        )
        close_btn.pack(side="right", padx=5)
       
        self.message_label = tk.Label(
            scrollable_frame,
            text="",
            font=("Segoe UI", 10),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["secondary"]
        )
        self.message_label.pack(pady=(10, 0))
   
    def show_triggers(self):
        ScheduleTriggersDialog(self, self.email_scheduler)
   
    def show_history(self):
        EmailHistoryDialog(self, self.email_scheduler)
   
    def update_trigger_summary(self):
        triggers = self.email_scheduler.get_active_triggers()
        if triggers:
            times = [t.get('time', '') for t in triggers]
            self.trigger_summary_label.config(
                text=f"🟢 Active Triggers: {', '.join(times)}",
                fg=Config.COLORS["success"]
            )
        else:
            self.trigger_summary_label.config(
                text="🔴 No active triggers",
                fg=Config.COLORS["gray"]
            )
   
    def set_placeholder(self, entry, placeholder):
        entry.insert(0, placeholder)
        entry.config(fg=Config.COLORS["gray"])
       
        def on_focus_in(e):
            if entry.get() == placeholder:
                entry.delete(0, tk.END)
                entry.config(fg=Config.COLORS["dark"])
       
        def on_focus_out(e):
            if not entry.get():
                entry.insert(0, placeholder)
                entry.config(fg=Config.COLORS["gray"])
       
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)
   
    def load_managers(self):
        all_data = self.get_data_func()
        managers = []
        manager_dict = {}
       
        if all_data and 'Score' in all_data:
            score_df = all_data['Score']
           
            for _, row in score_df.iterrows():
                site_manager = row.get('Site Manager', '')
                unit_manager = row.get('Unit Manager', '')
                area_manager = row.get('Area Manager', '')
                site_email = row.get('Site Manager Email', '')
                unit_email = row.get('Unit Manager Email', '')
                area_email = row.get('Area Manager Email', '')
               
                if site_manager and str(site_manager).strip() and site_email and '@' in str(site_email):
                    key = f"Site: {site_manager}"
                    if key not in manager_dict:
                        manager_dict[key] = str(site_email)
                        managers.append(key)
               
                if unit_manager and str(unit_manager).strip() and unit_email and '@' in str(unit_email):
                    key = f"Unit: {unit_manager}"
                    if key not in manager_dict:
                        manager_dict[key] = str(unit_email)
                        managers.append(key)
               
                if area_manager and str(area_manager).strip() and area_email and '@' in str(area_email):
                    key = f"Area: {area_manager}"
                    if key not in manager_dict:
                        manager_dict[key] = str(area_email)
                        managers.append(key)
       
        self.manager_dict = manager_dict
       
        if managers:
            self.manager_dropdown['values'] = managers
            self.manager_dropdown.set("Select a manager...")
        else:
            self.manager_dropdown['values'] = ["No managers available"]
            self.manager_dropdown.set("No managers available")
   
    def add_manual_recipient(self):
        name = self.manual_name_var.get().strip()
        email = self.manual_email_var.get().strip()
       
        if not name or name == "Name":
            self.show_message("⚠️ Please enter a name", "warning")
            return
       
        if not email or email == "email@example.com" or '@' not in email:
            self.show_message("⚠️ Please enter a valid email", "warning")
            return
       
        if any(r[1] == email for r in self.recipients_list):
            self.show_message(f"⚠️ {email} already exists", "warning")
            return
       
        self.recipients_list.append((name, email))
        self.update_recipient_display()
       
        self.manual_name_var.set("")
        self.manual_email_var.set("")
       
        self.show_message(f"✅ Added: {name} ({email})", "success")
   
    def add_manager_recipient(self):
        selected = self.manager_var.get()
       
        if not selected or selected == "Select a manager..." or selected == "No managers available":
            self.show_message("⚠️ Please select a manager", "warning")
            return
       
        name = selected
        email = self.manager_dict.get(selected, '')
       
        if not email:
            self.show_message("⚠️ No email found for this manager", "warning")
            return
       
        if any(r[1] == email for r in self.recipients_list):
            self.show_message(f"⚠️ {name} already exists", "warning")
            return
       
        self.recipients_list.append((name, email))
        self.update_recipient_display()
        self.show_message(f"✅ Added: {name} ({email})", "success")
   
    def remove_selected_recipients(self):
        to_remove = []
        for child in self.recipient_container.winfo_children():
            if hasattr(child, 'var') and child.var.get():
                if hasattr(child, 'email'):
                    to_remove.append(child.email)
       
        if not to_remove:
            self.show_message("⚠️ No recipients selected", "warning")
            return
       
        if messagebox.askyesno("Confirm Remove", f"Remove {len(to_remove)} selected recipients?"):
            self.recipients_list = [r for r in self.recipients_list if r[1] not in to_remove]
            self.update_recipient_display()
            self.show_message(f"🗑 Removed {len(to_remove)} recipients", "info")
   
    def clear_all_recipients(self):
        if not self.recipients_list:
            return
       
        if messagebox.askyesno("Confirm Clear", "Remove all recipients from the list?"):
            self.recipients_list = []
            self.update_recipient_display()
            self.show_message("🗑 All recipients cleared", "info")
   
    def update_recipient_display(self):
        for widget in self.recipient_container.winfo_children():
            widget.destroy()
       
        for idx, (name, email) in enumerate(self.recipients_list):
            item_frame = tk.Frame(
                self.recipient_container,
                bg=Config.COLORS["alternate_row"] if idx % 2 == 0 else Config.COLORS["white"],
                relief="solid",
                bd=1
            )
            item_frame.pack(fill="x", pady=1)
           
            var = tk.BooleanVar()
            item_frame.var = var
            item_frame.email = email
           
            cb = tk.Checkbutton(
                item_frame,
                variable=var,
                bg=item_frame['bg'],
                font=("Segoe UI", 10)
            )
            cb.pack(side="left", padx=(5, 0), pady=5)
           
            tk.Label(
                item_frame,
                text=f"{name}",
                font=("Segoe UI", 10, "bold"),
                bg=item_frame['bg'],
                fg=Config.COLORS["dark"],
                width=25,
                anchor="w"
            ).pack(side="left", padx=(5, 10), pady=5)
           
            tk.Label(
                item_frame,
                text=f"<{email}>",
                font=("Segoe UI", 9),
                bg=item_frame['bg'],
                fg=Config.COLORS["gray"]
            ).pack(side="left", padx=(0, 10), pady=5)
       
        self.recipient_count_label.config(text=f"{len(self.recipients_list)} recipients")
   
    def add_trigger(self):
        if not self.recipients_list:
            self.show_message("⚠️ Please add at least one recipient", "danger")
            return
       
        time_str = self.time_var.get().strip()
        try:
            datetime.strptime(time_str, "%H:%M")
        except ValueError:
            self.show_message("⚠️ Invalid time format. Use HH:MM (24-hour)", "danger")
            return
       
        recipients = [email for _, email in self.recipients_list]
        recipient_names = [name for name, _ in self.recipients_list]
        selected_date = datetime.now().strftime("%d/%m/%Y")
       
        recipient_display = '\n'.join([f"  • {name} ({email})" for name, email in self.recipients_list[:3]])
        if len(self.recipients_list) > 3:
            recipient_display += f"\n  • ... and {len(self.recipients_list) - 3} more"
       
        confirm_msg = (
            f"📧 Add Schedule Trigger\n\n"
            f"📤 From: {self.email_scheduler.get_sender_email()}\n"
            f"⏰ Time: {time_str} daily\n"
            f"👥 Recipients ({len(recipients)}):\n{recipient_display}\n\n"
            f"📊 Report: {self.company} with updated PSERM logic (>35 = 25 marks)\n"
            f"🔄 Runs 24/7 in background\n\n"
            f"Continue?"
        )
       
        if not messagebox.askyesno("Confirm Trigger", confirm_msg):
            return
       
        success, message, trigger = self.email_scheduler.add_trigger(
            time_str=time_str,
            recipients=recipients,
            recipient_names=recipient_names,
            company=self.company,
            selected_date=selected_date
        )
       
        if success:
            self.show_message(f"✅ {message}", "success")
            self.update_trigger_summary()
            messagebox.showinfo(
                "✅ Trigger Added",
                f"Schedule trigger added successfully!\n\n"
                f"📤 From: {self.email_scheduler.get_sender_email()}\n"
                f"⏰ Time: {time_str} daily\n"
                f"👥 Recipients: {len(recipients)}\n"
                f"🏢 Company: {self.company}\n"
                f"📊 Includes: Updated PSERM logic (>35 = 25 marks)\n"
                f"🔄 Runs 24/7 in background\n\n"
                f"💡 Click 'Triggers' to view all triggers and toggle Active/Inactive."
            )
        else:
            self.show_message(f"❌ {message}", "danger")
   
    def show_message(self, message, msg_type="info"):
        colors = {
            "info": Config.COLORS["secondary"],
            "success": Config.COLORS["success"],
            "warning": Config.COLORS["warning"],
            "danger": Config.COLORS["danger"]
        }
       
        self.message_label.config(text=message, fg=colors.get(msg_type, Config.COLORS["secondary"]))
       
        if msg_type != "danger":
            self.after(5000, lambda: self.message_label.config(text=""))
   
    def load_schedule_status(self):
        self.update_trigger_summary()
       
        triggers = self.email_scheduler.get_active_triggers()
        if triggers:
            times = [t.get('time', '') for t in triggers]
            self.status_label.config(
                text=f"✅ Active triggers: {', '.join(times)}",
                fg=Config.COLORS["success"]
            )
        else:
            self.status_label.config(
                text="📌 No active triggers. Add a trigger to schedule emails.",
                fg=Config.COLORS["secondary"]
            )


