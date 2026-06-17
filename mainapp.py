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

from advancedfilterdialog import AdvancedFilterDialog
from applogger import AppLogger
from config import Config
from databasehandler import DatabaseHandler
from emailscheduler import EmailScheduler
from enhanced_table import EnhancedTable
from excelformatter import ExcelFormatter
from grntargetmanager import GRNTargetManager
from querybuilder import QueryBuilder
from scheduleemaildialog import ScheduleEmailDialog
from ui_components import LoadingOverlay
from ui_components import ModernButton

# ==================== MAIN APPLICATION ====================

class MainApp(tk.Toplevel):
    def __init__(self, company: str, db_handler: DatabaseHandler, logger: AppLogger, login_page):
        super().__init__(login_page)
       
        self.company = company
        self.db_handler = db_handler
        self.logger = logger
        self.login_page = login_page
        self.all_data = {}
        self.selected_date = Config.DEFAULT_DATE
        self.selected_date_display = Config.DEFAULT_DATE_DISPLAY
        self.current_loading = None
        self.is_loading = False
       
        # ⭐ Store selected PUName for cross-tab navigation
        self.selected_pu_name = None
       
        self.email_scheduler = EmailScheduler(logger)
       
        self.setup_window()
        self.create_header()
        self.create_main_content()
        self.create_status_bar()
        self.load_initial_data()
       
        self.logger.info(f"Main application opened for {company}")
   
    def setup_window(self):
        self.title(f"{Config.APP_NAME} - {self.company}")
        self.configure(bg=Config.COLORS["light"])
        self.attributes('-fullscreen', True)
        self.bind("<Escape>", lambda e: self.attributes('-fullscreen', False))
        self.bind("<F11>", lambda e: self.toggle_fullscreen())
        self.minsize(1200, 700)
   
    def toggle_fullscreen(self):
        self.attributes('-fullscreen', not self.attributes('-fullscreen'))
   
    def create_header(self):
        header = tk.Frame(self, bg=Config.COLORS["white"], height=100)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        company_color = Config.COMPANY_COLORS.get(self.company, Config.COLORS["accent"])
        
        logo_frame = tk.Frame(header, bg=Config.COLORS["white"])
        logo_frame.pack(side="left", padx=30, pady=20)
        
        tk.Label(
            logo_frame,
            text=f"{Config.COMPANY_LOGO}",
            font=("Segoe UI", 28),
            bg=Config.COLORS["white"],
            fg=company_color
        ).pack(side="left", padx=(0, 15))
        
        company_label = tk.Label(
            logo_frame,
            text=f"{self.company.upper()} Portal",
            font=("Segoe UI", 20, "bold"),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["primary"]
        )
        company_label.pack(side="left")
        
        date_frame = tk.Frame(header, bg=Config.COLORS["white"])
        date_frame.pack(side="left", padx=50)
        
        tk.Label(
            date_frame,
            text="Analysis Date:",
            font=("Segoe UI", 11),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["secondary"]
        ).pack(side="left", padx=(0, 10))
        
        self.date_var = tk.StringVar(value=Config.DEFAULT_DATE_DISPLAY)
        date_display = tk.Entry(
            date_frame,
            textvariable=self.date_var,
            font=("Segoe UI", 11),
            width=15,
            justify="center",
            bg=Config.COLORS["light"],
            relief="solid",
            bd=1,
            state="readonly",
            readonlybackground=Config.COLORS["light"]
        )
        date_display.pack(side="left", padx=5, ipady=5)
        
        calendar_btn = ModernButton(
            date_frame,
            text="📅 Select Date",
            command=self.show_calendar,
            variant="primary",
            width=12
        )
        calendar_btn.pack(side="left", padx=5)
        
        action_frame = tk.Frame(header, bg=Config.COLORS["white"])
        action_frame.pack(side="right", padx=30, pady=5)
        
        self.refresh_btn = ModernButton(
            action_frame,
            text="🔄 Refresh All",
            command=self.refresh_all_data,
            variant="secondary"
        )
        self.refresh_btn.grid(row=0, column=0, padx=5, pady=3)
        
        self.export_btn = ModernButton(
            action_frame,
            text="📥 Export to Excel",
            command=self.export_to_excel,
            variant="success"
        )
        self.export_btn.grid(row=0, column=1, padx=5, pady=3)
        
        self.schedule_btn = ModernButton(
            action_frame,
            text="⏰ Schedule Report",
            command=self.open_schedule_dialog,
            variant="schedule"
        )
        self.schedule_btn.grid(row=0, column=2, padx=5, pady=3)
        
        # ⭐ Upload GRN Target Button (Background Upload)
        self.upload_grn_btn = ModernButton(
            action_frame,
            text="📤 Upload GRN Target",
            command=self.upload_grn_target,
            variant="upload",
            width=16
        )
        self.upload_grn_btn.grid(row=0, column=3, padx=5, pady=3)
        
        # ⭐ GRN Target Display
        self.grn_target_label = tk.Label(
            action_frame,
            text=f"Target: {GRNTargetManager.get_target():.1f}",
            font=("Segoe UI", 10, "bold"),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["success"]
        )
        self.grn_target_label.grid(row=0, column=4, padx=5, pady=3, sticky="w")
        
        self.progress_label = tk.Label(
            action_frame,
            text="0/5 Loaded",
            font=("Segoe UI", 10, "bold"),
            bg=Config.COLORS["white"],
            fg=company_color
        )
        self.progress_label.grid(row=1, column=0, padx=5, pady=3, sticky="e")
        
        self.progress_bar = ttk.Progressbar(
            action_frame,
            length=150,
            mode='determinate',
            style="modern.Horizontal.TProgressbar"
        )
        self.progress_bar.grid(row=1, column=1, padx=5, pady=3, sticky="w")
        
        self.schedule_indicator = tk.Label(
            action_frame,
            text="",
            font=("Segoe UI", 10),
            bg=Config.COLORS["white"]
        )
        self.schedule_indicator.grid(row=1, column=2, padx=5, pady=3, sticky="w", columnspan=2)
        
        self.logout_btn = ModernButton(
            action_frame,
            text="🚪 Logout",
            command=self.logout,
            variant="danger"
        )
        self.logout_btn.grid(row=1, column=4, padx=5, pady=3, sticky="e")
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("modern.Horizontal.TProgressbar", background=Config.COLORS["success"], thickness=8)
        
        self.update_schedule_indicator()
   
    def upload_grn_target(self):
        """Upload GRN target from Excel file (Background Thread)"""
        filepath = filedialog.askopenfilename(
            title="Select GRN Target Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
       
        if not filepath:
            return
       
        # ⭐ Show loading overlay
        loading = LoadingOverlay(self, "Uploading GRN Target...")
       
        def load_target():
            success, message, target = GRNTargetManager.load_from_excel(filepath)
           
            # Update UI in main thread
            self.after(0, lambda: self.on_target_loaded(loading, success, message, target))
       
        # ⭐ Run in background thread
        thread = threading.Thread(target=load_target)
        thread.daemon = True
        thread.start()
   
    def on_target_loaded(self, loading, success, message, target):
        """Handle target loaded completion"""
        loading.destroy()
       
        if success:
            self.grn_target_label.config(text=f"Target: {target:.1f}")
            self.logger.info(f"GRN Target loaded: {target}")
            messagebox.showinfo(
                "✅ Target Loaded",
                f"GRN Target loaded successfully!\n\n"
                f"📊 Target Value: {target:.1f}\n"
                f"📁 File: {message}\n\n"
                f"💡 This target will be used for ALL companies (SK, Maximus, EVPL)\n"
                f"📈 Scoring: > Target = 25, == Target = 20 (Grace), < Target = 0"
            )
            # Refresh data to apply new target
            self.refresh_all_data()
        else:
            self.logger.error(f"Failed to load GRN target: {message}")
            messagebox.showerror(
                "❌ Error",
                f"Failed to load GRN target:\n\n{message}\n\n"
                f"Please ensure the Excel file contains a numeric value."
            )
   
    def update_schedule_indicator(self):
        triggers = self.email_scheduler.get_active_triggers()
        if triggers:
            times = [t.get('time', '') for t in triggers]
            self.schedule_indicator.config(
                text=f"⏰ Active Triggers: {', '.join(times)}",
                fg=Config.COLORS["success"]
            )
        else:
            self.schedule_indicator.config(
                text="⏰ No active triggers",
                fg=Config.COLORS["gray"]
            )
   
    def open_schedule_dialog(self):
        if not self.all_data or not any(not df.empty for df in self.all_data.values()):
            messagebox.showwarning("No Data", "Please load data before scheduling reports.\n\nClick 'Refresh All' to load data.")
            return
       
        ScheduleEmailDialog(self, self.company, self.get_all_data, self.email_scheduler)
        self.update_schedule_indicator()
   
    def get_all_data(self):
        return self.all_data
   
    def show_calendar(self):
        calendar_window = tk.Toplevel(self)
        calendar_window.title("Select Date")
        calendar_window.configure(bg=Config.COLORS["white"])
        calendar_window.resizable(False, False)
        calendar_window.geometry("500x480")
       
        x = self.winfo_x() + self.winfo_width()//2 - 250
        y = self.winfo_y() + self.winfo_height()//2 - 240
        calendar_window.geometry(f"500x480+{x}+{y}")
       
        try:
            current_date = datetime.strptime(self.selected_date, "%d%m%Y")
        except:
            current_date = datetime.now()
       
        current_year = current_date.year
        current_month = current_date.month
        temp_selected_date = current_date
       
        main_container = tk.Frame(calendar_window, bg=Config.COLORS["white"])
        main_container.pack(fill="both", expand=True, padx=15, pady=15)
       
        header_frame = tk.Frame(main_container, bg=Config.COLORS["primary"], height=70)
        header_frame.pack(fill="x", pady=(0, 15))
        header_frame.pack_propagate(False)
       
        def change_month(delta):
            nonlocal current_year, current_month
            if delta == -1 and current_month == 1:
                current_month = 12
                current_year -= 1
            elif delta == 1 and current_month == 12:
                current_month = 1
                current_year += 1
            else:
                current_month += delta
            update_calendar()
       
        btn_left = tk.Button(header_frame, text="◀", font=("Segoe UI", 14, "bold"),
                             bg=Config.COLORS["primary"], fg="white", relief="flat",
                             command=lambda: change_month(-1), cursor="hand2", padx=15)
        btn_left.pack(side="left", padx=20, pady=15)
       
        self.month_label = tk.Label(header_frame, text="", font=("Segoe UI", 16, "bold"),
                                    bg=Config.COLORS["primary"], fg="white")
        self.month_label.pack(side="left", expand=True)
       
        btn_right = tk.Button(header_frame, text="▶", font=("Segoe UI", 14, "bold"),
                              bg=Config.COLORS["primary"], fg="white", relief="flat",
                              command=lambda: change_month(1), cursor="hand2", padx=15)
        btn_right.pack(side="right", padx=20, pady=15)
       
        day_header_frame = tk.Frame(main_container, bg=Config.COLORS["white"])
        day_header_frame.pack(fill="x", pady=(10, 5))
       
        day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        day_colors = [
            Config.COLORS["secondary"], Config.COLORS["secondary"], Config.COLORS["secondary"],
            Config.COLORS["secondary"], Config.COLORS["secondary"], Config.COLORS["warning"], Config.COLORS["danger"]
        ]
       
        for col, (day, color) in enumerate(zip(day_names, day_colors)):
            tk.Label(day_header_frame, text=day, font=("Segoe UI", 11, "bold"),
                     bg=Config.COLORS["white"], fg=color, width=8, pady=8).grid(row=0, column=col, padx=2, pady=2)
       
        calendar_frame = tk.Frame(main_container, bg=Config.COLORS["white"])
        calendar_frame.pack(fill="both", expand=True, pady=10)
       
        date_cells = []
        for row in range(6):
            row_cells = []
            for col in range(7):
                cell_frame = tk.Frame(calendar_frame, bg=Config.COLORS["light"],
                                      relief="solid", bd=1, width=55, height=50)
                cell_frame.grid(row=row, column=col, padx=2, pady=2)
                cell_frame.grid_propagate(False)
                cell_label = tk.Label(cell_frame, text="", font=("Segoe UI", 11),
                                      bg=Config.COLORS["light"], fg=Config.COLORS["dark"])
                cell_label.pack(expand=True, fill="both")
                row_cells.append((cell_frame, cell_label))
            date_cells.append(row_cells)
       
        selected_date_label = tk.Label(main_container, text="", font=("Segoe UI", 11),
                                        bg=Config.COLORS["white"], fg=Config.COLORS["success"])
        selected_date_label.pack(pady=(10, 5))
       
        def update_calendar():
            self.month_label.config(text=f"{cal.month_name[current_month]} {current_year}")
            month_cal = cal.monthcalendar(current_year, current_month)
           
            for row in range(6):
                for col in range(7):
                    date_cells[row][col][1].config(text="", bg=Config.COLORS["light"], fg=Config.COLORS["dark"])
                    date_cells[row][col][0].config(bg=Config.COLORS["light"])
           
            for row in range(len(month_cal)):
                for col in range(7):
                    day = month_cal[row][col]
                    if day != 0:
                        date_obj = datetime(current_year, current_month, day)
                        cell_frame, cell_label = date_cells[row][col]
                        cell_label.config(text=str(day))
                       
                        weekday = date_obj.weekday()
                        is_weekend = weekday >= 5
                       
                        if temp_selected_date and date_obj.date() == temp_selected_date.date():
                            cell_frame.config(bg=Config.COLORS["success"])
                            cell_label.config(bg=Config.COLORS["success"], fg="white", font=("Segoe UI", 11, "bold"))
                        elif date_obj.date() == datetime.now().date():
                            cell_frame.config(bg=Config.COLORS["info"])
                            cell_label.config(bg=Config.COLORS["info"], fg="white", font=("Segoe UI", 11, "bold"))
                        elif is_weekend:
                            if weekday == 5:
                                cell_frame.config(bg=Config.COLORS["warning"])
                                cell_label.config(bg=Config.COLORS["warning"], fg="white")
                            else:
                                cell_frame.config(bg=Config.COLORS["danger"])
                                cell_label.config(bg=Config.COLORS["danger"], fg="white")
                        else:
                            cell_frame.config(bg=Config.COLORS["light"])
                            cell_label.config(bg=Config.COLORS["light"], fg=Config.COLORS["dark"])
                       
                        cell_frame.bind("<Button-1>", lambda e, d=date_obj: select_date(d))
                        cell_label.bind("<Button-1>", lambda e, d=date_obj: select_date(d))
       
        def select_date(date):
            nonlocal temp_selected_date
            temp_selected_date = date
            update_calendar()
            selected_date_label.config(text=f"✓ Selected: {date.strftime('%A, %d %B %Y')}")
           
            date_sql_format = date.strftime("%d%m%Y")
            date_display_format = date.strftime("%d/%m/%Y")
           
            self.selected_date = date_sql_format
            self.selected_date_display = date_display_format
            self.date_var.set(date_display_format)
           
            calendar_window.after(500, calendar_window.destroy)
            self.refresh_current_tab_data()
            self.update_status(f"Loading all data for {date_display_format}...")
            self.load_all_modules()
       
        button_frame = tk.Frame(main_container, bg=Config.COLORS["white"])
        button_frame.pack(fill="x", pady=(10, 0))
       
        close_btn = ModernButton(button_frame, text="Close", command=calendar_window.destroy,
                                variant="secondary", width=10)
        close_btn.pack(side="right", padx=5)
       
        today_btn = ModernButton(button_frame, text="Today",
                                command=lambda: select_date(datetime.now()),
                                variant="primary", width=10)
        today_btn.pack(side="right", padx=5)
       
        update_calendar()
   
    def refresh_current_tab_data(self):
        current_tab_index = self.notebook.index(self.notebook.select())
        if current_tab_index < len(Config.MODULES):
            module_name = Config.MODULES[current_tab_index]['name']
           
            if module_name in self.tabs:
                tab_data = self.tabs[module_name]
                table = tab_data['table']
                for item in table.tree.get_children():
                    table.tree.delete(item)
           
            try:
                new_data = self.load_module_data(module_name)
                if module_name in self.all_data:
                    self.all_data[module_name] = new_data
                self.display_data_in_tab(module_name, new_data)
                self.update_status(f"✓ Refreshed {module_name} for {self.selected_date_display}")
            except Exception as e:
                self.logger.error(f"Failed to refresh {module_name}: {e}")
                self.update_status(f"�    def create_main_content(self):
        self.main_content = tk.Frame(self, bg=Config.COLORS["light"])
        self.main_content.pack(fill="both", expand=True)
        
        # --- Custom Sidebar ---
        self.sidebar_frame = tk.Frame(self.main_content, bg=Config.COLORS["white"], width=220)
        self.sidebar_frame.pack(side="left", fill="y")
        self.sidebar_frame.pack_propagate(False)
        
        # App Title in Sidebar
        tk.Label(self.sidebar_frame, text=f"{Config.COMPANY_LOGO} FOCUS PRISM", 
                 font=("Segoe UI", 14, "bold"), bg=Config.COLORS["white"], fg=Config.COLORS["primary"]).pack(pady=(30, 20))
                 
        tk.Frame(self.sidebar_frame, bg=Config.COLORS["border"], height=1).pack(fill="x", padx=15, pady=(0, 20))
        
        # Right Content Area
        self.content_frame = tk.Frame(self.main_content, bg=Config.COLORS["light"])
        self.content_frame.pack(side="right", fill="both", expand=True)
        
        self.notebook = ttk.Notebook(self.content_frame)
        self.notebook.pack(fill="both", expand=True, padx=25, pady=(25, 25))
        
        style = ttk.Style()
        style.configure("TNotebook", background=Config.COLORS["light"])
        style.layout("TNotebook.Tab", []) # Hide default tabs completely
        
        from ui_components import SidebarButton
        self.tabs = {}
        self.tab_buttons = []
        for i, module in enumerate(Config.MODULES):
            tab_frame = tk.Frame(self.notebook, bg=Config.COLORS["white"])
            self.notebook.add(tab_frame, text=module['name'])
            
            btn = SidebarButton(
                self.sidebar_frame,
                text=f"  {module['icon']}   {module['name']}",
                command=lambda idx=i: self.select_tab(idx)
            )
            btn.pack(fill="x", pady=2)
            self.tab_buttons.append(btn)
            
            # ⭐ No hidden columns - emails are completely removed
            self.tabs[module['name']] = self.create_enhanced_tab(tab_frame, module['name'])
        
        # Select first tab by default
        if self.tab_buttons:
            self.tab_buttons[0].set_active(True)me']] = self.create_enhanced_tab(tab_frame, module['name'])
        
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
       
        self.stats_frame = tk.Frame(self.main_content, bg=Config.COLORS["white"], height=100)
        self.stats_frame.pack(fill="x", padx=25, pady=(0, 25))
        self.stats_frame.pack_propagate(False)
       
        self.stats_labels = {}
        stats = ["📊 Total Records", "🕐 Last Updated", "✅ Status"]
        for i, stat in enumerate(stats):
            frame = tk.Frame(self.stats_frame, bg=Config.COLORS["white"])
            frame.pack(side="left", expand=True, fill="both", padx=15)
            tk.Label(frame, text=stat, font=("Segoe UI", 10), bg=Config.COLORS["white"], fg=Config.COLORS["gray"]).pack(pady=(15, 5))
            self.stats_labels[stat] = tk.Label(frame, text="--", font=("Segoe UI", 14, "bold"), bg=Config.COLORS["white"], fg=Config.COLORS["primary"])
            self.stats_labels[stat].pack()
   
    def create_enhanced_tab(self, parent, module_name):
        container = tk.Frame(parent, bg=Config.COLORS["white"])
        container.pack(fill="both", expand=True, padx=5, pady=5)
       
        button_frame = tk.Frame(container, bg=Config.COLORS["white"])
        button_frame.pack(fill="x", pady=(5, 10))
       
        left_button_frame = tk.Frame(button_frame, bg=Config.COLORS["white"])
        left_button_frame.pack(side="left")
       
        filter_btn = ModernButton(left_button_frame, text="🔧 Advanced Filters",
                                   command=lambda: self.show_advanced_filters(module_name), variant="primary", width=15)
        filter_btn.pack(side="left", padx=5)
       
        clear_filter_btn = ModernButton(left_button_frame, text="✗ Clear Filters",
                                         command=lambda: self.clear_filters(module_name), variant="warning", width=12)
        clear_filter_btn.pack(side="left", padx=5)
       
        rules_btn = ModernButton(
            button_frame,
            text="ℹ️ Rules",
            command=lambda: self.show_module_rules(module_name),
            variant="info",
            width=10
        )
        rules_btn.pack(side="right", padx=5)
       
        table = EnhancedTable(
            container,
            on_select_callback=self.on_row_selected
        )
        table.pack(fill="both", expand=True)
       
        return {'table': table, 'module_name': module_name, 'current_data': None}
   
    def show_module_rules(self, module_name):
        rules = {
            "Store Data": "📅 Store Data shows data for the SELECTED DATE only",
            "Actual Sales": "📅 Actual Sales shows data for the SELECTED DATE only",
            "Overall Estimate": "⚠️ Overall Estimate shows data for the NEXT DAY only",
            "Sessionwise Estimate": "⚠️ Sessionwise Estimate shows data for the NEXT DAY only",
            "Score": f"📅 Score shows data for the SELECTED DATE only\n\n💡 GRN Target: {GRNTargetManager.get_target():.1f}\n\n📊 GRN Scoring:\n• Today GRN > Target → 25 marks\n• Today GRN == Target → 20 marks (Grace)\n• Today GRN < Target → 0 marks\n\n📊 PSERM (Closing Stock) Scoring:\n• PSERM > 35 → 25 marks\n• 30-35 → 20 marks\n• 25-29 → 15 marks\n• 20-24 → 10 marks\n• < 20 → 0 marks"
        }
       
        rule = rules.get(module_name, f"ℹ️ {module_name} shows data for the SELECTED DATE only")
        messagebox.showinfo(f"📌 {module_name}", rule)
   
    def show_advanced_filters(self, module_name):
        if module_name in self.tabs:
            table = self.tabs[module_name]['table']
            if table.filtered_data is not None and not table.filtered_data.empty:
                columns = list(table.filtered_data.columns)
                AdvancedFilterDialog(self, columns,
                    lambda filters: self.apply_advanced_filters(module_name, filters),
                    lambda: self.clear_filters(module_name))
   
    def apply_advanced_filters(self, module_name, filters):
        if module_name in self.tabs:
            self.tabs[module_name]['table'].apply_advanced_filters(filters)
            self.update_status(f"Applied filters to {module_name}")
   
    def clear_filters(self, module_name):
        if module_name in self.tabs:
            self.tabs[module_name]['table'].clear_filters()
            self.update_status(f"Cleared all filters for {module_name}")
   
    def on_row_selected(self, row_values, pu_name=None):
        if pu_name and str(pu_name).strip():
            pu_name = str(pu_name).strip()
           
            if self.selected_pu_name == pu_name:
                self.selected_pu_name = None
                self.update_status(f"🔵 Deselected: {pu_name}")
            else:
                self.selected_pu_name = pu_name
                self.update_status(f"🔵 Selected: {self.selected_pu_name}")
           
            for module in Config.MODULES:
                module_name = module['name']
                if module_name in self.all_data and module_name in self.tabs:
                    self.display_data_in_tab(module_name, self.all_data[module_name])
        else:
            if self.selected_pu_name:
                self.selected_pu_name = None
                for module in Config.MODULES:
                    module_name = module['name']
                    if module_name in self.all_data and module_name in self.tabs:
                        self.display_data_in_tab(module_name, self.all_data[module_name])
                self.update_status("Ready")
   
    def create_status_bar(self):
        self.status_bar = tk.Frame(self, bg=Config.COLORS["primary"], height=35)
        self.status_bar.pack(side="bottom", fill="x")
        self.status_bar.pack_propagate(False)
       
        self.status_label = tk.Label(self.status_bar, text="Ready", font=("Segoe UI", 10),
                                      bg=Config.COLORS["primary"], fg=Config.COLORS["white"], anchor="w")
        self.status_label.pack(side="left", padx=15)
       
        self.time_label = tk.Label(self.status_bar, text=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    font=("Segoe UI", 10), bg=Config.COLORS["primary"], fg=Config.COLORS["white"])
        self.time_label.pack(side="right", padx=15)
       
        self.update_clock()
   
    def update_clock(self):
        self.time_label.config(text=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.after(1000, self.update_clock)
        
    def select_tab(self, idx):
        self.notebook.select(idx)
        # Update button visuals if needed (optional)
        # For a truly flat look, we can keep them colorful all the time
        
    def on_tab_changed(self, event):
        current_tab_index = self.notebook.index(self.notebook.select())
        if current_tab_index < len(Config.MODULES):
            module_name = Config.MODULES[current_tab_index]['name']
           
            if module_name in ["Overall Estimate", "Sessionwise Estimate"]:
                self.update_status(f"Viewing {module_name} module - Showing NEXT DAY data")
            else:
                self.update_status(f"Viewing {module_name} module - Showing SELECTED DATE data")
           
            if module_name in self.all_data and self.all_data[module_name] is not None:
                self.display_data_in_tab(module_name, self.all_data[module_name])
   
    def load_all_modules(self):
        if self.is_loading:
            return
       
        self.is_loading = True
        self.all_data = {}
        self.progress_bar['value'] = 0
        self.progress_bar['maximum'] = 5
        loaded_count = 0
       
        self.refresh_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.schedule_btn.config(state="disabled")
        self.upload_grn_btn.config(state="disabled")
       
        def load_in_background():
            nonlocal loaded_count
            for module in Config.MODULES:
                try:
                    data = self.load_module_data(module['name'])
                    self.all_data[module['name']] = data
                    loaded_count += 1
                    self.after(0, lambda: self.update_progress(loaded_count, module['name'], data))
                except Exception as e:
                    self.logger.error(f"Failed to load {module['name']}: {e}")
                    self.after(0, lambda m=module['name']: self.update_status(f"Error loading {m}", is_error=True))
            self.after(0, self.on_all_modules_loaded)
       
        thread = threading.Thread(target=load_in_background)
        thread.daemon = True
        thread.start()
   
    def load_module_data(self, module_name):
        query_builder = QueryBuilder(self.company)
       
        query_name_map = {
            "Store Data": "Store",
            "Overall Estimate": "Overall",
            "Sessionwise Estimate": "SessionWise",
            "Actual Sales": "Sale",
            "Score": "Score"
        }
       
        query_name = query_name_map.get(module_name, module_name)
       
        if module_name in ["Overall Estimate", "Sessionwise Estimate"]:
            current_date = datetime.strptime(self.selected_date, "%d%m%Y")
            next_date = current_date + timedelta(days=1)
            date_to_use = next_date.strftime("%d%m%Y")
            self.logger.info(f"{module_name} using next day date: {date_to_use} (selected was {self.selected_date})")
        else:
            date_to_use = self.selected_date
       
        sql = query_builder.get_query(query_name, date_to_use)
       
        try:
            df = self.db_handler.execute_query(sql, Config.DATABASES[self.company])
            # ⭐ No email columns to merge - they are completely removed
            if module_name == "Score":
                df = self.rename_score_columns(df)
            return df
        except Exception as e:
            self.logger.error(f"Query failed for {module_name}: {e}")
            return pd.DataFrame()
   
    def rename_score_columns(self, df):
        column_renames = {
            'GRN_Mark': 'GRN (25)',
            'Closing_Stock_Mark': 'Closing Stock (25)',
            'Estimate_Overall_Mark': 'Estimate Overall (15)',
            'SessionWise_Mark': 'Session Wise (15)',
            'Sales_Mark': 'Sales (20)',
            'Overall_Result': 'Score (100)'
        }
        return df.rename(columns=column_renames)
   
    def update_progress(self, loaded_count, module_name, data):
        progress = (loaded_count / 5) * 100
        self.progress_bar['value'] = progress
        self.progress_label.config(text=f"{loaded_count}/5 Loaded")
       
        current_tab_index = self.notebook.index(self.notebook.select())
        current_module = Config.MODULES[current_tab_index]['name'] if current_tab_index < len(Config.MODULES) else None
       
        if module_name == current_module:
            self.display_data_in_tab(module_name, data)
       
        self.update_status(f"Loaded {module_name} - {len(data)} records")
   
    def on_all_modules_loaded(self):
        self.update_status("All modules loaded successfully!")
        self.refresh_btn.config(state="normal")
        self.export_btn.config(state="normal")
        self.schedule_btn.config(state="normal")
        self.upload_grn_btn.config(state="normal")
        self.is_loading = False
       
        total_records = sum(len(df) for df in self.all_data.values())
        self.stats_labels["📊 Total Records"].config(text=f"{total_records:,}")
        self.stats_labels["🕐 Last Updated"].config(text=datetime.now().strftime("%H:%M:%S"))
        self.stats_labels["✅ Status"].config(text="Ready", fg=Config.COLORS["success"])
       
        current_tab_index = self.notebook.index(self.notebook.select())
        if current_tab_index < len(Config.MODULES):
            module_name = Config.MODULES[current_tab_index]['name']
            if module_name in self.all_data:
                self.display_data_in_tab(module_name, self.all_data[module_name])
       
        self.update_schedule_indicator()
   
    def display_data_in_tab(self, module_name, df):
        if module_name not in self.tabs:
            return
       
        table = self.tabs[module_name]['table']
        table.load_data(df, self.selected_pu_name)
   
    def refresh_all_data(self):
        if not self.is_loading:
            self.all_data.clear()
            self.load_all_modules()
   
    def export_to_excel(self):
        if not self.all_data:
            messagebox.showwarning("No Data", "Please load data before exporting")
            return
       
        loading = LoadingOverlay(self, "Preparing export...")
       
        def export():
            export_error = None
            filepath = None
            try:
                if sys.platform.startswith("win"):
                    downloads = os.path.join(os.environ["USERPROFILE"], "Downloads")
                else:
                    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
               
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"{self.company.lower()}_report_{self.selected_date}_{timestamp}.xlsx"
                filepath = os.path.join(downloads, filename)
               
                with pd.ExcelWriter(filepath, engine="openpyxl", mode='w') as writer:
                    for module_name, df in self.all_data.items():
                        if not df.empty:
                            # ⭐ No email columns to remove - they don't exist
                            df_to_export = df.copy()
                            if module_name == "Score":
                                df_to_export = self.rename_score_columns(df_to_export)
                           
                            df_to_export.to_excel(writer, sheet_name=module_name, index=False)
                           
                            worksheet = writer.sheets[module_name]
                            ExcelFormatter.apply_clean_formatting(worksheet, df_to_export)
                   
                    # Scoring Rules - Updated with new PSERM logic
                    rules_df = pd.DataFrame(Config.SCORING_RULES)
                    rules_df.to_excel(writer, sheet_name="Scoring_Rules", index=False, header=False)
                   
                    worksheet = writer.sheets["Scoring_Rules"]
                    worksheet.column_dimensions['A'].width = 35
                    worksheet.column_dimensions['B'].width = 25
                    worksheet.column_dimensions['C'].width = 15
               
            except Exception as e:
                export_error = str(e)
                self.logger.error(f"Export failed: {e}")
           
            if export_error:
                self.after(0, lambda: self.on_export_error(loading, export_error))
            else:
                self.after(0, lambda: self.on_export_complete(loading, filepath))
       
        thread = threading.Thread(target=export)
        thread.daemon = True
        thread.start()
   
    def on_export_complete(self, loading, filepath):
        loading.destroy()
        messagebox.showinfo(
            "Export Successful",
            f"Report exported successfully!\n\nLocation: {filepath}\n\nSheets included:\n• All data modules\n• Scoring Rules (GRN Target: {GRNTargetManager.get_target():.1f})\n• PSERM Logic: >35 = 25 marks, 30-35 = 20 marks, 25-29 = 15 marks, 20-24 = 10 marks, <20 = 0 marks\n\n📊 Clean formatting applied (No gridlines)"
        )
        self.update_status(f"Export completed: {filepath}")
   
    def on_export_error(self, loading, error):
        loading.destroy()
        messagebox.showerror("Export Failed", f"Failed to export data:\n{error}")
        self.update_status(f"Export failed: {error}", is_error=True)
   
    def update_status(self, message, is_error=False):
        self.status_label.config(text=message, fg=Config.COLORS["danger"] if is_error else Config.COLORS["white"])
        if is_error:
            self.after(3000, lambda: self.status_label.config(fg=Config.COLORS["white"]))
   
    def load_initial_data(self):
        self.after(100, self.load_all_modules)
   
    def logout(self):
        if messagebox.askyesno("Logout", "Are you sure you want to logout?"):
            self.logger.info(f"User logged out from {self.company}")
            self.email_scheduler.clear_schedule()
            self.destroy()
            self.login_page.deiconify()
            self.login_page.attributes('-fullscreen', True)
   
    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit the application?"):
            self.logger.info("Application closed by user")
            self.email_scheduler.clear_schedule()
            self.db_handler.close_all_connections()
            self.destroy()
            self.login_page.quit()


