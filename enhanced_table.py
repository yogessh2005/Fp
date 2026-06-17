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

# ==================== ENHANCED TABLE ====================

class EnhancedTable(ttk.Frame):
    """Enhanced table widget with toggle selection"""
   
    def __init__(self, parent, on_select_callback=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.on_select_callback = on_select_callback
        self.data = pd.DataFrame()
        self.filtered_data = pd.DataFrame()
        self.column_filters = {}
        self.selected_row = None
        self.selected_pu_name = None
        self.is_selected = False
       
        self.setup_ui()
   
    def setup_ui(self):
        self.tree_frame = tk.Frame(self, bg=Config.COLORS["white"])
        self.tree_frame.pack(fill="both", expand=True)
       
        self.tree = ttk.Treeview(
            self.tree_frame,
            show="headings",
            selectmode="browse"
        )
       
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
       
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")
       
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)
       
        self.tree.bind("<ButtonRelease-1>", self.on_row_select)
       
        style = ttk.Style()
        style.configure("Enhanced.Treeview",
            font=("Segoe UI", 10),
            rowheight=40,
            background=Config.COLORS["white"],
            foreground=Config.COLORS["dark"],
            fieldbackground=Config.COLORS["white"],
            borderwidth=0,
            relief="flat"
        )
        style.configure("Enhanced.Treeview.Heading",
            font=("Segoe UI", 11, "bold"),
            background=Config.COLORS["light"],
            foreground=Config.COLORS["primary"],
            relief="flat",
            borderwidth=0,
            padding=10
        )
        style.configure("Enhanced.Treeview",
            bordercolor=Config.COLORS["white"],
            lightcolor=Config.COLORS["white"],
            darkcolor=Config.COLORS["white"]
        )
        style.map("Enhanced.Treeview",
            background=[('selected', Config.COLORS["selected_row"])],
            foreground=[('selected', Config.COLORS["primary"])]
        )
        style.map("Enhanced.Treeview.Heading",
            background=[('active', Config.COLORS["border"])]
        )
        
        self.tree.tag_configure('selected_row', background=Config.COLORS["selected_row"], foreground=Config.COLORS["primary"])
        self.tree.tag_configure('evenrow', background=Config.COLORS["white"])
        self.tree.tag_configure('oddrow', background=Config.COLORS["alternate_row"])
        
        self.tree.configure(style="Enhanced.Treeview")
   
    def load_data(self, df: pd.DataFrame, selected_pu_name: str = None):
        self.data = df.copy()
        self.selected_pu_name = selected_pu_name
       
        display_df = df.copy()
       
        if selected_pu_name and 'PUName' in display_df.columns:
            selected_mask = display_df['PUName'].astype(str).str.strip() == selected_pu_name
            selected_rows = display_df[selected_mask].copy()
            other_rows = display_df[~selected_mask].copy()
           
            if not selected_rows.empty:
                self.filtered_data = pd.concat([selected_rows, other_rows], ignore_index=True)
                self.is_selected = True
            else:
                self.filtered_data = display_df.copy()
                self.is_selected = False
        else:
            self.filtered_data = display_df.copy()
            self.is_selected = False
       
        self.column_filters = {}
        self.refresh_display()
   
    def refresh_display(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
       
        if self.filtered_data.empty:
            return
       
        columns = list(self.filtered_data.columns)
        self.tree["columns"] = columns
       
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", minwidth=100)
       
        for idx, (_, row) in enumerate(self.filtered_data.iterrows()):
            values = [str(v) if pd.notna(v) else "" for v in row]
           
            if self.is_selected and 'PUName' in self.filtered_data.columns:
                pu_name = str(row.get('PUName', '')).strip()
                if pu_name == self.selected_pu_name and idx == 0:
                    tag = 'selected_row'
                else:
                    tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            else:
                tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
           
            self.tree.insert("", "end", values=values, tags=(tag,))
       
        self.auto_adjust_columns()
   
    def auto_adjust_columns(self):
        for col in self.tree["columns"]:
            max_width = len(str(col)) * 12
            for idx, item in enumerate(self.tree.get_children()):
                if idx > 100:
                    break
                values = self.tree.item(item)["values"]
                if values and self.tree["columns"].index(col) < len(values):
                    value = values[self.tree["columns"].index(col)]
                    if value:
                        width = len(str(value)) * 9
                        max_width = max(max_width, width)
            final_width = min(max_width + 30, 350)
            final_width = max(final_width, 100)
            self.tree.column(col, width=final_width)
   
    def apply_advanced_filters(self, filters):
        self.filtered_data = self.data.copy()
        display_df = self.data.copy()
       
        for col, filter_value in filters.items():
            if col in display_df.columns and filter_value:
                display_df = display_df[
                    display_df[col].astype(str).str.lower().str.contains(filter_value.lower(), na=False)
                ]
       
        self.filtered_data = display_df
        self.refresh_display()
   
    def clear_filters(self):
        self.filtered_data = self.data.copy()
        self.column_filters = {}
        self.refresh_display()
   
    def on_row_select(self, event):
        selection = self.tree.selection()
        if selection:
            self.selected_row = selection[0]
            values = self.tree.item(selection[0])["values"]
           
            pu_name = values[0] if values else None
           
            if self.on_select_callback:
                self.on_select_callback(values, pu_name)
   
    def get_full_data(self):
        return self.data
   
    def get_selected_pu(self):
        return self.selected_pu_name if self.is_selected else None


