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
       
        # Removed both vertical and horizontal scrollbars as requested
        # self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        # self.hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        # self.tree.configure(xscrollcommand=self.hsb.set)
       
        self.tree.grid(row=0, column=0, sticky="nsew")
        # self.vsb.grid(row=0, column=1, sticky="ns")
        # self.hsb.grid(row=1, column=0, sticky="ew")
       
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)
       
        self.tree.bind("<ButtonRelease-1>", self.on_row_select)
        self.tree.bind("<Button-1>", self.on_header_click)
        
        style = ttk.Style()
        style.configure("Enhanced.Treeview",
            font=("Segoe UI", 9),
            rowheight=26,
            background=Config.COLORS["white"],
            foreground=Config.COLORS["dark"],
            fieldbackground=Config.COLORS["white"],
            borderwidth=0,
            relief="flat"
        )
        style.configure("Enhanced.Treeview.Heading",
            font=("Segoe UI", 11, "bold"),
            background=Config.COLORS["primary"],
            foreground=Config.COLORS["white"],
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
            background=[('active', Config.COLORS["hover"])]
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
            header_text = col + " ▼"
            self.tree.heading(col, text=header_text)
            self.tree.column(col, anchor="center", minwidth=20, stretch=True)
       
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
        
        # Dynamically set height to exactly match the data to avoid excess empty rows
        num_rows = len(self.filtered_data)
        self.tree.configure(height=max(1, num_rows))
        
        self.auto_adjust_columns()
        
        # Adjust the treeview height to show all rows without vertical scrolling
        total_rows = len(self.filtered_data)
        if total_rows > 0:
            self.tree.configure(height=total_rows)
   
    def auto_adjust_columns(self):
        for col in self.tree["columns"]:
            max_width = len(str(col)) * 12 + 20
            for idx, item in enumerate(self.tree.get_children()):
                if idx > 100:
                    break
                values = self.tree.item(item)["values"]
                if values and self.tree["columns"].index(col) < len(values):
                    value = values[self.tree["columns"].index(col)]
                    if value:
                        width = len(str(value)) * 9
                        max_width = max(max_width, width)
            final_width = min(max_width + 10, 200)
            final_width = max(final_width, 30)
            self.tree.column(col, width=final_width, stretch=True)
            
    def on_header_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col_id = self.tree.identify_column(event.x)
            if col_id:
                col_idx = int(col_id.replace('#', '')) - 1
                col_name = self.tree["columns"][col_idx]
                
                self.open_column_filter(col_name, event.x_root, event.y_root)

    def open_column_filter(self, col_name, x, y):
        from advancedfilterdialog import ColumnFilterDialog
        existing = self.column_filters.get(col_name, None)
        ColumnFilterDialog(self, col_name, x, y, self.apply_column_filter, self.clear_column_filter, existing)

    def apply_column_filter(self, col_name, filter_data):
        self.column_filters[col_name] = filter_data
        self.apply_all_filters()
        
    def clear_column_filter(self, col_name):
        if col_name in self.column_filters:
            del self.column_filters[col_name]
            self.apply_all_filters()
            
    def _apply_condition(self, df, col, op, val):
        if not op:
            return pd.Series(True, index=df.index)
            
        op = op.lower()
        if "null" in op:
            if op == "is null":
                return df[col].isna() | (df[col] == "") | (df[col] == "nan") | (df[col] == "None")
            else:
                return df[col].notna() & (df[col] != "") & (df[col] != "nan") & (df[col] != "None")
                
        try:
            # Try numeric comparison first
            numeric_col = pd.to_numeric(df[col], errors='coerce')
            numeric_val = float(val)
            
            if op == "is equal to": return numeric_col == numeric_val
            elif op == "is not equal to": return numeric_col != numeric_val
            elif op == "is greater than": return numeric_col > numeric_val
            elif op == "is greater than or equal to": return numeric_col >= numeric_val
            elif op == "is less than": return numeric_col < numeric_val
            elif op == "is less than or equal to": return numeric_col <= numeric_val
        except:
            # Fallback to string comparison with whitespace stripped
            str_col = df[col].astype(str).str.strip().str.lower()
            val = str(val).strip().lower()
            
            if op == "is equal to": return str_col == val
            elif op == "is not equal to": return str_col != val
            elif op == "contains": return str_col.str.contains(val, na=False, regex=False)
            elif op == "does not contain": return ~str_col.str.contains(val, na=False, regex=False)
            elif op == "starts with": return str_col.str.startswith(val, na=False)
            elif op == "ends with": return str_col.str.endswith(val, na=False)
            elif op == "is greater than": return str_col > val
            elif op == "is greater than or equal to": return str_col >= val
            elif op == "is less than": return str_col < val
            elif op == "is less than or equal to": return str_col <= val
            
        return pd.Series(True, index=df.index)

    def apply_all_filters(self):
        display_df = self.data.copy()
        
        for col, f in self.column_filters.items():
            if col not in display_df.columns:
                continue
                
            mask1 = self._apply_condition(display_df, col, f.get("op1"), f.get("val1"))
            
            if "op2" in f and f.get("op2"):
                mask2 = self._apply_condition(display_df, col, f.get("op2"), f.get("val2"))
                if f.get("logic") == "And":
                    display_df = display_df[mask1 & mask2]
                else:
                    display_df = display_df[mask1 | mask2]
            else:
                display_df = display_df[mask1]
                
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


