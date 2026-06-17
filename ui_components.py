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

# ==================== MODERN UI COMPONENTS ====================

class ModernButton(tk.Button):
    def __init__(self, parent, text="", command=None, variant="primary", **kwargs):
        self.variant = variant
        self.colors = {
            "primary": (Config.COLORS["primary"], Config.COLORS["hover"]),
            "success": (Config.COLORS["success"], "#475927"),      # Darker Olive
            "danger": (Config.COLORS["danger"], "#6B0000"),        # Darker Red
            "warning": (Config.COLORS["warning"], "#B35A1A"),      # Darker Orange
            "secondary": (Config.COLORS["secondary"], "#8C6A50"),  # Brown
            "info": (Config.COLORS["info"], "#3B6E99"),            # Darker Steel Blue
            "schedule": (Config.COLORS["warning"], "#B35A1A"),     
            "upload": (Config.COLORS["primary"], Config.COLORS["hover"]),       
            "clear": ("#8C8C8C", "#6E6E6E"),
            "history": (Config.COLORS["primary"], Config.COLORS["hover"]),
            "toggle_on": (Config.COLORS["success"], "#475927"),
            "toggle_off": (Config.COLORS["gray"], "#6E6E6E"),
        }
       
        bg_color, hover_color = self.colors.get(variant, self.colors["primary"])
       
        # If font is passed in kwargs, use it; otherwise use default
        current_font = kwargs.pop("font", ("Segoe UI", 10, "bold"))
        
        super().__init__(
            parent,
            text=text,
            command=command,
            bg=bg_color,
            fg=Config.COLORS["white"],
            font=current_font,
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=20,
            pady=8,
            **kwargs
        )
       
        self.bind("<Enter>", lambda e: self.config(bg=hover_color))
        self.bind("<Leave>", lambda e: self.config(bg=bg_color))


class SidebarButton(tk.Button):
    def __init__(self, parent, text="", command=None, is_active=False, **kwargs):
        self.is_active = is_active
        self.default_bg = Config.COLORS["light"]
        self.hover_bg = Config.COLORS["border"]
        self.active_bg = Config.COLORS["secondary"]
        self.active_fg = Config.COLORS["white"]
        self.default_fg = Config.COLORS["primary"]
        
        bg_color = self.active_bg if is_active else self.default_bg
        fg_color = self.active_fg if is_active else self.default_fg
        
        super().__init__(
            parent,
            text=text,
            command=command,
            bg=bg_color,
            fg=fg_color,
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            bd=0,
            cursor="hand2",
            padx=20,
            pady=12,
            anchor="w",
            **kwargs
        )
        
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def on_enter(self, event):
        if not self.is_active:
            self.config(bg=self.hover_bg)

    def on_leave(self, event):
        if not self.is_active:
            self.config(bg=self.default_bg)

    def set_active(self, active):
        self.is_active = active
        self.config(
            bg=self.active_bg if active else self.default_bg,
            fg=self.active_fg if active else self.default_fg
        )


class ModernCard(tk.Frame):
    def __init__(self, parent, title="", **kwargs):
        super().__init__(parent, bg=Config.COLORS["white"], relief="flat", bd=1)
        self.configure(highlightbackground=Config.COLORS["border"], highlightthickness=1)
       
        if title:
            self.title_frame = tk.Frame(self, bg=Config.COLORS["white"])
            self.title_frame.pack(fill="x", padx=20, pady=(15, 10))
           
            tk.Label(
                self.title_frame,
                text=title,
                font=("Segoe UI", 14, "bold"),
                bg=Config.COLORS["white"],
                fg=Config.COLORS["primary"]
            ).pack(side="left")
       
        self.content_frame = tk.Frame(self, bg=Config.COLORS["white"])
        self.content_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))


class ModernEntry(tk.Frame):
    def __init__(self, parent, label="", placeholder="", **kwargs):
        super().__init__(parent, bg=Config.COLORS["white"])
       
        self.label = tk.Label(
            self,
            text=label,
            font=("Segoe UI", 10),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["secondary"]
        )
        self.label.pack(anchor="w", pady=(0, 5))
       
        self.entry = tk.Entry(
            self,
            font=("Segoe UI", 11),
            bg=Config.COLORS["light"],
            fg=Config.COLORS["dark"],
            relief="solid",
            bd=1,
            highlightthickness=0,
            **kwargs
        )
        self.entry.pack(fill="x", ipady=8)
       
        self.placeholder = placeholder
        self.original_fg = Config.COLORS["dark"]
        self.placeholder_fg = Config.COLORS["gray"]
       
        if placeholder:
            self.set_placeholder()
            self.entry.bind("<FocusIn>", self.on_focus_in)
            self.entry.bind("<FocusOut>", self.on_focus_out)
   
    def set_placeholder(self):
        if not self.entry.get():
            self.entry.insert(0, self.placeholder)
            self.entry.config(fg=self.placeholder_fg)
   
    def on_focus_in(self, event):
        if self.entry.get() == self.placeholder:
            self.entry.delete(0, tk.END)
            self.entry.config(fg=self.original_fg)
   
    def on_focus_out(self, event):
        if not self.entry.get():
            self.entry.insert(0, self.placeholder)
            self.entry.config(fg=self.placeholder_fg)
   
    def get(self):
        value = self.entry.get()
        return "" if value == self.placeholder else value
   
    def set(self, value):
        self.entry.delete(0, tk.END)
        self.entry.insert(0, value)
        self.entry.config(fg=self.original_fg)


class LoadingOverlay(tk.Toplevel):
    def __init__(self, parent, message="Loading..."):
        super().__init__(parent)
        self.title("")
        self.configure(bg=Config.COLORS["white"])
        self.overrideredirect(True)
       
        x = parent.winfo_x() + parent.winfo_width()//2 - 150
        y = parent.winfo_y() + parent.winfo_height()//2 - 100
        self.geometry(f"300x200+{x}+{y}")
       
        self.lift()
        self.attributes('-topmost', True)
       
        self.canvas = tk.Canvas(self, width=80, height=80, bg=Config.COLORS["white"], highlightthickness=0)
        self.canvas.pack(pady=30)
       
        self.arc = self.canvas.create_arc(10, 10, 70, 70, start=0, extent=120, fill=Config.COLORS["accent"], outline="")
        self.triangle = self.canvas.create_polygon(35, 5, 45, 15, 35, 25, fill=Config.COLORS["accent"])
       
        tk.Label(
            self,
            text=message,
            font=("Segoe UI", 12),
            bg=Config.COLORS["white"],
            fg=Config.COLORS["secondary"]
        ).pack()
       
        self.angle = 0
        self.animate()
       
        self.transient(parent)
        self.grab_set()
   
    def animate(self):
        self.angle = (self.angle + 10) % 360
        self.canvas.delete("all")
        self.canvas.create_arc(10, 10, 70, 70, start=self.angle, extent=120, fill=Config.COLORS["accent"], outline="")
        self.canvas.create_polygon(35 + 30, 5, 45 + 30, 15, 35 + 30, 25, fill=Config.COLORS["accent"])
        self.after(50, self.animate)
   
    def destroy(self):
        self.after_cancel(self.animate)
        super().destroy()


