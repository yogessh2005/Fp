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

from applogger import AppLogger
from config import Config
from databasehandler import DatabaseHandler
from mainapp import MainApp
from ui_components import ModernButton
from ui_components import ModernCard
from ui_components import ModernEntry

# ==================== LOGIN PAGE ====================

class LoginPage(tk.Tk):
    def __init__(self):
        super().__init__()
        self.logger = AppLogger()
        self.db_handler = DatabaseHandler(self.logger)
        self.current_app = None
       
        self.setup_window()
        self.create_gradient_background()
        self.create_login_interface()
       
        self.logger.info("Application started - Login page initialized")
   
    def setup_window(self):
        self.title(f"{Config.APP_NAME} - Login")
        self.attributes('-fullscreen', True)
        self.configure(bg=Config.COLORS["light"])
        self.bind("<Escape>", lambda e: self.attributes('-fullscreen', False))
        self.bind("<F11>", lambda e: self.toggle_fullscreen())
   
    def toggle_fullscreen(self):
        self.attributes('-fullscreen', not self.attributes('-fullscreen'))
   
    def create_gradient_background(self):
        self.configure(bg="#EBE7E0")

    def create_login_interface(self):
        # Main container to hold everything
        main_container = tk.Frame(self, bg="#EBE7E0", highlightthickness=0)
        main_container.place(relx=0.5, rely=0.5, anchor="center")
        
        # --- LOGO SECTION ---
        logo_frame = tk.Frame(main_container, bg="#EBE7E0")
        logo_frame.pack(pady=(0, 20))
        
        # Load and display the "logo.png" image
        logo_path = os.path.join(Config.BASE_DIR, "logo.png")
        try:
            from PIL import Image, ImageTk
            img = Image.open(logo_path).convert("RGBA")
            
            # Make white background transparent
            data = img.getdata()
            new_data = []
            tolerance = 40
            for item in data:
                if item[0] > 255 - tolerance and item[1] > 255 - tolerance and item[2] > 255 - tolerance:
                    new_data.append((255, 255, 255, 0))
                else:
                    new_data.append(item)
            img.putdata(new_data)
            
            # Resize image to fit nicely (max 120x120) while keeping aspect ratio
            img.thumbnail((120, 120), Image.Resampling.LANCZOS)
            self.logo_img = ImageTk.PhotoImage(img)
            tk.Label(logo_frame, image=self.logo_img, bg="#EBE7E0").pack()
        except ImportError:
            # Fallback to pure Tkinter if Pillow is not installed
            self.raw_logo_img = tk.PhotoImage(file=logo_path)
            # Use subsample to try to fit it roughly (optional, relies on raw image size)
            w, h = self.raw_logo_img.width(), self.raw_logo_img.height()
            scale = max(1, w // 120)
            self.logo_img = self.raw_logo_img.subsample(scale, scale)
            tk.Label(logo_frame, image=self.logo_img, bg="#EBE7E0").pack()
        except Exception as e:
            self.logger.error(f"Could not load logo.png: {e}")
            tk.Label(logo_frame, text="[Logo Image Error]", bg="#EBE7E0", fg="red").pack()

        logo_label = tk.Label(logo_frame, text="FOCUS PRISM", font=("Segoe UI", 12, "bold"), bg="#EBE7E0", fg="#5A5A5A")
        logo_label.pack()
        
        # --- TITLE SECTION ---
        title_label = tk.Label(
            main_container, 
            text="CLARITY", 
            font=("Times New Roman", 24, "bold"), 
            bg="#EBE7E0", 
            fg="#6B4F3B"
        )
        title_label.pack(pady=(10, 40))
        
        # --- INPUT FIELDS SECTION ---
        inputs_frame = tk.Frame(main_container, bg="#EBE7E0")
        inputs_frame.pack(fill="x", padx=20)
        
        # Custom Rounded Entry class
        class RoundedEntry(tk.Frame):
            def __init__(self, parent, placeholder, icon_text, right_icon=None, show=None, width=400):
                super().__init__(parent, bg="#EBE7E0")
                
                self.canvas = tk.Canvas(self, width=width, height=45, bg="#EBE7E0", highlightthickness=0)
                self.canvas.pack(fill="x", expand=True)
                
                # Draw rounded rectangle
                self.draw_rounded_rect(self.canvas, 2, 2, width-2, 43, 8, "#F7F5F0", "#A0998F")
                
                # Left Icon
                self.canvas.create_text(25, 22, text=icon_text, font=("Segoe UI", 14), fill="#5A5A5A")
                
                # Entry
                self.entry = tk.Entry(self, font=("Segoe UI", 11), bg="#F7F5F0", fg="#333333", bd=0, highlightthickness=0)
                if show:
                    self.entry.config(show=show)
                self.canvas.create_window(50, 22, window=self.entry, width=width-90, anchor="w")
                
                # Right Icon (Eye)
                if right_icon:
                    self.eye_text = self.canvas.create_text(width-25, 22, text=right_icon, font=("Segoe UI", 14), fill="#5A5A5A")
                    self.canvas.tag_bind(self.eye_text, "<Button-1>", self.toggle_show)
                    self.is_showing = False
                
                # Placeholder logic
                self.placeholder = placeholder
                self.show_char = show
                self.entry.insert(0, placeholder)
                self.entry.config(fg="#888888")
                if show:
                    self.entry.config(show="")
                
                self.entry.bind("<FocusIn>", self.on_focus_in)
                self.entry.bind("<FocusOut>", self.on_focus_out)
                
            def draw_rounded_rect(self, canvas, x1, y1, x2, y2, radius=25, fill="", outline=""):
                canvas.create_arc(x1, y1, x1+2*radius, y1+2*radius, start=90, extent=90, fill=fill, outline=outline)
                canvas.create_arc(x2-2*radius, y1, x2, y1+2*radius, start=0, extent=90, fill=fill, outline=outline)
                canvas.create_arc(x1, y2-2*radius, x1+2*radius, y2, start=180, extent=90, fill=fill, outline=outline)
                canvas.create_arc(x2-2*radius, y2-2*radius, x2, y2, start=270, extent=90, fill=fill, outline=outline)
                canvas.create_rectangle(x1+radius, y1, x2-radius, y2, fill=fill, outline="")
                canvas.create_rectangle(x1, y1+radius, x2, y2-radius, fill=fill, outline="")
                canvas.create_line(x1+radius, y1, x2-radius, y1, fill=outline)
                canvas.create_line(x1+radius, y2, x2-radius, y2, fill=outline)
                canvas.create_line(x1, y1+radius, x1, y2-radius, fill=outline)
                canvas.create_line(x2, y1+radius, x2, y2-radius, fill=outline)
                
            def on_focus_in(self, event):
                if self.entry.get() == self.placeholder:
                    self.entry.delete(0, tk.END)
                    self.entry.config(fg="#333333")
                    if self.show_char and not getattr(self, 'is_showing', False):
                        self.entry.config(show=self.show_char)
            
            def on_focus_out(self, event):
                if not self.entry.get():
                    self.entry.insert(0, self.placeholder)
                    self.entry.config(fg="#888888")
                    if self.show_char:
                        self.entry.config(show="")
                        
            def toggle_show(self, event):
                self.is_showing = not self.is_showing
                if self.is_showing:
                    self.entry.config(show="")
                    self.canvas.itemconfig(self.eye_text, text="🙈")
                else:
                    if self.entry.get() != self.placeholder:
                        self.entry.config(show="*")
                    self.canvas.itemconfig(self.eye_text, text="👁")
            
            def get(self):
                val = self.entry.get()
                return "" if val == self.placeholder else val
                
            def set(self, val):
                self.entry.delete(0, tk.END)
                self.entry.insert(0, val)
                self.entry.config(fg="#333333")
                if self.show_char and not getattr(self, 'is_showing', False):
                    self.entry.config(show=self.show_char)
        
        self.username_entry = RoundedEntry(inputs_frame, "Username", "👤")
        self.username_entry.pack(pady=(0, 15))
        
        self.password_entry = RoundedEntry(inputs_frame, "Password", "🔒", right_icon="👁", show="*")
        self.password_entry.pack(pady=(0, 15))
        
        # --- CHECKBOX SECTION ---
        check_frame = tk.Frame(inputs_frame, bg="#EBE7E0")
        check_frame.pack(fill="x", pady=(0, 25))
        
        self.save_creds_var = tk.BooleanVar(value=False)
        # Using a simple checkbutton that matches the background
        chk = tk.Checkbutton(
            check_frame, 
            text="Save login credentials", 
            variable=self.save_creds_var,
            bg="#EBE7E0", 
            activebackground="#EBE7E0", 
            fg="#5A5A5A", 
            selectcolor="#F7F5F0",
            font=("Segoe UI", 10),
            bd=0,
            highlightthickness=0
        )
        chk.pack(side="left")
        
        # --- LOGIN BUTTON ---
        btn_frame = tk.Frame(main_container, bg="#EBE7E0")
        btn_frame.pack(pady=(10, 0))
        
        btn_width, btn_height = 180, 40
        self.login_btn_canvas = tk.Canvas(btn_frame, width=btn_width, height=btn_height, bg="#EBE7E0", highlightthickness=0)
        self.login_btn_canvas.pack()
        
        def draw_rounded_btn(canvas, x1, y1, x2, y2, radius=10, fill="", outline=""):
            canvas.create_arc(x1, y1, x1+2*radius, y1+2*radius, start=90, extent=90, fill=fill, outline=outline)
            canvas.create_arc(x2-2*radius, y1, x2, y1+2*radius, start=0, extent=90, fill=fill, outline=outline)
            canvas.create_arc(x1, y2-2*radius, x1+2*radius, y2, start=180, extent=90, fill=fill, outline=outline)
            canvas.create_arc(x2-2*radius, y2-2*radius, x2, y2, start=270, extent=90, fill=fill, outline=outline)
            canvas.create_rectangle(x1+radius, y1, x2-radius, y2, fill=fill, outline="")
            canvas.create_rectangle(x1, y1+radius, x2, y2-radius, fill=fill, outline="")
            
        draw_rounded_btn(self.login_btn_canvas, 2, 2, btn_width-2, btn_height-2, radius=15, fill="#6B4F3B", outline="")
        self.login_btn_text = self.login_btn_canvas.create_text(btn_width/2, btn_height/2, text="Login", font=("Segoe UI", 11, "bold"), fill="white")
        
        def on_enter(e):
            self.login_btn_canvas.delete("all")
            draw_rounded_btn(self.login_btn_canvas, 2, 2, btn_width-2, btn_height-2, radius=15, fill="#8B6951", outline="")
            self.login_btn_canvas.create_text(btn_width/2, btn_height/2, text="Login", font=("Segoe UI", 11, "bold"), fill="white")
            self.login_btn_canvas.config(cursor="hand2")
            
        def on_leave(e):
            self.login_btn_canvas.delete("all")
            draw_rounded_btn(self.login_btn_canvas, 2, 2, btn_width-2, btn_height-2, radius=15, fill="#6B4F3B", outline="")
            self.login_btn_canvas.create_text(btn_width/2, btn_height/2, text="Login", font=("Segoe UI", 11, "bold"), fill="white")
            
        self.login_btn_canvas.bind("<Enter>", on_enter)
        self.login_btn_canvas.bind("<Leave>", on_leave)
        self.login_btn_canvas.bind("<Button-1>", lambda e: self.login())
        
        # Error Label
        self.error_label = tk.Label(
            main_container,
            text="",
            font=("Segoe UI", 10),
            bg="#EBE7E0",
            fg=Config.COLORS["danger"]
        )
        self.error_label.pack(pady=(10, 0))
        
        self.username_entry.entry.bind("<Return>", lambda e: self.password_entry.entry.focus())
        self.password_entry.entry.bind("<Return>", lambda e: self.login())
   
    def lighten_color(self, color):
        color = color.lstrip('#')
        r, g, b = tuple(int(color[i:i+2], 16) for i in (0, 2, 4))
        r = min(255, int(r * 1.2))
        g = min(255, int(g * 1.2))
        b = min(255, int(b * 1.2))
        return f"#{r:02x}{g:02x}{b:02x}"
   
    def demo_login(self, username):
        self.username_entry.set(username)
        demo_password = "sk123" if "sk" in username else "evpl123" if "evpl" in username else "maxi123"
        self.password_entry.set(demo_password)
        self.login()
   
    def login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        password_hash = hashlib.sha256(password.encode()).hexdigest()
       
        if not username or not password:
            self.error_label.config(text="Please enter username and password")
            return
       
        if username not in Config.USERS:
            self.error_label.config(text="Invalid username")
            return
       
        user_info = Config.USERS[username]
        if password_hash != user_info["password"]:
            self.error_label.config(text="Invalid password")
            return
       
        self.logger.info(f"User logged in: {username}")
        self.open_main_app(user_info["company"])
   
    def open_main_app(self, company):
        self.withdraw()
        if self.current_app:
            self.current_app.destroy()
        self.current_app = MainApp(company, self.db_handler, self.logger, self)
        self.current_app.protocol("WM_DELETE_WINDOW", self.on_main_app_close)
   
    def on_main_app_close(self):
        self.db_handler.close_all_connections()
        self.quit()


