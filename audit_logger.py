import tkinter as tk
from tkinter import ttk
import json
import os
from datetime import datetime
from config import Config

class AuditManager:
    """Centralized Audit Logging System for Enterprise Applications"""
    FILE_PATH = os.path.join(Config.BASE_DIR, "audit_logs.json")
    
    @classmethod
    def log_event(cls, action_type, username, details, status="Success"):
        """
        Logs an audit event securely.
        action_type: e.g. 'Login', 'Logout', 'Export', 'Email Sent'
        """
        history = cls._load_history()
        history.insert(0, {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "action_type": action_type,
            "username": username,
            "details": details,
            "status": status,
            "ip_address": "127.0.0.1" # Mocked for local apps, could be resolved
        })
        
        # Keep up to 1000 logs for enterprise retention locally
        if len(history) > 1000:
            history = history[:1000]
            
        cls._save_history(history)
        
    @classmethod
    def get_logs(cls):
        return cls._load_history()
        
    @classmethod
    def _load_history(cls):
        if os.path.exists(cls.FILE_PATH):
            try:
                with open(cls.FILE_PATH, 'r') as f:
                    return json.load(f)
            except Exception:
                return []
        return []
        
    @classmethod
    def _save_history(cls, history):
        try:
            with open(cls.FILE_PATH, 'w') as f:
                json.dump(history, f, indent=4)
        except Exception as e:
            print(f"Error saving audit log: {e}")

class AuditLogUI(tk.Toplevel):
    """Secure Enterprise UI for Admins to review Audit Logs."""
    
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Enterprise Audit Logs")
        self.geometry("850x500")
        self.configure(bg=Config.COLORS["light"])
        self.transient(parent)
        
        self.setup_ui()
        self.load_data()
        
    def setup_ui(self):
        header = tk.Frame(self, bg=Config.COLORS["primary"], height=60)
        header.pack(fill="x")
        tk.Label(header, text="🛡️ Enterprise Audit Logs", font=("Segoe UI", 16, "bold"), bg=Config.COLORS["primary"], fg="white").pack(pady=15)
        
        list_frame = tk.Frame(self, bg=Config.COLORS["white"], padx=10, pady=10)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        columns = ("Timestamp", "Action", "User", "Details", "Status", "IP Address")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        
        self.tree.heading("Timestamp", text="Date & Time")
        self.tree.heading("Action", text="Action Type")
        self.tree.heading("User", text="Username")
        self.tree.heading("Details", text="Details")
        self.tree.heading("Status", text="Status")
        self.tree.heading("IP Address", text="IP Address")
        
        self.tree.column("Timestamp", width=140, anchor="center")
        self.tree.column("Action", width=100, anchor="center")
        self.tree.column("User", width=100, anchor="center")
        self.tree.column("Details", width=250, anchor="w")
        self.tree.column("Status", width=80, anchor="center")
        self.tree.column("IP Address", width=100, anchor="center")
        
        # Zebra striping tag configurations
        self.tree.tag_configure("success", foreground=Config.COLORS.get("success", "#28a745"))
        self.tree.tag_configure("failed", foreground=Config.COLORS.get("danger", "#dc3545"))
        self.tree.tag_configure("evenrow", background=Config.COLORS.get("white", "#FFFFFF"))
        self.tree.tag_configure("oddrow", background=Config.COLORS.get("alternate_row", "#F5F5F5"))
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        btn_frame = tk.Frame(self, bg=Config.COLORS["light"])
        btn_frame.pack(fill="x", pady=10)
        tk.Button(btn_frame, text="Close Audit Trail", font=("Segoe UI", 10, "bold"), bg=Config.COLORS["gray"], fg="white", relief="flat", command=self.destroy).pack(pady=5)
        
    def load_data(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        logs = AuditManager.get_logs()
        for idx, record in enumerate(logs):
            status = record.get("status", "")
            tags = ("evenrow" if idx % 2 == 0 else "oddrow",)
            if status.lower() == "failed":
                tags += ("failed",)
            elif status.lower() == "success":
                tags += ("success",)
                
            self.tree.insert("", "end", values=(
                record.get("timestamp", ""),
                record.get("action_type", ""),
                record.get("username", ""),
                record.get("details", ""),
                status,
                record.get("ip_address", "")
            ), tags=tags)
