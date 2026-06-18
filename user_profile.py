import tkinter as tk
from tkinter import ttk, messagebox
import hashlib
from config import Config

class UserProfileUI(tk.Toplevel):
    """UI for users to manage their own profile (Change Name, Password)."""
    
    def __init__(self, parent, username, rbac_manager):
        super().__init__(parent)
        self.username = username
        self.rbac_manager = rbac_manager
        
        self.title("My Profile")
        self.geometry("400x350")
        self.configure(bg=Config.COLORS["light"])
        self.transient(parent)
        self.grab_set()
        
        # Load user info
        self.user_info = self.rbac_manager.get_user(self.username)
        if not self.user_info:
            messagebox.showerror("Error", "User info not found.")
            self.destroy()
            return
            
        self.setup_ui()
        
    def setup_ui(self):
        header = tk.Frame(self, bg=Config.COLORS["primary"], height=50)
        header.pack(fill="x")
        tk.Label(header, text="👤 My Profile", font=("Segoe UI", 14, "bold"), bg=Config.COLORS["primary"], fg="white").pack(pady=10)
        
        form = tk.Frame(self, bg=Config.COLORS["white"], padx=20, pady=20)
        form.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Username (Read-only)
        tk.Label(form, text="Username:", font=("Segoe UI", 10), bg=Config.COLORS["white"], fg=Config.COLORS["dark"]).grid(row=0, column=0, sticky="w", pady=10)
        tk.Label(form, text=self.username, font=("Segoe UI", 10, "bold"), bg=Config.COLORS["white"], fg=Config.COLORS["primary"]).grid(row=0, column=1, sticky="w", pady=10)
        
        # Role (Read-only)
        tk.Label(form, text="Role:", font=("Segoe UI", 10), bg=Config.COLORS["white"], fg=Config.COLORS["dark"]).grid(row=1, column=0, sticky="w", pady=10)
        role_label = self.user_info.get("role", "viewer").title()
        tk.Label(form, text=role_label, font=("Segoe UI", 10, "bold"), bg=Config.COLORS["white"], fg=Config.COLORS["success"]).grid(row=1, column=1, sticky="w", pady=10)
        
        # Full Name (Editable)
        tk.Label(form, text="Full Name:", font=("Segoe UI", 10), bg=Config.COLORS["white"], fg=Config.COLORS["dark"]).grid(row=2, column=0, sticky="w", pady=10)
        self.fullname_var = tk.StringVar(value=self.user_info.get("full_name", ""))
        tk.Entry(form, textvariable=self.fullname_var, font=("Segoe UI", 10), width=25).grid(row=2, column=1, sticky="w", pady=10)
        
        # New Password (Editable)
        tk.Label(form, text="New Password:", font=("Segoe UI", 10), bg=Config.COLORS["white"], fg=Config.COLORS["dark"]).grid(row=3, column=0, sticky="w", pady=10)
        self.password_var = tk.StringVar()
        tk.Entry(form, textvariable=self.password_var, show="*", font=("Segoe UI", 10), width=25).grid(row=3, column=1, sticky="w", pady=10)
        tk.Label(form, text="(Leave blank to keep current)", font=("Segoe UI", 8), bg=Config.COLORS["white"], fg=Config.COLORS["gray"]).grid(row=4, column=1, sticky="w")
        
        # Buttons
        btn_frame = tk.Frame(form, bg=Config.COLORS["white"])
        btn_frame.grid(row=5, column=0, columnspan=2, pady=20)
        
        tk.Button(btn_frame, text="Save Changes", font=("Segoe UI", 10, "bold"), bg=Config.COLORS["primary"], fg="white", relief="flat", command=self.save_profile).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Cancel", font=("Segoe UI", 10), bg=Config.COLORS["gray"], fg="white", relief="flat", command=self.destroy).pack(side="left", padx=10)
        
    def save_profile(self):
        new_name = self.fullname_var.get().strip()
        new_pass = self.password_var.get().strip()
        
        if not new_name:
            messagebox.showerror("Error", "Full Name cannot be empty.", parent=self)
            return
            
        # Update logic
        self.user_info["full_name"] = new_name
        if new_pass:
            self.user_info["password"] = hashlib.sha256(new_pass.encode()).hexdigest()
            
        # We must save this back to rbac_manager
        # Ensure user is in custom_users so changes persist
        self.rbac_manager.custom_users[self.username] = self.user_info
        self.rbac_manager.save_custom_users()
        
        messagebox.showinfo("Success", "Profile updated successfully!", parent=self)
        self.destroy()
