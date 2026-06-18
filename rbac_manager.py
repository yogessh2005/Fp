import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import hashlib
from config import Config

class RBACManager:
    """Role-Based Access Control Manager"""
    
    FILE_PATH = os.path.join(Config.BASE_DIR, "rbac_database.json")
    
    # Permission Table defining what each role can access
    PERMISSIONS = {
        "admin": {
            "allowed_tabs": ["Dashboard", "Store Data", "Overall Estimate", "Sessionwise Estimate", "Actual Sales", "Score"],
            "allowed_buttons": ["refresh_btn", "export_btn", "schedule_btn", "upload_grn_btn", "rbac_btn"]
        },
        "manager": {
            "allowed_tabs": ["Dashboard", "Store Data", "Overall Estimate", "Sessionwise Estimate", "Actual Sales", "Score"],
            "allowed_buttons": ["refresh_btn", "export_btn", "schedule_btn"]
        },
        "employee": {
            "allowed_tabs": ["Dashboard", "Store Data"],
            "allowed_buttons": ["refresh_btn"]
        },
        "viewer": {
            "allowed_tabs": ["Dashboard"],
            "allowed_buttons": ["refresh_btn"]
        }
    }
    
    def __init__(self):
        self.custom_users = {}
        self.load_custom_users()
        
    def load_custom_users(self):
        if os.path.exists(self.FILE_PATH):
            try:
                with open(self.FILE_PATH, 'r') as f:
                    self.custom_users = json.load(f)
            except Exception as e:
                print(f"Failed to load custom users: {e}")
                self.custom_users = {}
                
    def save_custom_users(self):
        try:
            with open(self.FILE_PATH, 'w') as f:
                json.dump(self.custom_users, f, indent=4)
        except Exception as e:
            print(f"Failed to save custom users: {e}")
            
    def get_user(self, username):
        """Retrieve user info from custom users first, then fallback to Config.USERS"""
        if username in self.custom_users:
            return self.custom_users[username]
        elif username in Config.USERS:
            return Config.USERS[username]
        return None
        
    def authenticate(self, username, password):
        """Check credentials against merged user list."""
        user = self.get_user(username)
        if not user:
            return False, "Invalid username"
            
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        if password_hash != user["password"]:
            return False, "Invalid password"
            
        return True, user
        
    def add_or_update_user(self, username, password, company, role, full_name):
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        self.custom_users[username] = {
            "password": password_hash,
            "company": company,
            "role": role.lower(),
            "full_name": full_name
        }
        self.save_custom_users()

    def enforce_rbac(self, app_instance, username):
        """Dynamically modify the MainApp UI based on the user's role"""
        user_info = self.get_user(username)
        if not user_info:
            return
            
        role = user_info.get("role", "viewer").lower()
        perms = self.PERMISSIONS.get(role, self.PERMISSIONS["viewer"])
        allowed_tabs = perms["allowed_tabs"]
        allowed_buttons = perms["allowed_buttons"]
        
        # 1. Enforce Tab / Module Restrictions
        # The notebook contains frames. We must hide the ones not allowed.
        # But ttk.Notebook only allows hide().
        if hasattr(app_instance, 'notebook'):
            notebook = app_instance.notebook
            tabs = notebook.tabs()
            
            for tab_id in tabs:
                tab_name = notebook.tab(tab_id, "text")
                if tab_name not in allowed_tabs:
                    notebook.hide(tab_id)
                    
            # Hide sidebar buttons for restricted tabs
            if hasattr(app_instance, 'tab_buttons'):
                for btn in app_instance.tab_buttons:
                    # Strip icons and spaces to match name
                    clean_name = btn.cget("text").split("   ")[-1].strip()
                    if clean_name not in allowed_tabs and "Dashboard" not in clean_name:
                        # btn is a SidebarButton which inherits from tk.Button or tk.Frame
                        btn.pack_forget()
                    elif "Dashboard" in btn.cget("text") and "Dashboard" not in allowed_tabs:
                        btn.pack_forget()
                        
        # 2. Enforce Button Restrictions in Header
        # MainApp has refresh_btn, export_btn, schedule_btn, upload_grn_btn
        if not hasattr(app_instance, 'refresh_btn'):
            return # App might not be fully initialized or structure changed
            
        buttons_map = {
            "refresh_btn": app_instance.refresh_btn,
            "export_btn": app_instance.export_btn,
            "schedule_btn": app_instance.schedule_btn,
            "upload_grn_btn": getattr(app_instance, "upload_grn_btn", None)
        }
        
        for btn_key, btn_obj in buttons_map.items():
            if btn_obj:
                if btn_key not in allowed_buttons:
                    # Disable or hide
                    btn_obj.grid_remove() # Hides it completely from the header action_frame
                    
        # 3. Add RBAC Admin Button if allowed
        if "rbac_btn" in allowed_buttons:
            # We add it next to the upload_grn_btn or at the end
            from ui_components import ModernButton
            
            # Find the action_frame parent
            action_frame = app_instance.refresh_btn.master
            
            app_instance.rbac_btn = ModernButton(
                action_frame,
                text="👥 Manage Roles",
                command=lambda: self.open_admin_ui(app_instance, app_instance.company),
                variant="primary"
            )
            app_instance.rbac_btn.grid(row=0, column=5, padx=5, pady=3)
            
    def open_admin_ui(self, parent, current_company):
        RBACAdminUI(parent, self, current_company)


class RBACAdminUI(tk.Toplevel):
    """UI for managing users and roles (Admin only)"""
    def __init__(self, parent, rbac_manager, current_company):
        super().__init__(parent)
        self.rbac_manager = rbac_manager
        self.current_company = current_company
        
        self.title("Role Assignment & User Management")
        self.geometry("600x450")
        self.configure(bg=Config.COLORS["light"])
        self.transient(parent)
        self.grab_set()
        
        self.setup_ui()
        self.load_users()
        
    def setup_ui(self):
        # Header
        header = tk.Frame(self, bg=Config.COLORS["primary"], height=50)
        header.pack(fill="x")
        tk.Label(header, text="Role Assignment", font=("Segoe UI", 14, "bold"), bg=Config.COLORS["primary"], fg="white").pack(pady=10)
        
        # Form Frame
        form_frame = tk.Frame(self, bg=Config.COLORS["white"], padx=20, pady=20)
        form_frame.pack(fill="x", padx=10, pady=10)
        
        tk.Label(form_frame, text="Username:", bg=Config.COLORS["white"]).grid(row=0, column=0, sticky="w", pady=5)
        self.username_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.username_var, width=25).grid(row=0, column=1, sticky="w", pady=5)
        
        tk.Label(form_frame, text="Full Name:", bg=Config.COLORS["white"]).grid(row=1, column=0, sticky="w", pady=5)
        self.fullname_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.fullname_var, width=25).grid(row=1, column=1, sticky="w", pady=5)
        
        tk.Label(form_frame, text="Password:", bg=Config.COLORS["white"]).grid(row=2, column=0, sticky="w", pady=5)
        self.password_var = tk.StringVar()
        tk.Entry(form_frame, textvariable=self.password_var, show="*", width=25).grid(row=2, column=1, sticky="w", pady=5)
        
        tk.Label(form_frame, text="Role:", bg=Config.COLORS["white"]).grid(row=3, column=0, sticky="w", pady=5)
        self.role_var = tk.StringVar(value="employee")
        ttk.Combobox(form_frame, textvariable=self.role_var, values=["admin", "manager", "employee", "viewer"], state="readonly", width=22).grid(row=3, column=1, sticky="w", pady=5)
        
        tk.Label(form_frame, text="Company:", bg=Config.COLORS["white"]).grid(row=4, column=0, sticky="w", pady=5)
        self.company_var = tk.StringVar(value=self.current_company)
        ttk.Combobox(form_frame, textvariable=self.company_var, values=list(Config.DATABASES.keys()), state="readonly", width=22).grid(row=4, column=1, sticky="w", pady=5)
        
        btn_frame = tk.Frame(form_frame, bg=Config.COLORS["white"])
        btn_frame.grid(row=5, column=0, columnspan=2, pady=15)
        
        tk.Button(btn_frame, text="Save User", bg=Config.COLORS["success"], fg="white", relief="flat", command=self.save_user).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Clear", bg=Config.COLORS["gray"], fg="white", relief="flat", command=self.clear_form).pack(side="left", padx=5)
        
        # User List
        list_frame = tk.Frame(self, bg=Config.COLORS["white"], padx=10, pady=10)
        list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        columns = ("Username", "Full Name", "Role", "Company", "Type")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
            
        self.tree.pack(fill="both", expand=True, side="left")
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.bind("<<TreeviewSelect>>", self.on_user_select)

    def save_user(self):
        username = self.username_var.get().strip()
        fullname = self.fullname_var.get().strip()
        password = self.password_var.get().strip()
        role = self.role_var.get().strip()
        company = self.company_var.get().strip()
        
        if not username or not fullname or not password or not role or not company:
            messagebox.showerror("Error", "All fields are required.", parent=self)
            return
            
        self.rbac_manager.add_or_update_user(username, password, company, role, fullname)
        messagebox.showinfo("Success", f"User '{username}' saved successfully.", parent=self)
        self.load_users()
        self.clear_form()
        
    def clear_form(self):
        self.username_var.set("")
        self.fullname_var.set("")
        self.password_var.set("")
        self.role_var.set("employee")
        
    def on_user_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        values = item["values"]
        
        self.username_var.set(values[0])
        self.fullname_var.set(values[1])
        self.role_var.set(values[2])
        self.company_var.set(values[3])
        # Password cannot be retrieved (it's hashed), leave blank or dummy
        self.password_var.set("")
        
    def load_users(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Load Config.USERS
        for username, info in Config.USERS.items():
            # Don't show users that are overridden in custom_users yet
            if username not in self.rbac_manager.custom_users:
                self.tree.insert("", "end", values=(username, info.get("full_name", ""), info.get("role", ""), info.get("company", ""), "Built-in"))
                
        # Load custom_users
        for username, info in self.rbac_manager.custom_users.items():
            self.tree.insert("", "end", values=(username, info.get("full_name", ""), info.get("role", ""), info.get("company", ""), "Custom"))
