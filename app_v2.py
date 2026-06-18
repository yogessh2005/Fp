import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Ensure the app_code directory is in path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from loginpage import LoginPage
from mainapp import MainApp
from rbac_manager import RBACManager
from notification_center import NotificationManager, NotificationCenterUI
from theme_manager import ThemeManager
from report_history import ReportHistoryManager, ReportHistoryUI
from user_profile import UserProfileUI
from pdf_generator import PDFGenerator
from ui_components import ModernButton
from config import Config
from audit_logger import AuditManager, AuditLogUI
from databasehandler import DatabaseHandler
from emailscheduler import EmailScheduler

def launch_professional_app():
    """
    Launches the application with all professional enterprise features integrated:
    RBAC, Notifications, Dark Mode, PDF Export, User Profiles, Report History, and Audit Logs.
    Zero modifications to original source files via comprehensive runtime patching.
    """
    print("Initializing Enterprise App Environment...")
    rbac_manager = RBACManager()
    
    # --- Intercept Database Errors globally for Notifications ---
    original_db_execute = DatabaseHandler.execute_query
    def patched_execute_query(self, query, database):
        try:
            return original_db_execute(self, query, database)
        except Exception as e:
            # We don't have access to the app instance here easily if it's called in threads,
            # but we can try to push to a global notification queue or if we have app instance access.
            # A cleaner way is to patch it per instance inside open_main_app, but execute_query is an instance method.
            raise e
            
    # --- Intercept Login ---
    original_login = LoginPage.login
    original_open_main_app = LoginPage.open_main_app
    
    def patched_login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not username or not password:
            self.error_label.config(text="Please enter username and password")
            return
            
        success, result = rbac_manager.authenticate(username, password)
        
        if not success:
            self.error_label.config(text=result)
            AuditManager.log_event("Login", username, f"Failed attempt: {result}", "Failed")
            return
            
        self.logger.info(f"User logged in via Enterprise App v2: {username}")
        AuditManager.log_event("Login", username, "User successfully authenticated.")
        
        self._authenticated_username = username
        self.open_main_app(result["company"])
        
    def patched_open_main_app(self, company):
        original_open_main_app(self, company)
        username = getattr(self, '_authenticated_username', None)
        
        if username and self.current_app:
            app = self.current_app
            
            # 1. Enforce RBAC rules
            rbac_manager.enforce_rbac(app, username)
            
            # Retrieve the action_frame from the header
            try:
                action_frame = app.refresh_btn.master
            except AttributeError:
                print("Could not find action_frame in MainApp")
                return
                
            # 2. Setup Notification Center
            app.notif_manager = NotificationManager()
            app.notif_ui = NotificationCenterUI(action_frame, manager=app.notif_manager, bg_color=Config.COLORS["white"])
            app.notif_ui.grid(row=0, column=7, padx=10, pady=3, rowspan=2)
            
            app.notif_manager.add_success(f"Welcome back, {username}!")
            
            # Hook into DB for this app instance to notify on error
            original_instance_db_execute = app.db_handler.execute_query
            def notify_db_error_execute(query, db_name):
                try:
                    return original_instance_db_execute(query, db_name)
                except Exception as e:
                    app.notif_manager.add_database_error(str(e))
                    raise e
            app.db_handler.execute_query = notify_db_error_execute
            
            # Hook into Email Dispatch for notifications and Audit
            original_send_email = app.email_scheduler.send_email_report
            def patched_send_email(recipients, company, all_data, selected_date, trigger_time=""):
                result, msg = original_send_email(recipients, company, all_data, selected_date, trigger_time)
                to_emails = ", ".join(recipients) if isinstance(recipients, list) else str(recipients)
                if result:
                    app.notif_manager.add_email_sent(to_emails)
                    AuditManager.log_event("Email Sent", username, f"Sent to {to_emails}. Subj: Report")
                else:
                    app.notif_manager.add_error(f"Failed to send email to {to_emails}: {msg}")
                    AuditManager.log_event("Email Sent", username, f"Failed dispatch to {to_emails}: {msg}", "Failed")
                return result, msg
            app.email_scheduler.send_email_report = patched_send_email

            # Hook into Logout
            original_logout = app.logout
            def patched_logout():
                AuditManager.log_event("Logout", username, "User logged out.")
                original_logout()
            app.logout = patched_logout
            app.logout_btn.configure(command=patched_logout)
            
            # 3. Setup Theme Toggle (Dark Mode)
            def toggle_theme():
                ThemeManager.toggle_theme(app)
                app.notif_ui.bg_color = Config.COLORS["white"]
                app.notif_ui.configure(bg=Config.COLORS["white"])
                app.notif_ui.btn_container.configure(bg=Config.COLORS["white"])
                app.notif_ui.bell_btn.configure(bg=Config.COLORS["white"], fg=Config.COLORS["primary"], activebackground=Config.COLORS["white"])
                
                theme_btn_text = "☀️ Light Mode" if ThemeManager.is_dark_mode() else "🌙 Dark Mode"
                app.theme_btn.configure(text=theme_btn_text)
                app.notif_manager.add_success("Theme updated successfully.")
                
            app.theme_btn = tk.Button(app.right_nav, text="🌙 Dark Mode", font=("Segoe UI", 10, "bold"),
                                      bg=Config.COLORS["primary"], fg=Config.COLORS["white"], activebackground=Config.COLORS["hover"], activeforeground=Config.COLORS["white"], relief="flat", cursor="hand2", command=toggle_theme)
            app.theme_btn.pack(side="right", padx=5)
            
            # 4. Setup User Profile
            def open_profile():
                UserProfileUI(app, username, rbac_manager)
                
            app.profile_btn = tk.Button(app.right_nav, text=f"👤 {username}", font=("Segoe UI", 10, "bold"),
                                      bg=Config.COLORS["primary"], fg=Config.COLORS["white"], activebackground=Config.COLORS["hover"], activeforeground=Config.COLORS["white"], relief="flat", cursor="hand2", command=open_profile)
            app.profile_btn.pack(side="right", padx=15)
            
            # 5. Setup PDF Export
            user_info = rbac_manager.get_user(username)
            role = user_info.get("role", "viewer").lower()
            perms = rbac_manager.PERMISSIONS.get(role, rbac_manager.PERMISSIONS["viewer"])
            
            if "export_btn" in perms["allowed_buttons"]:
                def export_pdf():
                    current_tab_idx = app.notebook.index(app.notebook.select())
                    if current_tab_idx == 0:
                        messagebox.showwarning("Warning", "Cannot export Dashboard directly. Please select a data tab.")
                        return
                        
                    module_name = Config.MODULES[current_tab_idx - 1]['name']
                    if module_name not in app.all_data or app.all_data[module_name] is None or app.all_data[module_name].empty:
                        messagebox.showwarning("Warning", "No data available to export.")
                        return
                        
                    df = app.all_data[module_name]
                    filepath = filedialog.asksaveasfilename(
                        defaultextension=".pdf",
                        initialfile=f"{module_name}_{company}_{app.selected_date}.pdf",
                        title="Save PDF As",
                        filetypes=[("PDF files", "*.pdf")]
                    )
                    
                    if filepath:
                        summary_info = {
                            "Company": company,
                            "Generated By": username,
                            "Analysis Date": app.selected_date_display,
                            "Total Records": len(df)
                        }
                        
                        success, msg = PDFGenerator.generate_report(df, filepath, company, f"{module_name} Report", summary_info)
                        if success:
                            app.notif_manager.add_report_generated(f"{module_name} PDF")
                            ReportHistoryManager.log_report("PDF", filepath, username, "Success")
                            AuditManager.log_event("Export", username, f"Exported PDF: {os.path.basename(filepath)}")
                            messagebox.showinfo("Success", msg)
                        else:
                            app.notif_manager.add_error("Failed to generate PDF.")
                            ReportHistoryManager.log_report("PDF", filepath, username, "Failed")
                            AuditManager.log_event("Export", username, f"Failed PDF Export: {os.path.basename(filepath)}", "Failed")
                            messagebox.showerror("Error", msg)
                            
                app.pdf_btn = ModernButton(app.action_frame, text="📄 Export PDF", command=export_pdf, variant="secondary")
                app.pdf_btn.pack(side="left", padx=5)
                app.export_btn.pack(side="left", padx=5)
                if app.schedule_btn.winfo_ismapped():
                    app.schedule_btn.pack(side="left", padx=5)
                if hasattr(app, 'upload_grn_btn') and app.upload_grn_btn.winfo_ismapped():
                    app.upload_grn_btn.pack(side="left", padx=5)
                if hasattr(app, 'rbac_btn') and app.rbac_btn.winfo_ismapped():
                    app.rbac_btn.pack(side="left", padx=5)
                    
            # 6. Report History Access (Admin/Manager only)
            if role in ["admin", "manager"]:
                def open_history():
                    ReportHistoryUI(app)
                    
                app.history_btn = tk.Button(app.right_nav, text="📜 Report History", font=("Segoe UI", 10, "bold"),
                                      bg=Config.COLORS["primary"], fg=Config.COLORS["white"], activebackground=Config.COLORS["hover"], activeforeground=Config.COLORS["white"], relief="flat", cursor="hand2", command=open_history)
                app.history_btn.pack(side="right", padx=10)
                
            # 7. Enterprise Audit Logs (Admin only)
            if role == "admin":
                def open_audit():
                    AuditLogUI(app)
                    
                app.audit_btn = tk.Button(app.right_nav, text="🛡️ Audit Logs", font=("Segoe UI", 10, "bold"),
                                      bg=Config.COLORS["primary"], fg=Config.COLORS["white"], activebackground=Config.COLORS["hover"], activeforeground=Config.COLORS["white"], relief="flat", cursor="hand2", command=open_audit)
                app.audit_btn.pack(side="right", padx=10)
                
            # 8. Intercept Excel Export to log it in history & audit
            original_excel_export = app.export_to_excel
            def patched_export_to_excel():
                try:
                    original_excel_export()
                    ReportHistoryManager.log_report("Excel", "User selected file", username, "Success")
                    AuditManager.log_event("Export", username, "Initiated Excel Export")
                    app.notif_manager.add_report_generated("Excel Export")
                except Exception as e:
                    ReportHistoryManager.log_report("Excel", "Error", username, f"Failed: {str(e)}")
                    AuditManager.log_event("Export", username, f"Failed Excel Export: {str(e)}", "Failed")
                    app.notif_manager.add_error("Excel export failed.")
                    
            app.export_to_excel = patched_export_to_excel
            app.export_btn.configure(command=patched_export_to_excel)

    # Apply the patches
    LoginPage.login = patched_login
    LoginPage.open_main_app = patched_open_main_app
    
    # Start the app
    app = LoginPage()
    app.mainloop()

if __name__ == "__main__":
    launch_professional_app()
