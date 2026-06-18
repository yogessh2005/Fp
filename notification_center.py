import tkinter as tk
from tkinter import ttk, font
import json
import os
from datetime import datetime
from config import Config

class NotificationManager:
    """Manager for storing, retrieving, and dispatching notifications."""
    
    FILE_PATH = os.path.join(Config.BASE_DIR, "notification_history.json")
    
    def __init__(self):
        self.notifications = []
        self.callbacks = []
        self.load_history()
        
    def load_history(self):
        if os.path.exists(self.FILE_PATH):
            try:
                with open(self.FILE_PATH, 'r') as f:
                    self.notifications = json.load(f)
            except Exception as e:
                print(f"Failed to load notifications: {e}")
                self.notifications = []
                
    def save_history(self):
        try:
            with open(self.FILE_PATH, 'w') as f:
                json.dump(self.notifications, f, indent=4)
        except Exception as e:
            print(f"Failed to save notifications: {e}")
            
    def register_callback(self, callback):
        if callback not in self.callbacks:
            self.callbacks.append(callback)
            
    def unregister_callback(self, callback):
        if callback in self.callbacks:
            self.callbacks.remove(callback)
            
    def _notify_callbacks(self):
        for callback in self.callbacks:
            callback()
            
    def get_unread_count(self):
        return sum(1 for n in self.notifications if not n.get('is_read', False))
        
    def mark_all_as_read(self):
        for n in self.notifications:
            n['is_read'] = True
        self.save_history()
        self._notify_callbacks()
        
    def clear_history(self):
        self.notifications = []
        self.save_history()
        self._notify_callbacks()
        
    def _add_notification(self, type_str, message, icon):
        notification = {
            "id": datetime.now().strftime("%Y%m%d%H%M%S%f"),
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "type": type_str,
            "message": message,
            "icon": icon,
            "is_read": False
        }
        self.notifications.insert(0, notification)
        
        # Keep only the last 100 notifications to prevent file bloat
        if len(self.notifications) > 100:
            self.notifications = self.notifications[:100]
            
        self.save_history()
        self._notify_callbacks()
        
    # --- Specific Notification Types ---
    
    def add_success(self, message):
        self._add_notification("success", message, "✅")
        
    def add_warning(self, message):
        self._add_notification("warning", message, "⚠️")
        
    def add_error(self, message):
        self._add_notification("error", message, "❌")
        
    def add_email_sent(self, recipient_info=""):
        msg = f"Email sent successfully. {recipient_info}".strip()
        self._add_notification("email", msg, "📧")
        
    def add_report_generated(self, report_name=""):
        msg = f"Report generated successfully: {report_name}".strip()
        self._add_notification("report", msg, "📊")
        
    def add_database_error(self, error_details=""):
        msg = f"Database Error: {error_details}".strip()
        self._add_notification("database", msg, "🗄️❌")


class NotificationCenterUI(tk.Frame):
    """UI Component for the Notification Bell and Popup History."""
    
    def __init__(self, parent, manager: NotificationManager, bg_color=Config.COLORS["white"]):
        super().__init__(parent, bg=bg_color)
        self.manager = manager
        self.bg_color = bg_color
        self.popup = None
        
        self.setup_ui()
        self.manager.register_callback(self.update_badge)
        self.update_badge()
        
    def setup_ui(self):
        # Container for the bell and badge
        self.btn_container = tk.Frame(self, bg=self.bg_color)
        self.btn_container.pack(padx=5, pady=5)
        
        # The Bell Button
        self.bell_btn = tk.Button(
            self.btn_container,
            text="🔔",
            font=("Segoe UI", 16),
            bg=self.bg_color,
            fg=Config.COLORS["primary"],
            relief="flat",
            activebackground=self.bg_color,
            bd=0,
            cursor="hand2",
            command=self.toggle_popup
        )
        self.bell_btn.pack(side="left")
        
        # The Badge (Red dot with count)
        self.badge_lbl = tk.Label(
            self.btn_container,
            text="0",
            font=("Segoe UI", 8, "bold"),
            bg="red",
            fg="white",
            relief="flat"
        )
        # We will dynamically place it using .place() relative to the bell button
        
    def update_badge(self):
        count = self.manager.get_unread_count()
        if count > 0:
            display_count = "9+" if count > 9 else str(count)
            self.badge_lbl.config(text=display_count)
            # Position badge at the top right of the bell button
            self.badge_lbl.place(x=22, y=0, width=16, height=16)
        else:
            self.badge_lbl.place_forget()
            
        # Also update popup if it's open
        if self.popup and self.popup.winfo_exists():
            self._populate_history()
            
    def toggle_popup(self):
        if self.popup and self.popup.winfo_exists():
            self.popup.destroy()
            self.popup = None
            # When closing the popup, mark all as read
            self.manager.mark_all_as_read()
        else:
            self.show_popup()
            
    def show_popup(self):
        self.popup = tk.Toplevel(self)
        self.popup.overrideredirect(True)  # Remove window decorations
        self.popup.configure(bg=Config.COLORS["border"])
        
        # Calculate position (below the bell button)
        x = self.bell_btn.winfo_rootx() - 300 + self.bell_btn.winfo_width()
        y = self.bell_btn.winfo_rooty() + self.bell_btn.winfo_height() + 5
        self.popup.geometry(f"320x400+{x}+{y}")
        
        # Bind clicking outside to close
        self.popup.bind("<FocusOut>", self._on_focus_out)
        self.popup.focus_set()
        
        # Main container with a border
        container = tk.Frame(self.popup, bg=Config.COLORS["white"], highlightbackground=Config.COLORS["border"], highlightthickness=1)
        container.pack(fill="both", expand=True)
        
        # Header
        header = tk.Frame(container, bg=Config.COLORS["primary"], height=40)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        tk.Label(header, text="Notifications", font=("Segoe UI", 11, "bold"), bg=Config.COLORS["primary"], fg="white").pack(side="left", padx=10)
        
        clear_btn = tk.Button(header, text="Clear All", font=("Segoe UI", 9), bg=Config.COLORS["primary"], fg="white", 
                              relief="flat", cursor="hand2", activebackground=Config.COLORS["accent"], command=self.clear_history)
        clear_btn.pack(side="right", padx=10)
        
        # Scrollable area for notifications
        self.canvas = tk.Canvas(container, bg=Config.COLORS["white"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        
        self.scrollable_frame = tk.Frame(self.canvas, bg=Config.COLORS["white"])
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", width=300)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Enable mousewheel scrolling
        self.popup.bind_all("<MouseWheel>", self._on_mousewheel)
        
        self._populate_history()
        
    def _populate_history(self):
        # Clear existing
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
            
        notifications = self.manager.notifications
        if not notifications:
            tk.Label(self.scrollable_frame, text="No notifications yet.", 
                     font=("Segoe UI", 10), bg=Config.COLORS["white"], fg=Config.COLORS["gray"]).pack(pady=30)
            return
            
        for n in notifications:
            bg_color = Config.COLORS["white"] if n.get("is_read") else Config.COLORS["alternate_row"]
            
            item_frame = tk.Frame(self.scrollable_frame, bg=bg_color)
            item_frame.pack(fill="x", pady=1, padx=2)
            
            # Type specific color coding
            border_color = Config.COLORS["border"]
            if n["type"] == "error" or n["type"] == "database":
                border_color = Config.COLORS["danger"]
            elif n["type"] == "success" or n["type"] == "email" or n["type"] == "report":
                border_color = Config.COLORS["success"]
            elif n["type"] == "warning":
                border_color = Config.COLORS["warning"]
                
            inner_frame = tk.Frame(item_frame, bg=bg_color, highlightbackground=border_color, highlightthickness=1)
            inner_frame.pack(fill="x", padx=5, pady=2)
            
            icon_lbl = tk.Label(inner_frame, text=n["icon"], font=("Segoe UI", 14), bg=bg_color)
            icon_lbl.pack(side="left", padx=(10, 5), pady=10)
            
            text_frame = tk.Frame(inner_frame, bg=bg_color)
            text_frame.pack(side="left", fill="both", expand=True, pady=5)
            
            msg_lbl = tk.Label(text_frame, text=n["message"], font=("Segoe UI", 10), bg=bg_color, fg=Config.COLORS["dark"], 
                               wraplength=230, justify="left", anchor="w")
            msg_lbl.pack(fill="x")
            
            time_lbl = tk.Label(text_frame, text=n["timestamp"], font=("Segoe UI", 8), bg=bg_color, fg=Config.COLORS["gray"], anchor="w")
            time_lbl.pack(fill="x")

    def clear_history(self):
        self.manager.clear_history()
        self._populate_history()
        
    def _on_focus_out(self, event):
        # We delay destruction slightly to allow clicks inside the popup to process
        if self.popup:
            try:
                # Check if the currently focused widget is a child of the popup
                focused_widget = self.winfo_toplevel().focus_displayof()
                if focused_widget and self.popup in self._get_all_parents(focused_widget):
                    return
            except Exception:
                pass
                
            # Need to schedule this to ensure it doesn't break click events on clear button
            self.after(100, self._check_and_close_popup)
            
    def _check_and_close_popup(self):
        if self.popup and self.popup.winfo_exists():
            # Double check focus
            focused = self.focus_displayof()
            if not focused or self.popup not in self._get_all_parents(focused):
                self.popup.destroy()
                self.popup = None
                self.manager.mark_all_as_read()
            
    def _get_all_parents(self, widget):
        parents = [widget]
        while widget.master:
            widget = widget.master
            parents.append(widget)
        return parents
        
    def _on_mousewheel(self, event):
        if self.popup and self.popup.winfo_exists():
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
    def destroy(self):
        self.manager.unregister_callback(self.update_badge)
        super().destroy()
