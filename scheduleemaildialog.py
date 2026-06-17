import tkinter as tk
from tkinter import ttk, messagebox, font, scrolledtext, filedialog
import os
from datetime import datetime, timedelta
import threading
import json
import time
import schedule
import warnings
from typing import Dict, Any, Optional, List, Tuple

warnings.filterwarnings("ignore", category=UserWarning)

from config import Config
from emailhistorydialog import EmailHistoryDialog
from emailscheduler import EmailScheduler
from schedule_triggers_dialog import ScheduleTriggersDialog
from ui_components import ModernButton

# ==================== SCHEDULE EMAIL DIALOG ====================

class ScheduleEmailDialog(tk.Toplevel):
    def __init__(self, parent, company: str, get_data_func, email_scheduler: EmailScheduler):
        super().__init__(parent)
        self.parent = parent
        self.company = company
        self.get_data_func = get_data_func
        self.email_scheduler = email_scheduler
        self.recipients_list = []
        
        # Placeholders and vars
        self.manual_name_var = tk.StringVar()
        self.manual_email_var = tk.StringVar()
        self.recipient_company_var = tk.StringVar(value=self.company)
        self.time_var = tk.StringVar(value="17:00")
       
        self.email_scheduler.set_data_fetch_function(get_data_func)
       
        self.setup_window()
        self.create_ui()
        self.load_schedule_status()
        self.update_countdown_timer()
   
    def setup_window(self):
        self.title(f"⏰ Schedule Email Reports - {self.company}")
        self.configure(bg=Config.COLORS["white"])
        self.geometry("1100x920")
       
        x = self.parent.winfo_x() + self.parent.winfo_width()//2 - 550
        y = self.parent.winfo_y() + self.parent.winfo_height()//2 - 460
        self.geometry(f"1100x920+{x}+{y}")
       
        self.resizable(False, False)
        self.transient(self.parent)
        self.grab_set()
   
    def create_ui(self):
        # --- Header ---
        header = tk.Frame(self, bg=Config.COLORS["primary"], height=90)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text="📧 Schedule Email Reports",
            font=("Segoe UI", 24, "bold"),
            bg=Config.COLORS["primary"],
            fg=Config.COLORS["white"]
        ).pack(pady=(15, 2))
        
        self.countdown_label = tk.Label(
            header,
            text="⌛ Preparing countdown...",
            font=("Segoe UI", 10, "italic"),
            bg=Config.COLORS["primary"],
            fg="#DED9D1"
        )
        self.countdown_label.pack(pady=(0, 10))

        # --- Main Layout (Two Columns) ---
        outer_container = tk.Frame(self, bg=Config.COLORS["white"])
        outer_container.pack(fill="both", expand=True, padx=25, pady=20)
        
        # Left Panel: Management (Fixed Width)
        left_panel = tk.Frame(outer_container, bg=Config.COLORS["white"], width=340)
        left_panel.pack(side="left", fill="both", padx=(0, 25))
        left_panel.pack_propagate(False)
        
        # Right Panel: Recipient list and Trigger Status
        right_panel = tk.Frame(outer_container, bg=Config.COLORS["white"])
        right_panel.pack(side="left", fill="both", expand=True)
        
        # --- LEFT PANEL CONTENT ---
        # 1. Sender Info Card
        sender_frame = tk.LabelFrame(left_panel, text="📤 SENDER INFORMATION", font=("Segoe UI", 9, "bold"),
                                   bg=Config.COLORS["white"], fg=Config.COLORS["secondary"], padx=15, pady=15)
        sender_frame.pack(fill="x", pady=(0, 20))
        
        tk.Label(sender_frame, text=f"Active Account:", font=("Segoe UI", 9), 
                 bg=Config.COLORS["white"], fg=Config.COLORS["gray"]).pack(anchor="w")
        tk.Label(sender_frame, text=self.email_scheduler.get_sender_email(), 
                 font=("Segoe UI", 10, "bold"), bg=Config.COLORS["white"], fg=Config.COLORS["dark"]).pack(anchor="w")
        
        # 2. Add Recipients Card
        add_frame = tk.LabelFrame(left_panel, text="➕ ADD RECIPIENTS", font=("Segoe UI", 9, "bold"),
                                bg=Config.COLORS["white"], fg=Config.COLORS["secondary"], padx=15, pady=15)
        add_frame.pack(fill="x", pady=(0, 20))
        
        tk.Label(add_frame, text="Target Company:", font=("Segoe UI", 9, "bold"), 
                 bg=Config.COLORS["white"], fg=Config.COLORS["primary"]).pack(anchor="w", pady=(0, 5))
        
        btn_row = tk.Frame(add_frame, bg=Config.COLORS["white"])
        btn_row.pack(fill="x", pady=(0, 15))
        for comp in ["SK", "EVPL", "Maximus"]:
            tk.Radiobutton(btn_row, text=comp, variable=self.recipient_company_var, value=comp,
                         font=("Segoe UI", 9), bg=Config.COLORS["white"], 
                         activebackground=Config.COLORS["white"]).pack(side="left", padx=5)
        
        tk.Label(add_frame, text="Name:", font=("Segoe UI", 9), bg=Config.COLORS["white"]).pack(anchor="w")
        self.manual_name_entry = tk.Entry(add_frame, textvariable=self.manual_name_var, font=("Segoe UI", 11), 
                             bg="#F8F9FA", relief="solid", bd=1)
        self.manual_name_entry.pack(fill="x", pady=(2, 12), ipady=6)
        self.set_placeholder(self.manual_name_entry, "Recipient Name")
        
        tk.Label(add_frame, text="Email Address:", font=("Segoe UI", 9), bg=Config.COLORS["white"]).pack(anchor="w")
        self.manual_email_entry = tk.Entry(add_frame, textvariable=self.manual_email_var, font=("Segoe UI", 11), 
                              bg="#F8F9FA", relief="solid", bd=1)
        self.manual_email_entry.pack(fill="x", pady=(2, 12), ipady=6)
        self.set_placeholder(self.manual_email_entry, "email@example.com")
        
        ModernButton(add_frame, text="Add to List", command=self.add_manual_recipient, 
                     variant="success").pack(fill="x", pady=(5, 5))
        
        # 3. Quick Schedule Card
        sched_frame = tk.LabelFrame(left_panel, text="⏰ QUICK SCHEDULE", font=("Segoe UI", 9, "bold"),
                                  bg=Config.COLORS["white"], fg=Config.COLORS["secondary"], padx=15, pady=15)
        sched_frame.pack(fill="x", pady=(0, 10))
        
        tk.Label(sched_frame, text="Daily Send Time:", font=("Segoe UI", 9), bg=Config.COLORS["white"]).pack(anchor="w")
        tk.Label(sched_frame, text="(HH:MM Format)", font=("Segoe UI", 8, "italic"), bg=Config.COLORS["white"], fg=Config.COLORS["gray"]).pack(anchor="w", pady=(0, 10))
        
        tk.Entry(sched_frame, textvariable=self.time_var, font=("Segoe UI", 22, "bold"), 
                 justify="center", bg="#F8F9FA", relief="solid", bd=1, width=6).pack(pady=5)
        
        ModernButton(sched_frame, text="Create Schedule", command=self.add_trigger, 
                     variant="schedule").pack(fill="x", pady=(10, 5))
        
        self.status_label = tk.Label(sched_frame, text="📌 Ready", font=("Segoe UI", 9, "italic"), 
                                     bg=Config.COLORS["white"], fg=Config.COLORS["secondary"], wraplength=300)
        self.status_label.pack(fill="x", pady=(10, 0))
        
        # --- RIGHT PANEL CONTENT ---
        # Important: Pack Bottom DASHBOARD first
        
        # BOTTOM: Active Trigger Dashboard
        trigger_section = tk.Frame(right_panel, bg=Config.COLORS["white"], height=350)
        trigger_section.pack(fill="x", side="bottom")
        trigger_section.pack_propagate(False)
        
        t_header_row = tk.Frame(trigger_section, bg=Config.COLORS["white"])
        t_header_row.pack(fill="x", pady=(5, 10))
        tk.Label(t_header_row, text="📅 ACTIVE SCHEDULE DASHBOARD", font=("Segoe UI", 11, "bold"), 
                 bg=Config.COLORS["white"], fg=Config.COLORS["primary"]).pack(side="left")
        
        ModernButton(t_header_row, text="🔄 REFRESH", command=self.load_schedule_status, variant="info", 
                     font=("Segoe UI", 8, "bold"), width=10).pack(side="right", padx=5)

        t_list_container = tk.Frame(trigger_section, bg=Config.COLORS["white"], relief="solid", bd=1)
        t_list_container.pack(fill="both", expand=True)
        
        t_canvas = tk.Canvas(t_list_container, bg=Config.COLORS["white"], highlightthickness=0)
        t_scrollbar = ttk.Scrollbar(t_list_container, orient="vertical", command=t_canvas.yview)
        self.trigger_container = tk.Frame(t_canvas, bg=Config.COLORS["white"])
        
        self.trigger_container.bind("<Configure>", lambda e: t_canvas.configure(scrollregion=t_canvas.bbox("all")))
        t_window = t_canvas.create_window((0, 0), window=self.trigger_container, anchor="nw")
        t_canvas.configure(yscrollcommand=t_scrollbar.set)
        
        t_canvas.pack(side="left", fill="both", expand=True)
        t_scrollbar.pack(side="right", fill="y")
        
        def on_t_canvas_configure(e):
            t_canvas.itemconfig(t_window, width=e.width)
        t_canvas.bind("<Configure>", on_t_canvas_configure)

        # TOP: Target Recipient List
        recipient_section = tk.Frame(right_panel, bg=Config.COLORS["white"])
        recipient_section.pack(fill="both", expand=True, pady=(0, 20))
        
        list_header = tk.Frame(recipient_section, bg=Config.COLORS["white"])
        list_header.pack(fill="x", pady=(0, 10))
        
        tk.Label(list_header, text="👥 TARGET RECIPIENTS", font=("Segoe UI", 11, "bold"), 
                 bg=Config.COLORS["white"], fg=Config.COLORS["primary"]).pack(side="left")
        
        self.recipient_count_label = tk.Label(list_header, text="0 recipients", font=("Segoe UI", 10), 
                                            bg=Config.COLORS["white"], fg=Config.COLORS["accent"])
        self.recipient_count_label.pack(side="right")
        
        # Recipient Toolbar
        toolbar = tk.Frame(recipient_section, bg=Config.COLORS["white"])
        toolbar.pack(fill="x", pady=(0, 10))
        
        ModernButton(toolbar, text="✅ SELECT ALL", command=self.select_all, variant="info", width=12, font=("Segoe UI", 9, "bold")).pack(side="left", padx=2)
        ModernButton(toolbar, text="🔲 DESELECT ALL", command=self.deselect_all, variant="secondary", width=12, font=("Segoe UI", 9, "bold")).pack(side="left", padx=2)
        ModernButton(toolbar, text="🗑 REMOVE", command=self.remove_selected_recipients, variant="danger", width=10, font=("Segoe UI", 9, "bold")).pack(side="left", padx=15)
        
        ModernButton(toolbar, text="📜 HISTORY", command=self.show_history, variant="history", width=10, font=("Segoe UI", 9, "bold")).pack(side="right", padx=2)
        ModernButton(toolbar, text="⏰ ADVANCED", command=self.show_triggers, variant="primary", width=12, font=("Segoe UI", 9, "bold")).pack(side="right", padx=2)
        
        # Recipient List Area
        list_container = tk.Frame(recipient_section, bg=Config.COLORS["white"], relief="solid", bd=1)
        list_container.pack(fill="both", expand=True)
        
        r_canvas = tk.Canvas(list_container, bg=Config.COLORS["white"], highlightthickness=0)
        r_scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=r_canvas.yview)
        self.recipient_container = tk.Frame(r_canvas, bg=Config.COLORS["white"])
        
        self.recipient_container.bind("<Configure>", lambda e: r_canvas.configure(scrollregion=r_canvas.bbox("all")))
        r_window_inst = r_canvas.create_window((0, 0), window=self.recipient_container, anchor="nw")
        r_canvas.configure(yscrollcommand=r_scrollbar.set)
        
        r_canvas.pack(side="left", fill="both", expand=True)
        r_scrollbar.pack(side="right", fill="y")
        
        def on_canvas_configure(e):
            r_canvas.itemconfig(r_window_inst, width=e.width)
        r_canvas.bind("<Configure>", on_canvas_configure)

        # Final Status Row
        bottom_row = tk.Frame(self, bg="#F0F2F5", height=50)
        bottom_row.pack(fill="x", side="bottom")
        
        self.message_label = tk.Label(bottom_row, text="", font=("Segoe UI", 10, "bold"), bg="#F0F2F5")
        self.message_label.pack(side="left", padx=25)
        
        ModernButton(bottom_row, text="Close Dashboard", command=self.destroy, variant="secondary", width=15).pack(side="right", padx=15, pady=8)

    def select_all(self):
        for child in self.recipient_container.winfo_children():
            if hasattr(child, 'var'): child.var.set(True)

    def deselect_all(self):
        for child in self.recipient_container.winfo_children():
            if hasattr(child, 'var'): child.var.set(False)

    def get_selected_recipients(self):
        selected = []
        for child in self.recipient_container.winfo_children():
            if hasattr(child, 'var') and child.var.get():
                if hasattr(child, 'data_tuple'):
                    selected.append(child.data_tuple)
        return selected
   
    def show_triggers(self):
        TriggersDialog(self, self.email_scheduler)
   
    def show_history(self):
        EmailHistoryDialog(self, self.email_scheduler)
   
    def update_trigger_summary(self):
        triggers = self.email_scheduler.get_active_triggers()
        if triggers:
            times = sorted(list(set([t.get('time', '') for t in triggers])))
            self.status_label.config(text=f"🟢 Next Run(s): {', '.join(times[:3])}", fg=Config.COLORS["success"])
        else:
            self.status_label.config(text="🔴 No active schedules", fg=Config.COLORS["gray"])
        self.update_trigger_display()
   
    def update_trigger_display(self):
        """Populate the dashboard list with premium design"""
        for widget in self.trigger_container.winfo_children():
            widget.destroy()
        
        triggers = self.email_scheduler.get_all_triggers()
        if not triggers:
            tk.Label(self.trigger_container, text="No active schedules found.", 
                     bg=Config.COLORS["white"], fg=Config.COLORS["gray"], font=("Segoe UI", 10, "italic")).pack(pady=40)
            return

        # Dashboard Column Headers
        h_frame = tk.Frame(self.trigger_container, bg="#E9ECEF")
        h_frame.pack(fill="x")
        
        # FIXED: Placed expand and fill in .pack() only
        tk.Label(h_frame, text=" TIME", font=("Segoe UI", 8, "bold"), bg="#E9ECEF", fg="#495057", width=10, anchor="w").pack(side="left", padx=10, pady=8)
        tk.Label(h_frame, text="COMPANY", font=("Segoe UI", 8, "bold"), bg="#E9ECEF", fg="#495057", width=12, anchor="w").pack(side="left")
        tk.Label(h_frame, text="STATUS", font=("Segoe UI", 8, "bold"), bg="#E9ECEF", fg="#495057", width=12, anchor="w").pack(side="left")
        
        # FIXED HERE: The Recipients header
        tk.Label(h_frame, text="RECIPIENTS SUMMARY", font=("Segoe UI", 8, "bold"), bg="#E9ECEF", fg="#495057", anchor="w").pack(side="left", expand=True, fill="x")
        
        tk.Label(h_frame, text="ACTIONS ", font=("Segoe UI", 8, "bold"), bg="#E9ECEF", fg="#495057", width=20).pack(side="right", padx=10)

        for idx, t in enumerate(triggers):
            enabled = t.get('enabled', False)
            bg = Config.COLORS["white"] if idx % 2 == 0 else Config.COLORS["alternate_row"]
            
            row = tk.Frame(self.trigger_container, bg=bg)
            row.pack(fill="x", pady=1)
            
            # 1. Time (Consolas)
            tk.Label(row, text=t.get('time'), font=("Consolas", 12, "bold"), bg=bg, width=10, anchor="w").pack(side="left", padx=10, pady=10)
            
            # 2. Company
            tk.Label(row, text=t.get('company'), font=("Segoe UI Black", 9), bg=bg, fg=Config.COLORS["primary"], width=12, anchor="w").pack(side="left")
            
            # 3. Status Badge
            status_color = Config.COLORS["success"] if enabled else Config.COLORS["danger"]
            status_text = "● ACTIVE" if enabled else "● INACTIVE"
            tk.Label(row, text=status_text, font=("Segoe UI", 8, "bold"), fg=status_color, bg=bg, width=12, anchor="w").pack(side="left")
            
            # 4. Recipients Summary (Truncated)
            rcount = len(t.get('recipients', []))
            rnames = ", ".join(t.get('recipient_names', []))
            if len(rnames) > 40: rnames = rnames[:37] + "..."
            tk.Label(row, text=f"[{rcount}] {rnames}", font=("Segoe UI", 9), fg=Config.COLORS["gray"], bg=bg, anchor="w").pack(side="left", expand=True, fill="x", padx=10)
            
            # 5. Row Actions
            btn_frame = tk.Frame(row, bg=bg)
            btn_frame.pack(side="right", padx=8)
            
            toggle_text = "DISABLE" if enabled else "ENABLE"
            toggle_variant = "toggle_off" if enabled else "toggle_on"
            
            # FIXED: font handled by rewritten ModernButton constructor
            ModernButton(btn_frame, text=toggle_text, variant=toggle_variant, width=9, font=("Segoe UI", 7, "bold"),
                         command=lambda tid=t.get('id'): self.toggle_trigger_from_ui(tid)).pack(side="left", padx=2)
            
            ModernButton(btn_frame, text="🗑", variant="danger", width=3, font=("Segoe UI", 7, "bold"),
                         command=lambda tid=t.get('id'): self.remove_trigger_from_ui(tid)).pack(side="left", padx=2)

    def toggle_trigger_from_ui(self, trigger_id):
        success, trigger = self.email_scheduler.toggle_trigger(trigger_id)
        if success:
            self.load_schedule_status()
            self.show_message(f"✅ Schedule {trigger['time']} {'activated' if trigger['enabled'] else 'deactivated'}", "success")

    def remove_trigger_from_ui(self, trigger_id):
        if messagebox.askyesno("Confirm Delete", "Delete this schedule forever?"):
            success, message = self.email_scheduler.remove_trigger(trigger_id)
            if success:
                self.load_schedule_status()
                self.show_message("🗑 Schedule removed", "info")

    def set_placeholder(self, entry, placeholder):
        entry.insert(0, placeholder)
        entry.config(fg=Config.COLORS["gray"])
        def on_focus_in(e):
            if entry.get() == placeholder:
                entry.delete(0, tk.END)
                entry.config(fg=Config.COLORS["dark"])
        def on_focus_out(e):
            if not entry.get():
                entry.insert(0, placeholder)
                entry.config(fg=Config.COLORS["gray"])
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)
   
    def show_message(self, message, msg_type="info"):
        colors = {"info": Config.COLORS["secondary"], "success": Config.COLORS["success"],
                  "warning": Config.COLORS["warning"], "danger": Config.COLORS["danger"]}
        if hasattr(self, 'message_label'):
            self.message_label.config(text=message, fg=colors.get(msg_type, Config.COLORS["secondary"]))
            if msg_type != "danger":
                self.after(5000, lambda: self.message_label.config(text="") if hasattr(self, 'message_label') else None)

    def add_manual_recipient(self):
        name = self.manual_name_var.get().strip()
        email = self.manual_email_var.get().strip()
        company = self.recipient_company_var.get()
       
        if not name or name == "Recipient Name" or name == "":
            self.show_message("⚠️ Please enter a name", "warning")
            return
        if not email or email == "email@example.com" or '@' not in email:
            self.show_message("⚠️ Please enter a valid email", "warning")
            return
        if any(r[1] == email and r[2] == company for r in self.recipients_list):
            self.show_message(f"⚠️ {email} already exists for {company}", "warning")
            return
       
        self.recipients_list.append((name, email, company))
        self.update_recipient_display()
        self.manual_name_var.set("")
        self.manual_email_var.set("")
        self.show_message(f"✅ Added {name} to {company}", "success")
   
    def update_recipient_display(self):
        if not hasattr(self, 'recipient_container'): return
        for widget in self.recipient_container.winfo_children():
            widget.destroy()
        for idx, data in enumerate(self.recipients_list):
            name, email, company = data
            item_frame = tk.Frame(self.recipient_container, relief="flat", bd=0,
                                bg=Config.COLORS["alternate_row"] if idx % 2 == 0 else Config.COLORS["white"])
            item_frame.pack(fill="x", pady=0)
            item_frame.data_tuple = data
            item_frame.email = email
            var = tk.BooleanVar(value=True)
            item_frame.var = var
            tk.Checkbutton(item_frame, variable=var, bg=item_frame['bg'], activebackground=item_frame['bg']).pack(side="left", padx=10, pady=8)
            tk.Label(item_frame, text=f"[{company}]", width=10, font=("Segoe UI", 9, "bold"),
                     bg=item_frame['bg'], fg=Config.COLORS["secondary"], anchor="w").pack(side="left")
            tk.Label(item_frame, text=name, font=("Segoe UI", 10, "bold"),
                     bg=item_frame['bg'], fg=Config.COLORS["dark"], width=20, anchor="w").pack(side="left", padx=5)
            tk.Label(item_frame, text=f"<{email}>", font=("Segoe UI", 9),
                     bg=item_frame['bg'], fg=Config.COLORS["gray"]).pack(side="left", padx=5)
        self.recipient_count_label.config(text=f"{len(self.recipients_list)} total")

    def remove_selected_recipients(self):
        selected_emails = [child.email for child in self.recipient_container.winfo_children() 
                          if hasattr(child, 'var') and child.var.get()]
        if not selected_emails:
            self.show_message("⚠️ No recipients selected", "warning")
            return
        if messagebox.askyesno("Confirm Remove", f"Remove {len(selected_emails)} selected recipients?"):
            self.recipients_list = [r for r in self.recipients_list if r[1] not in selected_emails]
            self.update_recipient_display()
            self.show_message(f"🗑 Removed {len(selected_emails)} recipients", "info")

    def add_trigger(self):
        selected_data = self.get_selected_recipients()
        if not selected_data:
            self.show_message("⚠️ Please select recipients from the list first", "danger")
            return
        time_str = self.time_var.get().strip()
        try:
            datetime.strptime(time_str, "%H:%M")
        except ValueError:
            self.show_message("⚠️ Invalid time format. Use HH:MM", "danger")
            return
        by_company = {}
        for name, email, company in selected_data:
            if company not in by_company: by_company[company] = {"recipients": [], "names": []}
            by_company[company]["recipients"].append(email)
            by_company[company]["names"].append(name)
        selected_date = datetime.now().strftime("%d/%m/%Y")
        if len(by_company) > 1:
            confirm = messagebox.askyesno("Confirm", f"Schedule for {len(by_company)} companies at {time_str}?")
            if not confirm: return
        else:
            company = list(by_company.keys())[0]
            confirm = messagebox.askyesno("Confirm", f"Schedule {company} for {time_str}?")
            if not confirm: return
        
        # Add triggers for each company
        results = []
        for company, data in by_company.items():
            success, message, _ = self.email_scheduler.add_trigger(
                time_str=time_str, recipients=data["recipients"],
                recipient_names=data["names"], company=company, selected_date=selected_date
            )
            results.append((company, success, message))
            
        if any(ok for _, ok, _ in results):
            self.show_message("✅ Schedules created successfully", "success")
            self.load_schedule_status()
            detail = "\n".join([f"• {c}: {m}" for c, ok, m in results])
            messagebox.showinfo("✅ Scheduled", f"Process complete!\n\n{detail}")
        else:
            self.show_message("❌ Scheduling failed", "danger")

    def load_schedule_status(self):
        self.update_trigger_summary()

    def update_countdown_timer(self):
        try:
            if not self.winfo_exists(): return
            info = self.email_scheduler.get_next_trigger_info()
            if info:
                rem = info['remaining']
                hours, remainder = divmod(rem.seconds, 3600)
                minutes, seconds = divmod(remainder, 60)
                time_str = f"{hours:02d}h {minutes:02d}m {seconds:02d}s"
                company = info['trigger'].get('company', 'Unknown')
                trigger_time = info['trigger'].get('time', '??:??')
                self.countdown_label.config(text=f"🕒 Next Report: {company} at {trigger_time} (In {time_str})", fg="#FFFFFF")
            else:
                self.countdown_label.config(text="⌛ No upcoming reports scheduled", fg="#DED9D1")
            self.after(1000, self.update_countdown_timer)
        except: pass
