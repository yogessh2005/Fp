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
from emailhistory import EmailHistory
from excelformatter import ExcelFormatter

# ==================== EMAIL SCHEDULER ====================

class EmailScheduler:
    """Handle scheduled email sending - Supports multiple triggers"""
   
    def __init__(self, logger: AppLogger):
        self.logger = logger
        self.scheduled_jobs = []
        self.is_running = False
        self.scheduler_thread = None
        self.config = Config.EMAIL_CONFIG.copy()
        self.active_triggers = []
       
        self.history = EmailHistory(logger)
        self.ensure_temp_dir()
        self.load_schedule_config()
        self.start_scheduler()
       
        atexit.register(self.cleanup)
   
    def ensure_temp_dir(self):
        try:
            if not os.path.exists(Config.TEMP_DIR):
                os.makedirs(Config.TEMP_DIR)
                self.logger.info(f"Created temp directory: {Config.TEMP_DIR}")
        except Exception as e:
            self.logger.warning(f"Could not create temp directory: {e}")
   
    def cleanup(self):
        self.is_running = False
        if self.scheduler_thread and self.scheduler_thread.is_alive():
            self.scheduler_thread.join(timeout=2)
        self.logger.info("Email scheduler cleaned up")
   
    def get_sender_email(self):
        return self.config.get('sender_email', '')
   
    def load_schedule_config(self):
        try:
            if os.path.exists(Config.SCHEDULE_FILE):
                with open(Config.SCHEDULE_FILE, 'r') as f:
                    self.schedule_config = json.load(f)
            else:
                self.schedule_config = {
                    "triggers": [],
                    "last_sent": None
                }
        except Exception as e:
            self.logger.error(f"Failed to load schedule config: {e}")
            self.schedule_config = {
                "triggers": [],
                "last_sent": None
            }
        self.restore_triggers()
   
    def save_schedule_config(self):
        try:
            with open(Config.SCHEDULE_FILE, 'w') as f:
                json.dump(self.schedule_config, f, indent=4)
            return True
        except Exception as e:
            self.logger.error(f"Failed to save schedule config: {e}")
            return False
   
    def restore_triggers(self):
        schedule.clear()
        self.active_triggers = []
       
        for trigger in self.schedule_config.get('triggers', []):
            if trigger.get('enabled', False):
                self._schedule_trigger(trigger)
   
    def _schedule_trigger(self, trigger: Dict):
        time_str = trigger.get('time', '')
        recipients = trigger.get('recipients', [])
        recipient_names = trigger.get('recipient_names', [])
        company = trigger.get('company', '')
        selected_date = trigger.get('selected_date', datetime.now().strftime("%d/%m/%Y"))
        trigger_id = trigger.get('id', str(int(time.time())))
       
        def scheduled_task():
            try:
                current_trigger = self.get_trigger_by_id(trigger_id)
                if not current_trigger or not current_trigger.get('enabled', False):
                    self.logger.info(f"Trigger {time_str} is disabled, skipping...")
                    return
               
                self.logger.info(f"Scheduled task triggered at {time_str}")
               
                if hasattr(self, '_get_data_func'):
                    all_data = self._get_data_func()
                    if all_data and any(not df.empty for df in all_data.values()):
                        success, message = self.send_email_report(
                            recipients=recipients,
                            company=company,
                            all_data=all_data,
                            selected_date=selected_date,
                            trigger_time=time_str
                        )
                        if success:
                            self.logger.info(f"Email sent successfully at {time_str}")
                        else:
                            self.logger.error(f"Email failed at {time_str}: {message}")
                    else:
                        self.logger.warning(f"No data available at {time_str}")
                else:
                    self.logger.warning("Data fetch function not set")
                   
            except Exception as e:
                self.logger.error(f"Scheduled task failed for {time_str}: {e}")
       
        try:
            schedule.every().day.at(time_str).do(scheduled_task)
            trigger['_job'] = scheduled_task
            self.active_triggers.append(trigger)
            self.logger.info(f"Trigger scheduled at {time_str}")
        except Exception as e:
            self.logger.error(f"Failed to schedule trigger at {time_str}: {e}")
   
    def add_trigger(self, time_str: str, recipients: List[str], recipient_names: List[str],
                    company: str, selected_date: str) -> Tuple[bool, str, Dict]:
        try:
            datetime.strptime(time_str, "%H:%M")
        except ValueError:
            return False, "Invalid time format. Use HH:MM (24-hour)", None
       
        for trigger in self.schedule_config.get('triggers', []):
            if trigger.get('time') == time_str and trigger.get('company') == company:
                if trigger.get('enabled', False):
                    return False, f"A trigger already exists at {time_str} for {company}", None
       
        trigger_id = str(int(time.time()))
        trigger = {
            "id": trigger_id,
            "time": time_str,
            "recipients": recipients,
            "recipient_names": recipient_names,
            "company": company,
            "selected_date": selected_date,
            "enabled": True,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "last_sent": None
        }
       
        self.schedule_config.setdefault('triggers', []).append(trigger)
        self.save_schedule_config()
        self._schedule_trigger(trigger)
       
        self.logger.info(f"Added new trigger at {time_str} for {company}")
        return True, f"Trigger added at {time_str}", trigger
   
    def remove_trigger(self, trigger_id: str) -> bool:
        triggers = self.schedule_config.get('triggers', [])
        for i, trigger in enumerate(triggers):
            if trigger.get('id') == trigger_id:
                self.active_triggers = [t for t in self.active_triggers if t.get('id') != trigger_id]
                triggers.pop(i)
                self.save_schedule_config()
                self.logger.info(f"Removed trigger {trigger_id}")
                return True
        return False
   
    def toggle_trigger(self, trigger_id: str) -> Tuple[bool, Dict]:
        triggers = self.schedule_config.get('triggers', [])
        for trigger in triggers:
            if trigger.get('id') == trigger_id:
                current_state = trigger.get('enabled', False)
                trigger['enabled'] = not current_state
                self.save_schedule_config()
               
                if trigger['enabled']:
                    self._schedule_trigger(trigger)
                else:
                    self.active_triggers = [t for t in self.active_triggers if t.get('id') != trigger_id]
               
                self.logger.info(f"Trigger {trigger_id} toggled to {trigger['enabled']}")
                return True, trigger
        return False, None
   
    def get_trigger_by_id(self, trigger_id: str) -> Optional[Dict]:
        triggers = self.schedule_config.get('triggers', [])
        for trigger in triggers:
            if trigger.get('id') == trigger_id:
                return trigger
        return None
   
    def get_all_triggers(self) -> List[Dict]:
        return self.schedule_config.get('triggers', [])
   
    def get_active_triggers(self) -> List[Dict]:
        return [t for t in self.schedule_config.get('triggers', []) if t.get('enabled', False)]
   
    def send_email_report(self, recipients: List[str], company: str,
                          all_data: Dict, selected_date: str,
                          trigger_time: str = "") -> Tuple[bool, str]:
        try:
            sender_email = self.config.get('sender_email', '')
            sender_password = self.config.get('sender_password', '')
           
            if not sender_email:
                return False, "Sender email not configured."
            if not sender_password:
                return False, "Email password not configured."
           
            filepath = self.generate_excel_report(company, all_data, selected_date)
           
            if not filepath or not os.path.exists(filepath):
                return False, "Failed to generate report file"
           
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = ', '.join(recipients)
            subject = f"{company} Report - {selected_date}"
            if trigger_time:
                subject += f" ({trigger_time})"
            msg['Subject'] = subject
           
            body = f"""Dear Team,

Please find attached the {company} report for {selected_date}.

The report includes:
• Store Data - GRN and stock analysis
• Overall Estimate - Performance metrics
• Sessionwise Estimate - Session-wise breakdown
• Actual Sales - Sales and revenue analysis
• Score - Performance scoring (Updated PSERM logic: >35 = 25 marks)
• Scoring Rules - Detailed scoring criteria

This is an automated report sent as per schedule.

Best regards,
{Config.APP_NAME} System
"""
            msg.attach(MIMEText(body, 'plain'))
           
            with open(filepath, 'rb') as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename={os.path.basename(filepath)}'
                )
                msg.attach(part)
           
            try:
                with smtplib.SMTP(self.config['smtp_server'], self.config['smtp_port']) as server:
                    if self.config.get('use_tls', True):
                        server.starttls()
                    server.login(sender_email, sender_password)
                    server.send_message(msg)
            except smtplib.SMTPAuthenticationError as auth_error:
                error_msg = f"Authentication failed: {str(auth_error)}"
                self.logger.error(error_msg)
                try:
                    if os.path.exists(filepath):
                        os.remove(filepath)
                except:
                    pass
                self.history.add_record(company, recipients, selected_date, "❌ FAILED", error_msg, trigger_time)
                return False, error_msg
           
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
            except:
                pass
           
            triggers = self.schedule_config.get('triggers', [])
            for trigger in triggers:
                if trigger.get('time') == trigger_time and trigger.get('company') == company:
                    trigger['last_sent'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    break
            self.save_schedule_config()
           
            self.history.add_record(company, recipients, selected_date, "✅ SUCCESS",
                                   f"Sent to {len(recipients)} recipients", trigger_time)
           
            self.logger.info(f"Email sent to {len(recipients)} recipients at {trigger_time}")
            return True, f"Email sent successfully to {len(recipients)} recipients"
           
        except Exception as e:
            error_msg = f"Failed to send: {str(e)}"
            self.logger.error(error_msg)
            self.history.add_record(company, recipients, selected_date, "❌ FAILED", error_msg, trigger_time)
            return False, error_msg
   
    def generate_excel_report(self, company: str, all_data: Dict, selected_date: str) -> str:
        if not all_data:
            raise ValueError("No data available to export")
       
        clean_date = selected_date.replace('/', '_').replace('\\', '_').replace(':', '_')
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"temp_scheduled_{company}_{clean_date}_{timestamp}.xlsx"
       
        self.ensure_temp_dir()
        filepath = os.path.join(Config.TEMP_DIR, filename)
       
        try:
            with pd.ExcelWriter(filepath, engine="openpyxl", mode='w') as writer:
                for module_name, df in all_data.items():
                    if not df.empty:
                        df.to_excel(writer, sheet_name=module_name, index=False)
                       
                        worksheet = writer.sheets[module_name]
                        ExcelFormatter.apply_clean_formatting(worksheet, df)
               
                # Scoring Rules - Updated with new PSERM logic
                rules_df = pd.DataFrame(Config.SCORING_RULES)
                rules_df.to_excel(writer, sheet_name="Scoring_Rules", index=False, header=False)
               
                worksheet = writer.sheets["Scoring_Rules"]
                worksheet.column_dimensions['A'].width = 35
                worksheet.column_dimensions['B'].width = 25
                worksheet.column_dimensions['C'].width = 15
           
            return filepath
           
        except Exception as e:
            self.logger.error(f"Failed to generate Excel report: {e}")
            raise
   
    def start_scheduler(self):
        if not self.scheduler_thread or not self.scheduler_thread.is_alive():
            self.is_running = True
            self.scheduler_thread = threading.Thread(target=self.run_scheduler, daemon=True)
            self.scheduler_thread.start()
            self.logger.info("Scheduler thread started")
   
    def run_scheduler(self):
        self.logger.info("Email scheduler running in background")
        while self.is_running:
            try:
                schedule.run_pending()
                time.sleep(60)
            except Exception as e:
                self.logger.error(f"Scheduler error: {e}")
                time.sleep(60)
   
    def set_data_fetch_function(self, func):
        self._get_data_func = func


