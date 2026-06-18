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

# ==================== DATABASE HANDLER ====================

class DatabaseHandler:
    """Database connection and query management"""
   
    def __init__(self, logger: AppLogger):
        self.logger = logger
        self.connection_pool = {}
        self.query_cache = {}
        self.lock = threading.Lock()
   
    def test_connection(self, database: str = None) -> Tuple[bool, str]:
        try:
            if not database:
                database = list(Config.DATABASES.values())[0]
           
            conn = pymssql.connect(
                server=Config.SERVER,
                user=Config.UID,
                password=Config.PWD,
                database=database,
                timeout=10,
                login_timeout=10,
                autocommit=True
            )
            conn.close()
            return True, f"✅ Successfully connected to {database}"
        except Exception as e:
            return False, f"❌ Connection failed: {str(e)}"
   
    def get_connection(self, database: str):
        if database not in self.connection_pool:
            try:
                conn = pymssql.connect(
                    server=Config.SERVER,
                    user=Config.UID,
                    password=Config.PWD,
                    database=database,
                    timeout=30,
                    login_timeout=30,
                    autocommit=True
                )
                self.connection_pool[database] = conn
                self.logger.info(f"Connected to database: {database}")
            except Exception as e:
                self.logger.error(f"Database connection failed: {e}")
                raise
        return self.connection_pool[database]
   
    def execute_query(self, sql: str, database: str, use_cache: bool = False) -> pd.DataFrame:
        cache_key = f"{database}_{hash(sql)}"
       
        if use_cache and cache_key in self.query_cache:
            return self.query_cache[cache_key].copy()
       
        try:
            with self.lock:
                conn = self.get_connection(database)
                df = pd.read_sql_query(sql, conn).fillna("")
           
            if use_cache:
                self.query_cache[cache_key] = df.copy()
           
            return df
        except Exception as e:
            self.logger.error(f"Query execution failed: {e}")
            return pd.DataFrame()
   
    def close_all_connections(self):
        for database, conn in self.connection_pool.items():
            try:
                conn.close()
                self.logger.info(f"Closed connection to {database}")
            except:
                pass
        self.connection_pool.clear()
        self.query_cache.clear()


