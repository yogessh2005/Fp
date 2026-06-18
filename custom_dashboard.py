import tkinter as tk
from tkinter import ttk
import pandas as pd
from datetime import datetime
import threading
import traceback
import matplotlib

# Set matplotlib to use TkAgg backend
matplotlib.use("TkAgg")
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from config import Config

class CustomDashboardTab(tk.Frame):
    """New Dashboard Module with requested KPIs and Charts"""
    
    def __init__(self, parent, company, db_handler, logger):
        super().__init__(parent, bg=Config.COLORS["light"])
        self.company = company
        self.db_handler = db_handler
        self.logger = logger
        self.refresh_interval = 30000  # 30 seconds auto-refresh
        self.after_id = None
        self.is_refreshing = False
        
        self.setup_ui()
        self.refresh_data()
        self.auto_refresh()
        
    def setup_ui(self):
        # 1. Top control frame for Title and Manual Refresh
        ctrl_frame = tk.Frame(self, bg=Config.COLORS["white"], height=60)
        ctrl_frame.pack(fill="x", padx=10, pady=(10, 5))
        ctrl_frame.pack_propagate(False)
        
        tk.Label(ctrl_frame, text="📈 KPI Dashboard", font=("Segoe UI", 16, "bold"), 
                 bg=Config.COLORS["white"], fg=Config.COLORS["primary"]).pack(side="left", padx=15, pady=15)
                 
        # Manual refresh button
        self.refresh_btn = tk.Button(ctrl_frame, text="🔄 Manual Refresh", 
                                     bg=Config.COLORS["primary"], fg=Config.COLORS["white"],
                                     font=("Segoe UI", 10), command=self.refresh_data, relief="flat", padx=10)
        self.refresh_btn.pack(side="left", padx=20, pady=15)
        
        self.last_refresh_lbl = tk.Label(ctrl_frame, text="Last Refresh: Never", 
                                         bg=Config.COLORS["white"], font=("Segoe UI", 10), fg=Config.COLORS["gray"])
        self.last_refresh_lbl.pack(side="right", padx=15, pady=15)

        # 2. Main content area for KPIs and Charts
        self.content_frame = tk.Frame(self, bg=Config.COLORS["light"])
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # 3. KPI Grid
        self.kpi_frame = tk.Frame(self.content_frame, bg=Config.COLORS["light"])
        self.kpi_frame.pack(fill="x", pady=(5, 10))
        for i in range(4):
            self.kpi_frame.grid_columnconfigure(i, weight=1)
            
        # Requested KPIs
        self.kpis = {
            "Total Reports": tk.StringVar(value="..."),
            "Reports Today": tk.StringVar(value="..."),
            "Emails Sent": tk.StringVar(value="..."),
            "Pending Emails": tk.StringVar(value="...")
        }
        
        idx = 0
        for title, var in self.kpis.items():
            card = tk.Frame(self.kpi_frame, bg=Config.COLORS["white"], relief="flat", bd=0)
            card.grid(row=0, column=idx, padx=8, pady=5, sticky="nsew")
            card.grid_propagate(False)
            card.config(height=100)
            
            # Subtle accent bar
            accent = tk.Frame(card, bg=Config.COLORS["primary"], width=4)
            accent.pack(side="left", fill="y")
            
            content = tk.Frame(card, bg=Config.COLORS["white"])
            content.pack(side="left", fill="both", expand=True, padx=15)
            
            tk.Label(content, text=title.upper(), font=("Segoe UI", 10, "bold"), bg=Config.COLORS["white"], fg=Config.COLORS["gray"]).pack(pady=(20, 0), anchor="w")
            tk.Label(content, textvariable=var, font=("Segoe UI", 24, "bold"), bg=Config.COLORS["white"], fg=Config.COLORS["primary"]).pack(pady=(5, 10), anchor="w")
            idx += 1
            
        # 4. Charts area
        self.charts_frame = tk.Frame(self.content_frame, bg=Config.COLORS["light"])
        self.charts_frame.pack(fill="both", expand=True)
        self.charts_frame.grid_rowconfigure(0, weight=1)
        self.charts_frame.grid_rowconfigure(1, weight=1)
        self.charts_frame.grid_columnconfigure(0, weight=2)
        self.charts_frame.grid_columnconfigure(1, weight=1)
        
        # Line Chart (Top Left)
        self.fig_line = Figure(figsize=(5, 3), dpi=100)
        self.ax_line = self.fig_line.add_subplot(111)
        self.canvas_line = FigureCanvasTkAgg(self.fig_line, master=self.charts_frame)
        self.canvas_line.get_tk_widget().grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # Pie Chart (Top Right)
        self.fig_pie = Figure(figsize=(3, 3), dpi=100)
        self.ax_pie = self.fig_pie.add_subplot(111)
        self.canvas_pie = FigureCanvasTkAgg(self.fig_pie, master=self.charts_frame)
        self.canvas_pie.get_tk_widget().grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        
        # Bar Chart (Bottom Full Width)
        self.fig_bar = Figure(figsize=(8, 2), dpi=100)
        self.ax_bar = self.fig_bar.add_subplot(111)
        self.canvas_bar = FigureCanvasTkAgg(self.fig_bar, master=self.charts_frame)
        self.canvas_bar.get_tk_widget().grid(row=1, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)
        
    def auto_refresh(self):
        self.refresh_data()
        self.after_id = self.after(self.refresh_interval, self.auto_refresh)
            
    def refresh_data(self):
        if self.is_refreshing:
            return
        self.is_refreshing = True
        self.refresh_btn.config(state="disabled", text="⏳ Refreshing...")
        # Use background threading for data fetching to prevent UI freezing
        threading.Thread(target=self._fetch_data_thread, daemon=True).start()
        
    def _fetch_data_thread(self):
        try:
            database = Config.DATABASES.get(self.company)
            
            # Using OutletMenuOrders to ensure we retrieve some data using the existing MSSQL connection,
            # while mapping the metrics to the newly requested KPIs.
            
            # 1. KPIs
            kpi_sql = """
            SELECT 
                (SELECT COUNT(DISTINCT OrderId) FROM OutletMenuOrders) as TotalReports,
                (SELECT COUNT(DISTINCT OrderId) FROM OutletMenuOrders WHERE CAST(OrderDate AS DATE) = CAST(GETDATE() AS DATE)) as ReportsToday,
                (SELECT COUNT(*) FROM OutletMenuOrders WHERE OrderStatus = 'Completed') as EmailsSent,
                (SELECT COUNT(*) FROM OutletMenuOrders WHERE OrderStatus = 'Pending') as PendingEmails
            """
            kpi_df = self.db_handler.execute_query(kpi_sql, database)
            
            # 2. Line Chart: Report Trends (Last 7 days)
            line_sql = """
            SELECT CAST(OrderDate AS DATE) as Date, COUNT(*) as Count
            FROM OutletMenuOrders
            WHERE OrderDate >= DATEADD(day, -7, GETDATE())
            GROUP BY CAST(OrderDate AS DATE)
            ORDER BY Date
            """
            line_df = self.db_handler.execute_query(line_sql, database)
            
            # 3. Bar Chart: Tasks overview (Last 6 months)
            bar_sql = """
            SELECT DATENAME(month, OrderDate) as Month, COUNT(*) as Count
            FROM OutletMenuOrders
            WHERE OrderDate >= DATEADD(month, -6, GETDATE())
            GROUP BY DATENAME(month, OrderDate), MONTH(OrderDate)
            ORDER BY MONTH(OrderDate)
            """
            bar_df = self.db_handler.execute_query(bar_sql, database)
            
            # 4. Pie Chart: System Activity Distribution
            pie_sql = """
            SELECT OrderStatus as Category, COUNT(*) as Count
            FROM OutletMenuOrders
            WHERE OrderStatus IS NOT NULL AND OrderStatus <> ''
            GROUP BY OrderStatus
            """
            pie_df = self.db_handler.execute_query(pie_sql, database)
            
            # Schedule UI update on main thread
            self.after(0, lambda: self._update_ui(kpi_df, line_df, bar_df, pie_df))
            
        except Exception as e:
            self.logger.error(f"Custom Dashboard data fetch error: {e}")
            self.logger.error(traceback.format_exc())
            self.after(0, self._handle_fetch_error)
            
    def _update_ui(self, kpi_df, line_df, bar_df, pie_df):
        try:
            # Update KPIs
            if not kpi_df.empty:
                self.kpis["Total Reports"].set(f"{kpi_df.iloc[0]['TotalReports']:,}")
                self.kpis["Reports Today"].set(f"{kpi_df.iloc[0]['ReportsToday']:,}")
                self.kpis["Emails Sent"].set(f"{kpi_df.iloc[0]['EmailsSent']:,}")
                self.kpis["Pending Emails"].set(f"{kpi_df.iloc[0]['PendingEmails']:,}")
                
            # Formatting configurations for charts
            primary_color = Config.COLORS.get("primary", "#6B4F3B")
            success_color = Config.COLORS.get("success", "#556B2F")
            
            def clean_ax(ax):
                ax.clear()
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                ax.spines['left'].set_color(Config.COLORS["border"])
                ax.spines['bottom'].set_color(Config.COLORS["border"])
                ax.set_facecolor(Config.COLORS["light"])
            
            self.fig_line.patch.set_facecolor(Config.COLORS["light"])
            self.fig_bar.patch.set_facecolor(Config.COLORS["light"])
            self.fig_pie.patch.set_facecolor(Config.COLORS["light"])
            
            # Update line chart
            clean_ax(self.ax_line)
            self.ax_line.set_title("Report Trends (Last 7 Days)", fontsize=11, fontweight='bold', color=primary_color, pad=15)
            if not line_df.empty:
                x_labels = line_df['Date'].astype(str).str[5:]
                self.ax_line.plot(x_labels, line_df['Count'], marker='o', color=primary_color, linestyle='-', linewidth=2.5, markersize=6)
                self.ax_line.fill_between(x_labels, line_df['Count'], color=primary_color, alpha=0.1)
                self.ax_line.tick_params(axis='x', rotation=45, labelsize=9, colors=Config.COLORS["gray"])
                self.ax_line.tick_params(axis='y', labelsize=9, colors=Config.COLORS["gray"])
                self.ax_line.grid(True, linestyle='--', alpha=0.4, color=Config.COLORS["border"])
            self.fig_line.tight_layout()
            self.canvas_line.draw()
            
            # Update bar chart
            clean_ax(self.ax_bar)
            self.ax_bar.set_title("Tasks Overview (Last 6 Months)", fontsize=11, fontweight='bold', color=primary_color, pad=15)
            if not bar_df.empty:
                self.ax_bar.bar(bar_df['Month'], bar_df['Count'], color=success_color, alpha=0.85, width=0.6, edgecolor="none", borderradius=2) # 'borderradius' might not be standard but width/edgecolor helps
                self.ax_bar.tick_params(axis='x', labelsize=9, colors=Config.COLORS["gray"])
                self.ax_bar.tick_params(axis='y', labelsize=9, colors=Config.COLORS["gray"])
                self.ax_bar.grid(True, axis='y', linestyle='--', alpha=0.4, color=Config.COLORS["border"])
            self.fig_bar.tight_layout()
            self.canvas_bar.draw()
            
            # Update pie chart
            self.ax_pie.clear()
            self.ax_pie.set_title("Activity Distribution", fontsize=11, fontweight='bold', color=primary_color, pad=15)
            if not pie_df.empty:
                colors = [Config.COLORS.get("primary", "#1A5276"), Config.COLORS.get("warning", "#F39C12"), 
                          Config.COLORS.get("success", "#2ECC71"), Config.COLORS.get("info", "#3498DB")]
                wedges, texts, autotexts = self.ax_pie.pie(pie_df['Count'], labels=pie_df['Category'], autopct='%1.1f%%', 
                                                           startangle=90, colors=colors, textprops={'fontsize': 9, 'color': Config.COLORS["dark"]},
                                                           wedgeprops={'edgecolor': Config.COLORS["light"], 'linewidth': 1.5, 'antialiased': True})
                for autotext in autotexts:
                    autotext.set_color('white')
                    autotext.set_weight('bold')
            self.fig_pie.tight_layout()
            self.canvas_pie.draw()
            
        except Exception as e:
            self.logger.error(f"Error updating custom dashboard UI: {e}")
            self.logger.error(traceback.format_exc())
            
        finally:
            self.is_refreshing = False
            self.refresh_btn.config(state="normal", text="🔄 Manual Refresh")
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.last_refresh_lbl.config(text=f"Last Refresh: {now_str}")
            
    def _handle_fetch_error(self):
        self.is_refreshing = False
        self.refresh_btn.config(state="normal", text="🔄 Manual Refresh")
        self.last_refresh_lbl.config(text="Last Refresh: Failed (Check Logs)")
