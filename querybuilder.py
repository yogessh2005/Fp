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

from config import Config
from grntargetmanager import GRNTargetManager

# ==================== QUERY BUILDER ====================

class QueryBuilder:
    def __init__(self, company: str):
        self.company = company
        self.dbname = Config.DATABASES[company]
   
    def get_query(self, module_name: str, date: str) -> str:
        sql_date = f"{date[4:8]}-{date[2:4]}-{date[0:2]}"
       
        queries = {
            "Store": self.get_store_query(sql_date),
            "Overall": self.get_overall_query(sql_date),
            "SessionWise": self.get_sessionwise_query(sql_date),
            "Sale": self.get_sale_query(sql_date),
            "Score": self.get_score_query(sql_date)  # Updated with new PSERM logic
        }
       
        return queries.get(module_name, "")
   
    def get_store_query(self, sql_date):
        from grntargetmanager import GRNTargetManager
        target_val = GRNTargetManager.get_target()
        return f"""
DECLARE @TargetDate DATE = '{sql_date}';

WITH TodayGRN AS (
    SELECT Gh.GrnStoreId, COUNT(*) AS TodayGRNCount
    FROM GrnHeaders Gh
    WHERE CAST(Gh.GRNDate AS DATE) = @TargetDate
    GROUP BY Gh.GrnStoreId
),
TodayPSE AS (
    SELECT Ph.PSEStoreId, COUNT(CASE WHEN Pd.PSEItemType = 'RM' THEN 1 END) AS PSERMCount
    FROM PhysicalStockEntryHeaders Ph
    JOIN PhysicalStockEntryDetails Pd ON Pd.PSEHeaderId = Ph.PSEHeaderId
    WHERE Ph.PSEEntryType <> 'ADHOC'
      AND CAST(Ph.PSEDate AS DATE) = @TargetDate
    GROUP BY Ph.PSEStoreId
)
SELECT
    PU.ProductionUnitName AS PUName,
    ISNULL(TG.TodayGRNCount, 0) AS TodayGRNCount,
    {target_val} AS Target,
    ISNULL(TG.TodayGRNCount, 0) - {target_val} AS GRNCountDifference,
    ISNULL(TP.PSERMCount, 0) AS PSERMCount
FROM ProductionUnits PU
LEFT JOIN TodayGRN TG ON TG.GrnStoreId = PU.ProductionUnitId
LEFT JOIN TodayPSE TP ON TP.PSEStoreId = PU.ProductionUnitId
WHERE PU.IsOutlet = 0 AND PU.IsActive = 1
ORDER BY PU.ProductionUnitName"""
   
    def get_overall_query(self, sql_date):
        return f"""
DECLARE @TargetDate DATE = '{sql_date}';

WITH FCData AS (
    SELECT
        PParent.ProductionUnitId,
        PParent.TargetFCPercent AS TargetValue,
        ROUND(SUM(ISNULL(O.OrderValue,0)),2) AS OrderValue,
        ROUND(SUM(ISNULL(O.SaleValue,0)),2) AS SaleValue,
        ROUND(SUM(ISNULL(O.OrderValue,0)) * 100.0 / NULLIF(SUM(ISNULL(O.SaleValue,0)),0),2) AS ESTPercent
    FROM (
        SELECT
            Om.OrderId,
            Om.OutletId,
            (Om.NoOfPax * Om.SaleRate) AS SaleValue,
            (SELECT SUM(Od.OrderQty * Od.IdealCost)
             FROM OutletMenuOrderDetails Od
             WHERE Od.OrderId = Om.OrderId) AS OrderValue
        FROM OutletMenuOrders Om
        WHERE Om.RefOrderNo > 0 AND Om.OrderStatus <> 'Cancelled'
          AND CAST(Om.OrderDate AS DATE) = @TargetDate
    ) O
    JOIN ProductionUnits Pu ON Pu.ProductionUnitId = O.OutletId
    JOIN ProductionUnits PParent ON PParent.ProductionUnitId = Pu.ParentProductionUnitId
    GROUP BY PParent.ProductionUnitId, PParent.TargetFCPercent
)
SELECT
    PU.ProductionUnitName AS PUName,
    ROUND(ISNULL(F.OrderValue,0),2) AS OrderValue,
    ROUND(ISNULL(F.SaleValue,0),2) AS SaleValue,
    CAST(ROUND(ISNULL(F.ESTPercent,0),2) AS VARCHAR(20)) + '%' AS ActualESTPercent,
    CAST(ROUND(ISNULL(F.TargetValue, PU.TargetFCPercent), 0) AS VARCHAR(20)) + '%' AS TargetFCPercent,
    CAST(ROUND(ISNULL(F.ESTPercent,0) - ISNULL(F.TargetValue, PU.TargetFCPercent), 2) AS VARCHAR(20)) + '%' AS Difference,
    CASE
        WHEN ISNULL(F.OrderValue,0) = 0 AND ISNULL(F.SaleValue,0) = 0 THEN 'Not Available'
        WHEN ISNULL(F.ESTPercent,0) BETWEEN (ISNULL(F.TargetValue, PU.TargetFCPercent) - 10)
             AND (ISNULL(F.TargetValue, PU.TargetFCPercent) + 5) THEN 'Achieved'
        ELSE 'Not Achieved'
    END AS Status
FROM ProductionUnits PU
LEFT JOIN FCData F ON F.ProductionUnitId = PU.ProductionUnitId
WHERE PU.IsActive = 1 AND PU.IsOutlet = 0
ORDER BY PU.ProductionUnitName"""
   
    def get_sessionwise_query(self, sql_date):
        return f"""
DECLARE @TargetDate DATE = '{sql_date}';

WITH OrderData AS (
    SELECT
        Om.OrderId,
        Om.OutletId,
        Om.SessionId,
        (Om.NoOfPax * Om.SaleRate) AS SaleValue,
        (SELECT SUM(Od.OrderQty * Od.IdealCost)
         FROM OutletMenuOrderDetails Od
         WHERE Od.OrderId = Om.OrderId) AS OrderValue
    FROM OutletMenuOrders Om
    WHERE Om.RefOrderNo > 0 AND Om.OrderStatus <> 'Cancelled'
      AND CAST(Om.OrderDate AS DATE) = @TargetDate
),
AggregatedData AS (
    SELECT
        Pu.ParentProductionUnitId AS ProductionUnitId,
        S.SessionId,
        S.SessionName,
        SUM(ISNULL(O.OrderValue, 0)) AS OrderValue,
        SUM(ISNULL(O.SaleValue, 0)) AS SaleValue,
        SUM(ISNULL(O.OrderValue, 0)) * 100.0 / NULLIF(SUM(ISNULL(O.SaleValue, 0)), 0) AS ActualESTPercent
    FROM OrderData O
    JOIN ProductionUnits Pu ON Pu.ProductionUnitId = O.OutletId
    JOIN Sessions S ON S.SessionId = O.SessionId
    GROUP BY Pu.ParentProductionUnitId, S.SessionId, S.SessionName
)
SELECT
    PU.ProductionUnitName AS PUName,
    A.SessionName,
    ROUND(A.OrderValue, 2) AS OrderValue,
    ROUND(A.SaleValue, 2) AS SaleValue,
    CAST(ROUND(A.ActualESTPercent, 2) AS VARCHAR(20)) + '%' AS ActualESTPercent
FROM ProductionUnits PU
INNER JOIN AggregatedData A ON A.ProductionUnitId = PU.ProductionUnitId
WHERE PU.IsActive = 1 AND PU.IsOutlet = 0
ORDER BY PU.ProductionUnitName, A.SessionId"""
   
    def get_sale_query(self, sql_date):
        return f"""
DECLARE @TargetDate DATE = '{sql_date}';

WITH OrderData AS (
    SELECT
        PParent.ProductionUnitId AS ParentId,
        Pu.ProductionUnitId AS OutletId,
        Pu.ProductionUnitName AS OutletName,
        Sn.SessionId,
        Sn.SessionName,
        ROUND(SUM(ISNULL(O.OrderValue, 0)), 2) AS ESTOrderValue,
        ROUND(SUM(ISNULL(O.SaleValue, 0)), 2) AS ActualSales
    FROM (
        SELECT
            Om.OrderId,
            Om.OutletId,
            Om.SessionId,
            (Om.NoOfPax * Om.SaleRate) AS SaleValue,
            (SELECT SUM(Od.OrderQty * Od.IdealCost)
             FROM OutletMenuOrderDetails Od
             WHERE Od.OrderId = Om.OrderId) AS OrderValue
        FROM OutletMenuOrders Om
        WHERE Om.RefOrderNo > 0 AND Om.OrderStatus <> 'Cancelled'
          AND CAST(Om.OrderDate AS DATE) = @TargetDate
    ) O
    JOIN ProductionUnits Pu ON Pu.ProductionUnitId = O.OutletId
    JOIN ProductionUnits PParent ON PParent.ProductionUnitId = Pu.ParentProductionUnitId
    JOIN Sessions Sn ON Sn.SessionId = O.SessionId
    GROUP BY PParent.ProductionUnitId, Pu.ProductionUnitId, Pu.ProductionUnitName, Sn.SessionId, Sn.SessionName
),
SalesData AS (
    SELECT
        PPu.ProductionUnitId AS ParentId,
        Pu.ProductionUnitId AS OutletId,
        Pu.ProductionUnitName AS OutletName,
        Sn.SessionId,
        Sn.SessionName,
        ROUND(SUM(fd.SalesCost), 2) AS Sales
    FROM FoodCostHeaders fh
    JOIN FoodCostDetails fd ON fd.FoodCostId = fh.FoodCostId
    JOIN ProductionUnits Pu ON Pu.ProductionUnitId = fh.ProductionUnitId
    JOIN ProductionUnits PPu ON PPu.ProductionUnitId = Pu.ParentProductionUnitId
    JOIN Sessions Sn ON Sn.SessionId = fd.SessionId
    WHERE CAST(fh.FoodCostDate AS DATE) = @TargetDate
      AND fd.SalesCost > 0
    GROUP BY PPu.ProductionUnitId, Pu.ProductionUnitId, Pu.ProductionUnitName, Sn.SessionId, Sn.SessionName
),
Combined AS (
    SELECT
        ISNULL(O.ParentId, S.ParentId) AS ParentId,
        ISNULL(O.OutletId, S.OutletId) AS OutletId,
        ISNULL(O.OutletName, S.OutletName) AS OutletName,
        ISNULL(O.SessionId, S.SessionId) AS SessionId,
        ISNULL(O.SessionName, S.SessionName) AS SessionName,
        ISNULL(O.ESTOrderValue, 0) AS ESTOrderValue,
        ISNULL(O.ActualSales, 0) AS ActualSales,
        ISNULL(S.Sales, 0) AS Sales
    FROM OrderData O
    FULL OUTER JOIN SalesData S ON S.ParentId = O.ParentId
        AND S.OutletId = O.OutletId
        AND S.SessionId = O.SessionId
)
SELECT
    PPu.ProductionUnitName AS PUName,
    C.OutletName,
    C.SessionName,
    C.ESTOrderValue,
    C.ActualSales,
    C.Sales,
    CASE
        WHEN C.Sales = 0 THEN 0
        ELSE ROUND(C.ESTOrderValue * 100.0 / C.Sales, 2)
    END AS FC_Percent
FROM ProductionUnits PPu
JOIN Combined C ON C.ParentId = PPu.ProductionUnitId
WHERE PPu.IsActive = 1 AND PPu.IsOutlet = 0
  AND (C.ESTOrderValue <> 0 OR C.ActualSales <> 0 OR C.Sales <> 0)
ORDER BY PPu.ProductionUnitName, C.OutletName, C.SessionId"""
   
    def get_score_query(self, sql_date):
        # ⭐ Get GRN target from file
        grn_target = GRNTargetManager.get_target()
       
        return f"""
DECLARE @TargetDate DATE = '{sql_date}';
DECLARE @PrevMonthStart DATE = DATEADD(MONTH, -1, @TargetDate);
DECLARE @MonthStart DATE = DATEADD(DAY, 1 - DAY(@TargetDate), @TargetDate);
DECLARE @GRN_Target FLOAT = {grn_target};

WITH StoreData AS (
    SELECT
        PU.ProductionUnitId,
        PU.ProductionUnitName AS PUName,
        ISNULL(TG.TodayGRNCount, 0) AS TodayGRNCount,
        ISNULL(TG.TodayGRNCount, 0) - CAST(ISNULL(PM.AvgGRNCount, 0) AS INT) AS GRNCountDifference,
        ISNULL(TP.PSERMCount, 0) AS PSERMCount,
        ISNULL(PM.AvgGRNCount, 0) AS PrevMonthAvgGRNCount
    FROM ProductionUnits PU
    LEFT JOIN (
        SELECT Gh.GrnStoreId, COUNT(*) AS TodayGRNCount
        FROM GrnHeaders Gh
        WHERE CAST(Gh.GRNDate AS DATE) = @TargetDate
        GROUP BY Gh.GrnStoreId
    ) TG ON TG.GrnStoreId = PU.ProductionUnitId
    LEFT JOIN (
        SELECT GrnStoreId, ROUND(AVG(CAST(DailyCount AS DECIMAL(18,2))), 0) AS AvgGRNCount
        FROM (
            SELECT Gh.GrnStoreId, CAST(Gh.GRNDate AS DATE) AS GRNDate, COUNT(*) AS DailyCount
            FROM GrnHeaders Gh
            WHERE CAST(Gh.GRNDate AS DATE) >= @PrevMonthStart
              AND CAST(Gh.GRNDate AS DATE) < @MonthStart
            GROUP BY Gh.GrnStoreId, CAST(Gh.GRNDate AS DATE)
        ) DailyCounts
        GROUP BY GrnStoreId
    ) PM ON PM.GrnStoreId = PU.ProductionUnitId
    LEFT JOIN (
        SELECT Ph.PSEStoreId, COUNT(CASE WHEN Pd.PSEItemType = 'RM' THEN 1 END) AS PSERMCount
        FROM PhysicalStockEntryHeaders Ph
        JOIN PhysicalStockEntryDetails Pd ON Pd.PSEHeaderId = Ph.PSEHeaderId
        WHERE Ph.PSEEntryType <> 'ADHOC' AND CAST(Ph.PSEDate AS DATE) = @TargetDate
        GROUP BY Ph.PSEStoreId
    ) TP ON TP.PSEStoreId = PU.ProductionUnitId
    WHERE PU.IsOutlet = 0 AND PU.IsActive = 1
),
OverallData AS (
    SELECT
        PU.ProductionUnitId,
        PU.ProductionUnitName AS PUName,
        PU.TargetFCPercent,
        ROUND(ISNULL(F.ESTPercent,0) - ISNULL(PU.TargetFCPercent,0), 2) AS OverallDiff
    FROM ProductionUnits PU
    LEFT JOIN (
        SELECT PParent.ProductionUnitId, ROUND(SUM(ISNULL(O.OrderValue,0)) * 100.0 / NULLIF(SUM(ISNULL(O.SaleValue,0)),0),2) AS ESTPercent
        FROM (
            SELECT Om.OrderId, Om.OutletId, (Om.NoOfPax * Om.SaleRate) AS SaleValue,
                   (SELECT SUM(Od.OrderQty * Od.IdealCost) FROM OutletMenuOrderDetails Od WHERE Od.OrderId = Om.OrderId) AS OrderValue
            FROM OutletMenuOrders Om
            WHERE Om.RefOrderNo > 0 AND Om.OrderStatus <> 'Cancelled'
              AND CAST(Om.OrderDate AS DATE) = @TargetDate
        ) O
        JOIN ProductionUnits Pu2 ON Pu2.ProductionUnitId = O.OutletId
        JOIN ProductionUnits PParent ON PParent.ProductionUnitId = Pu2.ParentProductionUnitId
        GROUP BY PParent.ProductionUnitId
    ) F ON F.ProductionUnitId = PU.ProductionUnitId
    WHERE PU.IsActive = 1 AND PU.IsOutlet = 0
),
SessionRaw AS (
    SELECT
        Pu2.ParentProductionUnitId AS ProductionUnitId,
        S.SessionId,
        SUM(ISNULL(O.OrderValue, 0)) * 100.0 / NULLIF(SUM(ISNULL(O.SaleValue, 0)), 0) AS ActualESTPercent
    FROM (
        SELECT Om.OrderId, Om.OutletId, Om.SessionId, (Om.NoOfPax * Om.SaleRate) AS SaleValue,
               (SELECT SUM(Od.OrderQty * Od.IdealCost) FROM OutletMenuOrderDetails Od WHERE Od.OrderId = Om.OrderId) AS OrderValue
        FROM OutletMenuOrders Om
        WHERE Om.RefOrderNo > 0 AND Om.OrderStatus <> 'Cancelled'
          AND CAST(Om.OrderDate AS DATE) = @TargetDate
    ) O
    JOIN ProductionUnits Pu2 ON Pu2.ProductionUnitId = O.OutletId
    JOIN Sessions S ON S.SessionId = O.SessionId
    GROUP BY Pu2.ParentProductionUnitId, S.SessionId
),
SessionWiseData AS (
    SELECT ProductionUnitId, AVG(ActualESTPercent) AS AvgSessionActualEST
    FROM SessionRaw
    GROUP BY ProductionUnitId
),
SalesData AS (
    SELECT
        PPu.ProductionUnitId,
        PPu.ProductionUnitName AS PUName,
        ROUND(SUM(fd.SalesCost), 2) AS TotalSales
    FROM FoodCostHeaders fh
    JOIN FoodCostDetails fd ON fd.FoodCostId = fh.FoodCostId
    JOIN ProductionUnits Pu ON Pu.ProductionUnitId = fh.ProductionUnitId
    JOIN ProductionUnits PPu ON PPu.ProductionUnitId = Pu.ParentProductionUnitId
    WHERE CAST(fh.FoodCostDate AS DATE) = @TargetDate AND fd.SalesCost > 0
    GROUP BY PPu.ProductionUnitId, PPu.ProductionUnitName
)
SELECT
    PUName,
    GRN_Mark,
    Closing_Stock_Mark,
    Estimate_Overall_Mark,
    SessionWise_Mark,
    Sales_Mark,
    GRN_Mark + Closing_Stock_Mark + Estimate_Overall_Mark + SessionWise_Mark + Sales_Mark AS Overall_Result
FROM (
    SELECT
        PU.ProductionUnitId,
        PU.ProductionUnitName AS PUName,
       
        -- ⭐ GRN Scoring with Target
        CASE
            WHEN ISNULL(S.TodayGRNCount, 0) > @GRN_Target THEN 25
            WHEN ISNULL(S.TodayGRNCount, 0) = @GRN_Target THEN 20
            WHEN ISNULL(S.TodayGRNCount, 0) < @GRN_Target THEN 0
            ELSE 0
        END AS GRN_Mark,
       
        -- ⭐ UPDATED CLOSING STOCK / PSERM LOGIC
        CASE
            WHEN ISNULL(S.PSERMCount, 0) > 35 THEN 25
            WHEN ISNULL(S.PSERMCount, 0) >= 30 THEN 20
            WHEN ISNULL(S.PSERMCount, 0) >= 25 THEN 15
            WHEN ISNULL(S.PSERMCount, 0) >= 20 THEN 10
            ELSE 0
        END AS Closing_Stock_Mark,
       
        CASE
            WHEN O.OverallDiff IS NULL THEN 0
            WHEN O.OverallDiff BETWEEN -5 AND 5 THEN 15
            WHEN O.OverallDiff BETWEEN -10 AND -6 OR O.OverallDiff BETWEEN 6 AND 10 THEN 10
            WHEN O.OverallDiff BETWEEN -15 AND -11 OR O.OverallDiff BETWEEN 11 AND 15 THEN 5
            ELSE 2
        END AS Estimate_Overall_Mark,
       
        -- ⭐ Session Wise mark now mirrors Estimate Overall mark
        CASE
            WHEN O.OverallDiff IS NULL THEN 0
            WHEN O.OverallDiff BETWEEN -5 AND 5 THEN 15
            WHEN O.OverallDiff BETWEEN -10 AND -6 OR O.OverallDiff BETWEEN 6 AND 10 THEN 10
            WHEN O.OverallDiff BETWEEN -15 AND -11 OR O.OverallDiff BETWEEN 11 AND 15 THEN 5
            ELSE 2
        END AS SessionWise_Mark,
       
        CASE
            WHEN ISNULL(SD.TotalSales, 0) = 0 THEN 0
            ELSE 20
        END AS Sales_Mark
       
    FROM ProductionUnits PU
    LEFT JOIN StoreData S ON S.ProductionUnitId = PU.ProductionUnitId
    LEFT JOIN OverallData O ON O.ProductionUnitId = PU.ProductionUnitId
    LEFT JOIN SessionWiseData SS ON SS.ProductionUnitId = PU.ProductionUnitId
    LEFT JOIN SalesData SD ON SD.ProductionUnitId = PU.ProductionUnitId
    WHERE PU.IsActive = 1 AND PU.IsOutlet = 0
) AS CombinedScore
ORDER BY PUName"""


