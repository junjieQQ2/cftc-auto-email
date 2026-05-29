"""配置常量"""

import os
from datetime import datetime, timedelta
from pathlib import Path

# ====================== 目录和文件 ======================
BASE_DIR = Path.cwd()
CFTC_XLS_FOLDER = BASE_DIR / "cftc_xls_all"
CFTC_ZIP_FOLDER = BASE_DIR / "cftc_zips"
OUTPUT_FOLDER = BASE_DIR / "cftc_plots_percentile"
OUTPUT_EXCEL = BASE_DIR / f"CFTC_Gold_Silver_Platinum_With_Prices_{datetime.now().strftime('%Y%m%d')}.xlsx"

# ====================== CFTC 数据源 ======================
BASE_URL = "https://www.cftc.gov/files/dea/history/"
HIST_ZIP = "com_disagg_xls_hist_2006_2016.zip"

METAL_FILTERS = {
    "Gold":     ["GOLD - COMMODITY EXCHANGE INC."],
    "Silver":   ["SILVER - COMMODITY EXCHANGE INC."],
    "Platinum": ["PLATINUM - NEW YORK MERCANTILE EXCHANGE"]
}

CATEGORIES = ["Producers", "Swap_Dealers", "Money_Managers", "Other_Reportables", "Non_Reportables"]

DATE_COL = 'Report_Date_as_MM_DD_YYYY'
OI_COL = 'Open_Interest_All'
MARKET_COL = 'Market_and_Exchange_Names'

TIME_PERIODS = {
    "1Y": timedelta(days=365),
    "3Y": timedelta(days=3*365),
    "5Y": timedelta(days=5*365),
    "10Y": timedelta(days=10*365)
}

# ====================== 回溯 sheet 名称 ======================
BACKTEST_SHEET = "Backtest_Analysis"

# ====================== 邮件配置（从环境变量读取，避免泄露） ======================
EMAIL_FROM = os.environ.get('EMAIL_FROM', '')
EMAIL_TO = os.environ.get('EMAIL_TO', '')
EMAIL_PASSWORD = os.environ.get('EMAIL_PASSWORD', '')
SMTP_SERVER = os.environ.get('SMTP_SERVER', 'smtp.qq.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '465'))
