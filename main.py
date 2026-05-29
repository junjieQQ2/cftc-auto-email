"""CFTC 贵金属持仓分析 - 入口"""

import warnings
from datetime import datetime

import matplotlib
import matplotlib.pyplot as plt
import pandas as pd

from cftc_report.config import (
    CFTC_XLS_FOLDER, CFTC_ZIP_FOLDER, OUTPUT_EXCEL, OUTPUT_FOLDER
)
from cftc_report.data import download_all_cftc_data, fetch_precious_metals_prices
from cftc_report.processing import add_2014_2020_sheets, process_and_write
from cftc_report.charts import generate_plots
from cftc_report.analysis import backtest_analysis, current_signal_analysis
from cftc_report.mail import send_email

warnings.filterwarnings('ignore')

# ====================== 主程序 ======================
if __name__ == "__main__":
    print(f"\n=== 开始 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n")
    for folder in [CFTC_XLS_FOLDER, CFTC_ZIP_FOLDER, OUTPUT_FOLDER]:
        folder.mkdir(parents=True, exist_ok=True)

    price_dfs = fetch_precious_metals_prices()
    current_year = datetime.now().year
    (CFTC_ZIP_FOLDER / f"com_disagg_xls_{current_year}.zip").unlink(missing_ok=True)
    download_all_cftc_data()

    with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
        all_data = process_and_write(writer, price_dfs)
        add_2014_2020_sheets(writer, all_data, price_dfs)

    generate_plots()
    backtest_analysis()
    current_analysis = current_signal_analysis()
    send_email(current_analysis)

    print("\n=== 完成 ===\n")
