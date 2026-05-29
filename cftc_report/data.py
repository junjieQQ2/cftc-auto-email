"""CFTC 数据下载与价格获取"""

import shutil
import zipfile
from datetime import datetime

import akshare as ak
import pandas as pd
import requests

from .config import (
    BASE_URL, CFTC_XLS_FOLDER, CFTC_ZIP_FOLDER,
    DATE_COL, HIST_ZIP, MARKET_COL,
)


# ====================== 获取价格数据（美元计价） ======================
def fetch_precious_metals_prices():
    """获取黄金 (XAU)、白银 (XAG)、铂金 (XPT) 国际现货美元价格"""
    print("AKShare 版本:", ak.__version__)
    price_dfs = {}
    symbol_map = [("XAU", "Gold"), ("XAG", "Silver"), ("XPT", "Platinum")]

    for symbol, metal in symbol_map:
        try:
            df = ak.futures_foreign_hist(symbol=symbol)
            date_col = next((c for c in df.columns if 'date' in c.lower()), None)
            close_col = next((c for c in df.columns if 'close' in c.lower() or 'settle' in c.lower()), None)
            if date_col and close_col:
                df = df.rename(columns={date_col: 'date', close_col: 'close'})
                df['date'] = pd.to_datetime(df['date'])
                price_dfs[metal] = df[['date', 'close']].sort_values('date').set_index('date')
                print(f"✓ {metal} 价格数据加载成功 (USD)")
        except Exception as e:
            print(f"{metal} price failed: {e}")

    return price_dfs


# ====================== 下载 CFTC 数据 ======================
def download_and_unzip(url, zip_dir, extract_dir, year=None, force=False):
    filename = url.split("/")[-1]
    zip_path = zip_dir / filename
    if not force and (extract_dir / f"com_disagg_{year or 'hist'}.xls").exists():
        return
    if not zip_path.exists() or force:
        r = requests.get(url, stream=True, timeout=60)
        r.raise_for_status()
        with open(zip_path, "wb") as f:
            for chunk in r.iter_content(8192):
                f.write(chunk)
    with zipfile.ZipFile(zip_path) as z:
        for member in z.namelist():
            if member.lower().endswith((".xls", ".xlsx")):
                z.extract(member, extract_dir)
                if year:
                    new_path = extract_dir / f"com_disagg_{year}.xls"
                    shutil.move(extract_dir / member, new_path)


def download_all_cftc_data():
    download_and_unzip(BASE_URL + HIST_ZIP, CFTC_ZIP_FOLDER, CFTC_XLS_FOLDER)
    current_year = datetime.now().year
    for year in range(2017, current_year):
        download_and_unzip(BASE_URL + f"com_disagg_xls_{year}.zip", CFTC_ZIP_FOLDER, CFTC_XLS_FOLDER, year=year)
    download_and_unzip(BASE_URL + f"com_disagg_xls_{current_year}.zip", CFTC_ZIP_FOLDER, CFTC_XLS_FOLDER, year=current_year, force=True)
