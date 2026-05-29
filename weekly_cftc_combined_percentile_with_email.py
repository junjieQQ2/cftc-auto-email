"""CFTC 贵金属持仓分析报告 - 自动下载数据、生成 Excel + 图表、发送邮件"""

import os
import shutil
import smtplib
import warnings
import zipfile
from datetime import datetime, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import akshare as ak
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
import pandas as pd
import requests
from matplotlib.dates import DateFormatter
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# =============================================================================
# 配置常量
# =============================================================================

BASE_DIR = Path.cwd()
CFTC_XLS_FOLDER = BASE_DIR / "cftc_xls_all"
CFTC_ZIP_FOLDER = BASE_DIR / "cftc_zips"
OUTPUT_FOLDER = BASE_DIR / "cftc_plots_percentile"
OUTPUT_EXCEL = BASE_DIR / f"CFTC_Gold_Silver_Platinum_With_Prices_{datetime.now().strftime('%Y%m%d')}.xlsx"

CFTC_BASE_URL = "https://www.cftc.gov/files/dea/history/"
CFTC_HIST_ZIP = "com_disagg_xls_hist_2006_2016.zip"

METALS = {
    "Gold": ["GOLD - COMMODITY EXCHANGE INC."],
    "Silver": ["SILVER - COMMODITY EXCHANGE INC."],
    "Platinum": ["PLATINUM - NEW YORK MERCANTILE EXCHANGE"],
}

TRADER_TYPES = ["Producers", "Swap_Dealers", "Money_Managers", "Other_Reportables", "Non_Reportables"]

# 各 trader 在 CFTC 原始文件中的多空列名
POSITION_COLUMNS = {
    "Producers": ("Prod_Merc_Positions_Long_ALL", "Prod_Merc_Positions_Short_ALL"),
    "Swap_Dealers": ("Swap_Positions_Long_All", "Swap__Positions_Short_All"),
    "Money_Managers": ("M_Money_Positions_Long_ALL", "M_Money_Positions_Short_ALL"),
    "Other_Reportables": ("Other_Rept_Positions_Long_ALL", "Other_Rept_Positions_Short_ALL"),
    "Non_Reportables": ("NonRept_Positions_Long_All", "NonRept_Positions_Short_All"),
}

ROLLING_PERIODS = {
    "1Y": timedelta(days=365),
    "3Y": timedelta(days=3 * 365),
    "5Y": timedelta(days=5 * 365),
    "10Y": timedelta(days=10 * 365),
}

DATE_COL = "Report_Date_as_MM_DD_YYYY"
OI_COL = "Open_Interest_All"
MARKET_COL = "Market_and_Exchange_Names"

EMAIL_FROM = os.environ.get("EMAIL_FROM", "")
EMAIL_TO = os.environ.get("EMAIL_TO", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
SMTP_SERVER = os.environ.get("SMTP_SERVER", "smtp.qq.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))

BACKTEST_SHEET = "Backtest_Analysis"


# =============================================================================
# 通用工具函数
# =============================================================================

def percentile_color(value: float) -> str:
    """分位值 → matplotlib 颜色名"""
    if value >= 80:
        return "darkred"
    elif value >= 50:
        return "orange"
    elif value >= 20:
        return "blue"
    else:
        return "darkgreen"


def percentile_fill(value: float) -> PatternFill:
    """分位值 → Excel 单元格填充色"""
    if value >= 80:
        return PatternFill("solid", start_color="FF0000")
    elif value >= 50:
        return PatternFill("solid", start_color="FFA500")
    elif value >= 20:
        return PatternFill("solid", start_color="0000FF")
    elif value <= 20:
        return PatternFill("solid", start_color="00FF00")
    return PatternFill("solid", start_color="D3D3D3")


def percentile_signal(pct_value: float) -> str:
    """分位值 → 信号文字描述"""
    mapping = {
        (80, 101): "极度拥挤多头",
        (50, 80): "强势多头 / 拥挤多头",
        (20, 50): "弱势空头 / 拥挤空头",
        (0, 20): "极度拥挤空头",
    }
    for (lo, hi), label in mapping.items():
        if lo <= pct_value < hi:
            return label
    return "中性"


def percentile_color_series(series: pd.Series) -> list:
    """分位数列 → matplotlib 散点颜色列表"""
    return [percentile_color(v) for v in series]


def apply_excel_percentile_colors(sheet, columns, n_rows):
    """给 Excel sheet 中所有分位列的单元格填色"""
    for i, col_name in enumerate(columns):
        if "Pct_OI_Percentile" not in col_name:
            continue
        letter = get_column_letter(i + 1)
        for r in range(2, n_rows + 2):
            cell = sheet[f"{letter}{r}"]
            try:
                cell.fill = percentile_fill(float(cell.value))
            except (ValueError, TypeError):
                pass


def match_price_to_cftc(df: pd.DataFrame, price_series: pd.Series) -> None:
    """根据 CFTC 报告日（周二回溯 3 天）匹配期货价格"""
    if price_series is None or price_series.empty:
        return
    price_series = price_series.sort_index()
    df["Tuesday_Date"] = df.index - pd.Timedelta(days=3)
    df["close"] = np.nan
    for idx, row in df.iterrows():
        tue = row["Tuesday_Date"]
        if tue in price_series.index:
            df.at[idx, "close"] = price_series.loc[tue]
        else:
            candidates = price_series.index[price_series.index < tue]
            if not candidates.empty:
                df.at[idx, "close"] = price_series.loc[candidates.max()]
    df["close"] = df["close"].ffill().bfill()


def compute_pct_oi(df: pd.DataFrame) -> None:
    """根据原始多空持仓列计算 Net / Pct_OI"""
    for cat, (long_col, short_col) in POSITION_COLUMNS.items():
        if long_col in df.columns and short_col in df.columns:
            df[f"{cat}_Long"] = df[long_col]
            df[f"{cat}_Short"] = df[short_col]
            df[f"{cat}_Net"] = df[long_col] - df[short_col]
            df[f"{cat}_Pct_OI"] = df[f"{cat}_Net"] / df[OI_COL] * 100


def compute_percentiles(df: pd.DataFrame) -> None:
    """添加全历史 + 滚动窗口（1Y/3Y/5Y/10Y）分位列"""
    for cat in TRADER_TYPES:
        pct_col = f"{cat}_Pct_OI"
        if pct_col not in df.columns:
            continue
        clipped = df[pct_col].clip(df[pct_col].quantile(0.01), df[pct_col].quantile(0.99))
        df[f"{cat}_Pct_OI_Percentile"] = clipped.rank(pct=True) * 100

    latest = df["Report_Date"].max()
    for p_name, delta in ROLLING_PERIODS.items():
        mask = df["Report_Date"] >= (latest - delta)
        for cat in TRADER_TYPES:
            pct_col = f"{cat}_Pct_OI"
            if pct_col not in df.columns:
                continue
            subset = df.loc[mask, pct_col]
            clipped = subset.clip(subset.quantile(0.01), subset.quantile(0.99))
            ranks = clipped.rank(pct=True) * 100
            col_name = f"{cat}_{p_name}_Pct_OI_Percentile"
            df[col_name] = np.nan
            df.loc[mask, col_name] = ranks


def clean_cftc_data(raw_df: pd.DataFrame, start_date: str = "2007-01-01") -> pd.DataFrame:
    """清洗：排序、去重、日期过滤、设索引"""
    df = raw_df.sort_values("Report_Date").drop_duplicates("Report_Date", keep="last").copy()
    df = df[df["Report_Date"] >= start_date]
    df["Report_Date"] = pd.to_datetime(df["Report_Date"])
    return df.set_index("Report_Date")


# =============================================================================
# 数据获取
# =============================================================================

def download_cftc_data() -> None:
    """下载 2006-至今所有 CFTC 周度持仓数据（.xls 格式）"""

    def _download_one(url, zip_dir, extract_dir, year=None, force=False):
        filename = url.split("/")[-1]
        zip_path = zip_dir / filename
        extract_flag = extract_dir / f"com_disagg_{year or 'hist'}.xls"
        if not force and extract_flag.exists():
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
                        shutil.move(extract_dir / member, extract_flag)

    _download_one(CFTC_BASE_URL + CFTC_HIST_ZIP, CFTC_ZIP_FOLDER, CFTC_XLS_FOLDER)
    current_year = datetime.now().year
    for year in range(2017, current_year):
        _download_one(CFTC_BASE_URL + f"com_disagg_xls_{year}.zip", CFTC_ZIP_FOLDER, CFTC_XLS_FOLDER, year=year)
    _download_one(
        CFTC_BASE_URL + f"com_disagg_xls_{current_year}.zip",
        CFTC_ZIP_FOLDER, CFTC_XLS_FOLDER, year=current_year, force=True,
    )


def fetch_price_data() -> dict:
    """获取黄金、白银、铂金期货价格 (akshare)"""
    prices = {}
    for symbol, metal in [("AU0", "Gold"), ("AG0", "Silver")]:
        try:
            df = ak.futures_main_sina(symbol=symbol).rename(columns={"日期": "date", "收盘价": "close"})
            df["date"] = pd.to_datetime(df["date"])
            prices[metal] = df.set_index("date")[["close"]].sort_index()
            print(f"✓ {metal} 价格数据")
        except Exception as e:
            print(f"✗ {metal} 价格获取失败: {e}")

    try:
        df = ak.futures_foreign_hist(symbol="XPT")
        date_col = next((c for c in df.columns if "date" in c.lower()), None)
        close_col = next((c for c in df.columns if "close" in c.lower() or "settle" in c.lower()), None)
        if date_col and close_col:
            df = df.rename(columns={date_col: "date", close_col: "close"})
            df["date"] = pd.to_datetime(df["date"])
            prices["Platinum"] = df.set_index("date")[["close"]].sort_index()
            print(f"✓ Platinum 价格数据")
    except Exception as e:
        print(f"✗ Platinum 价格获取失败: {e}")

    return prices


def load_cftc_raw_data() -> dict:
    """读取所有本地 CFTC .xls 文件，按金属拆分成 DataFrame 字典"""
    all_data = {metal: pd.DataFrame() for metal in METALS}
    for file_path in sorted(CFTC_XLS_FOLDER.glob("*.xls")):
        try:
            df = pd.read_excel(file_path, engine="xlrd")
            if DATE_COL not in df.columns or MARKET_COL not in df.columns:
                continue
            df["Report_Date"] = pd.to_datetime(df[DATE_COL], errors="coerce")
            df = df.dropna(subset=["Report_Date"])
            compute_pct_oi(df)
            for metal, keywords in METALS.items():
                mask = df[MARKET_COL].str.upper().isin([k.upper() for k in keywords])
                if mask.any():
                    all_data[metal] = pd.concat([all_data[metal], df[mask]]).copy()
        except Exception as e:
            print(f"跳过 {file_path.name}: {e}")
    return all_data


# =============================================================================
# Excel 输出
# =============================================================================

def build_main_sheet_output(df: pd.DataFrame, metal: str, price_data: dict) -> pd.DataFrame:
    """对一个金属的全历史数据做清洗、匹配价格、计算分位，返回输出 DataFrame"""
    df = clean_cftc_data(df)
    if metal in price_data and not price_data[metal].empty:
        match_price_to_cftc(df, price_data[metal]["close"])
    df = df.reset_index()
    compute_percentiles(df)

    cols = ["Report_Date", OI_COL]
    for cat in TRADER_TYPES:
        cols += [f"{cat}_Long", f"{cat}_Short", f"{cat}_Net", f"{cat}_Pct_OI",
                 f"{cat}_Pct_OI_Percentile"]
        for p in ROLLING_PERIODS:
            cols.append(f"{cat}_{p}_Pct_OI_Percentile")
    if "close" in df.columns:
        cols.append("close")
    return df[cols].sort_values("Report_Date", ascending=False)


def build_2014_2020_sheet_output(df: pd.DataFrame, metal: str, price_data: dict) -> pd.DataFrame | None:
    """对 2014-2020 区间数据做清洗、匹配价格、计算分位和回测信号"""
    df = clean_cftc_data(df, start_date="2014-01-01")
    df = df[(df.index >= "2014-01-01") & (df.index <= "2020-12-31")]
    if len(df) < 30:
        return None
    if metal in price_data and not price_data[metal].empty:
        match_price_to_cftc(df, price_data[metal]["close"])
    df = df.reset_index()
    compute_percentiles(df)

    if "close" in df.columns and df["close"].notna().sum() > 1:
        df["Weekly_Return"] = df["close"].pct_change(1) * 100
        df["Next_Weekly_Return"] = df["close"].pct_change(1).shift(-1) * 100
    else:
        df["Weekly_Return"] = np.nan
        df["Next_Weekly_Return"] = np.nan

    for cat in TRADER_TYPES:
        pct_col = f"{cat}_Pct_OI_Percentile"
        if pct_col not in df.columns:
            continue
        df[f"{cat}_Signal"] = df[pct_col].apply(percentile_signal)
        df[f"{cat}_Week_Match"] = (((df[pct_col] >= 50) & (df["Weekly_Return"] > 0)) |
                                   ((df[pct_col] < 50) & (df["Weekly_Return"] < 0))).astype(int)
        df[f"{cat}_Next_Match"] = (((df[pct_col] >= 50) & (df["Next_Weekly_Return"] > 0)) |
                                   ((df[pct_col] < 50) & (df["Next_Weekly_Return"] < 0))).astype(int)

    cols = ["Report_Date", OI_COL]
    if "close" in df.columns:
        cols.append("close")
    for cat in TRADER_TYPES:
        cols += [f"{cat}_Long", f"{cat}_Short", f"{cat}_Net", f"{cat}_Pct_OI",
                 f"{cat}_Pct_OI_Percentile", f"{cat}_Signal",
                 f"{cat}_Week_Match", f"{cat}_Next_Match"]
    cols += ["Weekly_Return", "Next_Weekly_Return"]
    return df[cols].sort_values("Report_Date", ascending=False)


def write_excel(raw_data: dict, price_data: dict) -> dict:
    """生成主 sheet + 2014-2020 sheet 到 Excel 文件"""
    with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
        all_data = {}
        for metal in METALS:
            df = raw_data.get(metal, pd.DataFrame())
            if df.empty:
                continue
            main = build_main_sheet_output(df, metal, price_data)
            all_data[metal] = main
            main.to_excel(writer, sheet_name=metal, index=False, float_format="%.2f")
            apply_excel_percentile_colors(writer.sheets[metal], main.columns, len(main))
            print(f"  {metal} 主 sheet → {len(main)} 行")

            sheet_1420 = build_2014_2020_sheet_output(df, metal, price_data)
            if sheet_1420 is not None:
                sheet_name = f"{metal}_2014_2020"
                sheet_1420.to_excel(writer, sheet_name=sheet_name, index=False, float_format="%.2f")
                apply_excel_percentile_colors(writer.sheets[sheet_name], sheet_1420.columns, len(sheet_1420))
                print(f"  {sheet_name} sheet → {len(sheet_1420)} 行")
    return all_data


# =============================================================================
# 图表生成
# =============================================================================

def draw_percentile_chart(metal: str, df: pd.DataFrame, title_suffix: str,
                          filename_suffix: str, col_suffix: str = "Pct_OI_Percentile") -> None:
    """绘制一个 6 面板分位散点图（价格 + 5 类 trader）"""
    fig, axes = plt.subplots(6, 1, figsize=(15, 26), sharex=True,
                             gridspec_kw={"height_ratios": [1] * 6})
    fig.suptitle(f"{metal} Percentiles & Spot Price {title_suffix}", fontsize=16, fontweight="bold")

    if "close" in df.columns and not df["close"].isna().all():
        axes[0].plot(df["Report_Date"], df["close"], color="black", lw=1.5)
        axes[0].set_ylabel("Spot Price")
    axes[0].grid(True, alpha=0.3)

    for i, cat in enumerate(TRADER_TYPES):
        ax = axes[i + 1]
        col = f"{cat}_{col_suffix}"
        if col not in df.columns or df[col].isna().all():
            ax.text(0.5, 50, "No data", ha="center", va="center")
            continue
        colors = percentile_color_series(df[col])
        ax.scatter(df["Report_Date"], df[col], c=colors, s=40, alpha=0.8)
        for y in [80, 50, 20]:
            ax.axhline(y, color={80: "darkred", 50: "gray", 20: "blue"}[y],
                       ls=":" if y != 50 else "--")
        ax.set_ylabel(f"{cat} (%)")
        ax.grid(True, alpha=0.3)
        ax.set_ylim(0, 100)

    legend = [mpatches.Patch(color="darkred", label="≥80%"),
              mpatches.Patch(color="orange", label="50–79%"),
              mpatches.Patch(color="blue", label="20–49%"),
              mpatches.Patch(color="darkgreen", label="≤20%")]
    fig.legend(handles=legend, loc="lower center", bbox_to_anchor=(0.5, -0.02), ncol=1, fontsize=9)
    axes[-1].xaxis.set_major_formatter(DateFormatter("%Y-%m"))
    plt.xticks(rotation=45)
    fig.tight_layout(rect=[0, 0.05, 1, 0.92])
    filepath = OUTPUT_FOLDER / f"{metal}_{filename_suffix}.png"
    plt.savefig(filepath, dpi=180, bbox_inches="tight")
    plt.close(fig)


def generate_plots() -> None:
    """从 Excel 读取数据，生成全历史 / 2014-2020 / 1Y-3Y-5Y-10Y 分位图"""
    if not OUTPUT_EXCEL.exists():
        print("Excel 文件不存在，跳过图表生成")
        return
    xl = pd.ExcelFile(OUTPUT_EXCEL, engine="openpyxl")
    for metal in METALS:
        if metal not in xl.sheet_names:
            continue
        df = xl.parse(metal)
        df["Report_Date"] = pd.to_datetime(df["Report_Date"], errors="coerce")
        df = df.dropna(subset=["Report_Date"]).sort_values("Report_Date")
        if df.empty:
            continue

        latest = df["Report_Date"].max()
        draw_percentile_chart(metal, df, "(Full History)", "Historical_Percentiles")

        sheet_name = f"{metal}_2014_2020"
        if sheet_name in xl.sheet_names:
            df1420 = xl.parse(sheet_name)
            df1420["Report_Date"] = pd.to_datetime(df1420["Report_Date"], errors="coerce")
            df1420 = df1420.dropna(subset=["Report_Date"]).sort_values("Report_Date")
            df1420 = df1420[(df1420["Report_Date"] >= "2014-01-01") &
                            (df1420["Report_Date"] <= "2020-12-31")]
            if not df1420.empty:
                draw_percentile_chart(metal, df1420, "(2014-2020)", "2014_2020_Historical_Percentiles")

        for period in ["1Y", "3Y", "5Y", "10Y"]:
            period_df = df[df["Report_Date"] >= latest - ROLLING_PERIODS[period]].copy()
            if len(period_df) < 10:
                continue
            col_suffix = f"{period}_Pct_OI_Percentile"
            if all(f"{cat}_{col_suffix}" not in period_df.columns or
                   period_df[f"{cat}_{col_suffix}"].isna().all() for cat in TRADER_TYPES):
                continue
            draw_percentile_chart(metal, period_df, f"({period})", f"{period}_Percentiles",
                                  col_suffix=col_suffix)
    print("图表生成完成")


# =============================================================================
# 回溯分析
# =============================================================================

def backtest_analysis() -> None:
    """基于 Money_Managers 分位信号的简单回溯统计，写入 Backtest_Analysis sheet"""
    xl = pd.ExcelFile(OUTPUT_EXCEL, engine="openpyxl")
    results = []
    for metal in METALS:
        if metal not in xl.sheet_names:
            continue
        df = xl.parse(metal)
        df["Report_Date"] = pd.to_datetime(df["Report_Date"])
        df = df.sort_values("Report_Date").set_index("Report_Date")

        pct_col = "Money_Managers_Pct_OI_Percentile"
        if pct_col not in df.columns or "close" not in df.columns:
            continue

        df["Overbought"] = (df[pct_col] >= 80).astype(int)
        df["Oversold"] = (df[pct_col] <= 20).astype(int)
        df["Overbought_Signal"] = df["Overbought"].rolling(2).sum() >= 2
        df["Oversold_Signal"] = df["Oversold"].rolling(2).sum() >= 2

        for h in [1, 2, 4]:
            df[f"Return_{h}w"] = df["close"].pct_change(h).shift(-h) * 100

        for sig, label in [("Overbought_Signal", "做多拥挤 (>=80%)"),
                           ("Oversold_Signal", "做空拥挤 (<=20%)")]:
            subset = df[df[sig]]
            if subset.empty:
                continue
            stats = {"Metal": metal, "Signal_Type": label, "Count": len(subset)}
            for h in [1, 2, 4]:
                rets = subset[f"Return_{h}w"].dropna()
                stats[f"Avg_Return_{h}w"] = rets.mean() if not rets.empty else np.nan
                stats[f"Win_Rate_{h}w"] = (rets > 0).mean() * 100 if not rets.empty else np.nan
            results.append(stats)

    if results:
        result_df = pd.DataFrame(results)
        with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl", mode="a",
                            if_sheet_exists="replace") as writer:
            result_df.to_excel(writer, sheet_name=BACKTEST_SHEET, index=False)
        print("回溯分析已写入")


# =============================================================================
# 当前信号分析
# =============================================================================

def current_signal_analysis() -> str:
    """基于 Money_Managers 全历史分位的最新一周信号分析"""
    xl = pd.ExcelFile(OUTPUT_EXCEL, engine="openpyxl")
    print("\n=== 当前历史分位分析 ===\n")
    lines = []

    for metal in METALS:
        if metal not in xl.sheet_names:
            continue
        df = xl.parse(metal)
        df["Report_Date"] = pd.to_datetime(df["Report_Date"], errors="coerce")
        df = df.dropna(subset=["Report_Date", "Money_Managers_Pct_OI_Percentile"])

        latest = df.sort_values("Report_Date", ascending=False).iloc[0]
        pct = latest["Money_Managers_Pct_OI_Percentile"]

        count_ge80 = (df["Money_Managers_Pct_OI_Percentile"] >= 80).sum()
        count_50_80 = ((df["Money_Managers_Pct_OI_Percentile"] >= 50) &
                       (df["Money_Managers_Pct_OI_Percentile"] < 80)).sum()
        count_20_50 = ((df["Money_Managers_Pct_OI_Percentile"] >= 20) &
                       (df["Money_Managers_Pct_OI_Percentile"] < 50)).sum()
        count_le20 = (df["Money_Managers_Pct_OI_Percentile"] <= 20).sum()

        print(f"\n{metal} 总历史 {len(df)} 周")
        print(f"  ≥80%: {count_ge80}  |  50-80%: {count_50_80}  |  20-50%: {count_20_50}  |  ≤20%: {count_le20}")

        msg = f"{metal} 当前处于**{percentile_signal(pct)}** ({pct:.1f}%) — {latest['Report_Date'].date()}\n"
        if pct >= 80:
            msg += (f">=80% 出现 {count_ge80} 次\n"
                    f"建议：偏谨慎看空，减持多头或建立空头，严格止损。")
        elif pct <= 20:
            msg += (f"<=20% 出现 {count_le20} 次\n"
                    f"建议：偏向看多，逐步建立多头，设止损防进一步下探。")
        elif pct >= 50:
            msg += (f"50-80% 出现 {count_50_80} 次\n"
                    f"建议：趋势跟随看多，但{'分位较高，设 trailing stop' if pct >= 70 else '风险可控，回调可加仓'}。")
        elif pct > 20:
            msg += (f"20-50% 出现 {count_20_50} 次\n"
                    f"建议：保持中性或轻仓做空，等待更极端信号。")
        else:
            msg += "建议：方向性不强，观望为主。"

        print(msg + "\n")
        lines.append(msg)

    return "\n\n".join(lines)


# =============================================================================
# 邮件发送
# =============================================================================

def send_email(body: str) -> None:
    """发送邮件，附带 Excel 和所有 PNG 图表"""
    msg = MIMEMultipart()
    msg["From"] = EMAIL_FROM
    msg["To"] = EMAIL_TO
    msg["Subject"] = f"CFTC 数据更新 {datetime.now().strftime('%Y-%m-%d')}"
    msg.attach(MIMEText(f"CFTC 数据更新完成。\n\n=== 当前分析 ===\n{body}\n祝交易顺利！", "plain", "utf-8"))

    for file in [OUTPUT_EXCEL, *OUTPUT_FOLDER.glob("*.png")]:
        if not file.exists():
            continue
        with open(file, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={file.name}")
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.send_message(msg)
        print("邮件已发送")
    except Exception as e:
        print(f"邮件发送失败: {e}")


# =============================================================================
# 主入口
# =============================================================================

def main():
    print(f"\n=== 开始 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

    # 准备目录
    for folder in [CFTC_XLS_FOLDER, CFTC_ZIP_FOLDER, OUTPUT_FOLDER]:
        folder.mkdir(parents=True, exist_ok=True)

    # 获取数据
    print("\n[1/5] 获取价格数据...")
    prices = fetch_price_data()

    print("\n[2/5] 下载 CFTC 持仓数据...")
    current_year = datetime.now().year
    (CFTC_ZIP_FOLDER / f"com_disagg_xls_{current_year}.zip").unlink(missing_ok=True)
    download_cftc_data()

    print("\n[3/5] 处理数据、生成 Excel...")
    raw_data = load_cftc_raw_data()
    write_excel(raw_data, prices)

    print("\n[4/5] 生成图表 & 回溯分析...")
    generate_plots()
    backtest_analysis()

    print("\n[5/5] 当前信号 & 发送邮件...")
    analysis = current_signal_analysis()
    send_email(analysis)

    print(f"\n=== 完成 {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")


if __name__ == "__main__":
    main()
