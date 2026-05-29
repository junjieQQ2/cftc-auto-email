"""图表生成"""

import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.dates import DateFormatter

from .config import CATEGORIES, OUTPUT_EXCEL, OUTPUT_FOLDER, TIME_PERIODS


# ====================== 图表生成 ======================
def draw_percentile_chart(metal, df, title_suffix, filename_suffix, col_suffix="Pct_OI_Percentile"):
    fig, axes = plt.subplots(6, 1, figsize=(15, 26), sharex=True, gridspec_kw={'height_ratios': [1]*6})
    fig.suptitle(f"{metal} Percentiles & Spot Price {title_suffix}", fontsize=16, fontweight='bold')
    ax_price = axes[0]
    if 'close' in df.columns and not df['close'].isna().all():
        ax_price.plot(df['Report_Date'], df['close'], color='black', lw=1.5)
        ax_price.set_ylabel('Spot Price')
        ax_price.grid(True, alpha=0.3)
    for i, cat in enumerate(CATEGORIES):
        ax = axes[i + 1]
        col = f"{cat}_{col_suffix}"
        if col not in df.columns or df[col].isna().all():
            ax.text(0.5, 50, "No data", ha='center', va='center')
            continue
        colors = ['darkred' if v >= 80 else 'orange' if 50 <= v < 80 else 'blue' if 20 <= v < 50 else 'darkgreen' for v in df[col]]
        ax.scatter(df['Report_Date'], df[col], c=colors, s=40, alpha=0.8)
        ax.axhline(80, color="darkred", ls=":")
        ax.axhline(50, color="orange", ls=":")
        ax.axhline(20, color="blue", ls=":")
        ax.axhline(50, color="gray", ls="--")
        ax.set_ylabel(f"{cat} (%)")
        ax.grid(True, alpha=0.3)
        ax.set_ylim(0, 100)
    legend_elements = [mpatches.Patch(color='darkred', label='≥80%'), mpatches.Patch(color='orange', label='50–79%'),
                       mpatches.Patch(color='blue', label='20–49%'), mpatches.Patch(color='darkgreen', label='≤20%')]
    fig.legend(handles=legend_elements, loc='lower center', bbox_to_anchor=(0.5, -0.02), ncol=1, fontsize=9)
    axes[-1].xaxis.set_major_formatter(DateFormatter('%Y-%m'))
    plt.xticks(rotation=45)
    fig.tight_layout(rect=[0, 0.05, 1, 0.92])
    plt.savefig(OUTPUT_FOLDER / f"{metal}_{filename_suffix}.png", dpi=180, bbox_inches='tight')
    plt.close(fig)


def generate_plots():
    if not OUTPUT_EXCEL.exists():
        print("Excel 文件不存在，跳过图表生成")
        return
    xl = pd.ExcelFile(OUTPUT_EXCEL, engine='openpyxl')
    for metal in ["Gold", "Silver", "Platinum"]:
        if metal not in xl.sheet_names:
            continue
        df = xl.parse(metal)
        df['Report_Date'] = pd.to_datetime(df['Report_Date'], errors='coerce')
        df = df.dropna(subset=['Report_Date']).sort_values('Report_Date')
        if df.empty:
            continue
        latest_date = df['Report_Date'].max()
        draw_percentile_chart(metal, df, "(Full History)", "Historical_Percentiles")
        sheet_name = f"{metal}_2014_2020"
        if sheet_name in xl.sheet_names:
            df_1420 = xl.parse(sheet_name)
            df_1420['Report_Date'] = pd.to_datetime(df_1420['Report_Date'], errors='coerce')
            df_1420 = df_1420.dropna(subset=['Report_Date']).sort_values('Report_Date')
            df_1420 = df_1420[(df_1420['Report_Date'] >= '2014-01-01') & (df_1420['Report_Date'] <= '2020-12-31')]
            if not df_1420.empty:
                draw_percentile_chart(metal, df_1420, "(2014-2020)", "2014_2020_Historical_Percentiles")
        for period in ["1Y", "3Y", "5Y", "10Y"]:
            period_start = latest_date - TIME_PERIODS[period]
            period_df = df[df['Report_Date'] >= period_start].copy()
            if len(period_df) < 10:
                continue
            col_suffix = f"{period}_Pct_OI_Percentile"
            if all(f"{cat}_{col_suffix}" not in period_df.columns or period_df[f"{cat}_{col_suffix}"].isna().all() for cat in CATEGORIES):
                continue
            draw_percentile_chart(metal, period_df, f"({period})", f"{period}_Percentiles", col_suffix=col_suffix)
    print("图表生成完成")
