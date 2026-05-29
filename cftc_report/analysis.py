"""回溯分析与当前信号分析"""

import numpy as np
import pandas as pd

from .config import BACKTEST_SHEET, METAL_FILTERS, OUTPUT_EXCEL


# ====================== 回溯分析 ======================
def backtest_analysis():
    xl = pd.ExcelFile(OUTPUT_EXCEL, engine='openpyxl')
    results = []
    for metal in METAL_FILTERS:
        if metal not in xl.sheet_names: continue
        df = xl.parse(metal)
        df['Report_Date'] = pd.to_datetime(df['Report_Date'])
        df = df.sort_values('Report_Date').set_index('Report_Date')
        pct_col = 'Money_Managers_Pct_OI_Percentile'
        if pct_col not in df.columns or 'close' not in df.columns: continue
        df['Overbought'] = (df[pct_col] >= 80).astype(int)
        df['Oversold'] = (df[pct_col] <= 20).astype(int)
        df['Overbought_Signal'] = df['Overbought'].rolling(2).sum() >= 2
        df['Oversold_Signal'] = df['Oversold'].rolling(2).sum() >= 2
        for h in [1, 2, 4]:
            df[f'Return_{h}w'] = df['close'].pct_change(h).shift(-h) * 100
        for sig, name in [('Overbought_Signal', '做多拥挤 (>=80%)'), ('Oversold_Signal', '做空拥挤 (<=20%)')]:
            instances = df[df[sig]]
            if instances.empty: continue
            stats = {'Metal': metal, 'Signal_Type': name, 'Count': len(instances)}
            for h in [1, 2, 4]:
                rets = instances[f'Return_{h}w'].dropna()
                stats[f'Avg_Return_{h}w'] = rets.mean() if not rets.empty else np.nan
                stats[f'Win_Rate_{h}w'] = (rets > 0).mean() * 100 if not rets.empty else np.nan
            results.append(stats)
    if results:
        result_df = pd.DataFrame(results)
        with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            result_df.to_excel(writer, sheet_name=BACKTEST_SHEET, index=False)
        print("回溯分析已写入")


# ====================== 当前历史分位分析（优化版） ======================
def current_signal_analysis():
    xl = pd.ExcelFile(OUTPUT_EXCEL, engine='openpyxl')
    print("\n=== 当前历史分位分析（最新报告周） ===\n")
    analysis_lines = []

    for metal in ["Gold", "Silver", "Platinum"]:
        if metal not in xl.sheet_names:
            continue

        full_df = xl.parse(metal)

        print(f"→ 正在读取 {metal} sheet，数据量: {len(full_df)} 行")   # 调试信息

        full_df['Report_Date'] = pd.to_datetime(full_df['Report_Date'], errors='coerce')
        full_df = full_df.dropna(subset=['Report_Date', 'Money_Managers_Pct_OI_Percentile'])

        latest_row = full_df.sort_values('Report_Date', ascending=False).head(1)
        pct = latest_row['Money_Managers_Pct_OI_Percentile'].iloc[0]
        latest_date = latest_row['Report_Date'].iloc[0]

        # 各区间统计
        count_ge80 = len(full_df[full_df['Money_Managers_Pct_OI_Percentile'] >= 80])
        count_le20 = len(full_df[full_df['Money_Managers_Pct_OI_Percentile'] <= 20])
        count_50_80 = len(full_df[(full_df['Money_Managers_Pct_OI_Percentile'] >= 50) & 
                                  (full_df['Money_Managers_Pct_OI_Percentile'] < 80)])
        count_20_50 = len(full_df[(full_df['Money_Managers_Pct_OI_Percentile'] > 20) & 
                                  (full_df['Money_Managers_Pct_OI_Percentile'] < 50)])
        # 出现次数统计
        total_weeks = len(full_df)
        print(f"\n{metal} 总历史周数: {total_weeks} 周")
        print(f"  >=80% : {count_ge80} 次")
        print(f"  50-80%: {count_50_80} 次")
        print(f"  20-50%: {count_20_50} 次")
        print(f"  <=20% : {count_le20} 次")

        if pct >= 80:
            msg = (f"{metal} 当前处于**极度拥挤多头** ({pct:.1f}%) - 日期: {latest_date.date()}\n"
                   f"历史>=80%出现次数: {count_ge80} 次\n"
                   f"专业建议：市场情绪极端乐观，历史上此类信号后回调或反转概率较高。\n"
                   f"**推荐操作**：偏谨慎看空，建议减持多头或建立空头头寸，设置严格止损。")

        elif pct <= 20:
            msg = (f"{metal} 当前处于**极度拥挤空头** ({pct:.1f}%) - 日期: {latest_date.date()}\n"
                   f"历史<=20%出现次数: {count_le20} 次\n"
                   f"专业建议：市场情绪极端悲观，历史上此类信号后反弹或反转概率较高。\n"
                   f"**推荐操作**：偏向看多，可逐步建立或加仓多头头寸，注意止损防范进一步下探。")

        elif 50 <= pct < 80:
            msg = (f"{metal} 当前处于**强势多头区间** ({pct:.1f}%) - 日期: {latest_date.date()}\n"
                   f"历史50-80%出现次数: {count_50_80} 次 | 20-50%出现次数: {count_20_50} 次\n"
                   f"专业建议：市场情绪偏强，趋势可能延续，但需关注分位高度。")

            if pct >= 70:
                msg += "\n**推荐操作**：趋势跟随看多，但分位已较高，注意过热风险，设置 trailing stop 保护利润，避免追高。"
            else:
                msg += "\n**推荐操作**：可趋势跟随看多，风险相对可控，回调时可考虑加仓。"

        elif 20 < pct < 50:
            msg = (f"{metal} 当前处于**弱势空头区间** ({pct:.1f}%) - 日期: {latest_date.date()}\n"
                   f"历史20-50%出现次数: {count_20_50} 次\n"
                   f"专业建议：市场情绪偏弱，短期仍有下行压力，但反弹机会也在增加。\n"
                   f"**推荐操作**：保持中性或轻仓做空，等待更极端信号或技术支撑位再加大仓位。")

        else:
            msg = (f"{metal} 当前处于**中性区间** ({pct:.1f}%) - 日期: {latest_date.date()}\n"
                   f"历史中性区间出现次数: {len(full_df) - count_ge80 - count_le20 - count_50_80 - count_20_50} 次\n"
                   f"专业建议：市场方向性不强，建议观望为主。")

        print(msg + "\n")
        analysis_lines.append(msg)

    return "\n\n".join(analysis_lines)
