# -*- coding: utf-8 -*-
"""
白银市场分析报告 - 单文件最终版（图表全部嵌入正文，只剩 Excel 附件，无 .bin）
"""

import pandas as pd
import akshare as ak
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication


# =====================================
#          配置区（从环境变量读取）
# =====================================
EMAIL_SENDER    = os.environ.get('EMAIL_FROM', '')
EMAIL_PASSWORD  = os.environ.get('EMAIL_PASSWORD', '')
EMAIL_RECEIVER  = os.environ.get('EMAIL_TO', '')
EMAIL_CC        = None

RSI_PERIOD      = 14

TODAY           = datetime.now().strftime("%Y%m%d")
EXCEL_FILENAME  = f"Silver_Analysis_Report_{TODAY}.xlsx"

# 调试开关
SEND_EMAIL      = True   # 测试时打开


# =====================================
#          RSI 计算
# =====================================
def calculate_rsi(series: pd.Series, period: int = 14) -> pd.Series:
    delta = series.diff()
    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)
    
    avg_gain = gain.rolling(window=period, min_periods=period).mean()
    avg_loss = loss.rolling(window=period, min_periods=period).mean()
    
    rs = avg_gain / avg_loss
    rsi = 100 - (100 / (1 + rs))
    return rsi.round(4)


# =====================================
#          数据获取 + 处理 + 回测 + 保存到 Excel
# =====================================
def get_and_process_data():
    print("正在获取数据...")
    
    df_slv = ak.stock_us_daily(symbol="SLV", adjust="qfq")
    date_col = 'date' if 'date' in df_slv.columns else '日期'
    df_slv[date_col] = pd.to_datetime(df_slv[date_col], errors='coerce')
    df_slv = df_slv.dropna(subset=[date_col]).set_index(date_col).sort_index()
    df_slv = df_slv[['close']].rename(columns={'close': 'Close'})

    df_xag_raw = ak.futures_foreign_hist(symbol="XAG")
    df_xag = df_xag_raw[['date', 'close']].copy()
    df_xag['date'] = pd.to_datetime(df_xag['date'], errors='coerce')
    df_xag = df_xag.dropna(subset=['date']).set_index('date').rename(columns={'close': 'Close'})

    df = pd.merge(
        df_slv.rename(columns={'Close': 'SLV_Close'}),
        df_xag.rename(columns={'Close': 'XAG_Close'}),
        left_index=True,
        right_index=True,
        how='inner'
    )
    
    df['SLV_RSI'] = calculate_rsi(df['SLV_Close'], RSI_PERIOD)
    df['XAG_RSI'] = calculate_rsi(df['XAG_Close'], RSI_PERIOD)
    
    df = df.reset_index(names='Date')
    df = df.sort_values('Date', ascending=False).reset_index(drop=True)
    df[['SLV_Close', 'SLV_RSI', 'XAG_Close', 'XAG_RSI']] = df[['SLV_Close', 'SLV_RSI', 'XAG_Close', 'XAG_RSI']].round(4)

    print(f"数据处理完成，有效行数: {len(df)}")

    # RSI 底部回测
    print("\n=== 开始 RSI(14) 底部回测 ===")
    df_bt = df.set_index('Date').sort_index().copy()
    
    for days in [5, 10, 20, 30]:
        df_bt[f'Return_{days}d'] = df_bt['SLV_Close'].pct_change(days).shift(-days) * 100

    def stats(th):
        sig = df_bt[df_bt['SLV_RSI'] <= th].copy()
        if len(sig) == 0:
            return pd.Series({'信号次数': 0})
        return pd.Series({
            '信号次数': len(sig),
            '5天平均涨幅%': round(sig['Return_5d'].mean(), 2),
            '5天胜率%': round((sig['Return_5d'] > 0).mean() * 100, 2),
            '10天平均涨幅%': round(sig['Return_10d'].mean(), 2),
            '10天胜率%': round((sig['Return_10d'] > 0).mean() * 100, 2),
            '20天平均涨幅%': round(sig['Return_20d'].mean(), 2),
            '20天胜率%': round((sig['Return_20d'] > 0).mean() * 100, 2),
            '30天平均涨幅%': round(sig['Return_30d'].mean(), 2),
            '30天胜率%': round((sig['Return_30d'] > 0).mean() * 100, 2),
        })

    backtest_df = pd.concat([stats(th) for th in [30, 28, 25, 22]], axis=1,
                            keys=[f'RSI <= {th}' for th in [30, 28, 25, 22]])

    # 保存到同一 Excel 的不同 sheet
    with pd.ExcelWriter(EXCEL_FILENAME, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Price_RSI_Data', index=False)
        backtest_df.to_excel(writer, sheet_name='RSI_Bottom_Backtest')
    
    print(f"所有数据已保存到: {EXCEL_FILENAME}")

    return df, backtest_df


# =====================================
#          生成图表
# =====================================
def create_multiple_charts(df: pd.DataFrame):
    print("正在生成多个时间范围的图表...")
    
    df['Date'] = pd.to_datetime(df['Date'])
    max_date = df['Date'].max()
    
    periods = [
        {"name": "Full History",     "days": None},
        {"name": "Last 5 Years",     "days": 5*365 + 1},
        {"name": "Last 2 Years",     "days": 2*365 + 1},
        {"name": "Last 1 Year",      "days": 365 + 1},
        {"name": "Last 6 Months",    "days": 180 + 1},
    ]
    
    generated_files = []
    
    for p in periods:
        if p["days"] is None:
            df_view = df.copy()
        else:
            cutoff = max_date - timedelta(days=p["days"])
            df_view = df[df['Date'] >= cutoff].copy()
        
        if len(df_view) < 30:
            print(f"跳过 {p['name']}：数据点太少 ({len(df_view)})")
            continue
        
        filename = f"Silver_Price_RSI_{TODAY}_{p['name'].replace(' ','_')}.png"
        
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 9), sharex=True,
                                      gridspec_kw={'height_ratios': [3, 1]})
        
        ax1.plot(df_view['Date'], df_view['SLV_Close'], label='SLV Close', color='blue', linewidth=1.2)
        ax1.plot(df_view['Date'], df_view['XAG_Close'], label='XAG Close', color='silver', ls='--', linewidth=1.2)
        ax1.legend(loc='upper left')
        ax1.grid(True, alpha=0.3)
        ax1.set_title(f"SLV & XAG Price - {p['name']}")
        
        ax2.plot(df_view['Date'], df_view['SLV_RSI'], label='SLV RSI(14)', color='blue', linewidth=1.2)
        ax2.plot(df_view['Date'], df_view['XAG_RSI'], label='XAG RSI(14)', color='silver', ls='--', linewidth=1.2)
        
        ax2.fill_between(df_view['Date'], 70, 100, where=(df_view['SLV_RSI'] > 70) | (df_view['XAG_RSI'] > 70),
                         color='red', alpha=0.12, interpolate=True)
        ax2.fill_between(df_view['Date'], 0, 30, where=(df_view['SLV_RSI'] < 30) | (df_view['XAG_RSI'] < 30),
                         color='green', alpha=0.12, interpolate=True)
        
        ax2.axhline(70, color='red', ls='--', alpha=0.6, label='Overbought 70')
        ax2.axhline(30, color='green', ls='--', alpha=0.6, label='Oversold 30')
        ax2.axhline(50, color='gray', ls='-', alpha=0.4)
        ax2.set_ylabel('RSI')
        ax2.legend(loc='upper left')
        ax2.grid(True, alpha=0.3)
        ax2.set_ylim(0, 100)
        
        ax2.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        fig.autofmt_xdate(rotation=45)
        plt.suptitle(f'Silver Market Analysis - {p["name"]} ({TODAY})', fontsize=16)
        plt.tight_layout()
        
        plt.savefig(filename, dpi=150, bbox_inches='tight')
        plt.close()
        
        generated_files.append(filename)
        print(f"已生成：{filename} ({len(df_view)} 条数据)")
    
    return generated_files


# =====================================
#          发送邮件（图表全部嵌入正文，只剩 Excel 附件）
# =====================================
def send_email(df: pd.DataFrame, backtest_df: pd.DataFrame, chart_files: list):
    if not SEND_EMAIL:
        print("邮件发送已关闭（调试模式）")
        return
    
    print("正在准备并发送邮件...")
    
    msg = MIMEMultipart('related')
    msg['From'] = EMAIL_SENDER
    msg['To'] = EMAIL_RECEIVER
    if EMAIL_CC:
        msg['Cc'] = EMAIL_CC
    msg['Subject'] = f"白银市场分析报告 {TODAY} - RSI底部回测 + 多时间范围图表"
    
    backtest_html = backtest_df.to_html(classes="table", border=1, float_format="%.2f")
    
    latest = df.iloc[0]
    latest_date = latest['Date'].date()
    latest_rsi = latest['SLV_RSI']
    
    suggestion = ""
    if latest_rsi <= 22:
        suggestion = "当前 RSI 极低（<=22），历史 10 天胜率约 60-61%，平均涨幅 1.6%以上 → 极强底部信号，可考虑轻仓试多"
    elif latest_rsi <= 25:
        suggestion = "当前 RSI 较低（<=25），历史 10 天胜率约 60%，平均涨幅 1.45% → 强底部信号，可关注反弹机会"
    elif latest_rsi <= 30:
        suggestion = "当前 RSI 中等超卖（<=30），历史 10 天胜率约 59-60%，平均涨幅 1.2-1.3% → 中等信号，建议观望或等待金叉"
    else:
        suggestion = "当前 RSI 未达超卖区（>30），无明确底部信号 → 建议继续观望，避免盲目抄底"

    # 构建正文 HTML + 内联所有图表
    img_tags = ""
    cid_counter = 1
    for chart_file in chart_files:
        cid = f"chart_{cid_counter}"
        with open(chart_file, 'rb') as f:
            img = MIMEImage(f.read())
            img.add_header('Content-ID', f'<{cid}>')
            msg.attach(img)
        img_tags += f'<p><strong>{os.path.basename(chart_file)}</strong><br><img src="cid:{cid}" style="max-width:100%;"></p>'
        cid_counter += 1

    html = f"""
    <html>
    <body>
    <h2>白银市场技术分析报告 ({TODAY})</h2>
    
    <p>最新数据 ({latest_date}):</p>
    <ul>
        <li>SLV 收盘价: {latest['SLV_Close']:.4f} | RSI(14): {latest_rsi:.2f}</li>
        <li>XAG 收盘价: {latest['XAG_Close']:.4f} | RSI(14): {latest['XAG_RSI']:.2f}</li>
    </ul>
    
    <p><strong>当前市场建议:</strong> {suggestion}</p>
    
    <h3>RSI(14) 底部回测统计表（历史数据）</h3>
    {backtest_df.to_html(classes="table", border=1, float_format="%.2f")}
    
    <p>详细数据见附件 Excel（多 sheet）。</p>
    
    <h3>多时间范围图表（全部嵌入正文）</h3>
    {img_tags}
    
    <p>祝交易顺利！</p>
    </body>
    </html>
    """
    msg.attach(MIMEText(html, 'html', 'utf-8'))
    
    # 只添加 Excel 作为附件
    with open(EXCEL_FILENAME, 'rb') as f:
        part = MIMEApplication(f.read(), _subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(EXCEL_FILENAME))
        part.add_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        msg.attach(part)
    
    try:
        server = smtplib.SMTP_SSL('smtp.qq.com', 465)
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print("邮件发送成功！（附件只剩 Excel，图表全部嵌入正文，无 .bin）")
    except Exception as e:
        print(f"邮件发送失败：{e}")


# =====================================
#          主程序
# =====================================
def main():
    try:
        df, backtest_df = get_and_process_data()
        
        chart_files = create_multiple_charts(df)
        
        if SEND_EMAIL:
            send_email(df, backtest_df, chart_files)
        else:
            print("调试模式：邮件已关闭。如需发送请把 SEND_EMAIL 改为 True")
        
        print("\n全部完成！")
        print(f"完整报告已保存到: {EXCEL_FILENAME} （包含两个 sheet）")
        
    except Exception as e:
        print(f"\n程序执行出错：{e}")


if __name__ == "__main__":
    main()
