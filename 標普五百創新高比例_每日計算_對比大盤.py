"""
標普五百（S&P 500）創新高價股數量比 — 每日計算腳本（含大盤指數比對版）
========================================================================
功能：
1. 從 Wikipedia 取得 S&P 500 所有成分股代碼
2. 透過 yfinance 下載過去一年的歷史資料
3. 計算「今日最高價 = 過去一年最高價」的股票佔比
4. 同時取得 S&P 500 指數（^GSPC）每日收盤價
5. 把每天的計算結果（含大盤指數）追加存入 CSV 檔案
6. 用最新的 CSV 資料產生「雙 Y 軸」互動式 HTML 圖表
   - 左 Y 軸：創新高比例 (%)
   - 右 Y 軸：S&P 500 指數收盤點數
"""

import yfinance as yf
import pandas as pd
import requests
import warnings
import json
import io
import os
from datetime import datetime

# ===== 關閉不必要的警告訊息 =====
warnings.filterwarnings('ignore')

# ===== 設定檔案路徑（跟這支腳本放在同一個資料夾） =====
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(SCRIPT_DIR, "標普五百創新高比例_歷史資料.csv")
HTML_PATH = os.path.join(SCRIPT_DIR, "標普五百創新高比例_圖表.html")


def get_sp500_tickers():
    """
    從 Wikipedia 取得 S&P 500 所有成分股代碼。
    資料來源：https://en.wikipedia.org/wiki/List_of_S%26P_500_companies
    回傳格式範例：['AAPL', 'MSFT', 'GOOGL', ...]
    """
    print("📡 正在從 Wikipedia 獲取 S&P 500 成分股清單...")

    url = "https://en.wikipedia.org/wiki/List_of_S%26P_500_companies"
    # 加上 User-Agent 避免被 Wikipedia 擋 403
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                      'AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/120.0.0.0 Safari/537.36'
    }

    try:
        res = requests.get(url, headers=headers, timeout=15)
        res.raise_for_status()

        # 用 io.StringIO 包住 HTML 字串，避免 pandas 誤判為檔案路徑
        tables = pd.read_html(io.StringIO(res.text))
        df = tables[0]  # 第一個表格就是成分股清單

        # Symbol 欄位存放股票代碼，有些代碼含有 '.' 需換成 '-'（yfinance 格式）
        # 例如 BRK.B → BRK-B
        tickers = df['Symbol'].str.replace('.', '-', regex=False).tolist()
        print(f"✅ 成功獲取 {len(tickers)} 檔 S&P 500 成分股。")
        return tickers

    except Exception as e:
        print(f"❌ 獲取 S&P 500 清單失敗: {e}")
        return []


def get_sp500_close():
    """
    從 Yahoo Finance 下載 S&P 500 指數（代碼 ^GSPC）的最新收盤價。

    什麼是 ^GSPC？
    - 就是「標準普爾五百指數」在 Yahoo Finance 上的代碼
    - 追蹤美國 500 家大型上市公司的加權平均表現
    - 例如新聞說「今天 S&P 500 收在 5,200 點」，指的就是這個

    回傳值：最新一個交易日的收盤點數（float），若下載失敗則回傳 None。
    """
    try:
        print("📡 正在下載 S&P 500 指數（^GSPC）資料...")

        # period="5d" 只抓最近 5 天，確保至少能拿到最新一個交易日的資料
        index_data = yf.download("^GSPC", period="5d", interval="1d", auto_adjust=True)

        if index_data.empty:
            print("⚠ 無法取得 S&P 500 指數資料")
            return None

        latest_close_data = index_data['Close'].iloc[-1]

        # 判斷如果是 Series（新版 yfinance），就取出第一個數值；否則直接轉 float
        if isinstance(latest_close_data, pd.Series):
            latest_close = float(latest_close_data.iloc[0])
        else:
            latest_close = float(latest_close_data)

        latest_close = round(latest_close, 2)
        print(f"✅ S&P 500 指數收盤價: {latest_close}")
        return latest_close

    except Exception as e:
        print(f"⚠ 下載 S&P 500 指數時發生錯誤: {e}")
        return None


def calculate_new_high_ratio():
    """
    核心計算函式：
    1. 取得所有 S&P 500 成分股代碼
    2. 下載過去一年的歷史最高價資料
    3. 計算「最新一天的最高價 = 過去一年最高價」的股票數量佔比
    4. 同時取得 S&P 500 指數收盤價
    5. 回傳 (日期, 有效股票數, 創新高數, 百分比, S&P500指數) 的 tuple
    """
    tickers = get_sp500_tickers()

    if not tickers:
        print("❌ 沒有獲取到股票代碼，程式結束。")
        return None

    print(f"⏳ 正在下載過去一年的歷史資料（這可能需要幾分鐘）...")

    # yf.download 一次批次下載所有股票，threads=True 啟用多線程加速
    data = yf.download(tickers, period="1y", interval="1d", auto_adjust=True, threads=True)

    # ===== 同步取得 S&P 500 指數收盤價 =====
    sp500_close = get_sp500_close()

    new_high_count = 0      # 創新高的股票數量
    valid_stock_count = 0   # 有足夠歷史資料的股票數量

    print("🔍 正在計算創新高比例...")

    for ticker in tickers:
        try:
            # 取出這檔股票每天的「最高價」，移除缺失值
            high_prices = data['High'][ticker].dropna()

            # 若交易天數不到 200 天，不納入計算
            if len(high_prices) < 200:
                continue

            valid_stock_count += 1

            latest_high = high_prices.iloc[-1]   # 最新一天的最高價
            year_high = high_prices.max()          # 過去一年的最高價

            # 最新最高價 >= 年度最高價，代表今天創了新高
            if latest_high >= year_high:
                new_high_count += 1

        except KeyError:
            continue

    if valid_stock_count == 0:
        print("❌ 沒有足夠的有效股票資料可供計算。")
        return None

    ratio = round((new_high_count / valid_stock_count) * 100, 2)
    today = datetime.now().strftime("%Y-%m-%d")

    print("-" * 40)
    print(f"📅 日期: {today}")
    print(f"📊 有效樣本數: {valid_stock_count} 檔")
    print(f"🔺 創新高股票數: {new_high_count} 檔")
    print(f"📈 S&P 500 創新高價股數量比: {ratio}%")
    if sp500_close is not None:
        print(f"📉 S&P 500 指數收盤: {sp500_close}")

    return today, valid_stock_count, new_high_count, ratio, sp500_close


def save_to_csv(date, valid_count, new_high_count, ratio, sp500_close):
    """
    把今天的計算結果追加到 CSV 檔案。
    如果 CSV 檔不存在，就自動建立並寫入表頭。
    如果今天已經有紀錄，就覆蓋（避免重複執行產生重複資料）。
    """
    new_row = pd.DataFrame([{
        "日期": date,
        "有效股票數": valid_count,
        "創新高股票數": new_high_count,
        "創新高比例(%)": ratio,
        "SP500指數": sp500_close
    }])

    if os.path.exists(CSV_PATH):
        df = pd.read_csv(CSV_PATH)

        # 相容舊版資料：若無 SP500指數 欄位，自動補上空值
        if "SP500指數" not in df.columns:
            df["SP500指數"] = None

        # 移除今天的舊紀錄（避免重複）
        df = df[df["日期"] != date]
        df = pd.concat([df, new_row], ignore_index=True)
    else:
        df = new_row

    df = df.sort_values("日期").reset_index(drop=True)
    df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
    print(f"💾 資料已儲存至: {CSV_PATH}")

    return df


def generate_html_chart(df):
    """
    根據 CSV 中的歷史資料，產生一個「雙 Y 軸」互動式 HTML 折線圖。
    - 左 Y 軸（金色線）：創新高比例 (%)
    - 右 Y 軸（青色線）：S&P 500 指數收盤點數
    """
    dates = df["日期"].tolist()
    ratios = df["創新高比例(%)"].tolist()
    new_highs = df["創新高股票數"].tolist()
    valid_counts = df["有效股票數"].tolist()

    if "SP500指數" in df.columns:
        sp500_values = df["SP500指數"].astype(object).where(pd.notnull(df["SP500指數"]), None).tolist()
    else:
        sp500_values = [None] * len(dates)

    html_content = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>S&P 500 創新高價股數量比 vs 大盤指數</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        body {{
            font-family: 'Segoe UI', 'Microsoft JhengHei', sans-serif;
            background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
            min-height: 100vh;
            padding: 20px;
            color: #e0e0e0;
        }}
        .container {{
            max-width: 1100px;
            margin: 0 auto;
        }}
        h1 {{
            text-align: center;
            font-size: 28px;
            margin-bottom: 8px;
            background: linear-gradient(90deg, #f7971e, #ffd200);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}
        .subtitle {{
            text-align: center;
            color: #888;
            font-size: 14px;
            margin-bottom: 24px;
        }}
        .stats-row {{
            display: flex;
            gap: 16px;
            margin-bottom: 24px;
            flex-wrap: wrap;
            justify-content: center;
        }}
        .stat-card {{
            background: rgba(255,255,255,0.06);
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 12px;
            padding: 16px 24px;
            text-align: center;
            min-width: 160px;
            flex: 1;
            max-width: 220px;
        }}
        .stat-card .label {{
            font-size: 13px;
            color: #999;
            margin-bottom: 6px;
        }}
        .stat-card .value {{
            font-size: 26px;
            font-weight: bold;
        }}
        .stat-card .value.highlight {{ color: #ffd200; }}
        .stat-card .value.green {{ color: #4caf50; }}
        .stat-card .value.blue {{ color: #42a5f5; }}
        .stat-card .value.cyan {{ color: #26c6da; }}
        .chart-wrapper {{
            background: rgba(255,255,255,0.04);
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 16px;
            padding: 24px;
            margin-bottom: 20px;
        }}
        canvas {{
            width: 100% !important;
        }}
        .table-wrapper {{
            background: rgba(255,255,255,0.04);
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 16px;
            padding: 20px;
            overflow-x: auto;
        }}
        .table-wrapper h2 {{
            font-size: 18px;
            margin-bottom: 12px;
            color: #ccc;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }}
        th {{
            background: rgba(255,255,255,0.08);
            padding: 10px 12px;
            text-align: center;
            color: #aaa;
            font-weight: 600;
        }}
        td {{
            padding: 10px 12px;
            text-align: center;
            border-bottom: 1px solid rgba(255,255,255,0.06);
        }}
        tr:hover td {{
            background: rgba(255,255,255,0.04);
        }}
        .footer {{
            text-align: center;
            color: #555;
            font-size: 12px;
            margin-top: 16px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>S&P 500 創新高價股數量比 vs 大盤指數</h1>
        <p class="subtitle">每日追蹤：S&P 500 成分股創 52 週新高佔比（金色）與 S&P 500 指數走勢（青色）</p>

        <div class="stats-row">
            <div class="stat-card">
                <div class="label">S&P 500 指數</div>
                <div class="value cyan" id="latest-sp500">—</div>
            </div>
            <div class="stat-card">
                <div class="label">創新高比例</div>
                <div class="value highlight" id="latest-ratio">—</div>
            </div>
            <div class="stat-card">
                <div class="label">創新高股票數</div>
                <div class="value green" id="latest-count">—</div>
            </div>
            <div class="stat-card">
                <div class="label">有效樣本數</div>
                <div class="value blue" id="latest-valid">—</div>
            </div>
            <div class="stat-card">
                <div class="label">資料天數</div>
                <div class="value" id="total-days">—</div>
            </div>
        </div>

        <div class="chart-wrapper">
            <canvas id="mainChart" height="420"></canvas>
        </div>

        <div class="table-wrapper">
            <h2>歷史資料</h2>
            <table id="dataTable">
                <thead>
                    <tr>
                        <th>日期</th>
                        <th>S&P 500 指數</th>
                        <th>有效股票數</th>
                        <th>創新高股票數</th>
                        <th>創新高比例 (%)</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
        </div>

        <p class="footer">資料來源：Yahoo Finance ｜ S&P 500 成分股清單：Wikipedia ｜ 最後更新：<span id="last-update"></span></p>
    </div>

    <script>
        const dates = {json.dumps(dates)};
        const ratios = {json.dumps(ratios)};
        const newHighs = {json.dumps(new_highs)};
        const validCounts = {json.dumps(valid_counts)};
        const sp500Values = {json.dumps(sp500_values)};

        const lastIdx = dates.length - 1;
        document.getElementById('latest-ratio').textContent = ratios[lastIdx] + '%';
        document.getElementById('latest-count').textContent = newHighs[lastIdx] + ' 檔';
        document.getElementById('latest-valid').textContent = validCounts[lastIdx] + ' 檔';
        document.getElementById('total-days').textContent = dates.length + ' 天';
        document.getElementById('last-update').textContent = dates[lastIdx];

        const latestSP500 = sp500Values[lastIdx];
        document.getElementById('latest-sp500').textContent =
            latestSP500 !== null ? latestSP500.toLocaleString() : '—';

        const ctx = document.getElementById('mainChart').getContext('2d');
        new Chart(ctx, {{
            type: 'line',
            data: {{
                labels: dates,
                datasets: [
                    {{
                        label: '創新高比例 (%)',
                        data: ratios,
                        borderColor: '#ffd200',
                        backgroundColor: 'rgba(255,210,0,0.08)',
                        borderWidth: 2,
                        pointRadius: 3,
                        pointHoverRadius: 6,
                        pointBackgroundColor: '#ffd200',
                        fill: true,
                        tension: 0.3,
                        yAxisID: 'y'
                    }},
                    {{
                        label: 'S&P 500 指數',
                        data: sp500Values,
                        borderColor: '#26c6da',
                        backgroundColor: 'rgba(38,198,218,0.08)',
                        borderWidth: 2,
                        pointRadius: 3,
                        pointHoverRadius: 6,
                        pointBackgroundColor: '#26c6da',
                        fill: true,
                        tension: 0.3,
                        yAxisID: 'y1',
                        spanGaps: true
                    }}
                ]
            }},
            options: {{
                responsive: true,
                maintainAspectRatio: false,
                interaction: {{
                    mode: 'index',
                    intersect: false
                }},
                plugins: {{
                    legend: {{
                        labels: {{ color: '#ccc', font: {{ size: 14 }} }}
                    }},
                    tooltip: {{
                        callbacks: {{
                            afterLabel: function(context) {{
                                const i = context.dataIndex;
                                if (context.datasetIndex === 0) {{
                                    return '創新高: ' + newHighs[i] + ' / ' + validCounts[i] + ' 檔';
                                }}
                                return '';
                            }}
                        }}
                    }}
                }},
                scales: {{
                    x: {{
                        ticks: {{ color: '#888', maxRotation: 45 }},
                        grid: {{ color: 'rgba(255,255,255,0.05)' }}
                    }},
                    y: {{
                        type: 'linear',
                        position: 'left',
                        ticks: {{
                            color: '#ffd200',
                            callback: function(v) {{ return v + '%'; }}
                        }},
                        grid: {{ color: 'rgba(255,255,255,0.05)' }},
                        beginAtZero: true,
                        title: {{
                            display: true,
                            text: '創新高比例 (%)',
                            color: '#ffd200',
                            font: {{ size: 13 }}
                        }}
                    }},
                    y1: {{
                        type: 'linear',
                        position: 'right',
                        ticks: {{
                            color: '#26c6da',
                            callback: function(v) {{ return v.toLocaleString(); }}
                        }},
                        grid: {{
                            drawOnChartArea: false
                        }},
                        title: {{
                            display: true,
                            text: 'S&P 500 指數',
                            color: '#26c6da',
                            font: {{ size: 13 }}
                        }}
                    }}
                }}
            }}
        }});

        const tbody = document.querySelector('#dataTable tbody');
        for (let i = dates.length - 1; i >= 0; i--) {{
            const tr = document.createElement('tr');
            const sp500Display = sp500Values[i] !== null
                ? sp500Values[i].toLocaleString()
                : '—';
            tr.innerHTML = `<td>${{dates[i]}}</td><td>${{sp500Display}}</td><td>${{validCounts[i]}}</td><td>${{newHighs[i]}}</td><td>${{ratios[i]}}%</td>`;
            tbody.appendChild(tr);
        }}
    </script>
</body>
</html>"""

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f"📊 互動式圖表已產生: {HTML_PATH}")


# ===== 主程式 =====
if __name__ == "__main__":
    print("=" * 55)
    print("  S&P 500 創新高價股數量比 — 每日計算（含大盤指數）")
    print("=" * 55)

    result = calculate_new_high_ratio()

    if result:
        date, valid_count, new_high_count, ratio, sp500_close = result
        df = save_to_csv(date, valid_count, new_high_count, ratio, sp500_close)
        generate_html_chart(df)
        print("\n✅ 全部完成！")
    else:
        print("\n❌ 計算失敗，請檢查網路連線或稍後再試。")
