import yfinance as yf
import pandas as pd
import numpy as np
import os
import json
import gspread
from google.oauth2.service_account import Credentials
from concurrent.futures import ThreadPoolExecutor
import datetime

# --- 設定: TOPIX-17業種 ETFリスト ---
SECTOR_ETFS = {
    "1617": "食品",
    "1618": "エネルギー・資源",
    "1619": "建設・資材",
    "1620": "素材・化学",
    "1621": "医薬品",
    "1622": "自動車・輸送機",
    "1623": "鉄鋼・非鉄",
    "1624": "機械",
    "1625": "電機・精密",
    "1626": "情報通信・サービス",
    "1627": "電力・ガス",
    "1628": "運輸・物流",
    "1629": "商社・卸売",
    "1630": "小売",
    "1631": "銀行",
    "1632": "金融(除く銀行)",
    "1633": "不動産"
}

def get_sector_data(code, name):
    """
    1つのETFのデータを取得し、トレンド指標を計算する
    """
    ticker = f"{code}.T"
    try:
        # 過去1年分のデータを取得
        stock = yf.Ticker(ticker)
        hist = stock.history(period="1y")
        
        if hist.empty:
            return None

        # 最新の終値
        current_price = hist['Close'].iloc[-1]
        
        # --- テクニカル指標の計算 ---
        
        # 1. 移動平均乖離率 (移動平均からどれくらい離れているか%)
        # 短期(5日), 中期(25日), 長期(75日)
        ma5 = hist['Close'].rolling(window=5).mean().iloc[-1]
        ma25 = hist['Close'].rolling(window=25).mean().iloc[-1]
        ma75 = hist['Close'].rolling(window=75).mean().iloc[-1]
        
        diff_short = ((current_price - ma5) / ma5) * 100
        diff_mid = ((current_price - ma25) / ma25) * 100
        diff_long = ((current_price - ma75) / ma75) * 100

        # 2. RSI (14日) - 買われすぎ/売られすぎ (0-100)
        delta = hist['Close'].diff()
        gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
        loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
        rs = gain / loss
        rsi = 100 - (100 / (1 + rs.iloc[-1]))

        # 3. 前日比
        prev_close = hist['Close'].iloc[-2]
        change_pct = ((current_price - prev_close) / prev_close) * 100

        return [
            code,           # A列: コード
            name,           # B列: セクター名
            round(current_price, 1), # C列: 現在値
            round(change_pct, 2),    # D列: 前日比(%)
            round(diff_short, 2),    # E列: 短期トレンド(5日線乖離%)
            round(diff_mid, 2),      # F列: 中期トレンド(25日線乖離%)
            round(diff_long, 2),     # G列: 長期トレンド(75日線乖離%)
            round(rsi, 1),           # H列: RSI(過熱感)
            datetime.datetime.now().strftime('%Y-%m-%d %H:%M') # I列: 更新日時
        ]

    except Exception as e:
        print(f"Error {code}: {e}")
        return None

def update_sheet():
    # --- Google Sheets認証 ---
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # GitHub Secretsまたはローカルファイルから認証情報を読み込む
    creds_json = None
    if 'GCP_SERVICE_ACCOUNT' in os.environ:
        creds_json = json.loads(os.environ['GCP_SERVICE_ACCOUNT'])
    elif os.path.exists('service_account.json'):
        with open('service_account.json', 'r') as f:
            creds_json = json.load(f)
            
    if not creds_json:
        print("認証情報が見つかりません")
        return

    creds = Credentials.from_service_account_info(creds_json, scopes=scope)
    gc = gspread.authorize(creds)

    # --- スプレッドシート取得 ---
    # GitHub公開用にハードコードされたURLを削除し、環境変数必須に変更
    sheet_url = os.environ.get('SHEET_URL')
    
    if not sheet_url:
        print("エラー: 環境変数 'SHEET_URL' が設定されていません。GitHub Secretsを確認してください。")
        return

    try:
        wb = gc.open_by_url(sheet_url)
        # 「業種分析」シートを指定
        worksheet = wb.worksheet("業種分析")
    except Exception as e:
        print(f"シートエラー: {e}")
        return

    print("セクターデータの取得を開始します...")

    # --- データ取得 (並列処理) ---
    results = []
    with ThreadPoolExecutor(max_workers=5) as executor:
        # 辞書のアイテムをリスト化して渡す
        futures = [executor.submit(get_sector_data, code, name) for code, name in SECTOR_ETFS.items()]
        for future in futures:
            res = future.result()
            if res:
                results.append(res)

    # 表示順序をコード順にソート
    results.sort(key=lambda x: x[0])

    # --- 書き込み ---
    headers = [
        "コード", "セクター名", "現在値", "前日比(%)", 
        "短期(5日乖離)", "中期(25日乖離)", "長期(75日乖離)", "RSI(過熱感)", "更新日時"
    ]
    
    # データ準備
    all_values = [headers] + results
    
    # クリアしてから書き込み
    worksheet.clear()
    worksheet.update(range_name='A1', values=all_values)
    
    print("書き込み完了！")

if __name__ == "__main__":
    update_sheet()
