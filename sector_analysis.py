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

def calculate_technical_indicators(df):
    """データフレーム全体に対してテクニカル指標を一括計算する"""
    df = df.copy()
    
    # 1. 移動平均乖離率
    df['ma5'] = df['Close'].rolling(window=5).mean()
    df['ma25'] = df['Close'].rolling(window=25).mean()
    df['ma75'] = df['Close'].rolling(window=75).mean()
    
    df['diff_short'] = ((df['Close'] - df['ma5']) / df['ma5']) * 100
    df['diff_mid'] = ((df['Close'] - df['ma25']) / df['ma25']) * 100
    df['diff_long'] = ((df['Close'] - df['ma75']) / df['ma75']) * 100

    # 2. RSI (14日)
    delta = df['Close'].diff()
    gain = (delta.where(delta > 0, 0)).rolling(window=14).mean()
    loss = (-delta.where(delta < 0, 0)).rolling(window=14).mean()
    rs = gain / loss
    df['rsi'] = 100 - (100 / (1 + rs))

    # 3. ボリンジャーバンド %B (20日, 2σ)
    # %B > 1.0 はバンド突破（過熱）、 < 0 はバンド割れ（売られすぎ）
    df['bb_ma'] = df['Close'].rolling(window=20).mean()
    df['bb_std'] = df['Close'].rolling(window=20).std()
    df['bb_up'] = df['bb_ma'] + (df['bb_std'] * 2)
    df['bb_low'] = df['bb_ma'] - (df['bb_std'] * 2)
    # ゼロ除算回避
    bb_range = df['bb_up'] - df['bb_low']
    df['bb_pct_b'] = np.where(bb_range == 0, 0, (df['Close'] - df['bb_low']) / bb_range)

    # 4. 出来高倍率 (直近5日平均との比較)
    df['vol_ma5'] = df['Volume'].rolling(window=5).mean()
    # ゼロ除算回避
    df['vol_ratio'] = np.where(df['vol_ma5'] == 0, 0, df['Volume'] / df['vol_ma5'])

    # 5. 前日比
    df['change_pct'] = df['Close'].pct_change() * 100

    return df

def get_sector_data(code, name, is_initial_run=False):
    """
    指定銘柄のデータを取得・計算
    is_initial_run=Trueなら過去データ全て、Falseなら最新1行のみ返す
    """
    ticker = f"{code}.T"
    try:
        stock = yf.Ticker(ticker)
        # 過去2年分取得（移動平均の計算用バッファ含む）
        hist = stock.history(period="2y")
        
        if hist.empty:
            return []

        # 指標計算
        df = calculate_technical_indicators(hist)
        
        # NaNを除去し、直近1年(250営業日)分に絞る
        df = df.dropna().tail(250) 

        # 行データ作成用ヘルパー関数
        def make_row(date_idx, row):
            return [
                code,                                            # A列: コード
                name,                                            # B列: セクター名
                date_idx.strftime('%Y-%m-%d'),                   # C列: 日付
                round(row['Close'], 1),                          # D列: 現在値
                round(row['change_pct'], 2),                     # E列: 前日比(%)
                round(row['diff_short'], 2),                     # F列: 短期乖離
                round(row['diff_mid'], 2),                       # G列: 中期乖離
                round(row['diff_long'], 2),                      # H列: 長期乖離
                round(row['rsi'], 1),                            # I列: RSI
                round(row['bb_pct_b'], 2),                       # J列: BB %B (過熱感)
                round(row['vol_ratio'], 2),                      # K列: 出来高倍率
                datetime.datetime.now().strftime('%Y-%m-%d %H:%M') # L列: 更新日時
            ]

        results = []
        if is_initial_run:
            # 過去すべての行をリスト化 (日付の新しい順)
            for date_idx, row in df.iloc[::-1].iterrows():
                results.append(make_row(date_idx, row))
        else:
            # 最新の1行のみ
            if len(df) > 0:
                last_date = df.index[-1]
                last_row = df.iloc[-1]
                results.append(make_row(last_date, last_row))
            
        return results

    except Exception as e:
        print(f"Error {code}: {e}")
        return []

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
    sheet_url = os.environ.get('SHEET_URL')
    
    if not sheet_url:
        print("エラー: 環境変数 'SHEET_URL' が設定されていません。GitHub Secretsを確認してください。")
        return

    try:
        wb = gc.open_by_url(sheet_url)
        # シート名指定（変更なし）
        worksheet = wb.worksheet("業種分析")
    except Exception as e:
        print(f"シートエラー: {e}")
        return

    print("スプレッドシートの状態を確認中...")
    
    # 初回チェック (A1セルが空なら初回とみなす)
    try:
        val_a1 = worksheet.acell('A1').value
        is_initial_run = not bool(val_a1)
    except:
        # シートが真っ白などの場合
        is_initial_run = True

    print(f"処理モード: {'初回一括作成(過去データ含む)' if is_initial_run else '通常更新(最新データ追記)'}")
    print("セクターデータの取得を開始します...")

    # --- データ取得 (並列処理) ---
    all_rows = []
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(get_sector_data, code, name, is_initial_run) for code, name in SECTOR_ETFS.items()]
        for future in futures:
            res = future.result()
            if res:
                all_rows.extend(res)

    # ソート: 日付(新しい順) > コード順
    # x[2]は日付文字列, x[0]はコード
    all_rows.sort(key=lambda x: (x[2], x[0]), reverse=True)

    # --- ヘッダー定義 ---
    headers = [
        "コード", "セクター名", "日付", "現在値", "前日比(%)", 
        "短期(5日乖離)", "中期(25日乖離)", "長期(75日乖離)", 
        "RSI", "BB%B(過熱)", "出来高倍率", "更新日時"
    ]
    
    # --- 書き込み ---
    if is_initial_run:
        print(f"{len(all_rows)}件のデータを書き込みます...")
        worksheet.clear()
        worksheet.update(range_name='A1', values=[headers] + all_rows)
        print("初期化書き込み完了！")
    else:
        if all_rows:
            print(f"{len(all_rows)}件の最新データを追記します...")
            # ヘッダー行(1行目)の次、つまり2行目に挿入
            worksheet.insert_rows(all_rows, row=2)
            print("追記完了！")
        else:
            print("更新データがありませんでした")

if __name__ == "__main__":
    update_sheet()
