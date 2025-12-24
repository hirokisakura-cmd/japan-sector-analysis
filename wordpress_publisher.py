import os
import json
import gspread
import requests
import base64
from google.oauth2.service_account import Credentials
import datetime

def get_sheet_data():
    """Googleスプレッドシートからデータを取得する"""
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    
    # --- 認証情報の読み込み (ここが重要) ---
    creds_json = None
    env_sa = os.environ.get('GCP_SERVICE_ACCOUNT')
    
    # 1. 環境変数(Secrets)から読み込み
    if env_sa:
        try:
            creds_json = json.loads(env_sa)
        except json.JSONDecodeError as e:
            print(f"JSON Decode Error: {e}")
    
    # 2. ローカルファイルから読み込み (フォールバック)
    if not creds_json and os.path.exists('service_account.json'):
        with open('service_account.json', 'r') as f:
            creds_json = json.load(f)

    if not creds_json:
        # エラーメッセージを修正: YAML設定漏れの可能性を示唆
        raise Exception("GCP認証情報が見つかりません。GitHub Actionsのワークフロー(YAML)ファイルで、このステップに環境変数 'GCP_SERVICE_ACCOUNT' が正しく渡されているか確認してください。")

    creds = Credentials.from_service_account_info(creds_json, scopes=scope)
    gc = gspread.authorize(creds)

    # --- シートを開く ---
    sheet_url = os.environ.get('SHEET_URL')
    if not sheet_url:
        raise Exception("SHEET_URLが設定されていません")

    wb = gc.open_by_url(sheet_url)
    worksheet = wb.worksheet("業種分析")
    
    # 全データを取得 (辞書形式のリスト)
    data = worksheet.get_all_records()
    return data

def generate_html_table(data):
    """取得したデータからHTMLテーブルを作成する"""
    if not data:
        return "<p>データがありません。</p>"

    # 更新日時（最新の行から取得）
    last_update = data[0].get("更新日時", datetime.datetime.now().strftime('%Y-%m-%d %H:%M'))

    html = f"<p>最終更新: {last_update}</p>"
    html += '<figure class="wp-block-table"><table>'
    html += '<thead><tr>'
    html += '<th>セクター</th><th>現在値</th><th>前日比</th><th>短期乖離</th><th>RSI</th><th>過熱感(BB)</th>'
    html += '</tr></thead><tbody>'

    for row in data:
        # 必要なカラムを抽出
        sector = row.get("セクター名", "")
        price = row.get("現在値", 0)
        change = row.get("前日比(%)", 0)
        diff_short = row.get("短期(5日乖離)", 0)
        rsi = row.get("RSI", 0)
        bb = row.get("BB%B(過熱)", 0)

        # 色付けのスタイル定義
        style_change = 'style="color: red;"' if float(change) > 0 else 'style="color: blue;"'
        
        # 過熱感の判定 (1.0以上は赤太字、0以下は青太字)
        bb_display = bb
        if float(bb) > 1.0:
            bb_display = f'<strong style="color: red;">{bb}</strong>'
        elif float(bb) < 0:
            bb_display = f'<strong style="color: blue;">{bb}</strong>'

        html += '<tr>'
        html += f'<td>{sector}</td>'
        html += f'<td>{price}</td>'
        html += f'<td {style_change}>{change}%</td>'
        html += f'<td>{diff_short}%</td>'
        html += f'<td>{rsi}</td>'
        html += f'<td>{bb_display}</td>'
        html += '</tr>'

    html += '</tbody></table></figure>'
    html += '<p><small>※TOPIX-17業種ETFのデータを元に算出</small></p>'
    
    return html

def update_wordpress(content):
    """WordPressの固定ページを更新する"""
    wp_url = os.environ.get("WP_URL")
    wp_user = os.environ.get("WP_USER")
    wp_pass = os.environ.get("WP_PASSWORD")
    page_id = os.environ.get("WP_PAGE_ID")

    if not all([wp_url, wp_user, wp_pass, page_id]):
        print("WordPressの設定情報が不足しています。")
        return

    # エンドポイントの構築 (末尾のスラッシュ処理など)
    api_url = f"{wp_url.rstrip('/')}/wp-json/wp/v2/pages/{page_id}"
    
    # 認証ヘッダー
    credentials = f"{wp_user}:{wp_pass}"
    token = base64.b64encode(credentials.encode())
    headers = {
        'Authorization': f'Basic {token.decode("utf-8")}',
        'Content-Type': 'application/json'
    }

    # 送信データ
    payload = {
        'content': content,
        # 必要であればタイトルも更新可能
        # 'title': f"セクター分析レポート ({datetime.datetime.now().strftime('%m/%d')})" 
    }

    print(f"WordPress ({api_url}) へ投稿中...")
    response = requests.post(api_url, headers=headers, json=payload)

    if response.status_code == 200:
        print("投稿成功！")
    else:
        print(f"投稿失敗: {response.status_code}")
        print(response.text)

if __name__ == "__main__":
    try:
        print("データを取得中...")
        data = get_sheet_data()
        
        print("HTMLを生成中...")
        # 簡易的に上位5件だけ表示するなら data[:5] 等にする
        html_content = generate_html_table(data)
        
        update_wordpress(html_content)
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        # GitHub Actionsでエラーを通知するためにexit code 1を返す
        exit(1)
