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
    
    # --- GCP認証情報の読み込み ---
    creds_json = None
    env_sa = os.environ.get('GCP_SERVICE_ACCOUNT')
    
    if env_sa:
        try:
            creds_json = json.loads(env_sa)
        except json.JSONDecodeError as e:
            print(f"JSON Decode Error: {e}")
    
    if not creds_json and os.path.exists('service_account.json'):
        with open('service_account.json', 'r') as f:
            creds_json = json.load(f)

    if not creds_json:
        raise Exception("GCP認証情報が見つかりません。SecretsのGCP_SERVICE_ACCOUNTを確認してください。")

    creds = Credentials.from_service_account_info(creds_json, scopes=scope)
    gc = gspread.authorize(creds)

    # --- シートを開く ---
    sheet_url = os.environ.get('SHEET_URL')
    if not sheet_url:
        raise Exception("SHEET_URLが設定されていません")

    wb = gc.open_by_url(sheet_url)
    worksheet = wb.worksheet("業種分析")
    data = worksheet.get_all_records()
    return data

def generate_html_table(data):
    """HTMLテーブル生成"""
    if not data:
        return "<p>データがありません。</p>"

    last_update = data[0].get("更新日時", datetime.datetime.now().strftime('%Y-%m-%d %H:%M'))
    html = f"<p>最終更新: {last_update}</p>"
    html += '<figure class="wp-block-table"><table>'
    html += '<thead><tr><th>セクター</th><th>現在値</th><th>前日比</th><th>短期乖離</th><th>RSI</th><th>過熱感(BB)</th></tr></thead><tbody>'

    for row in data:
        sector = row.get("セクター名", "")
        price = row.get("現在値", 0)
        change = row.get("前日比(%)", 0)
        diff_short = row.get("短期(5日乖離)", 0)
        rsi = row.get("RSI", 0)
        bb = row.get("BB%B(過熱)", 0)

        style_change = 'style="color: red;"' if float(change) > 0 else 'style="color: blue;"'
        
        bb_display = bb
        if float(bb) > 1.0:
            bb_display = f'<strong style="color: red;">{bb}</strong>'
        elif float(bb) < 0:
            bb_display = f'<strong style="color: blue;">{bb}</strong>'

        html += f'<tr><td>{sector}</td><td>{price}</td><td {style_change}>{change}%</td><td>{diff_short}%</td><td>{rsi}</td><td>{bb_display}</td></tr>'

    html += '</tbody></table></figure>'
    html += '<p><small>※TOPIX-17業種ETFのデータを元に算出</small></p>'
    return html

def get_wordpress_config():
    """
    環境変数 TOFU_WORDPRESS または 個別の環境変数から設定を取得する
    """
    config = {
        "url": os.environ.get("WP_URL"),
        "user": os.environ.get("WP_USER"),
        "password": os.environ.get("WP_PASSWORD"),
        "page_id": os.environ.get("WP_PAGE_ID"),
    }

    # まとめて登録された TOFU_WORDPRESS があれば解析して上書き
    tofu_secret = os.environ.get("TOFU_WORDPRESS")
    if tofu_secret:
        print("Secrets 'TOFU_WORDPRESS' を読み込み中...")
        for line in tofu_secret.splitlines():
            line = line.strip()
            if not line or "=" not in line:
                continue
            
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            
            if key == "WP_URL": config["url"] = value
            elif key == "WP_USER": config["user"] = value
            elif key == "WP_PASSWORD": config["password"] = value
            elif key == "WP_PAGE_ID": config["page_id"] = value

    return config

def update_wordpress(content):
    """WordPress更新"""
    # 設定を取得
    wp_config = get_wordpress_config()
    
    wp_url = wp_config["url"]
    wp_user = wp_config["user"]
    wp_pass = wp_config["password"]
    page_id = wp_config["page_id"]

    if not all([wp_url, wp_user, wp_pass, page_id]):
        print("エラー: WordPressの設定情報が不足しています。Secrets 'TOFU_WORDPRESS' の内容を確認してください。")
        return

    api_url = f"{wp_url.rstrip('/')}/wp-json/wp/v2/pages/{page_id}"
    credentials = f"{wp_user}:{wp_pass}"
    token = base64.b64encode(credentials.encode())
    headers = {
        'Authorization': f'Basic {token.decode("utf-8")}',
        'Content-Type': 'application/json'
    }
    payload = {'content': content}

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
        html_content = generate_html_table(data)
        update_wordpress(html_content)
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        exit(1)
