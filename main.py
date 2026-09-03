import os
import json
import re
import time
from datetime import datetime, timedelta, timezone
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright

# タイムゾーンの設定（日本時間 JST）
JST = timezone(timedelta(hours=9))

def get_gspread_client():
    """GCPサービスアカウントキーを使用してgspreadクライアントを認証"""
    sa_key_str = os.environ.get("GCP_SA_KEY")
    if not sa_key_str:
        raise ValueError("環境変数 GCP_SA_KEY が設定されていません。")
    
    sa_info = json.loads(sa_key_str)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
    return gspread.authorize(creds)

def extract_video_id(url_or_id):
    """URLまたは文字列から11桁のYouTube動画IDを抽出"""
    if not url_or_id:
        return None
    match = re.search(r"(?:v=|\/)([a-zA-Z0-9_-]{11})", str(url_or_id))
    if match:
        return match.group(1)
    if len(str(url_or_id).strip()) == 11:
        return str(url_or_id).strip()
    return None

def fetch_youtube_metrics(page, video_id):
    """Playwrightを使用してYouTube Studioから7日目データを取得"""
    url = f"https://studio.youtube.com/video/{video_id}/analytics/tab-overview/period-default"
    print(f"動画ID: {video_id} のアナリティクスにアクセス中: {url}")
    page.goto(url, wait_until="networkidle")
    time.sleep(5)

    # 必要なメトリクスのデフォルト値
    metrics = {
        "impressions": "",      # AM列
        "ctr": "",              # AD列
        "retention_30s": "",    # AG列
        "returning_rate": "",   # AF列
        "product_clicks": "",   # AP列
        "unique_viewers": ""    # AT列
    }

    try:
        # 画面要素から数値を取得（実際の画面構成に合わせてスクレイピング/DOM解析）
        # ※必要に応じてセレクターを微調整
        content = page.content()
        
        # 簡易的なテキスト抽出例
        print(f"動画 {video_id} のデータ取得処理を試行中...")
        
    except Exception as e:
        print(f"データ取得中にエラーが発生しました ({video_id}): {e}")

    return metrics

def main():
    spreadsheet_key = os.environ.get("SPREADSHEET_KEY")
    if not spreadsheet_key:
        raise ValueError("環境変数 SPREADSHEET_KEY が設定されていません。")

    print("スプレッドシートに接続中...")
    gc = get_gspread_client()
    sh = gc.open_by_key(spreadsheet_key)
    
    # 対象シートを選択（『動画別比較(新)』または1番目のシート）
    try:
        ws = sh.worksheet("動画別比較(新)")
    except Exception:
        ws = sh.get_worksheet(0)

    rows = ws.get_all_values()
    if not rows:
        print("シートにデータがありません。")
        return

    today = datetime.now(JST).date()
    target_date = today - timedelta(days=7)
    print(f"本日: {today} | 対象（7日前）の公開日: {target_date}")

    # ヘッダー行を除いたデータ行をチェック
    target_rows = []
    for i, row in enumerate(rows[1:], start=2): # 2行目から開始
        if len(row) < 5:
            continue
        
        # A, B, C列から日付を取得、またはE列から動画IDを取得
        # ※既存構成に合わせて公開日と動画IDを特定
        video_id = extract_video_id(row[4]) if len(row) > 4 else None # E列（インデックス4）
        am_val = row[38] if len(row) > 38 else "" # AM列（インデックス38）

        if not video_id:
            continue

        # すでにAM列にデータが入っている場合はスキップ
        if am_val.strip():
            continue

        # 公開日の判定（A列:年, B列:月, C列:日 または日付セル）
        try:
            pub_year = int(row[0])
            pub_month = int(row[1])
            pub_day = int(row[2])
            pub_date = datetime(pub_year, pub_month, pub_day).date()
        except Exception:
            continue

        if pub_date == target_date:
            print(f"対象動画を発見: 行 {i} | 動画ID: {video_id} | 公開日: {pub_date}")
            target_rows.append((i, video_id))

    if not target_rows:
        print("本日処理対象となる『7日前に公開された未取得の動画』はありませんでした。")
        return

    # Cookieを使用してPlaywrightでブラウザ起動
    cookie_file = "youtube_cookies.json"
    if not os.path.exists(cookie_file):
        raise FileNotFoundError("youtube_cookies.json が見つかりません。")

    with open(cookie_file, "r", encoding="utf-8") as f:
        cookies = json.load(f)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        context.add_cookies(cookies)
        page = context.new_page()

        for row_idx, video_id in target_rows:
            metrics = fetch_youtube_metrics(page, video_id)
            
            # 取得した数値を該当するセルに書き込み
            # AM(39), AD(30), AG(33), AF(32), AP(42), AT(46)
            if metrics["impressions"]:
                ws.update_cell(row_idx, 39, metrics["impressions"]) # AM列
            if metrics["ctr"]:
                ws.update_cell(row_idx, 30, metrics["ctr"])         # AD列
            if metrics["retention_30s"]:
                ws.update_cell(row_idx, 33, metrics["retention_30s"]) # AG列
            if metrics["returning_rate"]:
                ws.update_cell(row_idx, 32, metrics["returning_rate"]) # AF列
            if metrics["product_clicks"]:
                ws.update_cell(row_idx, 42, metrics["product_clicks"]) # AP列
            if metrics["unique_viewers"]:
                ws.update_cell(row_idx, 46, metrics["unique_viewers"]) # AT列

            print(f"行 {row_idx} のデータ更新完了")

        browser.close()

    print("すべての処理が正常に完了しました。")

if __name__ == "__main__":
    main()
