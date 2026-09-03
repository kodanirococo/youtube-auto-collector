import os
import json
import re
import time
from datetime import datetime, timedelta, timezone
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright

JST = timezone(timedelta(hours=9))

def get_gspread_client():
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
    if not url_or_id:
        return None
    match = re.search(r"(?:v=|\/)([a-zA-Z0-9_-]{11})", str(url_or_id))
    if match:
        return match.group(1)
    if len(str(url_or_id).strip()) == 11:
        return str(url_or_id).strip()
    return None

def fetch_youtube_metrics(page, video_id):
    """
    YouTube Studioから指定動画のアナリティクスデータを取得
    """
    base_url = f"https://studio.youtube.com/video/{video_id}/analytics"
    
    metrics = {
        "ctr": "",               # AD列 (30)
        "returning_rate": "",    # AF列 (32)
        "retention_30s": "",     # AG列 (33)
        "impressions": "",       # AM列 (39)
        "product_clicks": "",    # AP列 (42)
        "new_viewers": "",       # AR列 (44)
        "returning_viewers": "", # AS列 (45)
        "unique_viewers": ""     # AT列 (46)
    }

    try:
        # 1. リーチタブ（インプレッション数・クリック率）
        reach_url = f"{base_url}/tab-reach_tab/period-default"
        print(f"[{video_id}] リーチタブアクセス: {reach_url}")
        page.goto(reach_url, wait_until="domcontentloaded")
        time.sleep(5)

        # ログイン・アクセスの確認
        page_title = page.title()
        print(f"[{video_id}] ページタイトル: {page_title}")

        try:
            body_text = page.inner_text("body")
            
            # インプレッション数の抽出試行
            imp_match = re.search(r"インプレッション\s*([\d,]+)", body_text)
            if imp_match:
                metrics["impressions"] = imp_match.group(1).replace(",", "")
            
            # CTR（クリック率）の抽出試行
            ctr_match = re.search(r"インプレッションのクリック率\s*([\d\.]+%?)", body_text)
            if ctr_match:
                metrics["ctr"] = ctr_match.group(1)
                
        except Exception as e:
            print(f"[{video_id}] リーチ解析警告: {e}")

        # 2. 視聴者タブ（リピーター・新しい視聴者・ユニーク視聴者）
        audience_url = f"{base_url}/tab-audience/period-default"
        print(f"[{video_id}] 視聴者タブアクセス: {audience_url}")
        page.goto(audience_url, wait_until="domcontentloaded")
        time.sleep(5)
        
        try:
            body_text = page.inner_text("body")
            
            new_v_match = re.search(r"新しい視聴者\s*([\d,]+)", body_text)
            if new_v_match:
                metrics["new_viewers"] = new_v_match.group(1).replace(",", "")

            ret_v_match = re.search(r"リピーター\s*([\d,]+)", body_text)
            if ret_v_match:
                metrics["returning_viewers"] = ret_v_match.group(1).replace(",", "")

            uniq_v_match = re.search(r"ユニーク視聴者\s*([\d,]+)", body_text)
            if uniq_v_match:
                metrics["unique_viewers"] = uniq_v_match.group(1).replace(",", "")

            # リピート再生率の計算
            if metrics["returning_viewers"] and metrics["new_viewers"]:
                ret_v = float(metrics["returning_viewers"])
                new_v = float(metrics["new_viewers"])
                if (ret_v + new_v) > 0:
                    rate = (ret_v / (ret_v + new_v)) * 100
                    metrics["returning_rate"] = f"{rate:.2f}%"
        except Exception as e:
            print(f"[{video_id}] 視聴者解析警告: {e}")

        print(f"[{video_id}] 取得成功データ: {metrics}")

    except Exception as e:
        print(f"[{video_id}] エラーが発生しました: {e}")

    return metrics

def main():
    spreadsheet_key = os.environ.get("SPREADSHEET_KEY")
    if not spreadsheet_key:
        raise ValueError("環境変数 SPREADSHEET_KEY が設定されていません。")

    print("スプレッドシートに接続中...")
    gc = get_gspread_client()
    sh = gc.open_by_key(spreadsheet_key)
    
    try:
        ws = sh.worksheet("動画別比較(新)")
    except Exception:
        ws = sh.get_worksheet(0)

    rows = ws.get_all_values()
    if not rows:
        print("シートにデータがありません。")
        return

    today = datetime.now(JST).date()
    print(f"実行日: {today}")

    target_rows = []
    # ヘッダーを除外して順次処理（2行目以降）
    for i, row in enumerate(rows[1:], start=2):
        if len(row) < 5:
            continue

        video_id = extract_video_id(row[4]) if len(row) > 4 else None # E列
        am_val = row[38] if len(row) > 38 else ""                     # AM列 (インプレッション数)

        if not video_id:
            continue

        # すでにインプレッション数(AM列)が入っている場合はスキップ
        if am_val.strip():
            continue

        try:
            pub_year = int(row[1])  # B列: 年
            pub_month = int(row[2]) # C列: 月
            pub_day = int(row[3])   # D列: 日
            pub_date = datetime(pub_year, pub_month, pub_day).date()
        except Exception:
            continue

        # 公開から7日以上経過しているかを判定
        days_passed = (today - pub_date).days
        if days_passed >= 7:
            print(f"対象動画を発見: 行 {i} | 動画ID: {video_id} | 公開日: {pub_date} ({days_passed}日前)")
            target_rows.append((i, video_id))

    if not target_rows:
        print("処理対象となる未取得の動画はありませんでした。")
        return

    cookie_file = "youtube_cookies.json"
    if not os.path.exists(cookie_file):
        raise FileNotFoundError("youtube_cookies.json が見つかりません。")

    with open(cookie_file, "r", encoding="utf-8") as f:
        cookies = json.load(f)

    # クッキーデータのクレンジング
    valid_samesite = ["Strict", "Lax", "None"]
    for cookie in cookies:
        if "sameSite" in cookie:
            s_val = str(cookie["sameSite"]).capitalize()
            if s_val in valid_samesite:
                cookie["sameSite"] = s_val
            else:
                del cookie["sameSite"]
        
        for key in ["storeId", "hostOnly", "session", "id", "partitionKey"]:
            cookie.pop(key, None)

    # ブラウザの自動操作・データ更新
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        context.add_cookies(cookies)
        page = context.new_page()

        for row_idx, video_id in target_rows:
            print(f"\n--- 行 {row_idx} (動画ID: {video_id}) のデータ取得開始 ---")
            metrics = fetch_youtube_metrics(page, video_id)
            
            # 各列へのセル書き込み (値がある場合のみ)
            if metrics["ctr"]:
                ws.update_cell(row_idx, 30, metrics["ctr"])             # AD列 (30)
            if metrics["returning_rate"]:
                ws.update_cell(row_idx, 32, metrics["returning_rate"])  # AF列 (32)
            if metrics["retention_30s"]:
                ws.update_cell(row_idx, 33, metrics["retention_30s"])  # AG列 (33)
            if metrics["impressions"]:
                ws.update_cell(row_idx, 39, metrics["impressions"])    # AM列 (39)
            if metrics["product_clicks"]:
                ws.update_cell(row_idx, 42, metrics["product_clicks"]) # AP列 (42)
            if metrics["new_viewers"]:
                ws.update_cell(row_idx, 44, metrics["new_viewers"])     # AR列 (44)
            if metrics["returning_viewers"]:
                ws.update_cell(row_idx, 45, metrics["returning_viewers"]) # AS列 (45)
            if metrics["unique_viewers"]:
                ws.update_cell(row_idx, 46, metrics["unique_viewers"]) # AT列 (46)

            print(f"行 {row_idx} の書き込み処理が完了しました。")

        browser.close()

    print("\nすべての処理が正常に完了しました。")

if __name__ == "__main__":
    main()
