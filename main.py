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
        "ctr": "",              # AD列 (30)
        "returning_rate": "",   # AF列 (32)
        "retention_30s": "",    # AG列 (33)
        "impressions": "",      # AM列 (39)
        "product_clicks": "",   # AP列 (42)
        "new_viewers": "",      # AR列 (44)
        "returning_viewers": "",# AS列 (45)
        "unique_viewers": ""    # AT列 (46)
    }

    try:
        # 1. 概要 / リーチタブ（インプレッション数・CTR）
        overview_url = f"{base_url}/tab-overview/period-default"
        print(f"[{video_id}] 概要タブ読み込み中: {overview_url}")
        page.goto(overview_url, wait_until="networkidle")
        time.sleep(4)

        # 概要/リーチ要素の解析
        content = page.content()
        
        # リーチ/インプレッション解析
        reach_url = f"{base_url}/tab-reach_tab/period-default"
        page.goto(reach_url, wait_until="networkidle")
        time.sleep(4)
        
        # インプレッション数 & CTRの取得試行
        try:
            metric_cards = page.query_selector_all("ytcp-metric-visualizer")
            for card in metric_cards:
                text = card.inner_text()
                if "インプレッション" in text and "クリック率" not in text:
                    val = re.search(r"[\d,]+", text)
                    if val:
                        metrics["impressions"] = val.group(0).replace(",", "")
                elif "クリック率" in text:
                    val = re.search(r"[\d\.]+%?", text)
                    if val:
                        metrics["ctr"] = val.group(0)
        except Exception as e:
            print(f"[{video_id}] リーチメトリクス解析警告: {e}")

        # 2. エンゲージメントタブ（30秒時点の維持率）
        engagement_url = f"{base_url}/tab-engagement/period-default"
        page.goto(engagement_url, wait_until="networkidle")
        time.sleep(4)
        try:
            retention_elem = page.query_selector(".key-moments-retention, [id*='retention']")
            if retention_elem:
                val = re.search(r"[\d\.] process%", retention_elem.inner_text())
                if val:
                    metrics["retention_30s"] = val.group(0)
        except Exception as e:
            print(f"[{video_id}] 30秒維持率解析警告: {e}")

        # 3. 視聴者タブ（リピーター数・新しい視聴者数・ユニーク視聴者数）
        audience_url = f"{base_url}/tab-audience/period-default"
        page.goto(audience_url, wait_until="networkidle")
        time.sleep(4)
        try:
            audience_cards = page.query_selector_all("ytcp-metric-visualizer, .metric-value")
            for card in audience_cards:
                text = card.inner_text()
                if "新しい視聴者" in text:
                    val = re.search(r"[\d,]+", text)
                    if val:
                        metrics["new_viewers"] = val.group(0).replace(",", "")
                elif "リピーター" in text and "比率" not in text:
                    val = re.search(r"[\d,]+", text)
                    if val:
                        metrics["returning_viewers"] = val.group(0).replace(",", "")
                elif "ユニーク視聴者" in text:
                    val = re.search(r"[\d,]+", text)
                    if val:
                        metrics["unique_viewers"] = val.group(0).replace(",", "")
            
            # リピーター再生率の計算（リピーター数 / (リピーター数 + 新しい視聴者数)）
            if metrics["returning_viewers"] and metrics["new_viewers"]:
                ret_v = float(metrics["returning_viewers"])
                new_v = float(metrics["new_viewers"])
                if (ret_v + new_v) > 0:
                    rate = (ret_v / (ret_v + new_v)) * 100
                    metrics["returning_rate"] = f"{rate:.2f}%"
        except Exception as e:
            print(f"[{video_id}] 視聴者メトリクス解析警告: {e}")

        # 4. 商品クリック数（詳細アナリティクスより）
        # ※カード/概要欄クリックデータ
        try:
            product_elem = page.query_selector("[id*='product-click'], [class*='shopping']")
            if product_elem:
                val = re.search(r"[\d,]+", product_elem.inner_text())
                if val:
                    metrics["product_clicks"] = val.group(0).replace(",", "")
        except Exception as e:
            print(f"[{video_id}] 商品クリック数解析警告: {e}")

    except Exception as e:
        print(f"[{video_id}] データ取得中にエラーが発生しました: {e}")

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
    # 2行目から順にチェック
    for i, row in enumerate(rows[1:], start=2):
        if len(row) < 5:
            continue

        # E列（インデックス4）から動画IDを取得
        video_id = extract_video_id(row[4]) if len(row) > 4 else None
        # AM列（インデックス38）のインプレッション数値を確認
        am_val = row[38] if len(row) > 38 else ""

        if not video_id:
            continue

        # すでにインプレッションが入力済みの場合はスキップ
        if am_val.strip():
            continue

        # 公開日の判定（B列:年[1], C列:月[2], D列:日[3]）
        try:
            pub_year = int(row[1])
            pub_month = int(row[2])
            pub_day = int(row[3])
            pub_date = datetime(pub_year, pub_month, pub_day).date()
        except Exception:
            continue

        # 公開から7日以上経過している未取得動画を対象にする
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

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        context.add_cookies(cookies)
        page = context.new_page()

        for row_idx, video_id in target_rows:
            print(f"\n--- 行 {row_idx} (動画ID: {video_id}) のデータ取得開始 ---")
            metrics = fetch_youtube_metrics(page, video_id)
            
            # 各列への書き込み
            # AD列: 30, AF列: 32, AG列: 33, AM列: 39, AP列: 42, AR列: 44, AS列: 45, AT列: 46
            if metrics["ctr"]:
                ws.update_cell(row_idx, 30, metrics["ctr"])             # AD列
            if metrics["returning_rate"]:
                ws.update_cell(row_idx, 32, metrics["returning_rate"])  # AF列
            if metrics["retention_30s"]:
                ws.update_cell(row_idx, 33, metrics["retention_30s"])  # AG列
            if metrics["impressions"]:
                ws.update_cell(row_idx, 39, metrics["impressions"])    # AM列
            if metrics["product_clicks"]:
                ws.update_cell(row_idx, 42, metrics["product_clicks"]) # AP列
            if metrics["new_viewers"]:
                ws.update_cell(row_idx, 44, metrics["new_viewers"])     # AR列
            if metrics["returning_viewers"]:
                ws.update_cell(row_idx, 45, metrics["returning_viewers"]) # AS列
            if metrics["unique_viewers"]:
                ws.update_cell(row_idx, 46, metrics["unique_viewers"]) # AT列

            print(f"行 {row_idx} の全メトリクス更新完了")

        browser.close()

    print("\nすべての処理が正常に完了しました。")

if __name__ == "__main__":
    main()
