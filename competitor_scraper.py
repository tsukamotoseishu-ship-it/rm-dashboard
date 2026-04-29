"""
競合価格収集スクリプト - ホテル甲子園 RMシステム
================================================
楽天トラベル APIを使って競合12施設の価格を取得する。

【実行方法】
python competitor_scraper.py
"""

import requests
import csv
import os
import time
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

# ============================================================
# 設定
# ============================================================

APPLICATION_ID = "566f6f9f-f4c7-4ffc-a31a-eb977beaedf9"
ACCESS_KEY     = "pk_80hxyQNTsgpeIiaX1pN6G9p92lF6sWzDqRbQ76UDVXd"

FETCH_DAYS_AHEAD = 30  # 何日先まで取るか
MAX_WORKERS      = 5   # 並列数（同時リクエスト数）

OUTPUT_CSV = r"C:\Users\tsukamoto.seishu\rm_system\competitor_prices.csv"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Referer": "https://example.com",
    "Origin": "https://example.com",
}

# ============================================================
# 競合施設リスト
# ============================================================

COMPETITORS = [
    {"name": "【2食】銘庭の宿　ホテル甲子園",       "hotel_no": 18659},
    {"name": "【2食】糸柳別館　離れの邸　和穣苑",    "hotel_no": 72019},
    {"name": "【2食】石和名湯館　糸柳",              "hotel_no": 16390},
    {"name": "【2食】石和温泉　銘石の宿　かげつ",    "hotel_no": 16653},
    {"name": "【2食】別邸　花水晶",                  "hotel_no": 79333},
    {"name": "【2食】石和温泉　ホテル古柏園",        "hotel_no": 7024},
    {"name": "【2食】笛吹川温泉　坐忘",              "hotel_no": 72676},
    {"name": "【2食】石和温泉　糸柳こやど　ゆわ",    "hotel_no": 20033},
    {"name": "【2食】石和温泉　ホテルふじ",          "hotel_no": 27753},
    {"name": "【2食】川浦温泉　山県館",              "hotel_no": 67105},
    {"name": "【2食】石和温泉　華やぎの章　慶山",    "hotel_no": 19894},
    {"name": "【2食】湯めぐり宿　笛吹川",            "hotel_no": 32212},
    {"name": "【2食】石和温泉郷　旅館深雪温泉",      "hotel_no": 16067},
]

# ============================================================
# 価格取得
# ============================================================

# 2食付きプランと判断するキーワード（プラン名に含まれれば採用）
MEAL_KEYWORDS = [
    '2食', '二食', '朝夕', '夕朝', '朝食・夕食', '夕食・朝食',
    '食事付', '夕食付', '朝食付', 'ディナー', '懐石', '会席',
    'プレミアム', 'スタンダード', '特選', '旬の', '季節', '温泉',
]
# 素泊まり・朝食のみと判断するキーワード（含まれれば除外）
EXCLUDE_KEYWORDS = [
    '素泊', '室料', '朝のみ', '朝食のみ', '食事なし', 'without', '夕食なし',
    '0食', '1泊朝食', '一泊朝食', '朝食付き', '朝食付',  # 朝食だけのプランを除外
]
# 除外ワードが入っていても「夕食」も含む場合は2食なので再採用（朝夕食付き等）
INCLUDE_OVERRIDE = ['夕食', '朝夕', '2食', '二食']


def is_meal_plan(plan_name: str, breakfast_flag: int = 0, dinner_flag: int = 0) -> bool:
    """
    プラン名・食事フラグから2食付きかどうかを判定する。
    対象施設は全て旅館なので、「素泊まり系ワードがなければ2食とみなす」が基本方針。
    APIがメタデータ（朝夕フラグ）を正しく設定していない旅館が多いため。
    """
    # 除外ワードが含まれる場合は確実にNG（素泊まり・朝食のみ）
    if any(k in plan_name for k in EXCLUDE_KEYWORDS):
        return False
    # 除外ワードがない → 旅館の場合はほぼ2食付きとみなす
    return True


def _call_api(hotel_no, checkin, meal_condition=None):
    """楽天APIを呼び出してレスポンスを返す。エラー時は None。"""
    checkout = checkin + timedelta(days=1)
    url = (
        f"https://openapi.rakuten.co.jp/engine/api/Travel/VacantHotelSearch/20170426"
        f"?applicationId={APPLICATION_ID}&accessKey={ACCESS_KEY}"
        f"&hotelNo={hotel_no}&checkinDate={checkin.strftime('%Y-%m-%d')}"
        f"&checkoutDate={checkout.strftime('%Y-%m-%d')}"
        f"&adultNum=2&roomNum=1&hits=30&formatVersion=2"
    )
    if meal_condition is not None:
        url += f"&mealCondition={meal_condition}"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=10)
        data = resp.json()
        if "errors" in data or "error" in data:
            return None
        return data
    except Exception:
        return None


def _extract_min_price(data, strict_filter=True):
    """
    APIレスポンスから2食付き最低価格を抽出する。
    strict_filter=True  → 除外ワードありプランを弾く（素泊まりを除外）
    strict_filter=False → 除外ワードなければ全て2食とみなす（食事フラグ未設定旅館向け）
    """
    hotels = data.get("hotels", [])
    if not hotels:
        return None
    min_price = None
    for hotel in hotels:
        for item in hotel:
            if "roomInfo" not in item:
                continue
            rooms = item["roomInfo"]
            for j in range(0, len(rooms) - 1, 2):
                basic  = rooms[j].get("roomBasicInfo", {})
                charge = rooms[j + 1].get("dailyCharge", {}) if j + 1 < len(rooms) else {}
                p = charge.get("total")
                if not p:
                    continue
                breakfast = basic.get("withBreakfastFlag", 0)
                dinner    = basic.get("withDinnerFlag", 0)
                plan_name = basic.get("planName", "") or basic.get("roomName", "")
                if strict_filter:
                    # 除外ワードがあればNG、なければOK
                    if any(k in plan_name for k in EXCLUDE_KEYWORDS):
                        continue
                else:
                    # フラグ or 除外ワードなし で判定
                    if not is_meal_plan(plan_name, breakfast, dinner):
                        continue
                if min_price is None or p < min_price:
                    min_price = p
    return min_price


def _scrape_plan_page(page, hotel_no, checkin):
    """
    Playwright page オブジェクトで楽天トラベルのプランページを取得し
    2食付き最低価格を返す。素泊まり・朝食のみは除外。
    """
    d = checkin
    url = (
        f"https://hotel.travel.rakuten.co.jp/hotelinfo/plan/{hotel_no}"
        f"?f_flg=PLAN&f_hi1={d.day}&f_tuki1={d.month}&f_nen1={d.year}"
        f"&f_hi2={(d + timedelta(days=1)).day}&f_tuki2={(d + timedelta(days=1)).month}&f_nen2={(d + timedelta(days=1)).year}"
        f"&f_heya_su=1&f_otona_su=2&f_s1=0&f_s2=0&f_y1=0&f_y2=0&f_y3=0&f_y4=0"
    )
    try:
        # load で早めに切り上げ（networkidle は重いページでタイムアウトしやすい）
        page.goto(url, wait_until="load", timeout=30000)
        page.wait_for_timeout(2000)  # 動的コンテンツの読み込み待ち
    except Exception:
        return "×"

    plans = page.query_selector_all('[data-role="planArea"]')
    min_price = None
    for plan in plans:
        h4 = plan.query_selector("h4")
        name = h4.inner_text().strip() if h4 else ""
        # 素泊まり・朝食のみ除外
        # ただし除外ワードがあっても夕食も含む場合（朝夕食付き等）は2食として採用
        if any(k in name for k in EXCLUDE_KEYWORDS):
            if not any(k in name for k in INCLUDE_OVERRIDE):
                continue
        # 合計価格を複数セレクタで取得（施設によってHTML構造が異なる）
        price_text = None
        for sel in [".ndPrice strong", ".priceNum", ".prc strong", "[class*='price'] strong", "strong"]:
            price_el = plan.query_selector(sel)
            if price_el:
                txt = price_el.inner_text().replace(",", "").strip()
                if txt.isdigit() and int(txt) > 1000:
                    price_text = txt
                    break
        if not price_text:
            continue
        try:
            p = int(price_text) // 2  # 2名合計 → 1人あたり
            if min_price is None or p < min_price:
                min_price = p
        except Exception:
            continue
    return min_price if min_price else "×"


def fetch_hotel_all_dates(comp, target_dates, fetch_week, fetch_date):
    """
    1施設の全日付をPlaywrightで取得する。
    ブラウザを1施設で使い回してオーバーヘッドを削減。
    """
    from playwright.sync_api import sync_playwright
    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0"
        )
        for d in target_dates:
            price = _scrape_plan_page(page, comp["hotel_no"], d)
            results.append([
                fetch_week, fetch_date,
                d.strftime("%Y/%m/%d"),
                comp["name"], comp["hotel_no"],
                2, 1, 1, price
            ])
            time.sleep(0.3)
        browser.close()
    return results


def run():
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    fetch_date = today.strftime("%Y/%m/%d")
    fetch_week = (today - timedelta(days=today.weekday())).strftime("%Y/%m/%d")
    target_dates = [today + timedelta(days=i) for i in range(FETCH_DAYS_AHEAD)]

    file_exists = os.path.exists(OUTPUT_CSV)
    total = len(COMPETITORS) * len(target_dates)
    print(f"取得開始：{len(COMPETITORS)}施設 × {len(target_dates)}日 = {total}件")
    print(f"並列数: {MAX_WORKERS}  推定時間: {total * 0.15 / MAX_WORKERS / 60:.1f}〜{total * 0.3 / MAX_WORKERS / 60:.1f}分\n")

    all_rows = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(fetch_hotel_all_dates, comp, target_dates, fetch_week, fetch_date): comp
            for comp in COMPETITORS
        }
        done_hotels = 0
        for future in as_completed(futures):
            comp = futures[future]
            try:
                rows = future.result()
                all_rows.extend(rows)
                done_hotels += 1
                prices = [r[8] for r in rows if isinstance(r[8], int)]
                avg = f"¥{sum(prices)//len(prices):,}" if prices else "−"
                print(f"  [{done_hotels}/{len(COMPETITORS)}] {comp['name']}  平均{avg}")
            except Exception as e:
                print(f"  ERROR {comp['name']}: {e}")

    # CSV書き込み（取得完了後にまとめて書く）
    with open(OUTPUT_CSV, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["取得週", "取得日", "対象日", "施設名", "施設No.", "人数", "朝食", "夕食", "最低価格"])
        for row in sorted(all_rows, key=lambda r: (r[3], r[2])):  # 施設名・日付順
            writer.writerow(row)

    print(f"\n完了！{len(all_rows)}件を保存しました")
    print(f"保存先: {OUTPUT_CSV}")

    # サマリー表示
    print("\n=== 本日の競合価格サマリー（先10日） ===")
    print(f"{'対象日':<12} {'競合最低':>8} {'競合平均':>8}")
    print("-" * 35)
    import collections
    by_date = collections.defaultdict(list)
    with open(OUTPUT_CSV, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if row["取得日"] == fetch_date and row["最低価格"] not in ("×", ""):
                by_date[row["対象日"]].append(int(row["最低価格"]))
    for d in sorted(by_date)[:10]:
        vals = by_date[d]
        print(f"{d:<12} {min(vals):>8,} {int(sum(vals)/len(vals)):>8,}")

if __name__ == "__main__":
    run()
