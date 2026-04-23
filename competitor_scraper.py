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

# ============================================================
# 設定
# ============================================================

APPLICATION_ID = "566f6f9f-f4c7-4ffc-a31a-eb977beaedf9"
ACCESS_KEY     = "pk_80hxyQNTsgpeIiaX1pN6G9p92lF6sWzDqRbQ76UDVXd"

FETCH_DAYS_AHEAD = 60  # 何日先まで取るか

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

def fetch_price(hotel_no, checkin):
    """2食付きプランの最低価格を取得する。"""
    checkout = checkin + timedelta(days=1)
    url = (
        f"https://openapi.rakuten.co.jp/engine/api/Travel/VacantHotelSearch/20170426"
        f"?applicationId={APPLICATION_ID}&accessKey={ACCESS_KEY}"
        f"&hotelNo={hotel_no}&checkinDate={checkin.strftime('%Y-%m-%d')}"
        f"&checkoutDate={checkout.strftime('%Y-%m-%d')}"
        f"&adultNum=2&roomNum=1&hits=30&formatVersion=2&mealCondition=3"
    )
    try:
        resp = requests.get(url, headers=HEADERS, timeout=10)
        data = resp.json()
        if "errors" in data or "error" in data:
            return "×"
        hotels = data.get("hotels", [])
        if not hotels:
            return "×"
        min_price = None
        for hotel in hotels:
            for item in hotel:
                if "roomInfo" in item:
                    rooms = item["roomInfo"]
                    # roomBasicInfo → dailyCharge のペア構造
                    # mealCondition=3 で2食付きは API 側で絞り済みなので
                    # withBreakfastFlag/withDinnerFlag の二重チェックは行わない
                    for j in range(0, len(rooms) - 1, 2):
                        charge = rooms[j + 1].get("dailyCharge", {}) if j + 1 < len(rooms) else {}
                        p = charge.get("total")
                        if p and (min_price is None or p < min_price):
                            min_price = p
        return min_price if min_price else "×"
    except Exception:
        return "×"

# ============================================================
# メイン
# ============================================================

def run():
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    fetch_date = today.strftime("%Y/%m/%d")
    fetch_week = (today - timedelta(days=today.weekday())).strftime("%Y/%m/%d")
    target_dates = [today + timedelta(days=i) for i in range(FETCH_DAYS_AHEAD)]

    file_exists = os.path.exists(OUTPUT_CSV)
    total = len(COMPETITORS) * len(target_dates)
    done = 0

    with open(OUTPUT_CSV, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["取得週", "取得日", "対象日", "施設名", "施設No.", "人数", "朝食", "夕食", "最低価格"])

        for comp in COMPETITORS:
            print(f"\n>> {comp['name']}")
            for d in target_dates:
                price = fetch_price(comp["hotel_no"], d)
                writer.writerow([
                    fetch_week, fetch_date,
                    d.strftime("%Y/%m/%d"),
                    comp["name"], comp["hotel_no"],
                    2, 1, 1, price
                ])
                done += 1
                label = f"{price:,}円" if isinstance(price, int) else price
                print(f"  {d.strftime('%m/%d')} {label}  ({done}/{total}件完了)", end="\r")
                time.sleep(0.3)  # レート制限対策

    print(f"\n\n完了！{done}件を保存しました")
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
