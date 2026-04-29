"""
RM共通ロジック - ホテル甲子園
export_rm_excel.py と app.py の両方から利用される共通モジュール
"""

import csv, glob, io, math, os
from datetime import datetime, timedelta
from collections import defaultdict

# ============================================================
# 設定
# ============================================================
TODAY       = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
DAYS_AHEAD  = 30
TOTAL_ROOMS = 17

_BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
CSV_DIR       = r"C:\Users\tsukamoto.seishu\Downloads"
PMS_CSV       = r"C:\Users\tsukamoto.seishu\Downloads\a.csv"
COMP_PRICES_CSV = os.path.join(_BASE_DIR, "competitor_prices.csv")
COMP_PRICES_SAMPLE = os.path.join(_BASE_DIR, "competitor_prices_sample.csv")

RANKS = ['A1','B1','C1','D1','E1','F1','G1','H1','I1','J1',
         'K1','L1','M1','N1','O1','P1','Q1','R1']
RANK_PRICE = {
    'A1':13300,'B1':15300,'C1':16300,'D1':17800,'E1':19300,'F1':20800,
    'G1':22300,'H1':23800,'I1':25300,'J1':26800,'K1':28300,'L1':29800,
    'M1':31300,'N1':32800,'O1':34300,'P1':35800,'Q1':37300,'R1':38800,
}
ROOM_NAMES = ['金峰','笛吹','天神','天目','赤岳','風林']
ROOM_RANKS_BASE = {
    '金峰': ['E1','F1','G1','M1','F1','D1','B1','C1','E1','F1','G1','F1','E1','D1'],
    '笛吹': ['F1','G1','G1','L1','F1','E1','C1','D1','F1','G1','H1','G1','F1','E1'],
    '天神': ['E1','F1','G1','M1','F1','D1','B1','C1','E1','F1','G1','F1','E1','D1'],
    '天目': ['K1','L1','M1','R1','L1','K1','H1','I1','K1','L1','M1','L1','K1','J1'],
    '赤岳': ['D1','E1','F1','L1','E1','D1','B1','C1','D1','E1','F1','E1','D1','C1'],
    '風林': ['F1','G1','H1','N1','G1','E1','C1','D1','F1','G1','H1','G1','F1','E1'],
}

COMP_NAME_MAP = {
    '【2食】石和温泉　銘石の宿　かげつ':     'かげつ',
    '【2食】笛吹川温泉　坐忘':               '坐忘',
    '【2食】石和温泉　ホテルふじ':           'ホテルふじ',
    '【2食】石和温泉　華やぎの章　慶山':     '慶山',
    '【2食】石和温泉　ホテル古柏園':         '古柏園',
    '【2食】石和温泉　糸柳こやど　ゆわ':     '糸柳こやどゆわ',
    '【2食】川浦温泉　山県館':               '山県館',
    '【2食】糸柳別館　離れの邸　和穣苑':     '糸柳別館',
    '【2食】別邸　花水晶':                   '花水晶',
    '【2食】石和名湯館　糸柳':               '糸柳',
    '【2食】湯めぐり宿　笛吹川':             '笛吹川',
    '【2食】石和温泉郷　旅館深雪温泉':       '深雪温泉',
    '【2食】銘庭の宿　ホテル甲子園':         'ホテル甲子園（自社）',
}

MONTHLY_BUDGET = {
    '202404': 13_000_000, '202405': 14_000_000, '202406': 12_000_000,
    '202407': 18_000_000, '202408': 20_000_000, '202409': 16_000_000,
    '202410': 15_000_000, '202411': 13_000_000, '202412': 18_447_404,
    '202501': 17_200_984, '202502': 13_653_373, '202503': 18_618_641,
    '202504': 17_419_169, '202505': 15_498_589, '202506': 13_605_922,
    '202507': 18_000_000, '202508': 20_000_000, '202509': 16_000_000,
}

WDAYS = ['月','火','水','木','金','土','日']

# ============================================================
# ブッキングカーブ
# ============================================================
BOOKING_CURVE = {
    0: 1.000, 1: 0.953, 2: 0.938, 3: 0.826, 4: 0.812,
    5: 0.764, 6: 0.750, 7: 0.704, 8: 0.690, 9: 0.676,
    10: 0.637, 11: 0.625, 12: 0.613, 13: 0.601, 14: 0.540,
    15: 0.530, 16: 0.520, 17: 0.510, 18: 0.500, 19: 0.490,
    20: 0.480, 21: 0.407, 25: 0.360, 30: 0.270, 45: 0.138,
    60: 0.071, 90: 0.020,
}
_CURVE_KEYS = sorted(BOOKING_CURVE.keys())

TARGET_FINAL_OCC = {
    '土曜':      0.85,   # 土曜
    '金/祝前日/日': 0.75, # 金曜・祝前日・日曜
    '平日':      0.65,   # 月〜木・祝日当日
}

# ---- 祝日判定 ----
def _is_holiday(d):
    """日本の国民の祝日かどうか（振替休日を含む）"""
    mo, day, wd = d.month, d.day, d.weekday()

    # 固定祝日
    fixed = {
        (1,1),(1,2),(1,3),          # 元日・正月
        (2,11),(2,23),              # 建国記念日・天皇誕生日
        (4,29),                     # 昭和の日
        (5,3),(5,4),(5,5),          # 憲法・みどりの日・こどもの日
        (8,11),                     # 山の日
        (11,3),(11,23),             # 文化の日・勤労感謝の日
    }
    if (mo, day) in fixed:
        return True

    # ハッピーマンデー（月曜固定）
    if wd == 0:
        if mo == 1  and 8  <= day <= 14: return True   # 成人の日
        if mo == 7  and 15 <= day <= 21: return True   # 海の日
        if mo == 9  and 15 <= day <= 21: return True   # 敬老の日
        if mo == 10 and 8  <= day <= 14: return True   # スポーツの日

    # 春分の日（3/20 or 3/21）・秋分の日（9/22 or 9/23）近似
    if mo == 3 and day in (20, 21): return True
    if mo == 9 and day in (22, 23): return True

    # 振替休日：日曜が祝日 → 翌月曜が振替
    if wd == 0:
        prev = d - timedelta(days=1)
        pm, pd2 = prev.month, prev.day
        if (pm, pd2) in fixed: return True
        if pm == 3 and pd2 in (20, 21): return True
        if pm == 9 and pd2 in (22, 23): return True

    return False


def day_type(d):
    """
    曜日タイプを3段階で返す:
      '土曜'       : 土曜          → 目標稼働率 85%
      '金/祝前日/日': 金曜・祝前日・日曜 → 75%
      '平日'       : 月〜木・祝日当日   → 65%
    """
    wd       = d.weekday()
    tomorrow = d + timedelta(days=1)

    if wd == 5:               return '土曜'        # 土曜
    if wd == 6:               return '金/祝前日/日' # 日曜
    if wd == 4:               return '金/祝前日/日' # 金曜
    if _is_holiday(tomorrow): return '金/祝前日/日' # 祝前日（翌日が祝日）
    return '平日'

def suggest_rank(cur, action):
    if cur not in RANKS: return cur
    i = RANKS.index(cur)
    if action == 'UP'   and i < len(RANKS)-1: return RANKS[i+1]
    if action == 'DOWN' and i > 0:            return RANKS[i-1]
    return cur

def comp_avg_for_date(comp_prices, date_str, exclude='ホテル甲子園（自社）'):
    day_data = comp_prices.get(date_str, {})
    prices = [v for k, v in day_data.items() if k != exclude and v is not None]
    return round(sum(prices) / len(prices)) if prices else None

# ============================================================
# データ読み込み
# ============================================================
def _parse_pms(raw_bytes_or_path):
    """PMS CSVを読み込む。パスまたはBytesIOを受け取る"""
    if isinstance(raw_bytes_or_path, (str, bytes)) and not isinstance(raw_bytes_or_path, io.IOBase):
        # ファイルパス
        with open(raw_bytes_or_path, encoding='cp932') as f:
            raw = list(csv.DictReader(f))
    else:
        # BytesIO (Streamlitアップロード)
        text = raw_bytes_or_path.read().decode('cp932')
        raw = list(csv.DictReader(io.StringIO(text)))

    stay_rows = [r for r in raw
                 if r.get('利用有無','') == '有'
                 and '宿泊' in r.get('科目','')]

    daily_rooms = defaultdict(set)
    lead_list   = []
    seen_res    = set()

    for r in stay_rows:
        try:
            stay_date = datetime.strptime(r['利用日'], '%Y%m%d')
            nights    = int(r.get('泊数', 1) or 1)
            room      = r['宿泊部屋'].strip().lstrip('*')
            res_no    = r['予約番号'].strip()
            if not room or ',' in room: continue

            for n in range(nights):
                d = stay_date + timedelta(days=n)
                daily_rooms[d].add(room)

            res_key = f'{res_no}_{stay_date.strftime("%Y%m%d")}'
            if res_key not in seen_res:
                booked = datetime.strptime(r['予約日'], '%Y%m%d')
                ld = (stay_date - booked).days
                if 0 <= ld <= 180:
                    lead_list.append((stay_date, ld))
                seen_res.add(res_key)
        except: pass

    daily = {d: len(s) for d, s in daily_rooms.items()}

    lead_dist = defaultdict(lambda: defaultdict(int))
    for stay_date, ld in lead_list:
        dt = day_type(stay_date)
        lead_dist[dt][ld // 7] += 1

    monthly_rev = defaultdict(float)
    for r in raw:
        if r.get('利用有無', '') == '有':
            try:
                stay_date = datetime.strptime(r['利用日'], '%Y%m%d')
                month_key = stay_date.strftime('%Y%m')
                amount = float(r.get('金額', '0') or 0)
                monthly_rev[month_key] += amount
            except: pass

    # ---- 月次宿泊人数 ----
    def _get_persons(r):
        for key in ['大人人数', '大人数', '人数', 'M_ADL', '宿泊人数', '大人']:
            v = r.get(key, '')
            if v:
                try: return int(float(v))
                except: pass
        return 2  # PMS項目が取れない場合は2名として推定

    monthly_guests = defaultdict(int)
    seen_guest_res = set()
    for r in stay_rows:
        try:
            stay_date = datetime.strptime(r['利用日'], '%Y%m%d')
            res_no    = r.get('予約番号', '').strip()
            mk        = stay_date.strftime('%Y%m')
            gkey      = f'{res_no}_{mk}'
            if gkey not in seen_guest_res:
                monthly_guests[mk] += _get_persons(r)
                seen_guest_res.add(gkey)
        except: pass

    # ---- 部屋別月次集計 ----
    # nights: daily_rooms から集計（宿泊実績ベース）
    room_monthly_nights = defaultdict(lambda: defaultdict(int))
    for d, rooms_set in daily_rooms.items():
        mk = d.strftime('%Y%m')
        for room in rooms_set:
            room_monthly_nights[mk][room] += 1

    # revenue: stay_rows から重複なしで集計
    room_monthly_rev = defaultdict(lambda: defaultdict(float))
    seen_rev = set()
    for r in stay_rows:
        try:
            stay_date = datetime.strptime(r['利用日'], '%Y%m%d')
            room   = r['宿泊部屋'].strip().lstrip('*')
            res_no = r.get('予約番号', '').strip()
            if not room or ',' in room: continue
            mk     = stay_date.strftime('%Y%m')
            amount = float(r.get('金額', '0') or 0)
            rev_key = f'{res_no}_{stay_date.strftime("%Y%m%d")}_{room}'
            if rev_key not in seen_rev:
                room_monthly_rev[mk][room] += amount
                seen_rev.add(rev_key)
        except: pass

    # まとめる
    room_monthly = {}
    for mk in set(list(room_monthly_nights.keys()) + list(room_monthly_rev.keys())):
        all_rooms = set(list(room_monthly_nights.get(mk, {}).keys()) +
                        list(room_monthly_rev.get(mk, {}).keys()))
        room_monthly[mk] = {
            room: {
                'nights':  room_monthly_nights[mk].get(room, 0),
                'revenue': room_monthly_rev[mk].get(room, 0.0),
            }
            for room in all_rooms
        }

    return daily, lead_dist, dict(monthly_rev), room_monthly, dict(monthly_guests)


def _parse_rakutsuu(files_or_bytes):
    """ラクツウCSVを読み込む"""
    rows = {}
    if isinstance(files_or_bytes, list) and all(isinstance(f, str) for f in files_or_bytes):
        # ファイルパスリスト
        for fp in sorted(files_or_bytes):
            with open(fp, encoding='cp932') as f:
                for r in csv.DictReader(f):
                    k = r.get('予約番号','')
                    if k and k not in rows and r.get('区分','') != 'キャンセル':
                        rows[k] = r
    else:
        # BytesIOリスト
        for buf in (files_or_bytes if isinstance(files_or_bytes, list) else [files_or_bytes]):
            text = buf.read().decode('cp932')
            for r in csv.DictReader(io.StringIO(text)):
                k = r.get('予約番号','')
                if k and k not in rows and r.get('区分','') != 'キャンセル':
                    rows[k] = r

    daily = defaultdict(int)
    for r in rows.values():
        try:
            cin  = datetime.strptime(r['チェックイン'], '%Y%m%d')
            cout = datetime.strptime(r['チェックアウト'], '%Y%m%d')
            rms  = int(r.get('室数', 1))
            d = cin
            while d < cout:
                daily[d] += rms
                d += timedelta(days=1)
        except: pass

    lead_dist = defaultdict(lambda: defaultdict(int))
    for r in rows.values():
        try:
            cin = datetime.strptime(r['チェックイン'], '%Y%m%d')
            reg = datetime.strptime(r['受信日／登録日'][:8], '%Y%m%d')
            ld  = (cin - reg).days
            if 0 <= ld <= 180:
                dt = day_type(cin)
                lead_dist[dt][ld // 7] += 1
        except: pass

    return dict(daily), lead_dist, {}, {}, {}


def _read_comp_rows(source):
    """競合価格CSVの全行をリストで返す共通処理"""
    try:
        if isinstance(source, str):
            with open(source, encoding='utf-8-sig') as f:
                return list(csv.DictReader(f))
        else:
            text = source.read().decode('utf-8-sig')
            return list(csv.DictReader(io.StringIO(text)))
    except (FileNotFoundError, Exception):
        return []

def _parse_comp_prices(source):
    """競合価格CSVを読み込む（最新取得日の値を使用）"""
    rows = _read_comp_rows(source)
    # 取得日が最新のものだけ使う
    latest_date = max((r.get('取得日','') for r in rows), default='')
    comp_prices = defaultdict(dict)
    for r in rows:
        if r.get('取得日','') != latest_date:
            continue
        name      = r.get('施設名','').strip()
        short     = COMP_NAME_MAP.get(name, name)
        date      = r.get('対象日','').strip()
        price_str = r.get('最低価格','×').strip()
        price = None
        if price_str not in ('×', ''):
            try: price = int(price_str) // 2  # 2名合計 → 1人当たり
            except: pass
        comp_prices[date][short] = price
    return comp_prices

def load_comp_history(source=None):
    """
    競合価格の全履歴を返す。
    戻り値: list of dict
      [{取得日, 対象日, 施設名(短縮), 価格(int or None)}, ...]
    """
    if source:
        src = source
    elif os.path.exists(COMP_PRICES_CSV):
        src = COMP_PRICES_CSV
    else:
        src = COMP_PRICES_SAMPLE
    rows = _read_comp_rows(src)
    result = []
    for r in rows:
        name      = r.get('施設名','').strip()
        short     = COMP_NAME_MAP.get(name, name)
        price_str = r.get('最低価格','×').strip()
        price = None
        if price_str not in ('×', ''):
            try: price = int(price_str) // 2  # 2名合計 → 1人当たり
            except: pass
        result.append({
            '取得日': r.get('取得日','').strip(),
            '対象日': r.get('対象日','').strip(),
            '施設名': short,
            '価格':   price,
        })
    return result


def load_data(pms_file=None, rakutsuu_files=None, comp_file=None):
    """
    データを読み込む。引数なしならデフォルトパスを使用。
    pms_file: ファイルパス文字列 or BytesIO
    rakutsuu_files: ファイルパスリスト or BytesIOリスト
    comp_file: ファイルパス文字列 or BytesIO
    """
    # PMS or ラクツウ
    if pms_file is not None:
        daily, lead_dist, monthly_rev, room_monthly, monthly_guests = _parse_pms(pms_file)
        data_source = 'PMSデータ（アップロード）'
    elif rakutsuu_files is not None:
        daily, lead_dist, monthly_rev, room_monthly, monthly_guests = _parse_rakutsuu(rakutsuu_files)
        data_source = 'ラクツウCSV（アップロード）'
    elif os.path.exists(PMS_CSV):
        daily, lead_dist, monthly_rev, room_monthly, monthly_guests = _parse_pms(PMS_CSV)
        data_source = f'PMSデータ ({os.path.basename(PMS_CSV)})'
    else:
        files = glob.glob(f'{CSV_DIR}/ReserveList_*.csv')
        daily, lead_dist, monthly_rev, room_monthly, monthly_guests = _parse_rakutsuu(files)
        data_source = 'ラクツウCSV (ReserveList_*.csv)'

    # 競合価格（スクレイピング済み → サンプル の優先順）
    if comp_file is not None:
        comp_prices = _parse_comp_prices(comp_file)
    elif os.path.exists(COMP_PRICES_CSV):
        comp_prices = _parse_comp_prices(COMP_PRICES_CSV)
    elif os.path.exists(COMP_PRICES_SAMPLE):
        comp_prices = _parse_comp_prices(COMP_PRICES_SAMPLE)
    else:
        comp_prices = defaultdict(dict)

    return daily, lead_dist, comp_prices, data_source, monthly_rev, room_monthly, monthly_guests


# ============================================================
# 売上明細（科目別分析用）
# ============================================================

# 科目キーワードマッピング（実際の科目名に合わせて調整可能）
SALES_CAT_RULES = [
    ('宿泊',       lambda s: '宿泊' in s),
    ('昼休・日帰り', lambda s: '昼休' in s or '日帰' in s or '夜休' in s),
    ('ドリンク',    lambda s: any(k in s for k in [
        'ドリンク','飲料','ビール','ワイン','酒','サワー','ハイボール',
        'ウイスキー','カクテル','ジュース','コーヒー','お茶','ソフト','アルコール'
    ])),
    ('その他',     lambda s: True),   # fallback
]

def categorize_kamoku(kamoku):
    for cat, rule in SALES_CAT_RULES:
        if rule(kamoku):
            return cat
    return 'その他'


def load_sales_detail(pms_file=None):
    """
    売上明細を科目別に返す。
    戻り値: list of dict [{month, date_str, 科目, カテゴリ, 金額, 予約番号}]
    """
    src = pms_file or PMS_CSV
    try:
        if hasattr(src, 'read'):
            text = src.read().decode('cp932')
            raw = list(csv.DictReader(io.StringIO(text)))
        else:
            with open(src, encoding='cp932') as f:
                raw = list(csv.DictReader(f))
    except Exception:
        return []

    result = []
    for r in raw:
        if r.get('利用有無', '') != '有':
            continue
        try:
            date    = datetime.strptime(r['利用日'], '%Y%m%d')
            amount  = float(r.get('金額', '0') or 0)
            kamoku  = r.get('科目', '').strip()
            result.append({
                'month':    date.strftime('%Y%m'),
                'date_str': date.strftime('%Y/%m/%d'),
                '科目':     kamoku,
                'カテゴリ': categorize_kamoku(kamoku),
                '金額':     amount,
                '予約番号': r.get('予約番号', '').strip(),
            })
        except Exception:
            pass

    return result


# ============================================================
# 実績ブッキングカーブの計算
# ============================================================

def calc_actual_booking_curve(lead_dist):
    """
    _parse_pms が返す lead_dist から、日付タイプ別の実績ブッキングカーブを計算する。
    lead_dist[dt][week_bucket] = 予約件数
    戻り値: {'土曜': {...}, '金/祝前日/日': {...}, '平日': {...}}
    """
    # lead_dist のキーは day_type() の戻り値と一致
    merged = {'土曜': {}, '金/祝前日/日': {}, '平日': {}}
    for src_key, dst_key in [
        ('土曜',      '土曜'),
        ('金/祝前日/日', '金/祝前日/日'),
        ('平日',      '平日'),
        # 旧キー（既存CSVとの互換性）
        ('土/連休',   '土曜'),
        ('日曜',      '金/祝前日/日'),
    ]:
        for week_bucket, cnt in lead_dist.get(src_key, {}).items():
            merged[dst_key][week_bucket] = merged[dst_key].get(week_bucket, 0) + cnt

    result = {}
    for dtype, data in merged.items():
        total = sum(data.values())
        if total == 0:
            continue
        curve = {}
        for threshold_days in [0, 7, 14, 21, 30, 45, 60, 90]:
            week = threshold_days // 7
            # 指定リードタイム以上に予約した件数 = その時点ですでに入っていた予約
            count = sum(v for k, v in data.items() if k >= week)
            curve[threshold_days] = min(count / total, 1.0)
        result[dtype] = curve
    return result


def actual_curve_at(actual_curve, dtype, lead_days):
    """実績カーブから特定リードタイムの消化率を取得（線形補間）。"""
    curve = actual_curve.get(dtype) or actual_curve.get('平日')
    if not curve:
        return booking_curve_at(lead_days)  # fallback
    ld = max(0, lead_days)
    keys = sorted(curve.keys())
    if ld >= keys[-1]:
        return curve[keys[-1]]
    for i, k in enumerate(keys):
        if k >= ld:
            if k == ld or i == 0:
                return curve[k]
            k0, k1 = keys[i-1], k
            v0, v1 = curve[k0], curve[k1]
            return v0 + (v1 - v0) * (ld - k0) / (k1 - k0)
    return curve[keys[-1]]


# ============================================================
# 着地見込み計算（未来月）
# ============================================================

def calc_landing_forecast(daily, lead_dist, monthly_rev, room_monthly, today=None, months_ahead=6):
    """
    未来月の着地見込みを計算する。
    - 現状予約済み室数をブッキングカーブで割り戻して最終稼働率を推計
    - 昨年同月ADRを使って売上着地を推計

    戻り値: list of dict
    """
    import calendar as _cal
    today = today or TODAY

    actual_curve = calc_actual_booking_curve(lead_dist) if lead_dist else {}

    results = []

    for offset in range(0, months_ahead + 1):
        yr = today.year + (today.month - 1 + offset) // 12
        mo = (today.month - 1 + offset) % 12 + 1
        mk = f"{yr}{mo:02d}"
        days_in_month = _cal.monthrange(yr, mo)[1]
        avail = days_in_month * TOTAL_ROOMS

        # ── 現状室数（確定分）──────────────────────────
        # 今月: 過去分(確定) + 未来分(予約済み)
        # 未来月: 予約済みのみ
        past_nights   = sum(cnt for d, cnt in daily.items()
                            if d.year == yr and d.month == mo and d <= today)
        future_nights = sum(cnt for d, cnt in daily.items()
                            if d.year == yr and d.month == mo and d > today)
        cur_nights = past_nights + future_nights
        cur_occ    = cur_nights / avail if avail else 0

        # 残り日数
        future_days = [datetime(yr, mo, day)
                       for day in range(1, days_in_month + 1)
                       if datetime(yr, mo, day) > today]

        # 今月が完全に過去なら確定値として記録（見込み不要）
        is_past_month = (yr < today.year) or (yr == today.year and mo < today.month)
        is_completed  = not future_days and offset == 0

        # ── 着地見込み（ブッキングカーブ割り戻し）────────
        if future_days:
            forecast_nights = float(past_nights)  # 確定済みは固定
            for d in future_days:
                lead       = (d - today).days
                dtype      = day_type(d)
                curve_frac = actual_curve_at(actual_curve, dtype, lead)
                day_cur    = daily.get(d, 0)
                # カーブが低すぎる場合は最低5%で除算（過大推計を防ぐ）
                frac = max(curve_frac, 0.05)
                day_forecast = min(day_cur / frac, TOTAL_ROOMS)
                forecast_nights += day_forecast
            forecast_nights = min(round(forecast_nights), avail)
        else:
            # 未来日なし = 実績確定
            forecast_nights = cur_nights

        forecast_occ = forecast_nights / avail if avail else 0

        # ── 昨年同月実績 ──────────────────────────────
        prev_mk    = f"{yr-1}{mo:02d}"
        ly_rdata   = room_monthly.get(prev_mk, {})
        ly_nights  = sum(v['nights']  for v in ly_rdata.values())
        ly_rev     = monthly_rev.get(prev_mk, 0)
        ly_adr     = ly_rev / ly_nights if ly_nights else 0
        ly_occ     = ly_nights / avail  if avail     else 0

        # ── 売上着地見込み ────────────────────────────
        # 昨年ADR × 着地見込み室数
        if ly_adr and forecast_nights:
            forecast_rev = round(forecast_nights * ly_adr)
        else:
            forecast_rev = None

        results.append({
            'month':            mk,
            'label':            f"{yr}/{mo:02d}",
            'is_past':          is_past_month or is_completed,
            'avail':            avail,
            'past_nights':      past_nights,
            'cur_nights':       cur_nights,
            'cur_occ':          cur_occ,
            'forecast_nights':  forecast_nights,
            'forecast_occ':     forecast_occ,
            'forecast_rev':     forecast_rev,
            'last_year_nights': ly_nights,
            'last_year_occ':    ly_occ,
            'last_year_rev':    ly_rev,
            'budget':           MONTHLY_BUDGET.get(mk, 0),
            'remaining_days':   len(future_days),
        })

    return results


# ============================================================
# RM推奨計算（30日分）
# ============================================================
def calc_rm_rows(daily, comp_prices, today=None, days_ahead=None, lead_dist=None):
    """
    日別のRM推奨データを計算して返す。
    lead_dist が渡された場合は実績ブッキングカーブを使用する。
    戻り値: list of dict
    """
    today      = today or TODAY
    days_ahead = days_ahead or DAYS_AHEAD

    # 実績ブッキングカーブの計算
    actual_curve = calc_actual_booking_curve(lead_dist) if lead_dist else {}
    use_actual   = bool(actual_curve)

    rows = []

    for i in range(1, days_ahead + 1):
        d     = today + timedelta(days=i)
        dtype = day_type(d)
        lead  = i

        # 昨年同日の最終稼働率を取得
        try:
            d_ly = datetime(d.year - 1, d.month, d.day)
        except ValueError:
            d_ly = datetime(d.year - 1, d.month, 28)
        ly_final = daily.get(d_ly, 0) / TOTAL_ROOMS

        # 目標最終稼働率：昨年実績×1.1 を優先、データ不足時は固定値
        if ly_final > 0.01:
            final_occ = min(ly_final * 1.1, 1.0)
        else:
            final_occ = TARGET_FINAL_OCC.get(dtype, 0.65)

        # 目標消化率：実績カーブ or 固定カーブ × 目標最終稼働率
        if use_actual:
            curve_val = actual_curve_at(actual_curve, dtype, lead)
        else:
            curve_val = booking_curve_at(lead)
        tgt = final_occ * curve_val

        actual = daily.get(d, 0) / TOTAL_ROOMS
        note   = ''
        if daily.get(d, 0) == 0 and d > today:
            note = '予約なし'

        date_str  = d.strftime('%Y/%m/%d')
        cavg_real = comp_avg_for_date(comp_prices, date_str)
        cavg      = cavg_real if cavg_real is not None else None

        ri       = min(i - 1, len(ROOM_RANKS_BASE['金峰']) - 1)
        cur_rank = ROOM_RANKS_BASE['金峰'][ri]
        cur_p    = RANK_PRICE.get(cur_rank, 0)  # 1人当たり

        # ---- アクション判定（稼働率ベース）----
        diff = actual - tgt
        if diff >= 0.05:    action = 'UP'
        elif diff <= -0.10: action = 'DOWN'
        else:               action = 'STAY'

        # ---- 競合価格による補正（ハードブレーキのみ）----
        # 競合価格はアクションを「反転」させない。極端なケースのみ抑止。
        comp_reason = ''
        if cavg is not None and cur_p > 0:
            ratio = cur_p / cavg   # 1.0 = 競合平均と同額
            if action == 'DOWN' and ratio < 0.85:
                # 下げたいが既に競合より15%以上安い → これ以上下げない
                action = 'STAY'
                comp_reason = f'競合より安いため据え置き（{ratio:.0%}）'
            elif action == 'UP' and ratio > 1.40:
                # 上げたいが既に競合より40%以上高い → 過度な値上げ抑制
                action = 'STAY'
                comp_reason = f'競合より高すぎるため抑制（{ratio:.0%}）'

        sug_rank = suggest_rank(cur_rank, action)
        sug_p    = RANK_PRICE.get(sug_rank, 0)  # 1人当たり

        rows.append({
            'date':        d,
            'date_str':    d.strftime('%m/%d'),
            'wday':        WDAYS[d.weekday()],
            'lead':        lead,
            'dtype':       dtype,
            'target':      tgt,
            'actual':      actual,
            'diff':        diff,
            'cavg':        cavg,
            'cur_rank':    cur_rank,
            'cur_price':   cur_p,
            'action':      action,
            'sug_rank':    sug_rank,
            'sug_price':   sug_p,
            'note':        note,
            'comp_reason': comp_reason,
            'curve_src':   '実績' if use_actual else '固定',
            'ly_final':    ly_final,
            'final_occ':   final_occ,
        })

    return rows


# ============================================================
# RM設定スナップショット（履歴ログ）
# ============================================================
SNAPSHOT_CSV = r"C:\Users\tsukamoto.seishu\rm_system\rm_snapshot.csv"
SNAPSHOT_COLS = ['保存日', '対象日', '曜日', 'リードタイム',
                 '推奨ランク', '推奨価格', '実績消化率', '目標消化率', 'アクション', '競合平均']

def save_snapshot(rows, saved_date=None, path=None):
    """
    calc_rm_rows の結果をスナップショットCSVに追記する。
    同じ保存日のデータが既にあれば上書き（重複防止）。
    """
    path = path or SNAPSHOT_CSV
    saved_date = saved_date or datetime.now().strftime('%Y/%m/%d')

    # 既存データ読み込み
    existing = []
    if os.path.exists(path):
        with open(path, encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            existing = [r for r in reader if r.get('保存日') != saved_date]

    # 新しい行
    new_rows = []
    for r in rows:
        new_rows.append({
            '保存日':      saved_date,
            '対象日':      r['date'].strftime('%Y/%m/%d'),
            '曜日':        r['wday'],
            'リードタイム': r['lead'],
            '推奨ランク':  r['sug_rank'],
            '推奨価格':    r['sug_price'],
            '実績消化率':  f"{r['actual']:.3f}",
            '目標消化率':  f"{r['target']:.3f}",
            'アクション':  r['action'],
            '競合平均':    r['cavg'] if r['cavg'] else '',
        })

    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=SNAPSHOT_COLS)
        writer.writeheader()
        writer.writerows(existing + new_rows)

    return len(new_rows)


def load_snapshot(path=None):
    """スナップショットCSVを読み込んでリスト of dict で返す"""
    path = path or SNAPSHOT_CSV
    if not os.path.exists(path):
        return []
    with open(path, encoding='utf-8-sig') as f:
        return list(csv.DictReader(f))
