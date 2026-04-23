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
TODAY       = datetime(2026, 4, 9)
DAYS_AHEAD  = 30
TOTAL_ROOMS = 17

_BASE_DIR       = os.path.dirname(os.path.abspath(__file__))
CSV_DIR         = os.path.join(_BASE_DIR)
PMS_CSV         = os.path.join(_BASE_DIR, "a.csv")
COMP_PRICES_CSV = os.path.join(_BASE_DIR, "competitor_prices_sample.csv")

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

TARGET_FINAL_OCC = {'土/連休初日': 0.85, '平日/日': 0.65}

def booking_curve_at(lead_days):
    ld = max(0, lead_days)
    if ld >= _CURVE_KEYS[-1]:
        return BOOKING_CURVE[_CURVE_KEYS[-1]]
    for i, k in enumerate(_CURVE_KEYS):
        if k >= ld:
            if k == ld or i == 0:
                return BOOKING_CURVE[k]
            k0, k1 = _CURVE_KEYS[i-1], k
            v0, v1 = BOOKING_CURVE[k0], BOOKING_CURVE[k1]
            return v0 + (v1 - v0) * (ld - k0) / (k1 - k0)
    return BOOKING_CURVE[_CURVE_KEYS[-1]]

def target_occ(lead, dtype):
    final = TARGET_FINAL_OCC.get(dtype, 0.65)
    return final * booking_curve_at(lead)

def day_type(d):
    wd = d.weekday()
    hols = {(1,1),(1,2),(1,3),(4,29),(5,3),(5,4),(5,5),(7,21),(8,11),(9,15),(10,13),(11,3),(11,23),(12,23)}
    if wd == 5 or (d.month, d.day) in hols: return '土/連休初日'
    return '平日/日'

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
        wd = stay_date.weekday()
        dt = '土/連休' if wd == 5 else ('日曜' if wd == 6 else '平日')
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

    return daily, lead_dist, dict(monthly_rev)


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
                wd = cin.weekday()
                dt = '土/連休' if wd == 5 else ('日曜' if wd == 6 else '平日')
                lead_dist[dt][ld // 7] += 1
        except: pass

    return dict(daily), lead_dist, {}


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
            try: price = int(price_str)
            except: pass
        comp_prices[date][short] = price
    return comp_prices

def load_comp_history(source=None):
    """
    競合価格の全履歴を返す。
    戻り値: list of dict
      [{取得日, 対象日, 施設名(短縮), 価格(int or None)}, ...]
    """
    src = source or COMP_PRICES_CSV
    rows = _read_comp_rows(src)
    result = []
    for r in rows:
        name      = r.get('施設名','').strip()
        short     = COMP_NAME_MAP.get(name, name)
        price_str = r.get('最低価格','×').strip()
        price = None
        if price_str not in ('×', ''):
            try: price = int(price_str)
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
        daily, lead_dist, monthly_rev = _parse_pms(pms_file)
        data_source = 'PMSデータ（アップロード）'
    elif rakutsuu_files is not None:
        daily, lead_dist, monthly_rev = _parse_rakutsuu(rakutsuu_files)
        data_source = 'ラクツウCSV（アップロード）'
    elif os.path.exists(PMS_CSV):
        daily, lead_dist, monthly_rev = _parse_pms(PMS_CSV)
        data_source = f'PMSデータ ({os.path.basename(PMS_CSV)})'
    else:
        # データなし（クラウド環境など）→ 空データで起動
        daily, lead_dist, monthly_rev = defaultdict(int), defaultdict(lambda: defaultdict(int)), {}
        data_source = 'データ未アップロード'

    # 競合価格
    if comp_file is not None:
        comp_prices = _parse_comp_prices(comp_file)
    elif os.path.exists(COMP_PRICES_CSV):
        comp_prices = _parse_comp_prices(COMP_PRICES_CSV)
    else:
        comp_prices = defaultdict(dict)

    return daily, lead_dist, comp_prices, data_source, monthly_rev


# ============================================================
# RM推奨計算（30日分）
# ============================================================
def calc_rm_rows(daily, comp_prices, today=None, days_ahead=None):
    """
    日別のRM推奨データを計算して返す
    戻り値: list of dict
    """
    today      = today or TODAY
    days_ahead = days_ahead or DAYS_AHEAD
    rows = []

    for i in range(1, days_ahead + 1):
        d     = today + timedelta(days=i)
        dtype = day_type(d)
        lead  = i

        tgt    = target_occ(lead, dtype)
        actual = daily.get(d, 0) / TOTAL_ROOMS
        note   = ''
        if daily.get(d, 0) == 0 and d > today:
            note = '予約なし'

        date_str  = d.strftime('%Y/%m/%d')
        cavg_real = comp_avg_for_date(comp_prices, date_str)
        cavg      = cavg_real if cavg_real is not None else None

        ri       = min(i - 1, len(ROOM_RANKS_BASE['金峰']) - 1)
        cur_rank = ROOM_RANKS_BASE['金峰'][ri]
        cur_p    = RANK_PRICE.get(cur_rank, 0) * 2

        diff = actual - tgt
        if diff >= 0.05:    action = 'UP'
        elif diff <= -0.10: action = 'DOWN'
        else:               action = 'STAY'

        sug_rank = suggest_rank(cur_rank, action)
        sug_p    = RANK_PRICE.get(sug_rank, 0) * 2

        rows.append({
            'date':      d,
            'date_str':  d.strftime('%m/%d'),
            'wday':      WDAYS[d.weekday()],
            'lead':      lead,
            'dtype':     dtype,
            'target':    tgt,
            'actual':    actual,
            'diff':      diff,
            'cavg':      cavg,
            'cur_rank':  cur_rank,
            'cur_price': cur_p,
            'action':    action,
            'sug_rank':  sug_rank,
            'sug_price': sug_p,
            'note':      note,
        })

    return rows


# ============================================================
# RM設定スナップショット（履歴ログ）
# ============================================================
SNAPSHOT_CSV = os.path.join(_BASE_DIR, "rm_snapshot.csv")
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
