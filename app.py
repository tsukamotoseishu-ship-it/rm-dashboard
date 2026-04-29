"""
RM価格案 Webダッシュボード - ホテル甲子園
Streamlit アプリ本体
実行: streamlit run app.py
"""

import io
import calendar
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from datetime import datetime, timedelta
from collections import defaultdict

import rm_logic as rm

# ============================================================
# ページ設定
# ============================================================
st.set_page_config(
    page_title="RM ダッシュボード｜ホテル甲子園",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded",
)

# カスタムCSS
st.markdown("""
<style>
    .metric-card { background:#f8f9fa; border-radius:8px; padding:16px; text-align:center; }
    .action-up   { color:#E74C3C; font-weight:bold; }
    .action-stay { color:#27AE60; font-weight:bold; }
    .action-down { color:#2980b9; font-weight:bold; }
    thead tr th  { background:#1a3a5c !important; color:white !important; }
    .stDataFrame { font-size:13px; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# サイドバー：ファイルアップロード
# ============================================================
with st.sidebar:
    st.title("📂 データ入力")

    pms_file = st.file_uploader(
        "PMSデータ（a.csv）",
        type=["csv"],
        help="PMSから出力したa.csvをアップロード",
    )

    comp_file = st.file_uploader(
        "競合価格CSV（competitor_prices.csv）",
        type=["csv"],
        help="competitor_prices.csv をアップロード",
    )

    if st.button("🔍 競合価格を今すぐ取得"):
        with st.spinner("楽天APIから競合価格を取得中（約3分）..."):
            try:
                import competitor_scraper
                competitor_scraper.run()
                st.success("✅ 競合価格を取得しました！　ページをリロードしてください。")
                st.cache_data.clear()
            except Exception as e:
                st.error(f"取得エラー: {e}")

    if pms_file:
        st.success(f"✅ PMSデータ: {pms_file.name}")
    else:
        st.info("PMSデータ未アップロード\n（ローカルの a.csv を使用）")

    if comp_file:
        st.success(f"✅ 競合価格: {comp_file.name}")
    else:
        st.info("競合価格CSV未アップロード\n（ローカルファイルを使用）")

    st.caption(f"最終更新: {datetime.now().strftime('%Y/%m/%d %H:%M')}")

# ============================================================
# データ読み込み（キャッシュ）
# ============================================================
@st.cache_data(show_spinner="データを読み込み中...")
def load_cached(pms_bytes, comp_bytes):
    """ファイルの内容（bytes）をキーにキャッシュ"""
    pms_src  = io.BytesIO(pms_bytes)  if pms_bytes  else None
    comp_src = io.BytesIO(comp_bytes) if comp_bytes  else None

    daily, lead_dist, comp_prices, data_source, monthly_rev, room_monthly, monthly_guests = rm.load_data(
        pms_file=pms_src, comp_file=comp_src
    )
    rows = rm.calc_rm_rows(daily, comp_prices, lead_dist=lead_dist)
    daily         = dict(daily)
    lead_dist     = {k: dict(v) for k, v in lead_dist.items()}
    comp_prices   = {k: dict(v) for k, v in comp_prices.items()}
    monthly_rev   = dict(monthly_rev)
    monthly_guests= dict(monthly_guests)
    return daily, lead_dist, comp_prices, data_source, monthly_rev, rows, room_monthly, monthly_guests


pms_bytes  = pms_file.read()  if pms_file  else None
comp_bytes = comp_file.read() if comp_file else None

try:
    daily, lead_dist, comp_prices, data_source, monthly_rev, rows, room_monthly, monthly_guests = load_cached(
        pms_bytes, comp_bytes
    )
except Exception as e:
    st.error(f"データ読み込みエラー: {e}")
    st.stop()

@st.cache_data(show_spinner="売上明細を読み込み中...")
def load_sales_cached(pms_bytes):
    src = io.BytesIO(pms_bytes) if pms_bytes else None
    return rm.load_sales_detail(src)

sales_detail = load_sales_cached(pms_bytes)

# ============================================================
# 5タブ
# ============================================================
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📋 RM価格案",
    "📊 実績稼働率",
    "📈 ブッキングカーブ",
    "💴 競合価格",
    "💰 月次売上",
    "🍺 売上内訳",
])

# ------------------------------------------------------------------
# Tab 1: RM価格案
# ------------------------------------------------------------------
with tab1:
    st.header("RM推奨価格案（30日間）")
    st.caption(f"データソース: {data_source}")

    # ---- KPI カード ----
    today_rows = [r for r in rows if r['lead'] <= 7]
    avg_actual = sum(r['actual'] for r in today_rows) / len(today_rows) if today_rows else 0
    avg_target = sum(r['target'] for r in today_rows) / len(today_rows) if today_rows else 0
    up_count   = sum(1 for r in rows if r['action'] == 'UP')
    dn_count   = sum(1 for r in rows if r['action'] == 'DOWN')

    # 競合差額（直近7日平均）
    comp_diffs = [r['cur_price'] - r['cavg']
                  for r in today_rows if r['cavg'] is not None]
    comp_diff_avg = round(sum(comp_diffs) / len(comp_diffs)) if comp_diffs else None

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("直近7日　実績消化率", f"{avg_actual:.1%}", f"目標比 {avg_actual - avg_target:+.1%}")
    col2.metric("UP推奨日数", f"{up_count}日")
    col3.metric("DOWN推奨日数", f"{dn_count}日")
    if comp_diff_avg is not None:
        col4.metric("自社vs競合平均（直近7日）", f"¥{comp_diff_avg:+,}")
    else:
        col4.metric("自社vs競合平均", "データなし")

    st.divider()

    # ---- 折れ線グラフ: 期待進捗率 vs 実績消化率 ----
    df_curve = pd.DataFrame({
        '日付':     [r['date_str'] for r in rows],
        '期待進捗': [r['target']   for r in rows],
        '実績消化': [r['actual']   for r in rows],
    })
    fig_curve = px.line(
        df_curve.melt('日付', var_name='種別', value_name='消化率'),
        x='日付', y='消化率', color='種別',
        color_discrete_map={'期待進捗': '#2980b9', '実績消化': '#E74C3C'},
        title='期待進捗率 vs 実績消化率（30日間）',
        labels={'消化率': '消化率', '日付': '日付'},
    )
    fig_curve.update_layout(yaxis_tickformat='.0%', height=300)
    st.plotly_chart(fig_curve, use_container_width=True)

    # ---- 推奨アクション表 ----
    def action_label(a):
        icons = {'UP': '⬆ UP', 'STAY': '✔ STAY', 'DOWN': '⬇ DOWN'}
        return icons.get(a, a)

    def row_color(a):
        return {'UP': '#FFF5F5', 'DOWN': '#F0F8FF', 'STAY': '#F0FFF4'}.get(a, '#FFFFFF')

    curve_src = rows[0]['curve_src'] if rows else '固定'
    st.caption(f"ブッキングカーブ: **{curve_src}データ使用**　競合価格: {'あり（判定に反映）' if any(r['cavg'] for r in rows) else 'なし'}")

    # ---- 表示切替 ----
    view_mode = st.radio(
        "表示形式",
        ["📋 テーブル表示", "📅 カレンダー表示"],
        horizontal=True,
        label_visibility="collapsed",
    )

    if view_mode == "📋 テーブル表示":
        df_rm = pd.DataFrame([{
            '日付':             r['date_str'],
            '曜日':             r['wday'],
            'LT(日)':          r['lead'],
            '昨年最終稼働':     f"{r.get('ly_final', 0):.1%}" if r.get('ly_final', 0) > 0.01 else '—',
            '目標最終稼働':     f"{r.get('final_occ', 0):.1%}",
            '目標消化':         f"{r['target']:.1%}",
            '実績消化':         f"{r['actual']:.1%}",
            '差異':             f"{r['diff']:+.1%}",
            'アクション':       action_label(r['action']),
            '現在ランク':       r['cur_rank'],
            '現在価格':         f"¥{r['cur_price']:,}",
            '推奨ランク':       r['sug_rank'],
            '推奨価格':         f"¥{r['sug_price']:,}",
            '競合平均':         f"¥{r['cavg']:,}" if r['cavg'] else '—',
            '競合補正':         r['comp_reason'] or '—',
        } for r in rows])

        st.dataframe(
            df_rm,
            use_container_width=True,
            hide_index=True,
            height=600,
        )

        csv_bytes = df_rm.to_csv(index=False).encode('utf-8-sig')
        st.download_button("⬇ CSV ダウンロード", csv_bytes, "rm_plan.csv", "text/csv")

    else:
        # ---- カレンダービュー ----
        # rows を日付キーの dict に変換
        rows_by_date = {r['date'].date(): r for r in rows}
        today = datetime.now().date()

        # 対象月リストを収集
        months_in_view = sorted(set((r['date'].year, r['date'].month) for r in rows))

        ACTION_BG   = {'UP': '#FDECEA', 'STAY': '#E8F5E9', 'DOWN': '#E3F2FD'}
        ACTION_FG   = {'UP': '#C0392B', 'STAY': '#1E8449', 'DOWN': '#1565C0'}
        ACTION_ICON = {'UP': '⬆', 'STAY': '✔', 'DOWN': '⬇'}
        WDAY_JP     = ['月', '火', '水', '木', '金', '土', '日']

        cal_html_parts = []

        for (yr, mo) in months_in_view:
            # 月の1日の曜日インデックス（月=0）
            import calendar
            first_weekday, days_in_month = calendar.monthrange(yr, mo)

            # ヘッダー
            cal_html_parts.append(f"""
            <div style="margin-bottom:24px;">
            <div style="font-size:16px;font-weight:bold;color:#1a3a5c;margin-bottom:8px;">
                {yr}年{mo}月
            </div>
            <table style="border-collapse:collapse;width:100%;table-layout:fixed;">
            <colgroup>{"<col>" * 7}</colgroup>
            <thead>
            <tr>
            """)
            for i, w in enumerate(WDAY_JP):
                if i == 5:   # 土
                    color = '#1565C0'
                elif i == 6: # 日
                    color = '#C0392B'
                else:
                    color = '#555'
                cal_html_parts.append(
                    f'<th style="text-align:center;padding:4px;font-size:12px;'
                    f'color:{color};border-bottom:2px solid #1a3a5c;">{w}</th>'
                )
            cal_html_parts.append('</tr></thead><tbody>')

            # セルを生成
            day = 1
            cell_index = 0  # 0=月 〜 6=日
            # 最初の行：月曜始まりで空白埋め
            cal_html_parts.append('<tr>')
            for blank in range(first_weekday):
                cal_html_parts.append('<td style="border:1px solid #e0e0e0;padding:4px;min-height:80px;background:#fafafa;"></td>')
                cell_index += 1

            while day <= days_in_month:
                if cell_index % 7 == 0 and cell_index > 0:
                    cal_html_parts.append('</tr><tr>')

                weekday = (first_weekday + day - 1) % 7  # 0=月〜6=日
                date_obj = datetime(yr, mo, day).date()
                is_today = (date_obj == today)
                is_sat   = (weekday == 5)
                is_sun   = (weekday == 6)

                # セル背景
                if is_today:
                    cell_bg = '#FFF9C4'
                elif is_sat:
                    cell_bg = '#EEF4FF'
                elif is_sun:
                    cell_bg = '#FFF0F0'
                else:
                    cell_bg = '#FFFFFF'

                # 日付ラベル
                day_color = '#1565C0' if is_sat else ('#C0392B' if is_sun else '#333')
                today_badge = ' <span style="background:#FF6B35;color:white;border-radius:3px;font-size:9px;padding:1px 4px;">今日</span>' if is_today else ''
                day_label = f'<div style="font-size:13px;font-weight:bold;color:{day_color};margin-bottom:4px;">{day}{today_badge}</div>'

                # RMデータ
                r = rows_by_date.get(date_obj)
                if r:
                    action  = r['action']
                    ab      = ACTION_BG.get(action, '#FFFFFF')
                    af      = ACTION_FG.get(action, '#333')
                    ai      = ACTION_ICON.get(action, '')
                    action_badge = (
                        f'<div style="background:{ab};color:{af};border-radius:4px;'
                        f'font-size:11px;font-weight:bold;text-align:center;padding:2px 4px;margin-bottom:3px;">'
                        f'{ai} {action}</div>'
                    )
                    # 現在ランク → 推奨ランク
                    cur_r = r.get('cur_rank', '')
                    sug_r = r.get('sug_rank', '')
                    arrow = ' → ' if cur_r != sug_r else ' = '
                    rank_color = af if cur_r != sug_r else '#666'
                    rank_line = (
                        f'<div style="font-size:10px;color:{rank_color};text-align:center;'
                        f'font-weight:bold;margin-bottom:2px;">'
                        f'{cur_r}{arrow}{sug_r}</div>'
                    )
                    price_line = f'<div style="font-size:11px;color:#333;text-align:center;">¥{r["sug_price"]:,}</div>'
                    occ_color  = '#C0392B' if r['actual'] >= r['target'] else '#1565C0'
                    occ_line   = (
                        f'<div style="font-size:10px;color:{occ_color};text-align:center;margin-top:2px;">'
                        f'{r["actual"]:.0%}</div>'
                    )
                    cell_content = action_badge + rank_line + price_line + occ_line
                else:
                    cell_content = ''

                cal_html_parts.append(
                    f'<td style="border:1px solid #e0e0e0;padding:6px;vertical-align:top;'
                    f'background:{cell_bg};min-height:80px;">'
                    f'{day_label}{cell_content}</td>'
                )

                day += 1
                cell_index += 1

            # 最終行の残りを空白で埋める
            remaining = 7 - (cell_index % 7)
            if remaining < 7:
                for _ in range(remaining):
                    cal_html_parts.append('<td style="border:1px solid #e0e0e0;padding:4px;background:#fafafa;"></td>')
            cal_html_parts.append('</tr></tbody></table></div>')

        # 凡例
        legend_html = """
        <div style="display:flex;gap:16px;margin-bottom:16px;font-size:12px;flex-wrap:wrap;">
            <span style="background:#FDECEA;color:#C0392B;padding:3px 10px;border-radius:4px;font-weight:bold;">⬆ UP　価格引き上げ推奨</span>
            <span style="background:#E8F5E9;color:#1E8449;padding:3px 10px;border-radius:4px;font-weight:bold;">✔ STAY　現状維持</span>
            <span style="background:#E3F2FD;color:#1565C0;padding:3px 10px;border-radius:4px;font-weight:bold;">⬇ DOWN　価格引き下げ推奨</span>
            <span style="margin-left:16px;color:#555;">各セル：現在ランク→推奨ランク／推奨価格（1人）／実績消化率</span>
        </div>
        """
        st.html(legend_html + ''.join(cal_html_parts))

# ------------------------------------------------------------------
# Tab 2: 実績稼働率
# ------------------------------------------------------------------
with tab2:
    st.header("実績稼働率")

    if not daily or not room_monthly:
        st.info("PMSデータをアップロードすると実績が表示されます。")
    else:
        import calendar

        all_months_rm = sorted(room_monthly.keys(), reverse=True)
        sel_month = st.selectbox("月を選択", all_months_rm,
                                 format_func=lambda m: f"{m[:4]}/{m[4:]}", key='tab2_month')
        prev_year_month = str(int(sel_month[:4]) - 1) + sel_month[4:]

        yr, mo = int(sel_month[:4]), int(sel_month[4:])
        days_in_month = calendar.monthrange(yr, mo)[1]

        cur_data  = room_monthly.get(sel_month, {})
        prev_data = room_monthly.get(prev_year_month, {})

        # 全部屋リスト（今月 + 昨年同月の和集合）
        all_rooms = sorted(set(list(cur_data.keys()) + list(prev_data.keys())))

        # 合計行
        total_nights  = sum(v['nights']  for v in cur_data.values())
        total_revenue = sum(v['revenue'] for v in cur_data.values())
        total_occ     = total_nights / (days_in_month * rm.TOTAL_ROOMS)
        total_tanka   = total_revenue / total_nights if total_nights else 0

        prev_total_nights  = sum(v['nights']  for v in prev_data.values())
        prev_total_revenue = sum(v['revenue'] for v in prev_data.values())

        def yoy_pct(cur, prev):
            if prev and prev > 0:
                return cur / prev
            return None

        def bg_yoy(ratio):
            if ratio is None: return '#F5F5F5'
            if ratio >= 1.10: return '#27AE60'
            if ratio >= 1.00: return '#A9DFBF'
            if ratio >= 0.90: return '#F9EBEA'
            return '#E74C3C'

        def fmt_yoy(ratio):
            return f"{ratio:.1%}" if ratio is not None else '—'

        # ---- HTML テーブル ----
        th = 'padding:5px 10px;background:#2c3e50;color:white;text-align:right;white-space:nowrap;border:1px solid #4a4a4a;font-size:12px;'
        th_l = th.replace('text-align:right', 'text-align:left')
        td_style = 'padding:4px 10px;text-align:right;border:1px solid #ddd;font-size:12px;white-space:nowrap;'
        td_l = td_style.replace('text-align:right', 'text-align:left')
        td_total = td_style + 'font-weight:bold;background:#EBF5FB;'
        td_total_l = td_l + 'font-weight:bold;background:#EBF5FB;'

        def make_row(label, nights, revenue, room_count=1, is_total=False, prev_nights=None, prev_revenue=None):
            # 稼働率 = 販売室数 ÷ (月日数 × 室数)
            avail = days_in_month * room_count
            occ   = nights / avail if avail else 0
            tanka = revenue / nights if nights else 0
            ts  = td_total   if is_total else td_style
            tsl = td_total_l if is_total else td_l

            # 昨対
            n_yoy = yoy_pct(nights, prev_nights)
            r_yoy = yoy_pct(revenue, prev_revenue)

            row = (
                f'<td style="{tsl}">{label}</td>'
                f'<td style="{ts}">{revenue:,.0f}</td>'
                f'<td style="{ts}">{nights}</td>'
                f'<td style="{ts}">{occ:.1%}</td>'
                f'<td style="{ts}">{tanka:,.0f}</td>'
                f'<td style="{ts}background:{bg_yoy(r_yoy)};color:{"white" if r_yoy is not None and (r_yoy >= 1.1 or r_yoy < 0.9) else "inherit"}">{fmt_yoy(r_yoy)}</td>'
                f'<td style="{ts}background:{bg_yoy(n_yoy)};color:{"white" if n_yoy is not None and (n_yoy >= 1.1 or n_yoy < 0.9) else "inherit"}">{fmt_yoy(n_yoy)}</td>'
            )
            return f'<tr>{row}</tr>'

        header = (
            f'<tr>'
            f'<th style="{th_l}">部屋タイプ</th>'
            f'<th style="{th}">売上</th>'
            f'<th style="{th}">販売室数</th>'
            f'<th style="{th}">稼働率</th>'
            f'<th style="{th}">室単価</th>'
            f'<th style="{th}">売上<br>昨対</th>'
            f'<th style="{th}">室数<br>昨対</th>'
            f'</tr>'
        )

        rows_html = [header]
        # 合計行：TOTAL_ROOMS で割る
        rows_html.append(make_row('合計', total_nights, total_revenue,
                                  room_count=rm.TOTAL_ROOMS, is_total=True,
                                  prev_nights=prev_total_nights, prev_revenue=prev_total_revenue))
        # 部屋別行：各タイプ1室として計算
        for room in all_rooms:
            d = cur_data.get(room, {'nights': 0, 'revenue': 0.0})
            p = prev_data.get(room, {'nights': 0, 'revenue': 0.0})
            rows_html.append(make_row(room, d['nights'], d['revenue'],
                                      room_count=1,
                                      prev_nights=p['nights'] or None,
                                      prev_revenue=p['revenue'] or None))

        html = f"""
        <div style="overflow-x:auto;border:1px solid #ccc;border-radius:4px;">
        <table style="border-collapse:collapse;width:100%;min-width:600px;">
        <thead>{header}</thead>
        <tbody>{''.join(rows_html[1:])}</tbody>
        </table></div>
        <p style="font-size:11px;color:#888;margin-top:4px;">
        合計稼働率 = 販売室数 ÷ (月日数 × {rm.TOTAL_ROOMS}室)。部屋別稼働率 = 販売室数 ÷ 月日数（1室換算）。
        </p>
        """
        st.subheader(f"部屋別実績　{sel_month[:4]}/{sel_month[4:]}　（前年同月比較付き）")
        st.html(html)

        st.divider()

        # 月別稼働率トレンド（全体）
        df_occ = pd.DataFrame([
            {'年月': m, '稼働率': sum(v['nights'] for v in data.values()) /
                                  (calendar.monthrange(int(m[:4]), int(m[4:]))[1] * rm.TOTAL_ROOMS)}
            for m, data in sorted(room_monthly.items())
        ])
        fig_bar = px.bar(df_occ, x='年月', y='稼働率',
                         color='稼働率', color_continuous_scale='RdYlGn',
                         range_color=[0, 1], title='月別稼働率（全体）')
        fig_bar.update_layout(yaxis_tickformat='.0%', height=280, showlegend=False)
        st.plotly_chart(fig_bar, use_container_width=True)

# ------------------------------------------------------------------
# Tab 3: ブッキングカーブ
# ------------------------------------------------------------------
with tab3:
    st.header("週別入れ込み状況")

    today_dt = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    st.caption(
        "今日時点でPMSに入っている予約室数・稼働率を週ごとに集計。"
        "昨年同時点の入れ込みと比較して進捗を確認できます。"
    )

    # ── 週別データを集計 ────────────────────────────────────────────
    # 今日から8週先まで、1週ずつ集計
    WEEKS_AHEAD = 9   # 何週先まで表示するか

    week_rows = []
    for w in range(0, WEEKS_AHEAD):
        week_start = today_dt + timedelta(days=w * 7)
        week_end   = week_start + timedelta(days=6)
        label      = f"{week_start.strftime('%m/%d')}〜{week_end.strftime('%m/%d')}"
        lead_mid   = w * 7 + 3   # 週の中間リードタイム（代表値）

        # 今年の入れ込み室数
        cur_rooms = sum(daily.get(week_start + timedelta(days=d), 0) for d in range(7))
        cur_occ   = cur_rooms / (7 * rm.TOTAL_ROOMS)

        # 昨年同週の最終実績
        ly_rooms_final = sum(
            daily.get(week_start + timedelta(days=d - 365), 0) for d in range(7)
        )
        # ±1日ずれ対策：前後1日も考慮して昨年同週を特定
        ly_rooms_final2 = 0
        for d in range(7):
            arr = week_start + timedelta(days=d)
            try:
                arr_ly = datetime(arr.year - 1, arr.month, arr.day)
            except ValueError:
                arr_ly = datetime(arr.year - 1, arr.month, 28)
            ly_rooms_final2 += daily.get(arr_ly, 0)
        ly_occ_final = ly_rooms_final2 / (7 * rm.TOTAL_ROOMS)

        # 昨年同時点（同じリードタイム）での入れ込み
        # 昨年同週の同時点 = 昨年同週の最終 × 当時のカーブ進捗率
        actual_curve_t3 = rm.calc_actual_booking_curve(lead_dist)
        # 週の曜日構成で加重平均カーブ
        curve_fracs = []
        for d in range(7):
            arr   = week_start + timedelta(days=d)
            dtype = rm.day_type(arr)
            lead  = (arr - today_dt).days
            curve_fracs.append(rm.actual_curve_at(actual_curve_t3, dtype, lead))
        avg_curve = sum(curve_fracs) / len(curve_fracs) if curve_fracs else 0

        ly_rooms_same_point = round(ly_rooms_final2 * avg_curve)
        ly_occ_same_point   = ly_rooms_same_point / (7 * rm.TOTAL_ROOMS)

        # 昨対比（今年現状 vs 昨年同時点）
        yoy_ratio = cur_rooms / ly_rooms_same_point if ly_rooms_same_point > 0 else None

        week_rows.append({
            'label':             label,
            'week_start':        week_start,
            'cur_rooms':         cur_rooms,
            'cur_occ':           cur_occ,
            'ly_rooms_final':    ly_rooms_final2,
            'ly_occ_final':      ly_occ_final,
            'ly_rooms_same_pt':  ly_rooms_same_point,
            'ly_occ_same_pt':    ly_occ_same_point,
            'yoy_ratio':         yoy_ratio,
            'avail':             7 * rm.TOTAL_ROOMS,
        })

    # ── ① 週別入れ込みグラフ ─────────────────────────────────────────
    labels         = [r['label']           for r in week_rows]
    cur_rooms_list = [r['cur_rooms']       for r in week_rows]
    ly_final_list  = [r['ly_rooms_final']  for r in week_rows]
    ly_same_list   = [r['ly_rooms_same_pt'] for r in week_rows]

    # 稼働率 or 室数の切替
    show_occ = st.toggle("室数 → 稼働率で表示", value=False, key='bc_occ_toggle')

    if show_occ:
        cur_vals  = [r['cur_occ']        for r in week_rows]
        ly_f_vals = [r['ly_occ_final']   for r in week_rows]
        ly_s_vals = [r['ly_occ_same_pt'] for r in week_rows]
        y_fmt     = '.0%'
        y_title   = '稼働率'
        avail_line = None
    else:
        cur_vals  = cur_rooms_list
        ly_f_vals = ly_final_list
        ly_s_vals = ly_same_list
        y_fmt     = ',d'
        y_title   = '予約室数（週合計）'
        avail_line = 7 * rm.TOTAL_ROOMS

    fig_wk = go.Figure()
    fig_wk.add_bar(
        x=labels, y=cur_vals,
        name='今年現状',
        marker_color=['#2980b9' if r['yoy_ratio'] and r['yoy_ratio'] >= 1.0
                      else '#E74C3C' if r['yoy_ratio'] else '#AEB6BF'
                      for r in week_rows],
    )
    fig_wk.add_scatter(
        x=labels, y=ly_s_vals,
        name='昨年同時点',
        mode='lines+markers',
        line=dict(color='#E67E22', width=2),
        marker=dict(size=7, symbol='diamond'),
    )
    fig_wk.add_scatter(
        x=labels, y=ly_f_vals,
        name='昨年最終実績',
        mode='lines',
        line=dict(color='#BFC9CA', width=1.5, dash='dot'),
    )
    if avail_line:
        fig_wk.add_hline(
            y=avail_line, line_dash='dash', line_color='#E74C3C',
            annotation_text=f'満室（{avail_line}室）',
            annotation_position='top right',
        )

    fig_wk.update_layout(
        yaxis=dict(title=y_title, tickformat=y_fmt),
        xaxis_title='',
        height=380,
        barmode='overlay',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, x=0),
        hovermode='x unified',
    )
    st.plotly_chart(fig_wk, use_container_width=True)

    # ── ② 昨対比バーチャート ─────────────────────────────────────────
    st.subheader("昨年同時点比（今年現状 ÷ 昨年同リードタイム）")
    yoy_vals   = [r['yoy_ratio'] for r in week_rows]
    yoy_colors = ['#27AE60' if (v and v >= 1.0) else '#E74C3C' if v else '#ccc'
                  for v in yoy_vals]

    fig_yoy = go.Figure()
    fig_yoy.add_bar(x=labels, y=yoy_vals, marker_color=yoy_colors)
    fig_yoy.add_hline(y=1.0, line_dash='dash', line_color='gray', annotation_text='昨年同水準')
    fig_yoy.update_layout(
        yaxis=dict(tickformat='.0%', title=''),
        height=240,
        showlegend=False,
        hovermode='x unified',
    )
    st.plotly_chart(fig_yoy, use_container_width=True)

    # ── ③ 週別サマリーテーブル ───────────────────────────────────────
    st.subheader("週別サマリー")
    tbl_rows = []
    for r in week_rows:
        tbl_rows.append({
            '週':           r['label'],
            '今年予約室数': r['cur_rooms'],
            '今年稼働率':   f"{r['cur_occ']:.1%}",
            '昨年同時点室数': r['ly_rooms_same_pt'],
            '昨年最終室数': r['ly_rooms_final'],
            '昨年最終稼働率': f"{r['ly_occ_final']:.1%}",
            '昨対比':       f"{r['yoy_ratio']:.1%}" if r['yoy_ratio'] else '—',
        })
    st.dataframe(pd.DataFrame(tbl_rows), use_container_width=True, hide_index=True)

# ------------------------------------------------------------------
# Tab 4: 競合価格
# ------------------------------------------------------------------
with tab4:
    st.header("競合価格モニター")

    # 履歴データ読み込み（キャッシュ）
    @st.cache_data(show_spinner=False)
    def load_comp_history_cached(comp_bytes):
        src = io.BytesIO(comp_bytes) if comp_bytes else None
        return rm.load_comp_history(src)

    hist = load_comp_history_cached(comp_bytes)

    if not hist:
        st.info("competitor_prices.csv をアップロードすると競合価格が表示されます。")
    else:
        df_hist = pd.DataFrame(hist)
        fetch_dates  = sorted(df_hist['取得日'].unique())
        target_dates = sorted(df_hist['対象日'].unique())
        hotels       = sorted(df_hist['施設名'].unique())
        latest_fetch = fetch_dates[-1]
        df_latest    = df_hist[df_hist['取得日'] == latest_fetch].copy()

        WDAYS_JP = ['月','火','水','木','金','土','日']

        def date_label(d_str):
            """'2026/04/11' → '4/11(土)'"""
            try:
                d = datetime.strptime(d_str, '%Y/%m/%d')
                return f"{d.month}/{d.day}({WDAYS_JP[d.weekday()]})"
            except:
                return d_str

        # ---- ① Excelライク ピボット表（HTML直接レンダリング） ----
        st.caption(f"2食付き最安値（1人当たり）。× = 楽天に2食付きプランなし。　取得日: {latest_fetch}")

        # ピボット（施設 × 対象日）　※競合のみ
        pivot = df_latest.pivot_table(
            index='施設名', columns='対象日', values='価格', aggfunc='min'
        )

        # 自社行をスクレイピングデータから除外（RMデータで上書きするため）
        pivot = pivot.loc[[h for h in pivot.index if '甲子園' not in h]]

        # 今日以降の列だけに絞る
        today_str = datetime.now().strftime('%Y/%m/%d')
        future_cols = [c for c in pivot.columns if c >= today_str]
        if future_cols:
            pivot = pivot[future_cols]

        if pivot.empty or len(pivot.columns) == 0:
            st.warning("⚠️ 今日以降の競合価格データがありません。サイドバーの「🔍 競合価格を今すぐ取得」ボタンで最新データを取得してください。")
        else:
            # 競合平均行（自社除く）
            others = list(pivot.index)
            avg_row = pivot.mean(skipna=True)
            pivot.loc['競合平均（自社除く）'] = avg_row

            # 自社価格行：RMの推奨価格（2名合計）をマッピング
            own_prices = {
                r['date'].strftime('%Y/%m/%d'): r['sug_price']
                for r in rows
            }
            pivot.loc['ホテル甲子園（自社）'] = {
                col: own_prices.get(col) for col in pivot.columns
            }

            col_keys   = list(pivot.columns)           # 元の '2026/04/11' 形式
            col_labels_list = [date_label(c) for c in col_keys]

            # 行順：競合施設 → 競合平均 → 自社
            row_labels = (
                others
                + ['競合平均（自社除く）']
                + ['ホテル甲子園（自社）']
            )

            def cell_bg(row_name, col_label):
                if '甲子園' in str(row_name):   return '#FDE8E8'
                if '競合平均' in str(row_name): return '#E8F5E9'
                if '土' in col_label:            return '#FFF3E0'
                if '日' in col_label:            return '#FFEBEE'
                return '#FFFFFF'

            def cell_val(v, row_name):
                if pd.isna(v):
                    return '<span style="color:#CCCCCC">×</span>'
                n = int(v)
                bold = 'font-weight:bold;' if ('甲子園' in str(row_name) or '競合平均' in str(row_name)) else ''
                return f'<span style="{bold}">{n:,}</span>'

            # HTML組み立て
            th_base  = 'background:#1a3a5c;color:white;padding:4px 8px;white-space:nowrap;font-size:11px;text-align:center;border:1px solid #2c5f8a;position:sticky;top:0;z-index:2;'
            th_first = 'background:#1a3a5c;color:white;padding:4px 8px;min-width:130px;text-align:left;border:1px solid #2c5f8a;position:sticky;left:0;top:0;z-index:3;'
            td_first = 'padding:3px 10px;white-space:nowrap;font-size:12px;border:1px solid #ddd;position:sticky;left:0;z-index:1;font-weight:bold;'

            rows_html = []
            # ヘッダー行
            ths = f'<th style="{th_first}">施設名</th>'
            for lbl in col_labels_list:
                sat_bg = 'background:#d35400;' if '土' in lbl else ('background:#c0392b;' if '日' in lbl else '')
                ths += f'<th style="{th_base}{sat_bg}">{lbl}</th>'
            rows_html.append(f'<tr>{ths}</tr>')

            # データ行
            for rname in row_labels:
                if '甲子園' in str(rname):
                    row_bg = '#FDE8E8'; name_bg = '#f5b7b1;'
                elif '競合平均' in str(rname):
                    row_bg = '#E8F5E9'; name_bg = '#a9dfbf;'
                else:
                    row_bg = '#FFFFFF'; name_bg = '#2c3e50;color:white;'
                tds = f'<td style="{td_first}background:{name_bg if name_bg else row_bg}">{rname}</td>'
                for ck, cl in zip(col_keys, col_labels_list):
                    v   = pivot.loc[rname, ck]
                    bg  = cell_bg(rname, cl)
                    val = cell_val(v, rname)
                    tds += f'<td style="background:{bg};padding:3px 8px;text-align:right;font-size:12px;border:1px solid #eee;min-width:72px;">{val}</td>'
                rows_html.append(f'<tr>{tds}</tr>')

            html_table = f"""
            <div style="overflow-x:auto;max-height:480px;overflow-y:auto;border:1px solid #ccc;border-radius:4px;">
            <table style="border-collapse:collapse;width:max-content;">
            <thead>{''.join(rows_html[:1])}</thead>
            <tbody>{''.join(rows_html[1:])}</tbody>
            </table></div>
            """
            st.html(html_table)

        st.divider()

        # ---- ② 特定日の価格変動履歴 ----
        st.subheader("価格変動履歴（取得日ごと）")
        sel_target = st.selectbox("対象日を選択", target_dates, key='comp_target',
                                  format_func=date_label)
        df_chg = df_hist[df_hist['対象日'] == sel_target].dropna(subset=['価格'])

        if df_chg.empty:
            st.info("この日付は全施設「×（満室 or 非対応）」です。")
        else:
            fig_chg = px.line(
                df_chg, x='取得日', y='価格', color='施設名',
                markers=True,
                title=f"{date_label(sel_target)} の価格変動（取得日ごと）",
                labels={'価格': '最低価格(円)'},
            )
            for trace in fig_chg.data:
                if '甲子園' in trace.name:
                    trace.line.width = 3
                    trace.line.color = '#E74C3C'
            fig_chg.update_layout(height=360, yaxis_tickformat=',')
            st.plotly_chart(fig_chg, use_container_width=True)

            # 前回比テーブル
            if len(fetch_dates) >= 2:
                prev, curr = fetch_dates[-2], fetch_dates[-1]
                rows_delta = []
                for h in hotels:
                    def get_price(fd):
                        v = df_hist[(df_hist['取得日']==fd)&(df_hist['対象日']==sel_target)&(df_hist['施設名']==h)]['価格'].values
                        return int(v[0]) if len(v) and pd.notna(v[0]) else None
                    cv, pv = get_price(curr), get_price(prev)
                    delta = cv - pv if (cv and pv) else None
                    arrow = ('⬆' if delta > 0 else '⬇' if delta < 0 else '→') if delta is not None else ''
                    rows_delta.append({
                        '施設': h,
                        f'{prev}': f"¥{pv:,}" if pv else '×',
                        f'{curr}（最新）': f"¥{cv:,}" if cv else '×',
                        '変動': f"{arrow} {delta:+,}円" if delta is not None else '−',
                    })
                st.dataframe(pd.DataFrame(rows_delta), use_container_width=True, hide_index=True)

        st.divider()

        # ---- ③ 生データダウンロード ----
        st.subheader("生データ")
        col_dl1, col_dl2 = st.columns(2)
        try:
            with open(rm.COMP_PRICES_CSV, 'rb') as f:
                col_dl1.download_button("⬇ 全履歴CSV", f.read(),
                    "competitor_prices.csv", "text/csv")
        except FileNotFoundError:
            col_dl1.info("ローカルCSVが見つかりません")

        col_dl2.download_button(
            f"⬇ 最新日のみ（{latest_fetch}）",
            df_latest.to_csv(index=False).encode('utf-8-sig'),
            f"comp_{latest_fetch.replace('/','')}.csv", "text/csv",
        )

        with st.expander("全履歴テーブルを表示"):
            df_show = df_hist.copy()
            df_show['価格'] = df_show['価格'].apply(lambda x: f"¥{int(x):,}" if pd.notna(x) else '×')
            st.dataframe(df_show, use_container_width=True, hide_index=True, height=400)

# ------------------------------------------------------------------
# Tab 5: 月次売上サマリー
# ------------------------------------------------------------------
with tab5:
    st.header("月次売上サマリー")

    import calendar as cal_mod

    budget = rm.MONTHLY_BUDGET
    months = sorted(set(list(budget.keys()) + list(monthly_rev.keys())))


    rows_ms = []
    for m in months:
        bgt  = budget.get(m, 0)
        rev  = monthly_rev.get(m, 0)
        # 昨対：同月の前年
        prev_m = str(int(m[:4]) - 1) + m[4:]
        prev_rev = monthly_rev.get(prev_m, 0)
        yoy  = rev / prev_rev if (prev_rev and rev) else None
        yoy_diff = rev - prev_rev if (prev_rev and rev) else None
        diff = rev - bgt
        rate = rev / bgt if bgt > 0 else None
        label = f"{m[:4]}/{m[4:]}"
        rows_ms.append({
            '年月': label,
            '予算': bgt,
            '実績': rev,
            '差異': diff,
            '達成率': rate,
            '昨対実績': prev_rev,
            '昨対比': yoy,
            '昨対差額': yoy_diff,
        })

    df_ms = pd.DataFrame(rows_ms)

    # ================================================================
    # 着地見込み（未来月）
    # ================================================================
    st.subheader("🔮 着地見込み（未来月）")
    st.caption("現状予約 ÷ ブッキングカーブ進捗率 で最終稼働率を推計。売上は昨年同月ADR × 着地見込み室数。")

    landing = rm.calc_landing_forecast(daily, lead_dist, monthly_rev, room_monthly)
    future_landing = [r for r in landing if not r['is_past']]

    if not future_landing:
        st.info("PMSデータをアップロードすると着地見込みが表示されます。")
    else:
        # ── HTMLテーブルで表示 ────────────────────────────────
        th_s  = 'background:#1a3a5c;color:white;padding:5px 10px;white-space:nowrap;font-size:12px;text-align:center;border:1px solid #2c5f8a;'
        th_ls = th_s.replace('text-align:center', 'text-align:left')
        td_s  = 'padding:4px 10px;text-align:right;border:1px solid #ddd;font-size:12px;white-space:nowrap;'
        td_ls = td_s.replace('text-align:right', 'text-align:left')

        def occ_color(occ):
            if occ >= 0.85: return '#27AE60'
            if occ >= 0.70: return '#A9DFBF'
            if occ >= 0.55: return '#FAD7A0'
            return '#F1948A'

        def rev_vs_bgt_color(rev, bgt):
            if not bgt: return '#FFFFFF'
            ratio = rev / bgt
            if ratio >= 1.00: return '#A9DFBF'
            if ratio >= 0.90: return '#FAD7A0'
            return '#F1948A'

        rows_html_lf = [
            f'<tr>'
            f'<th style="{th_ls}">月</th>'
            f'<th style="{th_s}">残日数</th>'
            f'<th style="{th_s}">現状<br>予約室数</th>'
            f'<th style="{th_s}">現状<br>稼働率</th>'
            f'<th style="{th_s}">着地見込み<br>稼働率</th>'
            f'<th style="{th_s}">着地見込み<br>売上</th>'
            f'<th style="{th_s}">昨年<br>稼働率</th>'
            f'<th style="{th_s}">昨年<br>売上</th>'
            f'<th style="{th_s}">予算</th>'
            f'<th style="{th_s}">予算比<br>（見込み）</th>'
            f'</tr>'
        ]

        for r in future_landing:
            fc_occ = r['forecast_occ']
            ly_occ = r['last_year_occ']
            fc_rev = r['forecast_rev'] or 0
            bgt    = r['budget']

            # 着地見込み稼働率セル（色付き）
            fc_occ_bg  = occ_color(fc_occ)
            ly_occ_bg  = occ_color(ly_occ) if ly_occ else '#FFFFFF'
            fc_rev_bg  = rev_vs_bgt_color(fc_rev, bgt)

            # 予算比
            bgt_ratio = fc_rev / bgt if (fc_rev and bgt) else None
            bgt_str   = f"{bgt_ratio:.1%}" if bgt_ratio else '—'

            # 昨対（着地見込み÷昨年）
            yoy_fc = fc_occ / ly_occ if (fc_occ and ly_occ) else None
            yoy_str = f"{yoy_fc:.1%}" if yoy_fc else '—'

            rows_html_lf.append(
                f'<tr>'
                f'<td style="{td_ls}font-weight:bold;">{r["label"]}</td>'
                f'<td style="{td_s}">{r["remaining_days"]}日</td>'
                f'<td style="{td_s}">{r["cur_nights"]}室 / {r["avail"]}室</td>'
                f'<td style="{td_s}background:{occ_color(r["cur_occ"])};">{r["cur_occ"]:.1%}</td>'
                f'<td style="{td_s}background:{fc_occ_bg};font-weight:bold;">{fc_occ:.1%}</td>'
                f'<td style="{td_s}background:{fc_rev_bg};">{"¥{:,.0f}".format(fc_rev) if fc_rev else "—"}</td>'
                f'<td style="{td_s}background:{ly_occ_bg};">{"" if not ly_occ else f"{ly_occ:.1%}"}</td>'
                f'<td style="{td_s}">{"¥{:,.0f}".format(r["last_year_rev"]) if r["last_year_rev"] else "—"}</td>'
                f'<td style="{td_s}">{"¥{:,.0f}".format(bgt) if bgt else "—"}</td>'
                f'<td style="{td_s}font-weight:bold;">{bgt_str}</td>'
                f'</tr>'
            )

        landing_html = f"""
        <div style="overflow-x:auto;border:1px solid #ccc;border-radius:4px;margin-bottom:8px;">
        <table style="border-collapse:collapse;width:max-content;min-width:100%;">
        <thead>{''.join(rows_html_lf[:1])}</thead>
        <tbody>{''.join(rows_html_lf[1:])}</tbody>
        </table></div>
        <p style="font-size:11px;color:#888;margin-top:2px;">
        ※ 着地見込み = 確定予約 ÷ ブッキングカーブ進捗率（日別計算）。
        リードタイムが長い月ほど精度が下がります。昨年データがない場合は売上見込みは表示されません。
        </p>
        """
        st.html(landing_html)

        # ミニグラフ（着地見込み vs 昨年 vs 予算）
        df_lf = pd.DataFrame([{
            '月':             r['label'],
            '着地見込み稼働率': r['forecast_occ'],
            '昨年稼働率':      r['last_year_occ'],
            '現状稼働率':      r['cur_occ'],
        } for r in future_landing])
        fig_lf = go.Figure()
        fig_lf.add_bar(name='現状稼働率',      x=df_lf['月'], y=df_lf['現状稼働率'],      marker_color='#AEB6BF')
        fig_lf.add_bar(name='着地見込み稼働率', x=df_lf['月'], y=df_lf['着地見込み稼働率'], marker_color='#2980b9')
        fig_lf.add_scatter(name='昨年稼働率',   x=df_lf['月'], y=df_lf['昨年稼働率'],
                           mode='lines+markers', line=dict(color='#E67E22', dash='dot', width=2))
        fig_lf.update_layout(
            barmode='group', height=280, yaxis_tickformat='.0%',
            title='稼働率：現状 vs 着地見込み vs 昨年',
            yaxis=dict(range=[0, 1.1]),
        )
        st.plotly_chart(fig_lf, use_container_width=True)

    st.divider()

    # ================================================================
    # KPI ダッシュボード（月別）← チャートより先に表示
    # ================================================================
    st.subheader("📊 月別 KPI ダッシュボード")

    kpi_months = sorted(
        set(list(monthly_rev.keys()) + list(room_monthly.keys())),
        reverse=True
    )
    kpi_months = [m for m in kpi_months if monthly_rev.get(m, 0) > 0]

    if not kpi_months:
        st.info("PMSデータを読み込むとKPIが表示されます。")
    else:
        # 月ナビゲーター
        kc1, kc2, kc3 = st.columns([1, 3, 1])
        kpi_idx = kc2.selectbox(
            "月を選択", range(len(kpi_months)),
            format_func=lambda i: f"{kpi_months[i][:4]}年{kpi_months[i][4:]}月",
            key='kpi_month_sel'
        )
        sel_m  = kpi_months[kpi_idx]
        yr, mo = int(sel_m[:4]), int(sel_m[4:])
        days   = cal_mod.monthrange(yr, mo)[1]

        # 前月・前年同月
        prev_m_yr  = yr - 1 if mo == 1 else yr
        prev_m_mo  = 12     if mo == 1 else mo - 1
        prev_m     = f"{prev_m_yr}{prev_m_mo:02d}"
        prev_y     = f"{yr-1}{mo:02d}"

        def month_kpis(m):
            rev     = monthly_rev.get(m, 0)
            rdata   = room_monthly.get(m, {})
            nights  = sum(v['nights'] for v in rdata.values())
            guests  = monthly_guests.get(m, 0)
            d_yr, d_mo = int(m[:4]), int(m[4:])
            d_days  = cal_mod.monthrange(d_yr, d_mo)[1]
            avail   = d_days * rm.TOTAL_ROOMS
            occ     = nights / avail       if avail  else 0
            adr     = rev    / nights      if nights else 0
            revpar  = rev    / avail       if avail  else 0
            tanka   = rev    / guests      if guests else 0
            oper    = sum(1 for d, cnt in daily.items()
                         if d.strftime('%Y%m') == m and cnt > 0)
            return {
                '売上':     rev,
                '稼働率':   occ,
                'ADR':      adr,
                'RevPAR':   revpar,
                '販売室数': nights,
                '宿泊人数': guests,
                '客単価':   tanka,
                '営業日数': oper,
            }

        cur  = month_kpis(sel_m)
        prvy = month_kpis(prev_y)
        prvm = month_kpis(prev_m)

        def delta_fmt(cur_v, prev_v, is_pct=False, is_int=False):
            if not prev_v: return None
            diff = cur_v - prev_v
            pct  = diff / prev_v * 100
            sign = '+' if diff >= 0 else ''
            if is_pct:
                return f"{sign}{diff*100:.1f}pt（{pct:+.1f}%）"
            if is_int:
                return f"{sign}{diff:,.0f}（{pct:+.1f}%）"
            return f"{sign}¥{diff:,.0f}（{pct:+.1f}%）"

        st.caption(f"　昨対 = {yr-1}年{mo}月比　　前月 = {prev_m_yr}年{prev_m_mo}月比")

        # KPIカード 4列 × 2行
        kpis = [
            ('売上',     f"¥{cur['売上']:,.0f}",     delta_fmt(cur['売上'],     prvy['売上']),     delta_fmt(cur['売上'],     prvm['売上'])),
            ('稼働率',   f"{cur['稼働率']:.1%}",      delta_fmt(cur['稼働率'],   prvy['稼働率'],True), delta_fmt(cur['稼働率'],   prvm['稼働率'],True)),
            ('ADR',      f"¥{cur['ADR']:,.0f}",       delta_fmt(cur['ADR'],      prvy['ADR']),      delta_fmt(cur['ADR'],      prvm['ADR'])),
            ('RevPAR',   f"¥{cur['RevPAR']:,.0f}",    delta_fmt(cur['RevPAR'],   prvy['RevPAR']),   delta_fmt(cur['RevPAR'],   prvm['RevPAR'])),
            ('販売室数', f"{cur['販売室数']:,}室",     delta_fmt(cur['販売室数'], prvy['販売室数'],False,True), delta_fmt(cur['販売室数'], prvm['販売室数'],False,True)),
            ('宿泊人数', f"{cur['宿泊人数']:,}人",     delta_fmt(cur['宿泊人数'], prvy['宿泊人数'],False,True), delta_fmt(cur['宿泊人数'], prvm['宿泊人数'],False,True)),
            ('客単価',   f"¥{cur['客単価']:,.0f}",    delta_fmt(cur['客単価'],   prvy['客単価']),   delta_fmt(cur['客単価'],   prvm['客単価'])),
            ('営業日数', f"{cur['営業日数']}日/{days}日", None, None),
        ]

        for row_start in range(0, len(kpis), 4):
            cols = st.columns(4)
            for col, (label, val, yoy_d, mom_d) in zip(cols, kpis[row_start:row_start+4]):
                col.metric(
                    label=f"{label}　昨対",
                    value=val,
                    delta=yoy_d,
                )
            cols2 = st.columns(4)
            for col, (label, val, yoy_d, mom_d) in zip(cols2, kpis[row_start:row_start+4]):
                col.metric(
                    label=f"{label}　前月比",
                    value="",
                    delta=mom_d,
                )
            st.divider()

        st.caption("※ 宿泊人数はPMSの人数項目から取得。項目名が異なる場合は2名/室で推定しています。DORはPMSの定義をご確認の上、追加実装できます。")

    # ================================================================
    # 予算 vs 実績 グラフ・テーブル（直近12ヶ月）
    # ================================================================
    st.divider()
    st.subheader("📅 予算 vs 実績（直近12ヶ月）")

    # 実績がある月のうち最新12ヶ月に絞る
    rev_months_with_data = sorted([m for m in months if monthly_rev.get(m, 0) > 0])
    last_12 = rev_months_with_data[-12:] if len(rev_months_with_data) >= 12 else rev_months_with_data
    df_plot = df_ms[df_ms['年月'].isin([f"{m[:4]}/{m[4:]}" for m in last_12])].copy()

    if not df_plot.empty:
        # グループドバーチャート
        bar_colors = []
        for v, b in zip(df_plot['実績'], df_plot['予算']):
            if b > 0:
                bar_colors.append('#27AE60' if v >= b else '#E74C3C')
            else:
                bar_colors.append('#95A5A6')

        fig_ms = go.Figure()
        fig_ms.add_bar(name='昨対実績', x=df_plot['年月'], y=df_plot['昨対実績'], marker_color='#D5D8DC')
        fig_ms.add_bar(name='予算',     x=df_plot['年月'], y=df_plot['予算'],     marker_color='#AEB6BF')
        fig_ms.add_bar(name='実績',     x=df_plot['年月'], y=df_plot['実績'],     marker_color=bar_colors)
        fig_ms.update_layout(
            barmode='group', title='予算 vs 実績 vs 昨対（直近12ヶ月）', height=350,
            yaxis_tickformat=',.0f',
        )
        st.plotly_chart(fig_ms, use_container_width=True)

        # 達成率＋昨対比ライン
        df_rate = df_plot[df_plot['達成率'].notna() & (df_plot['実績'] > 0)].copy()
        if not df_rate.empty:
            fig_rate = go.Figure()
            fig_rate.add_scatter(x=df_rate['年月'], y=df_rate['達成率'],
                                 name='予算達成率', mode='lines+markers',
                                 line=dict(color='#2980b9'))
            df_yoy = df_rate[df_rate['昨対比'].notna()]
            if not df_yoy.empty:
                fig_rate.add_scatter(x=df_yoy['年月'], y=df_yoy['昨対比'],
                                     name='昨対比', mode='lines+markers',
                                     line=dict(color='#E67E22', dash='dot'))
            fig_rate.add_hline(y=1.0, line_dash='dash', line_color='gray', annotation_text='100%')
            fig_rate.update_layout(yaxis_tickformat='.0%', height=280,
                                   title='月次達成率 & 昨対比')
            st.plotly_chart(fig_rate, use_container_width=True)

    # テーブル（全期間）
    with st.expander("📋 全月データ一覧", expanded=False):
        df_disp = df_ms.copy()
        df_disp['予算']     = df_disp['予算'].map(lambda x: f"¥{x:,.0f}" if x else "—")
        df_disp['実績']     = df_disp['実績'].map(lambda x: f"¥{x:,.0f}" if x else "—")
        df_disp['差異']     = df_disp['差異'].map(lambda x: f"¥{x:+,.0f}" if x else "—")
        df_disp['達成率']   = df_disp['達成率'].map(lambda x: f"{x:.1%}" if x else "—")
        df_disp['昨対実績'] = df_disp['昨対実績'].map(lambda x: f"¥{x:,.0f}" if x else "—")
        df_disp['昨対比']   = df_disp['昨対比'].map(lambda x: f"{x:.1%}" if x else "—")
        df_disp['昨対差額'] = df_disp['昨対差額'].map(lambda x: f"¥{x:+,.0f}" if x else "—")
        st.dataframe(df_disp, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------
# Tab 6: 売上内訳（科目別）
# ------------------------------------------------------------------
with tab6:
    st.header("売上内訳（科目別）")

    if not sales_detail:
        st.info("PMSデータ（a.csv）が読み込まれていません。")
    else:
        df_sd = pd.DataFrame(sales_detail)

        # 負の金額があれば警告
        neg = df_sd[df_sd['金額'] < 0]
        if not neg.empty:
            st.warning(f"⚠️ マイナス金額の行が {len(neg)} 件あります（修正・返金分として集計に含めています）")

        CATS = ['宿泊', '昼休・日帰り', 'ドリンク', 'その他']
        CAT_COLORS = {
            '宿泊':        '#2980b9',
            '昼休・日帰り': '#27AE60',
            'ドリンク':    '#E67E22',
            'その他':      '#95A5A6',
        }
        df_sd['年']  = df_sd['month'].str[:4]
        df_sd['月']  = df_sd['month'].apply(lambda m: f"{m[:4]}/{m[4:]}")
        df_sd['月番'] = df_sd['month'].str[4:].astype(int)

        # ---- フィルター ----
        all_years = sorted(df_sd['年'].unique(), reverse=True)
        fc1, fc2 = st.columns([2, 3])
        sel_year = fc1.selectbox("年", all_years, key='sd_year')
        month_opts = ['全月'] + [f"{m}月" for m in range(1, 13)]
        sel_month_label = fc2.selectbox("月", month_opts, key='sd_month')
        sel_month_num = None if sel_month_label == '全月' else int(sel_month_label.replace('月',''))

        # フィルター適用
        df_f = df_sd[df_sd['年'] == sel_year]
        if sel_month_num:
            df_f = df_f[df_f['月番'] == sel_month_num]

        period_label = f"{sel_year}年" + (f"{sel_month_num}月" if sel_month_num else "（全月）")

        if df_f.empty:
            st.info(f"{period_label} のデータがありません。")
        else:
            # ---- 比較期間の計算ヘルパー ----
            def get_period_total(year_str, month_num, cat=None):
                mask = df_sd['年'] == year_str
                if month_num:
                    mask &= df_sd['月番'] == month_num
                df_tmp = df_sd[mask]
                if cat:
                    df_tmp = df_tmp[df_tmp['カテゴリ'] == cat]
                return df_tmp['金額'].sum()

            # 昨対期間
            yoy_year = str(int(sel_year) - 1)
            # 前月期間
            if sel_month_num:
                if sel_month_num == 1:
                    mom_year, mom_month = str(int(sel_year) - 1), 12
                else:
                    mom_year, mom_month = sel_year, sel_month_num - 1
            else:
                mom_year, mom_month = None, None

            def delta_str(cur, prev):
                if not prev:
                    return None
                diff = cur - prev
                pct  = diff / prev
                sign = '+' if diff >= 0 else ''
                return f"{sign}¥{diff:,.0f}（{pct:+.1%}）"

            # ---- 大項目KPIカード ----
            st.subheader(f"📊 大項目サマリー　{period_label}")

            # 合計行
            total      = df_f['金額'].sum()
            yoy_total  = get_period_total(yoy_year, sel_month_num)
            mom_total  = get_period_total(mom_year, mom_month) if mom_year else 0

            tc1, tc2, tc3 = st.columns(3)
            tc1.metric("合計売上",  f"¥{total:,.0f}")
            tc2.metric("昨対比（合計）",
                       f"¥{yoy_total:,.0f}" if yoy_total else "データなし",
                       delta_str(total, yoy_total))
            if sel_month_num:
                tc3.metric(f"前月比（{mom_year}/{mom_month:02d}）",
                           f"¥{mom_total:,.0f}" if mom_total else "データなし",
                           delta_str(total, mom_total))
            st.divider()

            # カテゴリ別KPIカード（4列）
            c1, c2, c3, c4 = st.columns(4)
            for col, cat in zip([c1, c2, c3, c4], CATS):
                amt     = df_f[df_f['カテゴリ'] == cat]['金額'].sum()
                yoy_amt = get_period_total(yoy_year, sel_month_num, cat)
                mom_amt = get_period_total(mom_year, mom_month, cat) if mom_year else 0
                # deltaは昨対優先、月選択時は前月も使える
                if sel_month_num and mom_amt:
                    d = delta_str(amt, mom_amt)
                    d_label = f"{cat}（前月比）"
                elif yoy_amt:
                    d = delta_str(amt, yoy_amt)
                    d_label = f"{cat}（昨対比）"
                else:
                    d = None
                    d_label = cat
                col.metric(d_label, f"¥{amt:,.0f}", d)
            st.divider()

            # ---- グラフ：月別（全月選択時）または 科目別円グラフ（月選択時）----
            if sel_month_num is None:
                # 月別積み上げ棒グラフ
                df_pivot = (
                    df_f.groupby(['月', 'カテゴリ'])['金額'].sum().reset_index()
                )
                df_pivot['カテゴリ'] = pd.Categorical(df_pivot['カテゴリ'], categories=CATS, ordered=True)
                df_pivot = df_pivot.sort_values(['月', 'カテゴリ'])
                fig_sd = px.bar(
                    df_pivot, x='月', y='金額', color='カテゴリ',
                    color_discrete_map=CAT_COLORS,
                    title=f'{sel_year}年　月別売上内訳',
                    labels={'金額': '売上（円）', '月': ''},
                    barmode='stack',
                )
                fig_sd.update_layout(height=360, yaxis_tickformat=',.0f')
                st.plotly_chart(fig_sd, use_container_width=True)
            else:
                # 月選択時：円グラフ
                cat_totals = df_f.groupby('カテゴリ')['金額'].sum().reset_index()
                fig_pie = px.pie(
                    cat_totals, names='カテゴリ', values='金額',
                    color='カテゴリ', color_discrete_map=CAT_COLORS,
                    title=f'{period_label}　売上構成',
                )
                fig_pie.update_layout(height=340)
                st.plotly_chart(fig_pie, use_container_width=True)

            # ---- ドリンクランキング ----
            st.subheader("🍺 ドリンク売れ筋ランキング")
            df_drink = df_f[df_f['カテゴリ'] == 'ドリンク']
            if df_drink.empty:
                st.info("この期間のドリンク売上データがありません。")
            else:
                rank = (
                    df_drink.groupby('科目')['金額']
                    .agg(売上合計='sum', 件数='count')
                    .reset_index()
                    .sort_values('売上合計', ascending=False)
                    .reset_index(drop=True)
                )
                rank.index += 1
                rank['売上合計'] = rank['売上合計'].map(lambda x: f"¥{x:,.0f}")
                st.dataframe(rank, use_container_width=True)

            # ---- 全科目一覧（確認用） ----
            with st.expander("📋 科目一覧（カテゴリ分類の確認・調整用）"):
                st.caption("「ドリンク」に分類されていない科目があれば教えてください。")
                summary = (
                    df_f.groupby(['カテゴリ', '科目'])['金額']
                    .agg(合計='sum', 件数='count')
                    .reset_index()
                    .sort_values(['カテゴリ', '合計'], ascending=[True, False])
                )
                summary['合計'] = summary['合計'].map(lambda x: f"¥{x:,.0f}")
                st.dataframe(summary, use_container_width=True, hide_index=True)
