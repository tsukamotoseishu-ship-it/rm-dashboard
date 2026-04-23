"""
RM価格案 Webダッシュボード - ホテル甲子園
Streamlit アプリ本体
実行: streamlit run app.py
"""

import io
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
        st.info("競合価格CSV未アップロード\n（サンプルデータを使用）")

    st.caption(f"最終更新: {datetime.now().strftime('%Y/%m/%d %H:%M')}")

# ============================================================
# データ読み込み（キャッシュ）
# ============================================================
@st.cache_data(show_spinner="データを読み込み中...")
def load_cached(pms_bytes, comp_bytes):
    """ファイルの内容（bytes）をキーにキャッシュ"""
    pms_src  = io.BytesIO(pms_bytes)  if pms_bytes  else None
    comp_src = io.BytesIO(comp_bytes) if comp_bytes  else None

    daily, lead_dist, comp_prices, data_source, monthly_rev = rm.load_data(
        pms_file=pms_src, comp_file=comp_src
    )
    rows = rm.calc_rm_rows(daily, comp_prices)
    # defaultdict(lambda:...) は pickle 不可なので通常の dict に変換
    daily       = dict(daily)
    lead_dist   = {k: dict(v) for k, v in lead_dist.items()}
    comp_prices = {k: dict(v) for k, v in comp_prices.items()}
    monthly_rev = dict(monthly_rev)
    return daily, lead_dist, comp_prices, data_source, monthly_rev, rows


pms_bytes  = pms_file.read()  if pms_file  else None
comp_bytes = comp_file.read() if comp_file else None

try:
    daily, lead_dist, comp_prices, data_source, monthly_rev, rows = load_cached(
        pms_bytes, comp_bytes
    )
except Exception as e:
    st.error(f"データ読み込みエラー: {e}")
    st.info("PMSデータ（a.csv）をサイドバーからアップロードしてください。")
    st.stop()

# ============================================================
# 5タブ
# ============================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 RM価格案",
    "📊 実績稼働率",
    "📈 ブッキングカーブ",
    "💴 競合価格",
    "💰 月次売上",
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

    df_rm = pd.DataFrame([{
        '日付':       r['date_str'],
        '曜日':       r['wday'],
        'LT(日)':    r['lead'],
        '目標消化':   f"{r['target']:.1%}",
        '実績消化':   f"{r['actual']:.1%}",
        '差異':       f"{r['diff']:+.1%}",
        'アクション': action_label(r['action']),
        '現在ランク': r['cur_rank'],
        '現在価格':   f"¥{r['cur_price']:,}",
        '推奨ランク': r['sug_rank'],
        '推奨価格':   f"¥{r['sug_price']:,}",
        '競合平均':   f"¥{r['cavg']:,}" if r['cavg'] else '—',
    } for r in rows])

    st.dataframe(
        df_rm,
        use_container_width=True,
        hide_index=True,
        height=600,
    )

    # CSV ダウンロード
    csv_bytes = df_rm.to_csv(index=False).encode('utf-8-sig')
    st.download_button("⬇ CSV ダウンロード", csv_bytes, "rm_plan.csv", "text/csv")

# ------------------------------------------------------------------
# Tab 2: 実績稼働率
# ------------------------------------------------------------------
with tab2:
    st.header("実績稼働率")

    if not daily:
        st.info("PMSデータをアップロードすると実績が表示されます。")
    else:
        occ_data = {d: c / rm.TOTAL_ROOMS for d, c in daily.items()}
        df_occ = pd.DataFrame([
            {'年月': d.strftime('%Y/%m'), '日': d.day, '稼働率': v}
            for d, v in sorted(occ_data.items())
        ])

        # ヒートマップ（月×日）
        if not df_occ.empty:
            pivot = df_occ.pivot_table(index='年月', columns='日', values='稼働率', aggfunc='mean')
            fig_hm = px.imshow(
                pivot,
                color_continuous_scale='RdYlGn',
                zmin=0, zmax=1,
                title='稼働率ヒートマップ（月×日付）',
                labels={'color': '稼働率'},
                aspect='auto',
            )
            fig_hm.update_layout(height=max(300, len(pivot) * 30 + 80))
            st.plotly_chart(fig_hm, use_container_width=True)

        # 月別平均バーチャート
        df_monthly_occ = df_occ.groupby('年月')['稼働率'].mean().reset_index()
        fig_bar = px.bar(
            df_monthly_occ,
            x='年月', y='稼働率',
            color='稼働率',
            color_continuous_scale='RdYlGn',
            range_color=[0, 1],
            title='月別平均稼働率',
        )
        fig_bar.update_layout(yaxis_tickformat='.0%', height=300, showlegend=False)
        st.plotly_chart(fig_bar, use_container_width=True)

# ------------------------------------------------------------------
# Tab 3: ブッキングカーブ
# ------------------------------------------------------------------
with tab3:
    st.header("ブッキングカーブ分析")

    # 理論カーブ
    theory_x = list(range(0, 91))
    theory_y = [rm.booking_curve_at(x) for x in theory_x]
    fig_bc = go.Figure()
    fig_bc.add_trace(go.Scatter(
        x=theory_x, y=theory_y,
        mode='lines', name='理論カーブ（設定値）',
        line=dict(color='#2980b9', dash='dash'),
    ))

    # 実績カーブ（曜日タイプ別）
    colors = {'土/連休': '#E74C3C', '日曜': '#E67E22', '平日': '#27AE60'}
    for dtype, weekly in lead_dist.items():
        if not weekly: continue
        total = sum(weekly.values())
        cumsum = 0
        xs, ys = [], []
        for w in sorted(weekly.keys(), reverse=True):
            cumsum += weekly[w]
            xs.append(w * 7)
            ys.append(cumsum / total)
        xs.reverse(); ys.reverse()
        fig_bc.add_trace(go.Scatter(
            x=xs, y=ys, mode='lines+markers',
            name=f'実績({dtype})',
            line=dict(color=colors.get(dtype, '#8E44AD')),
        ))

    fig_bc.update_layout(
        title='累積予約進捗（リードタイム別）',
        xaxis_title='宿泊日までの日数',
        yaxis_title='累積割合',
        yaxis_tickformat='.0%',
        height=400,
        xaxis=dict(autorange='reversed'),
    )
    st.plotly_chart(fig_bc, use_container_width=True)

    # リードタイム分布
    st.subheader("リードタイム週別分布")
    all_weeks = sorted(set(w for wdict in lead_dist.values() for w in wdict.keys()))
    if all_weeks:
        rows_lt = []
        for dtype, weekly in lead_dist.items():
            for w, cnt in weekly.items():
                rows_lt.append({'曜日タイプ': dtype, 'リードタイム(週)': w, '予約数': cnt})
        df_lt = pd.DataFrame(rows_lt)
        fig_lt = px.bar(
            df_lt, x='リードタイム(週)', y='予約数', color='曜日タイプ',
            barmode='group', title='週別リードタイム分布',
        )
        fig_lt.update_layout(height=300)
        st.plotly_chart(fig_lt, use_container_width=True)

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
        st.caption(f"2食付き最安値。× = 楽天に2食付きプランなし。　取得日: {latest_fetch}")

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

    budget = rm.MONTHLY_BUDGET
    months = sorted(set(list(budget.keys()) + list(monthly_rev.keys())))

    rows_ms = []
    for m in months:
        bgt = budget.get(m, 0)
        rev = monthly_rev.get(m, 0)
        diff = rev - bgt
        rate = rev / bgt if bgt > 0 else None
        label = f"{m[:4]}/{m[4:]}"
        rows_ms.append({
            '年月': label,
            '予算': bgt,
            '実績': rev,
            '差異': diff,
            '達成率': rate,
        })

    df_ms = pd.DataFrame(rows_ms)

    # KPIカード：最新月
    latest = [r for r in rows_ms if r['実績'] > 0]
    if latest:
        lm = latest[-1]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("最新月", lm['年月'])
        c2.metric("実績売上", f"¥{lm['実績']:,.0f}")
        c3.metric("予算", f"¥{lm['予算']:,.0f}")
        rate_str = f"{lm['達成率']:.1%}" if lm['達成率'] else "—"
        delta_str = f"¥{lm['差異']:+,.0f}"
        c4.metric("達成率", rate_str, delta_str)
        st.divider()

    # グループドバーチャート
    df_plot = df_ms[df_ms['予算'] > 0].copy()
    fig_ms = go.Figure()
    fig_ms.add_bar(name='予算',  x=df_plot['年月'], y=df_plot['予算'],  marker_color='#BDC3C7')
    fig_ms.add_bar(name='実績',  x=df_plot['年月'], y=df_plot['実績'],
                   marker_color=['#27AE60' if v >= b else '#E74C3C'
                                 for v, b in zip(df_plot['実績'], df_plot['予算'])])
    fig_ms.update_layout(
        barmode='group', title='予算 vs 実績', height=350,
        yaxis_tickformat=',.0f',
    )
    st.plotly_chart(fig_ms, use_container_width=True)

    # 達成率ライン
    df_rate = df_plot[df_plot['達成率'].notna() & (df_plot['実績'] > 0)]
    if not df_rate.empty:
        fig_rate = px.line(
            df_rate, x='年月', y='達成率',
            title='月次達成率推移',
            markers=True,
        )
        fig_rate.add_hline(y=1.0, line_dash='dash', line_color='gray', annotation_text='100%')
        fig_rate.update_layout(yaxis_tickformat='.0%', height=280)
        st.plotly_chart(fig_rate, use_container_width=True)

    # テーブル
    df_disp = df_ms.copy()
    df_disp['予算']   = df_disp['予算'].map(lambda x: f"¥{x:,.0f}" if x else "—")
    df_disp['実績']   = df_disp['実績'].map(lambda x: f"¥{x:,.0f}" if x else "—")
    df_disp['差異']   = df_disp['差異'].map(lambda x: f"¥{x:+,.0f}" if x else "—")
    df_disp['達成率'] = df_disp['達成率'].map(lambda x: f"{x:.1%}" if x else "—")
    st.dataframe(df_disp, use_container_width=True, hide_index=True)
