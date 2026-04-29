"""
Supabase データベース操作モジュール
=====================================
PMSデータ・競合価格の蓄積・取得を管理する。
"""

import os
import streamlit as st
from supabase import create_client, Client
from datetime import date, datetime
from collections import defaultdict


# ============================================================
# クライアント初期化
# ============================================================

def get_client() -> Client:
    """Supabaseクライアントを返す（Streamlit Secrets or 環境変数から認証情報取得）"""
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
    except Exception:
        url = os.environ.get("SUPABASE_URL", "")
        key = os.environ.get("SUPABASE_KEY", "")
    if not url or not key:
        raise ValueError("SUPABASE_URL / SUPABASE_KEY が設定されていません")
    return create_client(url, key)


# ============================================================
# PMS データ保存・取得
# ============================================================

def upsert_pms_daily(daily_dict: dict, monthly_rev: dict, room_monthly: dict, monthly_guests: dict):
    """
    daily辞書（{datetime: rooms}）と月次サマリーをSupabaseに保存（upsert）。
    既存レコードは上書き更新。
    """
    sb = get_client()

    # --- pms_daily テーブル ---
    rows = []
    for dt, rooms in daily_dict.items():
        d = dt.date() if isinstance(dt, datetime) else dt
        ym = d.strftime("%Y%m")
        rows.append({
            "stay_date":  d.isoformat(),
            "rooms":      int(rooms),
            "revenue":    int(monthly_rev.get(ym, 0)),   # 月次を日付に紐付け（概算）
            "is_monthly": False,
        })
    if rows:
        sb.table("pms_daily").upsert(rows, on_conflict="stay_date").execute()

    # --- pms_monthly テーブル ---
    months = set(list(monthly_rev.keys()) + list(room_monthly.keys()))
    mrows = []
    for ym in months:
        mrows.append({
            "year_month": ym,
            "revenue":    int(monthly_rev.get(ym, 0)),
            "nights":     int(room_monthly.get(ym, 0)),
            "guests":     int(monthly_guests.get(ym, 0)),
        })
    if mrows:
        sb.table("pms_monthly").upsert(mrows, on_conflict="year_month").execute()


def upsert_pms_reservations(reservations: list):
    """
    予約レコードリストをSupabaseに保存（ブッキングカーブ用）。
    reservations = [{"checkin": "YYYYMMDD", "booking_date": "YYYYMMDD", "rooms": 1}, ...]
    """
    sb = get_client()
    rows = []
    for r in reservations:
        rows.append({
            "checkin_date": r["checkin"],
            "booking_date": r["booking_date"],
            "rooms":        int(r.get("rooms", 1)),
        })
    if rows:
        # 重複は無視（同じ予約の再アップロード対応）
        sb.table("pms_reservations").upsert(rows, on_conflict="checkin_date,booking_date").execute()


def load_pms_daily() -> tuple[dict, dict, dict, dict]:
    """
    Supabaseから全期間の日次・月次データを取得してrm_logicが使う形式で返す。
    戻り値: (daily, {}, monthly_rev, room_monthly, monthly_guests)
    """
    sb = get_client()

    # pms_daily
    res = sb.table("pms_daily").select("stay_date,rooms").execute()
    daily = {}
    for row in res.data:
        dt = datetime.strptime(row["stay_date"], "%Y-%m-%d")
        daily[dt] = row["rooms"]

    # pms_monthly
    res2 = sb.table("pms_monthly").select("*").execute()
    monthly_rev    = {}
    room_monthly   = {}
    monthly_guests = {}
    for row in res2.data:
        ym = row["year_month"]
        monthly_rev[ym]    = row.get("revenue", 0)
        room_monthly[ym]   = row.get("nights", 0)
        monthly_guests[ym] = row.get("guests", 0)

    return daily, monthly_rev, room_monthly, monthly_guests


def load_pms_reservations() -> list:
    """ブッキングカーブ用の予約レコードをSupabaseから取得"""
    sb = get_client()
    res = sb.table("pms_reservations").select("checkin_date,booking_date,rooms").execute()
    return res.data


# ============================================================
# 競合価格 保存・取得
# ============================================================

def upsert_comp_prices(rows: list):
    """
    競合価格をSupabaseに保存。
    rows = [{"fetch_date": "YYYY/MM/DD", "target_date": "YYYY/MM/DD",
             "facility_name": "...", "price": 12345}, ...]
    """
    sb = get_client()
    records = []
    for r in rows:
        records.append({
            "fetch_date":    r["fetch_date"],
            "target_date":   r["target_date"],
            "facility_name": r["facility_name"],
            "price":         r["price"] if isinstance(r["price"], int) else None,
        })
    if records:
        sb.table("comp_prices").upsert(
            records,
            on_conflict="fetch_date,target_date,facility_name"
        ).execute()


def load_comp_prices_latest() -> dict:
    """
    最新取得日の競合価格を {対象日: {施設名: 価格}} 形式で返す（rm_logicと同形式）。
    """
    sb = get_client()
    # 最新の取得日を取得
    res = sb.table("comp_prices").select("fetch_date").order("fetch_date", desc=True).limit(1).execute()
    if not res.data:
        return {}
    latest = res.data[0]["fetch_date"]

    # その日のデータを全取得
    res2 = sb.table("comp_prices").select("*").eq("fetch_date", latest).execute()
    comp_prices = defaultdict(dict)
    for row in res2.data:
        comp_prices[row["target_date"]][row["facility_name"]] = row["price"]
    return dict(comp_prices)


def load_comp_history() -> list:
    """競合価格の全履歴を返す（Tab4用）"""
    sb = get_client()
    res = sb.table("comp_prices").select("*").order("fetch_date", desc=True).execute()
    return res.data


# ============================================================
# DB接続チェック
# ============================================================

def is_db_available() -> bool:
    """Supabaseに接続できるか確認"""
    try:
        get_client()
        return True
    except Exception:
        return False
