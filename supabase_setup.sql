-- ============================================================
-- ホテル甲子園 RMシステム Supabaseテーブル定義
-- Supabase の SQL Editor にコピペして実行してください
-- ============================================================

-- 1. 日次稼働データ
CREATE TABLE IF NOT EXISTS pms_daily (
    stay_date   DATE PRIMARY KEY,
    rooms       INTEGER NOT NULL DEFAULT 0,
    revenue     INTEGER NOT NULL DEFAULT 0,
    is_monthly  BOOLEAN DEFAULT FALSE,
    updated_at  TIMESTAMP DEFAULT NOW()
);

-- 2. 月次サマリー
CREATE TABLE IF NOT EXISTS pms_monthly (
    year_month  CHAR(6) PRIMARY KEY,  -- 'YYYYMM'
    revenue     INTEGER NOT NULL DEFAULT 0,
    nights      INTEGER NOT NULL DEFAULT 0,
    guests      INTEGER NOT NULL DEFAULT 0,
    updated_at  TIMESTAMP DEFAULT NOW()
);

-- 3. 予約レコード（ブッキングカーブ用）
CREATE TABLE IF NOT EXISTS pms_reservations (
    id           BIGSERIAL PRIMARY KEY,
    checkin_date DATE    NOT NULL,
    booking_date DATE    NOT NULL,
    rooms        INTEGER NOT NULL DEFAULT 1,
    UNIQUE(checkin_date, booking_date)
);

-- 4. 競合価格
CREATE TABLE IF NOT EXISTS comp_prices (
    id            BIGSERIAL PRIMARY KEY,
    fetch_date    CHAR(10) NOT NULL,   -- 'YYYY/MM/DD'
    target_date   CHAR(10) NOT NULL,   -- 'YYYY/MM/DD'
    facility_name TEXT     NOT NULL,
    price         INTEGER,             -- NULL = 満室or取得不可
    updated_at    TIMESTAMP DEFAULT NOW(),
    UNIQUE(fetch_date, target_date, facility_name)
);

-- インデックス
CREATE INDEX IF NOT EXISTS idx_pms_daily_date ON pms_daily(stay_date);
CREATE INDEX IF NOT EXISTS idx_pms_reservations_checkin ON pms_reservations(checkin_date);
CREATE INDEX IF NOT EXISTS idx_comp_prices_fetch ON comp_prices(fetch_date);
CREATE INDEX IF NOT EXISTS idx_comp_prices_target ON comp_prices(target_date);

-- Row Level Security（外部からの不正アクセス防止）
ALTER TABLE pms_daily        ENABLE ROW LEVEL SECURITY;
ALTER TABLE pms_monthly      ENABLE ROW LEVEL SECURITY;
ALTER TABLE pms_reservations ENABLE ROW LEVEL SECURITY;
ALTER TABLE comp_prices      ENABLE ROW LEVEL SECURITY;

-- サービスロールキーからのみ読み書き可能
CREATE POLICY "service_only" ON pms_daily        FOR ALL USING (true);
CREATE POLICY "service_only" ON pms_monthly      FOR ALL USING (true);
CREATE POLICY "service_only" ON pms_reservations FOR ALL USING (true);
CREATE POLICY "service_only" ON comp_prices      FOR ALL USING (true);
