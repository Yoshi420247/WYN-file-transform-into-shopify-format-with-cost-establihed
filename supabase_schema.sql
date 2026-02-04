-- Supabase Migration: WYN-to-Shopify Transformer
-- Run this in your Supabase SQL Editor (Dashboard → SQL Editor → New query)

-- ============================================================
-- 1. Transform runs — one row per execution
-- ============================================================
CREATE TABLE IF NOT EXISTS transform_runs (
    id              UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    created_at      TIMESTAMPTZ DEFAULT now() NOT NULL,
    source_file     TEXT NOT NULL,
    reference_file  TEXT NOT NULL,
    output_file     TEXT,                      -- Supabase Storage path
    total_products  INT DEFAULT 0,
    total_variants  INT DEFAULT 0,
    total_images    INT DEFAULT 0,
    new_products    INT DEFAULT 0,
    updated_products INT DEFAULT 0,
    removed_products INT DEFAULT 0,
    status          TEXT DEFAULT 'running' CHECK (status IN ('running','completed','failed')),
    error_message   TEXT,
    triggered_by    TEXT DEFAULT 'manual'       -- 'manual', 'github_actions', 'api'
);

-- ============================================================
-- 2. Products — canonical product table (latest state)
-- ============================================================
CREATE TABLE IF NOT EXISTS products (
    id              UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    handle          TEXT NOT NULL UNIQUE,
    title           TEXT NOT NULL,
    body_html       TEXT,
    vendor          TEXT DEFAULT 'What You Need',
    product_type    TEXT,
    tags            TEXT,
    seo_title       TEXT,
    seo_description TEXT,
    status          TEXT DEFAULT 'active',
    source_url      TEXT,                      -- WYN product URL
    wyn_category    TEXT,
    wyn_categories  TEXT,
    first_seen_run  UUID REFERENCES transform_runs(id),
    last_seen_run   UUID REFERENCES transform_runs(id),
    created_at      TIMESTAMPTZ DEFAULT now() NOT NULL,
    updated_at      TIMESTAMPTZ DEFAULT now() NOT NULL
);

-- ============================================================
-- 3. Variants — one row per variant
-- ============================================================
CREATE TABLE IF NOT EXISTS variants (
    id              UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    product_id      UUID REFERENCES products(id) ON DELETE CASCADE NOT NULL,
    sku             TEXT NOT NULL,
    option_name     TEXT DEFAULT 'Title',
    option_value    TEXT DEFAULT 'Default Title',
    wyn_price       NUMERIC(10,2),             -- raw WYN catalogue price
    shopify_cost    NUMERIC(10,2),             -- computed Cost per item
    variant_price   NUMERIC(10,2),             -- Shopify Variant Price (= wyn_price)
    inventory_qty   INT,                       -- parsed from WYN Stock field
    created_at      TIMESTAMPTZ DEFAULT now() NOT NULL,
    updated_at      TIMESTAMPTZ DEFAULT now() NOT NULL,

    UNIQUE(product_id, sku)
);

-- ============================================================
-- 4. Product images — ordered list
-- ============================================================
CREATE TABLE IF NOT EXISTS product_images (
    id              UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    product_id      UUID REFERENCES products(id) ON DELETE CASCADE NOT NULL,
    image_url       TEXT NOT NULL,
    position        INT NOT NULL,
    alt_text        TEXT,
    created_at      TIMESTAMPTZ DEFAULT now() NOT NULL,

    UNIQUE(product_id, position)
);

-- ============================================================
-- 5. Product change log — tracks diffs between runs
-- ============================================================
CREATE TABLE IF NOT EXISTS product_changelog (
    id              UUID DEFAULT gen_random_uuid() PRIMARY KEY,
    run_id          UUID REFERENCES transform_runs(id) ON DELETE CASCADE NOT NULL,
    product_id      UUID REFERENCES products(id) ON DELETE SET NULL,
    handle          TEXT NOT NULL,
    change_type     TEXT NOT NULL CHECK (change_type IN ('new','updated','removed','price_change')),
    field_name      TEXT,                      -- which field changed (NULL for new/removed)
    old_value       TEXT,
    new_value       TEXT,
    created_at      TIMESTAMPTZ DEFAULT now() NOT NULL
);

-- ============================================================
-- Indexes for common queries
-- ============================================================
CREATE INDEX IF NOT EXISTS idx_products_handle ON products(handle);
CREATE INDEX IF NOT EXISTS idx_products_last_run ON products(last_seen_run);
CREATE INDEX IF NOT EXISTS idx_variants_sku ON variants(sku);
CREATE INDEX IF NOT EXISTS idx_variants_product ON variants(product_id);
CREATE INDEX IF NOT EXISTS idx_changelog_run ON product_changelog(run_id);
CREATE INDEX IF NOT EXISTS idx_changelog_type ON product_changelog(change_type);
CREATE INDEX IF NOT EXISTS idx_images_product ON product_images(product_id);

-- ============================================================
-- Storage bucket for CSV outputs (run via Supabase Dashboard
-- if this doesn't work in SQL editor)
-- ============================================================
INSERT INTO storage.buckets (id, name, public)
VALUES ('transform-outputs', 'transform-outputs', true)
ON CONFLICT (id) DO NOTHING;

-- ============================================================
-- Row-Level Security (allow service_role full access)
-- ============================================================
ALTER TABLE transform_runs ENABLE ROW LEVEL SECURITY;
ALTER TABLE products ENABLE ROW LEVEL SECURITY;
ALTER TABLE variants ENABLE ROW LEVEL SECURITY;
ALTER TABLE product_images ENABLE ROW LEVEL SECURITY;
ALTER TABLE product_changelog ENABLE ROW LEVEL SECURITY;

-- Service role can do everything
CREATE POLICY "service_role_all" ON transform_runs FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "service_role_all" ON products FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "service_role_all" ON variants FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "service_role_all" ON product_images FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "service_role_all" ON product_changelog FOR ALL USING (true) WITH CHECK (true);
