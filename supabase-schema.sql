-- ============================================================
-- ShipX Billing — Supabase schema
-- Paste this whole file into: Supabase dashboard → SQL Editor → New query → Run
-- Safe to re-run (uses IF NOT EXISTS / OR REPLACE).
-- ============================================================

-- 1) Shared rate cards (one row per service card) -------------
create table if not exists public.rate_cards (
  service_id  text not null,
  card_id     text not null default 'main',
  rows        jsonb not null default '[]'::jsonb,
  updated_at  timestamptz not null default now(),
  updated_by  text,
  primary key (service_id, card_id)
);

-- 2) Shared billing records (one row per saved run) -----------
create table if not exists public.billing_records (
  id           uuid primary key default gen_random_uuid(),
  service_id   text not null,
  service      text,
  month        text,
  totals       jsonb,
  by_customer  jsonb,
  lines        jsonb,
  columns      jsonb,
  currency     text default '$',
  review_count int  default 0,
  files        jsonb default '[]'::jsonb,
  saved_by     text,
  saved_at     timestamptz not null default now()
);

create index if not exists billing_records_saved_at_idx on public.billing_records (saved_at desc);
create index if not exists billing_records_service_idx  on public.billing_records (service_id);

-- 3) Row-Level Security: only logged-in users, but all of them share the data
alter table public.rate_cards      enable row level security;
alter table public.billing_records enable row level security;

-- rate_cards policies
drop policy if exists rate_cards_rw on public.rate_cards;
create policy rate_cards_rw on public.rate_cards
  for all to authenticated using (true) with check (true);

-- billing_records policies
drop policy if exists billing_records_rw on public.billing_records;
create policy billing_records_rw on public.billing_records
  for all to authenticated using (true) with check (true);

-- 4) Realtime: broadcast changes to all connected clients
alter publication supabase_realtime add table public.rate_cards;
alter publication supabase_realtime add table public.billing_records;

-- ============================================================
-- After running this:
--   • Auth → Providers → Email: keep enabled.
--   • Auth → Providers → Email → DISABLE "Allow new users to sign up"
--     (so only invited teammates can log in).
--   • Auth → Users → "Invite user" (or "Add user") for each teammate + yourself.
-- ============================================================
