-- Quadro de Tarefas — database schema for Supabase (Postgres).
-- Run once in: Supabase Dashboard → SQL Editor → New query → paste → Run.
-- Safe to re-run (idempotent).
--
-- Design (built for offline-first sync):
--   * One "records" table holds every item (client, task, meeting, project, expense),
--     keyed by (user_id, collection, id). The item itself lives in `data` (jsonb),
--     matching the shape the app already stores locally.
--   * `updated_at` = client timestamp (ms) used for last-write-wins merges.
--   * `server_updated_at` = set by the server on every write; devices pull
--     "everything changed since my last sync" with it.
--   * Deletions are soft (`deleted_at`) so other devices learn about them.
--   * Row Level Security: each signed-in user can only ever see/modify their own rows.

-- ───────────────────────────── records ─────────────────────────────
create table if not exists public.records (
  user_id           uuid        not null default auth.uid() references auth.users (id) on delete cascade,
  collection        text        not null check (collection in ('clients','tasks','meetings','projects','expenses','monthlyValues')),
  id                text        not null check (char_length(id) between 1 and 64),
  data              jsonb       not null default '{}'::jsonb,
  updated_at        bigint      not null,                -- client clock, ms since epoch
  deleted_at        bigint,                              -- client clock, ms; null = alive
  server_updated_at timestamptz not null default now(),
  primary key (user_id, collection, id)
);

alter table public.records drop constraint if exists records_collection_check;
alter table public.records
  add constraint records_collection_check
  check (collection in ('clients','tasks','meetings','projects','expenses','monthlyValues'));

create index if not exists records_user_server_updated_idx
  on public.records (user_id, server_updated_at);

-- Always stamp server time on insert/update (clients can't fake it).
create or replace function public.records_touch()
returns trigger
language plpgsql
set search_path = ''
as $$
begin
  new.server_updated_at := now();
  return new;
end;
$$;

drop trigger if exists records_touch on public.records;
create trigger records_touch
  before insert or update on public.records
  for each row execute function public.records_touch();

-- ───────────────────────────── security ─────────────────────────────
alter table public.records enable row level security;

-- Only signed-in users; anonymous visitors get nothing.
revoke all on public.records from anon;
grant select, insert, update, delete on public.records to authenticated;

drop policy if exists "records: select own" on public.records;
create policy "records: select own" on public.records
  for select to authenticated
  using (user_id = (select auth.uid()));

drop policy if exists "records: insert own" on public.records;
create policy "records: insert own" on public.records
  for insert to authenticated
  with check (user_id = (select auth.uid()));

drop policy if exists "records: update own" on public.records;
create policy "records: update own" on public.records
  for update to authenticated
  using (user_id = (select auth.uid()))
  with check (user_id = (select auth.uid()));

drop policy if exists "records: delete own" on public.records;
create policy "records: delete own" on public.records
  for delete to authenticated
  using (user_id = (select auth.uid()));

-- ─────────────────── expenses view (for reports / SQL) ───────────────────
-- Read-only typed view over expense records. security_invoker → RLS of the caller applies.
create or replace view public.expenses_v
with (security_invoker = true) as
select
  r.user_id,
  r.id,
  (r.data->>'date')                          as date,
  r.data->>'category'                        as category,
  coalesce(r.data->>'kind', 'empresa')       as kind,        -- 'pessoal' | 'empresa'
  r.data->>'clientId'                        as client_id,
  r.data->>'supplier'                        as supplier,
  r.data->>'description'                     as description,
  nullif(r.data->>'amount', '')::numeric(12,2) as amount,
  r.updated_at
from public.records r
where r.collection = 'expenses' and r.deleted_at is null;

revoke all on public.expenses_v from anon;
grant select on public.expenses_v to authenticated;

-- ─────────────────── file storage for project documents ───────────────────
-- Private bucket; files live under "<user_id>/<doc_id>". 25 MB max per file.
insert into storage.buckets (id, name, public, file_size_limit)
values ('documents', 'documents', false, 26214400)
on conflict (id) do update set public = false, file_size_limit = excluded.file_size_limit;

drop policy if exists "documents: select own" on storage.objects;
create policy "documents: select own" on storage.objects
  for select to authenticated
  using (bucket_id = 'documents' and (storage.foldername(name))[1] = (select auth.uid())::text);

drop policy if exists "documents: insert own" on storage.objects;
create policy "documents: insert own" on storage.objects
  for insert to authenticated
  with check (bucket_id = 'documents' and (storage.foldername(name))[1] = (select auth.uid())::text);

drop policy if exists "documents: update own" on storage.objects;
create policy "documents: update own" on storage.objects
  for update to authenticated
  using (bucket_id = 'documents' and (storage.foldername(name))[1] = (select auth.uid())::text)
  with check (bucket_id = 'documents' and (storage.foldername(name))[1] = (select auth.uid())::text);

drop policy if exists "documents: delete own" on storage.objects;
create policy "documents: delete own" on storage.objects
  for delete to authenticated
  using (bucket_id = 'documents' and (storage.foldername(name))[1] = (select auth.uid())::text);
