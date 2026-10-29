-- Sara asszisztens – adatbázis létrehozása és frissítése
-- Új (üres) adatbázisban minden táblát létrehoz, meglévőben csak a hiányzó részeket pótolja.
-- Futtatás: Supabase → SQL Editor → illeszd be az egészet → Run.
-- Többször is lefuttatható, meglévő adatot nem töröl és nem módosít.

-- 0. Alaptáblák új telepítéshez: üres adatbázisban ez hozza létre őket.
--    Ahol már léteznek, ott ez a rész semmit nem változtat.
create table if not exists public.projects (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  status text not null default 'active',
  created_at timestamptz not null default now()
);
create table if not exists public.messages (
  id bigint generated always as identity primary key,
  project_id uuid not null,
  role text not null,
  content text not null default '',
  created_at timestamptz not null default now()
);
create table if not exists public.project_documents (
  id bigint generated always as identity primary key,
  project_id uuid not null,
  name text not null,
  mime_type text,
  size_bytes bigint,
  content_b64 text,
  created_at timestamptz not null default now()
);
create index if not exists messages_project_idx on public.messages (project_id, created_at);
create index if not exists project_documents_project_idx on public.project_documents (project_id, created_at);

-- 1. Projektek: kézi sorrend és a legutóbbi megnyitás ideje
alter table public.projects add column if not exists sort_order double precision;
alter table public.projects add column if not exists last_opened_at timestamptz;
-- 1/c. Bizalmas projekt: a dokumentumai más projektből (Sara kereső kollégájával) nem kereshetők
alter table public.projects add column if not exists confidential boolean not null default false;
-- 1/d. Dokumentumok rendezése: kategória (pl. Szerződés, Műszaki leírás) és elavult jelölés (régi változat:
--      Sara csak kérésre nézi). Amíg ez nem futott le, az alkalmazás rendezés nélkül működik.
alter table public.project_documents add column if not exists category text;
alter table public.project_documents add column if not exists obsolete boolean not null default false;
-- 1/e. Törzsakták (cég, ingatlan, személy, jármű): a projektek táblájában, kind = 'torzs'. A projekt a
--      links listában tartja a hozzá kapcsolt törzsakták azonosítóit. Az app_settings a saját profilt őrzi.
alter table public.projects add column if not exists kind text not null default 'projekt';
alter table public.projects add column if not exists record_type text;
alter table public.projects add column if not exists links jsonb not null default '[]'::jsonb;
create table if not exists public.app_settings (
  key text primary key,
  value jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

-- 1/b. Projektek állapota: a lezáráshoz a "closed" értéket is engedni kell.
--    A régi projects_status_check szabály ezt nem engedte (23514-es hiba lezáráskor).
alter table public.projects drop constraint if exists projects_status_check;
do $$
begin
  if not exists (select 1 from public.projects where status is null or status not in ('active', 'closed')) then
    alter table public.projects
      add constraint projects_status_check check (status in ('active', 'closed'));
  end if;
end $$;

-- 2. Dokumentumok: a fájlból kinyert szöveg (content_text) és annak kereshető,
--    kisbetűs, ékezet nélküli változata (search_text; amíg üres, a dokumentum nincs feldolgozva)
alter table public.project_documents add column if not exists content_text text;
alter table public.project_documents add column if not exists search_text text;
-- 2/b. Dokumentumok: rövid, 5–10 mondatos tartalmi kivonat (feltöltéskor készül)
alter table public.project_documents add column if not exists summary text;

-- 3. Új táblák: összefoglalók, a beszélgetésbe csatolt munkaanyagok és a feladatok.
--    A project_id és a message_id típusa automatikusan a meglévő táblákéhoz igazodik.
do $$
declare
  pid_type text;
  mid_type text;
begin
  select format_type(a.atttypid, a.atttypmod) into pid_type
    from pg_attribute a
   where a.attrelid = 'public.projects'::regclass and a.attname = 'id' and not a.attisdropped;
  select format_type(a.atttypid, a.atttypmod) into mid_type
    from pg_attribute a
   where a.attrelid = 'public.messages'::regclass and a.attname = 'id' and not a.attisdropped;
  if pid_type is null or mid_type is null then
    raise exception 'A projects vagy a messages táblában nincs id oszlop.';
  end if;

  execute format(
    'create table if not exists public.project_summaries (
       id bigint generated always as identity primary key,
       project_id %s not null,
       description text not null default '''',
       content text not null,
       search_text text,
       created_at timestamptz not null default now(),
       updated_at timestamptz
     )', pid_type);

  execute format(
    'create table if not exists public.message_attachments (
       id bigint generated always as identity primary key,
       project_id %s not null,
       message_id %s not null,
       name text not null,
       mime_type text not null,
       size_bytes bigint not null default 0,
       content_b64 text not null,
       created_at timestamptz not null default now()
     )', pid_type, mid_type);

  -- Feladatok: téma » tennivaló, állapottal és megjegyzéssel
  -- (állapotok: open = nyitott, waiting = választ várunk, done = kész, obsolete = elavult, cancelled = törölt)
  execute format(
    'create table if not exists public.project_tasks (
       id bigint generated always as identity primary key,
       project_id %s not null,
       topic text not null default '''',
       title text not null,
       status text not null default ''open'',
       note text not null default '''',
       created_at timestamptz not null default now(),
       updated_at timestamptz
     )', pid_type);

  -- Modellhívások tokenfelhasználása (a költség és a gyorsítótár hasznának elemzéséhez)
  execute format(
    'create table if not exists public.usage_log (
       id bigint generated always as identity primary key,
       project_id %s,
       kind text not null default '''',
       model text not null default '''',
       input_tokens integer not null default 0,
       cache_write_tokens integer not null default 0,
       cache_read_tokens integer not null default 0,
       output_tokens integer not null default 0,
       web_searches integer not null default 0,
       created_at timestamptz not null default now()
     )', pid_type);
end $$;

create index if not exists project_summaries_project_idx on public.project_summaries (project_id, created_at);
create index if not exists message_attachments_project_idx on public.message_attachments (project_id);
create index if not exists message_attachments_message_idx on public.message_attachments (message_id);
create index if not exists project_tasks_project_idx on public.project_tasks (project_id, created_at);
create index if not exists usage_log_created_idx on public.usage_log (created_at);

-- 4. Jogosultságok: a táblákat kizárólag az alkalmazás szerveroldali kulcsa érheti el
--    (Supabase: "service_role", az új elnevezés szerint "secret" kulcs). A nyilvános ("anon", "publishable")
--    kulcs semmit nem olvashat és nem írhat, így az adatbázis az alkalmazás megkerülésével sem érhető el.
--    FONTOS: előbb állítsd át az alkalmazás SUPABASE_KEY titkos beállítását a service_role kulcsra,
--    és csak utána futtasd ezt a szkriptet, különben az alkalmazás nem éri el az adatbázist.
do $$
declare
  t text;
  p record;
begin
  foreach t in array array['projects', 'messages', 'project_documents', 'project_summaries',
                           'message_attachments', 'project_tasks', 'usage_log', 'app_settings']
  loop
    execute format('alter table public.%I enable row level security', t);
    -- minden korábbi szabály törlése (szabály nélkül csak a service_role fér hozzá, mert rá az RLS nem vonatkozik)
    for p in select policyname from pg_policies where schemaname = 'public' and tablename = t
    loop
      execute format('drop policy %I on public.%I', p.policyname, t);
    end loop;
    execute format('revoke all on public.%I from anon, authenticated', t);
    execute format('grant select, insert, update, delete on public.%I to service_role', t);
  end loop;
  -- 2026.10.30-tól a Supabase az új táblákra nem ad magától API-jogot: a fenti grant ezt pótolja. ÚJ TÁBLÁT
  -- mindig vegyél fel a fenti listába, különben az alkalmazás „permission denied” hibát kap.
  execute 'grant usage, select on all sequences in schema public to service_role';
end $$;

-- 5. A Supabase API azonnal vegye észre az új oszlopokat és táblákat
notify pgrst, 'reload schema';
