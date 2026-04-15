create extension if not exists pgcrypto;

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = timezone('utc', now());
  return new;
end;
$$;

create table if not exists public.material_catalog (
  id uuid primary key default gen_random_uuid(),
  division text,
  category text not null unique,
  retail_price numeric(12, 2) not null check (retail_price >= 0),
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now())
);

create table if not exists public.quotes (
  id uuid primary key default gen_random_uuid(),
  client_name text not null,
  project_name text,
  quote_reference text,
  discount numeric(12, 2) not null default 0 check (discount >= 0),
  notes text,
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now())
);

create table if not exists public.quote_materials (
  id uuid primary key default gen_random_uuid(),
  quote_id uuid not null references public.quotes(id) on delete cascade,
  material_catalog_id uuid references public.material_catalog(id) on delete set null,
  division text,
  category text not null,
  retail_price numeric(12, 2) not null check (retail_price >= 0),
  asking_price numeric(12, 2) not null check (asking_price >= 0),
  pocket numeric(12, 2) generated always as (asking_price - retail_price) stored,
  sort_order integer not null default 0,
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now()),
  unique (quote_id, category)
);

create table if not exists public.quote_measurements (
  id uuid primary key default gen_random_uuid(),
  quote_id uuid not null references public.quotes(id) on delete cascade,
  quote_material_id uuid references public.quote_materials(id) on delete set null,
  room_section text,
  label text not null,
  width_mm numeric(12, 2) not null check (width_mm > 0),
  height_mm numeric(12, 2) not null check (height_mm > 0),
  material_label text not null,
  asking_price numeric(12, 2) not null check (asking_price >= 0),
  raw_sqft integer generated always as (
    round(((width_mm / 1000.0) * (height_mm / 1000.0)) * 10.76)
  ) stored,
  billed_sqft integer generated always as (
    greatest(
      round(((width_mm / 1000.0) * (height_mm / 1000.0)) * 10.76),
      15
    )
  ) stored,
  minimum_applied boolean generated always as (
    round(((width_mm / 1000.0) * (height_mm / 1000.0)) * 10.76) < 15
  ) stored,
  line_cost numeric(12, 2) generated always as (
    greatest(
      round(((width_mm / 1000.0) * (height_mm / 1000.0)) * 10.76),
      15
    ) * asking_price
  ) stored,
  sort_order integer not null default 0,
  created_at timestamptz not null default timezone('utc', now()),
  updated_at timestamptz not null default timezone('utc', now())
);

create index if not exists idx_quote_materials_quote_id
  on public.quote_materials (quote_id);

create index if not exists idx_quote_measurements_quote_id
  on public.quote_measurements (quote_id);

create index if not exists idx_quote_measurements_quote_material_id
  on public.quote_measurements (quote_material_id);

drop trigger if exists trg_material_catalog_updated_at on public.material_catalog;
create trigger trg_material_catalog_updated_at
before update on public.material_catalog
for each row
execute function public.set_updated_at();

drop trigger if exists trg_quotes_updated_at on public.quotes;
create trigger trg_quotes_updated_at
before update on public.quotes
for each row
execute function public.set_updated_at();

drop trigger if exists trg_quote_materials_updated_at on public.quote_materials;
create trigger trg_quote_materials_updated_at
before update on public.quote_materials
for each row
execute function public.set_updated_at();

drop trigger if exists trg_quote_measurements_updated_at on public.quote_measurements;
create trigger trg_quote_measurements_updated_at
before update on public.quote_measurements
for each row
execute function public.set_updated_at();

alter table public.material_catalog enable row level security;
alter table public.quotes enable row level security;
alter table public.quote_materials enable row level security;
alter table public.quote_measurements enable row level security;

drop policy if exists "authenticated users can read material catalog" on public.material_catalog;
create policy "authenticated users can read material catalog"
on public.material_catalog
for select
to authenticated
using (true);

drop policy if exists "authenticated users can manage material catalog" on public.material_catalog;
create policy "authenticated users can manage material catalog"
on public.material_catalog
for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can manage quotes" on public.quotes;
create policy "authenticated users can manage quotes"
on public.quotes
for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can manage quote materials" on public.quote_materials;
create policy "authenticated users can manage quote materials"
on public.quote_materials
for all
to authenticated
using (true)
with check (true);

drop policy if exists "authenticated users can manage quote measurements" on public.quote_measurements;
create policy "authenticated users can manage quote measurements"
on public.quote_measurements
for all
to authenticated
using (true)
with check (true);
