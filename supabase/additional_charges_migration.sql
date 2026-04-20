alter table public.quotes
  add column if not exists delivery_amount numeric(12, 2) not null default 0 check (delivery_amount >= 0),
  add column if not exists delivery_is_free boolean not null default false,
  add column if not exists install_steam_amount numeric(12, 2) not null default 0 check (install_steam_amount >= 0),
  add column if not exists install_steam_is_free boolean not null default false;

