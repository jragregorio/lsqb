alter table public.quotes
  add column if not exists installation_amount numeric(12, 2) not null default 0 check (installation_amount >= 0),
  add column if not exists installation_is_free boolean not null default false,
  add column if not exists steam_amount numeric(12, 2) not null default 0 check (steam_amount >= 0),
  add column if not exists steam_is_free boolean not null default false;

update public.quotes
set
  installation_amount = install_steam_amount,
  installation_is_free = install_steam_is_free
where installation_amount = 0
  and installation_is_free = false
  and (install_steam_amount <> 0 or install_steam_is_free = true);
