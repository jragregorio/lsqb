-- Per-quote disclaimer printed on the contract PDF after the quote totals.
alter table public.quotes
  add column if not exists disclaimer_enabled boolean not null default false;

alter table public.quotes
  add column if not exists disclaimer_text text not null default '';
