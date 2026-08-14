-- Optional per-line control type (SPLIT or FULL).
alter table public.quote_measurements
  add column if not exists control text;

comment on column public.quote_measurements.control is
  'Optional control type (SPLIT or FULL). Blank allowed.';
