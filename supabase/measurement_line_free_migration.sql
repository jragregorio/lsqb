-- Per-line "complimentary" flag: reference price still stored; excluded from totals when true.
alter table public.quote_measurements
  add column if not exists line_is_free boolean not null default false;

comment on column public.quote_measurements.line_is_free is
  'When true, line reference cost is shown but not added to the quote subtotal.';
