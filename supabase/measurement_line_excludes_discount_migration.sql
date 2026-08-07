-- Per-line "no discount" flag: line cost counts toward subtotal but not the discount base.
alter table public.quote_measurements
  add column if not exists line_excludes_discount boolean not null default false;

comment on column public.quote_measurements.line_excludes_discount is
  'When true, line cost is included in the subtotal but excluded from the discountable base.';
