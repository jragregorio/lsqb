alter table public.quote_measurements
  add column if not exists panel_count integer;

comment on column public.quote_measurements.panel_count is
  'Optional sewing panel count for labor-print rollup. Null/blank means the line is excluded from sewing labor.';
