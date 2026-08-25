alter table public.quote_measurements
  add column if not exists is_note boolean not null default false;

alter table public.quote_measurements
  add column if not exists line_note text;

comment on column public.quote_measurements.is_note is
  'When true, this row is a free-form labor-print note (not a billable measurement line).';

comment on column public.quote_measurements.line_note is
  'Optional note text for labor-print PDF display. Used when is_note is true.';

alter table public.quote_measurements
  drop constraint if exists quote_measurements_qty_or_dims;

alter table public.quote_measurements
  add constraint quote_measurements_qty_or_dims check (
    is_note = true
    or (unit_quantity is not null and unit_quantity > 0)
    or (width_mm > 0 and height_mm > 0)
  );
