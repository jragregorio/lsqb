-- Run in Supabase SQL editor after backup. Replaces sqft GENERATED columns with stored
-- line_cost and adds unit_quantity for quantity-priced lines (e.g. Curtains Motorized).

begin;

alter table public.quote_measurements
  add column if not exists line_cost_backup numeric(12, 2);

update public.quote_measurements
set line_cost_backup = line_cost
where line_cost_backup is null;

alter table public.quote_measurements drop column if exists raw_sqft cascade;
alter table public.quote_measurements drop column if exists billed_sqft cascade;
alter table public.quote_measurements drop column if exists minimum_applied cascade;
alter table public.quote_measurements drop column if exists line_cost cascade;

alter table public.quote_measurements rename column line_cost_backup to line_cost;

alter table public.quote_measurements
  alter column line_cost set not null;

alter table public.quote_measurements
  add column if not exists unit_quantity numeric(12, 4);

comment on column public.quote_measurements.unit_quantity is
  'When set, line_cost = unit_quantity * asking_price; width_mm/height_mm may be 0.';

alter table public.quote_measurements drop constraint if exists quote_measurements_width_mm_check;
alter table public.quote_measurements drop constraint if exists quote_measurements_height_mm_check;

alter table public.quote_measurements alter column width_mm drop not null;
alter table public.quote_measurements alter column height_mm drop not null;

update public.quote_measurements set width_mm = 0 where width_mm is null;
update public.quote_measurements set height_mm = 0 where height_mm is null;

alter table public.quote_measurements alter column width_mm set not null;
alter table public.quote_measurements alter column height_mm set not null;

alter table public.quote_measurements
  add constraint quote_measurements_width_mm_nonneg check (width_mm >= 0);

alter table public.quote_measurements
  add constraint quote_measurements_height_mm_nonneg check (height_mm >= 0);

alter table public.quote_measurements
  add constraint quote_measurements_qty_or_dims check (
    (unit_quantity is not null and unit_quantity > 0)
    or (width_mm > 0 and height_mm > 0)
  );

commit;
