alter table public.quotes
add column if not exists owner_user_id uuid references auth.users(id) on delete cascade;

alter table public.quotes
add column if not exists discount_type text;

alter table public.quotes
add column if not exists discount_value numeric(12, 2);

alter table public.quotes
add column if not exists subtotal_amount numeric(12, 2);

alter table public.quotes
add column if not exists applied_discount_amount numeric(12, 2);

alter table public.quotes
add column if not exists final_total_amount numeric(12, 2);

alter table public.quotes
add column if not exists quote_date date;

alter table public.quotes
add column if not exists project_architect text;

alter table public.quotes
add column if not exists contact_number text;

alter table public.quotes
add column if not exists email_address text;

alter table public.quote_measurements
add column if not exists measurement_type text;

alter table public.quote_measurements
add column if not exists material_code text;

update public.quotes
set
  discount_type = coalesce(discount_type, 'amount'),
  discount_value = coalesce(discount_value, discount, 0);

update public.quotes q
set
  subtotal_amount = coalesce(
    subtotal_amount,
    (
      select coalesce(sum(qm.line_cost), 0)
      from public.quote_measurements qm
      where qm.quote_id = q.id
    ),
    0
  ),
  applied_discount_amount = coalesce(applied_discount_amount, discount, 0),
  final_total_amount = coalesce(
    final_total_amount,
    greatest(
      coalesce(
        (
          select sum(qm.line_cost)
          from public.quote_measurements qm
          where qm.quote_id = q.id
        ),
        0
      ) - coalesce(discount, 0),
      0
    ),
    0
  );

alter table public.quotes
alter column discount_type set default 'amount';

alter table public.quotes
alter column discount_value set default 0;

alter table public.quotes
alter column discount_type set not null;

alter table public.quotes
alter column discount_value set not null;

alter table public.quotes
alter column subtotal_amount set default 0;

alter table public.quotes
alter column applied_discount_amount set default 0;

alter table public.quotes
alter column final_total_amount set default 0;

alter table public.quotes
alter column subtotal_amount set not null;

alter table public.quotes
alter column applied_discount_amount set not null;

alter table public.quotes
alter column final_total_amount set not null;

alter table public.quotes
drop constraint if exists quotes_discount_type_check;

alter table public.quotes
add constraint quotes_discount_type_check
check (discount_type in ('amount', 'percent'));

alter table public.quotes
drop constraint if exists quotes_discount_value_check;

alter table public.quotes
add constraint quotes_discount_value_check
check (discount_value >= 0);

alter table public.quotes
drop constraint if exists quotes_subtotal_amount_check;

alter table public.quotes
add constraint quotes_subtotal_amount_check
check (subtotal_amount >= 0);

alter table public.quotes
drop constraint if exists quotes_applied_discount_amount_check;

alter table public.quotes
add constraint quotes_applied_discount_amount_check
check (applied_discount_amount >= 0);

alter table public.quotes
drop constraint if exists quotes_final_total_amount_check;

alter table public.quotes
add constraint quotes_final_total_amount_check
check (final_total_amount >= 0);

create index if not exists idx_quotes_owner_user_id
  on public.quotes (owner_user_id);

drop policy if exists "authenticated users can manage quotes" on public.quotes;
drop policy if exists "authenticated users can manage quote materials" on public.quote_materials;
drop policy if exists "authenticated users can manage quote measurements" on public.quote_measurements;
drop policy if exists "owners can manage their quotes" on public.quotes;
drop policy if exists "owners can manage quote materials" on public.quote_materials;
drop policy if exists "owners can manage quote measurements" on public.quote_measurements;

create policy "owners can manage their quotes"
on public.quotes
for all
to authenticated
using (auth.uid() = owner_user_id)
with check (auth.uid() = owner_user_id);

create policy "owners can manage quote materials"
on public.quote_materials
for all
to authenticated
using (
  exists (
    select 1
    from public.quotes
    where public.quotes.id = quote_materials.quote_id
      and public.quotes.owner_user_id = auth.uid()
  )
)
with check (
  exists (
    select 1
    from public.quotes
    where public.quotes.id = quote_materials.quote_id
      and public.quotes.owner_user_id = auth.uid()
  )
);

create policy "owners can manage quote measurements"
on public.quote_measurements
for all
to authenticated
using (
  exists (
    select 1
    from public.quotes
    where public.quotes.id = quote_measurements.quote_id
      and public.quotes.owner_user_id = auth.uid()
  )
)
with check (
  exists (
    select 1
    from public.quotes
    where public.quotes.id = quote_measurements.quote_id
      and public.quotes.owner_user_id = auth.uid()
  )
);
