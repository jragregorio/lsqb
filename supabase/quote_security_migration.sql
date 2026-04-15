alter table public.quotes
add column if not exists owner_user_id uuid references auth.users(id) on delete cascade;

alter table public.quotes
add column if not exists discount_type text;

alter table public.quotes
add column if not exists discount_value numeric(12, 2);

update public.quotes
set
  discount_type = coalesce(discount_type, 'amount'),
  discount_value = coalesce(discount_value, discount, 0);

alter table public.quotes
alter column discount_type set default 'amount';

alter table public.quotes
alter column discount_value set default 0;

alter table public.quotes
alter column discount_type set not null;

alter table public.quotes
alter column discount_value set not null;

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
