alter table public.quotes
add column if not exists owner_user_id uuid references auth.users(id) on delete cascade;

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
