alter table public.quotes
  add column if not exists project_professional_role text not null default 'architect'
    check (project_professional_role in ('architect', 'interior_designer'));
