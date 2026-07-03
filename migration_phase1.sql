-- ============================================================
-- Migration Phase 1 — Fondations Auth + Persistance
-- À exécuter dans Supabase : SQL Editor → New query → coller → Run
-- ============================================================

-- Table profils utilisateur (étend auth.users géré par Supabase)
create table if not exists public.profiles (
  id          uuid references auth.users on delete cascade primary key,
  email       text not null,
  full_name   text,
  company     text default 'Savills',
  role        text default 'expert',
  created_at  timestamptz default now(),
  updated_at  timestamptz default now()
);

-- Row Level Security : chaque utilisateur ne voit QUE son propre profil
alter table public.profiles enable row level security;

create policy "profiles_select_own"
  on public.profiles for select
  using (auth.uid() = id);

create policy "profiles_update_own"
  on public.profiles for update
  using (auth.uid() = id);

-- Création automatique du profil à l'inscription
create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer set search_path = public
as $$
begin
  insert into public.profiles (id, email, full_name)
  values (new.id, new.email, coalesce(new.raw_user_meta_data->>'full_name', ''));
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute procedure public.handle_new_user();

-- ============================================================
-- Préparation Phase 2 (dossiers) — table créée mais pas encore utilisée
-- Évite une seconde migration à venir ; sans risque si vide.
-- ============================================================
create table if not exists public.dossiers (
  id           uuid primary key default gen_random_uuid(),
  user_id      uuid references auth.users on delete cascade not null,
  nom_client   text,
  adresse      text,
  type_bien    text,
  statut       text default 'brouillon',
  data         jsonb default '{}'::jsonb,   -- contenu libre (comparables, calculs, etc.)
  created_at   timestamptz default now(),
  updated_at   timestamptz default now()
);

alter table public.dossiers enable row level security;

create policy "dossiers_all_own"
  on public.dossiers for all
  using (auth.uid() = user_id)
  with check (auth.uid() = user_id);

-- Index pour les recherches par utilisateur (Phase 2)
create index if not exists idx_dossiers_user_id on public.dossiers(user_id);
