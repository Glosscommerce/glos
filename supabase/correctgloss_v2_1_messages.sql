-- CORRECTGLOSS V2.1: messages et commentaires
-- À exécuter dans Supabase > SQL Editor du projet CORRECTGLOSS.
-- Les messages sont insérés publiquement, mais leur lecture reste réservée au serveur administrateur.

create extension if not exists pgcrypto;

create table if not exists public.app_messages (
    id uuid primary key default gen_random_uuid(),
    created_at timestamptz not null default now(),
    sender_name text not null default 'Utilisateur',
    sender_email text,
    message_type text not null default 'message'
        check (message_type in ('message', 'commentaire', 'reclamation', 'suggestion')),
    message text not null
        check (char_length(message) between 5 and 3000),
    status text not null default 'nouveau'
        check (status in ('nouveau', 'lu', 'traite'))
);

alter table public.app_messages enable row level security;

drop policy if exists "Public peut envoyer un message" on public.app_messages;
create policy "Public peut envoyer un message"
on public.app_messages
for insert
to anon, authenticated
with check (
    char_length(message) between 5 and 3000
    and char_length(sender_name) between 1 and 100
    and (sender_email is null or char_length(sender_email) <= 254)
);

-- Aucun droit SELECT/UPDATE/DELETE public : le tableau de bord utilisera
-- SUPABASE_SERVICE_ROLE_KEY côté serveur Streamlit uniquement.
