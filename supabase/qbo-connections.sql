create extension if not exists pgcrypto;

create table if not exists public.qbo_connections (
  id uuid primary key default gen_random_uuid(),
  realm_id text not null unique,
  company_name text,
  environment text not null default 'sandbox',
  encrypted_access_token text not null,
  encrypted_refresh_token text not null,
  access_token_expires_at timestamptz,
  refresh_token_expires_at timestamptz,
  connected_by text,
  disconnected_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists qbo_connections_realm_id_idx
on public.qbo_connections (realm_id);

create index if not exists qbo_connections_environment_idx
on public.qbo_connections (environment);
