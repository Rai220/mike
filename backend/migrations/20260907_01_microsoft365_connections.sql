-- Migration date: 2026-09-07
-- Delegated credentials are backend-only. No browser policies are intentional.
begin;

create table if not exists public.user_microsoft365_connections (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  org_id uuid not null references public.organizations(id) on delete cascade,
  tenant_id uuid not null,
  client_id uuid not null,
  enabled boolean not null default true,
  status text not null default 'pending' check (status in ('pending', 'connected', 'reconnect_required')),
  account_id uuid,
  account_label text,
  token_ciphertext text,
  token_expires_at timestamptz,
  oauth_generation uuid not null default gen_random_uuid(),
  version uuid not null default gen_random_uuid(),
  refresh_lease_id uuid,
  refresh_lease_expires_at timestamptz,
  connected_at timestamptz,
  last_checked_at timestamptz,
  created_at timestamptz not null default now(),
  unique (user_id, org_id),
  unique (id, user_id, org_id, tenant_id, client_id),
  check ((refresh_lease_id is null) = (refresh_lease_expires_at is null)),
  check (status <> 'connected' or (account_id is not null and token_ciphertext is not null and token_expires_at is not null))
);

create table if not exists public.microsoft365_oauth_states (
  state_hash text primary key check (state_hash ~ '^[a-f0-9]{64}$'),
  session_hash text not null check (session_hash ~ '^[a-f0-9]{64}$'),
  connection_id uuid not null,
  user_id uuid not null references auth.users(id) on delete cascade,
  org_id uuid not null references public.organizations(id) on delete cascade,
  tenant_id uuid not null,
  client_id uuid not null,
  generation uuid not null,
  verifier_ciphertext text not null,
  expires_at timestamptz not null,
  created_at timestamptz not null default now(),
  foreign key (connection_id, user_id, org_id, tenant_id, client_id)
    references public.user_microsoft365_connections(id, user_id, org_id, tenant_id, client_id) on delete cascade
);
create index if not exists microsoft365_oauth_states_expiry_idx on public.microsoft365_oauth_states(expires_at);
create index if not exists microsoft365_oauth_states_connection_idx on public.microsoft365_oauth_states(connection_id);

alter table public.user_microsoft365_connections enable row level security;
alter table public.microsoft365_oauth_states enable row level security;
revoke all on public.user_microsoft365_connections, public.microsoft365_oauth_states from public, anon, authenticated;
grant select, insert, update, delete on public.user_microsoft365_connections, public.microsoft365_oauth_states to service_role;

commit;
