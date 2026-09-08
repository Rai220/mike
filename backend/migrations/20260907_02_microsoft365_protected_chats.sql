-- Migration date: 2026-09-07
-- Protected corporate history never enters ordinary chats, document exports or sharing.
begin;

-- Upgraded installs may retain credentials after membership revocation. Remove
-- those orphans before enforcing immediate token/state/history cascades.
delete from public.user_microsoft365_connections c
  where not exists (
    select 1 from public.org_members m where m.org_id = c.org_id and m.user_id = c.user_id
  );
alter table public.user_microsoft365_connections
  drop constraint if exists user_microsoft365_connections_membership_fkey;
alter table public.user_microsoft365_connections
  add constraint user_microsoft365_connections_membership_fkey
  foreign key (org_id, user_id) references public.org_members(org_id, user_id) on delete cascade;

create table if not exists public.microsoft365_chats (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  connection_id uuid not null references public.user_microsoft365_connections(id) on delete cascade,
  org_id uuid not null,
  tenant_id uuid not null,
  generation uuid not null,
  payload_ciphertext text not null check (octet_length(payload_ciphertext) <= 1500000),
  version uuid not null default gen_random_uuid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  expires_at timestamptz not null,
  foreign key (org_id, user_id) references public.org_members(org_id, user_id) on delete cascade,
  check (expires_at > created_at and expires_at <= created_at + interval '90 days')
);
create index if not exists microsoft365_chats_owner_idx on public.microsoft365_chats(user_id, connection_id, updated_at desc);
create index if not exists microsoft365_chats_expiry_idx on public.microsoft365_chats(expires_at);
alter table public.microsoft365_chats enable row level security;
revoke all on public.microsoft365_chats from public, anon, authenticated;
grant select, insert, update, delete on public.microsoft365_chats to service_role;

create or replace function public.guard_microsoft365_chat_write()
returns trigger language plpgsql security definer set search_path = public, pg_temp as $$
begin
  if tg_op = 'UPDATE' and (
    new.id <> old.id or new.user_id <> old.user_id or new.connection_id <> old.connection_id or
    new.org_id <> old.org_id or new.tenant_id <> old.tenant_id or new.generation <> old.generation or
    new.created_at <> old.created_at or new.expires_at > old.expires_at
  ) then
    raise exception 'Protected chat identity and retention cannot be extended' using errcode = '42501';
  end if;
  -- Serializes writes against disconnect/account-generation changes, closing the app check/write gap.
  perform 1 from public.user_microsoft365_connections c
    where c.id = new.connection_id and c.user_id = new.user_id and c.org_id = new.org_id
      and c.tenant_id = new.tenant_id and c.oauth_generation = new.generation
      and c.enabled and c.status = 'connected'
    for share;
  if not found or new.expires_at <= now() then
    raise exception 'Protected chat connection is unavailable' using errcode = '42501';
  end if;
  return new;
end;
$$;
revoke all on function public.guard_microsoft365_chat_write() from public, anon, authenticated;
drop trigger if exists microsoft365_chat_write_guard on public.microsoft365_chats;
create trigger microsoft365_chat_write_guard before insert or update on public.microsoft365_chats
  for each row execute function public.guard_microsoft365_chat_write();

create or replace function public.cleanup_microsoft365_chats(retention_days integer default 90)
returns integer language plpgsql security definer set search_path = public, pg_temp as $$
declare removed integer;
begin
  if retention_days < 1 or retention_days > 90 or retention_days is null then
    raise exception 'Invalid retention policy' using errcode = '22023';
  end if;
  delete from public.microsoft365_oauth_states where expires_at <= now();
  delete from public.microsoft365_chats h where h.expires_at <= now()
    or h.created_at <= now() - make_interval(days => retention_days)
    or not exists (
      select 1 from public.user_microsoft365_connections c where c.id = h.connection_id
        and c.user_id = h.user_id and c.org_id = h.org_id and c.tenant_id = h.tenant_id
        and c.oauth_generation = h.generation and c.enabled
    );
  get diagnostics removed = row_count;
  return removed;
end;
$$;
revoke all on function public.cleanup_microsoft365_chats(integer) from public, anon, authenticated;
grant execute on function public.cleanup_microsoft365_chats(integer) to service_role;

commit;
