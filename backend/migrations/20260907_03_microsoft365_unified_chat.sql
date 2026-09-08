-- Migration date: 2026-09-07
-- Keep the ordinary chat identity while all new corporate turns remain encrypted.
begin;

alter table public.chats add column if not exists microsoft365_protected boolean not null default false;
alter table public.microsoft365_chats add column if not exists ordinary_chat_id uuid references public.chats(id) on delete cascade;
create unique index if not exists microsoft365_chats_ordinary_chat_idx on public.microsoft365_chats(ordinary_chat_id) where ordinary_chat_id is not null;

create or replace function public.guard_microsoft365_ordinary_chat()
returns trigger language plpgsql security definer set search_path = public, pg_temp as $$
begin
  if tg_op = 'UPDATE' and old.microsoft365_protected and (
    not new.microsoft365_protected or new.user_id is distinct from old.user_id
  ) then
    raise exception 'Protected chat identity cannot be changed' using errcode = '42501';
  end if;
  if new.microsoft365_protected then
    if new.user_id is null or new.project_id is not null or new.org_id is not null or new.title is distinct from 'Microsoft 365'
      or exists (select 1 from public.chat_access_grants where chat_id = new.id) then
      raise exception 'Protected chat must remain private' using errcode = '42501';
    end if;
  end if;
  return new;
end;
$$;
revoke all on function public.guard_microsoft365_ordinary_chat() from public, anon, authenticated;
drop trigger if exists microsoft365_ordinary_chat_guard on public.chats;
create trigger microsoft365_ordinary_chat_guard before insert or update on public.chats
  for each row execute function public.guard_microsoft365_ordinary_chat();

create or replace function public.guard_microsoft365_ordinary_child()
returns trigger language plpgsql security definer set search_path = public, pg_temp as $$
declare protected boolean;
begin
  -- The same row lock serializes promotion, plaintext writes and sharing grants.
  select microsoft365_protected into protected from public.chats where id = new.chat_id for update;
  if protected then
    raise exception 'Protected chat cannot receive plaintext messages or sharing grants' using errcode = '42501';
  end if;
  return new;
end;
$$;
revoke all on function public.guard_microsoft365_ordinary_child() from public, anon, authenticated;
drop trigger if exists microsoft365_plaintext_message_guard on public.chat_messages;
create trigger microsoft365_plaintext_message_guard before insert or update on public.chat_messages
  for each row execute function public.guard_microsoft365_ordinary_child();
drop trigger if exists microsoft365_sharing_guard on public.chat_access_grants;
create trigger microsoft365_sharing_guard before insert or update on public.chat_access_grants
  for each row execute function public.guard_microsoft365_ordinary_child();

create or replace function public.guard_microsoft365_chat_write()
returns trigger language plpgsql security definer set search_path = public, pg_temp as $$
declare shell public.chats%rowtype;
begin
  if tg_op = 'UPDATE' and (
    new.id <> old.id or new.user_id <> old.user_id or new.connection_id <> old.connection_id or
    new.org_id <> old.org_id or new.tenant_id <> old.tenant_id or new.generation <> old.generation or
    new.created_at <> old.created_at or new.expires_at > old.expires_at or
    new.ordinary_chat_id is distinct from old.ordinary_chat_id
  ) then
    raise exception 'Protected chat identity and retention cannot be extended' using errcode = '42501';
  end if;
  perform 1 from public.user_microsoft365_connections c
    where c.id = new.connection_id and c.user_id = new.user_id and c.org_id = new.org_id
      and c.tenant_id = new.tenant_id and c.oauth_generation = new.generation
      and c.enabled and c.status = 'connected'
    for share;
  if not found or new.expires_at <= now() then
    raise exception 'Protected chat connection is unavailable' using errcode = '42501';
  end if;
  if new.ordinary_chat_id is not null then
    select * into shell from public.chats where id = new.ordinary_chat_id for update;
    if not found or shell.user_id is distinct from new.user_id or shell.project_id is not null or shell.org_id is not null
      or exists (select 1 from public.chat_access_grants where chat_id = shell.id) then
      raise exception 'Protected chat requires a private owned chat' using errcode = '42501';
    end if;
    if tg_op = 'INSERT' and shell.microsoft365_protected then
      raise exception 'Protected history cannot be replaced or retention restarted' using errcode = '42501';
    end if;
    if not shell.microsoft365_protected then
      update public.chats set microsoft365_protected = true, title = 'Microsoft 365' where id = shell.id;
    end if;
  end if;
  return new;
end;
$$;
revoke all on function public.guard_microsoft365_chat_write() from public, anon, authenticated;

commit;
