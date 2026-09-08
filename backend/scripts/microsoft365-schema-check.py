#!/usr/bin/env python3
"""Validate Microsoft 365 SQL using an owned, disposable local Docker container.

Usage: python3 backend/scripts/microsoft365-schema-check.py [--image postgres:16]
Requires a running local Docker daemon and an already cached PostgreSQL image.
No dependencies, image downloads, published ports, real credentials or remote DB.
Auth/role stubs validate PostgreSQL constraints, not the entire Supabase stack.
"""
import argparse
import difflib
import json
import os
from pathlib import Path
import subprocess
import time
import uuid

ROOT = Path(__file__).resolve().parents[2]
MIGRATIONS = [
    ROOT / "backend/migrations/20260907_01_microsoft365_connections.sql",
    ROOT / "backend/migrations/20260907_02_microsoft365_protected_chats.sql",
    ROOT / "backend/migrations/20260907_03_microsoft365_unified_chat.sql",
]
U, V, O, OTHER, C, T, G, H = [
    f"00000000-0000-4000-8000-{n:012d}" for n in range(1, 9)
]
TABLES = ("user_microsoft365_connections", "microsoft365_oauth_states", "microsoft365_chats")


def command(args, data=None, **kwargs):
    result = subprocess.run(args, input=data, capture_output=True, **kwargs)
    if result.returncode:
        raise RuntimeError(result.stderr.decode(errors="replace"))
    return result.stdout


class Fixture:
    def __init__(self, docker):
        self.docker = docker
        self.name = "mike-m365-sql-check-" + uuid.uuid4().hex[:12]
        self.checks = 0

    def sql(self, query):
        data = query.encode() if isinstance(query, str) else query
        return command(self.docker + ["exec", "-i", self.name, "psql", "-U", "postgres",
                                       "-XAtq", "-v", "ON_ERROR_STOP=1"], data).decode().strip()

    def check(self, name, query, expected="t"):
        result = self.sql(query)
        if result != expected:
            raise AssertionError(f"{name}: expected {expected!r}, got {result!r}")
        self.checks += 1
        print("PASS", name, flush=True)

    def rejects(self, name, query, code="42501"):
        self.check(name, "do $$ begin begin " + query + "; raise exception 'unexpected success'; "
                   "exception when sqlstate '" + code + "' then null; end; end $$; select true;")

    def reset(self):
        self.sql("drop schema if exists public cascade; drop schema if exists auth cascade; "
                 "create schema public; create schema auth; "
                 "grant usage on schema public to anon,authenticated,service_role; "
                 "create table auth.users(id uuid primary key,email text,"
                 "raw_user_meta_data jsonb default '{}',encrypted_password text);")


def connection(user=U, org=O, connection_id=C):
    return ("insert into public.user_microsoft365_connections"
            "(id,user_id,org_id,tenant_id,client_id,status,account_id,token_ciphertext,"
            "token_expires_at,oauth_generation) values "
            f"('{connection_id}','{user}','{org}','{T}','{T}','connected','{user}',"
            f"'synthetic',now()+interval '1 hour','{G}')")


def chat(user=U, org=O, tenant=T, generation=G, created="now()"):
    return ("insert into public.microsoft365_chats"
            "(id,user_id,connection_id,org_id,tenant_id,generation,payload_ciphertext,created_at,expires_at) "
            f"values ('{H}','{user}','{C}','{org}','{tenant}','{generation}',"
            f"'synthetic',{created},now()+interval '1 day')")


def state(letter="a", expired=False):
    expiry = "now()-interval '1 hour'" if expired else "now()+interval '1 hour'"
    return ("insert into public.microsoft365_oauth_states(state_hash,session_hash,connection_id,"
            "user_id,org_id,tenant_id,client_id,generation,verifier_ciphertext,expires_at) values "
            f"(repeat('{letter}',64),repeat('b',64),'{C}','{U}','{O}','{T}','{T}','{G}',"
            f"'synthetic',{expiry})")


def seed(fixture):
    fixture.sql(f"insert into auth.users(id) values ('{U}'),('{V}'); "
                f"insert into public.organizations(id,name) values ('{O}','synthetic'),('{OTHER}','synthetic2'); "
                f"insert into public.org_members(org_id,user_id) values ('{O}','{U}'),('{O}','{V}'),('{OTHER}','{U}');")


def assertions(f):
    seed(f)
    f.sql(connection())
    f.rejects("membership-less credentials rejected", connection(V, OTHER, OTHER), "23503")
    for key, kwargs in [("owner", {"user": V}), ("org", {"org": OTHER}),
                        ("tenant", {"tenant": OTHER}), ("generation", {"generation": OTHER})]:
        f.rejects("wrong " + key, chat(**kwargs))
    f.sql(f"update public.user_microsoft365_connections set enabled=false where id='{C}'")
    f.rejects("inactive connection write", chat())
    f.sql(f"update public.user_microsoft365_connections set enabled=true where id='{C}'")
    f.sql(chat())
    f.check("valid active write", f"select count(*)=1 from public.microsoft365_chats where id='{H}'")
    for name, assignment in [("retention extension", "expires_at=expires_at+interval '1 day'"),
                             ("created_at mutation", "created_at=created_at-interval '1 hour'"),
                             ("owner mutation", f"user_id='{V}'")]:
        f.rejects(name, f"update public.microsoft365_chats set {assignment} where id='{H}'")
    for role in ("anon", "authenticated"):
        for table in TABLES:
            f.check(f"{role} no grants {table}",
                    f"select not has_table_privilege('{role}','public.{table}','SELECT,INSERT,UPDATE,DELETE')")
            f.sql(f"set role {role}; do $$ begin begin perform 1 from public.{table}; "
                  "raise exception 'unexpected access'; exception when insufficient_privilege then null; end; end $$;")
            f.check(f"{role} actual SELECT denied {table}", "select true")
        f.check(f"{role} cleanup denied", "select not has_function_privilege("
                f"'{role}','public.cleanup_microsoft365_chats(integer)','EXECUTE')")
    names = ",".join("'" + table + "'" for table in TABLES)
    f.check("RLS enabled", f"select bool_and(relrowsecurity) from pg_class where relname in ({names})")
    f.check("no RLS policies", f"select count(*)=0 from pg_policies where tablename in ({names})")
    for days in ("0", "91", "null"):
        f.rejects("invalid retention " + days, f"perform public.cleanup_microsoft365_chats({days})", "22023")
    f.sql(state())
    f.sql(f"delete from public.org_members where org_id='{O}' and user_id='{U}'")
    for table in TABLES:
        f.check("membership cascades " + table, f"select count(*)=0 from public.{table}")
    f.sql(f"insert into public.org_members(org_id,user_id) values ('{O}','{U}')")
    f.sql(connection())
    f.sql(chat(created="now()-interval '2 days'"))
    f.check("shorter retention cleanup", "select public.cleanup_microsoft365_chats(1)=1")
    f.sql(chat())
    f.sql(f"update public.user_microsoft365_connections set oauth_generation='{OTHER}' where id='{C}'")
    f.check("generation cleanup", "select public.cleanup_microsoft365_chats(90)=1")
    f.sql(f"update public.user_microsoft365_connections set oauth_generation='{G}' where id='{C}'")
    f.sql(chat())
    f.sql(state())
    f.sql(state("c", expired=True))
    f.check("state cleanup keeps chat return count", "select public.cleanup_microsoft365_chats(90)=0")
    f.check("expired state removed", "select count(*)=0 from public.microsoft365_oauth_states where state_hash=repeat('c',64)")
    f.check("fresh state preserved", "select count(*)=1 from public.microsoft365_oauth_states where state_hash=repeat('a',64)")
    f.rejects("oauth owner FK enforced", f"update public.microsoft365_oauth_states set user_id='{V}'", "23503")
    f.sql(f"delete from public.user_microsoft365_connections where id='{C}'")
    for table in TABLES:
        f.check("disconnect cascades " + table, f"select count(*)=0 from public.{table}")


    f.sql(connection())
    shell = "00000000-0000-4000-8000-000000000009"
    other_shell = "00000000-0000-4000-8000-000000000010"
    f.sql(f"insert into public.chats(id,user_id,title) values ('{shell}','{U}','ordinary'),('{other_shell}','{V}','other');")
    f.sql(f"insert into public.chat_messages(chat_id,role,content) values ('{shell}','user',to_jsonb('pre-existing safe message'::text));")
    linked = chat().replace("payload_ciphertext,created_at", "ordinary_chat_id,payload_ciphertext,created_at").replace("'synthetic',now()", f"'{shell}','synthetic',now()")
    f.sql(f"insert into public.chat_access_grants(chat_id,email,role) values ('{shell}','synthetic@example.test','viewer')")
    f.rejects("cannot promote shared ordinary chat", linked)
    f.sql(f"delete from public.chat_access_grants where chat_id='{shell}'")
    f.rejects("cannot bind another owner shell", linked.replace(f"'{shell}'", f"'{other_shell}'"))
    f.sql(linked)
    f.check("promotion marks shell and fixes title", f"select microsoft365_protected and title='Microsoft 365' from public.chats where id='{shell}'")
    f.check("promotion preserves previous messages", f"select count(*)=1 from public.chat_messages where chat_id='{shell}'")
    f.rejects("no plaintext messages after promotion", f"insert into public.chat_messages(chat_id,role,content) values ('{shell}','assistant',to_jsonb('secret'::text))")
    f.rejects("no plaintext edits after promotion", f"update public.chat_messages set content=to_jsonb('secret'::text) where chat_id='{shell}'")
    f.rejects("no share after promotion", f"insert into public.chat_access_grants(chat_id,email,role) values ('{shell}','synthetic@example.test','viewer')")
    f.rejects("protection cannot be disabled", f"update public.chats set microsoft365_protected=false where id='{shell}'")
    f.rejects("protected title cannot carry content", f"update public.chats set title='secret' where id='{shell}'")
    f.rejects("protected owner immutable", f"update public.chats set user_id='{V}' where id='{shell}'")
    f.rejects("binding immutable", f"update public.microsoft365_chats set ordinary_chat_id=null where id='{H}'")
    f.sql(f"delete from public.microsoft365_chats where id='{H}'")
    f.check("protection survives expiry/deletion", f"select microsoft365_protected from public.chats where id='{shell}'")
    f.rejects("deleted history cannot restart retention", linked)
    f.sql(f"delete from public.chats where id='{shell}'")
    f.check("ordinary shell remains deletable", f"select count(*)=0 from public.chats where id='{shell}'")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--image", default="postgres:16", help="Already cached PostgreSQL image")
    parser.add_argument("--docker-context", help="Local Unix-socket context; current context by default")
    parser.add_argument("--baseline", default="HEAD", help="Git ref preceding migrations01/02")
    args = parser.parse_args()
    if os.environ.get("DOCKER_HOST") or os.environ.get("DOCKER_TLS_VERIFY"):
        parser.error("Unset DOCKER_HOST/DOCKER_TLS_VERIFY; only a verified local context is allowed")
    context = args.docker_context or command(["docker", "context", "show"]).decode().strip()
    details = json.loads(command(["docker", "context", "inspect", context]))
    if not details[0]["Endpoints"]["docker"]["Host"].startswith("unix://"):
        parser.error("Remote Docker contexts are not allowed")
    docker = ["docker", "--context", context]
    # Resolve cached image IDs: some Docker Desktop stores list a tag but cannot
    # inspect that tag directly. Run with the verified ID and never pull.
    images = command(docker + ["image", "ls", "--quiet", "--no-trunc", args.image]).decode().splitlines()
    if not images:
        parser.error("Requested PostgreSQL image is not cached locally")
    image_id = images[0]
    command(docker + ["image", "inspect", image_id])
    fixture = Fixture(docker)
    started = False
    try:
        command(docker + ["run", "--detach", "--pull", "never", "--name", fixture.name,
                          "--network", "none", "--memory", "512m", "--cpus", "1",
                          "--tmpfs", "/var/lib/postgresql/data", "-e", "POSTGRES_HOST_AUTH_METHOD=trust", image_id])
        started = True
        for _ in range(30):
            try:
                fixture.sql("select 1")
                break
            except RuntimeError:
                time.sleep(1)
        else:
            raise RuntimeError("Disposable PostgreSQL did not become ready")
        fixture.sql("create role anon; create role authenticated; create role service_role bypassrls;")
        fixture.reset()
        baseline = command(["git", "show", args.baseline + ":backend/schema.sql"], cwd=ROOT)
        fixture.sql(baseline)
        fixture.sql(MIGRATIONS[0].read_bytes())
        # Simulate migration01 credentials whose membership was already revoked.
        seed(fixture)
        fixture.sql(connection())
        fixture.sql(state())
        fixture.sql(f"delete from public.org_members where org_id='{O}' and user_id='{U}'")
        fixture.sql(MIGRATIONS[1].read_bytes())
        for table in TABLES[:2]:
            fixture.check("upgrade purges orphan " + table, f"select count(*)=0 from public.{table}")
        for migration in MIGRATIONS:
            fixture.sql(migration.read_bytes())
            print("PASS replay", migration.name)
        fingerprint_sql = (ROOT / "backend/scripts/schema-fingerprint.sql").read_bytes()
        upgraded = fixture.sql(fingerprint_sql)
        fixture.sql("delete from auth.users; delete from public.organizations;")
        assertions(fixture)
        fixture.reset()
        fixture.sql((ROOT / "backend/schema.sql").read_bytes())
        fresh = fixture.sql(fingerprint_sql)
        if fresh != upgraded:
            print("\n".join(difflib.unified_diff(upgraded.splitlines(), fresh.splitlines(),
                                                fromfile="upgraded", tofile="fresh")))
            raise AssertionError("Fresh and upgraded schemas differ")
        print("PASS identical fingerprints", len(fresh.splitlines()), "lines")
        assertions(fixture)
        print("TOTAL", fixture.checks, "SQL assertions passed; replay and fingerprint passed")
    finally:
        if started:
            command(docker + ["rm", "--force", fixture.name])
            print("Removed disposable container", fixture.name)


if __name__ == "__main__":
    main()
