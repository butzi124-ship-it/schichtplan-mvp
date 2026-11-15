-- ISOLATED REVIEW ONLY. No deployment registration, operator enrollment or Auth adapter.
begin;
set local lock_timeout='5s';
set local statement_timeout='30s';
do $$begin
 if not (pg_catalog.current_database()='humbel_isolated_test' or
  (pg_catalog.current_database()='postgres' and pg_catalog.version() like '%(PGlite %'))
  or length(coalesce(current_setting('humbel.isolated_sentinel',true),''))<32
 then raise exception 'ISOLATED 8B2 ONLY' using errcode='42501';end if;
 if pg_catalog.to_regclass('humbel_permissions.installation') is null
 then raise exception 'Verified 8B1 required' using errcode='55000';end if;
end$$;
create schema humbel_identity_admin;
revoke all on schema humbel_identity_admin from public,anon,authenticated,service_role;
create table humbel_identity_admin.runtime(singleton boolean primary key check(singleton),enabled boolean not null default false);
insert into humbel_identity_admin.runtime values(true,false);
-- EMPTY, independent DB-login authority. This is NOT an application role decision.
create table humbel_identity_admin.operators(
 role_oid oid primary key,role_name name not null unique,
 actor_user_id uuid not null unique references humbel_permissions.identities(user_id),
 binding_version bigint not null check(binding_version>0),authorization_epoch bigint not null check(authorization_epoch>0),
 enabled boolean not null default false,revoked boolean not null default false,
 valid_from timestamptz not null,valid_until timestamptz not null,
 evidence_ref uuid not null,check(valid_until>valid_from and isfinite(valid_from) and isfinite(valid_until))
);
create table humbel_identity_admin.scopes(
 operator_oid oid not null references humbel_identity_admin.operators(role_oid),
 target_id uuid not null references humbel_permissions.identities(user_id),
 action text not null check(action in('identity_activate','identity_deactivate','epoch_bump','session_allow','session_revoke','context_check')),
 enabled boolean not null default false,revoked boolean not null default false,
 valid_from timestamptz not null,valid_until timestamptz not null,evidence_ref uuid not null,
 primary key(operator_oid,target_id,action),check(valid_until>valid_from and isfinite(valid_from) and isfinite(valid_until))
);
-- EMPTY: only a separately approved trusted Auth verifier may supply evidence.
-- No token/claims payload or function accepting client assertions is provided.
create table humbel_identity_admin.session_attestations(
 session_id uuid primary key,user_id uuid not null references humbel_permissions.identities(user_id),
 binding_version bigint not null check(binding_version>0),authorization_epoch bigint not null check(authorization_epoch>0),
 verified_from timestamptz not null,verified_until timestamptz not null,evidence_ref uuid not null,
 check(verified_until>verified_from and isfinite(verified_from) and isfinite(verified_until))
);
create table humbel_identity_admin.identity_versions(
 user_id uuid primary key references humbel_permissions.identities(user_id),revision bigint not null check(revision>0)
);
create table humbel_identity_admin.requests(
 request_id uuid primary key,operator_oid oid not null references humbel_identity_admin.operators(role_oid),actor_user_id uuid not null references humbel_permissions.identities(user_id),
 target_id uuid not null references humbel_permissions.identities(user_id),payload jsonb not null,result jsonb not null
);
create table humbel_identity_admin.session_tombstones(
 session_id uuid primary key,user_id uuid not null references humbel_permissions.identities(user_id),
 revoked_at timestamptz not null,request_id uuid not null references humbel_identity_admin.requests(request_id) deferrable initially deferred
);
create table humbel_identity_admin.audit(
 request_id uuid primary key references humbel_identity_admin.requests(request_id),
 operator_oid oid not null references humbel_identity_admin.operators(role_oid),actor_user_id uuid not null references humbel_permissions.identities(user_id),target_id uuid not null,
 action text not null,session_id uuid,event_at timestamptz not null,
 before_revision bigint not null,after_revision bigint not null,
 before_epoch bigint not null,after_epoch bigint not null,
 operator_evidence uuid not null,scope_evidence uuid not null,outcome text not null check(outcome='applied'),
 check(after_revision=before_revision+1),check(after_epoch>=before_epoch)
);
create function humbel_identity_admin.immutable_event() returns trigger language plpgsql set search_path='' as $$
begin raise exception 'Append-only security history' using errcode='42501';end$$;
do $$declare t text;begin
 foreach t in array array['requests','session_tombstones','audit','session_attestations'] loop
  execute format('create trigger immutable_rows before update or delete on humbel_identity_admin.%I for each row execute function humbel_identity_admin.immutable_event()',t);
  execute format('create trigger immutable_truncate before truncate on humbel_identity_admin.%I for each statement execute function humbel_identity_admin.immutable_event()',t);
 end loop;
end$$;
-- Bounded legacy read/lock bridge; no public helper/policy is replaced.
create function humbel_identity_admin.binding_ok(p_user uuid) returns boolean
language plpgsql volatile security definer set search_path='' set lock_timeout='3s' as $$
declare eid uuid;
begin
 perform 1 from humbel_security.principals p where p.user_id=p_user for share;
 select p.employee_id into eid from humbel_security.principals p where p.user_id=p_user;
 perform 1 from public.employees e where e.id=eid for share;
 return exists(select 1 from humbel_permissions.identities i
  join humbel_security.principals p on p.user_id=i.user_id and p.employee_id=i.employee_id
  join public.employees e on e.id=i.employee_id
  where i.user_id=p_user and p.enabled and p.valid_until>clock_timestamp()
  and e.is_active and e.auth_user_id=p_user and e.role=p.role
  and e.department_id is not distinct from p.department_id
  and e.department_id is not distinct from i.primary_department_id
  and i.valid_from<=clock_timestamp() and (i.valid_until is null or i.valid_until>clock_timestamp())
  and (select count(*) from public.employees duplicate where duplicate.auth_user_id=p_user)=1);
end$$;
create function humbel_identity_admin.authority(p_action text,p_target uuid)
returns table(operator_oid oid,actor_user_id uuid,operator_evidence uuid,scope_evidence uuid)
language plpgsql volatile security definer set search_path='' set lock_timeout='3s' as $$
begin
 perform 1 from humbel_identity_admin.runtime r where r.singleton for share;
 perform 1 from humbel_permissions.runtime r where r.singleton for share;
 if current_setting('transaction_isolation')<>'read committed'
  or coalesce(current_setting('role',true),'none')<>'none'
  or not exists(select 1 from humbel_identity_admin.runtime r where r.singleton and r.enabled)
  or exists(select 1 from humbel_permissions.runtime r where r.enabled)
 then raise exception 'Management boundary denied' using errcode='42501';end if;
 -- session_user is the DB-authenticated login, never JWT sub/role/current_user.
 return query select o.role_oid,o.actor_user_id,o.evidence_ref,s.evidence_ref
 from humbel_identity_admin.operators o
 join humbel_permissions.identities actor on actor.user_id=o.actor_user_id and actor.binding_version=o.binding_version and actor.authorization_epoch=o.authorization_epoch and actor.enabled and not actor.blocked
 join pg_catalog.pg_roles r on r.oid=o.role_oid and r.rolname=o.role_name and r.rolname=session_user
 join humbel_identity_admin.scopes s on s.operator_oid=o.role_oid
 where r.rolcanlogin and not r.rolsuper and not r.rolbypassrls and not r.rolcreaterole and not r.rolcreatedb
 and r.rolname not in('authenticator','postgres','supabase_admin','anon','authenticated','service_role')
 and o.enabled and not o.revoked and o.valid_from<=clock_timestamp() and o.valid_until>clock_timestamp()
 and s.enabled and not s.revoked and s.target_id=p_target and s.action=p_action
 and s.valid_from<=clock_timestamp() and s.valid_until>clock_timestamp()
 and o.actor_user_id<>p_target for share of o,s;
 if not found then raise exception 'Management authority denied' using errcode='42501';end if;
end$$;
create function humbel_identity_admin.live_session(p_target uuid,p_session uuid) returns boolean
language sql volatile security definer set search_path='' as $$
 select exists(select 1 from humbel_permissions.identities i
 join humbel_permissions.sessions s on s.user_id=i.user_id
 join humbel_identity_admin.session_attestations a on a.session_id=s.session_id and a.user_id=i.user_id
 where i.user_id=p_target and s.session_id=p_session and i.enabled and not i.blocked
 and not s.revoked and s.binding_version=i.binding_version and s.authorization_epoch=i.authorization_epoch
 and s.valid_from<=clock_timestamp() and s.valid_until>clock_timestamp()
 and a.binding_version=i.binding_version and a.authorization_epoch=i.authorization_epoch
 and a.verified_from<=clock_timestamp() and a.verified_until>clock_timestamp()
 and s.valid_until<=a.verified_until
 and not exists(select 1 from humbel_identity_admin.session_tombstones t where t.session_id=p_session))
$$;
create function humbel_identity_admin.command(
 p_action text,p_target uuid,p_session uuid,p_expected_revision bigint,p_request uuid,p_until timestamptz default null)
returns jsonb language plpgsql volatile security definer set search_path='' set lock_timeout='3s' as $$
declare a record;i humbel_permissions.identities;old_receipt humbel_identity_admin.requests;
 proof humbel_identity_admin.session_attestations;rev bigint;epoch_before bigint;payload jsonb;result jsonb;
begin
 if p_action is null or p_action not in('identity_activate','identity_deactivate','epoch_bump','session_allow','session_revoke')
  or p_target is null or p_request is null or p_expected_revision is null or p_expected_revision<0
  or (p_action in('session_allow','session_revoke'))<>(p_session is not null)
  or (p_action='session_allow')<>(p_until is not null) or (p_until is not null and not isfinite(p_until))
 then raise exception 'Invalid management command' using errcode='22023';end if;
 select * into strict a from humbel_identity_admin.authority(p_action,p_target);
 -- Shared fixed order across commands/contexts; recheck AFTER waiting for locks.
 perform 1 from humbel_permissions.identities x where x.user_id in(a.actor_user_id,p_target) order by x.user_id for update;
 perform 1 from humbel_identity_admin.authority(p_action,p_target);
 if not exists(select 1 from humbel_permissions.identities x where x.user_id=a.actor_user_id and x.enabled and not x.blocked)
  or not humbel_identity_admin.binding_ok(a.actor_user_id)
 then raise exception 'Inactive management actor' using errcode='42501';end if;
 select * into strict i from humbel_permissions.identities x where x.user_id=p_target;
 select coalesce((select v.revision from humbel_identity_admin.identity_versions v where v.user_id=p_target),0) into rev;
 payload:=jsonb_build_object('action',p_action,'target',p_target,'session',p_session,'revision',p_expected_revision,'until_epoch',extract(epoch from p_until));
 select * into old_receipt from humbel_identity_admin.requests q where q.request_id=p_request;
 if found then
  if old_receipt.operator_oid<>a.operator_oid or old_receipt.actor_user_id<>a.actor_user_id
   or old_receipt.target_id<>p_target or old_receipt.payload<>payload
  then raise exception 'Conflicting request ID' using errcode='22023';end if;
  if (p_action='session_allow' and (not humbel_identity_admin.binding_ok(p_target) or not humbel_identity_admin.live_session(p_target,p_session)))
   or (p_action='identity_activate' and (not i.enabled or i.blocked or not humbel_identity_admin.binding_ok(p_target) or (old_receipt.result->>'authorization_epoch')::bigint<>i.authorization_epoch))
  then raise exception 'Stale success cannot authorize' using errcode='42501';end if;
  perform 1 from humbel_identity_admin.authority(p_action,p_target);
  return old_receipt.result||jsonb_build_object('replayed',true,'current_revision',rev);
 end if;
 if rev<>p_expected_revision then raise exception 'Management version conflict' using errcode='40001';end if;
 epoch_before:=i.authorization_epoch;
 if p_action in('identity_activate','session_allow') and not humbel_identity_admin.binding_ok(p_target)
 then raise exception 'Unverified identity binding' using errcode='42501';end if;
 perform 1 from humbel_identity_admin.authority(p_action,p_target);
 if p_action in('identity_activate','identity_deactivate','epoch_bump') then
  update humbel_permissions.identities x set
   enabled=case when p_action='identity_activate' then true when p_action='identity_deactivate' then false else x.enabled end,
   blocked=case when p_action='identity_activate' then false when p_action='identity_deactivate' then true else x.blocked end,
   authorization_epoch=x.authorization_epoch+1 where x.user_id=p_target;
  update humbel_permissions.sessions s set revoked=true where s.user_id=p_target;
  insert into humbel_identity_admin.session_tombstones
   select s.session_id,p_target,clock_timestamp(),p_request from humbel_permissions.sessions s where s.user_id=p_target
   on conflict(session_id) do nothing;
  -- Never reactivate prior organization/special grants as a side effect.
  update humbel_permissions.module_grants g set enabled=false,revoked=true where g.user_id=p_target;
  update humbel_permissions.organization_roles r set enabled=false where r.user_id=p_target;
 elsif p_action='session_allow' then
  if not i.enabled or i.blocked then raise exception 'Inactive target' using errcode='42501';end if;
  select * into proof from humbel_identity_admin.session_attestations s where s.session_id=p_session for share;
  if not found or proof.user_id<>p_target or proof.binding_version<>i.binding_version or proof.authorization_epoch<>i.authorization_epoch or proof.verified_from>clock_timestamp()
   or proof.verified_until<=clock_timestamp() or p_until<=clock_timestamp() or p_until>proof.verified_until
   or exists(select 1 from humbel_identity_admin.session_tombstones s where s.session_id=p_session)
   or exists(select 1 from humbel_permissions.sessions s where s.session_id=p_session)
  then raise exception 'Session evidence/reuse denied' using errcode='42501';end if;
  perform 1 from humbel_identity_admin.authority(p_action,p_target);
  insert into humbel_permissions.sessions values(p_session,p_target,i.binding_version,i.authorization_epoch,
   clock_timestamp(),p_until,false,'8B2 verified evidence '||proof.evidence_ref::text);
 elsif p_action='session_revoke' then
  perform 1 from humbel_permissions.sessions s where s.session_id=p_session and s.user_id=p_target for update;
  if not found then raise exception 'Unknown target session' using errcode='42501';end if;
  perform 1 from humbel_identity_admin.authority(p_action,p_target);
  insert into humbel_identity_admin.session_tombstones values(p_session,p_target,clock_timestamp(),p_request) on conflict(session_id) do nothing;
  update humbel_permissions.sessions s set revoked=true where s.session_id=p_session and s.user_id=p_target;
 end if;
 insert into humbel_identity_admin.identity_versions values(p_target,rev+1)
  on conflict(user_id) do update set revision=excluded.revision;
 result:=jsonb_build_object('request_id',p_request,'action',p_action,'outcome','applied','revision',rev+1,
  'authorization_epoch',(select x.authorization_epoch from humbel_permissions.identities x where x.user_id=p_target));
 insert into humbel_identity_admin.requests values(p_request,a.operator_oid,a.actor_user_id,p_target,payload,result);
 insert into humbel_identity_admin.audit values(p_request,a.operator_oid,a.actor_user_id,p_target,p_action,p_session,
  clock_timestamp(),rev,rev+1,epoch_before,(result->>'authorization_epoch')::bigint,a.operator_evidence,a.scope_evidence,'applied');
 return result;
end$$;
-- Internal server-only context check. IDs are resources, NOT authority claims.
-- Never hand this result to a later unprotected write as an authorization token.
create function humbel_identity_admin.server_context(p_target uuid,p_session uuid) returns jsonb
language plpgsql volatile security definer set search_path='' set lock_timeout='3s' as $$
declare a record;i humbel_permissions.identities;
begin
 select * into strict a from humbel_identity_admin.authority('context_check',p_target);
 perform 1 from humbel_permissions.identities x where x.user_id in(a.actor_user_id,p_target) order by x.user_id for share;
 perform 1 from humbel_identity_admin.authority('context_check',p_target);
 if not exists(select 1 from humbel_permissions.identities x where x.user_id=a.actor_user_id and x.enabled and not x.blocked)
  or not humbel_identity_admin.binding_ok(a.actor_user_id) or not humbel_identity_admin.binding_ok(p_target)
  or not humbel_identity_admin.live_session(p_target,p_session)
 then raise exception 'Context denied' using errcode='42501';end if;
 select * into strict i from humbel_permissions.identities x where x.user_id=p_target;
 perform 1 from humbel_identity_admin.authority('context_check',p_target);
 return jsonb_build_object('user_id',i.user_id,'employee_id',i.employee_id,'session_id',p_session,
  'binding_version',i.binding_version,'authorization_epoch',i.authorization_epoch);
end$$;
do $$declare t record;c record;f record;begin
 for t in select tablename from pg_catalog.pg_tables where schemaname='humbel_identity_admin' loop
  execute format('alter table humbel_identity_admin.%I enable row level security',t.tablename);
  execute format('revoke all on table humbel_identity_admin.%I from public,anon,authenticated,service_role',t.tablename);
  for c in select attname from pg_catalog.pg_attribute where attrelid=format('humbel_identity_admin.%I',t.tablename)::regclass and attnum>0 and not attisdropped loop
   execute format('revoke all privileges (%I) on table humbel_identity_admin.%I from public,anon,authenticated,service_role',c.attname,t.tablename);
  end loop;
 end loop;
 for f in select p.oid::regprocedure signature from pg_catalog.pg_proc p join pg_catalog.pg_namespace n on n.oid=p.pronamespace where n.nspname='humbel_identity_admin' loop
  execute format('revoke all on function %s from public,anon,authenticated,service_role',f.signature);
 end loop;
end$$;
commit;
