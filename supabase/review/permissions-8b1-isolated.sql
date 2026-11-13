-- REVIEW/ISOLATED TEST ONLY. Not registered under supabase/migrations.
-- Additive to the verified Step7 package. No public table/policy/helper replacement.
-- No identity enrollment, organizational appointment, API grants or business actions.
begin;
set local lock_timeout='5s';
set local statement_timeout='30s';
do $$begin
 if not (pg_catalog.current_database()='humbel_isolated_test' or
   (pg_catalog.current_database()='postgres' and pg_catalog.version() like '%(PGlite %'))
   or pg_catalog.length(coalesce(pg_catalog.current_setting('humbel.isolated_sentinel',true),''))<32
 then raise exception 'ISOLATED 8B1 ONLY' using errcode='42501'; end if;
 if pg_catalog.to_regclass('humbel_security.review_step7') is null
 then raise exception 'Verified Step7 baseline required' using errcode='55000'; end if;
end$$;
create schema humbel_permissions;
revoke all on schema humbel_permissions from public,anon,authenticated,service_role;
create table humbel_permissions.runtime (
 singleton boolean primary key default true check(singleton),
 enabled boolean not null default false
);
insert into humbel_permissions.runtime(singleton) values(true);
-- Catalog entries are names only, never assignments or inherited rank.
create table humbel_permissions.roles(code text primary key check(code in(
 'executive','production_manager','department_head','deputy','hr','employee','technical_admin')));
insert into humbel_permissions.roles values('executive'),('production_manager'),('department_head'),
 ('deputy'),('hr'),('employee'),('technical_admin');
create table humbel_permissions.identities (
 user_id uuid primary key references humbel_security.principals(user_id),
 employee_id uuid not null unique references public.employees(id),
 enabled boolean not null default false, blocked boolean not null default false,
 binding_version bigint not null default 1 check(binding_version>0),
 authorization_epoch bigint not null default 1 check(authorization_epoch>0),
 primary_epoch bigint not null default 1 check(primary_epoch>0),
 primary_department_id uuid references public.departments(id),
 valid_from timestamptz not null, valid_until timestamptz,
 evidence text not null check(length(trim(evidence))>0),
 check(valid_until is null or valid_until>valid_from)
);
-- Independent session allowlist, NOT a copy of JWT roles or a claim that auth.sessions
-- has been restored. No session enrollment/synchronization command in this phase.
create table humbel_permissions.sessions (
 session_id uuid primary key, user_id uuid not null references humbel_permissions.identities(user_id),
 binding_version bigint not null check(binding_version>0), authorization_epoch bigint not null check(authorization_epoch>0),
 valid_from timestamptz not null, valid_until timestamptz not null,
 revoked boolean not null default false, evidence text not null check(length(trim(evidence))>0),
 check(valid_until>valid_from)
);
create table humbel_permissions.memberships (
 employee_id uuid not null references public.employees(id), department_id uuid not null references public.departments(id),
 enabled boolean not null default false, revoked boolean not null default false,
 valid_from timestamptz not null, valid_until timestamptz,
 evidence text not null check(length(trim(evidence))>0),
 primary key(employee_id,department_id),check(valid_until is null or valid_until>valid_from)
);
create table humbel_permissions.organization_roles (
 id uuid primary key, user_id uuid not null references humbel_permissions.identities(user_id),
 role_code text not null references humbel_permissions.roles(code), department_id uuid references public.departments(id),
 enabled boolean not null default false, revoked boolean not null default false,
 valid_from timestamptz not null, valid_until timestamptz,
 evidence text not null check(length(trim(evidence))>0),
 unique(id,user_id),unique(id,user_id,department_id,role_code),
 check(role_code not in('department_head','deputy') or department_id is not null),
 check(valid_until is null or valid_until>valid_from)
);
-- One current designated head per department pending O03. History lives in role rows;
-- multiple departments can reference the SAME person. No pointer update in public.
create table humbel_permissions.leaderships (
 department_id uuid primary key references public.departments(id),
 role_assignment_id uuid not null unique,user_id uuid not null,
 role_code text not null default 'department_head' check(role_code='department_head'),
 foreign key(role_assignment_id,user_id,department_id,role_code)
 references humbel_permissions.organization_roles(id,user_id,department_id,role_code)
);
-- Empty registry, no actual module permission profile decided or enabled here.
create table humbel_permissions.capabilities (
 module text not null check(module ~ '^[a-z][a-z0-9_]{0,63}$' and module not in
 ('organization','identity','authorization','acting_logs','retention')),
 action text not null check(action ~ '^[a-z][a-z0-9_]{0,63}$'),
 scope_rule text not null check(scope_rule in('explicit_department','membership')),
 enabled boolean not null default false, evidence text not null check(length(trim(evidence))>0),
 primary key(module,action)
);
create table humbel_permissions.module_grants (
 id uuid primary key,user_id uuid not null references humbel_permissions.identities(user_id),
 role_assignment_id uuid not null,module text not null,action text not null,
 department_id uuid not null references public.departments(id),
 enabled boolean not null default false, revoked boolean not null default false,
 primary_epoch bigint not null check(primary_epoch>0),
 valid_from timestamptz not null,valid_until timestamptz,
 evidence text not null check(length(trim(evidence))>0),
 foreign key(role_assignment_id,user_id) references humbel_permissions.organization_roles(id,user_id),
 foreign key(module,action) references humbel_permissions.capabilities(module,action),
 check(valid_until is null or valid_until>valid_from)
);
-- Only bounded read evaluators are DEFINER, in a private schema with empty path.
-- Operator-owned in the isolated candidate. No API role receives USAGE/EXECUTE.
-- A later least-privilege owner/API/session review is mandatory before exposure.
create function humbel_permissions.actor_context()
returns table(user_id uuid,employee_id uuid,binding_version bigint,authorization_epoch bigint,primary_epoch bigint)
language plpgsql stable security definer set search_path='' as $$
declare uid uuid; sid uuid; claims jsonb;
begin
 if not exists(select 1 from humbel_permissions.runtime r where r.singleton and r.enabled) then return;end if;
 begin
  uid:=auth.uid(); claims:=auth.jwt(); sid:=nullif(claims->>'session_id','')::uuid;
  if uid is null or sid is null or coalesce((claims->>'is_anonymous')::boolean,false) then return;end if;
 exception when invalid_text_representation then return;
 end;
 return query
 select i.user_id,i.employee_id,i.binding_version,i.authorization_epoch,i.primary_epoch
 from humbel_permissions.identities i
 join public.employees e on e.id=i.employee_id
 join humbel_permissions.sessions s on s.user_id=i.user_id and s.session_id=sid
 where i.user_id=uid and i.enabled and not i.blocked and e.is_active
 and e.auth_user_id=uid and i.employee_id=humbel_security.actor_id()
 and (select count(*) from public.employees duplicate where duplicate.auth_user_id=uid)=1
 and e.department_id is not distinct from i.primary_department_id
 and i.valid_from<=statement_timestamp() and (i.valid_until is null or i.valid_until>statement_timestamp())
 and not s.revoked and s.binding_version=i.binding_version and s.authorization_epoch=i.authorization_epoch
 and s.valid_from<=statement_timestamp() and s.valid_until>statement_timestamp();
end$$;
create function humbel_permissions.can_perform(p_module text,p_action text,p_department uuid)
returns boolean language sql stable security definer set search_path='' as $$
 select coalesce((select count(*)=1
 from humbel_permissions.actor_context() a
 join humbel_permissions.module_grants g on g.user_id=a.user_id and g.primary_epoch=a.primary_epoch
 join humbel_permissions.capabilities c on c.module=g.module and c.action=g.action and c.enabled
 join humbel_permissions.organization_roles r on r.id=g.role_assignment_id and r.user_id=a.user_id
 join public.departments d on d.id=g.department_id and d.active
 where p_department is not null and c.module=p_module and c.action=p_action and g.department_id=p_department
 and g.enabled and not g.revoked and g.valid_from<=statement_timestamp() and (g.valid_until is null or g.valid_until>statement_timestamp())
 and r.enabled and not r.revoked and r.valid_from<=statement_timestamp() and (r.valid_until is null or r.valid_until>statement_timestamp())
 and (r.department_id is null or r.department_id=p_department)
 -- Deputy activation needs O04/O05 and is deliberately unavailable, not inferred.
 and r.role_code<>'deputy'
 and (r.role_code<>'department_head' or exists(select 1 from humbel_permissions.leaderships l
  where l.department_id=p_department and l.role_assignment_id=r.id and l.user_id=a.user_id))
 and (select count(*) from humbel_permissions.organization_roles duplicate
  where duplicate.user_id=r.user_id and duplicate.role_code=r.role_code
  and duplicate.department_id is not distinct from r.department_id
  and duplicate.enabled and not duplicate.revoked and duplicate.valid_from<=statement_timestamp()
  and (duplicate.valid_until is null or duplicate.valid_until>statement_timestamp()))=1
 and (c.scope_rule='explicit_department' or exists(select 1 from humbel_permissions.memberships m
  where m.employee_id=a.employee_id and m.department_id=p_department and m.enabled and not m.revoked
  and m.valid_from<=statement_timestamp() and (m.valid_until is null or m.valid_until>statement_timestamp())))
 ),false)
$$;
-- Explicit closure in the same transaction, including broad inherited defaults.
do $$declare t record;c record;f record;begin
 for t in select tablename from pg_catalog.pg_tables where schemaname='humbel_permissions' loop
  execute format('alter table humbel_permissions.%I enable row level security',t.tablename);
  execute format('revoke all on table humbel_permissions.%I from public,anon,authenticated,service_role',t.tablename);
  for c in select a.attname from pg_catalog.pg_attribute a where a.attrelid=format('humbel_permissions.%I',t.tablename)::regclass and a.attnum>0 and not a.attisdropped loop
   execute format('revoke all privileges (%I) on table humbel_permissions.%I from public,anon,authenticated,service_role',c.attname,t.tablename);
  end loop;
 end loop;
 for f in select p.oid::regprocedure signature from pg_catalog.pg_proc p join pg_catalog.pg_namespace n on n.oid=p.pronamespace where n.nspname='humbel_permissions' loop
  execute format('revoke all on function %s from public,anon,authenticated,service_role',f.signature);
 end loop;
end$$;
commit;
