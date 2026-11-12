-- PROPOSED CANDIDATE ONLY. NOT A MIGRATION. Never run against a hosted database.
-- Existing tables/columns: docs/qa/schema-inventory.json. New private structures
-- below are proposals, not recovered production definitions. No identity backfill.
begin;
set local lock_timeout='5s';
set local statement_timeout='30s';
do $$begin
 if not (pg_catalog.current_database()='humbel_isolated_test' or
   (pg_catalog.current_database()='postgres' and pg_catalog.version() like '%(PGlite %'))
   or pg_catalog.length(coalesce(pg_catalog.current_setting('humbel.isolated_sentinel',true),''))<32
 then raise exception 'ISOLATED SECURITY CANDIDATE ONLY' using errcode='42501'; end if;
end$$;
create schema humbel_security;
revoke all on schema humbel_security from public,anon,authenticated;
revoke create on schema public from public,anon,authenticated;
alter default privileges in schema humbel_security revoke execute on functions from public;
-- Independently reviewed enrollment, never copied from employee.role or JWT metadata.
create table humbel_security.principals(
 user_id uuid primary key references auth.users(id),employee_id uuid not null unique references public.employees(id),
 role text not null check(role in('admin','department_admin','employee','tool_scanner')),
 department_id uuid references public.departments(id),enabled boolean not null default false,
 valid_until timestamptz not null,evidence text not null check(length(trim(evidence))>0)
);
create table humbel_security.delegations(
 user_id uuid not null references humbel_security.principals(user_id),department_id uuid not null references public.departments(id),
 principal_user_id uuid not null references humbel_security.principals(user_id),
 valid_from timestamptz not null,valid_until timestamptz not null,revoked boolean not null default false,
 evidence text not null check(length(trim(evidence))>0),primary key(user_id,department_id),check(valid_until>valid_from)
);
-- Explicit read/publish scopes; no automatic publication right from planning rights.
create table humbel_security.schedule_grants(
 user_id uuid not null references humbel_security.principals(user_id),department_id uuid not null references public.departments(id),
 can_read boolean not null default false,can_publish boolean not null default false,
 valid_from timestamptz not null,valid_until timestamptz not null,revoked boolean not null default false,
 evidence text not null check(length(trim(evidence))>0),primary key(user_id,department_id),check(valid_until>valid_from)
);
-- Empty by default. Unresolved private absence/task workflows require explicit grants.
create table humbel_security.record_grants(
 user_id uuid not null references humbel_security.principals(user_id),department_id uuid not null references public.departments(id),
 table_name text not null,can_read boolean not null default false,can_write boolean not null default false,
 valid_from timestamptz not null,valid_until timestamptz not null,revoked boolean not null default false,
 evidence text not null check(length(trim(evidence))>0),primary key(user_id,department_id,table_name),check(valid_until>valid_from)
);
-- Trusted, explicitly reconciled mapping for name-based legacy records. No automatic backfill.
create table humbel_security.record_bindings(
 table_name text not null,row_id uuid not null,department_id uuid not null references public.departments(id),
 subject_employee_id uuid references public.employees(id),row_version bigint not null default 1 check(row_version>0),
 evidence text not null check(length(trim(evidence))>0),primary key(table_name,row_id)
);
create table humbel_security.publications(
 assignment_id uuid primary key,version bigint not null,published_by uuid not null references auth.users(id),published_at timestamptz not null
);
create table humbel_security.contact_fields(field text primary key,evidence text not null check(length(trim(evidence))>0));
create table humbel_security.contacts(employee_id uuid primary key references public.employees(id),data jsonb not null default '{}');
create table humbel_security.audit(
 id bigint generated always as identity primary key,actor uuid not null,event_at timestamptz not null default clock_timestamp(),
 action text not null,object_id uuid not null,before_state jsonb,after_state jsonb not null
);

create function humbel_security.actor_id() returns uuid language sql stable security definer set search_path='' as $$
 select p.employee_id from humbel_security.principals p join public.employees e on e.id=p.employee_id
 where p.user_id=auth.uid() and p.enabled and p.valid_until>statement_timestamp()
 and e.is_active and e.auth_user_id=p.user_id and e.role=p.role and e.department_id is not distinct from p.department_id
 and not coalesce((auth.jwt()->>'is_anonymous')::boolean,false)
$$;
create function humbel_security.actor_role() returns text language sql stable security definer set search_path='' as $$
 select p.role from humbel_security.principals p where p.employee_id=humbel_security.actor_id()
$$;
create function humbel_security.is_admin() returns boolean language sql stable security definer set search_path='' as $$
 select coalesce(humbel_security.actor_role()='admin',false)
$$;
create function humbel_security.can_plan(dept uuid) returns boolean language sql stable security definer set search_path='' as $$
 select humbel_security.actor_role() in('employee','department_admin','admin') and exists(
 select 1 from public.departments d where d.id=dept and d.active and (
 humbel_security.is_admin() or (humbel_security.actor_role()='department_admin' and d.leader_employee_id=humbel_security.actor_id())
 or exists(select 1 from humbel_security.delegations g
 join humbel_security.principals principal on principal.user_id=g.principal_user_id
 join public.employees e on e.id=principal.employee_id
 where g.user_id=auth.uid() and g.department_id=dept and not g.revoked
 and g.valid_from<=statement_timestamp() and g.valid_until>statement_timestamp()
 and principal.enabled and principal.valid_until>statement_timestamp() and principal.role='department_admin'
 and e.id=d.leader_employee_id and e.is_active and e.role=principal.role and e.auth_user_id=principal.user_id
 and e.department_id=dept and principal.department_id=dept)))
$$;
create function humbel_security.schedule_scope(dept uuid,publish boolean) returns boolean language sql stable security definer set search_path='' as $$
 select humbel_security.actor_role() in('employee','department_admin','admin') and exists(
 select 1 from humbel_security.schedule_grants g join public.departments d on d.id=g.department_id
 where g.user_id=auth.uid() and g.department_id=dept and d.active and not g.revoked
 and g.valid_from<=statement_timestamp() and g.valid_until>statement_timestamp()
 and case when publish then g.can_publish else g.can_read end)
$$;
create function humbel_security.record_scope(t text,r uuid,writing boolean) returns boolean language sql stable security definer set search_path='' as $$
 select humbel_security.actor_role() in('employee','department_admin','admin') and exists(
 select 1 from humbel_security.record_bindings b join public.departments d on d.id=b.department_id
 where b.table_name=t and b.row_id=r and d.active and (
 (t in('planner_assignments','planner_shift_cancellations') and humbel_security.can_plan(b.department_id))
 or exists(select 1 from humbel_security.record_grants g where g.user_id=auth.uid() and g.department_id=b.department_id
 and g.table_name=t and not g.revoked and g.valid_from<=statement_timestamp() and g.valid_until>statement_timestamp()
 and case when writing then g.can_write else g.can_read end)))
$$;
-- Known original helper signatures retained. No role derived from editable metadata.
create or replace function public.is_admin() returns boolean language sql stable security definer set search_path='' as $$select humbel_security.is_admin()$$;
create or replace function public.current_employee_id() returns uuid language sql stable security definer set search_path='' as $$select humbel_security.actor_id()$$;
create or replace function public.current_app_employee_id() returns uuid language sql stable security definer set search_path='' as $$select humbel_security.actor_id()$$;
create or replace function public.current_app_employee_role() returns text language sql stable security definer set search_path='' as $$select humbel_security.actor_role()$$;
-- Preserve original production-department contract; planning delegation is NOT a production grant.
create or replace function public.can_access_production_department(target_department_id uuid) returns boolean language sql stable security definer set search_path='' as $$
 select humbel_security.is_admin() or (humbel_security.actor_role()='department_admin' and exists(
 select 1 from public.departments d where d.id=target_department_id and d.active and d.leader_employee_id=humbel_security.actor_id()))
$$;

do $policies$
declare t text;p record;c record;
begin
 foreach t in array array['employees','departments','planner_absence_replacements','planner_absences','planner_assignments',
 'planner_availability','planner_saturday_requests','planner_shift_cancellations','planner_sick_leaves','planner_swaps','planner_vacations',
 'planner_special_days','planner_special_day_assignments','planner_tasks','tool_materials'] loop
  execute format('revoke all privileges on table public.%I from public,anon,authenticated',t);
  -- Table-level REVOKE alone does not remove pre-existing column privileges.
  for c in select attname from pg_attribute where attrelid=format('public.%I',t)::regclass and attnum>0 and not attisdropped loop
   execute format('revoke all privileges (%I) on table public.%I from public,anon,authenticated',c.attname,t);
  end loop;
  for p in select policyname from pg_policies where schemaname='public' and tablename=t loop
   execute format('drop policy %I on public.%I',p.policyname,t);
  end loop;
  execute format('alter table public.%I enable row level security',t);
  execute format('grant select,insert,update,delete on table public.%I to authenticated',t);
  if t like 'planner_%' then
   execute format('create policy scoped_read on public.%I for select to authenticated using(humbel_security.record_scope(%L,id,false))',t,t);
   execute format('create policy scoped_insert on public.%I for insert to authenticated with check(humbel_security.record_scope(%L,id,true))',t,t);
   execute format('create policy scoped_update on public.%I for update to authenticated using(humbel_security.record_scope(%L,id,true)) with check(humbel_security.record_scope(%L,id,true))',t,t,t);
   execute format('create policy scoped_delete on public.%I for delete to authenticated using(humbel_security.record_scope(%L,id,true))',t,t);
  end if;
 end loop;
end
$policies$;
-- Read policies do not grant raw writes. Name-based legacy writes need audited RPCs.
do $$declare t text;begin
 foreach t in array array['planner_absence_replacements','planner_absences','planner_assignments','planner_availability',
 'planner_saturday_requests','planner_shift_cancellations','planner_sick_leaves','planner_swaps','planner_vacations',
 'planner_special_days','planner_special_day_assignments','planner_tasks'] loop
  execute format('revoke insert,update,delete on public.%I from authenticated',t);
 end loop;
end$$;
create policy staff_self_or_admin on public.employees for select to authenticated using(humbel_security.is_admin() or id=humbel_security.actor_id());
create policy employee_admin_insert on public.employees for insert to authenticated with check(humbel_security.is_admin());
create policy employee_admin_update on public.employees for update to authenticated using(humbel_security.is_admin()) with check(humbel_security.is_admin());
create policy employee_admin_delete on public.employees for delete to authenticated using(humbel_security.is_admin());
create policy department_scoped_read on public.departments for select to authenticated using(humbel_security.can_plan(id) or humbel_security.schedule_scope(id,false));
create policy department_admin_write on public.departments for all to authenticated using(humbel_security.is_admin()) with check(humbel_security.is_admin());
create policy materials_staff_read on public.tool_materials for select to authenticated using(humbel_security.actor_id() is not null);
create policy materials_admin_write on public.tool_materials for all to authenticated using(humbel_security.is_admin()) with check(humbel_security.is_admin());

-- Direct legacy updates cannot rebind assignment identity, scope or publication.
create function humbel_security.guard_assignment() returns trigger language plpgsql security definer set search_path='' as $$
declare b humbel_security.record_bindings;
begin
 select * into b from humbel_security.record_bindings where table_name='planner_assignments' and row_id=new.id;
 if not found or not exists(select 1 from public.employees e where e.id=b.subject_employee_id and e.department_id=b.department_id
   and e.is_active and e.display_name=new.assigned_user)
 then raise exception 'Assignment requires verified person/department binding' using errcode='42501';end if;
 return new;
end$$;
create trigger security_assignment_guard before insert or update on public.planner_assignments for each row execute function humbel_security.guard_assignment();

create function public.security_assignment_command(p_id uuid,p_department uuid,p_subject uuid,p_shift_id text,p_date date,p_expected_version bigint)
returns jsonb language plpgsql security definer set search_path='' as $$
declare b humbel_security.record_bindings; name text; old_state jsonb; next_version bigint; result jsonb;
begin
 perform 1 from public.departments where id=p_department for share;
 perform 1 from humbel_security.principals where user_id=auth.uid() for share;
 perform 1 from public.employees where id=humbel_security.actor_id() or id=p_subject order by id for share;
 if p_expected_version is null or p_expected_version<0 or p_id is null or p_shift_id is null or length(trim(p_shift_id))=0 or p_date is null or not coalesce(humbel_security.can_plan(p_department),false)
 then raise exception 'Planning scope denied' using errcode='42501';end if;
 select display_name into name from public.employees where id=p_subject and is_active and department_id=p_department;
 if name is null or (select count(*) from public.employees where display_name=name and is_active)<>1
 then raise exception 'Unresolved or ambiguous subject' using errcode='42501';end if;
 perform pg_advisory_xact_lock(hashtextextended('humbel.assignment.'||p_id::text,0));
 select * into b from humbel_security.record_bindings where table_name='planner_assignments' and row_id=p_id for update;
 if found then
  if b.department_id<>p_department or b.row_version<>p_expected_version then raise exception 'Stale version or foreign department' using errcode='40001';end if;
  next_version=b.row_version+1;
 else
  if p_expected_version<>0 or exists(select 1 from public.planner_assignments where id=p_id) then raise exception 'Unbound legacy row' using errcode='42501';end if;
  next_version=1;
 end if;
 select to_jsonb(a) into old_state from public.planner_assignments a where id=p_id;
 insert into humbel_security.record_bindings values('planner_assignments',p_id,p_department,p_subject,next_version,'Explicit scoped assignment command')
 on conflict(table_name,row_id) do update set subject_employee_id=excluded.subject_employee_id,row_version=excluded.row_version;
 insert into public.planner_assignments(id,shift_id,shift_date,assigned_user,created_by_employee_id)
 values(p_id,p_shift_id,p_date,name,humbel_security.actor_id()) on conflict(id) do update set shift_id=excluded.shift_id,shift_date=excluded.shift_date,assigned_user=excluded.assigned_user;
 delete from humbel_security.publications where assignment_id=p_id;
 result=jsonb_build_object('id',p_id,'version',next_version);
 insert into humbel_security.audit(actor,action,object_id,before_state,after_state) values(auth.uid(),'assignment',p_id,old_state,result);
 return result;
end$$;
create function public.security_publish_assignment(p_id uuid,p_expected_version bigint) returns void language plpgsql security definer set search_path='' as $$
declare b humbel_security.record_bindings;
begin
 select * into b from humbel_security.record_bindings where table_name='planner_assignments' and row_id=p_id for update;
 if not found or p_expected_version is null or b.row_version<>p_expected_version or not coalesce(humbel_security.schedule_scope(b.department_id,true),false)
 then raise exception 'Publication scope/version denied' using errcode='42501';end if;
 insert into humbel_security.publications values(p_id,b.row_version,auth.uid(),clock_timestamp())
 on conflict(assignment_id) do update set version=excluded.version,published_by=excluded.published_by,published_at=excluded.published_at;
 insert into humbel_security.audit(actor,action,object_id,after_state) values(auth.uid(),'publish',p_id,jsonb_build_object('version',b.row_version));
end$$;
create function public.security_published_schedule() returns table(assignment_id uuid,shift_id text,shift_date date,employee_id uuid)
language sql stable security definer set search_path='' as $$
 select a.id,a.shift_id,a.shift_date,b.subject_employee_id from public.planner_assignments a
 join humbel_security.record_bindings b on b.table_name='planner_assignments' and b.row_id=a.id
 join humbel_security.publications p on p.assignment_id=a.id and p.version=b.row_version
 join public.employees e on e.id=b.subject_employee_id and e.is_active and e.department_id=b.department_id and e.display_name=a.assigned_user
 where humbel_security.schedule_scope(b.department_id,false)
$$;
-- No verified contact columns exist in the captured employees table. This proposed
-- separate store starts with an EMPTY field allowlist; the app is not switched to it.
create function public.security_own_contacts(p_values jsonb default null) returns jsonb language plpgsql security definer set search_path='' as $$
declare eid uuid;key text;old_state jsonb;result jsonb;
begin
 perform 1 from humbel_security.principals where user_id=auth.uid() for share;
 eid=humbel_security.actor_id();
 if eid is null or humbel_security.actor_role()='tool_scanner' then raise exception 'Selfservice denied' using errcode='42501';end if;
 if p_values is null then return coalesce((select data from humbel_security.contacts where employee_id=eid),'{}');end if;
 if jsonb_typeof(p_values)<>'object' then raise exception 'Contact object required' using errcode='22023';end if;
 for key in select jsonb_object_keys(p_values) loop
  if key in('role','auth_user_id','department_id','department','id','is_active','employee_id','user_id','permissions')
   or not exists(select 1 from humbel_security.contact_fields where field=key)
   or jsonb_typeof(p_values->key)<>'string' or length(p_values->>key)>1024
  then raise exception 'Contact field not approved' using errcode='42501';end if;
 end loop;
 insert into humbel_security.contacts(employee_id) values(eid) on conflict do nothing;
 select data into old_state from humbel_security.contacts where employee_id=eid for update;
 update humbel_security.contacts set data=data||p_values where employee_id=eid returning data into result;
 insert into humbel_security.audit(actor,action,object_id,before_state,after_state) values(auth.uid(),'own_contacts',eid,old_state,result);
 return result;
end$$;
-- All new internal tables are inaccessible directly, including the bootstrap trust root.
do $$declare t record;begin
 for t in select tablename from pg_tables where schemaname='humbel_security' loop
  execute format('alter table humbel_security.%I enable row level security',t.tablename);
  execute format('revoke all on table humbel_security.%I from public,anon,authenticated',t.tablename);
 end loop;
end$$;
revoke all on all sequences in schema humbel_security from public,anon,authenticated;
revoke execute on all functions in schema humbel_security from public,anon,authenticated;
grant usage on schema humbel_security to authenticated;
grant execute on function humbel_security.actor_id(),humbel_security.actor_role(),humbel_security.is_admin(),humbel_security.can_plan(uuid),humbel_security.schedule_scope(uuid,boolean),humbel_security.record_scope(text,uuid,boolean) to authenticated;
revoke execute on function public.security_assignment_command(uuid,uuid,uuid,text,date,bigint),public.security_publish_assignment(uuid,bigint),public.security_published_schedule(),public.security_own_contacts(jsonb) from public,anon,authenticated;
grant execute on function public.security_assignment_command(uuid,uuid,uuid,text,date,bigint),public.security_publish_assignment(uuid,bigint),public.security_published_schedule(),public.security_own_contacts(jsonb) to authenticated;
revoke execute on function public.is_admin(),public.current_employee_id(),public.current_app_employee_id(),public.current_app_employee_role(),public.can_access_production_department(uuid) from public,anon;
commit;
