-- ISOLATED REVIEW ONLY. No hosted use, no registered migration/deployment hook.
-- Original metadata captured 2026-11-12; no business data or identity backfill.
begin;
set local lock_timeout='5s';
set local statement_timeout='30s';
do $$begin
 if not (pg_catalog.current_database()='humbel_isolated_test' or
   (pg_catalog.current_database()='postgres' and pg_catalog.version() like '%(PGlite %'))
   or pg_catalog.length(coalesce(pg_catalog.current_setting('humbel.isolated_sentinel',true),''))<32
 then raise exception 'ISOLATED STEP6 ONLY' using errcode='42501';end if;
 if pg_catalog.to_regprocedure('public.security_actor_context()') is null then
  raise exception 'Exact reviewed Step3/4 prerequisites required' using errcode='55000';end if;
end$$;
-- Enabling a role already required by existing AL policies is not enrollment.
-- Only the observed original or known phase-2 candidate constraint is accepted.
do $role_check$
declare definition text;
begin
 select pg_catalog.pg_get_constraintdef(oid) into definition from pg_catalog.pg_constraint
 where conrelid='public.employees'::regclass and conname='employees_role_check';
 if definition = $$CHECK ((role = ANY (ARRAY['admin'::text, 'employee'::text, 'tool_scanner'::text])))$$ then
  alter table public.employees drop constraint employees_role_check;
  alter table public.employees add constraint employees_role_check check(role in('admin','employee','tool_scanner','department_admin'));
 elsif definition <> $$CHECK ((role = ANY (ARRAY['admin'::text, 'employee'::text, 'tool_scanner'::text, 'department_admin'::text])))$$ or definition is null then
  raise exception 'Unverified employee role constraint' using errcode='55000';
 end if;
end
$role_check$;
-- RLS does not protect TRUNCATE. Preserve ordinary CRUD behind existing policies.
do $acl$
declare t text;c record;creator text;
begin
 foreach t in array array['departments','employees','order_events','planner_absence_replacements','planner_absences','planner_assignments','planner_availability','planner_saturday_requests','planner_shift_cancellations','planner_sick_leaves','planner_special_day_assignments','planner_special_days','planner_swaps','planner_tasks','planner_vacations','production_checklist_templates','production_counts','production_machine50_storage','production_machines','production_order_checklist','production_order_employees','production_order_history','production_order_stations','production_orders','production_qa_causes','production_station_counts','production_station_events','shift_definitions','tool_journal','tool_materials','tools'] loop
  execute format('revoke truncate,references,trigger,maintain on public.%I from public,anon,authenticated',t);
  for c in select attname from pg_catalog.pg_attribute where attrelid=format('public.%I',t)::regclass and attnum>0 and not attisdropped loop
   execute format('revoke references (%I) on public.%I from public,anon,authenticated',c.attname,t);
  end loop;
 end loop;
 -- Both creators/public defaults are directly present in the observed catalog.
 foreach creator in array array['postgres','supabase_admin'] loop
  if not exists(select 1 from pg_catalog.pg_roles where rolname=creator) then
   raise exception 'Original default-ACL creator missing: %',creator using errcode='55000';end if;
  execute format('alter default privileges for role %I in schema public revoke truncate,references,trigger,maintain on tables from public,anon,authenticated',creator);
 end loop;
end
$acl$;
-- Trigger still runs on ordinary row updates. No direct API invocation is required.
revoke execute on function public.set_updated_at() from public,anon,authenticated;
commit;
