-- ISOLATED SECURITY REVIEW ONLY. Not a production migration.
-- Structural-catalog mirror only: requires exact restored function bodies, ACLs, and
-- explicit user-role access matrix before production adaptation.
-- The hard guard intentionally prevents executing this review file in production.
begin;
set local lock_timeout='5s';
set local statement_timeout='15s';
do $security_review$
declare t text; pol record;
begin
 if current_database()<>'humbel_isolated_test'
    or coalesce(current_setting('humbel.isolated_sentinel',true),'')=''
 then raise exception 'ISOLATED REVIEW ONLY; real or unknown database forbidden' using errcode='42501'; end if;
 -- Existing authenticated policies on these planner tables allow ALL rows/writes:
 -- simply enabling RLS would not fix the exposure. Replace them *in the test*.
 foreach t in array array[
   'planner_absence_replacements','planner_absences','planner_assignments',
   'planner_availability','planner_saturday_requests','planner_shift_cancellations',
   'planner_sick_leaves','planner_swaps','planner_vacations'
 ] loop
   execute format('revoke all privileges on table public.%I from public,anon',t);
   for pol in select policyname from pg_catalog.pg_policies
              where schemaname='public' and tablename=t loop
     execute format('drop policy %I on public.%I',pol.policyname,t);
   end loop;
   execute format('alter table public.%I enable row level security',t);
   -- Deliberately conservative candidate: admin-only until a verified roster/user
   -- binding and privacy-preserving read RPC are specified and regression-tested.
   execute format('create policy %I on public.%I for all to authenticated using (public.is_admin()) with check (public.is_admin())',
                  t||'_admin_only_review',t);
 end loop;
 -- Existing employee-admin policies remain, but permissive 'true' write policies
 -- must not override them through OR-composition.
 if not exists(select 1 from pg_catalog.pg_policies where schemaname='public'
   and tablename='employees' and policyname='employees_update_admin')
 then raise exception 'Required admin baseline policy missing'; end if;
 drop policy if exists employees_insert on public.employees;
 drop policy if exists employees_update on public.employees;
 drop policy if exists employees_delete on public.employees;
 -- Tool-materials is non-medical reference data: authenticated active staff may
 -- read; writes remain admin-only. Anonymous API access is denied.
 revoke all privileges on table public.tool_materials from public,anon;
 alter table public.tool_materials enable row level security;
 create policy tool_materials_active_employee_read_review
   on public.tool_materials for select to authenticated
   using (auth.uid() is not null and exists(
     select 1 from public.employees e
      where e.auth_user_id=auth.uid() and e.is_active));
 create policy tool_materials_admin_write_review
   on public.tool_materials for all to authenticated
   using (public.is_admin()) with check (public.is_admin());
end
$security_review$;
commit;
