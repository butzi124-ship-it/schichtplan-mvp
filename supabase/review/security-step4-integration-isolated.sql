-- ISOLATED REVIEW ONLY. Additive integration contract on the unchanged step-3 candidate.
begin;
set local lock_timeout='5s';
set local statement_timeout='15s';
do $$begin
 if not (pg_catalog.current_database()='humbel_isolated_test' or (pg_catalog.current_database()='postgres' and pg_catalog.version() like '%(PGlite %'))
 or pg_catalog.length(coalesce(pg_catalog.current_setting('humbel.isolated_sentinel',true),''))<32
 or pg_catalog.to_regclass('humbel_security.principals') is null
 then raise exception 'ISOLATED STEP3 CANDIDATE REQUIRED' using errcode='42501';end if;
end$$;
-- No user/employee ID parameter: identity and capabilities are always server derived.
create function public.security_actor_context() returns jsonb language plpgsql stable security definer set search_path='' as $$
declare eid uuid;role_name text;dept uuid;plan jsonb;fields jsonb;
begin
 eid=humbel_security.actor_id();role_name=humbel_security.actor_role();
 if eid is null then raise exception 'Verified active identity required' using errcode='42501';end if;
 select department_id into dept from humbel_security.principals where user_id=auth.uid();
 select coalesce(jsonb_agg(d.id order by d.id),'[]') into plan from public.departments d where humbel_security.can_plan(d.id);
 select coalesce(jsonb_agg(f.field order by f.field),'[]') into fields from humbel_security.contact_fields f
 where f.field not in('role','auth_user_id','department_id','department','id','is_active','employee_id','user_id','permissions');
 return jsonb_build_object('user_id',auth.uid(),'employee_id',eid,'role',role_name,'department_id',dept,
 'can_admin',humbel_security.is_admin(),'plan_departments',plan,'contact_fields',case when role_name='tool_scanner' then '[]'::jsonb else fields end);
end$$;
revoke execute on function public.security_actor_context() from public,anon,authenticated;
grant execute on function public.security_actor_context() to authenticated;
commit;
