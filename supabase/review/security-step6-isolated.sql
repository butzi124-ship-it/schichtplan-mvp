-- REVIEW/ISOLATED ONLY, additive to exact Step3 + Step4; never register as a migration.
-- Original column/constraint/owner definitions remain UNVERIFIED. No live use.
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
-- Source: addToolMaterial, renameToolMaterial, deactivateToolMaterial at PR30.
-- Keep SELECT for legacy returning rows; Step4 readMaterials uses four explicit columns.
-- No current application path physically deletes material references.
-- RLS still authorizes these writes ONLY to independently verified administrators.
revoke insert,update,delete on public.tool_materials from authenticated;
grant insert(name,sort_order,is_active) on public.tool_materials to authenticated;
grant update(name,is_active,updated_at) on public.tool_materials to authenticated;
-- No role/auth/employee/department/legacy data backfill and no new workflow permissions.
commit;
