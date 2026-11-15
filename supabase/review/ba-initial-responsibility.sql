-- REVIEW ONLY. Requires phases 1–5. No production application or activation.
-- Separate disabled gate preserves the original APIs and all existing records.
begin;
set local lock_timeout='5s';
set local statement_timeout='20s';
create table public.production_initial_responsibility_settings(singleton boolean primary key default true check(singleton),enabled boolean not null default false);
insert into public.production_initial_responsibility_settings values(true,false);
alter table public.production_initial_responsibility_settings enable row level security;
revoke all on public.production_initial_responsibility_settings from public,anon,authenticated;
create function article_master_private.initial_responsibility_enabled() returns boolean language sql stable security definer set search_path='' as $$
 select article_master_private.lifecycle_enabled() and coalesce((select enabled from public.production_initial_responsibility_settings where singleton),false)
$$;
-- Keep the already checked Phase-1 readers/commands, including all identity checks.
alter function article_master_private.ba_entry_read(uuid,integer) rename to ba_entry_read_before_initial_responsibility;
alter function article_master_private.ba_entry_command(text,text,text,integer,integer) rename to ba_entry_command_before_initial_responsibility;
revoke all on function article_master_private.ba_entry_read_before_initial_responsibility(uuid,integer),article_master_private.ba_entry_command_before_initial_responsibility(text,text,text,integer,integer) from public,anon,authenticated;
create function article_master_private.ba_entry_read(oid uuid,selected_op integer default null) returns jsonb language plpgsql stable security definer set search_path='' as $$
declare result jsonb;l public.production_lifecycle_orders;
begin
 result:=article_master_private.ba_entry_read_before_initial_responsibility(oid,selected_op);
 if not article_master_private.initial_responsibility_enabled() or coalesce((result->>'legacy')::boolean,false) then return result;end if;
 select * into l from public.production_lifecycle_orders where order_id=oid;
 return result||jsonb_build_object('responsibility_status',case when l.principal_user_id is null then 'open' else 'assigned' end,'principal_user_id',l.principal_user_id,'principal_department_id',l.principal_department_id,'principal_label',(select coalesce(nullif(to_jsonb(e)->>'display_name',''),nullif(trim(coalesce(to_jsonb(e)->>'first_name','')||' '||coalesce(to_jsonb(e)->>'last_name','')),''),'Zugeordnete Abteilungsleitung') from public.employees e join public.departments d on d.leader_employee_id=e.id where d.id=l.principal_department_id and e.auth_user_id=l.principal_user_id));
end;$$;
create function article_master_private.ba_entry_command(p_action text,p_ba_number text,p_article_number text,p_quantity integer,p_op_number integer) returns jsonb language plpgsql security definer set search_path='' as $$
declare result jsonb;oid uuid;op uuid;dept uuid;first_number integer;target uuid;candidate_count integer;request uuid;stamp timestamptz;before_data jsonb;outcome text;authority_evidence jsonb;
begin
 if not article_master_private.initial_responsibility_enabled() or p_action is distinct from 'create' then
  return article_master_private.ba_entry_command_before_initial_responsibility(p_action,p_ba_number,p_article_number,p_quantity,p_op_number);
 end if;
 -- Same global serialization and ordered authority locks as lifecycle commands.
 -- Hold them across creation, selection and audit; revocation cannot race assignment.
 perform pg_advisory_xact_lock(73103,1);
 perform 1 from public.production_initial_responsibility_settings for share;
 perform 1 from public.production_lifecycle_settings for share;
 perform 1 from public.production_setup_settings for share;
 perform 1 from public.production_routing_settings for share;
 perform 1 from public.production_quantity_settings for share;
 perform 1 from public.production_readiness_settings for share;
 -- Row locks prevent revocation; SHARE also prevents phantom competing grants/profiles.
 lock table public.departments,public.employees,public.production_routing_authorities in share mode;
 perform 1 from public.departments order by id for share;
 perform 1 from public.employees order by id for share;
 perform 1 from public.production_routing_authorities order by department_id,user_id for share;
 -- Recheck after potentially waiting for gate/authority locks.
 if not article_master_private.initial_responsibility_enabled() then
  return article_master_private.ba_entry_command_before_initial_responsibility(p_action,p_ba_number,p_article_number,p_quantity,p_op_number);
 end if;
 result:=article_master_private.ba_entry_command_before_initial_responsibility(p_action,p_ba_number,p_article_number,p_quantity,p_op_number);
 -- Opening/retrying an existing BA never assigns again or overwrites a transfer.
 if (result->>'existing')::boolean then return result;end if;
 oid:=(result->>'order_id')::uuid;
 select o.id,o.op_number,article_master_private.routing_department(o.id) into op,first_number,dept from public.production_order_operations o where o.order_id=oid order by o.position limit 1;
 if op is null or dept is null or first_number<>(result->>'selected_op')::integer then raise exception 'Missing verified first OP' using errcode='22023';end if;
 stamp:=clock_timestamp();request:=gen_random_uuid();before_data:=article_master_private.lifecycle_row(oid);
 -- Count all live independent leader grants, not just the desired profile role.
 -- Conflicting grants are ambiguous even if the department pointer favours one.
 select count(*),(array_agg(a.user_id order by a.user_id))[1] into candidate_count,target
 from public.production_routing_authorities a join public.employees e on e.auth_user_id=a.user_id
 join public.departments d on d.id=a.department_id
 where a.department_id=dept and a.authority='leader' and not a.revoked and a.valid_from<=stamp and a.valid_until>stamp
 and d.active and e.is_active and e.role='department_admin';
 outcome:=case when candidate_count=0 then 'missing_valid_leader' when candidate_count<>1 then 'ambiguous_leader' else 'assigned' end;
 if candidate_count=1 and not exists(select 1 from public.departments d join public.employees e on e.id=d.leader_employee_id where d.id=dept and d.active and e.auth_user_id=target and e.is_active and e.role='department_admin' and article_master_private.routing_manager(dept,target)) then outcome:='leader_mapping_mismatch';end if;
 if outcome<>'assigned' then target:=null;end if;
 select to_jsonb(a) into authority_evidence from public.production_routing_authorities a where a.user_id=target and a.department_id=dept;
 insert into public.production_lifecycle_orders(order_id,principal_user_id,principal_department_id,row_version) values(oid,target,case when target is not null then dept end,1);
 insert into public.production_lifecycle_events(request_id,order_id,operation_id,actor_id,happened_at,action,payload,before_value,after_value)
 values(request,oid,op,auth.uid(),stamp,'initial_responsibility',jsonb_build_object('rule','first_op_department_v1','outcome',outcome,'candidate_count',candidate_count,'first_op_number',first_number,'department_id',dept,'revision_id',result->'revision_id','authority',authority_evidence),before_data,article_master_private.lifecycle_row(oid));
 if target is not null then perform article_master_private.lifecycle_notify(oid,op,request,'responsibility_assigned',target);end if;
 return article_master_private.ba_entry_read(oid,first_number)||jsonb_build_object('existing',false);
end;$$;
-- Refresh the public invoker so pooled sessions cannot retain the renamed target.
create or replace function public.ba_entry_command(p_action text,p_ba_number text default null,p_article_number text default null,p_quantity integer default null,p_op_number integer default null) returns jsonb language sql security invoker set search_path='' as $$select article_master_private.ba_entry_command(p_action,p_ba_number,p_article_number,p_quantity,p_op_number)$$;
revoke all on function article_master_private.initial_responsibility_enabled(),article_master_private.ba_entry_read(uuid,integer),article_master_private.ba_entry_command(text,text,text,integer,integer) from public,anon,authenticated;
grant execute on function article_master_private.ba_entry_command(text,text,text,integer,integer) to authenticated;
commit;
