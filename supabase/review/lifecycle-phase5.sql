-- REVIEW ONLY: unregistered, disabled, never run on hosted/production databases.
-- Requires phases 1–4. Legacy counters, snapshots and time history remain the source.
begin;
set local lock_timeout='5s';
set local statement_timeout='20s';
create table public.production_lifecycle_settings(singleton boolean primary key default true check(singleton),enabled boolean not null default false);
insert into public.production_lifecycle_settings values(true,false);
-- No inferred initial AL. Pending responsibility must be explicitly provisioned/assigned.
create table public.production_lifecycle_orders(order_id uuid primary key references public.production_orders(id),principal_user_id uuid references auth.users(id),principal_department_id uuid references public.departments(id),row_version integer not null default 0,check((principal_user_id is null)=(principal_department_id is null)));
create table public.production_lifecycle_reasons(code text primary key,label text not null,requires_note boolean not null default false,active boolean not null default true);
create table public.production_lifecycle_events(request_id uuid primary key,order_id uuid not null references public.production_orders(id),operation_id uuid references public.production_order_operations(id),actor_id uuid not null references auth.users(id),happened_at timestamptz not null,action text not null,payload jsonb not null,before_value jsonb not null,after_value jsonb not null);
create table public.production_lifecycle_closures(id uuid primary key default gen_random_uuid(),request_id uuid not null unique references public.production_lifecycle_events(request_id) deferrable initially deferred,order_id uuid not null references public.production_orders(id),operation_id uuid references public.production_order_operations(id),closed_by uuid not null references auth.users(id),closed_at timestamptz not null,reason text,amounts jsonb not null);
create table public.production_lifecycle_rework(section_id uuid primary key references public.production_operation_machine_sections(id),request_id uuid references public.production_lifecycle_events(request_id) deferrable initially deferred,baseline_good bigint not null);
-- Each episode references the existing Phase-4 conflict. There is no second detector.
create table public.production_lifecycle_cases(episode bigint generated always as identity unique,id uuid primary key default gen_random_uuid(),conflict_id uuid not null references public.production_quantity_conflicts(id),quantity_request uuid not null references public.production_quantity_events(request_id) deferrable initially deferred,detected_at timestamptz not null default clock_timestamp(),after_completion boolean,unique(conflict_id,quantity_request));
create table public.production_lifecycle_case_entries(id bigint generated always as identity primary key,case_id uuid not null references public.production_lifecycle_cases(id),request_id uuid not null unique references public.production_lifecycle_events(request_id) deferrable initially deferred,actor_id uuid not null references auth.users(id),happened_at timestamptz not null,kind text not null check(kind in('comment','resolved','resolved_remaining')),reason text not null check(length(trim(reason)) between 1 and 4000),remaining bigint not null);
create table public.production_lifecycle_notifications(id uuid primary key default gen_random_uuid(),request_id uuid not null,recipient_user_id uuid not null references auth.users(id),order_id uuid not null references public.production_orders(id),kind text not null,created_at timestamptz not null default clock_timestamp(),unique(request_id,recipient_user_id,kind));
-- Private transaction capabilities. No user-settable GUC authorizes guarded writes.
create table public.production_lifecycle_context(transaction_id bigint primary key,actor_id uuid not null,order_id uuid not null,operation_id uuid,mode text not null);
create function article_master_private.lifecycle_enabled() returns boolean language sql stable security definer set search_path='' as $$select article_master_private.quantity_enabled() and article_master_private.setup_enabled() and coalesce((select enabled from public.production_lifecycle_settings where singleton),false)$$;
create function article_master_private.lifecycle_context(oid uuid,mode text) returns boolean language sql stable security definer set search_path='' as $$select article_master_private.lifecycle_enabled() and exists(select 1 from public.production_lifecycle_context c where c.transaction_id=txid_current() and c.actor_id=auth.uid() and c.order_id=oid and c.mode=mode)$$;
create function article_master_private.lifecycle_correct(op uuid) returns boolean language sql stable security definer set search_path='' as $$select exists(select 1 from public.production_lifecycle_context c where c.transaction_id=txid_current() and c.actor_id=auth.uid() and c.operation_id=op and c.mode='correct_count') and article_master_private.lifecycle_enabled()$$;
create function article_master_private.lifecycle_admin(oid uuid) returns boolean language sql stable security definer set search_path='' as $$
 select exists(select 1 from public.production_order_article_snapshots s join public.production_routing_authorities a on a.department_id=s.department_id or a.department_id=(select principal_department_id from public.production_lifecycle_orders where order_id=oid) where s.order_id=oid and a.user_id=auth.uid() and a.authority='admin' and article_master_private.routing_manager(a.department_id))
$$;
create function article_master_private.lifecycle_manage(oid uuid) returns boolean language sql stable security definer set search_path='' as $$
 select article_master_private.lifecycle_admin(oid) or exists(select 1 from public.production_lifecycle_orders l join public.production_routing_authorities a on a.department_id=l.principal_department_id where l.order_id=oid and a.user_id=auth.uid() and article_master_private.routing_manager(a.department_id) and ((a.authority='leader' and a.user_id=l.principal_user_id) or (a.authority='deputy' and a.principal_user_id=l.principal_user_id)))
$$;
create function article_master_private.lifecycle_visible(oid uuid) returns boolean language sql stable security definer set search_path='' as $$select article_master_private.lifecycle_enabled() and article_master_private.routing_active() and (article_master_private.lifecycle_manage(oid) or exists(select 1 from public.production_order_operations o where o.order_id=oid and article_master_private.routing_visible(o.id)))$$;
create function article_master_private.lifecycle_participant(op uuid) returns boolean language sql stable security definer set search_path='' as $$
 select article_master_private.routing_manager(article_master_private.routing_department(op)) or (exists(select 1 from public.production_operation_routing r where r.operation_id=op and article_master_private.routing_operator(op,r.machine_id)) and (exists(select 1 from public.production_quantity_attributions where operation_id=op and actor_id=auth.uid()) or exists(select 1 from public.production_operation_activity_segments a join public.production_operation_machine_sections x on x.id=a.section_id where x.operation_id=op and a.user_id=auth.uid())))
$$;
create function article_master_private.lifecycle_difference(conf uuid) returns bigint language sql stable security definer set search_path='' as $$select greatest(0,article_master_private.quantity_good(c.station_id)+s.scrap_total+s.clarify_total-article_master_private.quantity_available(c.station_id)) from public.production_quantity_conflicts c join public.production_order_stations s on s.id=c.station_id where c.id=conf$$;
create function article_master_private.lifecycle_case_open(cid uuid) returns boolean language sql stable security definer set search_path='' as $$select not exists(select 1 from public.production_lifecycle_case_entries where case_id=cid and kind in('resolved','resolved_remaining'))$$;
create function article_master_private.lifecycle_notify(oid uuid,op uuid,request uuid,kind text,target uuid default null) returns void language plpgsql security definer set search_path='' as $$begin
 insert into public.production_lifecycle_notifications(request_id,recipient_user_id,order_id,kind)
 select request,u,oid,kind from (select target u where target is not null union select l.principal_user_id from public.production_lifecycle_orders l where l.order_id=oid and l.principal_user_id is not null union select a.user_id from public.production_routing_authorities a where a.department_id=article_master_private.routing_department(op) and article_master_private.routing_manager(a.department_id,a.user_id)) recipients on conflict do nothing;
end;$$;
create function article_master_private.lifecycle_detected(conf uuid,request uuid) returns void language plpgsql security definer set search_path='' as $$
declare latest public.production_lifecycle_cases;op uuid;oid uuid;remaining bigint;prior_remaining bigint;begin
 if not article_master_private.lifecycle_enabled() then return;end if;
 select c.operation_id,o.order_id into op,oid from public.production_quantity_conflicts c join public.production_order_operations o on o.id=c.operation_id where c.id=conf;
 select * into latest from public.production_lifecycle_cases where conflict_id=conf order by episode desc limit 1;
 if latest.id is not null and article_master_private.lifecycle_case_open(latest.id) then return;end if;
 remaining:=article_master_private.lifecycle_difference(conf);
 select e.remaining into prior_remaining from public.production_lifecycle_case_entries e where e.case_id=latest.id and e.kind in('resolved','resolved_remaining');
 if latest.id is not null and remaining is not distinct from prior_remaining then return;end if;
 insert into public.production_lifecycle_cases(conflict_id,quantity_request,after_completion) select conf,request,b.status='completed' or exists(select 1 from public.production_operation_execution where operation_id=op and work_status='completed') from public.production_orders b where b.id=oid;
 perform article_master_private.lifecycle_notify(oid,op,request,'quantity_case');
end;$$;
-- Adopt detected Phase-4 cases without changing their original IDs or notifications.
-- A missing historical completion relation stays unknown, not inferred from current status.
insert into public.production_lifecycle_cases(conflict_id,quantity_request,detected_at,after_completion)
select c.id,c.first_request,c.first_detected,null::boolean from public.production_quantity_conflicts c join public.production_order_operations o on o.id=c.operation_id join public.production_orders b on b.id=o.order_id left join public.production_operation_execution e on e.operation_id=o.id;
create function article_master_private.lifecycle_adopt_cases() returns trigger language plpgsql security definer set search_path='' as $$begin
 if new.enabled and not old.enabled then
 insert into public.production_lifecycle_cases(conflict_id,quantity_request,detected_at,after_completion)
 select c.id,c.first_request,c.first_detected,null::boolean from public.production_quantity_conflicts c join public.production_order_operations o on o.id=c.operation_id join public.production_orders b on b.id=o.order_id left join public.production_operation_execution e on e.operation_id=o.id where not exists(select 1 from public.production_lifecycle_cases x where x.conflict_id=c.id);
 end if;return new;end;$$;
create trigger lifecycle_adopt_cases after update of enabled on public.production_lifecycle_settings for each row execute function article_master_private.lifecycle_adopt_cases();
create function article_master_private.lifecycle_row(oid uuid) returns jsonb language sql stable security definer set search_path='' as $$
 select jsonb_build_object('order_id',b.id,'ba_number',b.ba_number,'target',b.target_quantity,'original_target',s.quantity,'status',b.status,'row_version',coalesce(l.row_version,0),'principal_user_id',l.principal_user_id,'principal_department_id',l.principal_department_id,'principal_label',(select coalesce(nullif(to_jsonb(e)->>'display_name',''),nullif(trim(coalesce(to_jsonb(e)->>'first_name','')||' '||coalesce(to_jsonb(e)->>'last_name','')),''),'Abteilungsleitung') from public.employees e where e.auth_user_id=l.principal_user_id),'responsibility_pending',l.principal_user_id is null,'can_manage',article_master_private.lifecycle_manage(oid),'can_admin',article_master_private.lifecycle_admin(oid),
 'operations',coalesce((select jsonb_agg(jsonb_build_object('operation_id',o.id,'number',o.op_number,'work_status',coalesce(e.work_status,'planned'),'finished',article_master_private.quantity_finished(o.id),'overproduction',greatest(0,article_master_private.quantity_finished(o.id)-b.target_quantity),'target_reached',article_master_private.quantity_finished(o.id)>=b.target_quantity,'configured',article_master_private.quantity_linked(o.id),'can_close',article_master_private.lifecycle_participant(o.id),'can_manage',article_master_private.routing_manager(article_master_private.routing_department(o.id)), 'amounts',article_master_private.quantity_row(o.id),
 'active_segments',coalesce((select jsonb_agg(to_jsonb(a)) from public.production_operation_activity_segments a join public.production_operation_machine_sections x on x.id=a.section_id where x.operation_id=o.id and a.ended_at is null),'[]'::jsonb),
 'open_runs',coalesce((select jsonb_agg(to_jsonb(r)) from public.production_setup_runs r join public.production_operation_machine_sections x on x.id=r.section_id where x.operation_id=o.id and r.ended_at is null),'[]'::jsonb)) order by o.position) from public.production_order_operations o left join public.production_operation_execution e on e.operation_id=o.id where o.order_id=oid and (article_master_private.lifecycle_manage(oid) or article_master_private.routing_visible(o.id))),'[]'::jsonb),
 'cases',coalesce((select jsonb_agg(to_jsonb(c)||jsonb_build_object('operation_id',q.operation_id,'station_id',q.station_id,'open',article_master_private.lifecycle_case_open(c.id),'remaining',article_master_private.lifecycle_difference(c.conflict_id),'entries',coalesce((select jsonb_agg(to_jsonb(z) order by z.id) from public.production_lifecycle_case_entries z where z.case_id=c.id),'[]'::jsonb)) order by c.episode) from public.production_lifecycle_cases c join public.production_quantity_conflicts q on q.id=c.conflict_id join public.production_order_operations o on o.id=q.operation_id where o.order_id=oid and (article_master_private.lifecycle_manage(oid) or article_master_private.routing_visible(o.id))),'[]'::jsonb),
 'history',coalesce((select jsonb_agg(to_jsonb(ev)-'before_value'-'after_value' order by ev.happened_at,ev.request_id) from public.production_lifecycle_events ev where ev.order_id=oid and (article_master_private.lifecycle_manage(oid) or ev.operation_id is null or article_master_private.routing_visible(ev.operation_id))),'[]'::jsonb),
 'closures',coalesce((select jsonb_agg(case when article_master_private.lifecycle_manage(oid) or c.operation_id is not null then to_jsonb(c) else to_jsonb(c)-'amounts' end order by c.closed_at,c.id) from public.production_lifecycle_closures c where c.order_id=oid and (article_master_private.lifecycle_manage(oid) or c.operation_id is null or article_master_private.routing_visible(c.operation_id))),'[]'::jsonb))
 from public.production_orders b join public.production_order_article_snapshots s on s.order_id=b.id left join public.production_lifecycle_orders l on l.order_id=b.id where b.id=oid
$$;
create function article_master_private.lifecycle_baseline(op uuid) returns bigint language sql stable security definer set search_path='' as $$select coalesce((select r.baseline_good from public.production_lifecycle_rework r join public.production_operation_machine_sections x on x.id=r.section_id where x.operation_id=op and x.closed_at is null),0)$$;

-- Also retain the consumed baseline when routing creates another section mid-rework.
-- Routing already audits that section. A null lifecycle request denotes that source.
create function article_master_private.lifecycle_section_baseline() returns trigger language plpgsql security definer set search_path='' as $$declare oid uuid;begin
 select order_id into oid from public.production_order_operations where id=new.operation_id;
 if article_master_private.lifecycle_enabled() and article_master_private.quantity_linked(new.operation_id) and not article_master_private.lifecycle_context(oid,'quantity') then
 insert into public.production_lifecycle_rework(section_id,baseline_good) values(new.id,article_master_private.quantity_finished(new.operation_id));
 end if;return new;end;$$;
create trigger lifecycle_section_baseline after insert on public.production_operation_machine_sections for each row execute function article_master_private.lifecycle_section_baseline();
create function article_master_private.lifecycle_command(oid uuid,op uuid,action text,request uuid,expected integer,data jsonb) returns jsonb language plpgsql security definer set search_path='' as $$
declare b public.production_orders;l public.production_lifecycle_orders;prior public.production_lifecycle_events;before_data jsonb;after_data jsonb;body jsonb;stamp timestamptz;entry record;new_section uuid;quantity integer;target uuid;target_dept uuid;conf public.production_lifecycle_cases;conf_op uuid;difference bigint;matches integer;reopened boolean:=false;
begin
 if not article_master_private.lifecycle_visible(oid) then raise exception 'Denied' using errcode='42501';end if;
 if request is null or expected is null or data is null or jsonb_typeof(data)<>'object' or action is null or action not in('quantity','comment','resolve','complete_op','complete_ba','responsibility','correct_count','reason_config') then raise exception 'Invalid command' using errcode='22023';end if;
 if (action in('complete_op','correct_count')) is distinct from (op is not null) then raise exception 'Explicit operation scope required' using errcode='22023';end if;
 perform pg_advisory_xact_lock(73103,1);
 select * into b from public.production_orders where id=oid for update;
 -- Same order as routing; prevent concurrent revocation/role changes during authorization.
 perform 1 from public.production_lifecycle_settings for share;
 perform 1 from public.production_setup_settings for share;
 perform 1 from public.production_routing_settings for share;
 perform 1 from public.production_quantity_settings for share;
 perform 1 from public.production_readiness_settings for share;
 perform 1 from public.departments order by id for share;
 perform 1 from public.employees order by id for share;
 perform 1 from public.production_routing_authorities order by department_id,user_id for share;
 perform 1 from public.production_readiness_grants order by user_id,machine_id for share;
 if not article_master_private.lifecycle_visible(oid) then raise exception 'Authorization changed' using errcode='42501';end if;
 if op is not null and not exists(select 1 from public.production_order_operations where id=op and order_id=oid) then raise exception 'Foreign OP' using errcode='22023';end if;
 body:=jsonb_build_object('op',op,'action',action,'expected',expected,'data',data);
 select * into prior from public.production_lifecycle_events where request_id=request;
 if found then
  if prior.order_id<>oid or prior.actor_id<>auth.uid() or prior.payload<>body then raise exception 'Request reused' using errcode='22023';end if;
  -- A retry must not resurrect information whose scope was subsequently revoked.
  if coalesce((prior.after_value->>'can_manage')::boolean,false) and not article_master_private.lifecycle_manage(oid) or exists(select 1 from jsonb_array_elements(prior.after_value->'operations') item where not (article_master_private.lifecycle_manage(oid) or article_master_private.routing_visible((item->>'operation_id')::uuid))) then return article_master_private.lifecycle_row(oid);end if;
  return prior.after_value;
 end if;
 insert into public.production_lifecycle_orders(order_id) values(oid) on conflict do nothing;
 select * into l from public.production_lifecycle_orders where order_id=oid for update;
 if l.row_version<>expected then raise exception 'Stale BA version' using errcode='PT409';end if;
 before_data:=article_master_private.lifecycle_row(oid);stamp:=clock_timestamp();
 if length(coalesce(data->>'reason',''))>4000 then raise exception 'Reason too long' using errcode='22023';end if;
 insert into public.production_lifecycle_context values(txid_current(),auth.uid(),oid,op,action);
 if action='quantity' then
  if not article_master_private.lifecycle_manage(oid) then raise exception 'Main AL, active deputy or admin required' using errcode='42501';end if;
  if b.status not in('running','paused','completed') then raise exception 'Terminal cancelled BA' using errcode='PT409';end if;
  quantity:=(data->>'quantity')::integer;
  if quantity>b.target_quantity and exists(select 1 from public.production_order_operations o join public.production_operation_execution e on e.operation_id=o.id where o.order_id=oid and e.work_status='completed' and (not article_master_private.quantity_linked(o.id) or not exists(select 1 from public.production_operation_routing r where r.operation_id=o.id))) then raise exception 'Explicit historical OP adoption required before rework' using errcode='PT409';end if;
  if quantity is null or quantity<1 or quantity=b.target_quantity or (quantity<b.target_quantity and (length(trim(coalesce(data->>'reason','')))=0 or data->>'confirm_reduction' is distinct from 'true')) then raise exception 'Valid quantity; reduction warning and reason required' using errcode='22023';end if;
  update public.production_orders set ba_quantity=quantity,target_quantity=quantity where id=oid;
  if quantity>b.target_quantity then
   for entry in select e.*,r.executing_department_id from public.production_operation_execution e join public.production_order_operations o on o.id=e.operation_id join public.production_operation_routing r on r.operation_id=o.id where o.order_id=oid and e.work_status='completed' and article_master_private.quantity_finished(o.id)<quantity order by o.position for update of e loop
    reopened:=true;
    -- A historical closed section is never reopened. Preserve constraints/routing.
    if exists(select 1 from public.production_operation_machine_sections where operation_id=entry.operation_id and closed_at is null) then raise exception 'Unexpected open historical section' using errcode='PT409';end if;
    update public.production_operation_execution set work_status='planned',started_here=false,received_quantity=0,row_version=row_version+1,updated_at=stamp where operation_id=entry.operation_id;
    insert into public.production_operation_machine_sections(operation_id,machine_id,department_id,opened_at,opened_by,opening_execution) values(entry.operation_id,entry.machine_id,entry.executing_department_id,stamp,auth.uid(),to_jsonb(entry)) returning id into new_section;
    insert into public.production_lifecycle_rework values(new_section,request,article_master_private.quantity_finished(entry.operation_id));
   end loop;
   if reopened and b.status='completed' then update public.production_orders set status='running',completed_at=null where id=oid;end if;
  end if;
  perform article_master_private.lifecycle_notify(oid,null,request,'quantity_changed');
 elsif action='reason_config' then
  if not article_master_private.lifecycle_admin(oid) then raise exception 'Verified administrator required' using errcode='42501';end if;
  if coalesce(data->>'code','') !~ '^[a-zA-Z0-9_-]{1,80}$' or length(trim(coalesce(data->>'label',''))) not between 1 and 200 then raise exception 'Valid configured reason required' using errcode='22023';end if;
  perform 1 from public.production_lifecycle_reasons where code=data->>'code' for update;
  if (select to_jsonb(r) from public.production_lifecycle_reasons r where code=data->>'code') is distinct from nullif(data->'expected_reason','null'::jsonb) then raise exception 'Stale reason catalog' using errcode='PT409';end if;
  before_data:=before_data||jsonb_build_object('reason_config',(select to_jsonb(r) from public.production_lifecycle_reasons r where code=data->>'code'));
  insert into public.production_lifecycle_reasons values(data->>'code',trim(data->>'label'),coalesce((data->>'requires_note')::boolean,false),coalesce((data->>'active')::boolean,true)) on conflict(code) do update set label=excluded.label,requires_note=excluded.requires_note,active=excluded.active;
 elsif action='responsibility' then
  if not (article_master_private.lifecycle_manage(oid) and l.principal_user_id is not null or article_master_private.lifecycle_admin(oid)) then raise exception 'Main responsibility required' using errcode='42501';end if;
  target:=(data->>'user_id')::uuid;
  select (array_agg(a.department_id))[1],count(*) into target_dept,matches from public.production_routing_authorities a join public.departments d on d.id=a.department_id join public.employees e on e.id=d.leader_employee_id and e.auth_user_id=a.user_id where a.user_id=target and a.authority='leader' and article_master_private.routing_manager(a.department_id,target) and (nullif(lifecycle_command.data->>'department_id','') is null or a.department_id=(lifecycle_command.data->>'department_id')::uuid);
  if matches<>1 or target_dept is null or target is not distinct from l.principal_user_id or not exists(select 1 from public.production_lifecycle_reasons r where r.code=data->>'reason_code' and r.active and (not r.requires_note or length(trim(coalesce(data->>'reason','')))>0)) then raise exception 'Verified new AL and configured reason required' using errcode='22023';end if;
  before_data:=before_data||jsonb_build_object('reason_choice',(select to_jsonb(r) from public.production_lifecycle_reasons r where r.code=data->>'reason_code'));
  update public.production_lifecycle_orders set principal_user_id=target,principal_department_id=target_dept where order_id=oid;
  perform article_master_private.lifecycle_notify(oid,null,request,'responsibility_transferred',target);
 elsif action in('comment','resolve') then
  select c.* into conf from public.production_lifecycle_cases c join public.production_quantity_conflicts q on q.id=c.conflict_id join public.production_order_operations o on o.id=q.operation_id where c.id=(lifecycle_command.data->>'case_id')::uuid and o.order_id=oid;
  select operation_id into conf_op from public.production_quantity_conflicts where id=conf.conflict_id;
  if conf.id is null or not (article_master_private.routing_visible(conf_op) or article_master_private.lifecycle_manage(oid)) or length(trim(coalesce(data->>'reason','')))=0 then raise exception 'Accessible case and explanation required' using errcode='22023';end if;
  difference:=article_master_private.lifecycle_difference(conf.conflict_id);
  if action='resolve' and (not article_master_private.routing_manager(article_master_private.routing_department(conf_op))) then raise exception 'Responsible AL required' using errcode='42501';end if;
  if action='resolve' and (not article_master_private.lifecycle_case_open(conf.id) or (data->>'status' not in('resolved','resolved_remaining')) or data->>'status' is null or (data->>'status'='resolved' and difference<>0) or (data->>'status'='resolved_remaining' and difference=0)) then raise exception 'Use explicit resolution matching current deviation' using errcode='22023';end if;
  insert into public.production_lifecycle_case_entries(case_id,request_id,actor_id,happened_at,kind,reason,remaining) values(conf.id,request,auth.uid(),stamp,case when action='comment' then 'comment' else data->>'status' end,trim(data->>'reason'),difference);
  perform article_master_private.lifecycle_notify(oid,conf_op,request,action);
 elsif action='correct_count' then
  if op is null or length(trim(coalesce(data->>'reason','')))=0 then raise exception 'Correction reason required' using errcode='22023';end if;
  perform article_master_private.quantity_command(op,'count',request,(data->>'station_id')::uuid,(data->>'person_id')::uuid,data->>'kind',(data->>'delta')::integer,(data->>'section_id')::uuid,(data->>'cause_id')::uuid,data->>'reason',null,null);
 elsif action='complete_op' then
  if op is null or not article_master_private.quantity_linked(op) or not article_master_private.lifecycle_participant(op) then raise exception 'Authorized participant required' using errcode='42501';end if;
  if b.status not in('running','paused') or exists(select 1 from public.production_operation_execution where operation_id=op and work_status in('completed','cancelled')) then raise exception 'Already terminal' using errcode='PT409';end if;
  if exists(select 1 from public.production_lifecycle_cases c join public.production_quantity_conflicts q on q.id=c.conflict_id where q.operation_id=op and article_master_private.lifecycle_case_open(c.id)) then raise exception 'Open quantity case' using errcode='PT409';end if;
  if article_master_private.quantity_finished(op)<b.target_quantity and length(trim(coalesce(data->>'reason','')))=0 then raise exception 'Under-target reason required' using errcode='22023';end if;
  if (exists(select 1 from public.production_operation_activity_segments a join public.production_operation_machine_sections x on x.id=a.section_id where x.operation_id=op and a.ended_at is null) or exists(select 1 from public.production_setup_runs r join public.production_operation_machine_sections x on x.id=r.section_id where x.operation_id=op and r.ended_at is null)) and data->>'confirm_stop' is distinct from 'true' then raise exception 'Confirm stop of active/paused time bookings' using errcode='22023';end if;
  update public.production_operation_activity_segments a set ended_at=stamp,end_reason='op_complete' from public.production_operation_machine_sections x where x.id=a.section_id and x.operation_id=op and a.ended_at is null;
  update public.production_setup_runs r set ended_at=stamp,ended_by=auth.uid(),end_reason='op_complete' from public.production_operation_machine_sections x where x.id=r.section_id and x.operation_id=op and r.ended_at is null;
  update public.production_operation_execution set work_status='completed',good_quantity=article_master_private.quantity_finished(op),row_version=row_version+1,updated_at=stamp where operation_id=op;
  update public.production_operation_machine_sections set closed_at=stamp,closed_by=auth.uid(),close_reason='OP-Abschluss',closing_execution=(select to_jsonb(e) from public.production_operation_execution e where operation_id=op) where operation_id=op and closed_at is null;
  insert into public.production_lifecycle_closures(request_id,order_id,operation_id,closed_by,closed_at,reason,amounts) values(request,oid,op,auth.uid(),stamp,data->>'reason',jsonb_build_object('target',b.target_quantity,'operation',(select item from jsonb_array_elements(before_data->'operations') item where item->>'operation_id'=op::text),'cases',coalesce((select jsonb_agg(item) from jsonb_array_elements(before_data->'cases') item where item->>'operation_id'=op::text),'[]'::jsonb)));
  if article_master_private.quantity_finished(op)<b.target_quantity then perform article_master_private.lifecycle_notify(oid,op,request,'under_target_completion');end if;
 else
  if not article_master_private.lifecycle_manage(oid) then raise exception 'Main AL, active deputy or admin only' using errcode='42501';end if;
  if b.status not in('running','paused') or not exists(select 1 from public.production_order_operations where order_id=oid) or exists(select 1 from public.production_order_operations o left join public.production_operation_execution e on e.operation_id=o.id where o.order_id=oid and (e.work_status is distinct from 'completed' or not article_master_private.quantity_linked(o.id))) or exists(select 1 from public.production_lifecycle_cases c join public.production_quantity_conflicts q on q.id=c.conflict_id join public.production_order_operations o on o.id=q.operation_id where o.order_id=oid and article_master_private.lifecycle_case_open(c.id)) then raise exception 'All OPs and cases must be explicitly complete' using errcode='PT409';end if;
  update public.production_orders set status='completed',completed_at=stamp where id=oid;
  insert into public.production_lifecycle_closures(request_id,order_id,closed_by,closed_at,reason,amounts) values(request,oid,auth.uid(),stamp,data->>'reason',before_data-'closures'-'history');
 end if;
 update public.production_lifecycle_orders set row_version=row_version+1 where order_id=oid;
 after_data:=article_master_private.lifecycle_row(oid);
 if action='reason_config' then after_data:=after_data||jsonb_build_object('reason_config',(select to_jsonb(r) from public.production_lifecycle_reasons r where code=data->>'code'));end if;
 insert into public.production_lifecycle_events values(request,oid,case when action in('comment','resolve') then conf_op else op end,auth.uid(),stamp,action,body,before_data,after_data);
 delete from public.production_lifecycle_context where transaction_id=txid_current();
 return after_data;
end;$$;

-- Exact, gated adapters. The original Phase-4 API remains unchanged when disabled.
do $adapters$
declare d text;needle text;replacement text;
begin
 d:=pg_get_functiondef('article_master_private.protect_snapshot_order()'::regprocedure);
 needle:='new.id,new.ba_number,new.article_number,new.ba_quantity,new.target_quantity,new.department_id';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected snapshot guard';end if;
 execute replace(d,needle,'new.id,new.ba_number,new.article_number,case when article_master_private.lifecycle_context(old.id,''quantity'') then old.ba_quantity else new.ba_quantity end,case when article_master_private.lifecycle_context(old.id,''quantity'') then old.target_quantity else new.target_quantity end,new.department_id');
 d:=pg_get_functiondef('article_master_private.quantity_ba_guard()'::regprocedure);
 needle:='if article_master_private.quantity_tracked(old.id)';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected BA guard';end if;
 execute replace(d,needle,'if not (article_master_private.lifecycle_context(old.id,''quantity'') or article_master_private.lifecycle_context(old.id,''complete_ba'')) and article_master_private.quantity_tracked(old.id)');
 d:=pg_get_functiondef('article_master_private.quantity_command(uuid,text,uuid,uuid,uuid,text,integer,uuid,uuid,text,uuid[],integer)'::regprocedure);
 needle:='if b.status not in(''running'',''paused'') or exists';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected counter terminal guard';end if;
 d:=replace(d,needle,'if not article_master_private.lifecycle_correct(op) and (b.status not in(''running'',''paused'') or exists');
 d:=replace(d,'then raise exception ''Terminal operation''',') then raise exception ''Terminal operation''');
 needle:='where operation_id=op and closed_at is null for update;';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected counter section selection';end if;
 d:=replace(d,needle,'where operation_id=op and (closed_at is null or (id=section and article_master_private.lifecycle_correct(op))) order by opened_at desc limit 1 for update;');
 needle:='delta not in(-1,1)';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected counter delta';end if;
 d:=replace(d,needle,'(case when article_master_private.lifecycle_correct(op) then delta=0 or abs(delta::bigint)>1000000 else delta not in(-1,1) end)');
 needle:='case when delta<0 then ''_correction'' else '''' end';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected event kind';end if;
 d:=replace(d,needle,'case when delta<0 or article_master_private.lifecycle_correct(op) then ''_correction'' else '''' end');
 needle:='if row.counted>row.available then';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected Phase 4 detector';end if;
 d:=replace(d,needle,'if row.counted>row.available and not (article_master_private.lifecycle_enabled() and exists(select 1 from public.production_quantity_stations qs join public.production_order_operations qo on qo.id=qs.operation_id where qs.station_id=row.station_id and qs.position=1 and qo.predecessor is null)) then');
 needle:='  end if;'||chr(10)||' end loop;'||chr(10)||' return result;';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected Phase 4 detection tail';end if;
 d:=replace(d,needle,'   perform article_master_private.lifecycle_detected((select id from public.production_quantity_conflicts where station_id=row.station_id),request);'||chr(10)||needle);
 execute d;
 d:=pg_get_functiondef('article_master_private.readiness_command_before_setup(uuid,text,uuid,integer,integer,uuid,timestamptz,text)'::regprocedure);
 needle:='p_quantity>article_master_private.quantity_op_available(o.id)';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected physical transfer bound';end if;
 execute replace(d,needle,'p_quantity+(case when article_master_private.lifecycle_enabled() then article_master_private.lifecycle_baseline(o.id) else 0 end)>article_master_private.quantity_op_available(o.id)');
 d:=pg_get_functiondef('article_master_private.readiness_view()'::regprocedure);
 needle:='select s.ba_number,s.quantity,s.due_date';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected readiness target display';end if;
 d:=replace(d,needle,'select s.ba_number,case when article_master_private.lifecycle_enabled() and article_master_private.quantity_linked(o.id) then b.target_quantity else s.quantity end quantity,s.due_date');
 needle:='article_master_private.quantity_op_available(o.id) quantity_available';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected readiness availability display';end if;
 execute replace(d,needle,'case when article_master_private.lifecycle_enabled() then greatest(0,article_master_private.quantity_op_available(o.id)-article_master_private.lifecycle_baseline(o.id)) else article_master_private.quantity_op_available(o.id) end quantity_available,case when article_master_private.lifecycle_enabled() then article_master_private.lifecycle_baseline(o.id) else 0 end rework_baseline');
 d:=pg_get_functiondef('article_master_private.readiness_material(public.production_operation_execution)'::regprocedure);
 needle:='e.received_quantity>=s.quantity';
 if (length(d)-length(replace(d,needle,'')))/length(needle)<>1 then raise exception 'Unexpected readiness material prerequisite';end if;
 execute replace(d,needle,'(case when article_master_private.lifecycle_enabled() and article_master_private.quantity_linked(e.operation_id) then e.received_quantity>0 else e.received_quantity>=s.quantity end)');
end;$adapters$;
-- Ordinary writes can never reopen completed work or bypass private command contexts.
create function article_master_private.lifecycle_raw_guard() returns trigger language plpgsql set search_path='' as $$declare owner_name name;begin
 select pg_get_userbyid(relowner) into owner_name from pg_class where oid='public.production_lifecycle_orders'::regclass;
 if article_master_private.lifecycle_enabled() and article_master_private.quantity_tracked(old.id) and old.status='completed' and new.status<>'completed' and not article_master_private.lifecycle_context(old.id,'quantity') then raise exception 'Completed BA requires explicit rework release' using errcode='42501';end if;
 if article_master_private.lifecycle_enabled() and article_master_private.quantity_tracked(old.id) and current_user<>owner_name and (new.status is distinct from old.status or (new.ba_quantity,new.target_quantity) is distinct from (old.ba_quantity,old.target_quantity)) then raise exception 'Use checked lifecycle command' using errcode='42501';end if;return new;
end;$$;
create trigger lifecycle_raw_guard before update on public.production_orders for each row execute function article_master_private.lifecycle_raw_guard();
create function article_master_private.lifecycle_activity_guard() returns trigger language plpgsql security definer set search_path='' as $$begin
 if article_master_private.lifecycle_enabled() and exists(select 1 from public.production_operation_machine_sections x join public.production_order_operations o on o.id=x.operation_id join public.production_orders b on b.id=o.order_id left join public.production_operation_execution e on e.operation_id=o.id where x.id=new.section_id and (b.status='completed' or e.work_status='completed')) then raise exception 'Completed OP: corrections or released rework only' using errcode='42501';end if;return new;
end;$$;
create trigger lifecycle_activity_guard before insert on public.production_operation_activity_segments for each row execute function article_master_private.lifecycle_activity_guard();
create function public.ba_lifecycle_command(p_order_id uuid,p_action text,p_request_id uuid,p_expected_version integer,p_operation_id uuid default null,p_data jsonb default '{}'::jsonb) returns jsonb language sql security invoker set search_path='' as $$select article_master_private.lifecycle_command(p_order_id,p_operation_id,p_action,p_request_id,p_expected_version,p_data)$$;
create function article_master_private.lifecycle_view() returns jsonb language sql stable security definer set search_path='' as $$
 select case when article_master_private.lifecycle_enabled() and article_master_private.routing_active() then jsonb_build_object('rows',coalesce((select jsonb_agg(article_master_private.lifecycle_row(s.order_id) order by s.ba_number) from public.production_order_article_snapshots s where article_master_private.lifecycle_visible(s.order_id)),'[]'::jsonb),
 'reasons',case when article_master_private.lifecycle_enabled() then coalesce((select jsonb_agg(to_jsonb(r) order by label) from public.production_lifecycle_reasons r where active or exists(select 1 from public.production_routing_authorities a where a.user_id=auth.uid() and a.authority='admin' and article_master_private.routing_manager(a.department_id))),'[]'::jsonb) else '[]'::jsonb end,
 'leaders',case when article_master_private.lifecycle_enabled() and exists(select 1 from public.production_order_article_snapshots s where article_master_private.lifecycle_manage(s.order_id)) then coalesce((select jsonb_agg(jsonb_build_object('user_id',a.user_id,'department_id',a.department_id,'label',d.name||' · '||coalesce(nullif(to_jsonb(e)->>'display_name',''),nullif(trim(coalesce(to_jsonb(e)->>'first_name','')||' '||coalesce(to_jsonb(e)->>'last_name','')),''),'Abteilungsleitung'))) from public.production_routing_authorities a join public.departments d on d.id=a.department_id join public.employees e on e.id=d.leader_employee_id and e.auth_user_id=a.user_id where a.authority='leader' and article_master_private.routing_manager(a.department_id,a.user_id) and article_master_private.routing_active()),'[]'::jsonb) else '[]'::jsonb end,
 'notifications',coalesce((select jsonb_agg(to_jsonb(n) order by n.created_at desc) from public.production_lifecycle_notifications n where n.recipient_user_id=auth.uid() and article_master_private.lifecycle_visible(n.order_id)),'[]'::jsonb)) else jsonb_build_object('rows','[]'::jsonb,'reasons','[]'::jsonb,'leaders','[]'::jsonb,'notifications','[]'::jsonb) end
$$;
create function public.ba_lifecycle_view() returns jsonb language sql security invoker set search_path='' as $$select article_master_private.lifecycle_view()$$;
do $permissions$declare t text;f record;begin
 foreach t in array array['production_lifecycle_settings','production_lifecycle_orders','production_lifecycle_reasons','production_lifecycle_events','production_lifecycle_closures','production_lifecycle_rework','production_lifecycle_cases','production_lifecycle_case_entries','production_lifecycle_notifications','production_lifecycle_context'] loop
 execute format('alter table public.%I enable row level security',t);execute format('revoke all on public.%I from public,anon,authenticated',t);
 if t not in('production_lifecycle_settings','production_lifecycle_orders','production_lifecycle_reasons','production_lifecycle_context') then execute format('create trigger lifecycle_immutable before update or delete on public.%I for each row execute function article_master_private.immutable_snapshot()',t);end if;
 end loop;
 for f in select p.oid::regprocedure signature from pg_proc p join pg_namespace n on n.oid=p.pronamespace where n.nspname='article_master_private' and p.proname like 'lifecycle_%' loop execute format('revoke all on function %s from public,anon,authenticated',f.signature);end loop;
end;$permissions$;
revoke all on function public.ba_lifecycle_command(uuid,text,uuid,integer,uuid,jsonb),public.ba_lifecycle_view() from public,anon,authenticated;
grant execute on function article_master_private.lifecycle_command(uuid,uuid,text,uuid,integer,jsonb),article_master_private.lifecycle_view(),article_master_private.lifecycle_enabled(),article_master_private.lifecycle_context(uuid,text),public.ba_lifecycle_command(uuid,text,uuid,integer,uuid,jsonb),public.ba_lifecycle_view() to authenticated;
commit;
