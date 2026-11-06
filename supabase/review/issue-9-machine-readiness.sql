-- REVIEW ONLY; not registered or applied. Requires issue #7 and #8 migrations in order.
begin;
create table public.production_readiness_settings(singleton boolean primary key default true check(singleton),enabled boolean not null default false);
insert into public.production_readiness_settings values(true,false);
create table public.production_readiness_grants(
 user_id uuid not null references auth.users(id),machine_id uuid not null references public.production_machines(id),
 can_confirm boolean not null default false,can_work boolean not null default false,
 expires_at timestamptz not null,revoked boolean not null default false,issued_by uuid not null references auth.users(id),primary key(user_id,machine_id)
);
create table public.production_operation_execution(
 operation_id uuid primary key references public.production_order_operations(id) on delete restrict,
 machine_id uuid references public.production_machines(id) on delete restrict,planned_start timestamptz,
 work_status text not null default 'planned' check(work_status in('planned','ready','running','paused','completed','cancelled')),
 material_status text not null default 'planned' check(material_status in('planned','expected','ready','started')),
 received_quantity integer not null default 0 check(received_quantity>=0),good_quantity integer not null default 0 check(good_quantity>=0),
 started_at timestamptz,started_here boolean not null default false,row_version integer not null default 1 check(row_version>0),updated_at timestamptz not null default now()
);
-- No forecasting algorithm: only independently verified existing external data, provisioned server-side.
create table public.production_operation_arrivals(
 operation_id uuid primary key references public.production_order_operations(id),machine_id uuid not null references public.production_machines(id),
 expected_on date not null,expected_quantity integer not null check(expected_quantity>0),
 source_record text not null check(length(trim(source_record)) between 1 and 1000),evidence text not null check(length(trim(evidence)) between 1 and 4000),
 verified_by uuid not null references auth.users(id),verified_at timestamptz not null,valid_until timestamptz not null check(valid_until>verified_at)
);
create table public.production_readiness_events(
 request_id uuid primary key,operation_id uuid not null references public.production_order_operations(id),
 ba_number text not null,op_number integer not null,machine_id uuid references public.production_machines(id),
 actor_id uuid not null references auth.users(id),happened_at timestamptz not null default now(),action text not null,
 good_quantity integer,payload jsonb not null,before_state jsonb,after_state jsonb not null
);
create index readiness_events_operation_time on public.production_readiness_events(operation_id,happened_at);
create function article_master_private.readiness_enabled() returns boolean language sql stable security definer set search_path='' as $$select coalesce((select enabled from public.production_readiness_settings where singleton),false)$$;
create function article_master_private.readiness_allowed(op uuid,permission text) returns boolean language sql stable security definer set search_path='' as $$
 select auth.uid() is not null and article_master_private.readiness_enabled() and exists(
 select 1 from public.production_order_operations o join public.production_order_article_snapshots s on s.order_id=o.order_id
 join public.departments d on d.id=s.department_id left join public.production_operation_execution e on e.operation_id=o.id
 where o.id=op and d.active and (
 exists(select 1 from public.article_master_scopes a where a.user_id=auth.uid() and a.department_id=s.department_id)
 or (permission<>'plan' and exists(select 1 from public.production_machines rm where rm.id=coalesce(e.machine_id,(o.data->>'machine_id')::uuid) and rm.department_id=s.department_id) and exists(select 1 from public.production_readiness_grants g where g.user_id=auth.uid()
 and g.machine_id=coalesce(e.machine_id,(o.data->>'machine_id')::uuid) and not g.revoked and g.expires_at>now()
 and case permission when 'confirm' then g.can_confirm when 'work' then g.can_work when 'view' then g.can_confirm or g.can_work else false end))))
$$;
create function article_master_private.readiness_material(e public.production_operation_execution) returns text language sql stable security definer set search_path='' as $$
 select case when e.started_here then 'started' when e.received_quantity>=s.quantity then 'ready'
 when exists(select 1 from public.production_operation_arrivals f where f.operation_id=e.operation_id and f.machine_id=e.machine_id
 and f.expected_quantity>=s.quantity and f.expected_on=(now() at time zone 'Europe/Berlin')::date and f.verified_at<=now() and f.valid_until>now()) then 'expected' else 'planned' end
 from public.production_order_operations o join public.production_order_article_snapshots s on s.order_id=o.order_id where o.id=e.operation_id
$$;
create function article_master_private.readiness_event(op uuid,req uuid,kind text,body jsonb,previous jsonb,current_state jsonb,actor uuid default auth.uid()) returns void language plpgsql security definer set search_path='' as $$
begin
 insert into public.production_readiness_events(request_id,operation_id,ba_number,op_number,machine_id,actor_id,action,good_quantity,payload,before_state,after_state)
 select req,o.id,s.ba_number,o.op_number,(current_state->>'machine_id')::uuid,actor,kind,case when kind in ('handover','complete') then (body->>'quantity')::integer else (current_state->>'good_quantity')::integer end,body,previous,current_state
 from public.production_order_operations o join public.production_order_article_snapshots s on s.order_id=o.order_id where o.id=op;
end; $$;
create function article_master_private.readiness_initialize(p_order_id uuid) returns void language plpgsql security definer set search_path='' as $$
declare o public.production_order_operations;e public.production_operation_execution;
begin
 if not article_master_private.readiness_enabled() or auth.uid() is null or not exists(select 1 from public.production_order_operations init_op where init_op.order_id=p_order_id and article_master_private.readiness_allowed(init_op.id,'plan')) then raise exception 'Denied' using errcode='42501'; end if;
 perform 1 from public.production_orders where id=p_order_id for update;
 for o in select * from public.production_order_operations where order_id=p_order_id order by position loop
  insert into public.production_operation_execution(operation_id,machine_id) values(o.id,(o.data->>'machine_id')::uuid) on conflict do nothing returning * into e;
  if found then perform article_master_private.readiness_event(o.id,gen_random_uuid(),'plan','{}',null,to_jsonb(e)); end if;
 end loop;
end; $$;
create function article_master_private.readiness_new_operation() returns trigger language plpgsql security definer set search_path='' as $$
declare e public.production_operation_execution;actor uuid;
begin
 if not article_master_private.readiness_enabled() then return new; end if;
 select created_by into actor from public.production_order_article_snapshots where order_id=new.order_id;
 insert into public.production_operation_execution(operation_id,machine_id) values(new.id,(new.data->>'machine_id')::uuid) returning * into e;
 perform article_master_private.readiness_event(new.id,gen_random_uuid(),'plan','{}',null,to_jsonb(e),actor);return new;
end; $$;
create trigger readiness_new_operation after insert on public.production_order_operations for each row execute function article_master_private.readiness_new_operation();
create function article_master_private.readiness_command(p_operation_id uuid,p_action text,p_request_id uuid,p_expected_version integer,p_quantity integer,p_machine_id uuid,p_planned_start timestamptz,p_reason text)
returns jsonb language plpgsql security definer set search_path='' as $$
declare e public.production_operation_execution;old_state jsonb;o public.production_order_operations;s public.production_order_article_snapshots;b public.production_orders;prior public.production_readiness_events;body jsonb;pred public.production_operation_execution;
begin
 if p_action is null or p_action not in('handover','start','pause','resume','complete','assign','cancel') or p_request_id is null then raise exception 'Invalid command' using errcode='22023'; end if;
 if not article_master_private.readiness_allowed(p_operation_id,case when p_action in('assign','cancel') then 'plan' when p_action='handover' then 'confirm' else 'work' end) then raise exception 'Denied' using errcode='42501'; end if;
 select * into o from public.production_order_operations where id=p_operation_id;
 select * into b from public.production_orders where id=o.order_id for update;
 select * into s from public.production_order_article_snapshots where order_id=o.order_id;
 select * into e from public.production_operation_execution where operation_id=o.id for update;
 if e.operation_id is null then raise exception 'Explicit initialization required' using errcode='22023'; end if;
 body:=jsonb_build_object('action',p_action,'quantity',p_quantity,'machine_id',p_machine_id,'planned_start',p_planned_start,'reason',p_reason);
 select * into prior from public.production_readiness_events where request_id=p_request_id;
 if found then
  if prior.operation_id<>o.id or prior.actor_id<>auth.uid() or prior.payload<>body then raise exception 'Request identity reused' using errcode='22023'; end if;
  return to_jsonb(e);
 end if;
 if e.row_version is distinct from p_expected_version then raise exception 'Stale execution' using errcode='40001'; end if;
 if b.status not in('running','paused') or e.work_status in('completed','cancelled') then raise exception 'Terminal order or OP' using errcode='22023'; end if;
 if p_reason is null or length(trim(p_reason)) not between 1 and 4000 then raise exception 'Evidence/reason required' using errcode='22023'; end if;
 if b.status<>'running' and p_action in('handover','start','resume','complete') then raise exception 'Order is not running' using errcode='22023'; end if;
 old_state:=to_jsonb(e);
 if p_action='assign' then
  if e.work_status='running' then raise exception 'Pause before moving machine' using errcode='22023'; end if;
  if not exists(select 1 from public.production_machines where id=p_machine_id and department_id=s.department_id and active) then raise exception 'Invalid target machine' using errcode='42501'; end if;
  if p_machine_id is distinct from e.machine_id then e.received_quantity:=0;e.started_here:=false;e.work_status:=case when e.started_at is null then 'planned' else 'paused' end;end if;
  e.machine_id:=p_machine_id;e.planned_start:=p_planned_start;
 elsif p_action='cancel' then e.work_status:='cancelled';
 else
  if p_action<>'pause' and not exists(select 1 from public.production_machines where id=e.machine_id and department_id=s.department_id and active) then raise exception 'Machine unavailable' using errcode='22023'; end if;
  if p_action='handover' then
   if e.work_status not in('planned','paused') or e.received_quantity<>0 or p_quantity is distinct from s.quantity then raise exception 'Full explicit handover required; partial release disabled' using errcode='22023'; end if;
   if o.predecessor is not null then
    select x.* into pred from public.production_operation_execution x join public.production_order_operations op on op.id=x.operation_id where op.order_id=o.order_id and op.op_number=o.predecessor;
    if pred.operation_id is null or pred.work_status<>'completed' or pred.good_quantity<s.quantity then raise exception 'Predecessor not fully released' using errcode='22023'; end if;
   end if;
   e.received_quantity:=p_quantity;if e.started_at is null then e.work_status:='ready';end if;
  elsif p_action in('start','resume') then
   if e.received_quantity<>s.quantity or (p_action='start' and e.work_status<>'ready') or (p_action='resume' and e.work_status<>'paused') then raise exception 'Physical handover required' using errcode='22023'; end if;
   e.work_status:='running';e.started_at:=coalesce(e.started_at,now());e.started_here:=true;
  elsif p_action='pause' then
   if e.work_status<>'running' then raise exception 'Running OP required' using errcode='22023'; end if;e.work_status:='paused';
  elsif p_action='complete' then
   if e.work_status<>'running' or p_quantity is distinct from s.quantity then raise exception 'Full good quantity required; partial release disabled' using errcode='22023'; end if;
   e.good_quantity:=p_quantity;e.work_status:='completed';
  end if;
 end if;
 e.material_status:=article_master_private.readiness_material(e);e.row_version:=e.row_version+1;e.updated_at:=now();
 update public.production_operation_execution set machine_id=e.machine_id,planned_start=e.planned_start,work_status=e.work_status,material_status=e.material_status,received_quantity=e.received_quantity,good_quantity=e.good_quantity,started_at=e.started_at,started_here=e.started_here,row_version=e.row_version,updated_at=e.updated_at where operation_id=e.operation_id;
 perform article_master_private.readiness_event(o.id,p_request_id,p_action,body,old_state,to_jsonb(e));return to_jsonb(e);
end; $$;
create function article_master_private.readiness_view() returns jsonb language plpgsql security definer set search_path='' as $$
declare e public.production_operation_execution;old_state jsonb;material text;oid uuid;rows jsonb;visible boolean;
begin
 if not article_master_private.readiness_enabled() or auth.uid() is null then return jsonb_build_object('can_view',false,'rows','[]'::jsonb);end if;
 visible:=exists(select 1 from public.article_master_scopes a join public.departments d on d.id=a.department_id where a.user_id=auth.uid() and d.active)
 or exists(select 1 from public.production_readiness_grants g join public.production_machines gm on gm.id=g.machine_id join public.departments gd on gd.id=gm.department_id where gd.active and g.user_id=auth.uid() and not g.revoked and g.expires_at>now() and(g.can_confirm or g.can_work));
 -- Persist only explainable forecast display changes once; never release quantities through a read.
 for oid in select distinct o.order_id from public.production_order_operations o where article_master_private.readiness_allowed(o.id,'view') order by o.order_id loop
  perform 1 from public.production_orders where id=oid for update;
  for e in select x.* from public.production_operation_execution x join public.production_order_operations o on o.id=x.operation_id where o.order_id=oid and article_master_private.readiness_allowed(o.id,'view') order by o.position for update of x loop
   material:=article_master_private.readiness_material(e);
   if material<>e.material_status then
    old_state:=to_jsonb(e);e.material_status:=material;e.row_version:=e.row_version+1;e.updated_at:=now();
    update public.production_operation_execution set material_status=e.material_status,row_version=e.row_version,updated_at=e.updated_at where operation_id=e.operation_id;
    perform article_master_private.readiness_event(e.operation_id,gen_random_uuid(),'forecast_evaluated',jsonb_build_object('date',(now() at time zone 'Europe/Berlin')::date,'source',(select to_jsonb(f) from public.production_operation_arrivals f where f.operation_id=e.operation_id)),old_state,to_jsonb(e));
   end if;
  end loop;
 end loop;
 select coalesce(jsonb_agg(result order by machine_name,ba_number,position),'[]'::jsonb) into rows from(
  select s.ba_number,s.quantity,s.due_date,s.article_data->>'designation' designation,s.article_data->>'customer' customer,b.article_number,o.op_number,o.position,o.predecessor,o.data->>'activity' activity,
  coalesce(ex.machine_id,(o.data->>'machine_id')::uuid) machine_id,coalesce(m.name,'Nicht zugeordnet') machine_name,(m.active and m.department_id=s.department_id) machine_active,
  b.status order_status,o.order_id,o.id operation_id,(o.predecessor is null or exists(select 1 from public.production_order_operations po join public.production_operation_execution pe on pe.operation_id=po.id where po.order_id=o.order_id and po.op_number=o.predecessor and pe.work_status='completed' and pe.good_quantity>=s.quantity)) predecessor_released,coalesce(ex.row_version,0) row_version,ex.planned_start,coalesce(ex.work_status,'planned') work_status,coalesce(ex.material_status,'planned') material_status,
  coalesce(ex.received_quantity,0) received_quantity,coalesce(ex.good_quantity,0) good_quantity,ex.started_at,ex.operation_id is not null initialized,
  f.expected_on,f.valid_until forecast_valid_until,f.source_record forecast_source,f.evidence forecast_evidence,
  article_master_private.readiness_allowed(o.id,'plan') can_plan,article_master_private.readiness_allowed(o.id,'confirm') can_confirm,article_master_private.readiness_allowed(o.id,'work') can_work,
  coalesce((select jsonb_agg(ev) from(select actor_id,happened_at,action,good_quantity,machine_id,before_state,after_state,payload from public.production_readiness_events where operation_id=o.id order by happened_at desc,request_id limit 30) ev),'[]'::jsonb) events
  from public.production_order_operations o join public.production_order_article_snapshots s on s.order_id=o.order_id join public.production_orders b on b.id=o.order_id
  left join public.production_operation_execution ex on ex.operation_id=o.id left join public.production_machines m on m.id=coalesce(ex.machine_id,(o.data->>'machine_id')::uuid)
  left join public.production_operation_arrivals f on f.operation_id=o.id where article_master_private.readiness_allowed(o.id,'view')
 ) result;
 return jsonb_build_object('can_view',visible,'rows',rows,'evaluated_on',(now() at time zone 'Europe/Berlin')::date,'machines',coalesce((select jsonb_agg(jsonb_build_object('id',m.id,'name',m.name,'active',m.active,'department_id',m.department_id)) from public.production_machines m where exists(select 1 from public.article_master_scopes a where a.user_id=auth.uid() and a.department_id=m.department_id and exists(select 1 from public.departments d where d.id=m.department_id and d.active))),'[]'::jsonb));
end; $$;
create function article_master_private.readiness_order_status() returns trigger language plpgsql security definer set search_path='' as $$
declare e public.production_operation_execution;old_state jsonb;
begin
 if new.status is not distinct from old.status or not article_master_private.readiness_enabled() then return new;end if;
 for e in select x.* from public.production_operation_execution x join public.production_order_operations o on o.id=x.operation_id where o.order_id=new.id order by o.position for update of x loop
  if auth.uid() is null then raise exception 'Audit actor required for tracked BA status' using errcode='42501';end if;
  old_state:=to_jsonb(e);
  if new.status='cancelled' then e.work_status:='cancelled';elsif new.status in('paused','completed') and e.work_status='running' then e.work_status:='paused';end if;
  e.row_version:=e.row_version+1;e.updated_at:=now();
  update public.production_operation_execution set work_status=e.work_status,row_version=e.row_version,updated_at=e.updated_at where operation_id=e.operation_id;
  perform article_master_private.readiness_event(e.operation_id,gen_random_uuid(),'ba_status',jsonb_build_object('before',old.status,'after',new.status),old_state,to_jsonb(e));
 end loop;return new;
end; $$;
create trigger readiness_order_status after update of status on public.production_orders for each row execute function article_master_private.readiness_order_status();
create function article_master_private.readiness_forecast_audit() returns trigger language plpgsql security definer set search_path='' as $$
declare e public.production_operation_execution;
begin
 select * into e from public.production_operation_execution where operation_id=new.operation_id;
 if e.operation_id is not null then perform article_master_private.readiness_event(e.operation_id,gen_random_uuid(),'forecast_source',to_jsonb(new),to_jsonb(e),to_jsonb(e),new.verified_by);end if;return new;
end; $$;
create trigger readiness_retain_forecast before delete on public.production_operation_arrivals for each row execute function article_master_private.immutable_snapshot();
create trigger readiness_forecast_audit after insert or update on public.production_operation_arrivals for each row execute function article_master_private.readiness_forecast_audit();
create trigger readiness_immutable_events before update or delete on public.production_readiness_events for each row execute function article_master_private.immutable_snapshot();
create function public.machine_readiness_view() returns jsonb language sql security invoker set search_path='' as $$select article_master_private.readiness_view()$$;
create function public.machine_readiness_initialize(p_order_id uuid) returns void language sql security invoker set search_path='' as $$select article_master_private.readiness_initialize(p_order_id)$$;
create function public.machine_readiness_command(p_operation_id uuid,p_action text,p_request_id uuid,p_expected_version integer,p_quantity integer default null,p_machine_id uuid default null,p_planned_start timestamptz default null,p_reason text default null) returns jsonb language sql security invoker set search_path='' as $$select article_master_private.readiness_command(p_operation_id,p_action,p_request_id,p_expected_version,p_quantity,p_machine_id,p_planned_start,p_reason)$$;
alter table public.production_readiness_settings enable row level security;
alter table public.production_readiness_grants enable row level security;
alter table public.production_operation_execution enable row level security;
alter table public.production_operation_arrivals enable row level security;
alter table public.production_readiness_events enable row level security;
revoke all on public.production_readiness_settings,public.production_readiness_grants,public.production_operation_execution,public.production_operation_arrivals,public.production_readiness_events from public,anon,authenticated;
revoke all on all functions in schema article_master_private from public,anon;
revoke all on function article_master_private.readiness_enabled(),article_master_private.readiness_allowed(uuid,text),article_master_private.readiness_material(public.production_operation_execution),article_master_private.readiness_event(uuid,uuid,text,jsonb,jsonb,jsonb,uuid),article_master_private.readiness_new_operation(),article_master_private.readiness_order_status(),article_master_private.readiness_forecast_audit() from authenticated;
grant execute on function article_master_private.readiness_view(),article_master_private.readiness_initialize(uuid),article_master_private.readiness_command(uuid,text,uuid,integer,integer,uuid,timestamptz,text) to authenticated;
revoke all on function public.machine_readiness_view(),public.machine_readiness_initialize(uuid),public.machine_readiness_command(uuid,text,uuid,integer,integer,uuid,timestamptz,text) from public,anon;
grant execute on function public.machine_readiness_view(),public.machine_readiness_initialize(uuid),public.machine_readiness_command(uuid,text,uuid,integer,integer,uuid,timestamptz,text) to authenticated;
grant all on public.production_readiness_settings,public.production_readiness_grants,public.production_operation_arrivals to service_role;
commit;
