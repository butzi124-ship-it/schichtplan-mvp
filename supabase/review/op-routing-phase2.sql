-- REVIEW ONLY. Requires #7/#8, readiness #9, and Phase 1. Never applied to a hosted DB.
begin;
set local lock_timeout='5s';
set local statement_timeout='20s';
-- Captured legacy catalog allows only admin/employee/tool_scanner while the existing
-- app and original department helper already use department_admin. Explicit review
-- transition, NOT a fixture relaxation or a change applied to production.
do $role_contract$
declare definition text;
begin
 select pg_get_constraintdef(oid) into definition from pg_constraint where conrelid='public.employees'::regclass and conname='employees_role_check';
 if definition is null then raise exception 'Missing authoritative employees role constraint';end if;
 if definition not like '%department_admin%' then
  if definition <> 'CHECK ((role = ANY (ARRAY[''admin''::text, ''employee''::text, ''tool_scanner''::text])))' then raise exception 'Unknown employees role constraint; separate review required';end if;
  alter table public.employees drop constraint employees_role_check;
  alter table public.employees add constraint employees_role_check check(role in('admin','employee','tool_scanner','department_admin'));
 end if;
end;
$role_contract$;
create table public.production_routing_settings(singleton boolean primary key default true check(singleton),enabled boolean not null default false);
insert into public.production_routing_settings values(true,false);
-- Independently provisioned authority; mutable employee.role alone cannot confer AL/admin rights.
create table public.production_routing_authorities(
 user_id uuid not null references auth.users(id),department_id uuid not null references public.departments(id),
 authority text not null check(authority in('leader','deputy','admin')),principal_user_id uuid references auth.users(id),
 valid_from timestamptz not null,valid_until timestamptz not null,revoked boolean not null default false,
 issued_by uuid not null references auth.users(id),primary key(user_id,department_id),check(valid_until>valid_from)
);
create table public.production_routing_capabilities(machine_id uuid not null references public.production_machines(id),machine_group text not null,primary key(machine_id,machine_group));
create table public.production_routing_abort_reasons(code text primary key,label text not null,requires_note boolean not null default false,active boolean not null default true);
-- Immutable technical constraints, provisioned only from reviewed article/BA requirements.
-- No backfill or guessed binding is inferred from an optional historical machine_id.
create table public.production_operation_machine_constraints(operation_id uuid primary key references public.production_order_operations(id),machine_id uuid not null references public.production_machines(id),evidence text not null check(length(trim(evidence))>0));
create table public.production_operation_routing(
 operation_id uuid primary key references public.production_order_operations(id),original_department_id uuid not null references public.departments(id),
 executing_department_id uuid not null references public.departments(id),default_machine_id uuid references public.production_machines(id),
 machine_id uuid not null references public.production_machines(id),binding boolean not null default false,
 row_version integer not null default 1 check(row_version>0)
);
create table public.production_operation_machine_sections(
 id uuid primary key default gen_random_uuid(),operation_id uuid not null references public.production_order_operations(id),
 machine_id uuid not null references public.production_machines(id),department_id uuid not null references public.departments(id),
 opened_at timestamptz not null,closed_at timestamptz,opened_by uuid not null references auth.users(id),closed_by uuid references auth.users(id),
 close_reason text,opening_execution jsonb not null,closing_execution jsonb,
 check(closed_at is null or closed_at>=opened_at)
);
create unique index routing_one_open_section on public.production_operation_machine_sections(operation_id) where closed_at is null;
-- Minimal Phase-3 bridge. No public timer-start API, no attendance account and no new piece counter.
create table public.production_operation_activity_segments(
 id uuid primary key default gen_random_uuid(),section_id uuid not null references public.production_operation_machine_sections(id),
 user_id uuid not null references auth.users(id),activity text not null check(activity in('setup_internal','setup_external','production')),
 started_at timestamptz not null,ended_at timestamptz,end_reason text,
 check(ended_at is null or ended_at>=started_at)
);
create unique index routing_one_active_personal_segment on public.production_operation_activity_segments(user_id) where ended_at is null;
create table public.production_operation_routing_events(
 request_id uuid primary key,operation_id uuid not null references public.production_order_operations(id),actor_id uuid not null references auth.users(id),
 happened_at timestamptz not null,action text not null,payload jsonb not null,before_state jsonb,after_state jsonb not null,interrupted_segments jsonb not null
);
create table public.production_operation_routing_notifications(
 id uuid primary key default gen_random_uuid(),request_id uuid not null references public.production_operation_routing_events(request_id),
 recipient_user_id uuid not null references auth.users(id),department_id uuid not null references public.departments(id),
 kind text not null check(kind in('transferred','returned')),created_at timestamptz not null,
 unique(request_id,recipient_user_id)
);
create function article_master_private.routing_enabled() returns boolean language sql stable security definer set search_path='' as $$
 select coalesce((select enabled from public.production_routing_settings where singleton),false) and article_master_private.readiness_enabled()
$$;
create function article_master_private.routing_active() returns boolean language sql stable security definer set search_path='' as $$
 select auth.uid() is not null and exists(select 1 from public.employees where auth_user_id=auth.uid() and is_active and role in('employee','department_admin','admin'))
$$;
create function article_master_private.routing_manager(dept uuid,actor uuid default auth.uid()) returns boolean language sql stable security definer set search_path='' as $$
 select exists(select 1 from public.production_routing_authorities a join public.employees e on e.auth_user_id=a.user_id
 join public.departments d on d.id=a.department_id
 where a.user_id=actor and a.department_id=dept and d.active and e.is_active and not a.revoked and a.valid_from<=now() and a.valid_until>now()
 and ((a.authority='admin' and e.role='admin') or (a.authority='leader' and e.role='department_admin' and d.leader_employee_id=e.id)
 or(a.authority='deputy' and e.role in('employee','department_admin') and exists(select 1 from public.employees p where p.auth_user_id=a.principal_user_id and p.id=d.leader_employee_id and p.is_active and p.role='department_admin'))))
$$;
create function article_master_private.routing_department(op uuid) returns uuid language sql stable security definer set search_path='' as $$
 select coalesce(r.executing_department_id,s.department_id) from public.production_order_operations o
 join public.production_order_article_snapshots s on s.order_id=o.order_id left join public.production_operation_routing r on r.operation_id=o.id where o.id=op
$$;
create function article_master_private.routing_operator(op uuid,target uuid) returns boolean language sql stable security definer set search_path='' as $$
 select article_master_private.routing_active() and exists(select 1 from public.production_machines m
 join public.departments d on d.id=m.department_id join public.employees e on e.department_id=d.id and e.auth_user_id=auth.uid()
 join public.production_readiness_grants g on g.user_id=auth.uid() and g.machine_id=m.id
 where m.id=target and m.active and d.active and e.is_active and e.role in('employee','department_admin','admin') and m.department_id=article_master_private.routing_department(op)
 and not g.revoked and g.expires_at>now() and g.can_work)
$$;
create function article_master_private.routing_visible(op uuid) returns boolean language sql stable security definer set search_path='' as $$
 select article_master_private.routing_enabled() and article_master_private.routing_active() and exists(
 select 1 from public.production_order_operations o join public.production_order_article_snapshots s on s.order_id=o.order_id
 left join public.production_operation_routing r on r.operation_id=o.id join public.departments d on d.id=s.department_id
 where o.id=op and d.active and exists(select 1 from public.departments ed where ed.id=article_master_private.routing_department(op) and ed.active)
 and (article_master_private.routing_manager(s.department_id) or article_master_private.routing_manager(article_master_private.routing_department(op))
 or article_master_private.routing_operator(op,coalesce(r.machine_id,(o.data->>'machine_id')::uuid))
 or (article_master_private.routing_department(op)=s.department_id and article_master_private.ba_entry_allowed(s.department_id,(select article_number from public.production_orders where id=o.order_id),(s.operations->0->>'number')::integer))))
$$;
create function article_master_private.routing_suitable(op uuid,target uuid) returns boolean language sql stable security definer set search_path='' as $$
 select exists(select 1 from public.production_order_operations o join public.production_machines m on m.id=target join public.departments d on d.id=m.department_id
 where o.id=op and m.active and d.active and (m.id=(o.data->>'machine_id')::uuid or exists(select 1 from public.production_routing_capabilities c where c.machine_id=m.id and c.machine_group=o.data->>'machine_group')))
$$;
-- Existing assignment APIs cannot bypass a tracked BA-OP binding or desynchronize routing.
create function article_master_private.routing_execution_guard() returns trigger language plpgsql security definer set search_path='' as $$
begin
 if exists(select 1 from public.production_operation_routing where operation_id=new.operation_id and machine_id is distinct from new.machine_id)
 then raise exception 'Use checked OP routing command' using errcode='42501';end if;
 return new;
end; $$;
create trigger routing_execution_guard before insert or update of machine_id on public.production_operation_execution for each row execute function article_master_private.routing_execution_guard();
create function article_master_private.routing_section_guard() returns trigger language plpgsql set search_path='' as $$
begin
 if tg_op='DELETE' or old.closed_at is not null or (new.id,new.operation_id,new.machine_id,new.department_id,new.opened_at,new.opened_by,new.opening_execution) is distinct from
 (old.id,old.operation_id,old.machine_id,old.department_id,old.opened_at,old.opened_by,old.opening_execution) then raise exception 'Historical machine section retained' using errcode='42501';end if;
 return new;
end; $$;
create trigger routing_section_guard before update or delete on public.production_operation_machine_sections for each row execute function article_master_private.routing_section_guard();
create function article_master_private.routing_activity_guard() returns trigger language plpgsql set search_path='' as $$
begin
 if tg_op='DELETE' or old.ended_at is not null or (new.id,new.section_id,new.user_id,new.activity,new.started_at) is distinct from (old.id,old.section_id,old.user_id,old.activity,old.started_at)
 then raise exception 'Personal history retained' using errcode='42501';end if;return new;
end; $$;
create trigger routing_activity_guard before update or delete on public.production_operation_activity_segments for each row execute function article_master_private.routing_activity_guard();
create function article_master_private.routing_activity_insert_guard() returns trigger language plpgsql security definer set search_path='' as $$
declare oid uuid;closed timestamptz;
begin
 select o.order_id into oid from public.production_operation_machine_sections s join public.production_order_operations o on o.id=s.operation_id where s.id=new.section_id;
 perform 1 from public.production_orders where id=oid for update;
 select closed_at into closed from public.production_operation_machine_sections where id=new.section_id for update;
 if new.ended_at is null and closed is not null then raise exception 'Machine section already closed' using errcode='PT409';end if;
 return new;
end; $$;
create trigger routing_activity_insert_guard before insert on public.production_operation_activity_segments for each row execute function article_master_private.routing_activity_insert_guard();
create trigger routing_immutable_events before update or delete on public.production_operation_routing_events for each row execute function article_master_private.immutable_snapshot();
create trigger routing_immutable_notifications before update or delete on public.production_operation_routing_notifications for each row execute function article_master_private.immutable_snapshot();
create trigger routing_immutable_constraints before update or delete on public.production_operation_machine_constraints for each row execute function article_master_private.immutable_snapshot();

create function article_master_private.routing_command(op uuid,action text,request uuid,expected integer,target uuid,reason text,abort_code text)
returns jsonb language plpgsql security definer set search_path='' as $$
declare o public.production_order_operations;s public.production_order_article_snapshots;e public.production_operation_execution;r public.production_operation_routing;
 prior public.production_operation_routing_events; body jsonb; previous jsonb; interrupted jsonb:='[]';section uuid;target_dept uuid;stamp timestamptz;
 old_dept uuid;has_setup boolean;required_machine uuid;can_manage boolean;changed boolean;status text;
begin
 if not article_master_private.routing_visible(op) then raise exception 'Denied' using errcode='42501';end if;
 if action is null or action not in('select','fix','return') or request is null then raise exception 'Invalid routing action' using errcode='22023';end if;
 select * into o from public.production_order_operations where id=op;
 perform 1 from public.production_orders where id=o.order_id for update;
 select * into s from public.production_order_article_snapshots where order_id=o.order_id;
 select * into e from public.production_operation_execution where operation_id=op for update;
 select * into r from public.production_operation_routing where operation_id=op for update;
 -- Serialize authority changes with the business transaction; sorted department rows avoid cross-transfer deadlocks.
 perform 1 from public.production_routing_settings where singleton for share;
 perform 1 from public.departments order by id for share;
 perform 1 from public.employees where auth_user_id=auth.uid() or auth_user_id in(select user_id from public.production_routing_authorities union select principal_user_id from public.production_routing_authorities) order by id for share;
 perform 1 from public.production_routing_authorities order by department_id,user_id for share;
 perform 1 from public.production_readiness_grants where user_id=auth.uid() for share;
 if not article_master_private.routing_visible(op) then raise exception 'Authorization changed' using errcode='42501';end if;
 body:=jsonb_build_object('action',action,'target',target,'reason',reason,'abort_code',abort_code,'expected',expected);
 select * into prior from public.production_operation_routing_events where request_id=request;
 if found then
  if prior.operation_id<>op or prior.actor_id<>auth.uid() or prior.payload<>body then raise exception 'Request reused' using errcode='22023';end if;
  return prior.after_state;
 end if;
 if coalesce(r.row_version,0) is distinct from expected then raise exception 'Stale routing' using errcode='PT409';end if;
 select b.status into status from public.production_orders b where b.id=o.order_id;
 if status not in('running','paused') or e.work_status in('completed','cancelled') then raise exception 'Terminal BA/OP' using errcode='22023';end if;
 if reason is null or length(trim(reason)) not between 1 and 4000 then raise exception 'Reason required' using errcode='22023';end if;
 old_dept:=coalesce(r.executing_department_id,s.department_id);
 can_manage:=article_master_private.routing_manager(old_dept);
 if action in('fix','return') and not can_manage then raise exception 'AL or active deputy required' using errcode='42501';end if;
 if action='select' and (coalesce(r.binding,false) or not article_master_private.routing_operator(op,target)) then raise exception 'Bound or unauthorized machine' using errcode='42501';end if;
 if action='return' and old_dept=s.department_id then raise exception 'Already in original department' using errcode='22023';end if;
 select machine_id into required_machine from public.production_operation_machine_constraints where operation_id=op;
 if target is null or not article_master_private.routing_suitable(op,target) or required_machine is not null and target<>required_machine then raise exception 'Unsuitable or mandatory machine' using errcode='42501';end if;
 select department_id into target_dept from public.production_machines where id=target for share;
 perform 1 from public.production_routing_capabilities where machine_id=target for share;
 if not article_master_private.routing_suitable(op,target) then raise exception 'Target changed' using errcode='42501';end if;
 if action='return' and target_dept<>s.department_id then raise exception 'Return to original department required' using errcode='42501';end if;
 if target_dept<>old_dept and not exists(select 1 from public.production_routing_authorities a where a.department_id=target_dept and article_master_private.routing_manager(target_dept,a.user_id))
 then raise exception 'No verified receiving AL/notification recipient' using errcode='22023';end if;
 -- Explicit adoption retains preexisting execution and never guesses previous machine sections.
 previous:=jsonb_build_object('routing',to_jsonb(r),'execution',to_jsonb(e));
 if r.operation_id is null then
  if e.operation_id is null then
   insert into public.production_operation_execution(operation_id,machine_id) values(op,(o.data->>'machine_id')::uuid) returning * into e;
  end if;
  insert into public.production_operation_routing(operation_id,original_department_id,executing_department_id,default_machine_id,machine_id,binding)
   values(op,s.department_id,s.department_id,(o.data->>'machine_id')::uuid,coalesce(e.machine_id,target),required_machine is not null) returning * into r;
  if e.machine_id is not null then
   insert into public.production_operation_machine_sections(operation_id,machine_id,department_id,opened_at,opened_by,opening_execution)
    values(op,e.machine_id,s.department_id,clock_timestamp(),auth.uid(),to_jsonb(e)) returning id into section;
  end if;
 end if;
 changed:=e.machine_id is distinct from target;
 perform 1 from public.production_operation_activity_segments a join public.production_operation_machine_sections x on x.id=a.section_id where x.operation_id=op and a.ended_at is null for update of a;
 select exists(select 1 from public.production_operation_activity_segments a join public.production_operation_machine_sections x on x.id=a.section_id where x.operation_id=op and a.ended_at is null and a.activity in('setup_internal','setup_external')) into has_setup;
 if changed and has_setup and not exists(select 1 from public.production_routing_abort_reasons where code=abort_code and active and (not requires_note or length(trim(reason))>0))
 then raise exception 'Configured setup-abort reason required' using errcode='22023';end if;
 stamp:=clock_timestamp();
 if changed then
  with closed as(update public.production_operation_activity_segments a set ended_at=stamp,end_reason=case when a.activity like 'setup_%' then 'machine_change:'||abort_code else action end
   from public.production_operation_machine_sections x where x.id=a.section_id and x.operation_id=op and a.ended_at is null returning a.*)
  select coalesce(jsonb_agg(to_jsonb(closed)),'[]') into interrupted from closed;
  update public.production_operation_machine_sections set closed_at=stamp,closed_by=auth.uid(),close_reason=reason,closing_execution=to_jsonb(e) where operation_id=op and closed_at is null;
  -- Preserve cumulative good quantity and first actual start. A new machine is not physically released or automatically started.
  e.machine_id:=target;e.started_here:=false;e.received_quantity:=0;e.material_status:='planned';e.work_status:=case when e.started_at is null then 'planned' else 'paused' end;
  e.row_version:=e.row_version+1;e.updated_at:=stamp;
 end if;
 r.machine_id:=target;r.executing_department_id:=target_dept;r.binding:=r.binding or action='fix';r.row_version:=r.row_version+1;
 update public.production_operation_routing set machine_id=r.machine_id,executing_department_id=r.executing_department_id,binding=r.binding,row_version=r.row_version where operation_id=op;
 update public.production_operation_execution set machine_id=e.machine_id,started_here=e.started_here,received_quantity=e.received_quantity,material_status=e.material_status,work_status=e.work_status,row_version=e.row_version,updated_at=e.updated_at where operation_id=op;
 if changed or section is null and not exists(select 1 from public.production_operation_machine_sections where operation_id=op and closed_at is null) then
  insert into public.production_operation_machine_sections(operation_id,machine_id,department_id,opened_at,opened_by,opening_execution) values(op,target,target_dept,stamp,auth.uid(),to_jsonb(e));
 end if;
 insert into public.production_operation_routing_events values(request,op,auth.uid(),stamp,action,body,previous,to_jsonb(r),interrupted);
 if target_dept<>old_dept then
  insert into public.production_operation_routing_notifications(request_id,recipient_user_id,department_id,kind,created_at)
  select request,a.user_id,target_dept,case when action='return' then 'returned' else 'transferred' end,stamp from public.production_routing_authorities a
  where a.department_id=target_dept and article_master_private.routing_manager(target_dept,a.user_id);
 end if;
 return to_jsonb(r);
end; $$;
create function article_master_private.routing_view() returns jsonb language plpgsql stable security definer set search_path='' as $$
declare rows jsonb;notes jsonb;
begin
 if not article_master_private.routing_enabled() or not article_master_private.routing_active() then return jsonb_build_object('rows','[]'::jsonb,'notifications','[]'::jsonb);end if;
 select coalesce(jsonb_agg(result order by ba_number,op_number),'[]'::jsonb) into rows from(
 select o.id operation_id,o.order_id,s.ba_number,o.op_number,o.data->>'activity' activity,(o.data->>'machine_id')::uuid default_machine_id,
 s.department_id original_department_id,od.name original_department,article_master_private.routing_department(o.id) executing_department_id,ed.name executing_department,
 coalesce(r.machine_id,e.machine_id,(o.data->>'machine_id')::uuid) machine_id,(select name from public.production_machines where id=coalesce(r.machine_id,e.machine_id,(o.data->>'machine_id')::uuid)) machine_name,coalesce(r.row_version,0) row_version,coalesce(r.binding,false) binding,
 c.machine_id mandatory_machine_id,coalesce(e.work_status,'planned') work_status,e.good_quantity,
 (b.status in('running','paused') and coalesce(e.work_status,'planned') not in('completed','cancelled')) mutable,
 article_master_private.routing_manager(article_master_private.routing_department(o.id)) can_manage,
 (s.department_id<>article_master_private.routing_department(o.id) and article_master_private.routing_manager(article_master_private.routing_department(o.id))) can_return,
 coalesce((select jsonb_agg(jsonb_build_object('id',m.id,'name',m.name,'department_id',m.department_id) order by m.name)
 from public.production_machines m where article_master_private.routing_suitable(o.id,m.id) and (c.machine_id is null or m.id=c.machine_id)
 and (article_master_private.routing_manager(article_master_private.routing_department(o.id)) or not coalesce(r.binding,false) and article_master_private.routing_operator(o.id,m.id))),'[]'::jsonb) machines,
 exists(select 1 from public.production_operation_activity_segments a join public.production_operation_machine_sections x on x.id=a.section_id where x.operation_id=o.id and a.ended_at is null and a.activity like 'setup_%') active_setup,
 coalesce((select jsonb_agg(to_jsonb(x)||jsonb_build_object('machine_name',(select name from public.production_machines where id=x.machine_id)) order by x.opened_at,x.id) from public.production_operation_machine_sections x where x.operation_id=o.id),'[]'::jsonb) sections,
 coalesce((select jsonb_agg(to_jsonb(ev) order by ev.happened_at,ev.request_id) from public.production_operation_routing_events ev where ev.operation_id=o.id),'[]'::jsonb) events
 from public.production_order_operations o join public.production_order_article_snapshots s on s.order_id=o.order_id join public.production_orders b on b.id=o.order_id
 join public.departments od on od.id=s.department_id join public.departments ed on ed.id=article_master_private.routing_department(o.id)
 left join public.production_operation_routing r on r.operation_id=o.id left join public.production_operation_execution e on e.operation_id=o.id
 left join public.production_operation_machine_constraints c on c.operation_id=o.id where article_master_private.routing_visible(o.id)) result;
 select coalesce(jsonb_agg(jsonb_build_object('id',n.id,'kind',n.kind,'created_at',n.created_at,'ba_number',s.ba_number,'op_number',o.op_number,'operation_id',o.id) order by n.created_at desc),'[]') into notes
 from public.production_operation_routing_notifications n join public.production_operation_routing_events ev on ev.request_id=n.request_id
 join public.production_order_operations o on o.id=ev.operation_id join public.production_order_article_snapshots s on s.order_id=o.order_id
 where n.recipient_user_id=auth.uid() and article_master_private.routing_manager(n.department_id);
 return jsonb_build_object('rows',rows,'notifications',notes,'abort_reasons',coalesce((select jsonb_agg(to_jsonb(a) order by label) from public.production_routing_abort_reasons a where active),'[]'::jsonb));
end; $$;
-- Reuse existing work/physical-handover state machine, with checked current-department authority for tracked OPs.
alter function article_master_private.readiness_allowed(uuid,text) rename to readiness_allowed_before_routing;
create function article_master_private.readiness_allowed(op uuid,permission text) returns boolean language sql stable security definer set search_path='' as $$
 select case when exists(select 1 from public.production_operation_routing where operation_id=op) then
 article_master_private.routing_visible(op) and (case when permission='view' then true when permission='plan' then article_master_private.routing_manager(article_master_private.routing_department(op))
 else article_master_private.routing_manager(article_master_private.routing_department(op)) or exists(
 select 1 from public.production_operation_routing r join public.production_readiness_grants g on g.machine_id=r.machine_id
 join public.employees e on e.auth_user_id=g.user_id where r.operation_id=op and g.user_id=auth.uid() and e.is_active
 and e.department_id=r.executing_department_id and not g.revoked and g.expires_at>now() and case permission when 'work' then g.can_work when 'confirm' then g.can_confirm else false end) end)
 else article_master_private.readiness_allowed_before_routing(op,permission) end
$$;
-- Fail closed on an unexpected prerequisite body; adapt only the three explicit department checks.
do $routing_adapters$
declare definition text;needle text;
begin
 definition:=pg_get_functiondef('article_master_private.readiness_command(uuid,text,uuid,integer,integer,uuid,timestamptz,text)'::regprocedure);
 needle:='id=e.machine_id and department_id=s.department_id and active';
 if (length(definition)-length(replace(definition,needle,'')))/length(needle)<>1 then raise exception 'Unexpected readiness command prerequisite';end if;
 definition:=replace(definition,needle,'id=e.machine_id and department_id=article_master_private.routing_department(o.id) and active');
 needle:='id=p_machine_id and department_id=s.department_id and active';
 if (length(definition)-length(replace(definition,needle,'')))/length(needle)<>1 then raise exception 'Unexpected readiness planning prerequisite';end if;
 execute replace(definition,needle,'id=p_machine_id and department_id=article_master_private.routing_department(o.id) and active');
 definition:=pg_get_functiondef('article_master_private.readiness_view()'::regprocedure);
 needle:='(m.active and m.department_id=s.department_id) machine_active';
 if (length(definition)-length(replace(definition,needle,'')))/length(needle)<>1 then raise exception 'Unexpected readiness view prerequisite';end if;
 execute replace(definition,needle,'(m.active and m.department_id=article_master_private.routing_department(o.id)) machine_active');
end;
$routing_adapters$;
create function public.op_routing_view() returns jsonb language sql stable security invoker set search_path='' as $$select article_master_private.routing_view()$$;
create function public.op_routing_command(p_operation_id uuid,p_action text,p_request_id uuid,p_expected_version integer,p_machine_id uuid,p_reason text,p_abort_code text default null)
returns jsonb language sql security invoker set search_path='' as $$select article_master_private.routing_command(p_operation_id,p_action,p_request_id,p_expected_version,p_machine_id,p_reason,p_abort_code)$$;
do $permissions$
declare t text;f record;
begin
 foreach t in array array['production_routing_settings','production_routing_authorities','production_routing_capabilities','production_routing_abort_reasons','production_operation_machine_constraints','production_operation_routing','production_operation_machine_sections','production_operation_activity_segments','production_operation_routing_events','production_operation_routing_notifications'] loop
  execute format('alter table public.%I enable row level security',t);
  execute format('revoke all on public.%I from public,anon,authenticated',t);
 end loop;
 for f in select p.oid::regprocedure signature from pg_proc p join pg_namespace n on n.oid=p.pronamespace where n.nspname='article_master_private' and (p.proname like 'routing_%' or p.proname='readiness_allowed_before_routing') loop
  execute format('revoke all on function %s from public,anon,authenticated',f.signature);
 end loop;
end;
$permissions$;
revoke all on function article_master_private.readiness_allowed(uuid,text) from public,anon,authenticated;
revoke all on function public.op_routing_view(),public.op_routing_command(uuid,text,uuid,integer,uuid,text,text) from public,anon,authenticated;
grant execute on function article_master_private.routing_view(),article_master_private.routing_command(uuid,text,uuid,integer,uuid,text,text),public.op_routing_view(),public.op_routing_command(uuid,text,uuid,integer,uuid,text,text) to authenticated;
commit;
