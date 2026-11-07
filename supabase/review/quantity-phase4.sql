-- REVIEW ONLY. Requires phases 1–3; no hosted execution or activation.
-- Existing station counts/totals/events are the ONLY quantity store.
begin;
set local lock_timeout='5s';
set local statement_timeout='20s';
create table public.production_quantity_settings(singleton boolean primary key default true check(singleton),enabled boolean not null default false);
insert into public.production_quantity_settings values(true,false);
create table public.production_quantity_plans(operation_id uuid primary key references public.production_order_operations(id),created_by uuid not null references auth.users(id),created_at timestamptz not null default now());
create table public.production_quantity_stations(station_id uuid primary key references public.production_order_stations(id),operation_id uuid not null references public.production_quantity_plans(operation_id),position integer not null check(position>0),unique(operation_id,position));
-- Attribution references the existing event, does not duplicate its quantity.
create table public.production_quantity_attributions(event_id uuid primary key references public.production_station_events(id),operation_id uuid not null references public.production_order_operations(id),section_id uuid not null references public.production_operation_machine_sections(id),actor_id uuid not null references auth.users(id));
create table public.production_quantity_events(request_id uuid primary key,actor_id uuid not null references auth.users(id),operation_id uuid not null references public.production_order_operations(id),happened_at timestamptz not null default now(),body jsonb not null,before_value jsonb,after_value jsonb not null);
-- Detection only. No resolution, quantity-change or completion API in Phase 4.
create table public.production_quantity_conflicts(id uuid primary key default gen_random_uuid(),operation_id uuid not null references public.production_order_operations(id),station_id uuid not null unique references public.production_order_stations(id),first_request uuid not null references public.production_quantity_events(request_id),first_detected timestamptz not null default now(),initial_available bigint not null,initial_counted bigint not null);
create table public.production_quantity_notifications(conflict_id uuid not null references public.production_quantity_conflicts(id),recipient_user_id uuid not null references auth.users(id),primary key(conflict_id,recipient_user_id));
create function article_master_private.quantity_enabled() returns boolean language sql stable security definer set search_path='' as $$
 select article_master_private.routing_enabled() and coalesce((select enabled from public.production_quantity_settings where singleton),false)
$$;
create function article_master_private.quantity_good(station uuid) returns bigint language sql stable security definer set search_path='' as $$
 select coalesce(sum(good_qty),0)::bigint from public.production_station_counts where station_id=station
$$;
create function article_master_private.quantity_finished(op uuid) returns bigint language sql stable security definer set search_path='' as $$
 select article_master_private.quantity_good((select station_id from public.production_quantity_stations where operation_id=op order by position desc limit 1))
$$;
create function article_master_private.quantity_available(station uuid) returns bigint language sql stable security definer set search_path='' as $$
 select case when q.position>1 then article_master_private.quantity_good((select station_id from public.production_quantity_stations where operation_id=q.operation_id and position=q.position-1))
 when o.predecessor is not null then article_master_private.quantity_finished((select id from public.production_order_operations where order_id=o.order_id and op_number=o.predecessor))
 else b.target_quantity::bigint end
 from public.production_quantity_stations q join public.production_order_operations o on o.id=q.operation_id join public.production_orders b on b.id=o.order_id where q.station_id=station
$$;
create function article_master_private.quantity_row(op uuid) returns jsonb language sql stable security definer set search_path='' as $$
 select jsonb_build_object('operation_id',o.id,'order_id',o.order_id,'op_number',o.op_number,'ba_number',b.ba_number,'target',b.target_quantity,'original_target',s.quantity,
 'finished_good',article_master_private.quantity_finished(o.id),'ba_finished_good',article_master_private.quantity_finished((select last_op.id from public.production_order_operations last_op where last_op.order_id=o.order_id order by last_op.position desc limit 1)),'predecessor',o.predecessor,
 'available',case when o.predecessor is null then b.target_quantity::bigint else article_master_private.quantity_finished((select id from public.production_order_operations where order_id=o.order_id and op_number=o.predecessor)) end,
 'configured',exists(select 1 from public.production_quantity_plans where operation_id=o.id),
 'can_work',exists(select 1 from public.production_operation_machine_sections x where x.operation_id=o.id and x.closed_at is null and (article_master_private.routing_manager(x.department_id) or article_master_private.routing_operator(o.id,x.machine_id))),
 'can_manage',article_master_private.routing_manager(article_master_private.routing_department(o.id)),
 'section_id',(select id from public.production_operation_machine_sections where operation_id=o.id and closed_at is null),
 'stations',coalesce((select jsonb_agg(to_jsonb(st)||jsonb_build_object('counts',coalesce((select jsonb_agg(to_jsonb(pc)) from public.production_station_counts pc where pc.station_id=st.id),'[]'::jsonb),'position',q.position,'good',article_master_private.quantity_good(st.id),'available',article_master_private.quantity_available(st.id),'conflict_id',c.id) order by q.position)
 from public.production_quantity_stations q join public.production_order_stations st on st.id=q.station_id left join public.production_quantity_conflicts c on c.station_id=st.id where q.operation_id=o.id),'[]'::jsonb),
 'unbound_stations',coalesce((select jsonb_agg(jsonb_build_object('id',st.id,'station_no',st.station_no,'op_number',st.op_number)) from public.production_order_stations st where st.order_id=o.order_id and not exists(select 1 from public.production_quantity_stations q where q.station_id=st.id)),'[]'::jsonb),
 'sections',coalesce((select jsonb_agg(jsonb_build_object('id',x.id,'machine_id',x.machine_id,'machine_name',m.name,'closed_at',x.closed_at,
 'finished_good',coalesce((select sum(ev.qty) from public.production_quantity_attributions a join public.production_station_events ev on ev.id=a.event_id where a.section_id=x.id and ev.event_type in('good','good_correction') and ev.station_id=(select station_id from public.production_quantity_stations where operation_id=o.id order by position desc limit 1)),0)))
 from public.production_operation_machine_sections x join public.production_machines m on m.id=x.machine_id where x.operation_id=o.id),'[]'::jsonb),
 'own_person_id',(select pe.id from public.production_order_employees pe join public.employees e on e.id=pe.employee_id where pe.order_id=o.order_id and pe.active and e.auth_user_id=auth.uid() limit 1),
 'employees',coalesce((select jsonb_agg(to_jsonb(e)) from public.production_order_employees e where e.order_id=o.order_id),'[]'::jsonb))
 from public.production_order_operations o join public.production_orders b on b.id=o.order_id join public.production_order_article_snapshots s on s.order_id=o.order_id where o.id=op
$$;
-- Raw writes to a tracked BA cannot bypass atomic commands, even with old legacy grants.
-- SECURITY INVOKER: current_user retains the calling role. Checked commands run as
-- the review owner; activation requires verifying the owner/ACL as documented.
-- Use a definer read helper so legacy callers cannot gain raw reads of private plans.
create function article_master_private.quantity_tracked(oid uuid) returns boolean language sql stable security definer set search_path='' as $$
 select exists(select 1 from public.production_order_operations o join public.production_quantity_plans p on p.operation_id=o.id where o.order_id=oid)
$$;
create function article_master_private.quantity_station_tracked(sid uuid) returns boolean language sql stable security definer set search_path='' as $$select exists(select 1 from public.production_quantity_stations where station_id=sid)$$;
-- Replace guard's private-table query with the dedicated read-only helper.
create or replace function article_master_private.quantity_raw_guard() returns trigger language plpgsql set search_path='' as $$
declare tracked boolean;owner_name name;
begin
 select pg_get_userbyid(relowner) into owner_name from pg_class where oid='public.production_quantity_plans'::regclass;
 if tg_op<>'INSERT' then tracked:=article_master_private.quantity_tracked(old.order_id) or article_master_private.quantity_station_tracked((case when tg_table_name='production_order_stations' then to_jsonb(old)->>'id' else to_jsonb(old)->>'station_id' end)::uuid);end if;
 if tg_op<>'DELETE' then tracked:=coalesce(tracked,false) or article_master_private.quantity_tracked(new.order_id) or article_master_private.quantity_station_tracked((case when tg_table_name='production_order_stations' then to_jsonb(new)->>'id' else to_jsonb(new)->>'station_id' end)::uuid);end if;
 if tracked and current_user<>owner_name then raise exception 'Use checked Phase 4 command' using errcode='42501';end if;
 if tg_op='DELETE' then return old;end if;return new;
end; $$;
create trigger quantity_guard_counts before insert or update or delete on public.production_station_counts for each row execute function article_master_private.quantity_raw_guard();
create trigger quantity_guard_events before insert or update or delete on public.production_station_events for each row execute function article_master_private.quantity_raw_guard();
create trigger quantity_guard_stations before insert or update or delete on public.production_order_stations for each row execute function article_master_private.quantity_raw_guard();
create function article_master_private.quantity_command(op uuid,action text,request uuid,station uuid,person uuid,kind text,delta integer,section uuid,cause uuid,note text,stations uuid[],new_count integer) returns jsonb language plpgsql security definer set search_path='' as $$
declare o public.production_order_operations;b public.production_orders;x public.production_operation_machine_sections;st public.production_order_stations;pe public.production_order_employees;
 prior public.production_quantity_events;body jsonb;before_value jsonb;result jsonb;sid uuid;eid uuid;conf uuid;pos integer:=0;start_no integer;current_qty bigint;row record;event_kind text;
begin
 if not article_master_private.quantity_enabled() or not article_master_private.routing_visible(op) then raise exception 'Denied' using errcode='42501';end if;
 if request is null or action is null or action not in('configure','join','count') then raise exception 'Invalid command' using errcode='22023';end if;
 -- Same global lock order as Phase 3/routing; serializes BA flows and machine changes.
 perform pg_advisory_xact_lock(73103,1);
 select * into o from public.production_order_operations where id=op;
 select * into b from public.production_orders where id=o.order_id for update;
 body:=jsonb_build_object('action',action,'station',station,'person',person,'kind',kind,'delta',delta,'section',section,'cause',cause,'note',note,'stations',stations,'new_count',new_count);
 select * into prior from public.production_quantity_events where request_id=request;
 if prior.request_id is not null then
  if prior.actor_id<>auth.uid() or prior.operation_id<>op or prior.body<>body then raise exception 'Request reuse' using errcode='22023';end if;return prior.after_value;
 end if;
 if b.status not in('running','paused') or exists(select 1 from public.production_operation_execution where operation_id=op and work_status in('completed','cancelled')) then raise exception 'Terminal operation' using errcode='22023';end if;
 select * into x from public.production_operation_machine_sections where operation_id=op and closed_at is null for update;
 if x.id is null or not (article_master_private.routing_manager(x.department_id) or article_master_private.routing_operator(op,x.machine_id)) then raise exception 'No current machine permission' using errcode='42501';end if;
 before_value:=article_master_private.quantity_row(op);
 if action in('configure','join') then
  if not exists(select 1 from public.production_order_employees own_entry join public.employees e on e.id=own_entry.employee_id where own_entry.order_id=o.order_id and e.auth_user_id=auth.uid() and own_entry.active) then
   insert into public.production_order_employees(order_id,employee_id,employee_name)
   select o.order_id,e.id,coalesce(nullif(trim(coalesce(to_jsonb(e)->>'first_name','')||' '||coalesce(to_jsonb(e)->>'last_name','')),''),to_jsonb(e)->>'personnel_no','Mitarbeiter') from public.employees e where e.auth_user_id=auth.uid() and e.is_active
   on conflict(order_id,employee_id) do update set active=true,updated_at=now();
  end if;
 end if;
 if action='configure' then
  if exists(select 1 from public.production_quantity_plans where operation_id=op) then raise exception 'Plan immutable after confirmation' using errcode='PT409';end if;
  if coalesce(cardinality(stations),0)+coalesce(new_count,0) not between 1 and 99 or coalesce(new_count,0)<0 or length(trim(coalesce(note,'')))=0 then raise exception 'Explicit ordered station plan and reason required' using errcode='22023';end if;
  -- Existing historical quantities may only be adopted with independent AL authority.
  if exists(select 1 from public.production_order_stations z where z.id=any(stations) and (z.scrap_total<>0 or z.clarify_total<>0 or exists(select 1 from public.production_station_counts c where c.station_id=z.id and c.good_qty<>0) or exists(select 1 from public.production_station_events ev where ev.station_id=z.id))) and not article_master_private.routing_manager(x.department_id) then raise exception 'Historical adoption requires AL' using errcode='42501';end if;
  insert into public.production_quantity_plans(operation_id,created_by) values(op,auth.uid());
  foreach sid in array coalesce(stations,'{}'::uuid[]) loop
   if not exists(select 1 from public.production_order_stations where id=sid and order_id=o.order_id) then raise exception 'Foreign station' using errcode='22023';end if;
   if exists(select 1 from public.production_station_counts pc left join public.production_order_employees person on person.id=pc.order_employee_id where pc.station_id=sid and (pc.order_id<>o.order_id or person.order_id is distinct from o.order_id)) then raise exception 'Historical counter relationship conflict; preserve for separate review' using errcode='22023';end if;
   pos:=pos+1;insert into public.production_quantity_stations values(sid,op,pos);
  end loop;
  select coalesce(max(station_no),0) into start_no from public.production_order_stations where order_id=o.order_id;
  for i in 1..coalesce(new_count,0) loop
   insert into public.production_order_stations(order_id,station_no,op_number) values(o.order_id,start_no+i,lpad(o.op_number::text,3,'0')) returning id into sid;
   pos:=pos+1;insert into public.production_quantity_stations values(sid,op,pos);
  end loop;
 elsif action='join' then
  if not exists(select 1 from public.production_quantity_plans where operation_id=op) then raise exception 'Confirm plan first' using errcode='22023';end if;
 else
  if section is distinct from x.id then raise exception 'Machine section changed; reload before booking' using errcode='PT409';end if;
  select z.* into st from public.production_order_stations z join public.production_quantity_stations q on q.station_id=z.id where z.id=station and q.operation_id=op and z.order_id=o.order_id for update of z;
  if st.id is null or kind is null or kind not in('good','scrap','clarify') or delta is null or delta not in(-1,1) then raise exception 'Invalid count' using errcode='22023';end if;
  select * into pe from public.production_order_employees where id=person and order_id=o.order_id and active for share;
  if pe.id is null or not (pe.employee_id=(select id from public.employees where auth_user_id=auth.uid() and is_active) or article_master_private.routing_manager(x.department_id)) then raise exception 'Own assigned employee or current AL required' using errcode='42501';end if;
  if kind<>'good' and delta>0 and not exists(select 1 from public.production_qa_causes where id=cause and active) then raise exception 'Active 6M cause required' using errcode='22023';end if;
  if kind='good' then
   select coalesce(sum(good_qty),0) into current_qty from public.production_station_counts where station_id=station and order_employee_id=person;
   if current_qty+delta<0 then raise exception 'Negative good count' using errcode='22023';end if;
   if exists(select 1 from public.production_station_counts where station_id=station and order_employee_id=person) then
    update public.production_station_counts set good_qty=good_qty+delta,updated_at=now() where station_id=station and order_employee_id=person;
   else insert into public.production_station_counts(order_id,station_id,order_employee_id,good_qty) values(o.order_id,station,person,delta);end if;
  else
   current_qty:=case when kind='scrap' then st.scrap_total else st.clarify_total end;
   if current_qty+delta<0 then raise exception 'Negative amount' using errcode='22023';end if;
   if kind='scrap' then update public.production_order_stations set scrap_total=scrap_total+delta,scrap_lifetime=scrap_lifetime+greatest(delta,0),updated_at=now() where id=station;
   else update public.production_order_stations set clarify_total=clarify_total+delta,clarify_lifetime=clarify_lifetime+greatest(delta,0),updated_at=now() where id=station;end if;
  end if;
  event_kind:=kind||case when delta<0 then '_correction' else '' end;
  insert into public.production_station_events(order_id,station_id,order_employee_id,employee_id,event_type,qty,qa_cause_id,note) values(o.order_id,station,person,pe.employee_id,event_kind,delta,case when kind='good' then null else cause end,note) returning id into eid;
  insert into public.production_quantity_attributions values(eid,op,x.id,auth.uid());
 end if;
 result:=article_master_private.quantity_row(op);
 insert into public.production_quantity_events(request_id,actor_id,operation_id,body,before_value,after_value) values(request,auth.uid(),op,body,before_value,result);
 -- Re-evaluate the whole BA, including downstream OPs after upstream corrections.
 for row in select q.operation_id,q.station_id,article_master_private.quantity_available(q.station_id) available,article_master_private.quantity_good(q.station_id)+z.scrap_total+z.clarify_total counted
 from public.production_quantity_stations q join public.production_order_stations z on z.id=q.station_id where z.order_id=o.order_id loop
  if row.counted>row.available then
   conf:=null;
   insert into public.production_quantity_conflicts(operation_id,station_id,first_request,initial_available,initial_counted) values(row.operation_id,row.station_id,request,row.available,row.counted) on conflict(station_id) do nothing returning id into conf;
   if conf is not null then
    insert into public.production_quantity_notifications select conf,a.user_id from public.production_routing_authorities a where a.department_id=article_master_private.routing_department(row.operation_id) and article_master_private.routing_manager(a.department_id,a.user_id);
   end if;
  end if;
 end loop;
 return result;
end; $$;
create function article_master_private.quantity_view() returns jsonb language plpgsql stable security definer set search_path='' as $$
begin
 if not article_master_private.quantity_enabled() or not article_master_private.routing_active() then return jsonb_build_object('rows','[]'::jsonb,'notifications','[]'::jsonb);end if;
 return jsonb_build_object('rows',coalesce((select jsonb_agg(article_master_private.quantity_row(o.id) order by o.order_id,o.position) from public.production_order_operations o where article_master_private.routing_visible(o.id)),'[]'::jsonb),
 'notifications',coalesce((select jsonb_agg(to_jsonb(c)) from public.production_quantity_notifications n join public.production_quantity_conflicts c on c.id=n.conflict_id where n.recipient_user_id=auth.uid() and article_master_private.routing_visible(c.operation_id)),'[]'::jsonb));
end; $$;
create function public.op_quantity_view() returns jsonb language sql security invoker set search_path='' as $$select article_master_private.quantity_view()$$;
create function public.op_quantity_command(p_operation_id uuid,p_action text,p_request_id uuid,p_station_id uuid default null,p_order_employee_id uuid default null,p_kind text default null,p_delta integer default null,p_section_id uuid default null,p_qa_cause_id uuid default null,p_note text default null,p_station_ids uuid[] default null,p_new_count integer default null) returns jsonb language sql security invoker set search_path='' as $$select article_master_private.quantity_command(p_operation_id,p_action,p_request_id,p_station_id,p_order_employee_id,p_kind,p_delta,p_section_id,p_qa_cause_id,p_note,p_station_ids,p_new_count)$$;
do $$declare t text;begin
 foreach t in array array['production_quantity_settings','production_quantity_plans','production_quantity_stations','production_quantity_attributions','production_quantity_events','production_quantity_conflicts','production_quantity_notifications'] loop
 execute format('alter table public.%I enable row level security',t);
 execute format('revoke all on public.%I from public,anon,authenticated',t);
 if t<>'production_quantity_settings' then execute format('create trigger quantity_immutable before update or delete on public.%I for each row execute function article_master_private.immutable_snapshot()',t);end if;
 end loop;
end$$;
revoke all on function article_master_private.quantity_enabled(),article_master_private.quantity_good(uuid),article_master_private.quantity_finished(uuid),article_master_private.quantity_available(uuid),article_master_private.quantity_row(uuid),article_master_private.quantity_raw_guard(),article_master_private.quantity_tracked(uuid),article_master_private.quantity_station_tracked(uuid),article_master_private.quantity_command(uuid,text,uuid,uuid,uuid,text,integer,uuid,uuid,text,uuid[],integer),article_master_private.quantity_view(),public.op_quantity_view(),public.op_quantity_command(uuid,text,uuid,uuid,uuid,text,integer,uuid,uuid,text,uuid[],integer) from public,anon,authenticated;
grant execute on function article_master_private.quantity_tracked(uuid),article_master_private.quantity_station_tracked(uuid),article_master_private.quantity_command(uuid,text,uuid,uuid,uuid,text,integer,uuid,uuid,text,uuid[],integer),article_master_private.quantity_view(),public.op_quantity_view(),public.op_quantity_command(uuid,text,uuid,uuid,uuid,text,integer,uuid,uuid,text,uuid[],integer) to authenticated;
-- Narrow adapters to the reviewed readiness API; unlinked legacy behavior is exact.
create function article_master_private.quantity_linked(op uuid) returns boolean language sql stable security definer set search_path='' as $$select exists(select 1 from public.production_quantity_plans where operation_id=op)$$;
create function article_master_private.quantity_op_available(op uuid) returns bigint language sql stable security definer set search_path='' as $$
 select case when o.predecessor is null then b.target_quantity::bigint else article_master_private.quantity_finished((select id from public.production_order_operations where order_id=o.order_id and op_number=o.predecessor)) end from public.production_order_operations o join public.production_orders b on b.id=o.order_id where o.id=op
$$;
do $adapter$
declare definition text;needle text;replacement text;
begin
 definition:=pg_get_functiondef('article_master_private.readiness_command_before_setup(uuid,text,uuid,integer,integer,uuid,timestamptz,text)'::regprocedure);
 needle:='if e.work_status not in(''planned'',''paused'') or e.received_quantity<>0 or p_quantity is distinct from s.quantity then';
 replacement:='if (case when article_master_private.quantity_linked(o.id) then p_quantity is null or p_quantity<=e.received_quantity or p_quantity>article_master_private.quantity_op_available(o.id) else e.work_status not in(''planned'',''paused'') or e.received_quantity<>0 or p_quantity is distinct from s.quantity end) then';
 if position(needle in definition)=0 then raise exception 'Unexpected readiness handover prerequisite';end if;definition:=replace(definition,needle,replacement);
 needle:='if pred.operation_id is null or pred.work_status<>''completed'' or pred.good_quantity<s.quantity then';
 replacement:='if (case when article_master_private.quantity_linked(o.id) then article_master_private.quantity_op_available(o.id)<p_quantity else pred.operation_id is null or pred.work_status<>''completed'' or pred.good_quantity<s.quantity end) then';
 if position(needle in definition)=0 then raise exception 'Unexpected predecessor prerequisite';end if;definition:=replace(definition,needle,replacement);
 needle:='if e.received_quantity<>s.quantity or (p_action=''start'' and e.work_status<>''ready'') or (p_action=''resume'' and e.work_status<>''paused'') then';
 replacement:='if (case when article_master_private.quantity_linked(o.id) then e.received_quantity<=0 else e.received_quantity<>s.quantity end) or (p_action=''start'' and e.work_status<>''ready'') or (p_action=''resume'' and e.work_status<>''paused'') then';
 if position(needle in definition)=0 then raise exception 'Unexpected readiness start prerequisite';end if;definition:=replace(definition,needle,replacement);
 -- Do not reimplement Phase 5 closure or a second manual good-quantity channel.
 needle:='old_state:=to_jsonb(e);';
 replacement:='if article_master_private.quantity_linked(o.id) and (not article_master_private.quantity_enabled() or p_action in(''complete'',''cancel'')) then raise exception ''Phase 4 linked quantities: completion requires separate Phase 5'' using errcode=''42501'';end if; old_state:=to_jsonb(e);';
 if position(needle in definition)=0 then raise exception 'Unexpected readiness audit prerequisite';end if;execute replace(definition,needle,replacement);
 definition:=pg_get_functiondef('article_master_private.readiness_view()'::regprocedure);
 needle:='(o.predecessor is null or exists(select 1 from public.production_order_operations po join public.production_operation_execution pe on pe.operation_id=po.id where po.order_id=o.order_id and po.op_number=o.predecessor and pe.work_status=''completed'' and pe.good_quantity>=s.quantity)) predecessor_released';
 replacement:='(case when article_master_private.quantity_linked(o.id) then article_master_private.quantity_op_available(o.id)>0 else (o.predecessor is null or exists(select 1 from public.production_order_operations po join public.production_operation_execution pe on pe.operation_id=po.id where po.order_id=o.order_id and po.op_number=o.predecessor and pe.work_status=''completed'' and pe.good_quantity>=s.quantity)) end) predecessor_released, article_master_private.quantity_linked(o.id) quantity_linked, article_master_private.quantity_op_available(o.id) quantity_available';
 if position(needle in definition)=0 then raise exception 'Unexpected readiness view prerequisite';end if;definition:=replace(definition,needle,replacement);
 needle:='coalesce(ex.good_quantity,0) good_quantity';
 if position(needle in definition)=0 then raise exception 'Unexpected readiness quantity prerequisite';end if;execute replace(definition,needle,'case when article_master_private.quantity_linked(o.id) then article_master_private.quantity_finished(o.id) else coalesce(ex.good_quantity,0) end good_quantity');
end;$adapter$;
create function article_master_private.quantity_ba_guard() returns trigger language plpgsql security definer set search_path='' as $$
begin
 if article_master_private.quantity_tracked(old.id) and ((new.ba_quantity,new.target_quantity) is distinct from (old.ba_quantity,old.target_quantity) or new.status is distinct from old.status and new.status in('completed','cancelled')) then raise exception 'Linked BA quantity/closure requires Phase 5' using errcode='42501';end if;return new;
end;$$;
create trigger quantity_ba_guard before update of ba_quantity,target_quantity,status on public.production_orders for each row execute function article_master_private.quantity_ba_guard();
revoke all on function article_master_private.quantity_linked(uuid),article_master_private.quantity_op_available(uuid),article_master_private.quantity_ba_guard() from public,anon,authenticated;

commit;
