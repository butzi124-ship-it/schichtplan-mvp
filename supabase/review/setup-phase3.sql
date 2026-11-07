-- REVIEW ONLY. Unregistered proposal after Phase 2; never run against production.
begin;
set local lock_timeout='5s';
create table public.production_setup_settings(singleton boolean primary key default true check(singleton),enabled boolean not null default false);
insert into public.production_setup_settings values(true,false);
create table public.production_setup_runs(
 id uuid primary key default gen_random_uuid(),section_id uuid not null references public.production_operation_machine_sections(id),
 started_at timestamptz not null,ended_at timestamptz,ended_by uuid references auth.users(id),end_reason text,
 check(ended_at is null or ended_at>=started_at)
);
create unique index setup_one_open_run on public.production_setup_runs(section_id) where ended_at is null;
alter table public.production_operation_activity_segments add column setup_run_id uuid references public.production_setup_runs(id);
create table public.production_setup_interruptions(
 user_id uuid primary key references auth.users(id),paused boolean not null default false,last_run_id uuid references public.production_setup_runs(id)
);
-- Corrections overlay immutable original segments; each revision retains full old/new evidence.
create table public.production_setup_corrections(
 id uuid primary key default gen_random_uuid(),revision bigint generated always as identity unique,segment_id uuid not null references public.production_operation_activity_segments(id),
 request_id uuid not null unique,actor_id uuid not null references auth.users(id),happened_at timestamptz not null,
 before_state jsonb not null,after_state jsonb not null,reason text not null check(length(trim(reason)) between 1 and 4000)
);
create table public.production_setup_events(
 request_id uuid primary key,actor_id uuid not null references auth.users(id),happened_at timestamptz not null,
 action text not null,payload jsonb not null,before_state jsonb not null,after_state jsonb not null
);
create trigger setup_immutable_events before update or delete on public.production_setup_events for each row execute function article_master_private.immutable_snapshot();
create trigger setup_immutable_corrections before update or delete on public.production_setup_corrections for each row execute function article_master_private.immutable_snapshot();
create function article_master_private.setup_run_guard() returns trigger language plpgsql set search_path='' as $$
begin
 if tg_op='DELETE' or old.ended_at is not null or (new.id,new.section_id,new.started_at) is distinct from (old.id,old.section_id,old.started_at) then raise exception 'Setup history retained' using errcode='42501';end if;return new;
end;$$;
create trigger setup_run_guard before update or delete on public.production_setup_runs for each row execute function article_master_private.setup_run_guard();
create function article_master_private.setup_enabled() returns boolean language sql stable security definer set search_path='' as $$
 select article_master_private.routing_enabled() and coalesce((select enabled from public.production_setup_settings where singleton),false)
$$;
-- A single transaction lock covers cross-OP personal pauses as well as routing. It is
-- acquired BEFORE BA/section locks in every entrypoint; no opposing employee/order lock order.
-- Conservative serialization is intentional until load testing permits finer locks.
alter function article_master_private.routing_command(uuid,text,uuid,integer,uuid,text,text) rename to routing_command_before_setup;
create function article_master_private.routing_command(op uuid,action text,request uuid,expected integer,target uuid,reason text,abort_code text)
returns jsonb language plpgsql security definer set search_path='' as $$
begin
 perform pg_advisory_xact_lock(73103,1);
 return article_master_private.routing_command_before_setup(op,action,request,expected,target,reason,abort_code);
end;$$;
-- Include paused but still open shared runs in the existing configured-abort check and view.
do $adapter$
declare definition text;needle text:='a.ended_at is null and a.activity';
begin
 definition:=pg_get_functiondef('article_master_private.routing_command_before_setup(uuid,text,uuid,integer,uuid,text,text)'::regprocedure);
 if position(needle in definition)=0 then raise exception 'Unexpected routing prerequisite';end if;
 execute replace(definition,needle,'(a.ended_at is null or exists(select 1 from public.production_setup_runs sr where sr.section_id=x.id and sr.ended_at is null)) and a.activity');
 definition:=pg_get_functiondef('article_master_private.routing_view()'::regprocedure);
 if position(needle in definition)=0 then raise exception 'Unexpected routing view prerequisite';end if;
 execute replace(definition,needle,'(a.ended_at is null or exists(select 1 from public.production_setup_runs sr where sr.section_id=x.id and sr.ended_at is null)) and a.activity');
end;$adapter$;
-- Routing already ends personal segments at this exact server timestamp. Close its
-- shared run atomically, including paused runs, without touching a second OP.
create function article_master_private.setup_section_closed() returns trigger language plpgsql security definer set search_path='' as $$
declare old_runs jsonb;new_runs jsonb;
begin
 select coalesce(jsonb_agg(to_jsonb(r)),'[]') into old_runs from public.production_setup_runs r where section_id=new.id and ended_at is null;
 with closed as(update public.production_setup_runs set ended_at=new.closed_at,ended_by=new.closed_by,end_reason='machine_change' where section_id=new.id and ended_at is null returning *) select coalesce(jsonb_agg(to_jsonb(closed)),'[]') into new_runs from closed;
 if old_runs<>'[]'::jsonb then insert into public.production_setup_events values(gen_random_uuid(),auth.uid(),new.closed_at,'machine_change',jsonb_build_object('section_id',new.id,'reason',new.close_reason),old_runs,new_runs);end if;
 return new;
end;$$;
create trigger setup_section_closed after update of closed_at on public.production_operation_machine_sections for each row when(old.closed_at is null and new.closed_at is not null) execute function article_master_private.setup_section_closed();
create function article_master_private.setup_effective() returns table(id uuid,section_id uuid,user_id uuid,activity text,started_at timestamptz,ended_at timestamptz,setup_run_id uuid) language sql stable security definer set search_path='' as $$
 select a.id,a.section_id,a.user_id,coalesce(c.after_state->>'activity',a.activity),coalesce((c.after_state->>'started_at')::timestamptz,a.started_at),coalesce((c.after_state->>'ended_at')::timestamptz,a.ended_at),a.setup_run_id
 from public.production_operation_activity_segments a left join lateral(select after_state from public.production_setup_corrections where segment_id=a.id order by revision desc limit 1)c on true
$$;
create function article_master_private.setup_command(action text,request uuid,op uuid default null,run uuid default null,mode text default null,segment uuid default null,expected jsonb default null,change jsonb default null,reason text default null,previous_user uuid default null)
returns jsonb language plpgsql security definer set search_path='' as $$
declare stamp timestamptz;r public.production_setup_runs;x public.production_operation_machine_sections;booked public.production_operation_activity_segments;
 prior public.production_setup_events;body jsonb;before_data jsonb:='[]';after_data jsonb:='[]';target_user uuid;item record;effective record;start_value timestamptz;end_value timestamptz;activity_value text;oid uuid;
begin
 if not article_master_private.setup_enabled() or not article_master_private.routing_active() then raise exception 'Denied' using errcode='42501';end if;
 if request is null or action is null or action not in('start','pause_own','end_own','finish','pause','return','logout','takeover','correct') then raise exception 'Invalid setup action' using errcode='22023';end if;
 perform pg_advisory_xact_lock(73103,1);
 -- Freeze independently provisioned authority while checking and committing.
 perform 1 from public.production_setup_settings where singleton for share;
 perform 1 from public.production_routing_settings where singleton for share;
 perform 1 from public.departments order by id for share;
 perform 1 from public.employees order by id for share;
 perform 1 from public.production_routing_authorities order by department_id,user_id for share;
 perform 1 from public.production_readiness_grants order by user_id,machine_id for share;
 if not article_master_private.setup_enabled() or not article_master_private.routing_active() then raise exception 'Authorization changed' using errcode='42501';end if;
 body:=jsonb_build_object('op',op,'run',run,'mode',mode,'segment',segment,'expected',expected,'change',change,'reason',reason,'previous_user',previous_user);
 select * into prior from public.production_setup_events where request_id=request;
 if found then
  if prior.actor_id<>auth.uid() or prior.action<>action or prior.payload<>body then raise exception 'Request reused' using errcode='22023';end if;
  return prior.after_state;
 end if;
 -- Matching routing's order-first lock discipline; pause may span several OPs.
 perform 1 from public.production_orders order by id for update;
 stamp:=clock_timestamp();
 if action in('pause','logout','return') then
  target_user:=auth.uid();
 elsif action='takeover' then
  select * into x from public.production_operation_machine_sections where operation_id=op and closed_at is null for update;
  if x.id is null or not(article_master_private.routing_operator(op,x.machine_id) or article_master_private.routing_manager(x.department_id)) or previous_user is null or previous_user=auth.uid() then raise exception 'Takeover denied' using errcode='42501';end if;
  -- Explicit predecessor from this workspace only; cannot stop another OP.
  if not exists(select 1 from public.production_operation_activity_segments where user_id=previous_user and section_id=x.id and ended_at is null) then raise exception 'Predecessor already stopped' using errcode='PT409';end if;
  target_user:=previous_user;
 else
  if action='correct' then select * into booked from public.production_operation_activity_segments where id=segment;run:=booked.setup_run_id;end if;
  if run is not null then
   select * into r from public.production_setup_runs where id=run for update;
   select * into x from public.production_operation_machine_sections where id=r.section_id for update;
   if r.id is null or (op is not null and op<>x.operation_id) then raise exception 'Wrong run' using errcode='22023';end if;
   op:=x.operation_id;
  else
   select * into x from public.production_operation_machine_sections where operation_id=op and closed_at is null for update;
   select * into r from public.production_setup_runs where section_id=x.id and ended_at is null for update;
  end if;
  if x.id is null then raise exception 'Select machine explicitly first' using errcode='22023';end if;
  if action='correct' then
   if booked.user_id is distinct from auth.uid() and not article_master_private.routing_manager(article_master_private.routing_department(op)) then raise exception 'Correction denied' using errcode='42501';end if;
  elsif action='finish' then
   if not article_master_private.routing_manager(x.department_id) and not exists(select 1 from public.production_operation_activity_segments where setup_run_id=r.id and user_id=auth.uid()) then raise exception 'Participant required' using errcode='42501';end if;
  elsif not(article_master_private.routing_operator(op,x.machine_id) or article_master_private.routing_manager(x.department_id)) then raise exception 'Operator denied' using errcode='42501';end if;
 end if;
 if action in('pause','logout','takeover') then
  select coalesce(jsonb_agg(to_jsonb(s)),'[]') into before_data from public.production_operation_activity_segments s where s.user_id=target_user and s.ended_at is null and (action<>'takeover' or s.section_id=x.id);
  select s.setup_run_id into run from public.production_operation_activity_segments s where s.user_id=target_user and s.setup_run_id is not null and (action<>'takeover' or s.section_id=x.id) order by s.started_at desc limit 1;
  update public.production_operation_activity_segments set ended_at=stamp,end_reason=action where user_id=target_user and ended_at is null and (action<>'takeover' or section_id=x.id);
  insert into public.production_setup_interruptions values(target_user,action='pause',run) on conflict(user_id) do update set paused=excluded.paused,last_run_id=coalesce(excluded.last_run_id,public.production_setup_interruptions.last_run_id);
  after_data:=jsonb_build_object('interrupted',before_data,'ended_at',stamp,'resume_run',run);
 elsif action='return' then
  update public.production_setup_interruptions set paused=false where user_id=auth.uid();
  after_data:=jsonb_build_object('resumed',false);
 elsif action='correct' then
  if booked.ended_at is null or booked.setup_run_id is null or reason is null or length(trim(reason)) not between 1 and 4000 then raise exception 'Closed segment and reason required' using errcode='22023';end if;
  select * into effective from article_master_private.setup_effective() where id=segment;
  before_data:=to_jsonb(effective);
  if expected is distinct from before_data then raise exception 'Stale correction' using errcode='PT409';end if;
  start_value:=(change->>'started_at')::timestamptz;end_value:=(change->>'ended_at')::timestamptz;activity_value:=change->>'activity';
  if start_value is null or end_value is null or not isfinite(start_value) or not isfinite(end_value) or start_value<r.started_at or end_value>coalesce(r.ended_at,stamp) or end_value<start_value or activity_value is null or activity_value not in('setup_internal','setup_external') then raise exception 'Invalid corrected interval' using errcode='22023';end if;
  if exists(select 1 from article_master_private.setup_effective() e where e.user_id=booked.user_id and e.id<>booked.id and e.started_at<end_value and coalesce(e.ended_at,'infinity')>start_value) then raise exception 'Personal overlap' using errcode='22023';end if;
  after_data:=before_data||jsonb_build_object('started_at',start_value,'ended_at',end_value,'activity',activity_value);
  insert into public.production_setup_corrections(segment_id,request_id,actor_id,happened_at,before_state,after_state,reason) values(segment,request,auth.uid(),stamp,before_data,after_data,trim(reason));
 else
  if x.closed_at is not null or r.ended_at is not null or (action='start' and exists(select 1 from public.production_orders b join public.production_order_operations o on o.order_id=b.id left join public.production_operation_execution e on e.operation_id=o.id where o.id=op and (b.status not in('running','paused') or e.work_status in('completed','cancelled')))) then raise exception 'Run/OP closed' using errcode='PT409';end if;
  before_data:=jsonb_build_object('run',to_jsonb(r),'segments',coalesce((select jsonb_agg(to_jsonb(s)) from public.production_operation_activity_segments s where s.setup_run_id=r.id),'[]'::jsonb));
  if action<>'start' and r.id is null then raise exception 'No open setup' using errcode='PT409';end if;
  if action='start' then
   if mode is null or mode not in('internal','external') or exists(select 1 from public.production_setup_interruptions where user_id=auth.uid() and paused) then raise exception 'Select mode / return from pause first' using errcode='22023';end if;
   if mode='internal' and (exists(select 1 from public.production_operation_execution where machine_id=x.machine_id and work_status='running') or exists(select 1 from public.production_operation_activity_segments a join public.production_operation_machine_sections z on z.id=a.section_id where z.machine_id=x.machine_id and z.id<>x.id and a.ended_at is null and a.activity='setup_internal')) then raise exception 'Machine already occupied' using errcode='PT409';end if;
   -- An explicit mode change closes only this participant's old segment.
   if exists(select 1 from public.production_operation_activity_segments where user_id=auth.uid() and ended_at is null and setup_run_id is distinct from r.id) then raise exception 'Another personal activity running' using errcode='PT409';end if;
   if r.id is null then insert into public.production_setup_runs(section_id,started_at) values(x.id,stamp) returning * into r;end if;
   update public.production_operation_activity_segments set ended_at=stamp,end_reason='mode_change' where setup_run_id=r.id and user_id=auth.uid() and ended_at is null;
   insert into public.production_operation_activity_segments(section_id,user_id,activity,started_at,setup_run_id) values(x.id,auth.uid(),'setup_'||mode,stamp,r.id);
   insert into public.production_setup_interruptions values(auth.uid(),false,r.id) on conflict(user_id) do update set last_run_id=excluded.last_run_id;
  elsif action in('pause_own','end_own') then
   update public.production_operation_activity_segments set ended_at=stamp,end_reason=action where setup_run_id=r.id and user_id=auth.uid() and ended_at is null;
  elsif action='finish' then
   update public.production_operation_activity_segments set ended_at=stamp,end_reason='shared_finish' where setup_run_id=r.id and ended_at is null;
   update public.production_setup_runs set ended_at=stamp,ended_by=auth.uid(),end_reason='shared_finish' where id=r.id and ended_at is null;
  end if;
  after_data:=jsonb_build_object('run_id',r.id,'ended_at',case when action='finish' then stamp end);
 end if;
 insert into public.production_setup_events values(request,auth.uid(),stamp,action,body,before_data,after_data);
 return after_data;
end;$$;
create function article_master_private.setup_minutes(rid uuid,internal_only boolean default false) returns numeric language sql stable security definer set search_path='' as $$
 select coalesce(sum(extract(epoch from(upper(v)-lower(v)))/60),0) from unnest((select range_agg(tstzrange(started_at,coalesce(ended_at,statement_timestamp()),'[)')) from article_master_private.setup_effective() where setup_run_id=rid and (not internal_only or activity='setup_internal'))) v
$$;
create function article_master_private.setup_view() returns jsonb language plpgsql stable security definer set search_path='' as $$
declare rows jsonb;last_run uuid;
begin
 if not article_master_private.setup_enabled() or not article_master_private.routing_active() then return jsonb_build_object('rows','[]'::jsonb);end if;
 select last_run_id into last_run from public.production_setup_interruptions where user_id=auth.uid();
 select coalesce(jsonb_agg(to_jsonb(result) order by ba_number,op_number),'[]') into rows from(
 select x.closed_at,x.operation_id,x.id section_id,x.machine_id,m.name machine_name,b.ba_number,o.op_number,o.data->'tr' target_setup,
 (x.closed_at is null and b.status in('running','paused') and not exists(select 1 from public.production_operation_execution e where e.operation_id=o.id and e.work_status in('completed','cancelled')) and (article_master_private.routing_operator(o.id,x.machine_id) or article_master_private.routing_manager(x.department_id))) can_start,
 article_master_private.routing_manager(x.department_id) can_manage,
 coalesce((select jsonb_agg(to_jsonb(z) order by started_at) from(
 select r.*,
 (article_master_private.routing_manager(x.department_id) or exists(select 1 from public.production_operation_activity_segments a where a.setup_run_id=r.id and a.user_id=auth.uid())) can_finish,
 coalesce((select sum(extract(epoch from(coalesce(a.ended_at,statement_timestamp())-a.started_at))/60) from article_master_private.setup_effective() a where a.setup_run_id=r.id),0) person_minutes,
 extract(epoch from(coalesce(r.ended_at,statement_timestamp())-r.started_at))/60 calendar_minutes,
 article_master_private.setup_minutes(r.id) active_calendar_minutes,article_master_private.setup_minutes(r.id,true) machine_stop_minutes,
 coalesce((select jsonb_agg(to_jsonb(a)||jsonb_build_object('can_correct',a.ended_at is not null and (a.user_id=auth.uid() or article_master_private.routing_manager(article_master_private.routing_department(o.id))),'correction_count',(select count(*) from public.production_setup_corrections c where c.segment_id=a.id)) order by a.started_at) from article_master_private.setup_effective() a where a.setup_run_id=r.id),'[]') segments,
 coalesce((select jsonb_agg(to_jsonb(c) order by happened_at) from public.production_setup_corrections c join public.production_operation_activity_segments a on a.id=c.segment_id where a.setup_run_id=r.id and(a.user_id=auth.uid() or article_master_private.routing_manager(article_master_private.routing_department(o.id)))),'[]') corrections
 from public.production_setup_runs r where r.section_id=x.id)z),'[]') runs,
 coalesce((select jsonb_agg(jsonb_build_object('user_id',a.user_id,'display_name',coalesce((select coalesce(to_jsonb(e)->>'display_name',to_jsonb(e)->>'name') from public.employees e where e.auth_user_id=a.user_id limit 1),a.user_id::text))) from public.production_operation_activity_segments a where a.section_id=x.id and a.ended_at is null and a.user_id<>auth.uid()),'[]') predecessors
 from public.production_operation_machine_sections x join public.production_order_operations o on o.id=x.operation_id
 join public.production_orders b on b.id=o.order_id join public.production_machines m on m.id=x.machine_id
 where article_master_private.routing_visible(o.id) or exists(select 1 from public.production_operation_activity_segments own where own.section_id=x.id and own.user_id=auth.uid() and own.setup_run_id is not null)
 )result;
 return jsonb_build_object('rows',rows,'paused',coalesce((select paused from public.production_setup_interruptions where user_id=auth.uid()),false),
 'resume_run',case when exists(select 1 from public.production_setup_runs r join public.production_operation_machine_sections x on x.id=r.section_id where r.id=last_run and r.ended_at is null and x.closed_at is null and(article_master_private.routing_operator(x.operation_id,x.machine_id) or article_master_private.routing_manager(x.department_id))) then last_run end);
end;$$;
create function public.setup_time_view() returns jsonb language sql stable security invoker set search_path='' as $$select article_master_private.setup_view()$$;
create function public.setup_time_command(p_action text,p_request_id uuid,p_operation_id uuid default null,p_run_id uuid default null,p_mode text default null,p_segment_id uuid default null,p_expected jsonb default null,p_change jsonb default null,p_reason text default null,p_previous_user uuid default null)
returns jsonb language sql security invoker set search_path='' as $$select article_master_private.setup_command(p_action,p_request_id,p_operation_id,p_run_id,p_mode,p_segment_id,p_expected,p_change,p_reason,p_previous_user)$$;
-- Existing machine production starts use the same lock and cannot claim an internally
-- occupied machine. External setup never marks execution.work_status as stopped.
alter function article_master_private.readiness_command(uuid,text,uuid,integer,integer,uuid,timestamptz,text) rename to readiness_command_before_setup;
create function article_master_private.readiness_command(op uuid,action text,request uuid,expected integer,quantity integer,target uuid,planned timestamptz,reason text)
returns jsonb language plpgsql security definer set search_path='' as $$
declare machine uuid;
begin
 perform pg_advisory_xact_lock(73103,1);
 select machine_id into machine from public.production_operation_execution where operation_id=op;
 if action in('start','resume') and exists(select 1 from public.production_operation_activity_segments a join public.production_operation_machine_sections x on x.id=a.section_id where x.machine_id=machine and a.ended_at is null and a.activity='setup_internal') then raise exception 'Machine internally occupied' using errcode='PT409';end if;
 return article_master_private.readiness_command_before_setup(op,action,request,expected,quantity,target,planned,reason);
end;$$;
-- Rebind public wrappers explicitly after adapting their private implementations.
create or replace function public.op_routing_command(p_operation_id uuid,p_action text,p_request_id uuid,p_expected_version integer,p_machine_id uuid,p_reason text,p_abort_code text default null)
returns jsonb language sql security invoker set search_path='' as $$select article_master_private.routing_command(p_operation_id,p_action,p_request_id,p_expected_version,p_machine_id,p_reason,p_abort_code)$$;
create or replace function public.machine_readiness_command(p_operation_id uuid,p_action text,p_request_id uuid,p_expected_version integer,p_quantity integer default null,p_machine_id uuid default null,p_planned_start timestamptz default null,p_reason text default null)
returns jsonb language sql security invoker set search_path='' as $$select article_master_private.readiness_command(p_operation_id,p_action,p_request_id,p_expected_version,p_quantity,p_machine_id,p_planned_start,p_reason)$$;
do $permissions$
declare t text;f record;
begin
 foreach t in array array['production_setup_settings','production_setup_runs','production_setup_interruptions','production_setup_corrections','production_setup_events'] loop
 execute format('alter table public.%I enable row level security',t);execute format('revoke all on public.%I from public,anon,authenticated',t);end loop;
 for f in select p.oid::regprocedure signature from pg_proc p join pg_namespace n on n.oid=p.pronamespace where n.nspname='article_master_private' and(p.proname like 'setup_%' or p.proname in('routing_command_before_setup','readiness_command_before_setup','routing_command','readiness_command')) loop
 execute format('revoke all on function %s from public,anon,authenticated',f.signature);end loop;
end;$permissions$;
revoke all on function public.setup_time_view(),public.setup_time_command(text,uuid,uuid,uuid,text,uuid,jsonb,jsonb,text,uuid) from public,anon,authenticated;
grant execute on function public.setup_time_view(),public.setup_time_command(text,uuid,uuid,uuid,text,uuid,jsonb,jsonb,text,uuid),article_master_private.setup_view(),article_master_private.setup_command(text,uuid,uuid,uuid,text,uuid,jsonb,jsonb,text,uuid),article_master_private.routing_command(uuid,text,uuid,integer,uuid,text,text),article_master_private.readiness_command(uuid,text,uuid,integer,integer,uuid,timestamptz,text) to authenticated;
commit;
