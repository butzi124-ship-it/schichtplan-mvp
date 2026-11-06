-- REVIEW ONLY, unregistered. Apply reviewed #7, #8 and #9 first. No production application.
begin;
create table public.production_forecast_settings(
 department_id uuid primary key references public.departments(id),enabled boolean not null default false,
 zone text not null check(zone in('Europe/Berlin','Europe/Zurich')),start_date date not null,end_date date not null check(end_date>=start_date and end_date-start_date<=62),
 valid_until timestamptz not null,green_buffer_seconds integer not null check(green_buffer_seconds>=0),orange_buffer_seconds integer not null check(orange_buffer_seconds>=0 and orange_buffer_seconds<=green_buffer_seconds),
 max_age_seconds integer not null check(max_age_seconds between 1 and 3600),feedback_max_age_seconds integer not null check(feedback_max_age_seconds>0),dst_disambiguation text not null check(dst_disambiguation in('earlier','later','reject')),
 calculation_actor_id uuid not null references auth.users(id),staff_salt uuid not null default gen_random_uuid(),roster_complete boolean not null default false
);
create table public.production_forecast_machines(
 machine_id uuid primary key references public.production_machines(id),capacity_units integer not null check(capacity_units between 1 and 16),unmanned_capable boolean not null default false
);
create table public.production_forecast_shifts(
 id uuid primary key default gen_random_uuid(),machine_id uuid not null references public.production_forecast_machines(machine_id),shift_definition_id uuid not null references public.shift_definitions(id),
 planner_suffix text not null check(planner_suffix ~ '^-[a-z]{2}-[0-9]+$'),weekdays integer[] not null check(cardinality(weekdays)>0 and weekdays <@ array[1,2,3,4,5,6,7]),breaks jsonb not null check(jsonb_typeof(breaks)='array'),unmanned boolean not null default false,
 unique(machine_id,planner_suffix)
);
create table public.production_forecast_staff(
 employee_id uuid primary key references public.employees(id),planner_user text not null unique check(length(planner_user) between 1 and 200),qualifications text[] not null,machine_ids uuid[] not null
);
create table public.production_forecast_rules(
 operation_id uuid primary key references public.production_order_operations(id),priority integer not null,qualification text not null check(length(qualification) between 1 and 100),people integer not null check(people between 1 and 16),unmanned_allowed boolean not null default false,
 transfer_seconds integer check(transfer_seconds>=0),quantity_source text not null check(quantity_source in('reported','stations')),station_binding_verified boolean not null default false,quantity_station_id uuid references public.production_order_stations(id),station_times_are_processing boolean not null default false
);
create table public.production_forecast_actuals(
 operation_id uuid primary key references public.production_order_operations(id),setup_complete boolean,setup_elapsed_s numeric check(setup_elapsed_s>=0),setup_remaining_s numeric check(setup_remaining_s>=0),processing_elapsed_s numeric check(processing_elapsed_s>=0),current_piece_elapsed_s numeric check(current_piece_elapsed_s>=0),
 good integer check(good>=0),scrap integer check(scrap>=0),clarify integer check(clarify>=0),replacement_qty integer check(replacement_qty>=0),material_available_at timestamptz,wait_until timestamptz,
 observed_at timestamptz not null,reported_by uuid not null references auth.users(id),evidence text not null check(length(evidence) between 1 and 4000)
);
create table public.production_forecast_blocks(
 id uuid primary key default gen_random_uuid(),machine_id uuid not null references public.production_machines(id),starts_at timestamptz not null,ends_at timestamptz check(ends_at>starts_at),units integer not null check(units between 1 and 16),legacy_order_id uuid references public.production_orders(id),evidence text not null check(length(evidence) between 1 and 4000)
);
create table public.production_forecast_queue(department_id uuid primary key references public.departments(id),dirty_at timestamptz not null default now());
create table public.production_forecast_runs(
 id uuid primary key,department_id uuid not null references public.departments(id),actor_id uuid not null references auth.users(id),input_token text not null,calculated_at timestamptz not null,expires_at timestamptz not null,result jsonb not null
);
create index forecast_runs_department_time on public.production_forecast_runs(department_id,calculated_at);
create table public.production_forecast_latest(department_id uuid primary key references public.departments(id),run_id uuid not null references public.production_forecast_runs(id));
alter table public.production_operation_arrivals add column forecast_run_id uuid references public.production_forecast_runs(id);
create index forecast_arrivals_run on public.production_operation_arrivals(forecast_run_id) where forecast_run_id is not null;

create function article_master_private.forecast_authorized(dept uuid,actor uuid) returns boolean language sql stable security definer set search_path='' as $$
 select actor is not null and article_master_private.readiness_enabled() and exists(select 1 from public.production_forecast_settings f join public.departments d on d.id=f.department_id join public.article_master_scopes a on a.department_id=f.department_id where f.department_id=dept and f.enabled and d.active and a.user_id=actor)
$$;
create function article_master_private.forecast_payload(dept uuid) returns jsonb language plpgsql stable security definer set search_path='' as $$
declare cfg public.production_forecast_settings;result jsonb;
begin
 select * into cfg from public.production_forecast_settings where department_id=dept;
 if cfg.department_id is null then return '{}'::jsonb;end if;
 with days as(select d::date work_date from generate_series(cfg.start_date::timestamp,cfg.end_date::timestamp,interval '1 day') d),
 rules as(select ms.*,sd.start_time,sd.end_time,sd.active from public.production_forecast_shifts ms join public.production_machines m on m.id=ms.machine_id join public.shift_definitions sd on sd.id=ms.shift_definition_id where m.department_id=dept),
 roster as(select distinct d.work_date,r.planner_suffix,case when rep.mode='cancel' or can.shift_id is not null then null when rep.mode='replace' then rep.replacement_user else pa.assigned_user end assigned
 from days d cross join rules r left join public.planner_assignments pa on pa.shift_id=d.work_date::text||r.planner_suffix and pa.shift_date=d.work_date
 left join public.planner_absence_replacements rep on rep.shift_id=d.work_date::text||r.planner_suffix and rep.shift_date=d.work_date
 left join public.planner_shift_cancellations can on can.shift_id=d.work_date::text||r.planner_suffix and can.shift_date=d.work_date where extract(isodow from d.work_date)::integer=any(r.weekdays)),
 staffs as(select s.* from public.production_forecast_staff s join public.employees e on e.id=s.employee_id where e.department_id=dept and e.is_active)
 select jsonb_build_object('schema',1,'department_id',dept,'config',jsonb_build_object('zone',cfg.zone,'start_date',cfg.start_date,'end_date',cfg.end_date,'valid_until',cfg.valid_until,'green_buffer_seconds',cfg.green_buffer_seconds,'orange_buffer_seconds',cfg.orange_buffer_seconds,'max_age_seconds',cfg.max_age_seconds,'feedback_max_age_seconds',cfg.feedback_max_age_seconds,'dst_disambiguation',cfg.dst_disambiguation),
 'machines',coalesce((select jsonb_agg(jsonb_build_object('id',m.id,'active',m.active,'capacity_units',f.capacity_units,'unmanned_capable',f.unmanned_capable) order by m.id) from public.production_machines m left join public.production_forecast_machines f on f.machine_id=m.id where m.department_id=dept),'[]'::jsonb),
 'shifts',coalesce((select jsonb_agg(jsonb_build_object('machine_id',r.machine_id,'weekdays',r.weekdays,'suffix',r.planner_suffix,'start_time',r.start_time,'end_time',r.end_time,'breaks',r.breaks,'unmanned',r.unmanned,'excluded_dates',coalesce((select jsonb_agg(c.shift_date) from public.planner_shift_cancellations c where c.shift_id=c.shift_date::text||r.planner_suffix and c.shift_date between cfg.start_date and cfg.end_date),'[]'::jsonb)) order by r.machine_id,r.planner_suffix) from rules r where r.active),'[]'::jsonb),
 'closed_dates',coalesce((select jsonb_agg(day_date order by day_date) from public.planner_special_days where day_date between cfg.start_date and cfg.end_date and type in('holiday','bridge','company')),'[]'::jsonb),
 'staff',coalesce((select jsonb_agg(jsonb_build_object('id',md5(s.employee_id::text||cfg.staff_salt::text),'qualifications',s.qualifications,'machine_ids',s.machine_ids,
 'assignments',coalesce((select jsonb_agg(jsonb_build_object('date',r.work_date,'suffix',r.planner_suffix) order by r.work_date,r.planner_suffix) from roster r where r.assigned=s.planner_user),'[]'::jsonb),
 'unavailable_dates',coalesce((select jsonb_agg(d.work_date order by d.work_date) from days d where exists(select 1 from public.planner_absences a where a.user_name=s.planner_user and a.absence_date=d.work_date) or exists(select 1 from public.planner_vacations v where v.user_name=s.planner_user and d.work_date between v.from_date and v.to_date) or exists(select 1 from public.planner_sick_leaves v where v.user_name=s.planner_user and d.work_date between v.from_date and v.to_date)),'[]'::jsonb),
 'unavailable_shifts',coalesce((select jsonb_agg(a.shift_id order by a.shift_id) from public.planner_availability a where a.user_name=s.planner_user and a.status='no' and a.shift_date between cfg.start_date and cfg.end_date),'[]'::jsonb)) order by s.employee_id) from staffs s),'[]'::jsonb),
 'blocks',coalesce((select jsonb_agg(jsonb_build_object('machine_id',b.machine_id,'start',b.starts_at,'end',coalesce(b.ends_at,((cfg.end_date+1)::timestamp at time zone cfg.zone)),'units',b.units) order by b.starts_at,b.id) from public.production_forecast_blocks b join public.production_machines m on m.id=b.machine_id where m.department_id=dept),'[]'::jsonb),
 'jobs',coalesce((select jsonb_agg(jsonb_build_object('id',o.id,'order_id',o.order_id,'ba_number',s.ba_number,'op_number',o.op_number,'quantity',s.quantity,'tr',o.data->'tr','te',o.data->'te',
 'machine_id',coalesce(e.machine_id,(o.data->>'machine_id')::uuid),'predecessor_id',(select po.id from public.production_order_operations po where po.order_id=o.order_id and po.op_number=o.predecessor),
 'work_status',coalesce(e.work_status,'planned'),'order_status',b.status,'material_ready',coalesce(e.received_quantity,0)>=s.quantity,'completed_at',(select min(happened_at) from public.production_readiness_events where operation_id=o.id and action='complete'),
 'priority',r.priority,'qualification',r.qualification,'people',r.people,'unmanned_allowed',r.unmanned_allowed,'transfer_seconds',r.transfer_seconds,'planned_start',e.planned_start,'wait_until',a.wait_until,
 'due_at',((s.due_date+1)::timestamp at time zone cfg.zone),'actual',jsonb_build_object('setup_complete',a.setup_complete,'setup_elapsed_s',a.setup_elapsed_s,'setup_remaining_s',a.setup_remaining_s,
 'processing_elapsed_s',case when r.quantity_source='stations' and r.station_times_are_processing then (select case when count(*)>0 and count(st.actual_time_minutes)=count(*) then sum(st.actual_time_minutes)*60 else null end from public.production_order_stations st where st.order_id=o.order_id and st.id=r.quantity_station_id and st.op_number ~ '^[0-9]{1,3}$' and lpad(st.op_number,3,'0')=lpad(o.op_number::text,3,'0')) else a.processing_elapsed_s end,
 'current_piece_elapsed_s',a.current_piece_elapsed_s,
 'good',case when r.quantity_source='stations' and r.station_binding_verified and exists(select 1 from public.production_order_stations qs where qs.id=r.quantity_station_id and qs.order_id=o.order_id and qs.op_number ~ '^[0-9]{1,3}$' and lpad(qs.op_number,3,'0')=lpad(o.op_number::text,3,'0')) then (select coalesce(sum(pc.good_qty),0) from public.production_station_counts pc join public.production_order_stations st on st.id=pc.station_id where st.order_id=o.order_id and st.id=r.quantity_station_id and st.op_number ~ '^[0-9]{1,3}$' and lpad(st.op_number,3,'0')=lpad(o.op_number::text,3,'0')) when r.quantity_source='reported' then a.good else null end,
 'scrap',case when r.quantity_source='stations' and r.station_binding_verified and exists(select 1 from public.production_order_stations qs where qs.id=r.quantity_station_id and qs.order_id=o.order_id and qs.op_number ~ '^[0-9]{1,3}$' and lpad(qs.op_number,3,'0')=lpad(o.op_number::text,3,'0')) then (select coalesce(sum(st.scrap_total),0) from public.production_order_stations st where st.order_id=o.order_id and st.id=r.quantity_station_id and st.op_number ~ '^[0-9]{1,3}$' and lpad(st.op_number,3,'0')=lpad(o.op_number::text,3,'0')) when r.quantity_source='reported' then a.scrap else null end,
 'clarify',case when r.quantity_source='stations' and r.station_binding_verified and exists(select 1 from public.production_order_stations qs where qs.id=r.quantity_station_id and qs.order_id=o.order_id and qs.op_number ~ '^[0-9]{1,3}$' and lpad(qs.op_number,3,'0')=lpad(o.op_number::text,3,'0')) then (select coalesce(sum(st.clarify_total),0) from public.production_order_stations st where st.order_id=o.order_id and st.id=r.quantity_station_id and st.op_number ~ '^[0-9]{1,3}$' and lpad(st.op_number,3,'0')=lpad(o.op_number::text,3,'0')) when r.quantity_source='reported' then a.clarify else null end,
 'replacement_qty',a.replacement_qty,'material_available_at',a.material_available_at,'observed_at',a.observed_at)) order by s.ba_number,o.position) from public.production_order_operations o join public.production_order_article_snapshots s on s.order_id=o.order_id join public.production_orders b on b.id=o.order_id left join public.production_operation_execution e on e.operation_id=o.id left join public.production_forecast_rules r on r.operation_id=o.id left join public.production_forecast_actuals a on a.operation_id=o.id where s.department_id=dept and b.status in('running','paused')),'[]'::jsonb),
 'issues',coalesce((select jsonb_agg(i) from(
 select null::uuid machine_id,'Besetzungs-/Kalenderabdeckung nicht bestätigt'::text reason where not cfg.roster_complete
 union all select b.machine_id,'Aktiver Bestandsauftrag ohne dokumentierte Kapazitätsbindung' from public.production_orders b where b.department_id=dept and b.status in('running','paused') and not exists(select 1 from public.production_order_article_snapshots s where s.order_id=b.id) and not exists(select 1 from public.production_forecast_blocks k where k.legacy_order_id=b.id)
 union all select null,'Schichttausch muss vor Prognose in der Besetzung materialisiert und geprüft werden' where exists(select 1 from public.planner_swaps p join staffs s on s.planner_user in(p.user_a,p.user_b) where p.start_date<=cfg.end_date and coalesce(p.end_date,cfg.end_date)>=cfg.start_date)
 union all select null,'Mehrdeutige Personalzuordnung' where exists(select 1 from staffs s where (select count(*) from public.employees e where e.display_name=s.planner_user and e.is_active)<>1)
 union all select null,'Ungültige Verfügbarkeitsangabe' where exists(select 1 from public.planner_availability a join staffs s on s.planner_user=a.user_name where a.shift_date between cfg.start_date and cfg.end_date and a.status not in('yes','no'))
 ) i),'[]'::jsonb)) into result;
 return result;
end; $$;
create function article_master_private.forecast_input(p_department_id uuid,p_request_id uuid) returns jsonb language plpgsql security definer set search_path='' as $$
declare data jsonb;prior public.production_forecast_runs;
begin
 if not coalesce(article_master_private.forecast_authorized(p_department_id,auth.uid()),false) then raise exception 'Denied' using errcode='42501';end if;
 select * into prior from public.production_forecast_runs where id=p_request_id;
 if found then if prior.actor_id<>auth.uid() or prior.department_id<>p_department_id then raise exception 'Request reused' using errcode='22023';end if;return jsonb_build_object('existing_run',prior.id);end if;
 data:=article_master_private.forecast_payload(p_department_id);
 return data||jsonb_build_object('input_token',md5(data::text),'as_of',clock_timestamp());
end; $$;
create function article_master_private.forecast_source_valid(run uuid) returns boolean language sql stable security definer set search_path='' as $$
 select run is null or exists(select 1 from public.production_forecast_runs r join public.production_forecast_latest l on l.run_id=r.id join public.production_forecast_settings s on s.department_id=r.department_id join public.departments d on d.id=s.department_id where r.id=run and article_master_private.forecast_authorized(r.department_id,r.actor_id) and s.enabled and d.active and article_master_private.readiness_enabled() and r.expires_at>now() and r.input_token=md5(article_master_private.forecast_payload(r.department_id)::text))
$$;
create function article_master_private.forecast_publish(p_department_id uuid,p_actor_id uuid,p_run_id uuid,p_result jsonb) returns uuid language plpgsql security definer set search_path='' as $$
declare data jsonb;cfg public.production_forecast_settings;prior public.production_forecast_runs;v jsonb;job jsonb;stamp timestamptz;expires timestamptz;
begin
 perform pg_advisory_xact_lock(hashtextextended('forecast:all',0));
 if not coalesce(article_master_private.forecast_authorized(p_department_id,p_actor_id),false) then raise exception 'Denied' using errcode='42501';end if;
 select * into prior from public.production_forecast_runs where id=p_run_id;
 if found then if prior.actor_id<>p_actor_id or prior.department_id<>p_department_id then raise exception 'Request reused' using errcode='22023';end if;return prior.id;end if;
 select * into cfg from public.production_forecast_settings where department_id=p_department_id;
 data:=article_master_private.forecast_payload(p_department_id);stamp:=(p_result->>'as_of')::timestamptz;
 if p_result->>'input_token' is distinct from md5(data::text) or stamp is null or stamp>clock_timestamp() or stamp<clock_timestamp()-make_interval(secs=>cfg.max_age_seconds) or cfg.valid_until<=now() then raise exception 'Stale input' using errcode='40001';end if;
 if exists(select 1 from public.production_forecast_latest l join public.production_forecast_runs r on r.id=l.run_id where l.department_id=p_department_id and r.calculated_at>=stamp) then raise exception 'Older calculation' using errcode='40001';end if;
 if p_result->>'department_id' is distinct from p_department_id::text or p_result->>'schema' is distinct from '1' or jsonb_typeof(p_result->'results') is distinct from 'array' or jsonb_array_length(p_result->'results')<>jsonb_array_length(data->'jobs') or (select count(distinct x->>'operation_id') from jsonb_array_elements(p_result->'results') x)<>jsonb_array_length(data->'jobs') then raise exception 'Invalid result set' using errcode='22023';end if;
 for v in select value from jsonb_array_elements(p_result->'results') loop
  if not exists(select 1 from jsonb_array_elements(data->'jobs') j where j->>'id'=v->>'operation_id') or coalesce(v->>'color','') not in('green','orange','red','unknown') or coalesce(jsonb_typeof(v->'work_seconds'),'missing') not in('number','null') then raise exception 'Invalid result' using errcode='22023';end if;
  if v->>'color' in('green','orange') and (v->>'finish_at' is null or coalesce(v->>'status','') not in('calculated','completed')) then raise exception 'Unreliable green' using errcode='22023';end if;
 end loop;
 expires:=least(cfg.valid_until,stamp+make_interval(secs=>cfg.max_age_seconds));
 insert into public.production_forecast_runs values(p_run_id,p_department_id,p_actor_id,md5(data::text),stamp,expires,p_result);
 insert into public.production_forecast_latest values(p_department_id,p_run_id) on conflict(department_id) do update set run_id=excluded.run_id;
 for v in select value from jsonb_array_elements(p_result->'results') loop
  if v->>'arrival_today'='true' and v->>'arrival_at' is not null and v->>'color'<>'unknown' then
   select j into job from jsonb_array_elements(data->'jobs') j where j->>'id'=v->>'operation_id';
   insert into public.production_operation_arrivals(operation_id,machine_id,expected_on,expected_quantity,source_record,evidence,verified_by,verified_at,valid_until,forecast_run_id)
   values((v->>'operation_id')::uuid,(job->>'machine_id')::uuid,((v->>'arrival_at')::timestamptz at time zone cfg.zone)::date,(job->>'quantity')::integer,'forecast:'||p_run_id::text,'Berechnete Vorankündigung; keine physische Übergabe',p_actor_id,stamp,expires,p_run_id)
   on conflict(operation_id) do update set machine_id=excluded.machine_id,expected_on=excluded.expected_on,expected_quantity=excluded.expected_quantity,source_record=excluded.source_record,evidence=excluded.evidence,verified_by=excluded.verified_by,verified_at=excluded.verified_at,valid_until=excluded.valid_until,forecast_run_id=excluded.forecast_run_id where production_operation_arrivals.forecast_run_id is not null;
  end if;
 end loop;
 delete from public.production_forecast_queue where department_id=p_department_id;return p_run_id;
end; $$;
create function article_master_private.forecast_view() returns jsonb language sql security definer set search_path='' as $$
 select jsonb_build_object('departments',coalesce((select jsonb_agg(jsonb_build_object('id',s.department_id,'can_calculate',article_master_private.forecast_authorized(s.department_id,auth.uid()))) from public.production_forecast_settings s where s.enabled and article_master_private.forecast_authorized(s.department_id,auth.uid())),'[]'::jsonb),
 'rows',coalesce((select jsonb_agg(jsonb_build_object('operation_id',o.id,'department_id',os.department_id,'status',case when article_master_private.forecast_source_valid(r.id) then coalesce(v->>'status','missing') else 'stale' end,'color',case when article_master_private.forecast_source_valid(r.id) then coalesce(v->>'color','unknown') else 'unknown' end,'forecast',case when article_master_private.forecast_source_valid(r.id) then v||jsonb_build_object('zone',r.result->'zone','order_summary',(select z||jsonb_build_object('bottleneck_op',(select op_number from public.production_order_operations where id=(z->>'bottleneck')::uuid)) from jsonb_array_elements(r.result->'orders') z where z->>'order_id'=o.order_id::text)) else null end,'calculated_at',r.calculated_at,'expires_at',r.expires_at)) from public.production_order_operations o join public.production_order_article_snapshots os on os.order_id=o.order_id join public.production_forecast_settings s on s.department_id=os.department_id left join public.production_forecast_latest l on l.department_id=s.department_id left join public.production_forecast_runs r on r.id=l.run_id left join lateral(select x from jsonb_array_elements(r.result->'results') x where x->>'operation_id'=o.id::text) f(v) on true where s.enabled and article_master_private.readiness_allowed(o.id,'view')),'[]'::jsonb))
$$;
-- Source changes invalidate via token immediately; this queue is consumed by the disabled-by-default worker.
create function article_master_private.forecast_dirty() returns trigger language plpgsql security definer set search_path='' as $$
declare dept uuid;
begin
 for dept in select department_id from public.production_forecast_settings where enabled order by department_id loop
  perform pg_advisory_xact_lock(hashtextextended('forecast:all',0));
  insert into public.production_forecast_queue values(dept,clock_timestamp()) on conflict(department_id) do update set dirty_at=excluded.dirty_at;
 end loop;return null;
end; $$;
create table public.production_forecast_reports(request_id uuid primary key,operation_id uuid not null references public.production_order_operations(id),actor_id uuid not null references auth.users(id),payload jsonb not null,happened_at timestamptz not null default clock_timestamp());
create function article_master_private.forecast_report(p_operation_id uuid,p_request_id uuid,p_data jsonb) returns uuid language plpgsql security definer set search_path='' as $$
declare dept uuid;prior public.production_forecast_reports;a public.production_forecast_actuals;k text;
begin
 select s.department_id into dept from public.production_order_operations o join public.production_order_article_snapshots s on s.order_id=o.order_id where o.id=p_operation_id;
 perform pg_advisory_xact_lock(hashtextextended('forecast:all',0));
 if not coalesce(article_master_private.forecast_authorized(dept,auth.uid()),false) then raise exception 'Denied' using errcode='42501';end if;
 select * into prior from public.production_forecast_reports where request_id=p_request_id;
 if found then if prior.actor_id<>auth.uid() or prior.operation_id<>p_operation_id or prior.payload<>p_data then raise exception 'Request reused' using errcode='22023';end if;return p_request_id;end if;
 if jsonb_typeof(p_data) is distinct from 'object' then raise exception 'Invalid report' using errcode='22023';end if;
 for k in select jsonb_object_keys(p_data) loop
  if k not in('setup_complete','setup_elapsed_s','setup_remaining_s','processing_elapsed_s','current_piece_elapsed_s','good','scrap','clarify','replacement_qty','material_available_at','wait_until','evidence') then raise exception 'Invalid field' using errcode='22023';end if;
 end loop;
 a:=jsonb_populate_record(null::public.production_forecast_actuals,p_data);
 insert into public.production_forecast_actuals values(p_operation_id,a.setup_complete,a.setup_elapsed_s,a.setup_remaining_s,a.processing_elapsed_s,a.current_piece_elapsed_s,a.good,a.scrap,a.clarify,a.replacement_qty,a.material_available_at,a.wait_until,clock_timestamp(),auth.uid(),a.evidence)
 on conflict(operation_id) do update set setup_complete=excluded.setup_complete,setup_elapsed_s=excluded.setup_elapsed_s,setup_remaining_s=excluded.setup_remaining_s,processing_elapsed_s=excluded.processing_elapsed_s,current_piece_elapsed_s=excluded.current_piece_elapsed_s,good=excluded.good,scrap=excluded.scrap,clarify=excluded.clarify,replacement_qty=excluded.replacement_qty,material_available_at=excluded.material_available_at,wait_until=excluded.wait_until,observed_at=excluded.observed_at,reported_by=excluded.reported_by,evidence=excluded.evidence;
 insert into public.production_forecast_reports values(p_request_id,p_operation_id,auth.uid(),p_data,clock_timestamp());return p_request_id;
end; $$;
create function article_master_private.forecast_pending() returns jsonb language sql security definer set search_path='' as $$
 select coalesce(jsonb_agg(x),'[]'::jsonb) from(select s.department_id,s.calculation_actor_id from public.production_forecast_settings s left join public.production_forecast_queue q using(department_id) left join public.production_forecast_latest l using(department_id) left join public.production_forecast_runs r on r.id=l.run_id where s.enabled and (q.department_id is not null or r.id is null or r.expires_at<=now()) order by coalesce(q.dirty_at,r.expires_at) nulls first limit 10) x
$$;
create function article_master_private.forecast_worker_input(p_department_id uuid,p_actor_id uuid) returns jsonb language plpgsql security definer set search_path='' as $$
declare data jsonb;begin
 if not coalesce(article_master_private.forecast_authorized(p_department_id,p_actor_id),false) then raise exception 'Denied' using errcode='42501';end if;
 data:=article_master_private.forecast_payload(p_department_id);return data||jsonb_build_object('input_token',md5(data::text),'as_of',clock_timestamp());
end; $$;
create or replace function article_master_private.readiness_material(e public.production_operation_execution) returns text language sql stable security definer set search_path='' as $$
 select case when e.started_here then 'started' when e.received_quantity>=s.quantity then 'ready'
 when exists(select 1 from public.production_operation_arrivals f where f.operation_id=e.operation_id and f.machine_id=e.machine_id
 and f.expected_quantity>=s.quantity and f.expected_on=(now() at time zone 'Europe/Berlin')::date and f.verified_at<=now() and f.valid_until>now()
 and article_master_private.forecast_source_valid(f.forecast_run_id)) then 'expected' else 'planned' end
 from public.production_order_operations o join public.production_order_article_snapshots s on s.order_id=o.order_id where o.id=e.operation_id
$$;
create function public.production_forecast_input(p_department_id uuid,p_request_id uuid) returns jsonb language sql security invoker set search_path='' as $$select article_master_private.forecast_input(p_department_id,p_request_id)$$;
create function public.production_forecast_view() returns jsonb language sql security invoker set search_path='' as $$select article_master_private.forecast_view()$$;
create function public.production_forecast_report(p_operation_id uuid,p_request_id uuid,p_data jsonb) returns uuid language sql security invoker set search_path='' as $$select article_master_private.forecast_report(p_operation_id,p_request_id,p_data)$$;
create function public.production_forecast_publish(p_department_id uuid,p_actor_id uuid,p_run_id uuid,p_result jsonb) returns uuid language sql security invoker set search_path='' as $$select article_master_private.forecast_publish(p_department_id,p_actor_id,p_run_id,p_result)$$;
create function public.production_forecast_pending() returns jsonb language sql security invoker set search_path='' as $$select article_master_private.forecast_pending()$$;
create function public.production_forecast_worker_input(p_department_id uuid,p_actor_id uuid) returns jsonb language sql security invoker set search_path='' as $$select article_master_private.forecast_worker_input(p_department_id,p_actor_id)$$;
do $$declare tab text;fn record;begin
 foreach tab in array array['production_forecast_settings','production_forecast_machines','production_forecast_shifts','production_forecast_staff','production_forecast_rules','production_forecast_actuals','production_forecast_blocks','production_forecast_queue','production_forecast_runs','production_forecast_latest','production_forecast_reports'] loop
  execute format('alter table public.%I enable row level security',tab);
  execute format('revoke all on public.%I from public,anon,authenticated',tab);
 end loop;
 -- Source bindings use the existing catalog names. No existing RLS policy is changed.
 foreach tab in array array['production_forecast_settings','production_forecast_machines','production_forecast_shifts','production_forecast_staff','production_forecast_rules','production_forecast_actuals','production_forecast_blocks','planner_assignments','planner_availability','planner_absences','planner_vacations','planner_sick_leaves','planner_absence_replacements','planner_shift_cancellations','planner_swaps','planner_special_days','shift_definitions','employees','production_machines','production_orders','production_order_stations','production_station_counts','production_operation_execution'] loop
  execute format('create trigger forecast_source_dirty before insert or update or delete on public.%I for each statement execute function article_master_private.forecast_dirty()',tab);
 end loop;
 for fn in select p.oid::regprocedure signature from pg_proc p join pg_namespace n on n.oid=p.pronamespace where (n.nspname='article_master_private' and p.proname like 'forecast_%') or(n.nspname='public' and p.proname like 'production_forecast_%') loop
  execute format('revoke all on function %s from public,anon,authenticated',fn.signature);
 end loop;
end $$;
grant execute on function public.production_forecast_input(uuid,uuid),public.production_forecast_view(),public.production_forecast_report(uuid,uuid,jsonb),article_master_private.forecast_input(uuid,uuid),article_master_private.forecast_view(),article_master_private.forecast_report(uuid,uuid,jsonb) to authenticated;
grant usage on schema article_master_private to service_role;
grant execute on function public.production_forecast_publish(uuid,uuid,uuid,jsonb),article_master_private.forecast_publish(uuid,uuid,uuid,jsonb),public.production_forecast_pending(),article_master_private.forecast_pending(),public.production_forecast_worker_input(uuid,uuid),article_master_private.forecast_worker_input(uuid,uuid) to service_role;
create trigger forecast_runs_immutable before update or delete on public.production_forecast_runs for each row execute function article_master_private.immutable_snapshot();
create trigger forecast_reports_immutable before update or delete on public.production_forecast_reports for each row execute function article_master_private.immutable_snapshot();
commit;
