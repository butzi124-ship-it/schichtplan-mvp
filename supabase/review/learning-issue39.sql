-- REVIEW/ISOLATED INSTALLATION ONLY. Not registered in deployment migrations.
-- Requires reviewed humbel_security identity foundation; enrollments remain EMPTY.
begin;
set local lock_timeout='5s';
set local statement_timeout='30s';
do $$begin
 if not (current_database()='humbel_isolated_test' or (current_database()='postgres' and version() like '%(PGlite %'))
 or length(coalesce(current_setting('humbel.isolated_sentinel',true),''))<32
 then raise exception 'ISOLATED LEARNING REVIEW ONLY' using errcode='42501';end if;
 if to_regprocedure('humbel_security.actor_id()') is null then raise exception 'Reviewed identity prerequisite missing';end if;
end$$;
create schema humbel_learning;
revoke all on schema humbel_learning from public,anon,authenticated,service_role;
alter default privileges in schema humbel_learning revoke execute on functions from public;
create table humbel_learning.settings(singleton boolean primary key default true check(singleton),enabled boolean not null default false,
 safety_enabled boolean not null default false,thresholds integer[] not null default array[100,500,1000,2500,5000],
 thresholds_confirmed boolean not null default false,check(cardinality(thresholds)=5 and thresholds[1]>0 and thresholds[1]<thresholds[2] and thresholds[2]<thresholds[3] and thresholds[3]<thresholds[4] and thresholds[4]<thresholds[5]),participation_reward integer not null default 1 check(participation_reward between 0 and 5));
insert into humbel_learning.settings default values;
create table humbel_learning.people(
 employee_id uuid primary key references public.employees(id),user_id uuid not null unique references humbel_security.principals(user_id),
 department_id uuid not null references public.departments(id),access_enabled boolean not null default false,
 game_enabled boolean,visible boolean not null default false,learner boolean not null default false,
 curriculum text check(curriculum in('legacy2008','futuremem2026')),education_start date,
 next_offer timestamptz not null default now(),evidence text not null check(length(trim(evidence))>0),
 check(not learner or (curriculum is not null and education_start is not null)));
create table humbel_learning.department_settings(department_id uuid primary key references public.departments(id),enabled boolean not null default true);
create table humbel_learning.scopes(user_id uuid not null references humbel_security.principals(user_id),department_id uuid not null references public.departments(id),
 can_author boolean not null default false,can_manage boolean not null default false,can_global boolean not null default false,can_education boolean not null default false,
 valid_until timestamptz not null,revoked boolean not null default false,evidence text not null,primary key(user_id,department_id));
-- Trusted upstream adapter ONLY; no employee/manager endpoint can assert eligibility.
-- Unknown time/absence/safe-machine context => NO penalties and NO unsolicited offers.
create table humbel_learning.context_windows(id uuid primary key default gen_random_uuid(),employee_id uuid not null references humbel_learning.people,
 starts_at timestamptz not null,ends_at timestamptz not null,eligible boolean not null,safe_to_offer boolean not null default false,
 exempt boolean not null default false,evidence text not null,check(ends_at>starts_at));
create table humbel_learning.education_topics(employee_id uuid not null references humbel_learning.people,topic text not null,
 curriculum text not null,approved_until timestamptz not null,evidence text not null,primary key(employee_id,topic,curriculum));
create table humbel_learning.domains(code text primary key,label text not null,motif text not null);
insert into humbel_learning.domains values('milling','Fräsen','cutter'),('turning','Drehen','insert'),('grinding','Schleifen','wheel'),
 ('gearing','Verzahnen','gear'),('metrology','Messtechnik','caliper'),('quality','Qualitätssicherung','caliper'),
 ('tools','Werkzeugtechnik','cutter'),('machines','Maschinenbedienung','gear'),('safety','Arbeitssicherheit','shield');
create table humbel_learning.questions(id uuid primary key default gen_random_uuid(),department_id uuid not null references public.departments(id),
 concept text not null check(length(concept) between 1 and 100),domain text not null references humbel_learning.domains,topic text not null,
 difficulty integer not null check(difficulty between 1 and 3),prompt text not null check(length(trim(prompt)) between 1 and 4000),
 options jsonb not null,correct text[] not null,explanation text not null check(length(trim(explanation)) between 1 and 8000),
 source text not null check(length(trim(source)) between 1 and 2000),curriculum text not null default 'general' check(curriculum in('general','legacy2008','futuremem2026')),
 kind text not null default 'learning' check(kind in('learning','safety')),practical_required boolean not null default true,
 interval_months integer not null default 6 check(interval_months between 1 and 60),resource text not null default '',
 global_pool boolean not null default false,released boolean not null default false,disabled boolean not null default false,
 previous_id uuid references humbel_learning.questions(id),created_by uuid not null,created_at timestamptz not null default now());
-- Exact independently approved publication payload; no client enrollment endpoint.
create table humbel_learning.global_approvals(question_id uuid primary key references humbel_learning.questions,approved_payload jsonb not null,
 valid_until timestamptz not null,approver_evidence text not null check(length(trim(approver_evidence))>0));
create table humbel_learning.units(id uuid primary key default gen_random_uuid(),employee_id uuid not null references humbel_learning.people,
 question_id uuid not null references humbel_learning.questions,repeat_learned boolean not null,status text not null default 'open' check(status in('open','answered')),
 deferrals integer not null default 0,offered_at timestamptz not null default now(),available_at timestamptz not null default now(),answered_at timestamptz);
create unique index one_open_learning_unit on humbel_learning.units(employee_id) where status='open';
create table humbel_learning.mastery(employee_id uuid not null references humbel_learning.people,concept text not null,
 learned boolean not null default false,successes integer not null default 0,due_at timestamptz not null default now(),last_question uuid references humbel_learning.questions,
 primary key(employee_id,concept));
create table humbel_learning.events(id bigint generated always as identity primary key,employee_id uuid not null references humbel_learning.people,
 unit_id uuid references humbel_learning.units,domain text references humbel_learning.domains,kind text not null,competence integer not null default 0,
 participation integer not null default 0,experience integer not null default 0,details jsonb not null default '{}',created_at timestamptz not null default now(),actor uuid not null);
create table humbel_learning.badge_history(id bigint generated always as identity primary key,employee_id uuid not null references humbel_learning.people,
 domain text not null references humbel_learning.domains,stage integer not null check(stage between 0 and 5),created_at timestamptz not null default now());
create table humbel_learning.instruction_assignments(id uuid primary key default gen_random_uuid(),employee_id uuid not null references humbel_learning.people,
 question_id uuid not null references humbel_learning.questions,due_at timestamptz not null default now(),reason text not null,
 theory_passed_at timestamptz,failures integer not null default 0,confirmed_at timestamptz,confirmed_by uuid,practical_evidence text,next_due timestamptz);
create table humbel_learning.instruction_events(id bigint generated always as identity primary key,assignment_id uuid not null references humbel_learning.instruction_assignments,
 actor uuid not null,kind text not null,details jsonb not null,created_at timestamptz not null default now());
create table humbel_learning.receipts(user_id uuid not null,request_id uuid not null,payload jsonb not null,primary key(user_id,request_id));
-- Future reviewed quality adapter: no employee endpoint, no raw problem or blamed person.
create table humbel_learning.confirmed_topics(id uuid primary key default gen_random_uuid(),employee_id uuid not null references humbel_learning.people,
 topic text not null,confirmed_solution_evidence uuid not null,reviewed_by uuid not null,valid_until timestamptz not null);

create function humbel_learning.immutable() returns trigger language plpgsql set search_path='' set lock_timeout='3s' as $$begin
 raise exception 'Historical learning evidence is immutable' using errcode='42501';end$$;
do $$declare t text;begin
 foreach t in array array['events','badge_history','instruction_events','receipts'] loop
 execute format('create trigger immutable_row before update or delete on humbel_learning.%I for each row execute function humbel_learning.immutable()',t);
 execute format('create trigger immutable_truncate before truncate on humbel_learning.%I for each statement execute function humbel_learning.immutable()',t);
 end loop;
end$$;
create function humbel_learning.question_guard() returns trigger language plpgsql set search_path='' set lock_timeout='3s' as $$
declare v jsonb;k text;n integer;begin
 if tg_op='DELETE' or tg_op='TRUNCATE' then raise exception 'Question versions retained' using errcode='42501';end if;
 if tg_op='UPDATE' and (to_jsonb(new)-'disabled')<>(to_jsonb(old)-'disabled') then raise exception 'Create a new version' using errcode='42501';end if;
 if jsonb_typeof(new.options)<>'array' or jsonb_array_length(new.options) not between 2 and 8 or cardinality(new.correct) not between 1 and 8 then raise exception 'Invalid answers' using errcode='22023';end if;
 n=0;
 for v in select value from jsonb_array_elements(new.options) loop
 if jsonb_typeof(v)<>'object' or not(v ?& array['id','label']) or (select count(*) from jsonb_object_keys(v))<>2 or jsonb_typeof(v->'label')<>'string' or length(v->>'label') not between 1 and 1000 or (v->>'id') !~ '^[a-zA-Z0-9_-]{1,32}$' then raise exception 'Invalid option' using errcode='22023';end if;n=n+1;end loop;
 if (select count(distinct value->>'id') from jsonb_array_elements(new.options))<>n or
 (select count(distinct x) from unnest(new.correct)x)<>cardinality(new.correct) then raise exception 'Duplicate option' using errcode='22023';end if;
 foreach k in array new.correct loop if not exists(select 1 from jsonb_array_elements(new.options)z where z.value->>'id'=k) then raise exception 'Unknown solution' using errcode='22023';end if;end loop;
 return new;end$$;
create trigger validate_question before insert or update or delete on humbel_learning.questions for each row execute function humbel_learning.question_guard();
create trigger protect_question_truncate before truncate on humbel_learning.questions for each statement execute function humbel_learning.immutable();

create function humbel_learning.actor() returns uuid language plpgsql security definer set search_path='' set lock_timeout='3s' as $$
declare eid uuid;begin
 -- Lock the independently reviewed identity BEFORE trusting any mutable profile.
 perform 1 from humbel_security.principals where user_id=auth.uid() for share;
 eid=humbel_security.actor_id();
 perform 1 from public.employees where id=eid for share;
 if eid is null or humbel_security.actor_id() is distinct from eid or humbel_security.actor_role()='tool_scanner' then raise exception 'Verified active identity required' using errcode='42501';end if;
 perform 1 from humbel_learning.people p join public.departments d on d.id=p.department_id
 join humbel_security.principals r on r.user_id=p.user_id
 where p.employee_id=eid and p.user_id=auth.uid() and p.access_enabled and d.active and r.employee_id=eid and r.department_id=p.department_id for share of p,d;
 if not found then raise exception 'Independent learning enrollment required' using errcode='42501';end if;
 return eid;end$$;
create function humbel_learning.can_manage(dept uuid,cap text default 'manage') returns boolean language plpgsql security definer set search_path='' set lock_timeout='3s' as $$
declare eid uuid;ok boolean;begin
 eid=humbel_learning.actor();
 perform 1 from humbel_learning.scopes s where s.user_id=auth.uid() and s.department_id=dept for share;
 select exists(select 1 from humbel_learning.scopes s join public.departments d on d.id=s.department_id
 where s.user_id=auth.uid() and s.department_id=dept and not s.revoked and s.valid_until>clock_timestamp() and d.active
 and case cap when 'author' then s.can_author when 'global' then s.can_global when 'education' then s.can_education else s.can_manage end
 and (cap='global' or d.leader_employee_id=eid)) into ok;
 return ok;end$$;
create function humbel_learning.game_on(eid uuid) returns boolean language sql stable security definer set search_path='' set lock_timeout='3s' as $$
 select coalesce(p.game_enabled,d.enabled,true) from humbel_learning.people p left join humbel_learning.department_settings d using(department_id) where p.employee_id=eid
$$;
create function humbel_learning.eligible(eid uuid,offering boolean default false) returns boolean language sql stable security definer set search_path='' set lock_timeout='3s' as $$
 select humbel_learning.game_on(eid) and exists(select 1 from humbel_learning.context_windows w where w.employee_id=eid and w.starts_at<=clock_timestamp() and w.ends_at>clock_timestamp() and w.eligible and (not offering or w.safe_to_offer))
 and not exists(select 1 from humbel_learning.context_windows w where w.employee_id=eid and w.starts_at<=clock_timestamp() and w.ends_at>clock_timestamp() and (w.exempt or not w.eligible))
$$;
create function humbel_learning.allowed_question(eid uuid,qid uuid) returns boolean language sql stable security definer set search_path='' set lock_timeout='3s' as $$
 select exists(select 1 from humbel_learning.questions q join humbel_learning.people p on p.employee_id=eid
 where q.id=qid and q.released and not q.disabled and (q.department_id=p.department_id or q.global_pool)
 and (not p.learner or ((q.curriculum=p.curriculum or q.curriculum='general') and exists(select 1 from humbel_learning.education_topics t where t.employee_id=eid and t.topic=q.topic and t.curriculum=p.curriculum and t.approved_until>statement_timestamp()))))
$$;
create function humbel_learning.stage(eid uuid,dom text) returns integer language sql stable security definer set search_path='' set lock_timeout='3s' as $$
 select count(*)::integer from humbel_learning.settings s,unnest(s.thresholds)t where t <= coalesce((select sum(e.competence) from humbel_learning.events e where e.employee_id=eid and e.domain=dom),0)
$$;

create function humbel_learning.view() returns jsonb language plpgsql security definer set search_path='' set lock_timeout='3s' as $$
declare eid uuid;p humbel_learning.people;g boolean;s humbel_learning.settings;outcome jsonb;begin
 eid=humbel_learning.actor();select * into p from humbel_learning.people where employee_id=eid;select * into s from humbel_learning.settings;g=s.enabled;
 select jsonb_build_object('user_id',auth.uid(),'installed',true,'game_ready',g,'safety_ready',s.safety_enabled,'enabled',humbel_learning.game_on(eid),'visible',p.visible,'learner',p.learner,
 'provisional',not s.thresholds_confirmed,'competence',coalesce(sum(competence),0),'participation',coalesce(sum(participation),0),'experience',coalesce(sum(experience),0),
 'level',1+greatest(0,coalesce(sum(experience),0)) /100,'answered_count',(select count(*) from humbel_learning.units where employee_id=eid and status='answered'),'learned_count',(select count(*) from humbel_learning.mastery where employee_id=eid and learned),'next_offer',p.next_offer) into outcome from humbel_learning.events where employee_id=eid;
 return outcome||jsonb_build_object(
 'badges',(select coalesce(jsonb_agg(jsonb_build_object('code',d.code,'label',d.label,'motif',d.motif,'stage',humbel_learning.stage(eid,d.code),'points',(select coalesce(sum(competence),0) from humbel_learning.events where employee_id=eid and domain=d.code))),'[]') from humbel_learning.domains d),
 'unit',(select jsonb_build_object('id',u.id,'question_id',q.id,'prompt',q.prompt,'options',q.options,'domain',q.domain,'difficulty',q.difficulty,'available_at',u.available_at,'deferrals',u.deferrals,'repeat',u.repeat_learned) from humbel_learning.units u join humbel_learning.questions q on q.id=u.question_id where u.employee_id=eid and u.status='open' and g and humbel_learning.game_on(eid) and humbel_learning.allowed_question(eid,q.id)),
 'history',(select coalesce(jsonb_agg(to_jsonb(x) order by x.created_at desc),'[]') from (select kind,domain,competence,participation,experience,details,created_at from humbel_learning.events where employee_id=eid order by id desc limit 50)x),
 'badge_history',(select coalesce(jsonb_agg(to_jsonb(x) order by x.created_at desc),'[]') from (select domain,stage,created_at from humbel_learning.badge_history where employee_id=eid order by id desc limit 50)x),
 'review_count',(select count(*) from humbel_learning.mastery where employee_id=eid and due_at<=now()),
 'leaderboard',(select coalesce(jsonb_agg(to_jsonb(x) order by x.points desc,x.name),'[]') from (select e.display_name name,(select coalesce(sum(competence),0) from humbel_learning.events v where v.employee_id=b.employee_id) points,1+(select coalesce(sum(experience),0)/100 from humbel_learning.events v where v.employee_id=b.employee_id) level,(select jsonb_object_agg(dd.code,humbel_learning.stage(b.employee_id,dd.code)) from humbel_learning.domains dd) badges from humbel_learning.people b join public.employees e on e.id=b.employee_id join humbel_security.principals r on r.user_id=b.user_id
 where b.department_id=p.department_id and b.learner=p.learner and b.visible and b.access_enabled and e.is_active and r.enabled and r.valid_until>now() and e.auth_user_id=r.user_id and e.role=r.role and e.department_id is not distinct from r.department_id and s.enabled order by points desc limit 50)x),
 'instructions',(select coalesce(jsonb_agg(jsonb_build_object('id',a.id,'prompt',q.prompt,'options',q.options,'source',q.source,'resource',q.resource,'reason',a.reason,'due_at',a.due_at,'last_theory',(select ie.details from humbel_learning.instruction_events ie where ie.assignment_id=a.id and ie.kind='theory' order by ie.id desc limit 1),'theory_passed_at',a.theory_passed_at,'confirmed_at',a.confirmed_at,'next_due',a.next_due,'failures',a.failures,'practical_required',q.practical_required)),'[]') from humbel_learning.instruction_assignments a join humbel_learning.questions q on q.id=a.question_id where a.employee_id=eid and s.safety_enabled),
 'departments',(select coalesce(jsonb_agg(jsonb_build_object('id',d.id,'name',d.name)),'[]') from public.departments d where humbel_learning.can_manage(d.id)),
 'author_departments',(select coalesce(jsonb_agg(jsonb_build_object('id',d.id,'name',d.name)),'[]') from public.departments d where humbel_learning.can_manage(d.id,'author')));
end$$;

create function humbel_learning.admin_view(p_department uuid) returns jsonb language plpgsql security definer set search_path='' set lock_timeout='3s' as $$
declare eid uuid;begin
 eid=humbel_learning.actor();if not humbel_learning.can_manage(p_department) then raise exception 'Department authority required' using errcode='42501';end if;
 return jsonb_build_object('people',(select coalesce(jsonb_agg(jsonb_build_object('id',p.employee_id,'name',e.display_name,'enabled',humbel_learning.game_on(p.employee_id),'learner',p.learner,
 'offered',(select count(*) from humbel_learning.units where employee_id=p.employee_id),'answered',(select count(*) from humbel_learning.units where employee_id=p.employee_id and status='answered'),
 'deferrals',(select count(*) from humbel_learning.events where employee_id=p.employee_id and kind='deferred' and details->>'exempt'='false'),'participation',(select coalesce(sum(participation),0) from humbel_learning.events where employee_id=p.employee_id),
 'last_participation',(select max(answered_at) from humbel_learning.units where employee_id=p.employee_id),
 'support_topics',(select coalesce(jsonb_agg(concept),'[]') from humbel_learning.mastery where employee_id=p.employee_id and due_at<=now()))),'[]') from humbel_learning.people p join public.employees e on e.id=p.employee_id where p.department_id=p_department),
 'questions',(select coalesce(jsonb_agg(to_jsonb(q) order by q.created_at desc),'[]') from humbel_learning.questions q where q.department_id=p_department),
 'instructions',(select coalesce(jsonb_agg(to_jsonb(a)),'[]') from humbel_learning.instruction_assignments a join humbel_learning.people p on p.employee_id=a.employee_id where p.department_id=p_department));
end$$;

create function humbel_learning.command(p_action text,p_body jsonb,p_request uuid) returns jsonb language plpgsql security definer set search_path='' set lock_timeout='3s' as $$
declare eid uuid;p humbel_learning.people;t humbel_learning.people;q humbel_learning.questions;u humbel_learning.units;a humbel_learning.instruction_assignments;
 oldpayload jsonb;payload jsonb;sel text[];ok boolean;delta integer;reward integer;beforestage integer;afterstage integer;dept uuid;newid uuid;eligible boolean;begin
 perform pg_advisory_xact_lock(390039); -- one consistent lock order for learning mutations
 eid=humbel_learning.actor();select * into p from humbel_learning.people where employee_id=eid;
 if p_request is null or jsonb_typeof(p_body)<>'object' then raise exception 'Invalid command' using errcode='22023';end if;
 if p_action<>'visibility' and not exists(select 1 from humbel_learning.settings where case when p_action in('instruction_answer','instruction_assign','instruction_confirm') or (p_action='question_save' and p_body->>'kind'='safety') then safety_enabled else enabled end) then raise exception 'Not activated' using errcode='42501';end if;
 if p_action in('offer','defer') then lock table humbel_learning.context_windows in share mode;end if;
 payload=jsonb_build_object('action',p_action,'body',p_body);
 select r.payload into oldpayload from humbel_learning.receipts r where r.user_id=auth.uid() and request_id=p_request;
 if oldpayload is not null then
  if oldpayload<>payload then raise exception 'Request conflict' using errcode='40001';end if;
  -- Replay is read-only and returns only CURRENT personal authority/projection.
  return public.learning_view();
 end if;
 if p_action='visibility' then
  update humbel_learning.people set visible=(p_body->>'visible')::boolean where employee_id=eid;
  insert into humbel_learning.events(employee_id,kind,details,actor) values(eid,'visibility',jsonb_build_object('visible',(p_body->>'visible')::boolean),eid);
 elsif p_action='offer' then
  if not humbel_learning.eligible(eid,true) then raise exception 'No verified safe work context' using errcode='42501';end if;
  if not exists(select 1 from humbel_learning.units where employee_id=eid and status='open') and p.next_offer<=clock_timestamp() then
   select v.* into q from humbel_learning.questions v left join humbel_learning.mastery m on m.employee_id=eid and m.concept=v.concept
   where v.kind='learning' and humbel_learning.allowed_question(eid,v.id) and (m.due_at is null or m.due_at<=clock_timestamp())
   and (m.last_question is null or m.last_question<>v.id)
   order by (m.due_at is not null) desc,exists(select 1 from humbel_learning.confirmed_topics ct where ct.employee_id=eid and ct.topic=v.topic and ct.valid_until>clock_timestamp()) desc,
   coalesce(m.due_at,'epoch'),abs(v.difficulty-least(3,1+coalesce(m.successes,0)/3)),random() limit 1;
   if q.id is not null then insert into humbel_learning.units(employee_id,question_id,repeat_learned) values(eid,q.id,coalesce((select learned from humbel_learning.mastery where employee_id=eid and concept=q.concept),false));end if;
  end if;
 elsif p_action in('answer','defer') then
  select * into u from humbel_learning.units where id=(p_body->>'unit')::uuid and employee_id=eid for update;
  if u.id is null or u.status<>'open' then raise exception 'Unit already processed or foreign' using errcode='40001';end if;
  select * into q from humbel_learning.questions where id=u.question_id;
  if not humbel_learning.game_on(eid) or not humbel_learning.allowed_question(eid,q.id) then raise exception 'No current question permission' using errcode='42501';end if;
  if u.available_at>clock_timestamp() then raise exception 'Unit deferred' using errcode='22023';end if;
  if p_action='defer' then
   eligible=humbel_learning.eligible(eid,false);
   if p_body->>'until' not in('30m','today','next_workday') then raise exception 'Invalid defer option' using errcode='22023';end if;
   if p_body->>'until'='next_workday' then
    select min(w.starts_at) into a.due_at from humbel_learning.context_windows w where w.employee_id=eid and w.starts_at>clock_timestamp() and w.eligible and not w.exempt;
    if a.due_at is null then raise exception 'Next workday not verified' using errcode='22023';end if;
   else a.due_at=clock_timestamp()+case when p_body->>'until'='30m' then interval '30 minutes' else interval '3 hours' end;end if;
   update humbel_learning.units set deferrals=deferrals+case when eligible then 1 else 0 end,available_at=a.due_at where id=u.id;
   insert into humbel_learning.events(employee_id,unit_id,kind,participation,actor,details) values(eid,u.id,'deferred',case when eligible and u.deferrals>=2 then -5 else 0 end,eid,jsonb_build_object('count',u.deferrals+case when eligible then 1 else 0 end,'exempt',not eligible,'until',a.due_at));
  else
   select array_agg(x order by x) into sel from jsonb_array_elements_text(p_body->'answers') x;
   if sel is null or cardinality(sel)<>(select count(distinct x) from unnest(sel)x) or exists(select 1 from unnest(sel)x where not exists(select 1 from jsonb_array_elements(q.options)v where v->>'id'=x)) then raise exception 'Invalid answer selection' using errcode='22023';end if;
   ok=sel=(select array_agg(x order by x) from unnest(q.correct)x);
   delta=case when ok then (array[5,10,20])[q.difficulty] when u.repeat_learned then -(array[3,6,12])[q.difficulty] else 0 end;
   beforestage=humbel_learning.stage(eid,q.domain);select participation_reward into reward from humbel_learning.settings;
   insert into humbel_learning.events(employee_id,unit_id,domain,kind,competence,participation,experience,details,actor) values(eid,u.id,q.domain,'answered',delta,reward,greatest(delta,0),jsonb_build_object('correct',ok,'explanation',q.explanation,'source',q.source,'question',q.id),eid);
   insert into humbel_learning.mastery(employee_id,concept,learned,successes,due_at,last_question) values(eid,q.concept,ok,case when ok then 1 else 0 end,clock_timestamp()+case when ok then interval '14 days' else interval '1 day' end,q.id)
   on conflict(employee_id,concept) do update set learned=humbel_learning.mastery.learned or ok,successes=case when ok then humbel_learning.mastery.successes+1 else greatest(0,humbel_learning.mastery.successes-1) end,due_at=excluded.due_at,last_question=q.id;
   update humbel_learning.units set status='answered',answered_at=clock_timestamp(),deferrals=0 where id=u.id;
   update humbel_learning.people set next_offer=clock_timestamp()+make_interval(hours=>24+(random()*48)::integer) where employee_id=eid;
   afterstage=humbel_learning.stage(eid,q.domain);if beforestage<>afterstage then insert into humbel_learning.badge_history(employee_id,domain,stage) values(eid,q.domain,afterstage);end if;
  end if;
 elsif p_action in('person_enabled','department_enabled','question_save','question_disable','education_topic','instruction_assign','instruction_confirm') then
  dept=(p_body->>'department')::uuid;
  if not humbel_learning.can_manage(dept,case when p_action in('question_save','question_disable') then 'author' when p_action='education_topic' then 'education' else 'manage' end) then raise exception 'Own department permission required' using errcode='42501';end if;
  if p_action='department_enabled' then
   insert into humbel_learning.department_settings values(dept,(p_body->>'enabled')::boolean) on conflict(department_id) do update set enabled=excluded.enabled;
  elsif p_action in('person_enabled','education_topic','instruction_assign') then
   select * into t from humbel_learning.people where employee_id=(p_body->>'person')::uuid and department_id=dept;
   if t.employee_id is null then raise exception 'Foreign employee' using errcode='42501';end if;
   if p_action='person_enabled' then update humbel_learning.people set game_enabled=case when p_body->>'enabled' is null then null else (p_body->>'enabled')::boolean end where employee_id=t.employee_id;
   elsif p_action='education_topic' then
    if t.employee_id=eid then raise exception 'No self-approval of education topics' using errcode='42501';end if;
    if not t.learner then raise exception 'Not an apprentice' using errcode='22023';end if;
    insert into humbel_learning.education_topics values(t.employee_id,p_body->>'topic',t.curriculum,(p_body->>'until')::timestamptz,p_body->>'evidence') on conflict(employee_id,topic,curriculum) do update set approved_until=excluded.approved_until,evidence=excluded.evidence;
   else
    select * into q from humbel_learning.questions where id=(p_body->>'question')::uuid and kind='safety' and released and not disabled and (department_id=dept or global_pool);
    if q.id is null or not humbel_learning.allowed_question(t.employee_id,q.id) or length(trim(coalesce(p_body->>'reason','')))=0 then raise exception 'Approved appropriate instruction and trigger required' using errcode='22023';end if;
    insert into humbel_learning.instruction_assignments(employee_id,question_id,due_at,reason) values(t.employee_id,q.id,coalesce((p_body->>'due')::timestamptz,clock_timestamp()),p_body->>'reason');
   end if;
  elsif p_action='instruction_confirm' then
   select z.* into a from humbel_learning.instruction_assignments z join humbel_learning.people b on b.employee_id=z.employee_id where z.id=(p_body->>'assignment')::uuid and b.department_id=dept for update of z;
   if a.employee_id=eid then raise exception 'No self-confirmation of instruction' using errcode='42501';end if;
   if a.id is null or a.theory_passed_at is null or a.confirmed_at is not null or length(trim(coalesce(p_body->>'evidence','')))=0 then raise exception 'Theory and actual instruction evidence required' using errcode='22023';end if;
   select * into q from humbel_learning.questions where id=a.question_id;
   update humbel_learning.instruction_assignments set confirmed_at=clock_timestamp(),confirmed_by=eid,practical_evidence=p_body->>'evidence',next_due=clock_timestamp()+make_interval(months=>q.interval_months) where id=a.id;
   insert into humbel_learning.instruction_assignments(employee_id,question_id,due_at,reason) values(a.employee_id,a.question_id,clock_timestamp()+make_interval(months=>q.interval_months),'Regelmäßige Wiederholung');
   insert into humbel_learning.instruction_events(assignment_id,actor,kind,details) values(a.id,eid,'confirmed',jsonb_build_object('evidence',p_body->>'evidence','version',q.id));
  elsif p_action='question_disable' then
   update humbel_learning.questions set disabled=true where id=(p_body->>'question')::uuid and department_id=dept and (not global_pool or humbel_learning.can_manage(dept,'global'));
   if not found then raise exception 'Foreign question' using errcode='42501';end if;
  else
   if coalesce((p_body->>'global_pool')::boolean,false) and coalesce((p_body->>'released')::boolean,false) and
    (not humbel_learning.can_manage(dept,'global') or not exists(select 1 from humbel_learning.global_approvals ga where ga.question_id=(p_body->>'previous_id')::uuid and ga.approved_payload=p_body and ga.valid_until>clock_timestamp()))
    then raise exception 'Exact independent global publication approval required' using errcode='42501';end if;
   if p_body->>'previous_id' is not null and not exists(select 1 from humbel_learning.questions where id=(p_body->>'previous_id')::uuid and department_id=dept) then raise exception 'Foreign revision' using errcode='42501';end if;
   if p_body->>'kind'='safety' and not humbel_learning.can_manage(dept) then raise exception 'Instruction authority required' using errcode='42501';end if;
   insert into humbel_learning.questions(department_id,concept,domain,topic,difficulty,prompt,options,correct,explanation,source,curriculum,kind,interval_months,resource,global_pool,released,previous_id,created_by)
   values(dept,p_body->>'concept',p_body->>'domain',p_body->>'topic',(p_body->>'difficulty')::integer,p_body->>'prompt',p_body->'options',array(select jsonb_array_elements_text(p_body->'correct')),p_body->>'explanation',p_body->>'source',coalesce(p_body->>'curriculum','general'),coalesce(p_body->>'kind','learning'),coalesce((p_body->>'interval_months')::integer,6),coalesce(p_body->>'resource',''),coalesce((p_body->>'global_pool')::boolean,false),coalesce((p_body->>'released')::boolean,false),(p_body->>'previous_id')::uuid,eid) returning id into newid;
  end if;
  insert into humbel_learning.events(employee_id,kind,details,actor) values(eid,'administration',jsonb_build_object('action',p_action,'department',dept,'target',coalesce(p_body->>'person',p_body->>'question',p_body->>'assignment'), 'version',p_body->>'previous_id','new_revision',newid,'enabled',p_body->'enabled'),eid);
 elsif p_action='instruction_answer' then
  select * into a from humbel_learning.instruction_assignments where id=(p_body->>'assignment')::uuid and employee_id=eid for update;
  if a.id is null or a.confirmed_at is not null then raise exception 'Invalid instruction assignment' using errcode='42501';end if;
  select * into q from humbel_learning.questions where id=a.question_id;
  if not humbel_learning.allowed_question(eid,q.id) then raise exception 'Instruction version not available' using errcode='42501';end if;
  select array_agg(x order by x) into sel from jsonb_array_elements_text(p_body->'answers')x;
  ok=coalesce(sel=(select array_agg(x order by x) from unnest(q.correct)x),false);
  update humbel_learning.instruction_assignments set theory_passed_at=case when ok then clock_timestamp() else null end,failures=failures+case when ok then 0 else 1 end where id=a.id;
  insert into humbel_learning.instruction_events(assignment_id,actor,kind,details) values(a.id,eid,'theory',jsonb_build_object('passed',ok,'explanation',q.explanation,'version',q.id));
  if not ok and a.failures+1>=2 then insert into humbel_learning.instruction_events(assignment_id,actor,kind,details) values(a.id,eid,'support_required',jsonb_build_object('failures',a.failures+1,'department',p.department_id));end if;
 else raise exception 'Unknown learning action' using errcode='22023';end if;
 insert into humbel_learning.receipts values(auth.uid(),p_request,payload);
 return public.learning_view();
end$$;

do $$declare t text;begin
 for t in select tablename from pg_tables where schemaname='humbel_learning' loop
 execute format('alter table humbel_learning.%I enable row level security',t);
 execute format('revoke all on humbel_learning.%I from public,anon,authenticated,service_role',t);
 end loop;
end$$;
revoke all on all sequences in schema humbel_learning from public,anon,authenticated,service_role;
revoke all on all functions in schema humbel_learning from public,anon,authenticated,service_role;
grant usage on schema humbel_learning to authenticated;
grant execute on function humbel_learning.view(),humbel_learning.admin_view(uuid),humbel_learning.command(text,jsonb,uuid) to authenticated;
create function public.learning_view() returns jsonb language sql security invoker set search_path='' as $$select humbel_learning.view()$$;
create function public.learning_admin_view(p_department uuid) returns jsonb language sql security invoker set search_path='' as $$select humbel_learning.admin_view(p_department)$$;
create function public.learning_command(p_action text,p_body jsonb,p_request uuid) returns jsonb language sql security invoker set search_path='' as $$select humbel_learning.command(p_action,p_body,p_request)$$;
revoke all on function public.learning_view(),public.learning_admin_view(uuid),public.learning_command(text,jsonb,uuid) from public,anon,service_role;
grant execute on function public.learning_view(),public.learning_admin_view(uuid),public.learning_command(text,jsonb,uuid) to authenticated;
commit;
