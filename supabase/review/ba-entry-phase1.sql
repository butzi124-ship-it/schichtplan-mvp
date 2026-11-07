-- REVIEW PROPOSAL ONLY. Not registered/applied to any hosted project.
-- Requires #7/#8 schema; test only in an explicitly isolated database.
-- Leaves legacy policies, authentication, counters and existing rows untouched.
begin;
set local lock_timeout='5s';
set local statement_timeout='20s';
create table public.production_ba_entry_settings (
  singleton boolean primary key default true check(singleton), enabled boolean not null default false
);
insert into public.production_ba_entry_settings values(true,false);
create table public.production_ba_entry_records (
  order_id uuid primary key references public.production_order_article_snapshots(order_id) on delete restrict,
  actor_id uuid not null references auth.users(id), happened_at timestamptz not null default now(),
  first_op_number integer not null check(first_op_number between 1 and 999),
  authorization_source text not null check(authorization_source in ('scope','first_op_assignment')),
  responsibility_status text not null default 'pending_rule' check(responsibility_status='pending_rule'),
  before_value jsonb not null default 'null', after_value jsonb not null
);
-- Only proposed snapshot planning metadata becomes nullable; no stored value is rewritten.
-- The previous five-argument API retains its own mandatory-date validation.
alter table public.production_order_article_snapshots alter column due_date drop not null;
create trigger immutable_ba_entry_record before update or delete on public.production_ba_entry_records
for each row execute function article_master_private.immutable_snapshot();
alter table public.production_ba_entry_settings enable row level security;
alter table public.production_ba_entry_records enable row level security;
revoke all on public.production_ba_entry_settings,public.production_ba_entry_records from public,anon,authenticated;

-- Employee records restrict active/scanner access; independent, service-provisioned
-- scopes/first-OP assignments supply authority. Mutable employee roles alone never grant access.
create function article_master_private.ba_entry_allowed(dept uuid, number text, first_op integer)
returns boolean language sql stable security definer set search_path='' as $$
  select auth.uid() is not null
    and exists(select 1 from public.production_ba_entry_settings where singleton and enabled)
    and exists(select 1 from public.departments where id=dept and active)
    and exists(select 1 from public.employees where auth_user_id=auth.uid() and is_active and role in ('admin','department_admin','employee'))
    and (exists(select 1 from public.article_master_scopes where user_id=auth.uid() and department_id=dept)
      or exists(select 1 from public.article_master_assignments where user_id=auth.uid() and department_id=dept
        and number_key=upper(trim(number)) and first_op_number=first_op and not revoked and expires_at>now()));
$$;
create function article_master_private.ba_entry_read(oid uuid, selected_op integer default null)
returns jsonb language plpgsql stable security definer set search_path='' as $$
declare s public.production_order_article_snapshots; a public.article_master_items; result jsonb; legacy public.production_orders;
begin
  select * into s from public.production_order_article_snapshots where order_id=oid;
  if s.order_id is null then
    select * into legacy from public.production_orders where id=oid;
    -- No guessing the first OP of an old BA from a station number; scope-only legacy access.
    if legacy.id is null or selected_op is not null or not article_master_private.ba_entry_allowed(legacy.department_id,'',null)
      then raise exception 'Denied' using errcode='42501'; end if;
    return jsonb_build_object('order_id',legacy.id,'ba_number',legacy.ba_number,'legacy',true);
  end if;
  select * into a from public.article_master_items where id=(select article_id from public.article_master_revisions where id=s.revision_id);
  if not article_master_private.ba_entry_allowed(s.department_id,a.article_number,(s.operations->0->>'number')::integer)
    then raise exception 'Denied' using errcode='42501'; end if;
  if selected_op is not null and not exists(select 1 from public.production_order_operations where order_id=oid and op_number=selected_op)
    then raise exception 'Unknown OP' using errcode='22023'; end if;
  result:=to_jsonb(s)||jsonb_build_object('selected_op',selected_op,'responsibility_status','pending_rule',
    'op_rows',(select jsonb_agg(to_jsonb(o) order by position) from public.production_order_operations o where order_id=oid));
  return result;
end; $$;
create function article_master_private.ba_entry_command(p_action text,p_ba_number text,p_article_number text,p_quantity integer,p_op_number integer)
returns jsonb language plpgsql security definer set search_path='' as $$
declare a public.article_master_items; r public.article_master_revisions; plan public.article_operation_plans;
  oid uuid; existing_count integer; matches integer; first_op integer; pos integer:=0; o jsonb; result jsonb; source text;
begin
  if auth.uid() is null or p_action is null or p_action not in ('create','open') then raise exception 'Denied' using errcode='42501'; end if;
  if p_ba_number is null or p_ba_number !~ '^[0-9]{1,24}$' or p_op_number is not null and p_op_number not between 1 and 999
    then raise exception 'Invalid BA/OP' using errcode='22023'; end if;
  if p_action='open' then
    select count(*),(array_agg(id))[1] into existing_count,oid from public.production_orders where trim(ba_number)=p_ba_number;
    if existing_count<>1 then raise exception 'Unknown or ambiguous BA' using errcode='42501'; end if;
    return article_master_private.ba_entry_read(oid,p_op_number);
  end if;
  if p_quantity is null or p_quantity<=0 or length(trim(coalesce(p_article_number,''))) not between 1 and 80 or p_op_number is not null
    then raise exception 'Three valid fields required' using errcode='22023'; end if;
  -- Parent -> revision -> BA lock order agrees with the existing order_from_article API.
  select count(*),(array_agg(i.id))[1] into matches,oid from public.article_master_items i
    where i.number_key=upper(trim(p_article_number)) and exists(
      select 1 from public.article_master_revisions v join public.article_operation_plans p on p.revision_id=v.id
      where v.article_id=i.id and v.status='released' and jsonb_array_length(p.operations)>0
        and article_master_private.ba_entry_allowed(i.department_id,i.article_number,(p.operations->0->>'number')::integer));
  if matches<>1 then raise exception 'Unknown or ambiguous released article / first OP' using errcode='22023'; end if;
  select * into a from public.article_master_items where id=oid for update;
  select v.* into r from public.article_master_revisions v where v.article_id=a.id and v.status='released' order by v.revision desc limit 1 for update;
  select * into plan from public.article_operation_plans where revision_id=r.id;
  first_op:=(plan.operations->0->>'number')::integer;
  -- Lock independently provisioned authority until the creation transaction completes.
  perform 1 from public.departments where id=a.department_id for share;
  perform 1 from public.employees where auth_user_id=auth.uid() for share;
  perform 1 from public.article_master_scopes where user_id=auth.uid() and department_id=a.department_id for share;
  perform 1 from public.article_master_assignments where user_id=auth.uid() and department_id=a.department_id and number_key=a.number_key for share;
  if plan.revision_id is null or not article_master_private.valid_operations(plan.operations) or first_op is null
    or not article_master_private.ba_entry_allowed(a.department_id,a.article_number,first_op)
    then raise exception 'Denied' using errcode='42501'; end if;
  perform pg_advisory_xact_lock(hashtextextended('operations-ba:'||p_ba_number,0));
  select count(*),(array_agg(id))[1] into existing_count,oid from public.production_orders where trim(ba_number)=p_ba_number;
  if existing_count>1 then raise exception 'Ambiguous legacy BA' using errcode='22023'; end if;
  if existing_count=1 then return article_master_private.ba_entry_read(oid)||jsonb_build_object('existing',true); end if;
  oid:=gen_random_uuid();
  -- Unplanned BA: no arbitrary machine, deadline, AL or synthetic station/counter.
  insert into public.production_orders(id,machine_id,department_id,ba_number,article_number,ba_quantity,target_quantity,pallet_count,pieces_per_pallet,use_chain_logic,status)
    values(oid,null,a.department_id,p_ba_number,a.article_number,p_quantity,p_quantity,0,0,false,'running');
  insert into public.production_order_article_snapshots(order_id,revision_id,department_id,ba_number,quantity,due_date,article_revision,release_status,released_by,released_at,article_data,operations,created_by)
    values(oid,r.id,a.department_id,p_ba_number,p_quantity,null,r.revision,r.status,r.released_by,r.released_at,r.data,plan.operations,auth.uid());
  insert into public.article_master_ba_links values(oid,r.id);
  for o in select value from jsonb_array_elements(plan.operations) loop
    pos:=pos+1;
    insert into public.production_order_operations(order_id,position,op_number,predecessor,data) values(oid,pos,(o->>'number')::integer,(o->>'predecessor')::integer,o);
  end loop;
  source:=case when exists(select 1 from public.article_master_scopes where user_id=auth.uid() and department_id=a.department_id) then 'scope' else 'first_op_assignment' end;
  result:=article_master_private.ba_entry_read(oid,first_op);
  insert into public.production_ba_entry_records(order_id,actor_id,first_op_number,authorization_source,after_value)
    values(oid,auth.uid(),first_op,source,result);
  return result||jsonb_build_object('existing',false);
end; $$;
create function public.ba_entry_command(p_action text,p_ba_number text default null,p_article_number text default null,p_quantity integer default null,p_op_number integer default null)
returns jsonb language sql security invoker set search_path='' as $$
  select article_master_private.ba_entry_command(p_action,p_ba_number,p_article_number,p_quantity,p_op_number);
$$;
revoke all on function article_master_private.ba_entry_allowed(uuid,text,integer),article_master_private.ba_entry_read(uuid,integer),article_master_private.ba_entry_command(text,text,text,integer,integer),public.ba_entry_command(text,text,text,integer,integer) from public,anon,authenticated;
-- Only the checked command is externally callable; read/permission helpers are private.
grant execute on function article_master_private.ba_entry_command(text,text,text,integer,integer),public.ba_entry_command(text,text,text,integer,integer) to authenticated;
commit;
