-- REVIEW ONLY. Requires the un-applied issue #7 migration first. Never run on production by CI.
begin;
create table public.article_operation_plans (
  revision_id uuid primary key references public.article_master_revisions(id) on delete restrict,
  operations jsonb not null default '[]', row_version integer not null default 1 check(row_version>0),
  updated_by uuid not null references auth.users(id), updated_at timestamptz not null default now()
);
create table public.production_order_article_snapshots (
  order_id uuid primary key references public.production_orders(id) on delete restrict,
  revision_id uuid not null references public.article_master_revisions(id) on delete restrict,
  department_id uuid not null references public.departments(id),
  ba_number text not null unique check(ba_number ~ '^[0-9]{1,24}$'),
  quantity integer not null check(quantity>0), due_date date not null,
  article_revision integer not null check(article_revision>0),
  release_status text not null check(release_status='released'),
  released_by uuid not null references auth.users(id), released_at timestamptz not null,
  article_data jsonb not null, operations jsonb not null,
  created_by uuid not null references auth.users(id), created_at timestamptz not null default now()
);
create table public.production_order_operations (
  id uuid primary key default gen_random_uuid(),
  order_id uuid not null references public.production_order_article_snapshots(order_id) on delete restrict,
  position integer not null check(position>0), op_number integer not null check(op_number between 1 and 999),
  predecessor integer, data jsonb not null,
  unique(order_id,op_number), unique(order_id,position),
  foreign key(order_id,predecessor) references public.production_order_operations(order_id,op_number) deferrable initially deferred
);
-- A position is not a station. No legacy station/counter rows are created or converted.
create function article_master_private.valid_operations(ops jsonb)
returns boolean language plpgsql immutable set search_path='' as $$
declare o jsonb; t jsonb; k text; seen integer[] := '{}'; n integer; pred integer;
begin
  if ops is null or jsonb_typeof(ops)<>'array' or jsonb_array_length(ops)>99 then return false; end if;
  for o in select value from jsonb_array_elements(ops) loop
    if jsonb_typeof(o)<>'object' or not (o ?& array['number','predecessor','activity','description','machine_group','machine_id','tr','te','technical_notes','inspection_notes','programs','fixtures','tools','drawings','equipment'])
      or exists(select 1 from jsonb_object_keys(o) x where x <> all(array['number','predecessor','activity','description','machine_group','machine_id','tr','te','technical_notes','inspection_notes','programs','fixtures','tools','drawings','equipment'])) then return false; end if;
    if jsonb_typeof(o->'number')<>'number' or (o->>'number') !~ '^[0-9]{1,3}$' then return false; end if;
    n := (o->>'number')::integer;
    if n<1 or n>999 or n=any(seen) then return false; end if;
    if o->'predecessor'<>'null'::jsonb then
      if jsonb_typeof(o->'predecessor')<>'number' or (o->>'predecessor') !~ '^[0-9]{1,3}$' then return false; end if;
      pred := (o->>'predecessor')::integer;
      if not pred=any(seen) then return false; end if;
    end if;
    foreach k in array array['activity','description','machine_group','technical_notes','inspection_notes','programs','fixtures','tools','drawings','equipment'] loop
      if jsonb_typeof(o->k)<>'string' or length(o->>k)>12000 then return false; end if;
    end loop;
    if length(trim(o->>'activity'))=0 or length(trim(o->>'machine_group'))=0 then return false; end if;
    if o->'machine_id'<>'null'::jsonb and (jsonb_typeof(o->'machine_id')<>'string' or (o->>'machine_id') !~ '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') then return false; end if;
    foreach k in array array['tr','te'] loop
      t := o->k;
      if jsonb_typeof(t)<>'object' or not(t ?& array['value','unit']) or (select count(*) from jsonb_object_keys(t))<>2 then return false; end if;
      if t->'value'='null'::jsonb then
        if t->'unit'<>'null'::jsonb then return false; end if;
      else
        if jsonb_typeof(t->'value')<>'number' or (t->>'value')::numeric<0 or (t->>'value')::numeric>1000000 or coalesce(t->>'unit','') not in ('s','min','h') then return false; end if;
      end if;
    end loop;
    seen := array_append(seen,n);
  end loop;
  return true;
end; $$;
alter table public.article_operation_plans add constraint valid_plan check(article_master_private.valid_operations(operations));
alter table public.production_order_article_snapshots add constraint valid_snapshot_plan check(article_master_private.valid_operations(operations) and jsonb_array_length(operations)>0);
create function article_master_private.protect_plan() returns trigger language plpgsql security definer set search_path='' as $$
declare v public.article_master_revisions;
begin
  if tg_op='DELETE' then raise exception 'Plans are retained' using errcode='42501'; end if;
  if tg_op='UPDATE' and new.revision_id<>old.revision_id then raise exception 'Immutable plan identity' using errcode='42501'; end if;
  select * into v from public.article_master_revisions where id=new.revision_id;
  if v.status<>'draft' or exists(select 1 from public.article_master_ba_links where revision_id=v.id) then raise exception 'Frozen plan' using errcode='42501'; end if;
  return new;
end; $$;
create trigger protect_operation_plan before insert or update or delete on public.article_operation_plans for each row execute function article_master_private.protect_plan();
create function article_master_private.check_snapshot_operation() returns trigger language plpgsql security definer set search_path='' as $$
declare expected jsonb;
begin
  select operations->(new.position-1) into expected from public.production_order_article_snapshots where order_id=new.order_id;
  if expected is null or new.data is distinct from expected or new.op_number is distinct from (expected->>'number')::integer or new.predecessor is distinct from (expected->>'predecessor')::integer then
    raise exception 'Operation must match its immutable snapshot position' using errcode='23514';
  end if;
  return new;
end; $$;
create trigger check_snapshot_operation before insert on public.production_order_operations for each row execute function article_master_private.check_snapshot_operation();
create function article_master_private.immutable_snapshot() returns trigger language plpgsql set search_path='' as $$
begin raise exception 'Immutable order snapshot' using errcode='42501'; end; $$;
create trigger immutable_order_article_snapshot before update or delete on public.production_order_article_snapshots for each row execute function article_master_private.immutable_snapshot();
create trigger immutable_order_operation before update or delete on public.production_order_operations for each row execute function article_master_private.immutable_snapshot();

-- Protect only newly snapshot-backed orders; do not rewrite legacy orders or their RLS.
-- Shared lock covers legacy inserts/BA renames racing with the new transactional API.
create function article_master_private.protect_snapshot_order() returns trigger language plpgsql security definer set search_path='' as $$
begin
  if tg_op='UPDATE' and exists(select 1 from public.production_order_article_snapshots s where s.order_id=old.id)
    and (new.id,new.ba_number,new.article_number,new.ba_quantity,new.target_quantity,new.department_id) is distinct from
        (old.id,old.ba_number,old.article_number,old.ba_quantity,old.target_quantity,old.department_id) then
    raise exception 'Snapshot order identity and quantity are immutable' using errcode='42501';
  end if;
  perform pg_advisory_xact_lock(hashtextextended('operations-ba:'||trim(new.ba_number),0));
  if exists(select 1 from public.production_order_article_snapshots s where s.ba_number=trim(new.ba_number) and s.order_id<>new.id) then
    raise exception 'BA already reserved by snapshot order' using errcode='23505';
  end if;
  return new;
end; $$;
create trigger protect_snapshot_order before insert or update on public.production_orders for each row execute function article_master_private.protect_snapshot_order();

create function article_master_private.operation_plan(p_revision_id uuid,p_operations jsonb,p_expected_version integer)
returns jsonb language plpgsql security definer set search_path='' as $$
declare a public.article_master_items; r public.article_master_revisions; p public.article_operation_plans; o jsonb;
begin
  select i.* into a from public.article_master_items i join public.article_master_revisions v on v.article_id=i.id where v.id=p_revision_id for update of i;
  select * into r from public.article_master_revisions where id=p_revision_id for update;
  if r.id is null or not article_master_private.allowed(a.department_id,a.article_number,case when p_operations is null then 'read' else 'update' end,r.created_by) then raise exception 'Denied' using errcode='42501'; end if;
  select * into p from public.article_operation_plans where revision_id=r.id;
  if p_operations is not null then
    if r.status<>'draft' or r.row_version is distinct from p_expected_version then raise exception 'Stale or frozen revision' using errcode='40001'; end if;
    if not article_master_private.valid_operations(p_operations) then raise exception 'Invalid operations' using errcode='22023'; end if;
    for o in select value from jsonb_array_elements(p_operations) loop
      if o->>'machine_id' is not null and not exists(select 1 from public.production_machines m where m.id=(o->>'machine_id')::uuid and m.department_id=a.department_id and m.active) then raise exception 'Foreign or inactive machine' using errcode='42501'; end if;
    end loop;
    insert into public.article_operation_plans(revision_id,operations,updated_by) values(r.id,p_operations,auth.uid())
      on conflict(revision_id) do update set operations=excluded.operations,row_version=article_operation_plans.row_version+1,updated_by=auth.uid(),updated_at=now() returning * into p;
    update public.article_master_revisions set row_version=row_version+1,updated_at=now() where id=r.id returning * into r;
    insert into public.article_master_audit(revision_id,actor_id,action,row_version) values(r.id,auth.uid(),'operations',r.row_version);
  end if;
  return jsonb_build_object('operations',coalesce(p.operations,'[]'::jsonb),'row_version',r.row_version,'plan_version',coalesce(p.row_version,0),'status',r.status,'revision',r.revision);
end; $$;

create function article_master_private.operations_article_command(p_action text,p_revision_id uuid,p_department_id uuid,p_article_number text,p_data jsonb,p_expected_version integer)
returns jsonb language plpgsql security definer set search_path='' as $$
declare result jsonb; plan jsonb;
begin
  -- Original command performs parent/revision locks, authorization, stale checks and audit.
  result := article_master_private.command(p_action,p_revision_id,p_department_id,p_article_number,p_data,p_expected_version);
  if p_action='release' and not exists(select 1 from public.article_operation_plans where revision_id=p_revision_id and jsonb_array_length(operations)>0) then
    raise exception 'A complete operation plan is required' using errcode='22023';
  end if;
  if p_action='revise' then
    select operations into plan from public.article_operation_plans where revision_id=p_revision_id;
    if plan is not null then insert into public.article_operation_plans(revision_id,operations,updated_by) values((result->>'id')::uuid,plan,auth.uid()); end if;
  end if;
  return result;
end; $$;
create function article_master_private.order_from_article(p_revision_id uuid,p_ba_number text,p_quantity integer,p_due_date date,p_machine_id uuid)
returns jsonb language plpgsql security definer set search_path='' as $$
declare a public.article_master_items; r public.article_master_revisions; ops jsonb; o jsonb; oid uuid := gen_random_uuid(); pos integer:=0;
begin
  select i.* into a from public.article_master_items i join public.article_master_revisions v on v.article_id=i.id where v.id=p_revision_id for update of i;
  select * into r from public.article_master_revisions where id=p_revision_id for update;
  if r.id is null or not article_master_private.allowed(a.department_id,a.article_number,'release',r.created_by) then raise exception 'Denied' using errcode='42501'; end if;
  if r.status<>'released' then raise exception 'Released revision required' using errcode='22023'; end if;
  if p_ba_number is null or p_ba_number !~ '^[0-9]{1,24}$' or p_quantity is null or p_quantity<=0 or p_due_date is null or p_due_date<current_date then raise exception 'Invalid new order fields' using errcode='22023'; end if;
  if not exists(select 1 from public.production_machines where id=p_machine_id and department_id=a.department_id and active) then raise exception 'Foreign or inactive machine' using errcode='42501'; end if;
  select operations into ops from public.article_operation_plans where revision_id=r.id;
  if ops is null or not article_master_private.valid_operations(ops) or jsonb_array_length(ops)=0 then raise exception 'Operation plan required' using errcode='22023'; end if;
  perform pg_advisory_xact_lock(hashtextextended('operations-ba:'||p_ba_number,0));
  if exists(select 1 from public.production_orders where trim(ba_number)=p_ba_number) then raise exception 'Duplicate BA' using errcode='23505'; end if;
  insert into public.production_orders(id,machine_id,department_id,ba_number,article_number,ba_quantity,target_quantity,pallet_count,pieces_per_pallet,use_chain_logic,status)
    values(oid,p_machine_id,a.department_id,p_ba_number,a.article_number,p_quantity,p_quantity,0,0,false,'running');
  insert into public.production_order_article_snapshots(order_id,revision_id,department_id,ba_number,quantity,due_date,article_revision,release_status,released_by,released_at,article_data,operations,created_by)
    values(oid,r.id,a.department_id,p_ba_number,p_quantity,p_due_date,r.revision,r.status,r.released_by,r.released_at,r.data,ops,auth.uid());
  insert into public.article_master_ba_links values(oid,r.id);
  for o in select value from jsonb_array_elements(ops) loop
    pos:=pos+1;
    insert into public.production_order_operations(order_id,position,op_number,predecessor,data) values(oid,pos,(o->>'number')::integer,(o->>'predecessor')::integer,o);
  end loop;
  return article_master_private.operation_order_read(oid);
end; $$;
create function article_master_private.operation_order_read(p_order_id uuid)
returns jsonb language plpgsql stable security definer set search_path='' as $$
declare s public.production_order_article_snapshots;
begin
  select * into s from public.production_order_article_snapshots where order_id=p_order_id;
  if s.order_id is null or not article_master_private.allowed(s.department_id,'','release') then raise exception 'Denied' using errcode='42501'; end if;
  return to_jsonb(s)||jsonb_build_object('op_rows',coalesce((select jsonb_agg(to_jsonb(o) order by position) from public.production_order_operations o where o.order_id=s.order_id),'[]'::jsonb));
end; $$;
create function article_master_private.operation_scan(p_barcode text)
returns jsonb language plpgsql stable security definer set search_path='' as $$
declare ba text; op integer; oid uuid; result jsonb;
begin
  if p_barcode is null or p_barcode !~ '^[0-9]{1,24}(\.[0-9]{3})?$' then raise exception 'Invalid barcode' using errcode='22023'; end if;
  ba:=split_part(p_barcode,'.',1);
  if position('.' in p_barcode)>0 then op:=split_part(p_barcode,'.',2)::integer; if op<1 then raise exception 'Invalid OP' using errcode='22023'; end if; end if;
  select order_id into oid from public.production_order_article_snapshots where ba_number=ba;
  result:=article_master_private.operation_order_read(oid);
  if op is not null and not exists(select 1 from public.production_order_operations where order_id=oid and op_number=op) then raise exception 'Unknown OP' using errcode='22023'; end if;
  return result||jsonb_build_object('selected_op',op);
end; $$;

create function public.operation_plan(p_revision_id uuid,p_operations jsonb default null,p_expected_version integer default null) returns jsonb language sql security invoker set search_path='' as $$select article_master_private.operation_plan(p_revision_id,p_operations,p_expected_version)$$;
create function public.operations_article_command(p_action text,p_revision_id uuid default null,p_department_id uuid default null,p_article_number text default null,p_data jsonb default null,p_expected_version integer default null) returns jsonb language sql security invoker set search_path='' as $$select article_master_private.operations_article_command(p_action,p_revision_id,p_department_id,p_article_number,p_data,p_expected_version)$$;
create function public.order_from_article(p_revision_id uuid,p_ba_number text,p_quantity integer,p_due_date date,p_machine_id uuid) returns jsonb language sql security invoker set search_path='' as $$select article_master_private.order_from_article(p_revision_id,p_ba_number,p_quantity,p_due_date,p_machine_id)$$;
create function public.operation_order_read(p_order_id uuid) returns jsonb language sql stable security invoker set search_path='' as $$select article_master_private.operation_order_read(p_order_id)$$;
create function public.operation_scan(p_barcode text) returns jsonb language sql stable security invoker set search_path='' as $$select article_master_private.operation_scan(p_barcode)$$;
alter table public.article_operation_plans enable row level security;
alter table public.production_order_article_snapshots enable row level security;
alter table public.production_order_operations enable row level security;
-- Deny-by-default RLS intentionally has no client table policies/grants; all access via checked RPCs.
revoke all on public.article_operation_plans,public.production_order_article_snapshots,public.production_order_operations from public,anon,authenticated;
revoke all on all functions in schema article_master_private from public,anon;
revoke all on function article_master_private.protect_plan(),article_master_private.immutable_snapshot(),article_master_private.protect_snapshot_order(),article_master_private.check_snapshot_operation() from authenticated;
grant execute on function article_master_private.operation_plan(uuid,jsonb,integer),article_master_private.operations_article_command(text,uuid,uuid,text,jsonb,integer),article_master_private.order_from_article(uuid,text,integer,date,uuid),article_master_private.operation_order_read(uuid),article_master_private.operation_scan(text),article_master_private.valid_operations(jsonb) to authenticated;
revoke all on function public.operation_plan(uuid,jsonb,integer),public.operations_article_command(text,uuid,uuid,text,jsonb,integer),public.order_from_article(uuid,text,integer,date,uuid),public.operation_order_read(uuid),public.operation_scan(text) from public,anon;
grant execute on function public.operation_plan(uuid,jsonb,integer),public.operations_article_command(text,uuid,uuid,text,jsonb,integer),public.order_from_article(uuid,text,integer,date,uuid),public.operation_order_read(uuid),public.operation_scan(text) to authenticated;
commit;
