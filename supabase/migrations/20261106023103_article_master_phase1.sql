-- REVIEW ONLY: never applied by CI or this task. See docs/articles/SECURITY.md.
-- New tables only; existing production orders/stations/counts and policies stay unchanged.
begin;
create schema article_master_private;
revoke all on schema article_master_private from public, anon;
grant usage on schema article_master_private to authenticated;

-- Authoritative access provisioning: not derived from writable employees or JWT user_metadata.
create table public.article_master_scopes (
  user_id uuid not null references auth.users(id),
  department_id uuid not null references public.departments(id),
  primary key (user_id, department_id)
);
create table public.article_master_settings (
  department_id uuid primary key references public.departments(id),
  require_second_reviewer boolean not null default false
);
create table public.article_master_assignments (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id),
  department_id uuid not null references public.departments(id),
  article_number text not null check (length(trim(article_number)) between 1 and 80),
  number_key text generated always as (upper(trim(article_number))) stored,
  first_op_number integer not null check (first_op_number > 0),
  -- Must refer to independently verified first OP, never infer OP from a station number.
  first_op_evidence text not null check (length(trim(first_op_evidence)) > 0),
  expires_at timestamptz not null,
  revoked boolean not null default false,
  issued_by uuid not null references auth.users(id),
  issued_at timestamptz not null default now()
);
create index article_master_assignment_lookup on public.article_master_assignments(user_id, department_id, number_key) where not revoked;
create table public.article_master_items (
  id uuid primary key default gen_random_uuid(),
  department_id uuid not null references public.departments(id),
  article_number text not null check (length(trim(article_number)) between 1 and 80),
  number_key text generated always as (upper(trim(article_number))) stored,
  unique (department_id, number_key)
);
create index article_master_number_prefix on public.article_master_items(number_key text_pattern_ops);
create table public.article_master_revisions (
  id uuid primary key default gen_random_uuid(),
  article_id uuid not null references public.article_master_items(id) on delete restrict,
  revision integer not null check (revision > 0),
  status text not null default 'draft' check (status in ('draft','released','archived')),
  data jsonb not null,
  row_version integer not null default 1 check (row_version > 0),
  created_by uuid not null references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  released_by uuid references auth.users(id), released_at timestamptz,
  archived_by uuid references auth.users(id), archived_at timestamptz,
  unique (article_id, revision),
  check ((released_by is null) = (released_at is null)),
  check ((archived_by is null) = (archived_at is null))
);
create unique index article_master_one_draft on public.article_master_revisions(article_id) where status='draft';
create table public.article_master_audit (
  id bigint generated always as identity primary key,
  revision_id uuid not null references public.article_master_revisions(id) on delete restrict,
  actor_id uuid not null references auth.users(id), action text not null,
  row_version integer not null, happened_at timestamptz not null default now()
);
-- Future phase can create links via its separately reviewed write API. No legacy BA alteration.
create table public.article_master_ba_links (
  order_id uuid primary key references public.production_orders(id) on delete restrict,
  revision_id uuid not null references public.article_master_revisions(id) on delete restrict
);
create index article_master_ba_revision on public.article_master_ba_links(revision_id);

create function article_master_private.allowed(dept uuid, number text, operation text, creator uuid default null)
returns boolean language sql stable security definer set search_path = '' as $$
  select auth.uid() is not null and exists(select 1 from public.departments d where d.id=dept and d.active) and (
    exists(select 1 from public.article_master_scopes s where s.user_id=auth.uid() and s.department_id=dept)
    or (operation in ('read','create','update') and (operation='create' or creator=auth.uid()) and exists(
      select 1 from public.article_master_assignments a where a.user_id=auth.uid() and a.department_id=dept
        and a.number_key=upper(trim(number)) and not a.revoked and a.expires_at>now()
    ))
  );
$$;

-- Validate exact reusable-field allowlist, including hostile direct RPC requests.
create function article_master_private.valid_data(value jsonb)
returns boolean language sql immutable set search_path = '' as $$
  select case when jsonb_typeof(value)='object' then not exists (
    select 1 from jsonb_each(value) e where e.key <> all(array[
      'designation','customer','drawing_number','drawing_revision','material','dimensions','specification',
      'technical_notes','programs','fixtures','tools','equipment','inspection_notes','internal_notes'
    ]) or jsonb_typeof(e.value)<>'string' or length(e.value #>> '{}')>12000
  ) else false end;
$$;
alter table public.article_master_revisions add constraint article_master_reusable_fields check (article_master_private.valid_data(data));

-- Defense in depth for maintenance mistakes; even linked drafts cannot be rewritten.
create function article_master_private.protect_revision()
returns trigger language plpgsql security definer set search_path = '' as $$
begin
  if tg_op='DELETE' then raise exception 'Article revisions are retained' using errcode='42501'; end if;
  if new.id<>old.id or new.article_id<>old.article_id or new.revision<>old.revision
    or new.created_by<>old.created_by or new.created_at<>old.created_at then
    raise exception 'Immutable revision identity' using errcode='42501';
  end if;
  if new.data is distinct from old.data and (old.status<>'draft' or exists(
    select 1 from public.article_master_ba_links l where l.revision_id=old.id)) then
    raise exception 'Frozen revision contents' using errcode='42501';
  end if;
  if (old.status='archived' and new.status<>old.status) or (old.status='released' and new.status not in ('released','archived')) then
    raise exception 'Invalid revision transition' using errcode='42501';
  end if;
  return new;
end; $$;
create trigger article_master_frozen_revision before update or delete on public.article_master_revisions
for each row execute function article_master_private.protect_revision();

create function article_master_private.capabilities()
returns jsonb language sql stable security definer set search_path = '' as $$
  select jsonb_build_object(
    'scopes',coalesce((select jsonb_agg(jsonb_build_object('department_id',s.department_id,'department_name',d.name))
      from public.article_master_scopes s join public.departments d on d.id=s.department_id
      where s.user_id=auth.uid() and d.active),'[]'::jsonb),
    'assignments',coalesce((select jsonb_agg(jsonb_build_object('department_id',a.department_id,'department_name',d.name,
      'article_number',a.article_number,'expires_at',a.expires_at)) from public.article_master_assignments a
      join public.departments d on d.id=a.department_id where a.user_id=auth.uid() and not a.revoked and a.expires_at>now() and d.active),'[]'::jsonb)
  );
$$;
create function article_master_private.list(p_query text)
returns jsonb language sql stable security definer set search_path = '' as $$
  select coalesce(jsonb_agg(result),'[]'::jsonb) from (
    select to_jsonb(r)||jsonb_build_object('department_id',a.department_id,'article_number',a.article_number) as result
    from public.article_master_revisions r join public.article_master_items a on a.id=r.article_id
    where article_master_private.allowed(a.department_id,a.article_number,'read',r.created_by)
      and a.number_key like replace(replace(replace(upper(trim(coalesce(p_query,''))),'\','\\'),'%','\%'),'_','\_')||'%' escape '\'
    order by a.number_key,r.revision desc limit 200
  ) q;
$$;
create function article_master_private.command(p_action text,p_revision_id uuid,p_department_id uuid,p_article_number text,p_data jsonb,p_expected_version integer)
returns jsonb language plpgsql security definer set search_path = '' as $$
declare
  a public.article_master_items; r public.article_master_revisions; result public.article_master_revisions;
  actor uuid := auth.uid(); next_revision integer;
begin
  if actor is null then raise exception 'Authentication required' using errcode='42501'; end if;
  if p_action is null or p_action not in ('create','update','revise','release','archive') then raise exception 'Unknown action' using errcode='22023'; end if;
  if p_action='create' then
    if not article_master_private.allowed(p_department_id,p_article_number,'create') then raise exception 'Denied' using errcode='42501'; end if;
    if not article_master_private.valid_data(p_data) or p_data is null then raise exception 'Invalid reusable fields' using errcode='22023'; end if;
    insert into public.article_master_items(department_id,article_number) values(p_department_id,trim(p_article_number)) returning * into a;
    insert into public.article_master_revisions(article_id,revision,data,created_by) values(a.id,1,p_data,actor) returning * into result;
  else
    -- Lock parent first for consistent lock order and serialized revision-number allocation.
    select i.* into a from public.article_master_items i join public.article_master_revisions v on v.article_id=i.id where v.id=p_revision_id for update of i;
    select * into r from public.article_master_revisions where id=p_revision_id for update;
    if r.id is null or not article_master_private.allowed(a.department_id,a.article_number,p_action,r.created_by) then
      raise exception 'Denied' using errcode='42501'; end if;
    if r.row_version is distinct from p_expected_version then raise exception 'Stale revision' using errcode='40001'; end if;
    if p_action='update' then
      if r.status<>'draft' or not article_master_private.valid_data(p_data) or p_data is null then raise exception 'Draft required' using errcode='22023'; end if;
      update public.article_master_revisions set data=p_data,row_version=row_version+1,updated_at=now() where id=r.id returning * into result;
    elsif p_action='revise' then
      if r.status<>'released' then raise exception 'Released source required' using errcode='22023'; end if;
      select coalesce(max(revision),0)+1 into next_revision from public.article_master_revisions where article_id=a.id;
      insert into public.article_master_revisions(article_id,revision,data,created_by) values(a.id,next_revision,r.data,actor) returning * into result;
    elsif p_action='release' then
      if r.status<>'draft' then raise exception 'Draft required' using errcode='22023'; end if;
      if r.created_by=actor and exists(select 1 from public.article_master_settings where department_id=a.department_id and require_second_reviewer) then
        raise exception 'Independent reviewer required' using errcode='42501'; end if;
      update public.article_master_revisions set status='released',released_by=actor,released_at=now(),updated_at=now(),row_version=row_version+1 where id=r.id returning * into result;
    elsif p_action='archive' then
      if r.status='archived' then raise exception 'Already archived' using errcode='22023'; end if;
      update public.article_master_revisions set status='archived',archived_by=actor,archived_at=now(),updated_at=now(),row_version=row_version+1 where id=r.id returning * into result;
    end if;
  end if;
  insert into public.article_master_audit(revision_id,actor_id,action,row_version) values(result.id,actor,p_action,result.row_version);
  return to_jsonb(result)||jsonb_build_object('department_id',a.department_id,'article_number',a.article_number);
end; $$;

-- Exposed API wrappers are invoker-only. Privileged implementation stays in unexposed schema.
create function public.article_master_capabilities() returns jsonb language sql stable security invoker set search_path = ''
as $$ select article_master_private.capabilities(); $$;
create function public.article_master_list(p_query text default '') returns jsonb language sql stable security invoker set search_path = ''
as $$ select article_master_private.list(p_query); $$;
create function public.article_master_command(p_action text,p_revision_id uuid default null,p_department_id uuid default null,p_article_number text default null,p_data jsonb default null,p_expected_version integer default null)
returns jsonb language sql security invoker set search_path = ''
as $$ select article_master_private.command(p_action,p_revision_id,p_department_id,p_article_number,p_data,p_expected_version); $$;

-- RLS everywhere, no client DML, no anonymous technical data, provisioning service-only.
alter table public.article_master_scopes enable row level security;
alter table public.article_master_settings enable row level security;
alter table public.article_master_assignments enable row level security;
alter table public.article_master_items enable row level security;
alter table public.article_master_revisions enable row level security;
alter table public.article_master_audit enable row level security;
alter table public.article_master_ba_links enable row level security;
revoke all on public.article_master_scopes,public.article_master_settings,public.article_master_assignments,
  public.article_master_items,public.article_master_revisions,public.article_master_audit,public.article_master_ba_links from public,anon,authenticated;
revoke all on all functions in schema article_master_private from public,anon,authenticated;
revoke all on function public.article_master_capabilities(),public.article_master_list(text),public.article_master_command(text,uuid,uuid,text,jsonb,integer) from public,anon;
grant execute on function article_master_private.capabilities(),article_master_private.list(text),article_master_private.command(text,uuid,uuid,text,jsonb,integer) to authenticated;
grant execute on function public.article_master_capabilities(),public.article_master_list(text),public.article_master_command(text,uuid,uuid,text,jsonb,integer) to authenticated;
-- Grants for allowed() / valid_data() used by constraints/policies, neither grants data mutation.
grant execute on function article_master_private.allowed(uuid,text,text,uuid),article_master_private.valid_data(jsonb) to authenticated;
create policy article_master_revisions_read on public.article_master_revisions for select to authenticated using (
  exists(select 1 from public.article_master_items a where a.id=article_id and article_master_private.allowed(a.department_id,a.article_number,'read',created_by))
);
-- Data API reads are deliberately via the scoped RPC only, no table SELECT grant.
-- Service role remains server-only for reviewed provisioning; no client/UI grant endpoint.
grant all on public.article_master_scopes,public.article_master_settings,public.article_master_assignments to service_role;
commit;
