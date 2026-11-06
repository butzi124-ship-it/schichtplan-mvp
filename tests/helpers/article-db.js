// Disposable in-memory PostgreSQL only. No URLs, clients, live credentials or exported rows.
const { PGlite } = require('@electric-sql/pglite');
const fs = require('node:fs');
const path = require('node:path');
const manager = '00000000-0000-0000-0000-000000000001';
const worker = '00000000-0000-0000-0000-000000000002';
const outsider = '00000000-0000-0000-0000-000000000003';
const reviewer = '00000000-0000-0000-0000-000000000004';
const department = '10000000-0000-0000-0000-000000000001';
const foreignDepartment = '10000000-0000-0000-0000-000000000002';
const order = '20000000-0000-0000-0000-000000000001';
const migration = path.resolve(__dirname, '../../supabase/migrations/20261106023103_article_master_phase1.sql');
async function articleDatabase() {
  const db = new PGlite();
  await db.exec(`create role anon; create role authenticated; create role service_role bypassrls;
    create schema auth; create table auth.users(id uuid primary key);
    create function auth.uid() returns uuid language sql stable as $$select nullif(current_setting('request.jwt.claim.sub',true),'')::uuid$$;
    grant usage on schema auth to authenticated,anon; grant execute on function auth.uid() to authenticated,anon;
    create table public.departments(id uuid primary key, name text not null, active boolean not null default true);
    create table public.production_orders(id uuid primary key,ba_number text,ba_quantity integer);
    insert into auth.users values ('${manager}'),('${worker}'),('${outsider}'),('${reviewer}');
    insert into public.departments(id,name) values ('${department}','Fertigung'),('${foreignDepartment}','Andere Abteilung');
    insert into public.production_orders values ('${order}','812150',48);`);
  await db.exec(fs.readFileSync(migration, 'utf8'));
  await db.query('insert into public.article_master_scopes values ($1,$2),($3,$2)', [manager,department,reviewer]);
  await db.query(`insert into public.article_master_assignments(user_id,department_id,article_number,first_op_number,first_op_evidence,expires_at,issued_by)
    values($1,$2,'32505',10,'synthetic first-OP assignment; no live BA',now()+interval '1 hour',$3)`,[worker,department,manager]);
  async function asUser(uid, fn, dbRole='authenticated') {
    await db.query("select set_config('request.jwt.claim.sub',$1,false)",[uid || '']);
    await db.exec(`set role ${dbRole}`);
    try { return await fn(); } finally { await db.exec('reset role'); }
  }
  async function rpc(uid, name, params={}) {
    const signatures = {
      article_master_capabilities: [], article_master_list: ['p_query'],
      article_master_command: ['p_action','p_revision_id','p_department_id','p_article_number','p_data','p_expected_version'],
    };
    if (!Object.hasOwn(signatures,name)) throw Error('Unexpected RPC');
    const keys=signatures[name];
    const result=await asUser(uid,()=>db.query(`select public.${name}(${keys.map((_,i)=>'$'+(i+1)).join(',')}) as result`,keys.map(k=>params[k]??null)));
    return result.rows[0].result;
  }
  return { db, rpc, asUser };
}
module.exports = { articleDatabase, manager,worker,outsider,reviewer,department,foreignDepartment,order };
