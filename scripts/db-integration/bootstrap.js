const fs=require('node:fs');const {Client}=require('pg');const {guard}=require('./guard');
async function main(){const config=guard(JSON.parse(fs.readFileSync('/tmp/humbel-isolated-db/runtime.json','utf8')));
 const db=new Client(config);await db.connect();try{
 const identity=(await db.query('select current_database() name')).rows[0];if(identity.name!==config.database)throw Error('Wrong isolated database');
 await db.query(`alter database humbel_isolated_test set "humbel.isolated_sentinel"='${config.sentinel}'`);
 await db.query(`do $$begin if not exists(select 1 from pg_roles where rolname='anon') then create role anon nologin;end if;if not exists(select 1 from pg_roles where rolname='authenticated') then create role authenticated nologin;end if;if not exists(select 1 from pg_roles where rolname='service_role') then create role service_role nologin bypassrls;end if;if not exists(select 1 from pg_roles where rolname='authenticator') then create role authenticator login noinherit;end if;if not exists(select 1 from pg_roles where rolname='supabase_auth_admin') then create role supabase_auth_admin login;end if;end$$;grant anon,authenticated,service_role to authenticator;`);
 await db.query(`alter role authenticator password '${config.password}';alter role supabase_auth_admin password '${config.password}';create schema if not exists auth authorization supabase_auth_admin;grant all on schema auth to supabase_auth_admin;grant create on database humbel_isolated_test to supabase_auth_admin;`);
 await db.query('grant usage on schema auth to anon,authenticated,service_role');
 // Auth accessors are supplied by the real Supabase image / GoTrue migrations, never test substitutes.

 }finally{await db.end();}

}
main().then(()=>console.log('Container-local database bootstrap passed')).catch(e=>{console.error('Bootstrap failed '+(e.code||e.message));process.exitCode=1;});
