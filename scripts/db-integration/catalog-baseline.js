// A structural catalog mirror, NOT an authoritative pg_dump. No production connection.
const fs=require('node:fs');const crypto=require('node:crypto');
const inventory=JSON.parse(fs.readFileSync('docs/qa/schema-inventory.json','utf8'));
const quote=x=>'"'+String(x).replaceAll('"','""')+'"';
const helpers=`
-- TEST SUBSTITUTES: original public function bodies/ACLs were not captured.
create function public.set_updated_at() returns trigger language plpgsql as $$begin new.updated_at=now();return new;end$$;
create function public.is_admin() returns boolean language sql stable security definer set search_path='' as $$select exists(select 1 from public.employees where auth_user_id=auth.uid() and is_active and role='admin')$$;
create function public.current_app_employee_role() returns text language sql stable security definer set search_path='' as $$select role from public.employees where auth_user_id=auth.uid() and is_active order by id limit 1$$;
create function public.can_access_production_department(dept uuid) returns boolean language sql stable security definer set search_path='' as $$select exists(select 1 from public.employees e where e.auth_user_id=auth.uid() and e.is_active and e.role<>'tool_scanner' and(e.role='admin' or e.department_id=dept))$$;
`;
function sql(){const out=['begin;set local search_path=public,pg_catalog;'];
 for(const t of inventory.tables){const cols=inventory.columns.filter(c=>c.table===t.name);out.push(`create table public.${quote(t.name)}(${cols.map(c=>`${quote(c.name)} ${c.type}${c.default?' default '+c.default:''}${c.nullable==='NO'?' not null':''}`).join(',')});`);}
 for(const c of inventory.constraints.filter(c=>!c.definition.startsWith('FOREIGN KEY')))out.push(`alter table public.${quote(c.table)} add constraint ${quote(c.name)} ${c.definition};`);
 for(const c of inventory.constraints.filter(c=>c.definition.startsWith('FOREIGN KEY')))out.push(`alter table public.${quote(c.table)} add constraint ${quote(c.name)} ${c.definition};`);
 out.push(helpers);
 for(const p of inventory.policies)out.push(`create policy ${quote(p.policyname)} on public.${quote(p.tablename)} as ${p.permissive} for ${p.cmd} to ${p.roles.map(r=>r==='public'?'public':quote(r)).join(',')}${p.qual?' using ('+p.qual+')':''}${p.with_check?' with check ('+p.with_check+')':''};`);
 for(const t of inventory.tables){if(t.rls_enabled)out.push(`alter table public.${quote(t.name)} enable row level security;`);if(t.rls_forced)out.push(`alter table public.${quote(t.name)} force row level security;`);}
 for(const t of inventory.triggers)out.push(t.definition+';');
 // Only A/B table ACLs were verified by the supplied read-only report. Other ACLs are explicit test assumptions.
 out.push('grant usage on schema public to anon,authenticated,service_role;grant select,insert,update,delete on all tables in schema public to anon,authenticated;grant all on all tables in schema public to service_role;commit;');return out.join('\n');}
async function verify(db){
 // Stable catalog deparsing: auth must not become an implicitly visible namespace.
 await db.query("set search_path=public,pg_catalog");
 const tables=(await db.query("select tablename from pg_tables where schemaname='public' order by tablename")).rows.map(r=>r.tablename);
 if(JSON.stringify(tables)!==JSON.stringify(inventory.tables.map(t=>t.name).sort()))throw Error('31-table catalog mismatch');
 const cols=(await db.query("select table_name, column_name,data_type,is_nullable,column_default from information_schema.columns where table_schema='public' order by table_name,ordinal_position")).rows;
 if(cols.length!==inventory.columns.length)throw Error('308-column count mismatch');
 for(const c of inventory.columns){const r=cols.find(r=>r.table_name===c.table&&r.column_name===c.name);if(!r||r.data_type!==c.type||r.is_nullable!==c.nullable||r.column_default!==c.default)throw Error('Column metadata mismatch '+c.table+'.'+c.name);}
 const constraints=(await db.query("select c.conname,t.relname,pg_get_constraintdef(c.oid) definition from pg_constraint c join pg_class t on t.oid=c.conrelid join pg_namespace n on n.oid=t.relnamespace where n.nspname='public' and c.contype<>'n'")).rows;
 for(const c of inventory.constraints){const r=constraints.find(r=>r.conname===c.name&&r.relname===c.table);if(!r||r.definition.replaceAll('public.','')!==c.definition.replaceAll('public.',''))throw Error('Constraint mismatch '+c.name+': expected '+c.definition+'; actual '+r?.definition);}
 if(constraints.length!==inventory.constraints.length)throw Error('134-constraint count mismatch');
 const policies=(await db.query("select * from pg_policies where schemaname='public'")).rows;if(policies.length!==127)throw Error('127-policy count mismatch');
 for(const p of inventory.policies){const r=policies.find(r=>r.tablename===p.tablename&&r.policyname===p.policyname);for(const k of['cmd','permissive','qual','with_check'])if(r?.[k]!==p[k])throw Error('Policy mismatch '+p.policyname+' '+k);if(JSON.stringify(r.roles)!==JSON.stringify(p.roles))throw Error('Policy role mismatch');}
 const triggers=(await db.query("select t.tgname,c.relname,pg_get_triggerdef(t.oid,false)definition from pg_trigger t join pg_class c on c.oid=t.tgrelid join pg_namespace n on n.oid=c.relnamespace where n.nspname='public' and not t.tgisinternal")).rows;if(triggers.length!==11)throw Error('11-trigger count mismatch');
 for(const t of inventory.triggers)if(!triggers.some(r=>r.tgname===t.name&&r.relname===t.table&&r.definition===t.definition))throw Error('Trigger mismatch');
 const flags=(await db.query("select c.relname,c.relrowsecurity,c.relforcerowsecurity from pg_class c join pg_namespace n on n.oid=c.relnamespace where n.nspname='public'and c.relkind='r'")).rows;for(const t of inventory.tables){const r=flags.find(r=>r.relname===t.name);if(!r||r.relrowsecurity!==t.rls_enabled||r.relforcerowsecurity!==t.rls_forced)throw Error('RLS metadata mismatch '+t.name);}
 return{status:'STRUCTURAL_CATALOG_MIRROR_VERIFIED',authoritativeRestore:false,tables:31,columns:308,constraints:134,policies:127,triggers:11,inventory_sha256:crypto.createHash('sha256').update(fs.readFileSync('docs/qa/schema-inventory.json')).digest('hex'),missing:['Original 5 SECURITY DEFINER bodies/signatures/ACLs (3 inferred helpers substituted)','Original set_updated_at body substituted','Complete table/sequence/default ACLs and standalone indexes','Authoritative schema-only dump and newer catalog evidence']};
}
module.exports={sql,verify,inventory};
