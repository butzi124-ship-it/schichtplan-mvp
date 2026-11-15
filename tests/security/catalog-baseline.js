// A structural catalog mirror, NOT an authoritative pg_dump. No production connection.
const fs=require('node:fs');const crypto=require('node:crypto');
const inventory=JSON.parse(fs.readFileSync('docs/qa/schema-inventory.json','utf8'));
const quote=x=>'"'+String(x).replaceAll('"','""')+'"';
const {functions: verifiedFunctions,standaloneIndexes}=require('./source-catalog');
function sql(){const out=['begin;set local search_path=public,pg_catalog;'];
 for(const t of inventory.tables){const cols=inventory.columns.filter(c=>c.table===t.name);out.push(`create table public.${quote(t.name)}(${cols.map(c=>`${quote(c.name)} ${c.type}${c.default?' default '+c.default:''}${c.nullable==='NO'?' not null':''}`).join(',')});`);}
 for(const c of inventory.constraints.filter(c=>!c.definition.startsWith('FOREIGN KEY')))out.push(`alter table public.${quote(c.table)} add constraint ${quote(c.name)} ${c.definition};`);
 for(const c of inventory.constraints.filter(c=>c.definition.startsWith('FOREIGN KEY')))out.push(`alter table public.${quote(c.table)} add constraint ${quote(c.name)} ${c.definition};`);
 for(const i of standaloneIndexes)out.push(i.definition+';');
 for(const fn of verifiedFunctions){out.push(fn.definition+';');out.push('revoke execute on function public.'+fn.name+'('+fn.arguments+') from public;');out.push('grant execute on function public.'+fn.name+'('+fn.arguments+') to anon,authenticated,service_role;');}
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
 const policies=(await db.query("select policyname,tablename,cmd,permissive,qual,with_check,roles::text[] roles from pg_policies where schemaname='public'")).rows;if(policies.length!==127)throw Error('127-policy count mismatch');
 for(const p of inventory.policies){const r=policies.find(r=>r.tablename===p.tablename&&r.policyname===p.policyname);for(const k of['cmd','permissive','qual','with_check'])if(r?.[k]!==p[k])throw Error('Policy mismatch '+p.policyname+' '+k);if(JSON.stringify(r.roles)!==JSON.stringify(p.roles))throw Error('Policy role mismatch');}
 const triggers=(await db.query("select t.tgname,c.relname,pg_get_triggerdef(t.oid,false)definition from pg_trigger t join pg_class c on c.oid=t.tgrelid join pg_namespace n on n.oid=c.relnamespace where n.nspname='public' and not t.tgisinternal")).rows;if(triggers.length!==11)throw Error('11-trigger count mismatch');
 for(const t of inventory.triggers)if(!triggers.some(r=>r.tgname===t.name&&r.relname===t.table&&r.definition===t.definition))throw Error('Trigger mismatch');
 const idx=(await db.query("select ci.relname as name,pg_get_indexdef(i.indexrelid) definition from pg_index i join pg_class t on t.oid=i.indrelid join pg_class ci on ci.oid=i.indexrelid join pg_namespace n on n.oid=t.relnamespace where n.nspname='public'")).rows;
 if(idx.length!==76)throw Error('76-index live schema count mismatch: '+idx.length);
 for(const reference of standaloneIndexes){const found=idx.find(i=>i.name===reference.name);if(!found||found.definition!==reference.definition)throw Error('Live standalone index mismatch '+reference.name);}
 const fn=(await db.query("select p.proname name,pg_get_functiondef(p.oid) definition,p.prosecdef is_definer,coalesce(p.proconfig::text,'') settings from pg_proc p join pg_namespace n on n.oid=p.pronamespace where n.nspname='public'")).rows;
 if(fn.length!==6)throw Error('6-original-function count mismatch');
 for(const original of verifiedFunctions){const got=fn.find(x=>x.name===original.name);if(!got||got.is_definer!==original.security_definer||got.settings!==original.settings||got.definition.replaceAll('\\r\\n','\\n').trim()!==original.definition.replaceAll('\\r\\n','\\n').trim())throw Error('Original function metadata/definition mismatch '+original.name);}
 const flags=(await db.query("select c.relname,c.relrowsecurity,c.relforcerowsecurity from pg_class c join pg_namespace n on n.oid=c.relnamespace where n.nspname='public'and c.relkind='r'")).rows;for(const t of inventory.tables){const r=flags.find(r=>r.relname===t.name);if(!r||r.relrowsecurity!==t.rls_enabled||r.relforcerowsecurity!==t.rls_forced)throw Error('RLS metadata mismatch '+t.name);}
 return{status:'STRUCTURAL_CATALOG_MIRROR_VERIFIED',authoritativeRestore:false,tables:31,columns:308,constraints:134,policies:127,triggers:11,inventory_sha256:crypto.createHash('sha256').update(fs.readFileSync('docs/qa/schema-inventory.json')).digest('hex'),missing:['Original function ownership/grants/default ACLs are not restored exactly','Complete table and sequence/default ACLs remain assumptions','Authoritative schema-only dump/restore, Realtime and load tests remain outstanding']};
}
module.exports={sql,verify,inventory};
