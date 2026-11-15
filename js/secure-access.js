/* Isolated integration adapter. Not loaded by the production HTML; default disabled.
   No storage, offline fallback, local role authority, or legacy table-write fallback. */
(function(root){
 'use strict';
 function create({client,enabled=()=>false,online=()=>root.navigator?.onLine!==false,events=root,cleared=()=>{}}){
  let generation=0,disposed=false;
  let data={identity:null,schedule:[],contacts:null,materials:[]};
  const clone=value=>JSON.parse(JSON.stringify(value));
  function reset(){const hadData=data.identity!==null||data.schedule.length>0||data.contacts!==null||data.materials.length>0;generation++;data={identity:null,schedule:[],contacts:null,materials:[]};if(hadData)cleared();}
  const change=()=>reset();
  const subscription=client.auth.onAuthStateChange(change)?.data?.subscription;
  events?.addEventListener?.('offline',change);
  function check(token){if(disposed||token!==generation)throw Error('Sicherer Zugriff verworfen: Session geändert.');if(!enabled()||!online()){reset();throw Error('Sicherer Zugriff verworfen: Freigabe oder Verbindung geändert.');}}
  async function call(name,args,token){check(token);const result=await client.rpc(name,args);check(token);if(result.error)throw Error('Sicherer Serverzugriff abgewiesen.');return result.data;}
  async function identity(token){
   check(token);const response=await client.auth.getUser();check(token);
   if(response.error||!response.data?.user?.id)throw Error('Keine verifizierte Session.');
   const context=await call('security_actor_context',{},token);
   if(!context||context.user_id!==response.data.user.id||!context.employee_id||!['admin','department_admin','employee','tool_scanner'].includes(context.role)
    ||!Array.isArray(context.plan_departments)||!Array.isArray(context.contact_fields))throw Error('Widersprüchliche Serveridentität.');
   return context;
  }
  async function operation(work){
   const token=generation;
   try{const who=await identity(token);const result=await work(who,token);check(token);data.identity=clone(who);return clone(result);}
   catch(error){if(token===generation)reset();throw error;}
  }
  function getState(){if(disposed||!enabled()||!online())reset();return clone(data);}
  function refreshIdentity(){return operation(who=>who);}
  function readSchedule(){return operation(async(who,token)=>{
   if(who.role==='tool_scanner')throw Error('Schichtzugriff nicht vorgesehen.');
   const rows=await call('security_published_schedule',{},token);
   if(!Array.isArray(rows))throw Error('Ungültige Schichtantwort.');
   const allowed=['assignment_id','shift_id','shift_date','employee_id'];
   if(rows.some(row=>!row||Object.keys(row).some(k=>!allowed.includes(k))||allowed.some(k=>!Object.hasOwn(row,k))))throw Error('Unerwartete Schichtfelder.');
   data.schedule=clone(rows);return rows;
  });}
  function ownContacts(values=null){return operation(async(who,token)=>{
   if(who.role==='tool_scanner')throw Error('Kontaktzugriff nicht vorgesehen.');
   if(values!==null&&(typeof values!=='object'||Array.isArray(values)||Object.keys(values).some(k=>!who.contact_fields.includes(k))))throw Error('Kontaktfeld nicht freigegeben.');
   const result=await call('security_own_contacts',{p_values:values},token);
   if(!result||typeof result!=='object'||Array.isArray(result)||Object.keys(result).some(k=>!who.contact_fields.includes(k)||typeof result[k]!=='string'))throw Error('Ungültige Kontaktantwort.');
   data.contacts=clone(result);return result;
  });}
  function assign({id,department,subject,shiftId,date,expectedVersion}){return operation(async(who,token)=>{
   if(!who.plan_departments.includes(department))throw Error('Planungsbereich nicht freigegeben.');
   const result=await call('security_assignment_command',{p_id:id,p_department:department,p_subject:subject,p_shift_id:shiftId,p_date:date,p_expected_version:expectedVersion},token);
   data.schedule=[];return result;
  });}
  function publish({id,expectedVersion}){return operation(async(who,token)=>{
   // Publication is not inferred from role or plan_departments; SQL checks its distinct grant.
   await call('security_publish_assignment',{p_id:id,p_expected_version:expectedVersion},token);data.schedule=[];return true;
  });}
  function readMaterials(){return operation(async(who,token)=>{
   check(token);const result=await client.from('tool_materials').select('id,name,sort_order,is_active').eq('is_active',true).order('sort_order',{ascending:true}).order('name',{ascending:true});check(token);
   if(result.error||!Array.isArray(result.data))throw Error('Materialreferenzen nicht verfügbar.');
   const allowed=['id','name','sort_order','is_active'];
   if(result.data.some(row=>Object.keys(row).some(k=>!allowed.includes(k))))throw Error('Unerwartete Materialfelder.');
   data.materials=clone(result.data);return result.data;
  });}
  function dispose(){reset();disposed=true;subscription?.unsubscribe();events?.removeEventListener?.('offline',change);}
  return{reset,dispose,getState,refreshIdentity,readSchedule,ownContacts,assign,publish,readMaterials};
 }
 root.HumbelSecureAccess={create};if(typeof module!=='undefined')module.exports=root.HumbelSecureAccess;
})(typeof window==='undefined'?globalThis:window);
