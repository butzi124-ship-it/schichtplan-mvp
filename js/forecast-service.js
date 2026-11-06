/* Trusted server adapter. Never loaded by the application HTML. Credentials are injected server-side. */
(function(root){'use strict';
 function create({enabled,originAllowed,authenticate,rpc,calculate,uuid=()=>crypto.randomUUID()}){
  return async request=>{
   const origin=request.headers.get('origin');const headers={'content-type':'application/json','cache-control':'no-store','vary':'Origin'};
   if(origin){if(!originAllowed(origin))return new Response('{}',{status:403,headers});headers['access-control-allow-origin']=origin;headers['access-control-allow-headers']='authorization,apikey,content-type,x-client-info';headers['access-control-allow-methods']='POST,OPTIONS';}
   if(request.method==='OPTIONS')return new Response(null,{status:204,headers});
   const reply=(body,status=200)=>new Response(JSON.stringify(body),{status,headers});
   if(!enabled())return reply({error:'Forecast disabled'},403);if(request.method!=='POST')return reply({error:'POST required'},405);
   try{
    const actor=await authenticate(request.headers.get('authorization'));if(!actor)return reply({error:'Unauthorized'},401);
    const raw=await request.text();if(raw.length>12000)return reply({error:'Request too large'},413);const body=JSON.parse(raw);
    if(Object.keys(body).some(k=>!['department_id','request_id','report'].includes(k))||!/^[0-9a-f-]{36}$/i.test(body.department_id)||!/^[0-9a-f-]{36}$/i.test(body.request_id))return reply({error:'Invalid request'},400);
    if(body.report)await rpc('production_forecast_report',{p_operation_id:body.report.operation_id,p_request_id:body.report.request_id,p_data:body.report.data},{request});
    const input=await rpc('production_forecast_input',{p_department_id:body.department_id,p_request_id:body.request_id},{request});
    if(input.existing_run)return reply({run_id:input.existing_run,duplicate:true});
    const result=calculate(input);const run_id=await rpc('production_forecast_publish',{p_department_id:body.department_id,p_actor_id:actor,p_run_id:body.request_id,p_result:result},{service:true});
    return reply({run_id});
   }catch(e){const code=e.code;return reply({error:code==='40001'?'Input changed; recalculate':code==='42501'?'Denied':'Calculation unavailable'},code==='40001'?409:code==='42501'?403:400);}
  };
 }
 async function worker({enabled,rpc,calculate,uuid=()=>crypto.randomUUID()}){
  if(!enabled())throw Error('Forecast disabled');const pending=await rpc('production_forecast_pending',{}, {service:true}),out=[];
  for(const item of pending){try{const input=await rpc('production_forecast_worker_input',{p_department_id:item.department_id,p_actor_id:item.calculation_actor_id},{service:true});const result=calculate(input);const id=uuid();await rpc('production_forecast_publish',{p_department_id:item.department_id,p_actor_id:item.calculation_actor_id,p_run_id:id,p_result:result},{service:true});out.push({department_id:item.department_id,status:'calculated',run_id:id});}catch{out.push({department_id:item.department_id,status:'blocked'});}}return out;
 }
 root.HumbelForecastService={create,worker};if(typeof module!=='undefined')module.exports=root.HumbelForecastService;
})(globalThis);
