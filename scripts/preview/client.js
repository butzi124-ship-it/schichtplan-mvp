// Browser-only legacy fixture state; NOT native RLS evidence or shared persistence.
const make=window.supabase.createClient;
window.supabase.createClient=(...args)=>({...make(...args),rpc:async(name,params={})=>{const r=await fetch((window.__PREVIEW_READONLY?'/documentation':'')+'/rpc',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,params})});if(r.status===401){location.reload();throw Error('Test session expired');}return r.json();}});
