import {enabled,rpc,calculate,service,workerAuthorized} from '../_shared/forecast-runtime.ts';
Deno.serve(async (request) => {
 if(request.method!=='POST'||!enabled()||!await workerAuthorized(request))return new Response(null,{status:403});
 try{return Response.json({results:await service.worker({enabled,rpc,calculate})});}
 catch{return Response.json({error:'Worker unavailable'},{status:503});}
});
