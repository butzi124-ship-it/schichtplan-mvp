import {enabled,origins,authenticate,rpc,calculate,service} from '../_shared/forecast-runtime.ts';
Deno.serve(service.create({enabled,originAllowed:(origin)=>origins.includes(origin),authenticate,rpc,calculate}));
