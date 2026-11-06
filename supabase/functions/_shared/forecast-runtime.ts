// Review-only server entry dependencies. Nothing is deployed or scheduled by adding these files.
import { Temporal } from 'npm:@js-temporal/polyfill@0.5.1';
globalThis.HumbelTemporal = Temporal;
await import('../../../js/production-forecast.js');
await import('../../../js/forecast-service.js');
const url = Deno.env.get('SUPABASE_URL') || '';
const key = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') || '';
const anon = Deno.env.get('SUPABASE_ANON_KEY') || '';
export const enabled = () => Deno.env.get('HUMBEL_FORECAST_ENABLED') === 'true';
export const origins = (Deno.env.get('HUMBEL_FORECAST_ALLOWED_ORIGINS') || '').split(',').filter(Boolean);
export async function authenticate(authorization) {
 if (!authorization?.startsWith('Bearer ')) return null;
 const r = await fetch(url + '/auth/v1/user', {headers: {authorization, apikey: anon}});
 if (!r.ok) return null;
 const user = await r.json(); return typeof user.id === 'string' ? user.id : null;
}
export async function rpc(name, params, context) {
 const authorization = context.service ? 'Bearer ' + key : context.request?.headers.get('authorization') || '';
 const r = await fetch(url + '/rest/v1/rpc/' + name, {method:'POST',headers:{authorization,apikey:context.service?key:anon,'content-type':'application/json'},body:JSON.stringify(params)});
 const data = await r.json(); if (!r.ok) throw {code:data.code}; return data;
}
export const calculate = globalThis.HumbelProductionForecast.calculate;
export const service = globalThis.HumbelForecastService;
export async function workerAuthorized(request) {
 const expected = Deno.env.get('HUMBEL_FORECAST_WORKER_SECRET');
 const got = request.headers.get('x-forecast-worker-secret');
 if (!expected || !got) return false;
 const digest = (s) => crypto.subtle.digest('SHA-256',new TextEncoder().encode(s));
 const a = new Uint8Array(await digest(expected)), b = new Uint8Array(await digest(got));
 let diff=0; for(let i=0;i<a.length;i++)diff |= a[i]^b[i]; return diff===0;
}
