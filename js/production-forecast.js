/* Pure, bounded, deterministic forecast engine. No network, storage, names or medical reasons. */
(function(root){
 'use strict';
 const Temporal=root.HumbelTemporal||(typeof require==='function'?require('@js-temporal/polyfill').Temporal:null);
 const fail=m=>{throw Error(m);};
 const finite=(n,min=0,max=1e12)=>typeof n==='number'&&Number.isFinite(n)&&n>=min&&n<=max;
 const instant=v=>{if(typeof v!=='string'||!/(Z|[+-]\d\d:\d\d)$/.test(v))fail('Zeitpunkt mit expliziter Zeitzone fehlt');return Temporal.Instant.from(v).epochMilliseconds;};
 const iso=ms=>Temporal.Instant.fromEpochMilliseconds(Math.round(ms)).toString();
 function seconds(time){if(!time||!finite(time.value)||!Object.hasOwn({s:1,min:60,h:3600},time.unit))fail('Tr/Te mit expliziter Einheit fehlen');const s=time.value*({s:1,min:60,h:3600}[time.unit]);if(!Number.isSafeInteger(Math.round(s*1000))||s>0&&s<0.001)fail('Zeitwert außerhalb Millisekundenauflösung');return Math.round(s*1000)/1000;}
 function remainder(job,asOf,maxAge){
  const a=job.actual||{};if(job.work_status==='completed'){if(!job.completed_at)fail('Abschlusszeit nicht belegt');return{setup_s:0,processing_s:0,work_s:0,remaining_pieces:0,variance_s:null};}
  const tr=seconds(job.tr),te=seconds(job.te);
  if(!a.observed_at||instant(asOf)-instant(a.observed_at)>maxAge*1000||instant(a.observed_at)>instant(asOf))fail('Ist-Daten fehlen oder sind veraltet');
  for(const k of ['setup_elapsed_s','processing_elapsed_s','current_piece_elapsed_s','good','scrap','clarify','replacement_qty'])if(!finite(a[k]))fail('Ist-Wert fehlt: '+k);
  for(const k of ['good','scrap','clarify','replacement_qty'])if(!Number.isSafeInteger(a[k]))fail('Stückzahlen müssen ganzzahlig sein');
  if(!Number.isSafeInteger(job.quantity)||job.quantity<1||a.good>job.quantity||a.good+a.scrap+a.clarify>job.quantity+a.replacement_qty)fail('Mengen widersprüchlich');
  if(a.clarify>0)fail('Ungeklärte Mengen: Disposition fehlt');
  const remaining=job.quantity-a.good-a.scrap-a.clarify+a.replacement_qty;
  if(a.good+remaining<job.quantity)fail('Ausschuss ohne bestätigte Ersatzmenge');
  if(a.current_piece_elapsed_s>te||a.current_piece_elapsed_s>a.processing_elapsed_s||remaining===0&&a.current_piece_elapsed_s>0)fail('Angebrochene Bearbeitung nicht plausibel');
  let setup;
  if(a.setup_complete===true)setup=0;
  else if(finite(a.setup_remaining_s))setup=a.setup_remaining_s;
  else {if(a.setup_complete!==false)fail('Rüststatus fehlt');if(tr>0&&a.setup_elapsed_s>=tr)fail('Rüst-Soll überschritten: offene Rüstzeit bestätigen');setup=Math.max(0,tr-a.setup_elapsed_s);}
  const processing=Math.max(0,remaining*te-a.current_piece_elapsed_s);
  return{setup_s:setup,processing_s:processing,work_s:setup+processing,remaining_pieces:remaining,variance_s:a.processing_elapsed_s-(a.good+a.scrap+a.clarify)*te-a.current_piece_elapsed_s};
 }
 function local(date,time,zone,policy){return Temporal.ZonedDateTime.from(Temporal.PlainDate.from(date).toPlainDateTime(Temporal.PlainTime.from(time)).toString()+'['+zone+']',{disambiguation:policy}).epochMilliseconds;}
 const overlap=(a,b)=>a.start<b.end&&b.start<a.end;
 function subtract(interval,cuts){let parts=[interval];for(const c of cuts){parts=parts.flatMap(p=>!overlap(p,c)?[p]:[{...p,end:Math.min(p.end,c.start)},{...p,start:Math.max(p.start,c.end)}].filter(x=>x.end>x.start));}return parts;}
 function calendar(input){
  const c=input.config;if(!c||!['Europe/Berlin','Europe/Zurich'].includes(c.zone)||!['earlier','later','reject'].includes(c.dst_disambiguation))fail('Kalender/Zeitzonenregel fehlt');
  const start=Temporal.PlainDate.from(c.start_date),end=Temporal.PlainDate.from(c.end_date),days=start.until(end).days;
  if(days<0||days>62||!finite(c.green_buffer_seconds)||!finite(c.orange_buffer_seconds)||c.orange_buffer_seconds>c.green_buffer_seconds||!finite(c.feedback_max_age_seconds,1)||!finite(c.max_age_seconds,1,3600))fail('Prognoseregeln fehlen oder sind ungültig');
  if(instant(c.valid_until)<=instant(input.as_of))fail('Datenfreigabe abgelaufen');
  const until=local(end.add({days:1}).toString(),'00:00',c.zone,c.dst_disambiguation),from=local(start.toString(),'00:00',c.zone,c.dst_disambiguation),now=instant(input.as_of);
  if(now<from||now>=until)fail('Berechnungszeit außerhalb bestätigtem Kalender');
  const closed=(input.closed_dates||[]).map(d=>({start:local(d,'00:00',c.zone,c.dst_disambiguation),end:local(Temporal.PlainDate.from(d).add({days:1}).toString(),'00:00',c.zone,c.dst_disambiguation)}));
  const windows=[],staffWindows=[];
  for(const m of input.machines){if(!Number.isInteger(m.capacity_units)||m.capacity_units<1||m.capacity_units>16||typeof m.unmanned_capable!=='boolean')fail('Maschinenkapazität/Fähigkeit nicht konfiguriert');}
  for(const rule of input.shifts){
   if(!Array.isArray(rule.weekdays)||!rule.weekdays.length||!Array.isArray(rule.breaks)||typeof rule.unmanned!=='boolean'||!rule.suffix||!rule.start_time||!rule.end_time||!rule.machine_id)fail('Schichtzuordnung unvollständig');
   for(let d=start;Temporal.PlainDate.compare(d,end)<=0;d=d.add({days:1})){
    if(!rule.weekdays.includes(d.dayOfWeek)||rule.excluded_dates?.includes(d.toString()))continue;
    const day=d.toString(),next=d.add({days:1}).toString();let s=local(day,rule.start_time,c.zone,c.dst_disambiguation),e=local(rule.end_time<=rule.start_time?next:day,rule.end_time,c.zone,c.dst_disambiguation);
    if(e<=s)fail('Schicht hat keine reale Dauer');
    const breaks=rule.breaks.map(b=>{const bs=local(b.start<rule.start_time?next:day,b.start,c.zone,c.dst_disambiguation),be=local(b.end<=b.start||b.end<rule.start_time?next:day,b.end,c.zone,c.dst_disambiguation);if(bs<s||be>e||be<=bs)fail('Pause liegt außerhalb Schicht');return{start:bs,end:be};});
    const pieces=subtract({start:s,end:e,machine_id:rule.machine_id,unmanned:rule.unmanned},[...breaks,...closed]);windows.push(...pieces);
    if(!rule.unmanned)for(const person of input.staff){
     if(!person.machine_ids.includes(rule.machine_id)||!person.assignments.some(a=>a.date===day&&a.suffix===rule.suffix)||person.unavailable_shifts?.includes(day+rule.suffix))continue;
     const absent=person.unavailable_dates.map(ad=>({start:local(ad,'00:00',c.zone,c.dst_disambiguation),end:local(Temporal.PlainDate.from(ad).add({days:1}).toString(),'00:00',c.zone,c.dst_disambiguation)}));
     for(const p of pieces)staffWindows.push(...subtract({...p,person:person.id,qualifications:person.qualifications},absent));
    }
   }
  }
  return{windows,staffWindows,from,until,now};
 }
 function calculate(input){
  if(!Temporal)fail('Temporal nicht verfügbar');
  if(!input||input.schema!==1||!Array.isArray(input.jobs)||input.jobs.length>200||!Array.isArray(input.staff)||input.staff.length>128||!Array.isArray(input.machines)||input.machines.length>64||!Array.isArray(input.shifts)||input.shifts.length>512)fail('Ungültige oder zu große Datenbasis');
  const jobs=input.jobs.map(j=>({...j})),byId=new Map(jobs.map(j=>[j.id,j]));if(byId.size!==jobs.length)fail('Doppelte OP-Identität');
  const errors=new Map(jobs.map(j=>[j.id,[]]));let cal;
  try{cal=calendar(input);}catch(e){for(const j of jobs)errors.get(j.id).push(e.message);}
  const machines=new Map(input.machines.map(m=>[m.id,m]));
  for(const j of jobs){try{
    if(!j.id||!machines.has(j.machine_id)||!machines.get(j.machine_id).active)fail('Maschine fehlt oder ist inaktiv');
    if(!Number.isInteger(j.priority)||!Number.isInteger(j.people)||j.people<1||j.people>16||!j.qualification||typeof j.unmanned_allowed!=='boolean')fail('OP-Kapazitäts-/Qualifikationsregeln fehlen');
    j.rest=remainder(j,input.as_of,input.config?.feedback_max_age_seconds);j.due=instant(j.due_at);
    if(cal&&j.due>cal.until)fail('Bestätigter Kalender reicht nicht bis Endtermin');
    if(j.predecessor_id&&!byId.has(j.predecessor_id))fail('Vorgänger fehlt');
    if(j.order_status==='paused'||j.work_status==='cancelled'||j.order_status==='cancelled')fail('BA/OP gesperrt, Wiederaufnahme nicht belegt');
    if(j.work_status==='paused'&&!j.wait_until)fail('Unterbrechung ohne bestätigten Wiederaufnahmezeitpunkt');
    if(j.planned_start)instant(j.planned_start);if(j.wait_until)instant(j.wait_until);
    if(j.work_status!=='completed'&&!j.predecessor_id&&j.material_ready!==true&&!j.actual?.material_available_at)fail('Materialverfügbarkeit für erste OP fehlt');
    if(j.predecessor_id&&(!finite(j.transfer_seconds)))fail('Dokumentierte Übergabevorlaufzeit fehlt');
   }catch(e){errors.get(j.id).push(e.message);}}
  for(const issue of input.issues||[])for(const j of jobs)if(!issue.machine_id||j.machine_id===issue.machine_id)errors.get(j.id).push(issue.reason);
  // Unknown workload contaminates any connected machine/personnel/dependency component; never assume it consumes zero capacity.
  const share=(a,b)=>a.machine_id===b.machine_id||a.predecessor_id===b.id||b.predecessor_id===a.id||input.staff.some(p=>p.machine_ids.includes(a.machine_id)&&p.machine_ids.includes(b.machine_id));
  let changed=true;while(changed){changed=false;for(const a of jobs)if(errors.get(a.id).length)for(const b of jobs)if(!errors.get(b.id).length&&share(a,b)){errors.get(b.id).push('Unvollständige konkurrierende OP: '+(a.op_number?String(a.op_number).padStart(3,'0'):'gemeinsame Kapazität'));changed=true;}}
  const result=new Map(),bookings=[],staffBookings=[];
  const unknown=j=>({operation_id:j.id,order_id:j.order_id,color:'unknown',status:'incomplete',finish_at:null,arrival_at:null,arrival_today:false,work_seconds:j.rest?.work_s??null,capacity_until_due_seconds:null,buffer_seconds:null,issues:errors.get(j.id),allocations:[]});
  for(const j of jobs)if(errors.get(j.id).length)result.set(j.id,unknown(j));
  const windows=cal?.windows||[],sw=cal?.staffWindows||[];
  const blocks=(input.blocks||[]).map(b=>({...b,start:instant(b.start),end:instant(b.end)}));
  for(const b of blocks)if(b.end<=b.start||!Number.isInteger(b.units)||b.units<1)fail('Stillstand/Reservierung ungültig');
  function slots(job,from,to,phase){
   const wm=windows.filter(w=>w.machine_id===job.machine_id&&(!w.unmanned||phase==='process'&&job.unmanned_allowed&&machines.get(job.machine_id).unmanned_capable));
   const sm=sw.filter(w=>w.machine_id===job.machine_id&&w.qualifications.includes(job.qualification));
   const relevant=[...wm,...sm,...blocks.filter(b=>b.machine_id===job.machine_id),...bookings.filter(b=>b.machine_id===job.machine_id),...staffBookings];
   const edges=[...new Set([from,to,...relevant.flatMap(w=>[Math.max(from,w.start),Math.min(to,w.end)]).filter(t=>t>from&&t<to)])].sort((a,b)=>a-b),out=[];
   for(let i=0;i<edges.length-1;i++){
    const start=edges[i],end=edges[i+1],span={start,end};const covering=wm.filter(w=>w.start<=start&&w.end>=end);if(!covering.length)continue;
    const used=[...blocks,...bookings].filter(b=>b.machine_id===job.machine_id&&overlap(b,span)).reduce((n,b)=>n+(b.units||1),0);if(used>=machines.get(job.machine_id).capacity_units)continue;
    let people=[];
    const unmanned=phase==='process'&&job.unmanned_allowed&&machines.get(job.machine_id).unmanned_capable&&covering.some(w=>w.unmanned);
    if(!unmanned){people=[...new Set(sm.filter(w=>w.start<=start&&w.end>=end&&!staffBookings.some(b=>b.person===w.person&&overlap(b,span))).map(w=>w.person))].sort();if(people.length<job.people)continue;people=people.slice(0,job.people);}
    out.push({start,end,people,unmanned});
   }return out;
  }
  function capacity(j,from,to,phase){return to<=from?0:slots(j,from,to,phase).reduce((n,s)=>n+(s.end-s.start)/1000,0);}
  function allocate(j,from,seconds,phase){let left=Math.round(seconds*1000),cursor=from;const allocations=[];if(left===0)return{finish:cursor,allocations};for(const s of slots(j,from,cal.until,phase)){const amount=Math.min(left,s.end-s.start),end=s.start+amount;const b={machine_id:j.machine_id,start:s.start,end,units:1};bookings.push(b);for(const person of s.people)staffBookings.push({person,start:s.start,end});allocations.push({start:iso(s.start),end:iso(end),phase,unmanned:s.unmanned});left-=amount;cursor=end;if(left===0)return{finish:cursor,allocations};}return{finish:null,allocations};}
  let pending=jobs.filter(j=>!result.has(j.id));
  while(pending.length){
   const eligible=pending.filter(j=>!j.predecessor_id||result.has(j.predecessor_id)).sort((a,b)=>(a.work_status==='running'?-1:0)-(b.work_status==='running'?-1:0)||a.priority-b.priority||(a.planned_start||'').localeCompare(b.planned_start||'')||a.id.localeCompare(b.id));
   if(!eligible.length){for(const j of pending){errors.get(j.id).push('Zyklische OP-Abhängigkeit');result.set(j.id,unknown(j));}break;}
   const j=eligible[0];pending=pending.filter(x=>x.id!==j.id);let release=cal.now,arrival=null;
   if(j.predecessor_id){const pred=result.get(j.predecessor_id);if(!pred.finish_at){errors.get(j.id).push('Vorgänger nicht verlässlich fertigstellbar');result.set(j.id,unknown(j));continue;}arrival=instant(pred.finish_at)+j.transfer_seconds*1000;release=Math.max(release,arrival);}
   else if(j.actual?.material_available_at)release=Math.max(release,instant(j.actual.material_available_at));
   if(j.planned_start)release=Math.max(release,instant(j.planned_start));if(j.wait_until)release=Math.max(release,instant(j.wait_until));
   if(j.work_status==='completed'){const finish=instant(j.completed_at);result.set(j.id,{operation_id:j.id,order_id:j.order_id,status:'completed',color:finish>j.due?'red':'green',finish_at:iso(finish),arrival_at:null,arrival_today:false,work_seconds:0,capacity_until_due_seconds:0,buffer_seconds:0,issues:[],allocations:[]});continue;}
   const available=capacity(j,release,j.due,'process');const setup=allocate(j,release,j.rest.setup_s,'setup');const process=setup.finish===null?{finish:null,allocations:[]}:allocate(j,setup.finish,j.rest.processing_s,'process');const finish=process.finish;
   const buffer=finish===null?null:finish<=j.due?capacity(j,finish,j.due,'process'):-capacity(j,j.due,finish,'process');
   const color=finish===null||finish>j.due?'red':buffer>=input.config.green_buffer_seconds?'green':'orange';
   const day=arrival===null?null:Temporal.Instant.fromEpochMilliseconds(Math.round(arrival)).toZonedDateTimeISO(input.config.zone).toPlainDate().toString();
   const today=Temporal.Instant.from(input.as_of).toZonedDateTimeISO(input.config.zone).toPlainDate().toString();
   const workingToday=arrival!==null&&windows.some(w=>w.machine_id===j.machine_id&&Temporal.Instant.fromEpochMilliseconds(w.start).toZonedDateTimeISO(input.config.zone).toPlainDate().toString()===today&&w.end>arrival);
   result.set(j.id,{operation_id:j.id,order_id:j.order_id,status:finish===null?'capacity_exhausted':'calculated',color,finish_at:finish===null?null:iso(finish),arrival_at:arrival===null?null:iso(arrival),arrival_today:day===today&&workingToday,work_seconds:j.rest.work_s,remaining_setup_seconds:j.rest.setup_s,remaining_processing_seconds:j.rest.processing_s,remaining_pieces:j.rest.remaining_pieces,recorded_variance_seconds:j.rest.variance_s,capacity_until_due_seconds:available,buffer_seconds:buffer,critical_reserve:buffer!==null&&buffer<input.config.orange_buffer_seconds,bottleneck:j.machine_id,issues:finish===null?['Bestätigte Maschinen-/Personalkapazität im Horizont erschöpft']:[],allocations:[...setup.allocations,...process.allocations]});
  }
  const results=jobs.map(j=>result.get(j.id)),orders=[];for(const id of [...new Set(jobs.map(j=>j.order_id))]){const rr=results.filter(r=>r.order_id===id),finish=rr.every(r=>r.finish_at)?rr.map(r=>r.finish_at).sort((a,b)=>instant(a)-instant(b)).at(-1):null;orders.push({order_id:id,color:rr.some(r=>r.color==='unknown')?'unknown':rr.some(r=>r.color==='red')?'red':rr.some(r=>r.color==='orange')?'orange':'green',finish_at:finish,work_seconds:rr.every(r=>r.work_seconds!==null)?rr.reduce((n,r)=>n+r.work_seconds,0):null,bottleneck:rr.find(r=>['unknown','red','orange'].includes(r.color))?.operation_id||null});}
  return{schema:1,input_token:input.input_token,department_id:input.department_id,as_of:input.as_of,zone:input.config?.zone||null,results,orders};
 }
 const api={seconds,remainder,calendar,calculate};root.HumbelProductionForecast=api;if(typeof module!=='undefined')module.exports=api;
})(globalThis);
