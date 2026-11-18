'use strict';
// FABOS-only synthetic preview data. No production database connection.
const BASE = '2026-11-18T08:00:00.000Z'; // Wednesday 09:00 CET, reproducible demo reference
const ts = t => new Date(t).toISOString();
const depSpec = [
  ['DR','Dreherei'],['FR','Fräserei'],['VZ','Verzahnerei'],['EN','Entgraterei'],
  ['HT','Härterei'],['SL','Schleiferei'],['QS','Qualitätssicherung'],['MO','Montage']
];
const nameBank = [
  ['Nina','Reuter'],['Emil','Krämer'],['Mara','Seidel'],['Felix','Binder'],['Leonie','Hartmann'],
  ['Jonas','Lindner'],['Mila','Fischer'],['Timo','Bachmann'],['Paula','Renz'],['Theo','Neumann'],
  ['Elisa','Bergmann'],['Noah','Ahrens'],['Amira','Keller'],['Ben','Sommer'],['Nora','Wiegand'],
  ['Mika','Scholz'],['Greta','Wolff'],['Jannis','Albrecht'],['Luisa','Voss'],['Erik','Horn'],
  ['Clara','Jäger'],['David','Schuster'],['Anika','Maurer'],['Lennart','Kruse'],['Jule','Henning'],
  ['Leon','Bühler'],['Zoe','Kunert'],['Max','Roth'],['Selina','Brandt'],['Oskar','Frey'],
  ['Ada','Hartwig'],['Robin','Walter'],['Lia','Petersen'],['Sven','Kopp'],['Helena','Graf'],
  ['Fabian','Beck'],['Alina','Riedel'],['Niklas','Urban'],['Sina','Winter'],['Anton','Döring'],
  ['Laura','Meier'],['Tim','Schreiber'],['Sophie','Winkler'],['Paul','Franke'],['Melina','Voigt'],
  ['Luca','Reich'],['Hannah','Kuhn'],['Moritz','Fuchs'],['Jana','Lorenz'],['Simon','Koch']
];
const machinesSpec = [
  ['DR','M136','Maschine 136','known'],['DR','DEMO-D02','DEMO Drehzentrum D02','synthetic'],['DR','DEMO-D03','DEMO Drehzentrum D03','synthetic'],
  ['FR','M50','Maschine 50','known'],['FR','C62','Hermle C62','known'],['FR','C22','Hermle C22','known'],['FR','DEMO-F04','DEMO Fräszentrum F04','synthetic'],
  ['VZ','DEMO-V01','DEMO Wälzfräsmaschine V01','synthetic'],['VZ','DEMO-V02','DEMO Wälzstoßmaschine V02','synthetic'],
  ['EN','DEMO-E01','DEMO Entgratplatz E01','synthetic'],['EN','DEMO-E02','DEMO Entgratplatz E02','synthetic'],
  ['HT','DEMO-H01','DEMO Härteofen H01','synthetic'],['HT','DEMO-H02','DEMO Härtezelle H02','synthetic'],
  ['SL','DEMO-S01','DEMO Rundschleifmaschine S01','synthetic'],['SL','DEMO-S02','DEMO Flachschleifmaschine S02','synthetic'],
  ['QS','DEMO-Q01','DEMO Messraum Q01','synthetic'],['QS','DEMO-Q02','DEMO Zahnradmessplatz Q02','synthetic'],
  ['MO','DEMO-A01','DEMO Montageplatz A01','synthetic'],['MO','DEMO-A02','DEMO Montageplatz A02','synthetic']
];
// Six machines per existing department, retaining all 19 already named machines.
const extraMachineLabels={DR:'Drehzentrum',FR:'Fräszentrum',VZ:'Wälzfräsmaschine',EN:'Entgratplatz',HT:'Härtezelle',SL:'Schleifmaschine',QS:'Messplatz',MO:'Montageplatz'};
const extraMachineLetters={DR:'D',FR:'F',VZ:'V',EN:'E',HT:'H',SL:'S',QS:'Q',MO:'A'};
for(const [dept] of depSpec){
  const count=machinesSpec.filter(x=>x[0]===dept).length;
  for(let n=count+1;n<=6;n++){
    const short=extraMachineLetters[dept]+String(n).padStart(2,'0');
    machinesSpec.push([dept,'DEMO-'+short,'DEMO '+extraMachineLabels[dept]+' '+short,'synthetic']);
  }
}
const products = [
  {name:'Getriebewelle',material:'16MnCr5',route:['DR','FR','VZ','EN','HT','SL','QS'],setup:[50,75,90,15,60,65,35],cycle:[240,370,310,75,110,280,155]},
  {name:'Stirnrad',material:'16MnCr5',route:['DR','VZ','EN','HT','SL','QS'],setup:[50,85,15,65,75,40],cycle:[200,340,65,100,220,145]},
  {name:'Getriebegehäuse',material:'EN AW-6082',route:['FR','EN','QS','MO'],setup:[100,20,55,40],cycle:[880,100,170,185]},
  {name:'Lagerflansch',material:'42CrMo4',route:['DR','FR','EN','QS','MO'],setup:[65,85,15,40,35],cycle:[350,425,80,140,130]},
  {name:'Verzahntes Ritzel',material:'18CrNiMo7-6',route:['DR','VZ','EN','HT','SL','QS'],setup:[60,100,20,65,80,45],cycle:[300,420,80,100,280,175]},
  {name:'Schaltgabel',material:'C45',route:['FR','EN','QS','MO'],setup:[85,20,45,40],cycle:[510,100,180,145]},
  {name:'Abtriebswelle',material:'42CrMo4',route:['DR','FR','EN','SL','QS'],setup:[60,80,20,75,50],cycle:[330,390,85,230,150]},
  {name:'Planetenradträger',material:'16MnCr5',route:['DR','FR','EN','HT','SL','QS','MO'],setup:[85,115,20,65,70,50,50],cycle:[460,630,95,100,260,220,170]}
];
const opDescriptions = {
 DR:(p)=>`Rohteil ${p.material} auf Anschlag aufnehmen; Stirnflächen planen, Außendurchmesser vor- und fertigdrehen, Zentrierungen und Lagersitze mit Schleifaufmaß herstellen. Erststück messen.`,
 FR:(p)=>`Werkstück nach Bezug A/B aufspannen; Pass- und Querbohrungen, Nuten und Befestigungsgewinde gemäß synthetischer Zeichnung fertigen; Schnittkanten anfasen. Erststückfreigabe dokumentieren.`,
 VZ:(p)=>`Zahnform nach freigegebenem Demo-Zahndatenblatt wälzfräsen/stoßen; Verzahnungsbreite und Rundlauf prüfen, erforderliches Schleifaufmaß belassen.`,
 EN:()=>`Scharfe Kanten handentgraten, Bohrungskanten und Zahnflanken auf Grate prüfen; Bauteil reinigen und für Folge-OP bereitstellen.`,
 HT:(p)=>`Bauteile nach synthetischem Härteplan für ${p.material} wärmebehandeln; Chargen-Nachweis anlegen, Verzug und Härteprüfwerte dokumentieren.`,
 SL:()=>`Funktions- und Lagersitze auf Endmaß schleifen; Rundlauf, Maßhaltigkeit und Oberfläche prüfen; Schleifprotokoll ergänzen.`,
 QS:()=>`Erst- und Letztteilprüfung nach Demo-Prüfplan; Zeichnungsmaße, Form/Lage und ggf. Zahnradmessung dokumentieren. Abweichungen zur Klärung stellen.`,
 MO:()=>`Bauteil reinigen, Baugruppe nach Montagezeichnung fügen; Drehmomente und Leichtgängigkeit prüfen, fertig montierte Einheit kennzeichnen.`
};
const uid = (kind,n)=>`${kind}0000000-0000-4000-8000-${String(n).padStart(12,'0')}`;
const weekdays = t=>![0,6].includes(new Date(t).getUTCDay());
function nextShift(time){
  let d = new Date(time);d.setUTCSeconds(0,0);
  while(true){
    const dow=d.getUTCDay(),h=d.getUTCHours()+d.getUTCMinutes()/60;
    if(dow===0||dow===6||h>=21){d.setUTCDate(d.getUTCDate()+1);d.setUTCHours(5,0,0,0);continue;}
    if(h<5){d.setUTCHours(5,0,0,0);continue;}
    return +d;
  }
}
function finishWork(time,minutes){let x=nextShift(time),remaining=minutes;while(remaining>0){x=nextShift(x);const d=new Date(x);d.setUTCHours(21,0,0,0);const chunk=Math.min(remaining,(+d-x)/60000);if(chunk<=0)throw Error('Invalid work slot');x+=chunk*60000;remaining-=chunk;}return x;}
function weekdayOffset(days){let t=Date.parse('2026-11-13T05:00:00.000Z');const step=days<0?-1:1;for(let i=0;i<Math.abs(days);){t+=step*86400000;if(weekdays(t))i++;}return t;}
function makeDemoData(){
 const asOf=Date.parse(BASE);
 const departments=depSpec.map(([code,name],i)=>({id:uid('a',i+1),code,name,active:true,leader_employee_id:uid('b',i*5+1)}));
 const byDept=Object.fromEntries(departments.map(x=>[x.code,x]));
 const employees=nameBank.map(([first_name,last_name],i)=>({id:uid('b',i+1),auth_user_id:null,first_name,last_name,display_name:`${first_name} ${last_name}`,personnel_no:`DEMO-${String(i+1).padStart(3,'0')}`,department_id:departments[i<40?Math.floor(i/5):(i-40)%8].id,department:departments[i<40?Math.floor(i/5):(i-40)%8].name,employee_type:'core',role:i%5===0?'department_admin':'employee',is_active:true,color_key:['blue','green','amber','slate','purple'][i%5],created_at:BASE,updated_at:BASE}));
 const machines=machinesSpec.map(([dept,code,name,origin],i)=>({id:uid('c',i+1),machine_code:code,name,department_id:byDept[dept].id,active:true,origin}));
 const machinesByDept=Object.fromEntries(departments.map(d=>[d.code,machines.filter(m=>m.department_id===d.id)]));
 const machineReady=new Map(machines.map(m=>[m.id,Date.parse('2026-11-09T05:00:00.000Z')]));
 const operations=[],orders=[];
 const assignmentCounter=Object.fromEntries(depSpec.map(([code])=>[code,0]));
 const releases=[-6,-5,-4,-3,-3,-2,-2,-1,-1,-1,0,0,0,0,0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,3,3,4];
 for(let i=0;i<50;i++){
   const part=products[i%products.length],qty=[12,18,24,30,36,48,60,72][(i*3)%8];
   const orderId=uid('d',i+1),release=nextShift(weekdayOffset(releases[i]??(2+Math.floor((i-32)/4)))+((i*3)%6)*3600000);
   const order={id:orderId,ba_number:`DEMO-BA-${String(810001+i)}`,article_number:`DEMO-ART-${String(31001+i%8)}`,article_name:part.name,material:part.material,ba_quantity:qty,release_at:ts(release),operations:[]};
   let predecessor=release;
   part.route.forEach((code,j)=>{
      const ms=machinesByDept[code],m=ms[(assignmentCounter[code]++)%ms.length];
      const scheduledStart=nextShift(Math.max(predecessor,machineReady.get(m.id)));
      const setup=part.setup[j],cycle=part.cycle[j],duration=setup+Math.ceil(cycle*qty/60);
      const scheduledEnd=finishWork(scheduledStart,duration);
      machineReady.set(m.id,scheduledEnd);predecessor=scheduledEnd;
      const state=scheduledEnd<=asOf?'completed':scheduledStart<=asOf?'running':'queued';
      const elapsed=state==='completed'?duration:state==='running'?Math.min(duration,Math.max(0,(asOf-scheduledStart)/60000)):0;
      const progressQty=state==='completed'?qty:state==='running'?Math.min(qty,Math.max(0,Math.floor(qty*elapsed/duration))):0;
      const eligible=employees.filter(e=>e.department_id===byDept[code].id&&e.role==='employee');
      const employee=eligible[(i+j)%eligible.length];
      const op={id:uid('e',operations.length+1),order_id:orderId,op_number:String((j+1)*10).padStart(3,'0'),sequence:j+1,department_id:byDept[code].id,department:byDept[code].name,machine_id:m.id,machine:m.name,machine_code:m.machine_code,description:opDescriptions[code](part),setup_minutes:setup,cycle_seconds:cycle,planned_minutes:duration,planned_start:ts(scheduledStart),planned_end:ts(scheduledEnd),state,good_qty:progressQty,scrap_qty:0,assigned_employee_id:employee.id,predecessor_op_id:order.operations.length?order.operations[order.operations.length-1]:null};
      operations.push(op);order.operations.push(op.id);
   });
   order.status=operations.filter(op=>op.order_id===orderId).some(op=>op.state==='running')?'running':operations.filter(op=>op.order_id===orderId).every(op=>op.state==='completed')?'completed':'queued';
   order.next_operation=operations.find(op=>op.order_id===orderId&&op.state!=='completed')?.id||null;
   order.promised_date=ts(finishWork(predecessor,960));
   orders.push(order);
 }
 const machineQueue=machines.map(m=>({machine_id:m.id,machine:m.name,machine_code:m.machine_code,department_id:m.department_id,operations:operations.filter(op=>op.machine_id===m.id&&op.state!=='completed').sort((a,b)=>a.planned_start.localeCompare(b.planned_start)).map(op=>op.id)}));
 const data={schema_version:1,purpose:'isolated synthetic FABOS presentation, NOT production data',as_of:BASE,shift_model:'Mo-Fr 06:00-22:00 (Demo-Kapazitätsfenster; Feiertage und Personenabwesenheit nicht simuliert)',departments,employees,machines,orders,operations,machine_queue:machineQueue};
 validate(data);return data;
}
function validate(data){
 if(data.departments.length!==8||data.employees.length!==50||data.orders.length!==50||data.machines.length!==48)throw Error('Demo entity counts incorrect');
 const deptIDs=new Set(data.departments.map(x=>x.id)),machineIDs=new Set(data.machines.map(x=>x.id)),empIDs=new Set(data.employees.map(x=>x.id)),opIDs=new Set(data.operations.map(x=>x.id));
 for(const d of data.departments){if(![6,7].includes(data.employees.filter(e=>e.department_id===d.id).length)||!empIDs.has(d.leader_employee_id)||data.machines.filter(m=>m.department_id===d.id).length!==6)throw Error('Incorrect department staffing');}
 for(const m of data.machines)if(!deptIDs.has(m.department_id))throw Error('Orphan machine');
 for(const o of data.orders){let prev=null;for(const id of o.operations){const op=data.operations.find(x=>x.id===id);if(!op||op.order_id!==o.id||!machineIDs.has(op.machine_id)||!empIDs.has(op.assigned_employee_id)||op.good_qty>o.ba_quantity||!(op.setup_minutes>0)||!(op.cycle_seconds>0))throw Error('Invalid operation');if(prev&&op.planned_start<prev.planned_end)throw Error('Non-sequential routing');prev=op;}}
 for(const m of data.machines){const a=data.operations.filter(op=>op.machine_id===m.id).sort((x,y)=>x.planned_start.localeCompare(y.planned_start));for(let i=1;i<a.length;i++)if(a[i].planned_start<a[i-1].planned_end)throw Error('Machine double booked');}
 if(opIDs.size!==data.operations.length)throw Error('Duplicate operation');
 for(const m of data.machines)if(!data.operations.some(op=>op.machine_id===m.id))throw Error('Machine without demo job: '+m.name);
}
module.exports={makeDemoData,validate};