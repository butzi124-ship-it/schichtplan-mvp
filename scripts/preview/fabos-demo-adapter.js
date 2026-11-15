'use strict';
// Presentation-only bridge into the legacy in-memory mock. Never contacts Supabase.
const { makeDemoData } = require('./fabos-demo-data');
function buildDemoPreview(baseRows) {
  const data = makeDemoData();
  const rows = JSON.parse(JSON.stringify(baseRows));
  rows.departments = data.departments.map(d => ({ ...d }));
  rows.employees = [rows.employees[0], ...data.employees];
  rows.production_machines = data.machines.map(({origin, ...m}) => m);
  // Demo inventory is independently fictitious. No real manufacturer/tool IDs.
  const toolSpec=[
    ['Schaftfräser','10','HSK 63','F01','SixSigma',5],
    ['Schaftfräser','16','HSK 63','F02','SixSigma',3],
    ['Kugelfräser','8','HSK 63','F03','THAA',4],
    ['Radiusfräser','12','HSK 63','F04','THAA',3],
    ['Trochodialfräser','12','HSK 63','F05','SFS',3],
    ['NC Anbohrer','6','HSK 63','F06','SFS',6],
    ['Bohrer','8.5','HSK 63','F07','SFS',6],
    ['Bohrer','10.2','HSK 63','F08','THAA',4],
    ['Gewindebohrer M5','5','HSK 63','F09','SFS',4],
    ['Gewindebohrer M8','8','HSK 63','F10','SFS',3],
    ['Gewindefräser M12','12','HSK 63','F11','THAA',2],
    ['Ausdrehkopf','32','HSK 100','F12','SixSigma',2],
    ['Drehmeißel Schruppen','25','VDI 40','D01','SFS',5],
    ['Drehmeißel Schlichten','16','VDI 40','D02','SFS',4],
    ['Stechhalter','3','VDI 40','D03','THAA',3],
    ['Verzahnungsfräser Modul 2','55','Hob','V01','SixSigma',2],
    ['Entgratbürste','20','Hand','E01','THAA',7],
    ['Messuhr','0.01','Messraum','Q01','SFS',3],
    ['Grenzlehrdorn','12','Messraum','Q02','SFS',2],
    ['Schleifscheibe','200','Schleifspindel','S01','SixSigma',2]
  ];
  const toolBase=rows.tools[0]||{};
  rows.tools=toolSpec.map(([label,diameter,holder,shelf,manufacturer,stock],i)=>({
    ...toolBase,id:'demo-tool-'+String(i+1).padStart(3,'0'),
    t_number:'DEMO-T'+String(i+1).padStart(3,'0'),
    label,diameter,holder,shelf,manufacturer,stock,
    article_no:'DEMO-WZ-'+String(51001+i),min_stock:Math.max(1,Math.floor(stock/3)),
    optimal_stock:stock+2,ordered:false,ordered_qty:0,is_borrowed:false
  }));

  const current = new Map();
  for (const order of data.orders) {
    const ops = data.operations.filter(op => op.order_id === order.id);
    const active = ops.find(op => op.state === 'running');
    const closed = ops[ops.length - 1];
    if (active || order.status === 'completed') current.set(order.id, active || closed);
  }
  rows.production_orders = data.orders.filter(o => current.has(o.id)).map(o => {
    const op = current.get(o.id);
    const completed = o.status === 'completed';
    return {
      id: o.id, machine_id: op.machine_id, department_id: op.department_id,
      ba_number: o.ba_number, article_number: o.article_number,
      ba_quantity: o.ba_quantity, target_quantity: o.ba_quantity,
      pallet_count: Math.ceil(o.ba_quantity / 12), pieces_per_pallet: 12,
      use_chain_logic: false, status: completed ? 'completed' : 'running',
      started_at: op.planned_start, completed_at: completed ? op.planned_end : null,
      created_by_employee_id: rows.employees[0].id,
      updated_by_employee_id: rows.employees[0].id,
      created_at: o.release_at, updated_at: data.as_of,
    };
  });
  rows.production_order_stations = [...current.values()].map(op => ({
    id: op.id, order_id: op.order_id, station_no: 1,
    name: 'Spannung 1 · OP ' + op.op_number, lock_name: true,
    op_number: op.op_number, time_status: 'ok',
    actual_time_minutes: op.state === 'completed' ? op.planned_minutes : null,
    scrap_total: op.scrap_qty, clarify_total: 0, scrap_lifetime: op.scrap_qty, clarify_lifetime: 0,
    created_at: op.planned_start, updated_at: data.as_of,
  }));
  rows.production_order_employees = [...current.values()].map(op => {
    const e = data.employees.find(x => x.id === op.assigned_employee_id);
    return {id: 'demo-staff-' + op.id,order_id: op.order_id,
      employee_id:e.id,employee_name:e.display_name,personnel_no:e.personnel_no,
      role:'worker',sort_order:1,active:true,created_at:op.planned_start,updated_at:data.as_of};
  });
  rows.production_station_counts = [...current.values()].map(op => ({
    id:'demo-count-' + op.id,order_id:op.order_id,station_id:op.id,
    order_employee_id:'demo-staff-' + op.id,good_qty:op.good_qty,
    created_at:op.planned_start,updated_at:data.as_of
  }));
  // Never leave fixtures from the original unrelated four test orders behind.
  rows.production_station_events = [];
  rows.production_order_history = [];
  rows.production_counts = [];
  rows.production_order_checklist = [];
  return { data, rows };
}
const esc = value => String(value ?? '').replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function renderDemoPlan(data){
  const orderById=new Map(data.orders.map(x=>[x.id,x]));
  const opsById=new Map(data.operations.map(x=>[x.id,x]));
  const nameByDept=new Map(data.departments.map(x=>[x.id,x.name]));
  const displayDate=x=>new Date(x).toLocaleString('de-DE',{timeZone:'Europe/Berlin',day:'2-digit',month:'2-digit',hour:'2-digit',minute:'2-digit'});
  const status={running:'Läuft',queued:'Wartet',completed:'Fertig'};
  const machineRows=data.machines.map(m=>{
    const ops=data.machine_queue.find(q=>q.machine_id===m.id).operations.map(id=>opsById.get(id));
    const live=ops.find(x=>x.state==='running'),next=ops.find(x=>x.state==='queued');
    const cell=op=>op?esc(orderById.get(op.order_id).ba_number+' · OP '+op.op_number+' · '+displayDate(op.planned_start)):'–';
    return '<tr><td>'+esc(nameByDept.get(m.department_id))+'</td><td>'+esc(m.name)+'</td><td>'+cell(live)+'</td><td>'+cell(next)+'</td></tr>';
  }).join('');
  const orderRows=data.orders.map(o=>{
    const ops=o.operations.map(id=>opsById.get(id));
    const detail=ops.map(op=>'<tr><td>'+esc(op.op_number)+'</td><td>'+esc(op.department)+'</td><td>'+esc(op.machine)+'</td><td>'+esc(op.description)+'</td><td>'+esc(op.setup_minutes)+' min</td><td>'+esc(op.cycle_seconds)+' s</td><td>'+esc(displayDate(op.planned_start))+'</td><td>'+esc(displayDate(op.planned_end))+'</td><td>'+esc(status[op.state])+'</td></tr>').join('');
    return '<details><summary><strong>'+esc(o.ba_number)+'</strong> · '+esc(o.article_name)+' · '+esc(o.material)+' · '+esc(o.ba_quantity)+' Stk. · '+esc(status[o.status])+'</summary><div class="scroll"><table><thead><tr><th>OP</th><th>Abteilung</th><th>Maschine</th><th>Arbeitsanweisung</th><th>Rüsten</th><th>Zyklus</th><th>Start</th><th>Ende</th><th>Status</th></tr></thead><tbody>'+detail+'</tbody></table></div></details>';
  }).join('');
  return '<!doctype html><html lang="de"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>FABOS · Synthetische Maschinenbelegung</title><style>body{font:15px system-ui;background:#f3f7fc;color:#14304a;margin:2rem}h1{color:#091e37}a{color:#0871dc}table{width:100%;border-collapse:collapse;background:white}th,td{border-bottom:1px solid #dce5f1;padding:10px;text-align:left;vertical-align:top}th{background:#eaf1fa}details{background:white;margin:14px 0;border:1px solid #dce5f1;border-radius:8px;padding:14px}summary{cursor:pointer}summary:focus-visible{outline:3px solid #0871dc}.scroll{overflow-x:auto}aside{background:#fff2c8;padding:12px;border-radius:8px}</style><h1>FABOS – Maschinenbelegung und OP-Vorschau (nur Demo)</h1><aside>Alle Personen, Aufträge und Zeitbuchungen sind synthetisch. Keine Original-Firmendaten. Diese Übersicht ist ein lesbarer Demo-Plan, keine echte PlanOS-/RPC-Maschinenplanung. Die Fertigungszähler der Testvorschau verwenden eine lokale Browserfixture und werden nicht gemeinsam auf einem Server gespeichert.</aside><p><a href="/">Zur Testanwendung</a> · Bezugszeit: '+esc(displayDate(data.as_of))+' · 8 Abteilungen · 40 fiktive Personen · 32 BA · '+data.operations.length+' OP</p><h2>Maschinen: aktueller und nächster Auftrag</h2><div class="scroll"><table><thead><tr><th>Abteilung</th><th>Maschine</th><th>Jetzt laufend</th><th>Nächster Auftrag</th></tr></thead><tbody>'+machineRows+'</tbody></table></div><h2>Alle Betriebsaufträge mit OP-Folge, Rüst- und Stückzeit</h2>'+orderRows+'</html>';
}
module.exports={buildDemoPreview,renderDemoPlan};
