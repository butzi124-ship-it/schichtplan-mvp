'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const {rowsByTable}=require('../helpers/supabase-mock');
const {makeDemoData}=require('../../scripts/preview/fabos-demo-data');
const {buildDemoPreview}=require('../../scripts/preview/fabos-demo-adapter');
test('FABOS 50/48/50 synthetic seed has valid departmental assignments, operations and machine coverage',()=>{
 const d=makeDemoData();
 assert.equal(d.departments.length,8);
 assert.equal(d.employees.length,50);
 assert.equal(d.machines.length,48);
 assert.equal(d.orders.length,50);
 assert(d.operations.length>=250);
 assert.equal(new Set(d.employees.map(x=>x.id)).size,50);
 assert.equal(new Set(d.machines.map(x=>x.id)).size,48);
 assert.equal(new Set(d.orders.map(x=>x.ba_number)).size,50);
 for(const dep of d.departments){
  assert.equal(d.machines.filter(x=>x.department_id===dep.id).length,6,'six machines for '+dep.name);
  assert(d.employees.filter(x=>x.department_id===dep.id).length>=6,'staff for '+dep.name);
  const lead=d.employees.find(x=>x.id===dep.leader_employee_id);
  assert.equal(lead.department_id,dep.id);
  assert.equal(lead.role,'department_admin');
 }
 for(const machine of d.machines)assert(d.operations.some(x=>x.machine_id===machine.id),'machine has assigned operation: '+machine.name);
 for(const op of d.operations){
  const machine=d.machines.find(x=>x.id===op.machine_id),employee=d.employees.find(x=>x.id===op.assigned_employee_id);
  assert(machine&&employee);
  assert.equal(machine.department_id,op.department_id);
  assert.equal(employee.department_id,op.department_id);
 }
 const {rows}=buildDemoPreview(rowsByTable);
 assert.equal(rows.departments.length,8);
 assert.equal(rows.employees.length,51,'50 staff plus one synthetic administrator');
 assert.equal(rows.production_machines.length,48);
 assert.equal(rows.production_orders.length,50);
 assert(rows.production_orders.some(o=>o.status==='queued'));
});
