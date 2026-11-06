const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const {BitArray,Code128Reader}=require('@zxing/library');
const {opNumber,barcode,parseBarcode,blank,normalize}=require('../../js/operations');
const ref=require('../fixtures/operations-reference.json');
// Exercise exactly the vendored browser encoder, decode using an independent implementation.
const sandbox={window:{}};vm.runInNewContext(fs.readFileSync('js/vendor/JsBarcode.code128.min.js','utf8'),sandbox);
function decode(bits){const quiet='0'.repeat(20);const s=quiet+bits+quiet;const row=new BitArray(s.length);[...s].forEach((b,i)=>{if(b==='1')row.set(i);});return new Code128Reader().decodeRow(0,row).getText();}
test('Code128 browser vendor independently decodes BA and all mandatory OP barcodes including leading zeros',()=>{
 assert.equal(fs.readFileSync('js/vendor/JsBarcode.code128.min.js','utf8'),fs.readFileSync('node_modules/jsbarcode/dist/barcodes/JsBarcode.code128.min.js','utf8'));
 for(const ba of ['812150','000812150'])for(const n of [null,10,20,100,110]){
  const value=barcode(ba,n),out={};sandbox.window.JsBarcode(out,value,{format:'CODE128'});
  assert.equal(decode(out.encodings.map(e=>e.data).join('')),value);
 }
});
test('barcode parser keeps BA text exact, canonical OP width and scanner CR/LF terminator; rejects ambiguous formats',()=>{
 assert.equal(opNumber(10),'010');assert.equal(opNumber(100),'100');
 assert.deepEqual(parseBarcode('000812150.010\r\n'),{ba:'000812150',op:10,value:'000812150.010'});
 assert.equal(barcode('000812150',110),'000812150.110');
 for(const value of ['1.1','1.10','1.0010','1.000','1.010.020','1e3',' 812150','812150 ','１２３','BA812150','https://example.test/812150','812150\n.010',null])assert.throws(()=>parseBarcode(value));
 for(const n of [0,1000,1.1,'010',NaN])assert.throws(()=>opNumber(n));
});
test('default OP sequence increments by ten; reference contains only known OP 100 and no invented times',()=>{
 assert.equal(blank().number,10);assert.equal(blank(20,10).predecessor,10);
 assert.equal(ref.provenance.complete_operation_plan,false);assert.equal(ref.operations.length,1);
 assert.equal(ref.operations[0].number,100);assert.deepEqual(ref.operations[0].tr,{value:null,unit:null});assert.deepEqual(ref.operations[0].te,{value:null,unit:null});
 assert.deepEqual(normalize(ref.operations),ref.operations);
});
test('Tr/Te validate explicit units, zero, decimal and unknown pair without mutating source',()=>{
 const a={...ref.operations[0],tr:{value:1.5,unit:'h'},te:{value:0,unit:'s'}};assert.deepEqual(normalize([a])[0].tr,a.tr);
 for(const t of [{value:-1,unit:'min'},{value:Infinity,unit:'min'},{value:'2',unit:'s'},{value:2,unit:null},{value:null,unit:'min'},{value:1,unit:'day'},{value:1,unit:'h',other:2}])assert.throws(()=>normalize([{...a,tr:t}]));
});
test('OP identity, ordered predecessors, unknown order/counter fields and invalid machine IDs fail closed',()=>{
 const a=ref.operations[0];
 for(const ops of [[a,a],[{...a,predecessor:100}],[{...a,number:0}],[{...a,ba_number:'812150'}],[{...a,machine_id:'bad'}],[{...a,activity:''}],[{...a,description:'x'.repeat(12001)}]])assert.throws(()=>normalize(ops));
 const plan=[{...a,number:10},{...a,number:20,predecessor:10}];assert.deepEqual(normalize(plan),plan);
});
