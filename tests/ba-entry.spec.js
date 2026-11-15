const {test,expect}=require('@playwright/test');
const {prepareMockPage,appUrl}=require('./helpers/mock-page');
const {baEntryDatabase,manager,worker,reviewer,outsider,department}=require('./helpers/ba-entry-db');
const {blank}=require('../js/operations');
const article=require('./fixtures/article-reference.json');
test.describe('Operations plan phase 1 BA entry',()=>{
  let api,uid,calls,failNext,delay;
  test.beforeEach(async({page})=>{
    api=await baEntryDatabase();uid=worker;calls=[];failNext=false;delay=0;
    let r=await api.rpc(manager,'operations_article_command',{p_action:'create',p_department_id:department,p_article_number:'32505',p_data:article.data});
    const operations=[10,100].map((n,i)=>({...blank(n,i?10:null),activity:'Synthetic OP '+n,machine_group:'52'}));
    const p=await api.rpc(manager,'operation_plan',{p_revision_id:r.id,p_operations:operations,p_expected_version:r.row_version});
    await api.rpc(reviewer,'operations_article_command',{p_action:'release',p_revision_id:r.id,p_expected_version:p.row_version});
    await page.exposeFunction('__entryRpc',async(name,params)=>{
      calls.push({name,params});if(failNext&&name==='ba_entry_command'){failNext=false;return{error:{code:'08006'}};}
      if(delay)await new Promise(resolve=>setTimeout(resolve,delay));
      try{return{data:await api.rpc(uid,name,params),error:null};}catch(e){return{data:null,error:{code:e.code}};}
    });
    await prepareMockPage(page);
    await page.addInitScript(()=>{const make=window.supabase.createClient;window.supabase.createClient=(...args)=>({...make(...args),rpc:(name,params)=>window.__entryRpc(name,params)});});
  });
  test.afterEach(async()=>api?.db.close());
  async function login(page,{enabled=true,role='employee',user=worker}={}){
    uid=user;await page.goto(appUrl);await page.getByRole('button',{name:'Mido/Admin'}).click();
    await page.evaluate(({enabled,role,user})=>{window.HUMBEL_ARTICLE_MASTER_ENABLED=true;window.HUMBEL_OPERATIONS_ENABLED=true;window.HUMBEL_BA_ENTRY_ENABLED=enabled;currentSupabaseUser={id:user};currentUser.role=role;render();},{enabled,role,user});
    await page.locator('#tabs').getByRole('button',{name:'Produktion',exact:true}).click();
    await expect(page.getByRole('button',{name:'Artikelstamm',exact:true}).first()).toBeVisible();
  }
  async function panel(page){await page.getByRole('button',{name:'Aufträge / BA',exact:true}).click();return page.getByRole('region',{name:'BA-Erstanlage'});}
  async function fill(p,ba='000812160',quantity='48'){
    await p.getByLabel('BA-Nummer',{exact:true}).fill(ba);await p.getByLabel('Artikelnummer',{exact:true}).fill('32505');await p.getByLabel('Stückzahl',{exact:true}).fill(quantity);
  }
  async function submit(p){await p.getByRole('button',{name:'BA anlegen / vorhandenen öffnen'}).click();}
  test('employee enters exactly three required fields; server supplies released snapshot; no counter or role opening',async({page})=>{
    await login(page);const p=await panel(page);await expect(p.locator('#ba-entry-create [required]')).toHaveCount(3);
    await expect(p.getByLabel('Liefer-/Endtermin')).toHaveCount(0);await expect(p.getByLabel('BA-Maschine')).toHaveCount(0);
    await fill(p);await submit(p);await expect(p.getByRole('status')).toContainText('angelegt');await expect(p).toContainText('Termin Nicht geplant');await expect(p).toContainText('Zuordnungsregel noch offen');
    await expect(p.getByRole('region',{name:'BA OP 010'})).toHaveAttribute('data-selected-op','true');
    expect(await page.evaluate(()=>canAccessProduction())).toBe(false);
    const request=calls.find(c=>c.name==='ba_entry_command'&&c.params.p_action==='create');expect(request.params).toEqual({p_action:'create',p_ba_number:'000812160',p_article_number:'32505',p_quantity:48});
    expect((await api.db.query('select count(*)::int n from public.production_station_counts')).rows[0].n).toBe(1);
    expect(await page.evaluate(()=>Object.values(localStorage).join(' '))).not.toContain('Synthetic OP');
  });
  test('duplicate and repeated barcode/manual OP lookup open existing BA without modifying its quantity',async({page})=>{
    await login(page);const p=await panel(page);await fill(p);await submit(p);await expect(p.getByRole('status')).toContainText('angelegt');
    await fill(p,'000812160','999');await submit(p);await expect(p.getByRole('status')).toContainText('Vorhandenen BA geöffnet');await expect(p).toContainText('48 Stück');
    for(const value of ['000812160.100','000812160.100']){await p.getByLabel('BA-/OP-Barcode',{exact:true}).fill(value);await p.getByRole('button',{name:'Barcode öffnen',exact:true}).click();await expect(p.getByRole('status')).toContainText('Keine Mengenbuchung');await expect(p.getByRole('region',{name:'BA OP 100'})).toHaveAttribute('data-selected-op','true');}
    await p.getByLabel('BA für Suche').fill('000812160');await p.getByLabel('OP-Nummer (optional)').fill('10');await p.getByRole('button',{name:'BA / OP suchen',exact:true}).click();await expect(p.getByRole('region',{name:'BA OP 010'})).toHaveAttribute('data-selected-op','true');
    expect((await api.db.query("select count(*)::int n from public.production_orders where ba_number='000812160'")).rows[0].n).toBe(1);
  });
  test('transport failure preserves input and rapid submit creates only one BA',async({page})=>{
    await login(page);const p=await panel(page);await fill(p);failNext=true;await submit(p);await expect(p.getByRole('alert')).toContainText('fehlgeschlagen');await expect(p.getByLabel('BA-Nummer',{exact:true})).toHaveValue('000812160');
    delay=200;await p.getByRole('button',{name:'BA anlegen / vorhandenen öffnen'}).dblclick();await expect(p.getByRole('status')).toContainText('angelegt');
    expect(calls.filter(c=>c.name==='ba_entry_command'&&c.params.p_action==='create')).toHaveLength(2);
    expect((await api.db.query("select count(*)::int n from public.production_orders where ba_number='000812160'")).rows[0].n).toBe(1);
  });
  test('revoked permission is checked again at submit; scanner/outsider/disabled do not get entry access',async({page})=>{
    await login(page);const p=await panel(page);await fill(p);await api.db.query('update public.article_master_assignments set revoked=true where user_id=$1',[worker]);await submit(p);await expect(p.getByRole('alert')).toBeVisible();
    await page.evaluate(()=>{currentUser.role='tool_scanner';render();});await expect(page.getByRole('region',{name:'BA-Erstanlage'})).toHaveCount(0);
    await page.evaluate(()=>{currentUser.role='employee';window.HUMBEL_BA_ENTRY_ENABLED=false;render();});await expect(page.getByRole('region',{name:'BA-Erstanlage'})).toHaveCount(0);
    uid=outsider;await page.evaluate(user=>{window.HUMBEL_BA_ENTRY_ENABLED=true;currentSupabaseUser={id:user};render();},outsider);await expect(page.getByRole('region',{name:'BA-Erstanlage'})).toHaveCount(0);
  });
  test('geometry, session reset and delayed response cannot expose previous employee snapshot',async({page})=>{
    await login(page);const p=await panel(page);await fill(p);await submit(p);await expect(p.getByRole('status')).toContainText('angelegt');
    expect(await p.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1)).toBe(true);
    delay=300;await p.getByLabel('BA-/OP-Barcode',{exact:true}).fill('000812160.100');await p.getByRole('button',{name:'Barcode öffnen',exact:true}).click();
    await page.evaluate(()=>logoutSupabase());await expect(page.getByRole('region',{name:'BA-Erstanlage'})).toHaveCount(0);
    await expect(page.getByRole('region',{name:'BA-Operationssnapshot'})).toHaveCount(0);
    expect(await page.evaluate(()=>Object.values(localStorage).join(' '))).not.toContain('000812160');
  });
});
