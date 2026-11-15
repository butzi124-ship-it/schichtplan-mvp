const { test, expect } = require('@playwright/test');
const { prepareMockPage, appUrl } = require('./helpers/mock-page');
const { rowsByTable, installSupabaseMock } = require('./helpers/supabase-mock');
const { baEntryDatabase, manager, worker, reviewer, department } = require('./helpers/ba-entry-db');
const { blank } = require('../js/operations');
const article = require('./fixtures/article-reference.json');
async function target(button) {
  await button.scrollIntoViewIfNeeded();
  await expect(button).toBeVisible();
  expect(await button.evaluate(el=>{const r=el.getBoundingClientRect(),top=document.elementFromPoint(r.x+r.width/2,r.y+r.height/2);return top===el||el.contains(top);})).toBe(true);
}
test.describe('Dashboard home approved production reference',()=>{
  test('actual BA overview, filters and detail selection preserve source data and open the existing workflow',async({page},info)=>{
    const errors=[];page.on('pageerror',e=>errors.push(e.message));
    await prepareMockPage(page);
    const rows=JSON.parse(JSON.stringify(rowsByTable));
    rows.production_order_checklist=rows.production_checklist_templates.map(i=>({id:'ref-'+i.id,order_id:'order-one',item_key:i.item_key,item_label:i.item_label,sort_order:i.sort_order,checked:false}));
    await page.addInitScript(installSupabaseMock,rows);
    await page.goto(appUrl);await page.getByRole('button',{name:'Mido/Admin',exact:true}).click();
    await page.locator('#tabs').getByRole('button',{name:'Produktion',exact:true}).click();
    await page.locator('#view').getByRole('button',{name:'Stückzahl',exact:true}).click();
    await expect(page.getByRole('heading',{name:'Fertigung – Produktionsübersicht',exact:true})).toBeVisible();
    await expect(page.locator('.h-kpi')).toHaveCount(5);
    const preview=page.getByRole('button',{name:'BA BA-100 Vorschau',exact:true});await target(preview);await preview.click();
    await expect(page.getByRole('complementary',{name:'Auftragsvorschau'})).toContainText('BA BA-100');
    const search=page.getByRole('searchbox',{name:'Fertigung durchsuchen',exact:true});await search.fill('NO MATCH');
    await expect(page.locator('[data-filter-empty]')).toBeVisible();await search.fill('BA-100');
    await expect(preview).toBeVisible();await search.fill('');
    expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth)).toBe(true);
    expect(await page.evaluate(()=>window.__SUPABASE_WRITE_LOG)).toEqual([]);
    if(await page.evaluate(()=>innerWidth<=800))await page.getByRole('button',{name:'Menü',exact:true}).click();
    await page.screenshot({path:info.outputPath('reference-production-overview.png'),fullPage:true});
    const open=page.locator('.h-reference-detail').getByRole('button',{name:/Stückzahlen \/ Auftrag öffnen/});await target(open);await open.click();
    await expect(page.locator('#productionActiveOrderPanel, .production-touch-view').first()).toBeVisible();
    await page.getByRole('group',{name:'Zähleransicht'}).getByRole('button',{name:'Übersicht',exact:true}).click();
    await page.screenshot({path:info.outputPath('reference-ba-detail.png'),fullPage:true});
    const plus=page.getByRole('button',{name:'Ausschuss +1',exact:true}).first();await target(plus);await plus.click();
    await expect(page.getByRole('heading',{name:'6M-Ursache für Ausschuss',exact:true})).toBeVisible();
    await page.screenshot({path:info.outputPath('reference-6m-dialog.png'),fullPage:true});
    await page.locator('.fixed.inset-0.z-50').last().getByRole('button',{name:'Abbrechen',exact:true}).first().click();
    expect(await page.evaluate(()=>window.__SUPABASE_WRITE_LOG)).toEqual([]);expect(errors).toEqual([]);
  });
  test('BA forms and immutable OP detail retain their real isolated SQL workflow',async({page},info)=>{
    const api=await baEntryDatabase();
    try {
      const r=await api.rpc(manager,'operations_article_command',{p_action:'create',p_department_id:department,p_article_number:'32505',p_data:article.data});
      const p=await api.rpc(manager,'operation_plan',{p_revision_id:r.id,p_operations:[{...blank(10),activity:'Fräsen',machine_group:'52'}],p_expected_version:r.row_version});
      await api.rpc(reviewer,'operations_article_command',{p_action:'release',p_revision_id:r.id,p_expected_version:p.row_version});
      await page.exposeFunction('__referenceRpc',async(name,params)=>{try{return{data:await api.rpc(worker,name,params),error:null};}catch(e){return{data:null,error:{code:e.code}};}});
      await prepareMockPage(page);
      await page.addInitScript(()=>{const make=window.supabase.createClient;window.supabase.createClient=(...args)=>({...make(...args),rpc:(name,params)=>window.__referenceRpc(name,params)});});
      await page.goto(appUrl);await page.getByRole('button',{name:'Mido/Admin',exact:true}).click();
      await page.evaluate(user=>{window.HUMBEL_ARTICLE_MASTER_ENABLED=true;window.HUMBEL_OPERATIONS_ENABLED=true;window.HUMBEL_BA_ENTRY_ENABLED=true;currentSupabaseUser={id:user};currentUser.role='employee';render();},worker);
      await page.locator('#tabs').getByRole('button',{name:'Produktion',exact:true}).click();
      await page.locator('#view').getByRole('button',{name:'Aufträge / BA',exact:true}).click();
      const panel=page.getByRole('region',{name:'BA-Erstanlage'});
      await panel.getByLabel('BA-Nummer',{exact:true}).fill('52978');await panel.getByLabel('Artikelnummer',{exact:true}).fill('32505');await panel.getByLabel('Stückzahl',{exact:true}).fill('360');
      if(await page.evaluate(()=>innerWidth<=800))await page.getByRole('button',{name:'Menü',exact:true}).click();
      await page.screenshot({path:info.outputPath('reference-ba-form.png'),fullPage:true});
      const submit=panel.getByRole('button',{name:'BA anlegen / vorhandenen öffnen'});await target(submit);await submit.click();
      await expect(panel.getByRole('status')).toContainText('angelegt');
      await expect(panel.getByRole('region',{name:'BA OP 010',exact:true})).toBeVisible();
      expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth)).toBe(true);
      await page.screenshot({path:info.outputPath('reference-op-snapshot.png'),fullPage:true});
    } finally { await api.db.close(); }
  });
});
