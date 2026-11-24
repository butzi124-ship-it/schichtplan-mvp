'use strict';
const {test,expect}=require('@playwright/test');
const {prepareMockPage,appUrl}=require('./helpers/mock-page');

test('mandatory shift-end checklist remains accessible without covering employee dashboard navigation',async({page})=>{
 await prepareMockPage(page);
 await page.goto(appUrl);
 await page.getByRole('button',{name:'Lavdrim (A)',exact:true}).click();
 await page.evaluate(()=>openChecklist('synthetic-navigation-check'));
 const checklist=page.locator('body > .fixed.inset-0').filter({has:page.locator('#cancelBtn')});
 await expect(checklist).toBeVisible();
 await expect(checklist).toHaveCSS('pointer-events','none');
 await expect(checklist.locator('#cancelBtn')).toHaveCSS('pointer-events','auto');
 await expect(checklist.getByRole('heading',{name:'Pflicht-Checkliste'})).toBeVisible();
 // An actual user click is required; no forced clicks or reduced hit-test assertions.
 await page.locator('#view').getByRole('button',{name:'Meine Schichten',exact:true}).click();
 await expect(page.getByRole('heading',{name:'Meine Schichten (90 Tage)',exact:true})).toBeVisible();
 await expect(checklist).toBeVisible();
 await checklist.getByRole('button',{name:'Abbrechen',exact:true}).click();
 await expect(checklist).toHaveCount(0);
 await page.locator('#view').getByRole('button',{name:'To-Do',exact:true}).click();
 await expect(page.getByRole('region',{name:'Offene Aufgaben',exact:true})).toBeVisible();
 expect(await page.evaluate(()=>window.__SUPABASE_WRITE_LOG)).toEqual([]);
});
