const { test, expect } = require('@playwright/test');
const { prepareMockPage, appUrl } = require('./helpers/mock-page');
const { rowsByTable, installSupabaseMock } = require('./helpers/supabase-mock');

async function hitTarget(button) {
  await expect(button).toBeVisible();
  await button.scrollIntoViewIfNeeded();
  expect(await button.evaluate(el => {
    const r = el.getBoundingClientRect();
    const top = document.elementFromPoint(r.x + r.width / 2, r.y + r.height / 2);
    return top === el || el.contains(top);
  })).toBe(true);
}

test.describe('Dashboard home display corrections', () => {
  test.beforeEach(async ({ page }) => {
    await prepareMockPage(page);
    const rows = JSON.parse(JSON.stringify(rowsByTable));
    rows.production_orders[0].target_quantity = 50;
    Object.assign(rows.production_order_stations[0], { scrap_total: 2, clarify_total: 2 });
    for (let i = 2; i <= 4; i++) {
      rows.production_order_stations.push({ ...rows.production_order_stations[0], id: 'display-station-' + i, station_no: i, name: 'Spannung ' + i });
    }
    // Existing first-open checklist initialization is outside this cancellation test.
    // Match the original visual fixture: a fully initialized synthetic order.
    rows.production_order_checklist = rows.production_checklist_templates.map(item => ({
      id: 'display-' + item.id, order_id: 'order-one', item_key: item.item_key,
      item_label: item.item_label, sort_order: item.sort_order, checked: false,
    }));
    await page.addInitScript(installSupabaseMock, rows);
    // The loopback documentation preview enables these same presentation modules.
    await page.addInitScript(() => {
      for (const name of ['HUMBEL_CONTEXT_HELP_ENABLED', 'HUMBEL_LEARNING_UI_ENABLED']) {
        Object.defineProperty(window, name, { configurable: true, get: () => true, set() {} });
      }
    });
    await page.goto(appUrl);
    await page.getByRole('button', { name: 'Mido/Admin', exact: true }).click();
  });

  test('amount controls stay inside their cards and real clicks open the correct cause dialog', async ({ page }, info) => {
    await page.locator('#tabs').getByRole('button', { name: 'Produktion', exact: true }).click();
    await page.locator('#view').getByRole('button', { name: 'Aufträge / BA', exact: true }).click();
    await page.locator('tr[data-production-order-id="order-one"]').getByRole('button', { name: 'BA BA-100 öffnen', exact: true }).click();
    await page.getByRole('group', { name: 'Zähleransicht' }).getByRole('button', { name: 'Übersicht', exact: true }).click();
    const cards = page.locator('.h-amount-controls > div');
    await expect(cards).toHaveCount(8);
    for (const card of await cards.all()) {
      for (const button of await card.getByRole('button').all()) {
        const bounds = await button.boundingBox();
        const parent = await card.boundingBox();
        expect(bounds.x).toBeGreaterThanOrEqual(parent.x);
        expect(bounds.x + bounds.width).toBeLessThanOrEqual(parent.x + parent.width);
        expect(bounds.width).toBeGreaterThanOrEqual(44);
        expect(bounds.height).toBeGreaterThanOrEqual(44);
        await hitTarget(button);
      }
    }
    for (const label of ['Ausschuss', 'In Abklärung']) {
      await cards.getByRole('button', { name: label + ' +1', exact: true }).first().click();
      await expect(page.getByRole('heading', { name: '6M-Ursache für ' + label, exact: true })).toBeVisible();
      await page.locator('.fixed.inset-0.z-50').last().getByRole('button', { name: 'Abbrechen', exact: true }).first().click();
    }
    expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    await page.screenshot({ path: info.outputPath('production-controls.png'), fullPage: true });
  });

  test('help and version controls stay in flow and remain clickable without covering the dashboard', async ({ page }, info) => {
    const dock = page.locator('#app-utilities');
    await expect(dock).toBeVisible();
    const viewBounds = await page.locator('#view').boundingBox();
    const dockBounds = await dock.boundingBox();
    expect(dockBounds.y).toBeGreaterThanOrEqual(viewBounds.y + viewBounds.height);
    const help = page.locator('#context-help-launcher');
    const version = page.locator('#versionFooter button');
    for (const button of [help, version]) await hitTarget(button);
    const h = await help.boundingBox(), v = await version.boundingBox();
    expect(h.x + h.width <= v.x || v.x + v.width <= h.x || h.y + h.height <= v.y || v.y + v.height <= h.y).toBe(true);
    await help.click();
    await expect(page.locator('#context-help-dialog')).toBeVisible();
    await page.locator('#context-help-dialog').getByRole('button', { name: 'Schließen', exact: true }).click();
    await version.click();
    await expect(page.getByRole('heading', { name: 'Versionslog', exact: true })).toBeVisible();
    await page.getByRole('button', { name: 'Schließen', exact: true }).click();
    await page.screenshot({ path: info.outputPath('dashboard-utilities.png'), fullPage: true });
  });

  test('preview banner resizes without covering desktop or tablet navigation', async ({ page }, info) => {
    await page.evaluate(() => {
      const banner = document.createElement('aside');
      banner.id = 'isolated-preview-banner';
      banner.style = 'position:sticky;top:0;z-index:90;background:#fff3c4;padding:12px;font:14px system-ui';
      banner.textContent = 'SYNTHETISCHE ENTWICKLUNGSANSICHT · ausschließlich synthetische Daten · keine native Supabase-Abnahme · Neustart setzt Daten zurück';
      document.body.prepend(banner);
    });
    await expect.poll(() => page.evaluate(() => parseFloat(document.documentElement.style.getPropertyValue('--h-preview-offset')) === Math.ceil(document.getElementById('isolated-preview-banner').getBoundingClientRect().height))).toBe(true);
    if (await page.evaluate(() => innerWidth > 800)) {
      const banner = await page.locator('#isolated-preview-banner').boundingBox();
      const nav = await page.locator('#tabs').boundingBox();
      expect(nav.y).toBeGreaterThanOrEqual(banner.y + banner.height);
    }
    await hitTarget(page.locator('#tabs').getByRole('button', { name: 'Start / Dashboard', exact: true }));
    await page.screenshot({ path: info.outputPath('preview-header.png'), fullPage: true });
  });

  test('navigation has a semantic compound-word break without clipping menu labels', async ({ page }) => {
    const button = page.locator('#tabs').getByRole('button', { name: 'Sicherheitsunterweisungen', exact: true });
    await expect(button.locator('wbr')).toHaveCount(1);
    for (const entry of await page.locator('#tabs button').all()) {
      expect(await entry.evaluate(el => el.scrollWidth <= el.clientWidth)).toBe(true);
    }
    await hitTarget(button);
    await button.click();
    await expect(button).toHaveAttribute('aria-current', 'page');
  });

  test('tool table actions are reachable without utility overlays on every viewport', async ({ page }, info) => {
    await page.locator('#tabs').getByRole('button', { name: 'Werkzeugverwaltung', exact: true }).click();
    const actions = page.locator('#view table button');
    await expect(actions.first()).toBeVisible();
    for (const button of await actions.all()) await hitTarget(button);
    await page.locator('#view').screenshot({ path: info.outputPath('tool-actions.png') });
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  });

  test('the existing downtime control stays in the header flow and cannot cover employee shift actions', async ({ page }, info) => {
    await page.locator('#tabs').getByRole('button', { name: 'Abmelden', exact: true }).click();
    await page.getByRole('button', { name: 'Lavdrim (A)', exact: true }).click();
    const timer = page.getByRole('button', { name: 'Stillstand läuft – Stoppen', exact: true });
    await expect(timer).toBeVisible();
    await expect(page.locator('.h-app-header #downtimeWidget')).toHaveCSS('position', 'static');
    await expect(timer).toHaveAttribute('onclick', 'stopDowntimeTimer()');
    await hitTarget(timer);
    const h = await page.locator('.h-app-header').boundingBox();
    const t = await timer.boundingBox();
    expect(t.y).toBeGreaterThanOrEqual(h.y);
    expect(t.y + t.height).toBeLessThanOrEqual(h.y + h.height);
    await page.locator('#view').getByRole('button', { name: 'Meine Schichten', exact: true }).click();
    await hitTarget(page.getByRole('button', { name: 'Abwesenheit', exact: true }).first());
    await page.screenshot({ path: info.outputPath('employee-header.png'), fullPage: true });
  });
});
