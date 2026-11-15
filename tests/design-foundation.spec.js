const { test, expect } = require('@playwright/test');
const { prepareMockPage, appUrl } = require('./helpers/mock-page');

test.describe('Dashboard home design foundation', () => {
  test.beforeEach(async ({ page }) => {
    await prepareMockPage(page);
    await page.goto(appUrl);
  });

  test('login retains controls, labels and keyboard access in the shared shell', async ({ page }, info) => {
    await expect(page.getByRole('heading', { name: 'Anmeldung', exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Menü', exact: true })).toBeHidden();
    await page.getByLabel('E-Mail', { exact: true }).fill('test@example.invalid');
    await page.getByLabel('Passwort', { exact: true }).fill('local-fixture-only');
    await page.getByLabel('Passwort', { exact: true }).press('Tab');
    await expect(page.getByRole('button', { name: 'Mit Supabase anmelden', exact: true })).toBeFocused();
    await expect(page.getByLabel('E-Mail', { exact: true })).toHaveValue('test@example.invalid');
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
    await page.screenshot({ path: info.outputPath('login-design.png'), fullPage: true });
  });

  test('navigation keeps the dark shared shell, active state and keyboard focus across modules', async ({ page }, info) => {
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.getByRole('button', { name: 'Mido/Admin', exact: true }).click();
    const nav = page.getByRole('navigation', { name: 'Hauptnavigation' });
    for (const name of ['Produktion', 'Werkzeugverwaltung', 'Start / Dashboard']) {
      const button = nav.getByRole('button', { name, exact: true });
      await button.focus();
      await button.press('Enter');
      await expect(nav.getByRole('button', { name, exact: true })).toBeFocused();
      await expect(nav.getByRole('button', { name, exact: true })).toHaveAttribute('aria-current', 'page');
      await expect(page.locator('body')).toHaveClass(/h-app-signed-in/);
      expect(await nav.evaluate(n => getComputedStyle(n).backgroundImage)).toContain('gradient');
    }
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
    expect(errors).toEqual([]);
    expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    await page.screenshot({ path: info.outputPath('dashboard-design.png'), fullPage: true });
  });

  test('mobile menu collapses, Escape restores focus and logout restores the login layout', async ({ page }) => {
    await page.getByRole('button', { name: 'Mido/Admin', exact: true }).click();
    const toggle = page.getByRole('button', { name: 'Menü', exact: true });
    const mobile = await page.evaluate(() => matchMedia('(max-width: 800px)').matches);
    if (mobile) {
      await toggle.click();
      await expect(toggle).toHaveAttribute('aria-expanded', 'false');
      await expect(page.locator('#tabs')).toBeHidden();
      await toggle.press('Enter');
      await expect(page.locator('#tabs')).toBeVisible();
      await page.locator('#tabs button').first().focus();
      await page.keyboard.press('Escape');
      await expect(toggle).toBeFocused();
      await expect(toggle).toHaveAttribute('aria-expanded', 'false');
      await toggle.press('Enter');
    } else {
      await expect(toggle).toBeHidden();
    }
    await page.locator('#tabs').getByRole('button', { name: 'Abmelden', exact: true }).click();
    await expect(page.locator('#loginBox')).toBeVisible();
    await expect(page.locator('#tabs')).toBeHidden();
    await expect(page.locator('body')).not.toHaveClass(/h-app-signed-in/);
    await expect(page.locator('#app-shell')).toHaveCSS('margin-left', '0px');
  });

  test('own shifts and tasks use labelled scroll regions and preserve employee permissions', async ({ page }, info) => {
    await page.getByRole('button', { name: 'Lavdrim (A)', exact: true }).click();
    await page.locator('#view').getByRole('button', { name: 'Meine Schichten', exact: true }).click();
    await expect(page.getByRole('heading', { name: 'Meine Schichten (90 Tage)', exact: true })).toBeVisible();
    const region = page.getByRole('region', { name: 'Meine Schichten', exact: true });
    await region.focus();
    await expect(region).toBeFocused();
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
    await page.screenshot({ path: info.outputPath('my-shifts-design.png'), fullPage: true });
    await page.locator('#view').getByRole('button', { name: 'To-Do', exact: true }).click();
    await expect(page.getByRole('region', { name: 'Offene Aufgaben', exact: true })).toBeVisible();
    await expect(page.getByRole('region', { name: 'Erledigte Aufgaben', exact: true })).toBeVisible();
    await expect(page.locator('#taskTitle')).toHaveCount(0);
    await expect(page.locator('#tabs').getByRole('button', { name: 'Personalverwaltung', exact: true })).toHaveCount(0);
    expect(await page.evaluate(() => window.__SUPABASE_WRITE_LOG)).toEqual([]);
    await page.screenshot({ path: info.outputPath('my-tasks-design.png'), fullPage: true });
  });

  test('admin task form keeps values, labelled fields and the original save action', async ({ page }) => {
    await page.getByRole('button', { name: 'Mido/Admin', exact: true }).click();
    await page.locator('#view').getByRole('button', { name: 'To-Do', exact: true }).click();
    await page.getByLabel('Aufgabe', { exact: true }).fill('Synthetische Aufgabe');
    await page.getByLabel('Frist', { exact: true }).fill('2026-12-01');
    await expect(page.getByLabel('Mitarbeiter', { exact: true })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Speichern', exact: true })).toHaveAttribute('onclick', 'createTask()');
    await expect(page.getByLabel('Aufgabe', { exact: true })).toHaveValue('Synthetische Aufgabe');
    await expect(page.getByLabel('Frist', { exact: true })).toHaveValue('2026-12-01');
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  });
});
