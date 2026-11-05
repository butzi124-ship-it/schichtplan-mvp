import { test, expect } from '@playwright/test';
const { prepareMockPage, appUrl } = require('./helpers/mock-page');

test.beforeEach(async ({ page }) => {
  await prepareMockPage(page);
});

test('Humbel safe mock seed', async ({ page }) => {
  await page.goto(appUrl);
  await expect(page.getByRole('heading', { name: 'Anmeldung' })).toBeVisible();
  await page.getByRole('button', { name: 'Mido/Admin', exact: true }).click();
  await expect(page.locator('#tabs')).toBeVisible();
  expect(await page.evaluate(() => (window as any).__SUPABASE_WRITE_LOG)).toEqual([]);
});
