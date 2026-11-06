const { test, expect } = require('@playwright/test');

test('CI gate probe: intentional failure must block merge', () => {
  expect(true).toBe(false);
});
