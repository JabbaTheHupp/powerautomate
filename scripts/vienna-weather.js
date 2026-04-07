#!/usr/bin/env bun
const { chromium } = require('playwright');

async function run() {
  const browser = await chromium.launch({ headless: false, slowMo: 100 });
  const context = await browser.newContext();
  const page = await context.newPage();

  await page.goto('https://tipicoltd.sharepoint.com/', { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(5000);
  await browser.close();
}

run().catch((error) => {
  console.error('Automation failed:', error);
  process.exitCode = 1;
});