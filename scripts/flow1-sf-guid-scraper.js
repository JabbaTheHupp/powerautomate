/**
 * Flow 1 — Salesforce GUID → Hyperlink Scraper
 *
 * KONFIGURATION: Nur diesen Block anpassen
 */
const CONFIG = {
  // Pfad zur Excel-Datei mit den GUIDs
  excelInputPath: './CommsLogs_UUID_Lists/Template_Automation_Feeder_List_CommsLogs.xlsm',

  // Name des Tabellenblatts (oder null für erstes Sheet)
  sheetName: 'Sheet1',

  // Spaltenname für die PersonaID
  personaIdColumn: 'PersonaID',

  // Spaltenname in der Excel-Datei, der die GUIDs enthält (ggf. mehrere per : getrennt)
  guidColumn: 'UUID',

  // Spaltenname, in den der Hyperlink geschrieben wird
  linkColumn: 'Hyperlink',

  // Salesforce Login-URL (Microsoft SSO)
  loginUrl: 'https://tipico.lightning.force.com',

  // SF Home URL — nach Login wird hierhin navigiert
  homeUrl: 'https://tipico.lightning.force.com/lightning/page/home',

  // Browser sichtbar lassen (MFA erfordert manuellen Login)
  headless: false,
};

// ─────────────────────────────────────────────────────────────────────────────

import * as XLSX from 'xlsx';
import { chromium } from 'playwright';
import fs from 'fs';

const OUTPUT_PATH = CONFIG.excelInputPath.replace(/(\.\w+)$/, '_output.xlsx');

async function readGuids() {
  const workbook = XLSX.readFile(CONFIG.excelInputPath);
  const sheetName = CONFIG.sheetName ?? workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet);
  console.log(`📋 ${rows.length} GUIDs geladen aus "${sheetName}"`);
  return { rows, workbook, sheetName };
}

async function saveProgress(outputRows) {
  const ws = XLSX.utils.json_to_sheet(outputRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Results');
  XLSX.writeFile(wb, OUTPUT_PATH);
}

async function closeAllSfTabs(page) {
  await page.evaluate(() => document.body.dispatchEvent(new MouseEvent('click', { bubbles: true })));
  await page.waitForTimeout(500);
  await page.keyboard.press('Shift+W');
  try {
    const confirmBtn = page.getByRole('button', { name: 'Close All' });
    await confirmBtn.waitFor({ timeout: 5000 });
    await confirmBtn.click();
    await page.locator('text=/tabs were closed/').waitFor({ state: 'hidden', timeout: 10000 }).catch(() => {});
    console.log('  🧹 Alle SF-Tabs geschlossen');
  } catch {
    console.log('  🧹 Keine Tabs offen');
  }
  await page.waitForTimeout(500);
}

async function processGuid(page, guid) {
  // / öffnet die globale Suchleiste zuverlässig, auch wenn sie schon offen war
  await page.keyboard.press('/');
  const searchBox = page.getByRole('searchbox', { name: 'Search...' });
  await searchBox.waitFor({ timeout: 10000 });
  // Explizit klicken um Fokus sicherzustellen, dann Inhalt leeren und neu befüllen
  await searchBox.click();
  await searchBox.fill('');
  await searchBox.fill(guid);
  await searchBox.press('Enter');

  // Warten bis Suchergebnisse geladen sind
  await page.getByRole('heading', { name: 'Search Results' }).waitFor({ timeout: 15000 });

  // Race: entweder Tabelle mit Ergebnis oder "Don't give up yet!" — wer zuerst erscheint
  const resultLink = page.locator('table').first().getByRole('link').first();
  const noResultMsg = page.getByText("Don't give up yet!", { exact: true });

  const winner = await Promise.race([
    resultLink.waitFor({ timeout: 15000 }).then(() => 'found'),
    noResultMsg.waitFor({ timeout: 15000 }).then(() => 'not_found'),
  ]).catch(() => 'not_found');

  if (winner === 'not_found') {
    throw new Error('NOT_FOUND: Kein Account für diese UUID in SF');
  }
  console.log(`  🔍 Account gefunden — klicke`);
  await resultLink.click();

  // Warten bis der Salesforce Flow vollständig geladen ist bevor Checkbox gesucht wird
  await page.locator('flowruntime-lwc-body').waitFor({ timeout: 20000 });
  await page.waitForTimeout(1000);

  // Zweite Checkbox im Flow = "Conversation Log" (erste = DSAR)
  const convLogCheckbox = page.locator('flowruntime-lwc-body .slds-checkbox_faux').nth(1);
  await convLogCheckbox.waitFor({ timeout: 15000 });
  await convLogCheckbox.click();

  // 2x Next — nach erstem Next warten bis zweiter Next-Button wirklich da ist
  await page.getByText('Next', { exact: true }).click();
  await page.getByText('Next', { exact: true }).waitFor({ timeout: 10000 });
  await page.getByText('Next', { exact: true }).click();

  // Nach Finish: Record-ID Link erscheint als anklickbarer Text
  await page.getByText('Finish', { exact: true }).waitFor({ timeout: 15000 });
  const recordLink = page.locator('a[href*="/lightning/r/"]').last();
  await recordLink.waitFor({ timeout: 15000 });
  const hyperlink = await recordLink.getAttribute('href');
  const fullUrl = hyperlink.startsWith('http') ? hyperlink : `https://tipico.lightning.force.com${hyperlink}`;
  console.log(`  🔗 Link: ${fullUrl}`);

  await page.getByText('Finish', { exact: true }).click();

  // Tabs nach dem Finish schließen — sauberer Start für nächste UUID
  await closeAllSfTabs(page);

  return fullUrl;
}

async function run() {
  const { rows } = await readGuids();

  // Input-Zeilen aufsplitten: pro UUID eine eigene Zeile, PersonaID bleibt erhalten
  const outputRows = [];
  for (const row of rows) {
    const personaId = row[CONFIG.personaIdColumn];
    const uuidCell = String(row[CONFIG.guidColumn] || '').trim();
    if (!uuidCell) continue;
    const uuids = uuidCell.split(':').map(u => u.trim()).filter(Boolean);
    for (const uuid of uuids) {
      // Resume: bereits verarbeitete überspringen
      const existing = outputRows.find(r => r.UUID === uuid);
      if (existing) continue;
      outputRows.push({ UUID: uuid, URL: '', PersonaID: personaId });
    }
  }

  // Bereits gespeicherte Ergebnisse laden (Resume bei Abbruch)
  if (fs.existsSync(OUTPUT_PATH)) {
    const savedWb = XLSX.readFile(OUTPUT_PATH);
    const savedWs = savedWb.Sheets['Results'];
    const existing = savedWs ? XLSX.utils.sheet_to_json(savedWs) : [];
    for (const saved of existing) {
      const match = outputRows.find(r => r.UUID === saved.UUID);
      if (match && saved.URL && !String(saved.URL).startsWith('ERROR')) {
        match.URL = saved.URL;
      }
    }
    console.log(`📂 Bestehende Ergebnisse geladen (Resume)`);
  }

  console.log(`📋 ${outputRows.length} UUIDs total (aus ${rows.length} Zeilen)\n`);

  const browser = await chromium.launch({ headless: CONFIG.headless });
  const context = await browser.newContext();
  const page = await context.newPage();

  console.log(`\n🌐 Browser öffnet sich — bitte einloggen (MFA abschließen)...`);
  await page.goto(CONFIG.loginUrl);
  await page.waitForURL(/lightning\.force\.com/, { timeout: 180000 });
  await page.goto(CONFIG.homeUrl);
  await page.waitForLoadState('domcontentloaded');
  console.log('✅ Eingeloggt — räume Tabs auf...');
  await closeAllSfTabs(page);
  console.log('✅ Tabs bereinigt — starte Verarbeitung\n');

  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < outputRows.length; i++) {
    const item = outputRows[i];

    // Resume: bereits erfolgreich verarbeitet
    if (item.URL && !String(item.URL).startsWith('ERROR')) {
      console.log(`[${i + 1}/${outputRows.length}] ${item.UUID} → bereits vorhanden, überspringe`);
      continue;
    }

    console.log(`\n[${i + 1}/${outputRows.length}] UUID: ${item.UUID} (PersonaID: ${item.PersonaID})`);

    try {
      const hyperlink = await processGuid(page, item.UUID);
      item.URL = hyperlink;
      successCount++;
    } catch (err) {
      if (err.message.startsWith('NOT_FOUND')) {
        console.log(`  ⚪ Kein Account gefunden — überspringe`);
        item.URL = 'NOT_FOUND';
        errorCount++;
      } else {
        console.error(`  ❌ Fehler: ${err.message}`);
        item.URL = `ERROR: ${err.message}`;
        errorCount++;

        const screenshotPath = `./data/error_${i + 1}_${Date.now()}.png`;
        await page.screenshot({ path: screenshotPath, fullPage: true });
        console.log(`  📸 Screenshot: ${screenshotPath}`);

        await page.goto(CONFIG.homeUrl);
        await page.waitForLoadState('domcontentloaded');
        await closeAllSfTabs(page);
      }
    }

    // Nach jeder UUID speichern
    await saveProgress(outputRows);
  }

  await browser.close();

  console.log(`\n${'─'.repeat(50)}`);
  console.log(`✅ Fertig: ${successCount} erfolgreich, ${errorCount} Fehler`);
  console.log(`📁 Ergebnis gespeichert: ${OUTPUT_PATH}`);
  if (errorCount > 0) {
    console.log(`⚠️  Fehlerhafte GUIDs: führe das Skript erneut aus — sie werden automatisch wiederholt`);
  }
}

run().catch(console.error);
