/**
 * Flow 2 — Salesforce UUID → Conversation Log File Downloader
 *
 * KONFIGURATION: Nur diesen Block anpassen
 */
import * as XLSX from 'xlsx';
import { chromium } from 'playwright';
import fs from 'fs';
import os from 'os';
import path from 'path';

const CONFIG = {
  // Bereinigtes Output-Workbook aus Flow 1.
  // Es dient nur noch als Filter: verarbeitet werden nur UUIDs mit URL in Spalte B.
  excelInputPath: './CommsLogs_UUID_Lists/Template_Automation_Feeder_List_CommsLogs_output.xlsx',

  // Name des Tabellenblatts (oder null für erstes Sheet)
  sheetName: 'Results',

  // Spaltennamen
  uuidColumn: 'UUID',
  personaIdColumn: 'PersonaID',
  linkColumn: 'URL',

  // Salesforce Login-URL (Microsoft SSO)
  loginUrl: 'https://tipico.lightning.force.com',

  // SF Home URL — nach Login wird hierhin navigiert
  homeUrl: 'https://tipico.lightning.force.com/lightning/page/home',

  // Browser sichtbar lassen (MFA erfordert manuellen Login)
  headless: false,

  // Zielpfad für Downloads: ~/Downloads/<PersonaID>/<UUID>/
  downloadRootPath: path.join(os.homedir(), 'Downloads'),

  // Fortschritt für Resume
  progressPath: './data/flow2-conversation-log-download-progress.json',

  // Optionales Download-Limit; `null` = alle angezeigten Files herunterladen
  maxFilesPerUuid: null,

  // Timeouts
  searchTimeoutMs: 15000,
  pageTimeoutMs: 20000,
  downloadTimeoutMs: 180000,
};

function isUrl(value) {
  return /^https?:\/\//i.test(String(value ?? '').trim());
}

function sanitizePathSegment(value) {
  return String(value ?? '')
    .trim()
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
    .replace(/\s+/g, ' ')
    .slice(0, 120);
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function loadProgress() {
  if (!fs.existsSync(CONFIG.progressPath)) {
    return {};
  }

  return JSON.parse(fs.readFileSync(CONFIG.progressPath, 'utf8'));
}

function saveProgress(progress) {
  ensureDir(path.dirname(CONFIG.progressPath));
  fs.writeFileSync(CONFIG.progressPath, JSON.stringify(progress, null, 2));
}

function uniqueFilePath(dirPath, fileName) {
  const parsed = path.parse(fileName);
  let candidate = path.join(dirPath, fileName);
  let counter = 1;

  while (fs.existsSync(candidate)) {
    const nextName = `${parsed.name}_${counter}${parsed.ext}`;
    candidate = path.join(dirPath, nextName);
    counter++;
  }

  return candidate;
}

function screenshotPath(index) {
  ensureDir('./data');
  return `./data/flow2_error_${index}_${Date.now()}.png`;
}

function setSheetCell(sheet, address, value) {
  if (value === '' || value === null || value === undefined) {
    delete sheet[address];
    return;
  }

  sheet[address] = { t: typeof value === 'number' ? 'n' : 's', v: value };
}

function ensureSheetRangeIncludes(sheet, address) {
  const cell = XLSX.utils.decode_cell(address);

  if (!sheet['!ref']) {
    sheet['!ref'] = XLSX.utils.encode_range({ s: cell, e: cell });
    return;
  }

  const range = XLSX.utils.decode_range(sheet['!ref']);
  range.s.r = Math.min(range.s.r, cell.r);
  range.s.c = Math.min(range.s.c, cell.c);
  range.e.r = Math.max(range.e.r, cell.r);
  range.e.c = Math.max(range.e.c, cell.c);
  sheet['!ref'] = XLSX.utils.encode_range(range);
}

function writeCountHeaders(sheet) {
  ensureSheetRangeIncludes(sheet, 'D1');
  ensureSheetRangeIncludes(sheet, 'E1');
  setSheetCell(sheet, 'D1', 'IST');
  setSheetCell(sheet, 'E1', 'SOLL');
}

function writeWorkbook(workbook) {
  XLSX.writeFile(workbook, CONFIG.excelInputPath);
}

function updateWorkbookCounts(workbook, rowNumber, actualCount, expectedCount) {
  const sheetName = CONFIG.sheetName ?? workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  writeCountHeaders(sheet);

  ensureSheetRangeIncludes(sheet, `D${rowNumber}`);
  ensureSheetRangeIncludes(sheet, `E${rowNumber}`);
  setSheetCell(sheet, `D${rowNumber}`, actualCount);
  setSheetCell(sheet, `E${rowNumber}`, expectedCount);

  writeWorkbook(workbook);
}

function readInputRows() {
  const workbook = XLSX.readFile(CONFIG.excelInputPath);
  const sheetName = CONFIG.sheetName ?? workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const deduped = new Map();
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');

  writeCountHeaders(sheet);
  writeWorkbook(workbook);

  for (let rowIndex = range.s.r + 1; rowIndex <= range.e.r; rowIndex++) {
    const rowNumber = rowIndex + 1;
    const uuid = String(sheet[`A${rowNumber}`]?.v ?? '').trim();
    const linkValue = String(sheet[`B${rowNumber}`]?.v ?? '').trim();
    const personaId = String(sheet[`C${rowNumber}`]?.v ?? '').trim();

    if (!uuid || !personaId || !isUrl(linkValue)) {
      continue;
    }

    if (!deduped.has(uuid)) {
      deduped.set(uuid, {
        UUID: uuid,
        PersonaID: personaId,
        rowNumber,
      });
    }
  }

  const filteredRows = [...deduped.values()];
  console.log(`📋 ${filteredRows.length} UUIDs mit befüllter URL-Spalte geladen`);
  return {
    workbook,
    rows: filteredRows,
  };
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
    console.log('  🧹 Keine zusätzlichen SF-Tabs offen');
  }

  await page.waitForTimeout(500);
}

async function searchUuid(page, uuid) {
  await page.keyboard.press('/');

  const searchBox = page.getByRole('searchbox', { name: 'Search...' });
  await searchBox.waitFor({ timeout: CONFIG.searchTimeoutMs });
  await searchBox.click();
  await searchBox.fill('');
  await searchBox.fill(uuid);
  await searchBox.press('Enter');

  await page.getByRole('heading', { name: 'Search Results' }).waitFor({ timeout: CONFIG.pageTimeoutMs });
}

async function openFirstConversationLog(page) {
  const objectLink = page.getByRole('link', { name: 'Conversation Logs', exact: true }).first();

  try {
    await objectLink.waitFor({ timeout: 10000 });
    await objectLink.click();
  } catch {
    console.log('  ℹ️ "Conversation Logs"-Link nicht separat klickbar — versuche Direktzugriff auf ersten Treffer');
  }

  const directLink = page.getByRole('link', { name: /^CONV-/ }).first();

  if (await directLink.isVisible().catch(() => false)) {
    await directLink.click();
    return;
  }

  const section = page.locator('section, article, div').filter({
    has: page.getByText('Conversation Logs', { exact: true }),
  }).first();
  const sectionLink = section.getByRole('link', { name: /^CONV-/ }).first();

  await sectionLink.waitFor({ timeout: CONFIG.pageTimeoutMs });
  await sectionLink.click();
}

async function openRelatedTab(page) {
  const relatedTab = page.getByRole('tab', { name: 'Related' });
  await relatedTab.waitFor({ timeout: CONFIG.pageTimeoutMs });
  await relatedTab.click();
  await page.waitForTimeout(1500);
}

async function getAvailableFileCount(page) {
  const filesText = await page.locator('text=/Files\\s*\\(\\d+\\)/').first().textContent().catch(() => null);

  if (!filesText) {
    return null;
  }

  const match = filesText.match(/Files\s*\((\d+)\)/i);
  return match ? Number(match[1]) : null;
}

function getFileButtons(page) {
  return page.getByRole('button', { name: /(EmailLogs|ChatLogs)/i });
}

async function getFileButtonLabels(page) {
  const fileButtons = getFileButtons(page);
  const availableCount = await fileButtons.count();
  const count = CONFIG.maxFilesPerUuid == null
    ? availableCount
    : Math.min(availableCount, CONFIG.maxFilesPerUuid);
  const labels = [];

  for (let i = 0; i < count; i++) {
    labels.push((await fileButtons.nth(i).innerText()).trim());
  }

  return labels;
}

async function closePreview(page) {
  const closeButton = page.getByRole('button', { name: 'Close' });
  await closeButton.waitFor({ timeout: 10000 });
  await closeButton.click();
  await page.waitForTimeout(500);
}

async function downloadCurrentFile(page, destinationDir, fallbackName) {
  const downloadButton = page.locator('#previewStatus').getByRole('button', { name: 'Download' });
  await downloadButton.waitFor({ timeout: CONFIG.pageTimeoutMs });

  const [popup] = await Promise.all([
    page.waitForEvent('popup'),
    downloadButton.click(),
  ]);

  const download = await popup.waitForEvent('download', {
    timeout: CONFIG.downloadTimeoutMs,
  });

  const suggestedName = sanitizePathSegment(download.suggestedFilename()) || `${fallbackName}.bin`;
  const targetPath = uniqueFilePath(destinationDir, suggestedName);
  await download.saveAs(targetPath);

  await popup.close().catch(() => {});

  return targetPath;
}

async function downloadFilesForUuid(page, item) {
  const destinationDir = path.join(
    CONFIG.downloadRootPath,
    sanitizePathSegment(`CommsLogs_${item.PersonaID}`),
    sanitizePathSegment(item.UUID),
  );
  ensureDir(destinationDir);

  await searchUuid(page, item.UUID);
  await openFirstConversationLog(page);
  await openRelatedTab(page);

  const expectedFileCount = await getAvailableFileCount(page);
  const labels = await getFileButtonLabels(page);

  if (expectedFileCount === null) {
    throw new Error('FEHLSCHLAG: Files(n) konnte nicht gelesen werden');
  }

  console.log(`  📊 Files-SOLL laut Salesforce: ${expectedFileCount}`);

  if (labels.length === 0) {
    return {
      status: 'no_files',
      files: [],
      destinationDir,
      actualFileCount: 0,
      expectedFileCount,
    };
  }

  const downloadedFiles = [];

  for (let i = 0; i < labels.length; i++) {
    const fileButton = getFileButtons(page).nth(i);
    const label = labels[i];
    const fallbackName = `file_${i + 1}`;

    console.log(`  📄 Datei ${i + 1}/${labels.length}: ${label}`);
    await fileButton.click();

    const savedPath = await downloadCurrentFile(page, destinationDir, fallbackName);
    downloadedFiles.push(savedPath);
    console.log(`  ⬇️ Gespeichert: ${savedPath}`);

    await closePreview(page);
  }

  return {
    status: 'done',
    files: downloadedFiles,
    destinationDir,
    actualFileCount: downloadedFiles.length,
    expectedFileCount,
  };
}

async function run() {
  const { workbook, rows } = readInputRows();
  const progress = loadProgress();

  const browser = await chromium.launch({ headless: CONFIG.headless });
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();

  console.log('\n🌐 Browser öffnet sich — bitte einloggen (MFA abschließen)...');
  await page.goto(CONFIG.loginUrl);
  await page.waitForURL(/lightning\.force\.com/, { timeout: 180000 });
  await page.goto(CONFIG.homeUrl);
  await page.waitForLoadState('domcontentloaded');
  console.log('✅ Eingeloggt — räume Tabs auf...');
  await closeAllSfTabs(page);
  console.log('✅ Tabs bereinigt — starte Downloads\n');

  let successCount = 0;
  let noFilesCount = 0;
  let errorCount = 0;

  for (let i = 0; i < rows.length; i++) {
    const item = rows[i];
    const existing = progress[item.UUID];

    if (
      (existing?.status === 'done' || existing?.status === 'no_files')
      && typeof existing.actualFileCount === 'number'
      && typeof existing.expectedFileCount === 'number'
    ) {
      updateWorkbookCounts(workbook, item.rowNumber, existing.actualFileCount, existing.expectedFileCount);
      console.log(`[${i + 1}/${rows.length}] ${item.UUID} → bereits verarbeitet, überspringe`);
      continue;
    }

    console.log(`\n[${i + 1}/${rows.length}] UUID: ${item.UUID} (PersonaID: ${item.PersonaID})`);

    try {
      const result = await downloadFilesForUuid(page, item);
      progress[item.UUID] = {
        personaId: item.PersonaID,
        status: result.status,
        destinationDir: result.destinationDir,
        files: result.files,
        actualFileCount: result.actualFileCount,
        expectedFileCount: result.expectedFileCount,
        updatedAt: new Date().toISOString(),
      };
      updateWorkbookCounts(workbook, item.rowNumber, result.actualFileCount, result.expectedFileCount);

      if (result.status === 'done') {
        successCount++;
      } else {
        noFilesCount++;
        console.log('  ⚪ Keine Dateien gefunden');
      }
    } catch (err) {
      const error = err instanceof Error ? err : new Error(String(err));
      const shotPath = screenshotPath(i + 1);
      await page.screenshot({ path: shotPath, fullPage: true });

      progress[item.UUID] = {
        personaId: item.PersonaID,
        status: 'error',
        files: [],
        actualFileCount: 0,
        expectedFileCount: 'FEHLSCHLAG',
        error: error.message,
        screenshotPath: shotPath,
        updatedAt: new Date().toISOString(),
      };
      updateWorkbookCounts(workbook, item.rowNumber, 0, 'FEHLSCHLAG');

      errorCount++;
      console.error(`  ❌ Fehler: ${error.message}`);
      console.log(`  📸 Screenshot: ${shotPath}`);
    } finally {
      saveProgress(progress);
      await page.goto(CONFIG.homeUrl).catch(() => {});
      await page.waitForLoadState('domcontentloaded').catch(() => {});
      await closeAllSfTabs(page).catch(() => {});
    }
  }

  await browser.close();

  console.log(`\n${'─'.repeat(50)}`);
  console.log(`✅ Fertig: ${successCount} UUIDs mit Downloads, ${noFilesCount} ohne Dateien, ${errorCount} Fehler`);
  console.log(`📁 Download-Root: ${CONFIG.downloadRootPath}`);
  console.log(`📝 Progress-Datei: ${CONFIG.progressPath}`);
}

if (import.meta.main) {
  run().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}
