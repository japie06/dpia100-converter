// UI tests for dpia100.html with Playwright in Microsoft Edge: npm run test:ui.
// The main route is the one that always works: the plain file input and a Blob download, with the File
// System Access dialogs switched off (as under a blocking policy or in Firefox). The dialog route
// (Edge and Chrome) is tested separately with stubbed dialogs, because no tool can drive the native ones.
// Byte-for-byte parity on all cases is tested by tests/run-tests.ps1.
import { test, expect } from '@playwright/test';
import { readFileSync, copyFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import path from 'node:path';
import { pagePath, casePath, repoRoot, loadCore } from './core.mjs';

const pageUrl = pathToFileURL(pagePath).href;
const bytesOf = file => readFileSync(file);
const base64Of = file => bytesOf(file).toString('base64');

// A copy of a case file under a chosen name, as a user would have it.
function fixture(testInfo, caseFile, name) {
  const target = testInfo.outputPath(name);
  copyFileSync(caseFile, target);
  return target;
}

// Every test: the page must start under its own CSP, without script errors or CSP violations.
test.beforeEach(async ({ page }, testInfo) => {
  testInfo.problems = [];
  page.on('pageerror', e => testInfo.problems.push(`page error: ${e.message}`));
  page.on('console', m => {
    if (/Content Security Policy|Refused to|TrustedHTML/.test(m.text())) testInfo.problems.push(`console: ${m.text()}`);
  });
});
test.afterEach(async ({}, testInfo) => {
  expect(testInfo.problems, 'no CSP violations or script errors').toEqual([]);
});

async function withoutDialogs(page) {
  await page.addInitScript(() => {
    for (const name of ['showOpenFilePicker', 'showSaveFilePicker']) {
      Object.defineProperty(window, name, { value: undefined, configurable: true });
    }
  });
}

async function chooseFile(page, file) {
  const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.locator('#drop').click()]);
  await chooser.setFiles(file);
  await expect(page.locator('#check-status .banner')).toBeVisible();
}

async function fillForm(page, gebruiker = 'TSTUSER') {
  await page.fill('#f-proces', 'cmt');
  await page.fill('#f-datum', '2026-10-07');
  await page.fill('#f-gebruiker', gebruiker);
}

async function downloadFile(page) {
  const [download] = await Promise.all([page.waitForEvent('download'), page.locator('#btn-make').click()]);
  return { name: download.suggestedFilename(), bytes: readFileSync(await download.path()) };
}

test.describe('file input and download (works everywhere)', () => {
  test.beforeEach(async ({ page }) => {
    await withoutDialogs(page);
    await page.goto(pageUrl);
  });

  test('starts, shows the version, privacy section collapsed until the header link is used', async ({ page }) => {
    await expect(page.locator('#startup-error')).toHaveCount(0);
    await expect(page.locator('#version')).toHaveText(`v${loadCore().VERSION}`);
    await expect(page.locator('#privacy')).not.toHaveAttribute('open');
    await page.locator('.privacy-pill').click();
    await expect(page.locator('#privacy')).toHaveAttribute('open');
    await expect(page.locator('#step-check')).toBeHidden();
  });

  test('voorbeeld.csv gives the exact bytes of the Excel tool layout; Onthoud survives a reload', async ({ page }, testInfo) => {
    await chooseFile(page, fixture(testInfo, path.join(repoRoot, 'voorbeeld.csv'), 'voorbeeld.csv'));
    await expect(page.locator('#check-status')).toContainText('3 regels gecontroleerd, geen fouten');
    await expect(page.locator('.tbl-mapping th[scope="col"]')).toHaveCount(7);
    await expect(page.locator('#check-body')).toContainText('125,50 → 125.50');
    await expect(page.locator('#confirm-box')).toBeHidden();
    await expect(page.locator('#btn-make')).toHaveText('Maak DPIA100_voorbeeld.txt');

    await fillForm(page);
    await page.check('#f-remember');
    const file = await downloadFile(page);
    expect(file.name).toBe('DPIA100_voorbeeld.txt');
    expect(file.bytes.equals(bytesOf(casePath('ok-voorbeeld/expected.txt')))).toBe(true);
    await expect(page.locator('#make-result')).toContainText('Gedownload als DPIA100_voorbeeld.txt');
    await expect(page.locator('#make-result')).toContainText('DPIA100_voorbeeld (1).txt');

    // After a reload step 3 is hidden until a CSV is loaded, but the fields are already filled in.
    await page.reload();
    await expect(page.locator('#f-remember')).toBeChecked();
    await expect(page.locator('#f-gebruiker')).toHaveValue('TSTUSER');
    await expect(page.locator('#f-proces')).toHaveValue('CMT');
    await page.locator('#f-remember').dispatchEvent('click');
    expect(await page.evaluate(() => localStorage.getItem('dpia100.instellingen'))).toBeNull();
  });

  test('a CSV with errors shows every error and no step 3', async ({ page }, testInfo) => {
    await chooseFile(page, fixture(testInfo, casePath('err-persnr/input.csv'), 'fouten.csv'));
    await expect(page.locator('#check-status')).toContainText('6 fouten gevonden. Er is geen bestand gemaakt.');
    await expect(page.locator('.tbl-errors tbody tr')).toHaveCount(6);
    await expect(page.locator('.tbl-errors')).toContainText('Alleen cijfers toegestaan');
    await expect(page.locator('#step-make')).toBeHidden();
  });

  test('a missing column is reported with the mapping', async ({ page }, testInfo) => {
    await chooseFile(page, fixture(testInfo, casePath('err-missing-column/input.csv'), 'kolom.csv'));
    await expect(page.locator('#check-status')).toContainText('Er ontbreekt een verplichte kolom: S of V');
    await expect(page.locator('.tbl-mapping .missing')).toHaveText('(niet gevonden)');
  });

  test('warnings need the confirmation before the file is made', async ({ page }, testInfo) => {
    await chooseFile(page, fixture(testInfo, casePath('ok-waarde-truncate/input.csv'), 'lang.csv'));
    await expect(page.locator('.warn-list li')).toHaveCount(1);
    await expect(page.locator('#confirm-box')).toBeVisible();
    await fillForm(page);
    let downloaded = false;
    page.on('download', () => { downloaded = true; });
    await page.locator('#btn-make').click();
    await expect(page.locator('#f-confirm-err')).toBeVisible();
    expect(downloaded).toBe(false);

    await page.check('#f-confirm');
    const file = await downloadFile(page);
    expect(file.name).toBe('DPIA100_lang.txt');
    expect(file.bytes.equals(bytesOf(casePath('ok-waarde-truncate/expected.txt')))).toBe(true);
  });

  test('the form checks procescode and creatiegebruiker', async ({ page }, testInfo) => {
    await chooseFile(page, fixture(testInfo, path.join(repoRoot, 'voorbeeld.csv'), 'voorbeeld.csv'));
    await page.fill('#f-proces', 'TO?');
    await page.fill('#f-gebruiker', '');
    await page.locator('#btn-make').click();
    await expect(page.locator('#f-proces-err')).toContainText('1 tot 3 letters of cijfers');
    await expect(page.locator('#f-gebruiker-err')).toContainText('Vul je gebruikersnaam');
    await expect(page.locator('#f-proces')).toBeFocused();
  });

  test('a file dropped on the page is loaded', async ({ page }) => {
    await page.evaluate(b64 => {
      const dt = new DataTransfer();
      dt.items.add(new File([Uint8Array.from(atob(b64), c => c.charCodeAt(0))], 'gesleept.csv'));
      document.getElementById('drop').dispatchEvent(new DragEvent('drop', { dataTransfer: dt, bubbles: true, cancelable: true }));
    }, base64Of(path.join(repoRoot, 'voorbeeld.csv')));
    await expect(page.locator('#check-status')).toContainText('geen fouten');
    await expect(page.locator('#file-name')).toHaveText('gesleept.csv');
  });
});

test.describe('open and save dialogs (Edge and Chrome, stubbed)', () => {
  // Fake file handles: the open handle reads window.__disk, so a test can 'fix the file in Excel'; the
  // save handle records what is written.
  test.beforeEach(async ({ page }) => {
    await page.addInitScript(() => {
      const fromBase64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
      window.__disk = { name: 'voorbeeld.csv', base64: '', modified: 1000 };
      window.__save = { name: null, written: null, removed: false, suggested: null };
      const openHandle = {
        kind: 'file',
        get name() { return window.__disk.name; },
        getFile: async () => new File([fromBase64(window.__disk.base64)], window.__disk.name, { lastModified: window.__disk.modified })
      };
      window.showOpenFilePicker = async () => [openHandle];
      window.showSaveFilePicker = async opts => {
        window.__save.suggested = opts.suggestedName;
        const name = window.__save.name || opts.suggestedName;
        return {
          kind: 'file', name,
          getFile: async () => new File([], name, { lastModified: Date.now() }),
          remove: async () => { window.__save.removed = true; },
          createWritable: async () => {
            const chunks = [];
            return {
              write: async d => { chunks.push(d); },
              close: async () => { window.__save.written = Array.from(new Uint8Array(await new Blob(chunks).arrayBuffer())); }
            };
          }
        };
      };
    });
    await page.goto(pageUrl);
  });

  const putOnDisk = (page, name, file) =>
    page.evaluate(([n, b64]) => { Object.assign(window.__disk, { name: n, base64: b64 }); window.__disk.modified += 1000; }, [name, base64Of(file)]);
  const saved = page => page.evaluate(() => window.__save);

  test('saves through the dialog with the exact bytes', async ({ page }) => {
    await putOnDisk(page, 'voorbeeld.csv', path.join(repoRoot, 'voorbeeld.csv'));
    await page.locator('#drop').click();
    await expect(page.locator('#check-status')).toContainText('geen fouten');
    await fillForm(page);
    await page.locator('#btn-make').click();
    await expect(page.locator('#make-result')).toContainText('Opgeslagen als DPIA100_voorbeeld.txt');
    const s = await saved(page);
    expect(s.suggested).toBe('DPIA100_voorbeeld.txt');
    expect(Buffer.from(s.written).equals(bytesOf(casePath('ok-voorbeeld/expected.txt')))).toBe(true);
  });

  test('Opnieuw controleren reads the fixed file again', async ({ page }) => {
    await putOnDisk(page, 'fouten.csv', casePath('err-persnr/input.csv'));
    await page.locator('#drop').click();
    await expect(page.locator('#check-status')).toContainText('6 fouten gevonden');
    await putOnDisk(page, 'fouten.csv', path.join(repoRoot, 'voorbeeld.csv'));
    await page.locator('#btn-recheck').click();
    await expect(page.locator('#check-status')).toContainText('geen fouten');
    await expect(page.locator('#step-make')).toBeVisible();
  });

  test('refuses a name without DPIA100 and removes the empty file the dialog made', async ({ page }) => {
    await putOnDisk(page, 'voorbeeld.csv', path.join(repoRoot, 'voorbeeld.csv'));
    await page.locator('#drop').click();
    await fillForm(page);
    await page.evaluate(() => { window.__save.name = 'mutaties.txt'; });
    await page.locator('#btn-make').click();
    await expect(page.locator('#make-result')).toContainText('begint niet met DPIA100');
    const s = await saved(page);
    expect(s.written).toBeNull();
    expect(s.removed).toBe(true);
  });

  test('refuses to build from a CSV that changed after the check', async ({ page }) => {
    await putOnDisk(page, 'voorbeeld.csv', path.join(repoRoot, 'voorbeeld.csv'));
    await page.locator('#drop').click();
    await fillForm(page);
    await page.evaluate(() => { window.__disk.modified += 1000; });
    await page.locator('#btn-make').click();
    await expect(page.locator('#make-result')).toContainText('gewijzigd sinds de controle');
    expect((await saved(page)).written).toBeNull();
  });
});
