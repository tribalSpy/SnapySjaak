/**
 * fust_out_submit.js
 *
 * Leest de "Overzicht" tab uit Fust_Week_27_overzicht.xlsx en zet voor elke
 * dag x carrier combinatie een "OUT" record in https://snapysjaak.onrender.com/
 * (Country = FR, Klantnaam/carrier per kolomgroep, Connect = auto, DC/DCS/DCO
 * uit de tabel, No CMR aangevinkt).
 *
 * Vereisten:
 *   npm install playwright exceljs
 *   npx playwright install chromium   (eenmalig, download browser binary)
 *
 * Gebruik:
 *   SNAPPY_USER=Master SNAPPY_PASS=*** node fust_out_submit.js pad/naar/Overzicht.xlsx
 *
 * Credentials NOOIT hardcoden in dit bestand — altijd via environment variables.
 *
 * Gedrag:
 *   - Combinaties met DC=DCS=DCO=0 worden overgeslagen (geen zin een leeg record te posten).
 *   - Bij een fout op 1 combinatie stopt het script niet; het logt de fout en gaat door.
 *   - Aan het einde staat een samenvatting: hoeveel gelukt, overgeslagen, gefaald.
 *   - DRY_RUN=1 vult het formulier wel maar klikt niet op "Save OUT" (handig om een keer te controleren
 *     voordat je het echt laat opslaan).
 */

const { chromium } = require('playwright');
const ExcelJS = require('exceljs');
const path = require('path');

const BASE_URL = 'https://snapysjaak.onrender.com/';
const COUNTRY = 'FR';
const DRY_RUN = process.env.DRY_RUN === '1';

// Kolomgroep in de Overzicht-tab -> klantnaam zoals de site die kent.
// Let op: de site noemt "ML Express" klant "ML Express Parijs".
const CARRIER_GROUPS = [
  { sheetLabel: 'Breewel',    siteKlant: 'Breewel' },
  { sheetLabel: 'ML Express', siteKlant: 'ML Express Parijs' },
  { sheetLabel: 'De Wit',     siteKlant: 'De Wit' },
  { sheetLabel: 'De Wit 2',   siteKlant: 'De Wit 2' },
];

async function readOverzicht(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.getWorksheet('Overzicht');
  if (!ws) throw new Error('Tab "Overzicht" niet gevonden in ' + filePath);

  // Header row 1 = groepnamen (gemerged), row 2 = FustDC/FustDCS/FustDCO per groep.
  // Kolom A = datum-label, data start op rij 3.
  const headerRow1 = ws.getRow(1).values; // 1-indexed sparse array
  const groupStartCol = {}; // sheetLabel -> kolomindex van FustDC binnen die groep
  let col = 2; // B
  for (const g of ['Breewel', 'ML Express', 'De Wit', 'De Wit 2', 'Totaal']) {
    groupStartCol[g] = col;
    col += 3;
  }

  const records = [];
  const dateRegex = /(\d{2})-(\d{2})-(\d{4})/;

  ws.eachRow((row, rowNumber) => {
    if (rowNumber < 3) return; // skip headers
    const labelCell = row.getCell(1).value;
    if (!labelCell || typeof labelCell !== 'string') return;
    if (labelCell.toLowerCase().startsWith('totaal')) return; // skip totaalrij

    const m = labelCell.match(dateRegex);
    if (!m) return;
    const isoDate = `${m[3]}-${m[2]}-${m[1]}`; // YYYY-MM-DD voor <input type=date>

    for (const g of CARRIER_GROUPS) {
      const startCol = groupStartCol[g.sheetLabel];
      const dc = Number(row.getCell(startCol).value) || 0;
      const dcs = Number(row.getCell(startCol + 1).value) || 0;
      const dco = Number(row.getCell(startCol + 2).value) || 0;
      records.push({
        date: isoDate,
        dateLabel: labelCell,
        siteKlant: g.siteKlant,
        dc, dcs, dco,
      });
    }
  });

  return records;
}

async function submitOne(page, rec) {
  await page.goto(BASE_URL, { waitUntil: 'networkidle' });

  // Date
  await page.locator('form input[type=date]').fill(rec.date);

  // Country -> FR (dit triggert dat Klantnaam de juiste opties toont)
  await page.locator('form select').nth(0).selectOption(COUNTRY);

  // Klantnaam/carrier -> wachten tot de juiste optie beschikbaar is
  const klantSelect = page.locator('form select').nth(1);
  await klantSelect.selectOption({ label: rec.siteKlant });

  // Connect -> automatisch maar 1 optie beschikbaar, moet nog wel expliciet gekozen worden
  const connectSelect = page.locator('form select').nth(2);
  const connectOptions = await connectSelect.locator('option').all();
  // index 0 = "Choose connect" placeholder, index 1 = de enige echte optie
  if (connectOptions.length < 2) {
    throw new Error(`Geen Connect-optie gevonden voor klant "${rec.siteKlant}"`);
  }
  const connectValue = await connectOptions[1].getAttribute('value');
  await connectSelect.selectOption(connectValue);

  // DC / DCS / DCO (Remark, Fustbon, Fustfactuur, CCTAG, PAL, VK blijven leeg)
  const numberInputs = page.locator('form input[type=number]');
  await numberInputs.nth(0).fill(String(rec.dc));   // DC
  await numberInputs.nth(2).fill(String(rec.dcs));  // DCS
  await numberInputs.nth(3).fill(String(rec.dco));  // DCO

  // No CMR aanvinken
  await page.locator('form input[type=checkbox]').check();

  if (DRY_RUN) {
    return { status: 'dry-run' };
  }

  await page.locator('form button[type=submit]', { hasText: 'Save OUT' }).click();
  // Wacht op een reactie van de pagina (success message of hervalidatie van het formulier).
  await page.waitForTimeout(800);
  return { status: 'submitted' };
}

async function login(page) {
  const user = process.env.SNAPPY_USER;
  const pass = process.env.SNAPPY_PASS;
  await page.goto(BASE_URL, { waitUntil: 'networkidle' });

  const isLoginPage = await page.locator('input[type=password]').count() > 0;
  if (!isLoginPage) return; // al ingelogd (bestaande sessie/cookie)

  if (!user || !pass) {
    throw new Error(
      'Niet ingelogd en geen SNAPPY_USER/SNAPPY_PASS environment variables gezet. ' +
      'Start het script als: SNAPPY_USER=... SNAPPY_PASS=... node fust_out_submit.js <bestand.xlsx>'
    );
  }

  await page.locator('input[autocomplete=username]').fill(user);
  await page.locator('input[type=password]').fill(pass);
  await page.getByRole('button', { name: 'Log in' }).click();
  await page.waitForLoadState('networkidle');
}

async function main() {
  const filePath = process.argv[2];
  if (!filePath) {
    console.error('Gebruik: node fust_out_submit.js pad/naar/Overzicht.xlsx');
    process.exit(1);
  }

  const records = await readOverzicht(path.resolve(filePath));
  console.log(`${records.length} combinaties gevonden in het Overzicht-bestand.`);

  const browser = await chromium.launch({ headless: !process.env.HEADFUL });
  const page = await browser.newPage();

  await login(page);

  const summary = { submitted: [], skipped: [], failed: [] };

  for (const rec of records) {
    const label = `${rec.dateLabel} | ${rec.siteKlant} | DC=${rec.dc} DCS=${rec.dcs} DCO=${rec.dco}`;

    if (rec.dc === 0 && rec.dcs === 0 && rec.dco === 0) {
      console.log(`SKIP (alles 0): ${label}`);
      summary.skipped.push(label);
      continue;
    }

    try {
      const result = await submitOne(page, rec);
      console.log(`OK (${result.status}): ${label}`);
      summary.submitted.push(label);
    } catch (err) {
      console.error(`FOUT: ${label} -> ${err.message}`);
      summary.failed.push({ label, error: err.message });
    }
  }

  await browser.close();

  console.log('\n--- Samenvatting ---');
  console.log(`Verzonden: ${summary.submitted.length}`);
  console.log(`Overgeslagen (alles 0): ${summary.skipped.length}`);
  console.log(`Gefaald: ${summary.failed.length}`);
  if (summary.failed.length) {
    console.log('Gefaalde combinaties:');
    summary.failed.forEach(f => console.log(`  - ${f.label}: ${f.error}`));
  }
}

main().catch(err => {
  console.error('Onverwachte fout:', err);
  process.exit(1);
});
