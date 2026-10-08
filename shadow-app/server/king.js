// King ERP export for FloraHolland invoices (Inkoop Controle > "Import naar
// King"). Pure functions only -- no I/O -- so the journal a user previews is
// byte-for-byte the journal that gets delivered.
//
// Formats (finance-provided sample king_journaal.xml + King's own docs):
// - Journal: KING_JOURNAAL > BOEKINGSGANGEN > BOEKINGSGANG > JOURNAALPOSTEN >
//   JOURNAALPOST (one per invoice) > JOURNAALREGELS > JOURNAALREGEL.
//   Line 1 CRED = the company's creditor account for the invoice total incl.
//   BTW; line 2/3 DEB = BTW laag/hoog; lines 4+ DEB = one per summary line,
//   to the ledger the "kingkoppel" table assigns.
// - Archive (KING_DIGITAAL_ARCHIEF): each PDF is referenced by its path on
//   the King server (DAR_BESTANDSNAAM), linked to the journal post through
//   DAR_EXTERN_ID == JR_ARCHIEFSTUK_EXTERN_ID (max 20 chars).

// The "kingkoppel" table as finance had it in the old SQL tool (PDF
// "kingkoppel info", 156 rows): FloraHolland Prd Id + BTW category (H hoog,
// S laag, O nul, E leeg) -> ledger for a flowers (GB FL) or plants (GB PL)
// booking. Plus finance's own per-klok-location keys for "Product
// aankopen" (1015_<Location>_klok), replacing Frank's GLN-based codes.
export const KING_LEDGER_SEED = [
  ["2395", "H", "140200", "140200", "Boxcontracten stapelwagens N Q1"],
  ["2036", "H", "260012", "260012", "Parkeren"],
  ["1433", "O", "302100", "302100", "Emballage statiegeld klok"],
  ["AI2_CT", "O", "302100", "302100", "Emballage meermalig statiegeld AI2"],
  ["CT", "O", "302100", "302100", "Verrekencode Fc566+0,90"],
  ["1411", "O", "303000", "303000", "Emballage statiegeld Connect"],
  ["1431", "O", "303000", "303000", "Emballage statiegeld klok"],
  ["1461", "O", "303000", "303000", "Emballage statiegeld depot"],
  ["1481", "O", "303000", "303000", "Emballage statiegeld Handel"],
  ["1491", "O", "303000", "303000", "Emballage statiegeld Handel"],
  ["2032", "H", "410000", "410000", "Huur"],
  ["2035", "H", "410000", "410000", "Gebruik Extra ruimte m2"],
  ["2043", "H", "410100", "410100", "Elektra Naaldwijk"],
  ["2049", "H", "410200", "410200", "Dienstverlening Beveiliging Alarm 1x"],
  ["2516", "H", "420300", "420300", "Trekker Vignet"],
  ["2263", "S", "433000", "433000", "Telefonie"],
  ["2216", "H", "460100", "460100", "Administratiekosten betalingen"],
  ["1381", "E", "460137", "460137", "Rente Betalingstermijn"],
  ["1381", "O", "460137", "460137", "Rente Betalingstermijn"],
  ["Connect", "O", "500000", "500000", "Product Aankopen Connect"],
  ["1005", "O", "500000", "510000", "Product aankopen"],
  ["1005", "S", "500000", "510000", "Product aankopen"],
  ["1015", "S", "500000", "510000", "Product aankopen"],
  ["1035", "S", "500000", "510000", "Product aankopen"],
  ["1535", "S", "500000", "510000", "Product aankopen"],
  ["8713783439074", "S", "500000", "510000", "Product Aankopen  NAALDWIJK KLOK"],
  ["8713783439074", "H", "500010", "500010", "Product Aankopen  NAALDWIJK KLOK"],
  ["1005", "H", "500010", "510010", "Product aankopen"],
  ["1015", "H", "500010", "510010", "Product aankopen"],
  ["1035", "H", "500010", "510010", "Product aankopen"],
  ["1535", "H", "500010", "510010", "Product aankopen"],
  ["8713783439043", "S", "500020", "510020", "Product Aankopen  AALSMEER KLOK"],
  ["8713783439043", "H", "500025", "510025", "Product Aankopen  AALSMEER KLOK"],
  ["8713783439081", "S", "500030", "510030", "Product Aankopen  RIJNSBURG KLOK"],
  ["8713783439081", "H", "500035", "500035", "Product Aankopen  RIJNSBURG KLOK"],
  ["Connect", "S", "500200", "510200", "Product Aankopen Connect"],
  ["Connect", "H", "500210", "510210", "Product Aankopen Connect"],
  ["8713783439043_KVV", "S", "500320", "510300", "KVV Product Aankopen  AALSMEER KLOK"],
  ["8713783439074_KVV", "S", "500320", "510300", "KVV Product Aankopen  NAALDWIJK KLOK"],
  ["8713783439081_KVV", "S", "500320", "510300", "KVV Product Aankopen  RIJNSBURG KLOK"],
  ["8713783439043_KVV", "H", "500325", "500325", "KVV Product Aankopen  AALSMEER KLOK"],
  ["8713783439074_KVV", "H", "500325", "510310", "KVV Product Aankopen  NAALDWIJK KLOK"],
  ["8713783439081_KVV", "H", "500325", "510310", "KVV Product Aankopen  RIJNSBURG KLOK"],
  ["Handelsregeling", "S", "500400", "500400", "Product Aankopen Handelsregeling"],
  ["Handelsregeling", "H", "500410", "500410", "Product Aankopen Handelsregeling"],
  ["AI2", "O", "510915", "510915", "Product Aankopen AI2"],
  ["AI2", "S", "510920", "510920", "Product Aankopen AI2"],
  ["AI2", "H", "510930", "510930", "Product Aankopen AI2"],
  ["1410", "H", "520100", "520100", "Emballage eenmalig Connect"],
  ["1410", "O", "520100", "520100", "Emballage eenmalig Connect"],
  ["1410", "S", "520100", "520100", "Emballage eenmalig Connect"],
  ["1430", "H", "520100", "520100", "Emballage eenmalig klok"],
  ["1430", "S", "520100", "520100", "Emballage eenmalig klok"],
  ["1460", "S", "520100", "520100", "Emballage eenmalig depot"],
  ["1480", "S", "520100", "520100", "Emballage eenmalig Handel"],
  ["1490", "H", "520100", "520100", "Emballage eenmalig Handel"],
  ["1490", "S", "520100", "520100", "Emballage eenmalig Handel"],
  ["2413", "H", "520100", "520100", "CL Opslag gekoeld N S.van der Vijver"],
  ["2486", "H", "520105", "520105", "Overveilkosten eigen rekening 1000"],
  ["AI2_PE", "S", "520105", "520105", "Emballage eenmalig AI2"],
  ["PE", "H", "520105", "520105", "Verrekencode 0,90"],
  ["PE", "O", "520105", "520105", "Statiegeldlegbord"],
  ["PE", "S", "520105", "520105", "Verrekencode Fc566+0,90"],
  ["1316", "S", "520500", "520500", "CC heffing klok"],
  ["2025", "H", "520500", "520500", "CC vergoeding (huur) 63155"],
  ["2026", "H", "520500", "520500", "CC opslag 63155"],
  ["2026", "S", "520500", "520500", "CC opslag 63155"],
  ["2027", "H", "520500", "520500", "CC handling 63155"],
  ["2027", "S", "520500", "520500", "CC handling 63155"],
  ["1412", "H", "520600", "520600", "Emballage huur Connect"],
  ["1412", "O", "520600", "520600", "Emballage huur Connect"],
  ["1412", "S", "520600", "520600", "Emballage huur Connect"],
  ["1432", "H", "520600", "520600", "Emballage huur klok"],
  ["1432", "S", "520600", "520600", "Emballage huur klok"],
  ["1462", "S", "520600", "520600", "Emballage huur depot"],
  ["1482", "S", "520600", "520600", "Emballage huur Handel"],
  ["1492", "H", "520600", "520600", "Emballage huur Handel"],
  ["1492", "S", "520600", "520600", "Emballage huur Handel"],
  ["2916", "S", "520600", "520600", "Korting Emballage huur 2024"],
  ["AI2_PC", "H", "520605", "520605", "Emballage meermalig huur AI2"],
  ["PC", "H", "520605", "520605", "Meermalig tray(Fc716)+2 bekers"],
  ["PC", "O", "520605", "520605", "Medium container"],
  ["PC", "S", "520605", "520605", "Verrekencode Fc566+0,90"],
  ["1320", "H", "570000", "570000", "Serviceheffing < 100.000"],
  ["1321", "H", "570000", "570000", "Serviceheffing > 100.000"],
  ["1322", "H", "570000", "570000", "Serviceheffing > 200.000"],
  ["1323", "H", "570000", "570000", "Serviceheffing > 300.000"],
  ["1324", "H", "570000", "570000", "Serviceheffing > 2.500.000"],
  ["1325", "H", "570000", "570000", "Service"],
  ["1325", "O", "570000", "570000", "Serviceheffing > 10.000.000"],
  ["1325", "S", "570000", "570000", "Serviceheffing > 10.000.000"],
  ["1328", "H", "570000", "570000", "Promotieheffing"],
  ["1329", "H", "570000", "570000", "Promotieheffing"],
  ["1330", "H", "570000", "570000", "Inschrijfgeld Direct"],
  ["1540", "H", "570000", "570000", "Handel Provisie"],
  ["2028", "H", "570000", "570000", "Slotplaten huur"],
  ["2072", "H", "570000", "570000", "Connect Logistiek zonder brief"],
  ["2099", "H", "570000", "570000", "EKT gebruik"],
  ["2099", "S", "570000", "570000", "EKT gebruik"],
  ["2128", "H", "570000", "570000", "D&O koeling A 28"],
  ["2146", "H", "570000", "570000", "Abonnement e-transacties"],
  ["2155", "H", "570000", "570000", "Klokleveringbericht"],
  ["2158", "H", "570000", "570000", "Abonnement LM Online Basic"],
  ["2159", "H", "570000", "570000", "Abonnement LM Online Premium"],
  ["2161", "H", "570000", "570000", "Floriday Exporteursmodule"],
  ["2263", "H", "570000", "570000", "Abonnement telefonie"],
  ["2299", "H", "570000", "570000", "Webshopkoppeling"],
  ["2300", "H", "570000", "570000", "Floriday VMP-koppeling"],
  ["2301", "H", "570000", "570000", "Floriday Koppeling"],
  ["2370", "H", "570000", "570000", "Extra distributienummers N Planten"],
  ["2403", "H", "570000", "570000", "Klokleveringsbericht Naaldwijk"],
  ["2403", "S", "570000", "570000", "Klokleveringsbericht Naaldwijk"],
  ["2404", "H", "570000", "570000", "Klokleveringsbericht Rijnsburg"],
  ["2405", "H", "570000", "570000", "Klokleveringsbericht Aalsmeer"],
  ["2405", "S", "570000", "570000", "Klokleveringsbericht Aalsmeer"],
  ["2792", "H", "570000", "570000", "Select Delivery KVV"],
  ["2793", "H", "570000", "570000", "Select Delivery"],
  ["2806", "H", "570000", "570000", "Inleveren Denen"],
  ["2806", "S", "570000", "570000", "Inleveren Denen"],
  ["2843", "S", "570000", "570000", "Restitutie promotieheffing"],
  ["2913", "H", "570000", "570000", "Minimum distributie A A Planten"],
  ["2913", "S", "570000", "570000", "Minimum distributie A A Planten"],
  ["2914", "H", "570000", "570000", "Minimum distributie N N Planten"],
  ["2914", "S", "570000", "570000", "Minimum distributie N N Planten"],
  ["2915", "H", "570000", "570000", "Minimum distributie R R Bloemen"],
  ["2916", "H", "570000", "570000", "Korting Emballage huur 2024"],
  ["2918", "H", "570000", "570000", "Aanbod markeren via App"],
  ["2955", "H", "570000", "570000", "Abonnement Insights"],
  ["2221", "H", "570010", "570010", "MPN KOA Box"],
  ["2222", "H", "570010", "570010", "MPN KOA Box Extra"],
  ["2625", "H", "570010", "570010", "KOA"],
  ["1317", "S", "570100", "570100", "Select Delivery Priority"],
  ["1320", "S", "570100", "570100", "Serviceheffing < 100.000"],
  ["1321", "S", "570100", "570100", "Serviceheffing > 100.000"],
  ["1322", "S", "570100", "570100", "Serviceheffing > 200.000"],
  ["1323", "S", "570100", "570100", "Serviceheffing > 300.000"],
  ["1324", "S", "570100", "570100", "Serviceheffing > 2.500.000"],
  ["1328", "S", "570100", "570100", "Promotieheffing"],
  ["1329", "S", "570100", "570100", "Promotieheffing"],
  ["2792", "S", "570100", "570100", "Select Delivery KVV"],
  ["1314", "H", "570120", "570150", "Transactieheffing"],
  ["2917", "H", "570120", "570150", "Korting Transactieheffing 2024"],
  ["1314", "S", "570135", "570165", "Transactieheffing"],
  ["2917", "S", "570135", "570165", "Korting Transactieheffing 2024"],
  ["1320", "O", "570200", "570200", "Serviceheffing < 100.000"],
  ["1321", "O", "570200", "570200", "Serviceheffing > 100.000"],
  ["1322", "O", "570200", "570200", "Serviceheffing > 200.000"],
  ["1323", "O", "570200", "570200", "Serviceheffing > 300.000"],
  ["1324", "O", "570200", "570200", "Serviceheffing > 2.500.000"],
  ["1328", "O", "570200", "570200", "Promotieheffing"],
  ["2031", "H", "570500", "570500", "Afkoop slotplaten 63155"],
  ["1030", "S", "890000", "890000", "Product verkopen"],
  ["1530", "S", "890000", "890000", "Product verkopen"],
  ["1030", "H", "890100", "890100", "Product verkopen"],
  ["1530", "H", "890100", "890100", "Product verkopen"],
  // Finance (Marten): FK "Product aankopen" (1015) books per klok location.
  ["1015_Aalsmeer_klok", "S", "500020", "510020", "Product Aankopen  AALSMEER KLOK"],
  ["1015_Aalsmeer_klok", "H", "500025", "510025", "Product Aankopen  AALSMEER KLOK"],
  ["1015_Naaldwijk_klok", "S", "500000", "510000", "Product Aankopen  NAALDWIJK KLOK"],
  ["1015_Naaldwijk_klok", "H", "500010", "500010", "Product Aankopen  NAALDWIJK KLOK"],
  ["1015_Rijnsburg_klok", "S", "500030", "510030", "Product Aankopen  RIJNSBURG KLOK"],
  ["1015_Rijnsburg_klok", "H", "500035", "500035", "Product Aankopen  RIJNSBURG KLOK"],
];

const KING_BTW_CODES = new Set(["H", "S", "O", "E"]);

function text(value) {
  return String(value ?? "").trim();
}

export function kingLedgerKey(prdId, btw) {
  return `${text(prdId).toUpperCase()}|${text(btw).toUpperCase()}`;
}

export function normalizeKingLedgerRow(row) {
  const btw = text(row?.btw).toUpperCase();
  return {
    prd_id: text(row?.prd_id),
    btw: KING_BTW_CODES.has(btw) ? btw : btw.slice(0, 1),
    gb_fl: text(row?.gb_fl),
    gb_pl: text(row?.gb_pl) || text(row?.gb_fl),
    omschrijving: text(row?.omschrijving).slice(0, 40),
  };
}

export function seedKingLedgerMap() {
  return KING_LEDGER_SEED.map(([prd_id, btw, gb_fl, gb_pl, omschrijving]) => normalizeKingLedgerRow({ prd_id, btw, gb_fl, gb_pl, omschrijving }));
}

// An empty/missing map means "never configured" -> seeded; once saved (even
// edited down) the stored map is used as-is.
export function normalizeKingLedgerMap(source) {
  if (!Array.isArray(source)) {
    return seedKingLedgerMap();
  }
  const byKey = new Map();
  for (const row of source.map(normalizeKingLedgerRow)) {
    if (row.prd_id && row.btw && row.gb_fl) {
      byKey.set(kingLedgerKey(row.prd_id, row.btw), row);
    }
  }
  return [...byKey.values()];
}

// Prefilled from finance's sample journal (king_journaal.xml): dagboek
// "Veilin", the creditor account per buying company, BTW accounts 150200
// (laag, S) / 150250 (hoog, H). archiefsoort and the King folder path are
// finance/King-side settings left to fill in.
export const DEFAULT_KING_SETTINGS = {
  dagboek: "Veilin",
  btw_laag_account: "150200",
  btw_hoog_account: "150250",
  creditor_accounts: { "049876": "250051", "057390": "250050", "060708": "250055", "063155": "250052" },
  // flowers ("FL") or plants ("PL") per company -- used for product lines
  // that can't be split by product group.
  // AI2 (Sjaak van der Vijver B.V.'s Ai2 incasso, klant 90985) has its own
  // creditor in King -- listed here so its field shows in the settings.
  company_ledger_column: { "049876": "FL", "057390": "FL", "060708": "FL", "063155": "PL", "AI2-90985": "FL" },
  // "split": product amounts split over flowers (sales account 001) vs
  // plants (002/003) in proportion to the invoice's own product lines;
  // "company": always the company's column above.
  fl_pl_mode: "split",
  archiefsoort: "",
  king_pdf_dir: "",
  next_stuknummer: 0,
  bg_definitief: false,
};

export function normalizeKingSettings(source) {
  const merged = { ...DEFAULT_KING_SETTINGS, ...(source && typeof source === "object" ? source : {}) };
  const cleanMap = (value, fallback) => {
    const out = {};
    for (const [key, entry] of Object.entries(value && typeof value === "object" ? value : fallback)) {
      if (text(key) && text(entry)) {
        out[text(key)] = text(entry);
      }
    }
    return out;
  };
  // Defaults merged per company, so a company added later (AI2) also shows
  // up for settings saved before it existed.
  const columns = cleanMap({ ...DEFAULT_KING_SETTINGS.company_ledger_column, ...(merged.company_ledger_column || {}) }, DEFAULT_KING_SETTINGS.company_ledger_column);
  for (const key of Object.keys(columns)) {
    columns[key] = columns[key].toUpperCase() === "PL" ? "PL" : "FL";
  }
  return {
    dagboek: text(merged.dagboek).slice(0, 10) || DEFAULT_KING_SETTINGS.dagboek,
    btw_laag_account: text(merged.btw_laag_account) || DEFAULT_KING_SETTINGS.btw_laag_account,
    btw_hoog_account: text(merged.btw_hoog_account) || DEFAULT_KING_SETTINGS.btw_hoog_account,
    creditor_accounts: cleanMap(merged.creditor_accounts, DEFAULT_KING_SETTINGS.creditor_accounts),
    company_ledger_column: columns,
    fl_pl_mode: merged.fl_pl_mode === "company" ? "company" : "split",
    archiefsoort: text(merged.archiefsoort).slice(0, 10),
    king_pdf_dir: text(merged.king_pdf_dir),
    next_stuknummer: Math.max(0, Math.floor(Number(merged.next_stuknummer) || 0)),
    bg_definitief: merged.bg_definitief === true,
  };
}

// "049876.FC.2025.0085" -> "049876FC20250085" (the sample's own format).
export function kingExternId(invoiceNumber) {
  return text(invoiceNumber).replace(/[^A-Za-z0-9]/g, "").slice(0, 20);
}

function round2(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

const KING_PRODUCT_PRD_IDS = new Set(["1005", "1015", "1035", "1535", "1030", "1530"]);

function isProductSummaryLine(line) {
  return text(line?.product_type_code) === "57" || KING_PRODUCT_PRD_IDS.has(text(line?.prd_id));
}

// Detail sales account 001 = cut flowers, 002/003 = (house/garden) plants.
function ledgerColumnForAccount(accountId) {
  const id = text(accountId);
  return id === "002" || id === "003" ? "PL" : "FL";
}

// Splits `amount` over FL/PL in proportion to `weights` ({FL, PL}), keeping
// the exact total (rounding remainder on the larger share).
function splitAmount(amount, weights) {
  const fl = Math.abs(Number(weights.FL) || 0);
  const pl = Math.abs(Number(weights.PL) || 0);
  const sum = fl + pl;
  if (!sum || !pl) {
    return [{ column: "FL", amount: round2(amount) }];
  }
  if (!fl) {
    return [{ column: "PL", amount: round2(amount) }];
  }
  const plAmount = round2((amount * pl) / sum);
  const flAmount = round2(amount - plAmount);
  return [{ column: "FL", amount: flAmount }, { column: "PL", amount: plAmount }];
}

// One King journal post for one invoice. `header` = inkoop_invoice_headers
// row, `lines` = that invoice's stored lines (raw), `ledgerMap` = the
// kingkoppel rows, `settings` = normalizeKingSettings output.
// Returns { lines: [...], problems: [...], warnings: [...], totals }.
export function buildKingJournalPost(header, lines, ledgerMap, settings) {
  const problems = [];
  const warnings = [];
  const invoiceNumber = text(header?.invoice_number);
  const company = text(header?.company_number) || invoiceNumber.slice(0, 6);
  const invoiceDate = text(header?.invoice_date).slice(0, 10);
  const grandTotal = round2(header?.grand_total);
  const externId = kingExternId(invoiceNumber);
  const byKey = new Map((ledgerMap || []).map((row) => [kingLedgerKey(row.prd_id, row.btw), row]));
  const creditor = settings.creditor_accounts[company] || "";
  if (!creditor) {
    problems.push(`No creditor account for company ${company} (King settings).`);
  }
  if (header?.grand_total === null || header?.grand_total === undefined) {
    problems.push("Invoice total missing -- re-upload this invoice so its XML header is read.");
  }
  if (Array.isArray(header?.klok_locations) && header.klok_locations.length > 1) {
    warnings.push(`Invoice has products from several klok locations (${header.klok_locations.join(", ")}); booked to ${header.klok_location}.`);
  }

  const journalLines = [];
  const pushLine = (volgnummer, account, side, amount, omschrijving) => {
    journalLines.push({
      volgnummer,
      rekeningnummer: account,
      boekzijde: side,
      valutacode: text(header?.currency) || "EUR",
      bedrag: round2(amount),
      omschrijving: text(omschrijving).slice(0, 40),
      factuurnummer: invoiceNumber,
      factuurdatum: invoiceDate,
      archiefstuk_extern_id: externId,
    });
  };

  pushLine(1, creditor, "CRED", grandTotal, invoiceNumber);
  const vatBy = (category) => round2((header?.vat_subtotals || []).filter((entry) => text(entry.category).toUpperCase() === category).reduce((sum, entry) => sum + (Number(entry.amount) || 0), 0));
  const vatLaag = vatBy("S");
  const vatHoog = vatBy("H");
  if (vatLaag) pushLine(2, settings.btw_laag_account, "DEB", vatLaag, "Verlaagd tarief");
  if (vatHoog) pushLine(3, settings.btw_hoog_account, "DEB", vatHoog, "Standaard tarief");

  const summaryLines = (lines || []).filter((line) => text(line?.trigger_code) === "0");
  if (!summaryLines.length) {
    problems.push("No summary lines found -- re-upload this invoice so its King fields are read.");
  }

  // Product-group weights per BTW category, from the invoice's own detail
  // product lines (trigger 1, type 57).
  const weightsByVat = {};
  for (const line of lines || []) {
    if (text(line?.trigger_code) !== "1" || text(line?.product_type_code) !== "57") continue;
    const vat = text(line?.vat_category).toUpperCase();
    weightsByVat[vat] = weightsByVat[vat] || { FL: 0, PL: 0 };
    weightsByVat[vat][ledgerColumnForAccount(line?.account_id)] += Math.abs(Number(line?.total) || 0);
  }
  const companyColumn = settings.company_ledger_column[company] === "PL" ? "PL" : "FL";

  const merged = new Map();
  for (const line of summaryLines) {
    const vat = text(line?.vat_category).toUpperCase();
    const amount = round2(line?.total);
    if (!amount) continue;
    let key = text(line?.prd_id);
    if (key === "1015" && text(header?.invoice_type) === "klokfactuur" && header?.klok_location) {
      key = `1015_${header.klok_location}_klok`;
    }
    const mapping = byKey.get(kingLedgerKey(key, vat)) || byKey.get(kingLedgerKey(line?.prd_id, vat));
    if (!mapping) {
      problems.push(`No King mapping for Prd ${key || "?"} / BTW ${vat || "?"} (${text(line?.description)}, ${amount.toFixed(2)}).`);
      continue;
    }
    const parts = isProductSummaryLine(line)
      ? (settings.fl_pl_mode === "split" && weightsByVat[vat]
        ? splitAmount(amount, weightsByVat[vat])
        : [{ column: companyColumn, amount }])
      : [{ column: "FL", amount }];
    for (const part of parts) {
      const account = part.column === "PL" ? mapping.gb_pl : mapping.gb_fl;
      const mergeKey = `${account}|${mapping.omschrijving}`;
      const entry = merged.get(mergeKey) || { account, omschrijving: mapping.omschrijving, amount: 0, prd_keys: new Set() };
      entry.amount = round2(entry.amount + part.amount);
      entry.prd_keys.add(`${key}/${vat}`);
      merged.set(mergeKey, entry);
    }
  }
  let volgnummer = 4;
  for (const entry of merged.values()) {
    pushLine(volgnummer, entry.account, "DEB", entry.amount, entry.omschrijving);
    journalLines[journalLines.length - 1].prd_keys = [...entry.prd_keys];
    volgnummer += 1;
  }

  const debit = round2(journalLines.filter((line) => line.boekzijde === "DEB").reduce((sum, line) => sum + line.bedrag, 0));
  const credit = round2(journalLines.filter((line) => line.boekzijde === "CRED").reduce((sum, line) => sum + line.bedrag, 0));
  if (Math.abs(debit - credit) > 0.011 && !problems.length) {
    problems.push(`Journal does not balance: DEB ${debit.toFixed(2)} vs CRED ${credit.toFixed(2)}.`);
  }
  return {
    invoice_number: invoiceNumber,
    invoice_type: text(header?.invoice_type),
    company_number: company,
    invoice_date: invoiceDate,
    extern_id: externId,
    lines: journalLines,
    totals: { debit, credit, grand_total: grandTotal },
    problems,
    warnings,
  };
}

function xmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function amount(value) {
  return round2(value).toFixed(2);
}

// Same element names, order and formats as finance's accepted sample.
export function buildKingJournaalXml(posts, { boekdatum, dagboek, definitief }) {
  const parts = ["<?xml version='1.0' encoding='UTF-8'?>", "<KING_JOURNAAL><BOEKINGSGANGEN><BOEKINGSGANG>"];
  parts.push(`<BG_OMSCHRIJVING>${xmlEscape(`VeilingNr nota loaded on - ${String(boekdatum).replace(/-/g, ".")}`)}</BG_OMSCHRIJVING>`);
  parts.push(`<BG_DEFINITIEF>${definitief ? "true" : "false"}</BG_DEFINITIEF><JOURNAALPOSTEN>`);
  for (const post of posts) {
    parts.push("<JOURNAALPOST>");
    parts.push(`<JP_DAGBOEKCODE>${xmlEscape(dagboek)}</JP_DAGBOEKCODE>`);
    parts.push(`<JP_BOEKDATUM>${xmlEscape(post.invoice_date || boekdatum)}</JP_BOEKDATUM>`);
    if (post.stuknummer) {
      parts.push(`<JP_STUKNUMMER>${post.stuknummer}</JP_STUKNUMMER>`);
    }
    parts.push(`<JP_OMSCHRIJVING>${xmlEscape(post.invoice_number)}</JP_OMSCHRIJVING><JOURNAALREGELS>`);
    for (const line of post.lines) {
      parts.push("<JOURNAALREGEL>");
      parts.push(`<JR_VOLGNUMMER>${line.volgnummer}</JR_VOLGNUMMER>`);
      parts.push(`<JR_REKENINGNUMMER>${xmlEscape(line.rekeningnummer)}</JR_REKENINGNUMMER>`);
      parts.push(`<JR_BOEKZIJDE>${line.boekzijde}</JR_BOEKZIJDE>`);
      parts.push(`<JR_VALUTACODE>${xmlEscape(line.valutacode)}</JR_VALUTACODE>`);
      parts.push(`<JR_VALUTABEDRAG>${amount(line.bedrag)}</JR_VALUTABEDRAG>`);
      parts.push(`<JR_OMSCHRIJVING>${xmlEscape(line.omschrijving)}</JR_OMSCHRIJVING>`);
      parts.push(`<JR_FACTUURNUMMER>${xmlEscape(line.factuurnummer)}</JR_FACTUURNUMMER>`);
      parts.push(`<JR_FACTUURDATUM>${xmlEscape(line.factuurdatum)}</JR_FACTUURDATUM>`);
      parts.push("<JR_VERVALDATUM></JR_VERVALDATUM><JR_BETALINGSKENMERK></JR_BETALINGSKENMERK>");
      parts.push(`<JR_ARCHIEFSTUK_EXTERN_ID>${xmlEscape(line.archiefstuk_extern_id)}</JR_ARCHIEFSTUK_EXTERN_ID>`);
      parts.push("</JOURNAALREGEL>");
    }
    parts.push("</JOURNAALREGELS></JOURNAALPOST>");
  }
  parts.push("</JOURNAALPOSTEN></BOEKINGSGANG></BOEKINGSGANGEN></KING_JOURNAAL>");
  return parts.join("");
}

// King docs ("Eisen aan het XML-bestand met digitale archiefstukken"):
// DAR_ARCHIEFSOORT + DAR_BESTANDSNAAM are mandatory; DAR_EXTERN_ID is what
// a journal line's JR_ARCHIEFSTUK_EXTERN_ID links to.
export function buildKingArchiefXml(posts, { archiefsoort, pdfDir }) {
  const dir = String(pdfDir || "").replace(/[\\/]+$/, "");
  const parts = ["<?xml version='1.0' encoding='UTF-8'?>", "<KING_DIGITAAL_ARCHIEF><DIGITAAL_ARCHIEF>"];
  for (const post of posts) {
    parts.push("<DIGITAAL_ARCHIEFSTUK>");
    parts.push(`<DAR_ARCHIEFSOORT>${xmlEscape(archiefsoort)}</DAR_ARCHIEFSOORT>`);
    parts.push(`<DAR_BESTANDSNAAM>${xmlEscape(`${dir}\\${post.extern_id}.pdf`)}</DAR_BESTANDSNAAM>`);
    parts.push(`<DAR_DATUM>${xmlEscape(post.invoice_date)}</DAR_DATUM>`);
    parts.push(`<DAR_EXTERNE_CODE>${xmlEscape(post.invoice_number)}</DAR_EXTERNE_CODE>`);
    parts.push(`<DAR_OPMERKING>${xmlEscape(`FloraHolland ${post.invoice_number}`)}</DAR_OPMERKING>`);
    parts.push("<DAR_VERWERKSOORT>GEEN</DAR_VERWERKSOORT>");
    parts.push("<DAR_NAW_SOORT>NOG_TOEWIJZEN</DAR_NAW_SOORT>");
    parts.push(`<DAR_EXTERN_ID>${xmlEscape(post.extern_id)}</DAR_EXTERN_ID>`);
    parts.push("</DIGITAAL_ARCHIEFSTUK>");
  }
  parts.push("</DIGITAAL_ARCHIEF></KING_DIGITAAL_ARCHIEF>");
  return parts.join("");
}
