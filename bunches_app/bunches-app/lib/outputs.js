// outputs.js — genereer downloadbare bestanden uit pipeline-resultaat.

// CSV escape: omsluit met quotes als nodig, escape interne quotes.
function csvCell(v) {
  if (v == null) return '';
  const s = String(v);
  if (/[",;\n\r]/.test(s)) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function csvRow(cells, sep = ';') {
  return cells.map(csvCell).join(sep);
}

/**
 * Inlezen CSV: één regel per (variant, broncode_inlezen, total).
 * Formaat volgens doelsysteem:
 *   - Geen header
 *   - 7 kolommen per rij: variant (met trailing space) | totaal_stems | broncode_inlezen | 0 | 0 | 0 | 0
 *   - CRLF regeleindes
 *   - Alfabetisch gesorteerd op variant, dan op broncode_inlezen
 */
function generateInlezenCsv(inlezen) {
  const sorted = [...inlezen].sort((a, b) => {
    const va = String(a.variant || '');
    const vb = String(b.variant || '');
    if (va !== vb) return va.localeCompare(vb);
    return (a.broncode_inlezen || 0) - (b.broncode_inlezen || 0);
  });
  const lines = sorted.map(r =>
    [`${r.variant} `, r.total, r.broncode_inlezen, 0, 0, 0, 0].join(';')
  );
  return lines.join('\r\n') + '\r\n';
}

/**
 * Format ISO date (yyyy-mm-dd) to dd-mm-yyyy for UNI / printing.
 */
function formatDateNL(isoDate) {
  if (!isoDate) return '';
  const m = String(isoDate).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return isoDate;
  return `${m[3]}-${m[2]}-${m[1]}`;
}

/**
 * YYBU* file: Universal/UNI formaat.
 * Header rijen + data rijen, semicolon-separated.
 */
function generateUniFile(sheetName, lines, dateStr) {
  const out = [];
  out.push('UNI_VERSION:3.6.34');
  out.push(`UNI_CUST_ID:${sheetName}`);
  out.push(`UNI_DATE:${formatDateNL(dateStr) || 'DATUM VERTREK'}`);
  out.push('UNI_FTERM:;');
  out.push('UNI_STANDING:J');
  out.push('UNI_HEADER:');
  out.push(['int_item_number', 'group', 'description', 'remark', 'amount'].join(';'));

  for (const l of lines) {
    out.push([
      l.broncode,
      l.broncode_inlezen ?? '',
      (l.naam || '').trim(),
      '',
      l.total_stems,
    ].map(v => v == null ? '' : String(v)).join(';'));
  }
  return out.join('\n') + '\n';
}

/**
 * Printlijst als HTML (voor print/preview). Klein en zelf-contained.
 * Elke rij heeft een checkbox voor het afvinken bij het picken.
 */
function generatePrintlijstHtml(title, items, dateStr, options = {}) {
  const displayDate = formatDateNL(dateStr) || dateStr || '';
  const autoPrint = options?.autoPrint === true;
  const rows = items.map(it => `
    <tr>
      <td class="check"><input type="checkbox"></td>
      <td>${escapeHtml(it.naam)}</td>
      <td class="num">${escapeHtml(it.tak || '')}</td>
      <td class="num">${it.lengte != null ? it.lengte : ''}</td>
      <td class="num">${it.aantal_eenheden != null ? formatNum(it.aantal_eenheden) : '?'}</td>
      <td class="num">${it.ape || '?'}</td>
      <td class="num">${it.totaal_bossen}</td>
    </tr>
  `).join('');

  return `<!doctype html>
<html lang="nl">
<head>
<meta charset="utf-8">
<title>${escapeHtml(title)}</title>
<style>
  body { font-family: Arial, sans-serif; font-size: 11pt; margin: 20px; }
  h1 { font-size: 14pt; margin: 0 0 10px; }
  .meta { color: #666; font-size: 10pt; margin-bottom: 15px; }
  table { width: 100%; border-collapse: collapse; }
  th, td { border: 1px solid #999; padding: 4px 8px; text-align: left; }
  th { background: #eee; }
  .num { text-align: right; }
  .check { text-align: center; width: 30px; }
  .check input { width: 18px; height: 18px; cursor: pointer; }
  @media print {
    body { margin: 10mm; }
    h1 { font-size: 12pt; }
    /* Checkboxes zichtbaar houden bij printen */
    .check input { -webkit-appearance: checkbox; appearance: checkbox; }
  }
</style>
</head>
<body>
<h1>${escapeHtml(title)}</h1>
<div class="meta">${escapeHtml(displayDate || '')} — ${items.length} regels</div>
<table>
  <thead>
    <tr>
      <th class="check">✓</th>
      <th>Naam</th>
      <th class="num">Tak</th>
      <th class="num">Lengte</th>
      <th class="num">Aantal</th>
      <th class="num">APE</th>
      <th class="num">Totaal bossen</th>
    </tr>
  </thead>
  <tbody>${rows}</tbody>
</table>
${autoPrint ? `<script>
window.addEventListener('load', () => {
  setTimeout(() => {
    window.focus();
    window.print();
  }, 150);
}, { once: true });
</script>` : ''}
</body>
</html>`;
}

function escapeHtml(s) {
  if (s == null) return '';
  return String(s).replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

function formatNum(n) {
  if (n == null) return '';
  // Print als integer als het rond is, anders met 1 decimaal
  return Math.abs(n - Math.round(n)) < 0.001 ? String(Math.round(n)) : n.toFixed(1);
}

module.exports = {
  generateInlezenCsv,
  generateUniFile,
  generatePrintlijstHtml,
  generatePrintlijstPdf,
  formatDateNL,
};

/**
 * Printlijst als PDF (A4 portret). Layout:
 *   - Header met titel + datum
 *   - Tabel met kolommen: ✓ | Naam | Tak | Lengte | Aantal | APE | Totaal
 *   - Checkbox-vierkantje aan het begin van elke rij (leeg, om met de hand aan te vinken)
 *   - Auto pagina-break, kolom-headers worden herhaald op elke pagina
 * Returns een Promise<Buffer>.
 */
function generatePrintlijstPdf(title, items, dateStr) {
  const PDFDocument = require('pdfkit');
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ size: 'A4', margin: 36, bufferPages: true });
    const chunks = [];
    doc.on('data', c => chunks.push(c));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const displayDate = formatDateNL(dateStr) || dateStr || '';
    const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    const leftX = doc.page.margins.left;

    // Kolomindeling (in punten, totaal moet ≤ pageWidth = 523 bij margin 36)
    const cols = [
      { label: '✓',     width: 22, align: 'center' },
      { label: 'Naam',  width: 235, align: 'left'  },
      { label: 'Tak',   width: 40, align: 'center'},
      { label: 'Lengte',width: 45, align: 'right' },
      { label: 'Aantal',width: 55, align: 'right' },
      { label: 'APE',   width: 40, align: 'right' },
      { label: 'Totaal',width: 76, align: 'right' },
    ];
    const rowHeight = 20;
    const headerHeight = 22;

    function drawTableHeader() {
      const y = doc.y;
      doc.rect(leftX, y, pageWidth, headerHeight).fillAndStroke('#eeeeee', '#999999');
      doc.fillColor('#000000').font('Helvetica-Bold').fontSize(9);
      let x = leftX;
      for (const c of cols) {
        doc.text(c.label, x + 4, y + 7, {
          width: c.width - 8,
          align: c.align,
          lineBreak: false,
        });
        x += c.width;
      }
      doc.y = y + headerHeight;
    }

    function drawRow(item, rowIndex) {
      // Check pagina-break
      if (doc.y + rowHeight > doc.page.height - doc.page.margins.bottom) {
        doc.addPage();
        drawPageHeader();
        drawTableHeader();
      }
      const y = doc.y;
      // Zebra achtergrond
      if (rowIndex % 2 === 1) {
        doc.rect(leftX, y, pageWidth, rowHeight).fill('#f7f7f7');
      }
      // Rand
      doc.strokeColor('#cccccc').lineWidth(0.5)
        .moveTo(leftX, y + rowHeight).lineTo(leftX + pageWidth, y + rowHeight).stroke();

      // Checkbox vierkantje
      const cbSize = 11;
      const cbX = leftX + (cols[0].width - cbSize) / 2;
      const cbY = y + (rowHeight - cbSize) / 2;
      doc.strokeColor('#333333').lineWidth(1).rect(cbX, cbY, cbSize, cbSize).stroke();

      // Data-cellen
      doc.fillColor('#000000').font('Helvetica').fontSize(9);
      const values = [
        '',  // checkbox al getekend
        item.naam || '',
        item.tak || '',
        item.lengte != null ? String(item.lengte) : '',
        item.aantal_eenheden != null ? formatNum(item.aantal_eenheden) : '?',
        item.ape ? String(item.ape) : '?',
        item.totaal_bossen != null ? String(item.totaal_bossen) : '',
      ];
      let x = leftX;
      for (let i = 0; i < cols.length; i++) {
        if (i === 0) { x += cols[0].width; continue; }
        doc.text(values[i], x + 4, y + 6, {
          width: cols[i].width - 8,
          align: cols[i].align,
          lineBreak: false,
          ellipsis: true,
        });
        x += cols[i].width;
      }
      doc.y = y + rowHeight;
    }

    function drawPageHeader() {
      doc.fillColor('#000000').font('Helvetica-Bold').fontSize(14)
        .text(title, leftX, doc.page.margins.top);
      doc.font('Helvetica').fontSize(9).fillColor('#666666')
        .text(`${displayDate}  —  ${items.length} regels`, leftX);
      doc.moveDown(0.5);
    }

    // Eerste pagina
    drawPageHeader();
    drawTableHeader();
    items.forEach((it, i) => drawRow(it, i));

    doc.end();
  });
}
