// Small dependency-free chart renderers — SVG only, no charting library.
// Categorical order below is the dataviz skill's validated default theme,
// used in fixed order (never cycled/reassigned) so a truck keeps its color
// across re-renders as long as it keeps appearing in the same relative order.
window.WC = (function () {
  const CATEGORICAL = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#4a3aa7', '#e34948'];

  function colorFor(index) {
    return CATEGORICAL[index % CATEGORICAL.length];
  }

  function renderDonut(container, scannedPct) {
    const pct = Math.max(0, Math.min(100, scannedPct));
    const r = 60, c = 2 * Math.PI * r;
    const dash = (pct / 100) * c;
    container.innerHTML = `
      <svg viewBox="0 0 160 160" width="200" height="200">
        <circle cx="80" cy="80" r="${r}" fill="none" stroke="var(--pending-soft)" stroke-width="18"></circle>
        <circle cx="80" cy="80" r="${r}" fill="none" stroke="var(--st-good)" stroke-width="18"
          stroke-dasharray="${dash} ${c - dash}" stroke-linecap="round"
          transform="rotate(-90 80 80)"></circle>
        <text x="80" y="88" text-anchor="middle" font-family="'Big Shoulders',sans-serif"
          font-weight="800" font-size="26" fill="var(--ink)">${pct}%</text>
      </svg>`;
  }

  // rows: [{ x, series, value }]. One bar per (x, series) pair, grouped by x.
  function renderGroupedBarChart(container, rows, { height = 160 } = {}) {
    if (!rows.length) {
      container.innerHTML = '<p class="chart-empty">No data for this selection.</p>';
      return;
    }
    const xValues = Array.from(new Set(rows.map(r => r.x))).sort();
    const seriesValues = Array.from(new Set(rows.map(r => r.series)));
    const maxVal = Math.max(1, ...rows.map(r => r.value));

    const seriesColor = {};
    seriesValues.forEach((s, i) => { seriesColor[s] = colorFor(i); });

    const groupW = 46, barW = Math.max(6, Math.floor(30 / Math.max(seriesValues.length, 1)));
    const chartW = xValues.length * groupW + 20;
    const chartH = height;

    let bars = '';
    xValues.forEach((x, gi) => {
      const gx = 10 + gi * groupW;
      seriesValues.forEach((s, si) => {
        const row = rows.find(r => r.x === x && r.series === s);
        const val = row ? row.value : 0;
        const barH = (val / maxVal) * (chartH - 30);
        const bx = gx + si * barW;
        const by = chartH - 20 - barH;
        bars += `<rect x="${bx}" y="${by}" width="${barW - 1}" height="${barH}" fill="${seriesColor[s]}" rx="1"><title>${escapeXml(x)} · ${escapeXml(s)}: ${val}</title></rect>`;
      });
      bars += `<text x="${gx + (groupW - 10) / 2}" y="${chartH - 6}" text-anchor="middle" font-size="9" fill="var(--text-muted)" font-family="'IBM Plex Mono',monospace">${escapeXml(x)}</text>`;
    });

    const legend = seriesValues.map((s, i) =>
      `<span class="legend-item"><span class="swatch" style="background:${colorFor(i)}"></span>${escapeXml(s)}</span>`
    ).join('');

    container.innerHTML = `
      <div class="chart-legend">${legend}</div>
      <svg viewBox="0 0 ${chartW} ${chartH}" style="width:100%;height:${chartH}px;">${bars}</svg>`;
  }

  // rows: [{ label, value }]. Single-series horizontal bars, longest first.
  function renderSimpleBarChart(container, rows) {
    if (!rows.length) {
      container.innerHTML = '<p class="chart-empty">No data for this selection.</p>';
      return;
    }
    const sorted = rows.slice().sort((a, b) => b.value - a.value);
    const maxVal = Math.max(1, ...sorted.map(r => r.value));

    container.innerHTML = sorted.map(r => `
      <div class="hbar-row">
        <span class="hbar-label">${escapeXml(r.label)}</span>
        <div class="hbar-track"><div class="hbar-fill" style="width:${(r.value / maxVal) * 100}%"></div></div>
        <span class="hbar-value">${r.value}</span>
      </div>
    `).join('');
  }

  function escapeXml(str) {
    return String(str).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  return { renderDonut, renderGroupedBarChart, renderSimpleBarChart, colorFor };
})();
