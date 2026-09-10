(function () {
  const D = window.WD, C = window.WC;

  const state = {
    liveStatus: {},
    events: [],
    excludedReferences: new Set(),
    filters: { date: D.todayStr(), truck: 'All', group: 'All', status: 'All' }
  };

  const el = {};
  document.querySelectorAll('[id]').forEach(node => { el[toCamel(node.id)] = node; });

  function toCamel(id) {
    return id.replace(/-([a-z0-9])/g, (_, c) => c.toUpperCase());
  }

  // ---------------------------------------------------------------
  // Fetching
  // ---------------------------------------------------------------
  async function loadConfig() {
    try {
      const res = await fetch('api/config');
      const data = await res.json();
      el.backendHost.textContent = data.backendHost || 'unknown';
    } catch { el.backendHost.textContent = 'unknown'; }
  }

  async function loadSettings() {
    try {
      const res = await fetch('api/settings');
      const data = await res.json();
      state.excludedReferences = new Set(data.excludedReferences || []);
    } catch { state.excludedReferences = new Set(); }
  }

  async function saveSettings(list) {
    const res = await fetch('api/settings', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ excludedReferences: list })
    });
    const data = await res.json();
    state.excludedReferences = new Set(data.excludedReferences || []);
  }

  async function loadStatus() {
    try {
      const res = await fetch('api/status');
      const payload = await res.json();
      if (!payload.ok) throw new Error(payload.error || 'backend error');
      state.liveStatus = payload.data || {};
      setConnected(true);
    } catch (err) {
      setConnected(false, err.message);
    }
  }

  async function loadActivity() {
    try {
      const res = await fetch('api/activity-log?limit=20000');
      const payload = await res.json();
      state.events = D.prepareEvents(Array.isArray(payload.events) ? payload.events : []);
    } catch {
      // keep whatever we had; a transient failure shouldn't blank the page
    }
  }

  function setConnected(ok, detail) {
    el.connDot.classList.toggle('ok', ok);
    el.connText.textContent = ok
      ? 'connected · ' + new Date().toLocaleTimeString()
      : 'disconnected' + (detail ? ' — ' + detail : '');
  }

  // ---------------------------------------------------------------
  // Generic table helper
  // ---------------------------------------------------------------
  function renderTable(tableEl, columns, rows) {
    tableEl.querySelector('thead').innerHTML =
      '<tr>' + columns.map(c => `<th>${escapeHtml(c.label)}</th>`).join('') + '</tr>';
    tableEl.querySelector('tbody').innerHTML = rows.map(row =>
      '<tr>' + columns.map(c => `<td>${c.render ? c.render(row) : escapeHtml(fmtCell(row[c.key]))}</td>`).join('') + '</tr>'
    ).join('');
  }

  function fmtCell(v) {
    if (v === null || v === undefined || v === '') return '—';
    if (Array.isArray(v)) return v.length ? v.join(', ') : '—';
    return String(v);
  }

  function statusPill(status) {
    const known = ['scanned', 'partial', 'pending', 'extra', 'scheduled'];
    const key = known.includes(status) ? status : 'pending';
    const label = { scanned: 'Scanned', partial: 'Partial', pending: 'Pending', extra: 'Extra', scheduled: 'Scheduled' }[key];
    return `<span class="pill ${key}">${label}</span>`;
  }

  function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function localTime(iso) {
    const d = D.parseTimestamp(iso);
    return d ? d.toLocaleString() : '—';
  }

  // ---------------------------------------------------------------
  // Filter option population
  // ---------------------------------------------------------------
  function fillSelect(selectEl, values, keepAll) {
    const current = selectEl.value;
    selectEl.innerHTML = '<option value="All">All</option>' +
      values.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
    if (values.includes(current) || (keepAll && current === 'All')) selectEl.value = current;
  }

  function populateFilterOptions(stateRows) {
    fillSelect(el.filterTruck, D.uniqueSorted(stateRows.map(r => r.truck)), true);
    fillSelect(el.filterGroup, D.uniqueSorted(stateRows.map(r => r.group)), true);
    fillSelect(el.filterStatus, D.uniqueSorted(stateRows.map(r => r.status)), true);
  }

  // ---------------------------------------------------------------
  // Stats + donut
  // ---------------------------------------------------------------
  function renderStats(summary) {
    const items = [
      ['Expected refs', summary.expected_refs],
      ['Expected trolleys', summary.expected_trolleys],
      ['Scanned refs', summary.scanned_refs],
      ['Scanned count', summary.scanned_trolleys],
      ['Pending refs', summary.pending_refs],
      ['Pending trolleys', summary.pending_trolleys],
      ['Extra refs', summary.extra_refs]
    ];
    el.stats.innerHTML = items.map(([label, val]) =>
      `<div class="stat"><span class="num">${val}</span><span class="label">${label}</span></div>`
    ).join('');
  }

  function renderDonutFor(summary) {
    const expected = Math.max(summary.expected_trolleys, 0);
    const scanned = Math.min(Math.max(summary.effective_scanned_trolleys, 0), expected);
    const pct = expected ? Math.round((scanned / expected) * 1000) / 10 : 0;
    C.renderDonut(el.donut, pct);
  }

  // ---------------------------------------------------------------
  // Issue banner (Google Drive / upload problems)
  // ---------------------------------------------------------------
  function renderIssueBanner() {
    const cutoff = Date.now() - 30 * 60 * 1000;
    const issues = state.events.filter(e => e.type === 'upload_issue' && e._localTs && e._localTs.getTime() >= cutoff);
    if (!issues.length) { el.issueBanner.hidden = true; return; }

    const latest = issues[issues.length - 1];
    const resolved = latest.status === 'resolved';
    el.issueBanner.hidden = false;
    el.issueBanner.classList.toggle('resolved', resolved);
    const sourceLabel = (latest.source || 'upload').replace(/_/g, ' ');
    el.issueBanner.textContent = resolved
      ? `✅ ${sourceLabel} reconnected on ${latest.scannerId || 'unknown scanner'} — upload succeeded.`
      : `🚨 ${sourceLabel} issue on ${latest.scannerId || 'unknown scanner'} (${latest.status || 'retrying'}) — ${issues.length} in the last 30 min.`;
  }

  // ---------------------------------------------------------------
  // Tabs: Expected Work
  // ---------------------------------------------------------------
  function renderExpectedTab(filteredState, selectedDate) {
    el.expectedCaption.textContent =
      `Expected Work on ${selectedDate}. For older dates this is rebuilt from the activity log and current expected list.`;
    el.expectedEmpty.hidden = filteredState.length > 0;
    renderTable(el.tblExpected, [
      { key: 'location', label: 'Location' },
      { key: 'reference', label: 'Reference' },
      { key: 'truck', label: 'Truck' },
      { key: 'group', label: 'Group' },
      { key: 'status', label: 'Status', render: r => statusPill(r.status) },
      { key: 'scannedCount', label: 'Scanned', render: r => `${r.scannedCount} / ${r.trolleyCount}` },
      { key: 'lastScannedBy', label: 'Last scanned by' },
      { key: 'lastScannedAt', label: 'Last scanned at', render: r => escapeHtml(r.lastScannedAt ? localTime(r.lastScannedAt) : '—') }
    ], filteredState.slice().sort((a, b) => a.location.localeCompare(b.location)));
  }

  // ---------------------------------------------------------------
  // Tabs: Truck Schedule
  // ---------------------------------------------------------------
  function renderTrucksTab(filteredState, selectedActivity, visibleRefs, selectedDate) {
    let scansForGraph = selectedActivity.filter(e => e.type === 'scan_complete');
    scansForGraph = D.filterEventsToVisibleReferences(scansForGraph, visibleRefs);
    scansForGraph = D.enrichScansWithState(scansForGraph, filteredState);
    if (state.filters.truck !== 'All') scansForGraph = scansForGraph.filter(e => e.truck === state.filters.truck);
    if (state.filters.group !== 'All') scansForGraph = scansForGraph.filter(e => e.group === state.filters.group);

    // Per-truck summary
    const byTruck = {};
    for (const r of D.dedupeByReference(filteredState)) {
      const key = (r.truck || '') + '|' + (r.group || '');
      if (!byTruck[key]) byTruck[key] = { truck: r.truck || '', group: r.group || '', references: 0, expected_trolleys: 0, scanned_count: 0 };
      byTruck[key].references++;
      byTruck[key].expected_trolleys += r.trolleyCount;
      byTruck[key].scanned_count += r.scannedCount;
    }
    const durations = D.truckScanDurations(scansForGraph);
    const durationByTruck = Object.fromEntries(durations.map(d => [d.truck, d]));
    const perTruck = Object.values(byTruck).map(t => ({ ...t, ...(durationByTruck[t.truck] || {}) }));

    renderTable(el.tblTruckSummary, [
      { key: 'truck', label: 'Truck', render: r => escapeHtml(r.truck || 'No truck') },
      { key: 'group', label: 'Group' },
      { key: 'references', label: 'References' },
      { key: 'expected_trolleys', label: 'Expected trolleys' },
      { key: 'scanned_count', label: 'Scanned count' },
      { key: 'first_scan', label: 'First scan' },
      { key: 'last_scan', label: 'Last scan' },
      { key: 'load_minutes', label: 'Load minutes' }
    ], perTruck.sort((a, b) => (a.truck || '').localeCompare(b.truck || '')));

    // Truck detail dropdown
    const truckChoices = D.uniqueSorted(filteredState.map(r => r.truck));
    const prev = el.truckDetailSelect.value;
    el.truckDetailSelect.innerHTML = truckChoices.length
      ? truckChoices.map(t => `<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`).join('')
      : '<option value="">—</option>';
    if (truckChoices.includes(prev)) el.truckDetailSelect.value = prev;
    renderTruckDetail(el.truckDetailSelect.value, filteredState, durationByTruck);

    // All truck lines
    renderTable(el.tblTruckLines, [
      { key: 'truck', label: 'Truck' }, { key: 'reference', label: 'Reference' },
      { key: 'location', label: 'Location' }, { key: 'group', label: 'Group' },
      { key: 'trolleyCount', label: 'Trolleys' }, { key: 'scannedCount', label: 'Scanned' },
      { key: 'status', label: 'Status', render: r => statusPill(r.status) }
    ], filteredState.slice().sort((a, b) => (a.truck || '').localeCompare(b.truck || '') || a.reference.localeCompare(b.reference)));

    // Loading timeframe chart
    const timeframeRows = D.scanTimeframe(scansForGraph).map(r => ({ x: r.hour, series: r.truck, value: r.scans }));
    C.renderGroupedBarChart(el.chartTruckTimeframe, timeframeRows);
  }

  function renderTruckDetail(truck, filteredState, durationByTruck) {
    if (!truck) { el.truckDetail.innerHTML = ''; return; }
    const detailRows = filteredState.filter(r => r.truck === truck)
      .sort((a, b) => a.status.localeCompare(b.status) || a.reference.localeCompare(b.reference));
    const summary = D.statusSummary(detailRows);
    const dur = durationByTruck[truck];

    let html = `<div class="detail-stats">
      <div class="stat"><span class="num">${summary.expected_refs}</span><span class="label">References</span></div>
      <div class="stat"><span class="num">${summary.expected_trolleys}</span><span class="label">Expected trolleys</span></div>
      <div class="stat"><span class="num">${summary.scanned_trolleys}</span><span class="label">Scanned count</span></div>
      <div class="stat"><span class="num">${summary.pending_refs}</span><span class="label">Pending refs</span></div>
    </div>`;
    if (dur) {
      html += `<div class="detail-stats">
        <div class="stat"><span class="num">${dur.load_minutes}</span><span class="label">Load minutes</span></div>
        <div class="stat"><span class="num">${dur.first_scan}</span><span class="label">First scan</span></div>
        <div class="stat"><span class="num">${dur.last_scan}</span><span class="label">Last scan</span></div>
      </div>`;
    }

    const pending = detailRows.filter(r => r.status === 'pending');
    const partial = detailRows.filter(r => r.status === 'partial');
    if (pending.length) html += `<h3 class="section-title">Pending in ${escapeHtml(truck)}</h3>` + tableHtml(pending);
    if (partial.length) html += `<h3 class="section-title">Partially scanned in ${escapeHtml(truck)}</h3>` + tableHtml(partial);
    html += `<h3 class="section-title">All references in ${escapeHtml(truck)}</h3>` + tableHtml(detailRows);

    el.truckDetail.innerHTML = html;
  }

  function tableHtml(rows) {
    const cols = [
      ['reference', 'Reference'], ['location', 'Location'], ['trolleyCount', 'Trolleys'],
      ['scannedCount', 'Scanned'], ['group', 'Group']
    ];
    return `<div class="table-wrap"><table><thead><tr>${cols.map(c => `<th>${c[1]}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(r => `<tr>${cols.map(c => `<td>${escapeHtml(fmtCell(r[c[0]]))}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
  }

  // ---------------------------------------------------------------
  // Tabs: Scans
  // ---------------------------------------------------------------
  function renderScansTab(selectedActivity, filteredState, visibleRefs, selectedDate) {
    el.scansTitle.textContent = `Scans on ${selectedDate}`;

    let scans = selectedActivity.filter(e => e.type === 'scan_complete');
    scans = D.filterEventsToVisibleReferences(scans, visibleRefs);
    scans = D.enrichScansWithState(scans, filteredState);
    if (state.filters.truck !== 'All') scans = scans.filter(e => e.truck === state.filters.truck);
    if (state.filters.group !== 'All') scans = scans.filter(e => e.group === state.filters.group);

    el.scansEmpty.hidden = scans.length > 0 || selectedActivity.length > 0;
    if (!selectedActivity.length) {
      el.scansEmpty.textContent = 'No activity log exists yet, or no scan events were recorded for this date.';
    }

    renderTable(el.tblScans, [
      { key: 'timestamp', label: 'Time', render: r => escapeHtml(localTime(r.timestamp)) },
      { key: 'reference', label: 'Reference' }, { key: 'truck', label: 'Truck' }, { key: 'group', label: 'Group' },
      { key: 'status', label: 'Status', render: r => statusPill(r.status) },
      { key: 'scannedCount', label: 'Scanned', render: r => `${r.scannedCount ?? '?'} / ${r.trolleyCount ?? '?'}` },
      { key: 'uniqueTagCount', label: 'Tags' }, { key: 'hasExtension', label: 'Extension' },
      { key: 'scannerId', label: 'Scanner' }, { key: 'locations', label: 'Locations' }
    ], scans.slice().sort((a, b) => (b._localTs?.getTime() || 0) - (a._localTs?.getTime() || 0)));

    if (scans.length) {
      const byScanner = {};
      for (const s of scans) byScanner[s.scannerId || 'unknown'] = (byScanner[s.scannerId || 'unknown'] || 0) + 1;
      C.renderSimpleBarChart(el.chartScansByScanner, Object.entries(byScanner).map(([label, value]) => ({ label, value })));

      const timeRows = D.scanTimeframe(scans).map(r => ({ x: r.hour, series: r.truck, value: r.scans }));
      C.renderGroupedBarChart(el.chartScansByTime, timeRows);
    } else {
      el.chartScansByScanner.innerHTML = '<p class="chart-empty">No scans for this selection.</p>';
      el.chartScansByTime.innerHTML = '';
    }

    const missing = D.collapseMissingReferenceEvents(
      selectedActivity.filter(e => e.type === 'scan_missing_reference')
    );
    el.missingFaultsBlock.hidden = missing.length === 0;
    if (missing.length) {
      renderTable(el.tblMissing, [
        { key: 'timestamp', label: 'Time', render: r => escapeHtml(localTime(r.timestamp)) },
        { key: 'reference', label: 'Reference' }, { key: 'scannerId', label: 'Scanner' },
        { key: 'uniqueTagCount', label: 'Tags' }, { key: 'repeatCount', label: 'Repeat count' }
      ], missing.sort((a, b) => (b._localTs?.getTime() || 0) - (a._localTs?.getTime() || 0)));
    }
  }

  // ---------------------------------------------------------------
  // Tabs: Canceled Scans
  // ---------------------------------------------------------------
  function renderCanceledTab(selectedActivity, filteredState, selectedDate) {
    el.canceledTitle.textContent = `Canceled Scans on ${selectedDate}`;

    let canceled = selectedActivity.filter(e => e.type === 'scan_canceled');
    canceled = D.enrichScansWithState(canceled, filteredState);
    if (state.filters.truck !== 'All') canceled = canceled.filter(e => e.truck === state.filters.truck);
    if (state.filters.group !== 'All') canceled = canceled.filter(e => e.group === state.filters.group);

    canceled = canceled.map(e => {
      const filteredTagCount = Number(e.filteredTagCount) || 0;
      const expectedCount = Number(e.expectedCount) || 0;
      return { ...e, filteredTagCount, expectedCount, countDelta: filteredTagCount - expectedCount, countHint: D.classifyCanceledScan({ filteredTagCount, expectedCount }) };
    });

    el.canceledEmpty.hidden = canceled.length > 0;
    el.canceledEmpty.textContent = selectedActivity.length ? 'No canceled scans were recorded for this date.' : 'No activity log exists yet for this date.';

    const hintCounts = { too_many: 0, missing: 0, no_tags: 0, count_matches: 0, unknown_expected: 0 };
    for (const c of canceled) hintCounts[c.countHint] = (hintCounts[c.countHint] || 0) + 1;

    el.canceledStats.innerHTML = [
      ['Canceled scans', canceled.length], ['Too many', hintCounts.too_many],
      ['Missing', hintCounts.missing], ['No tags', hintCounts.no_tags]
    ].map(([label, val]) => `<div class="stat"><span class="num">${val}</span><span class="label">${label}</span></div>`).join('');

    C.renderSimpleBarChart(el.chartCanceledHints,
      Object.entries(hintCounts).filter(([, v]) => v > 0).map(([label, value]) => ({ label, value })));

    renderTable(el.tblCanceled, [
      { key: 'timestamp', label: 'Time', render: r => escapeHtml(localTime(r.timestamp)) },
      { key: 'reference', label: 'Reference' }, { key: 'truck', label: 'Truck' }, { key: 'group', label: 'Group' },
      { key: 'location', label: 'Location' }, { key: 'filteredTagCount', label: 'Filtered tags' },
      { key: 'expectedCount', label: 'Expected' }, { key: 'countDelta', label: 'Delta' },
      { key: 'countHint', label: 'Hint' }, { key: 'scannerId', label: 'Scanner' }, { key: 'sessionId', label: 'Session' }
    ], canceled.sort((a, b) => (b._localTs?.getTime() || 0) - (a._localTs?.getTime() || 0)));
  }

  // ---------------------------------------------------------------
  // Tabs: Timeline
  // ---------------------------------------------------------------
  function renderTimelineTab(selectedActivity, visibleRefs, selectedDate) {
    el.timelineTitle.textContent = `Activity Timeline on ${selectedDate}`;
    let timeline = D.collapseMissingReferenceEvents(selectedActivity);
    timeline = timeline.filter(e => !e.reference || visibleRefs.has(String(e.reference).toUpperCase()));

    el.timelineEmpty.hidden = timeline.length > 0;
    el.timelineEmpty.textContent = 'No activity has been recorded for this date yet. The dashboard still shows the current backend state above.';

    renderTable(el.tblTimeline, [
      { key: 'timestamp', label: 'Time', render: r => escapeHtml(localTime(r.timestamp)) },
      { key: 'type', label: 'Type' }, { key: 'reference', label: 'Reference' },
      { key: 'scannerId', label: 'Scanner' }, { key: 'truck', label: 'Truck' }, { key: 'group', label: 'Group' },
      { key: 'summary', label: 'Details', render: r => escapeHtml(describeEvent(r)) }
    ], timeline.sort((a, b) => (b._localTs?.getTime() || 0) - (a._localTs?.getTime() || 0)));
  }

  function describeEvent(e) {
    switch (e.type) {
      case 'scan_complete': return `${e.scannedCount ?? '?'}/${e.trolleyCount ?? '?'} scanned, status ${e.status || '?'}`;
      case 'scan_canceled': return `cancel reason: ${e.cancelReason || 'unknown'}`;
      case 'scan_missing_reference': return `unknown reference${e.repeatCount ? ` (x${e.repeatCount})` : ''}`;
      case 'upload_issue': return `${e.source || 'upload'} — ${e.status || ''} — ${e.message || ''}`;
      case 'sheet_sync': case 'upload_locations': return `${e.referenceCount ?? '?'} references, ${e.locationCount ?? '?'} locations`;
      case 'reset': return `reason: ${e.reason || 'manual'}`;
      case 'split_update': return `→ truck ${e.truck || '?'}, group ${e.group || '?'}`;
      case 'expected_reference': return `expected ${e.trolleyCount ?? '?'} trolleys`;
      default: return '';
    }
  }

  // ---------------------------------------------------------------
  // Tabs: Settings
  // ---------------------------------------------------------------
  function renderSettingsTab(availableStateRows) {
    const byRef = {};
    for (const r of availableStateRows) {
      const key = String(r.reference || '').toUpperCase();
      if (!key) continue;
      if (!byRef[key]) byRef[key] = { reference: key, locations: new Set(), truck: r.truck, group: r.group, trolleyCount: r.trolleyCount, status: r.status, scannedCount: r.scannedCount };
      byRef[key].locations.add(r.location);
      byRef[key].trolleyCount = Math.max(byRef[key].trolleyCount, r.trolleyCount);
      byRef[key].scannedCount = Math.max(byRef[key].scannedCount, r.scannedCount);
    }
    const refRows = Object.values(byRef).map(r => ({ ...r, locations: Array.from(r.locations).filter(Boolean).sort().join(', ') }));

    const availableRefs = D.uniqueSorted([...refRows.map(r => r.reference), ...state.excludedReferences]);
    el.settingsHidden.innerHTML = availableRefs.map(ref =>
      `<option value="${escapeHtml(ref)}" ${state.excludedReferences.has(ref) ? 'selected' : ''}>${escapeHtml(ref)}</option>`
    ).join('');

    el.settingsSummary.textContent =
      `Visible references: ${availableRefs.length - state.excludedReferences.size} | Hidden references: ${state.excludedReferences.size}`;

    const withHidden = refRows.map(r => ({ ...r, hidden: state.excludedReferences.has(r.reference) }))
      .sort((a, b) => (b.hidden - a.hidden) || a.reference.localeCompare(b.reference));

    renderTable(el.tblSettingsRefs, [
      { key: 'reference', label: 'Reference' }, { key: 'locations', label: 'Locations' },
      { key: 'truck', label: 'Truck' }, { key: 'group', label: 'Group' },
      { key: 'trolleyCount', label: 'Trolleys' }, { key: 'status', label: 'Status', render: r => statusPill(r.status) },
      { key: 'scannedCount', label: 'Scanned' },
      {
        key: 'hidden', label: 'Hidden',
        render: r => `<input type="checkbox" class="hide-toggle" data-ref="${escapeHtml(r.reference)}" ${r.hidden ? 'checked' : ''}>`
      }
    ], withHidden);
  }

  el.btnSaveHidden.addEventListener('click', async () => {
    const selected = Array.from(el.settingsHidden.selectedOptions).map(o => o.value);
    await saveSettings(selected);
    recomputeAndRender();
  });
  el.btnClearHidden.addEventListener('click', async () => {
    await saveSettings([]);
    recomputeAndRender();
  });

  // Delegated listener (the table body is fully re-rendered on every update,
  // so a per-checkbox listener would be destroyed each time) — this is the
  // direct way to add or remove ONE reference without fighting the native
  // multi-select's Ctrl/Cmd+click requirement above.
  el.tblSettingsRefs.addEventListener('change', async (e) => {
    if (!e.target.matches('.hide-toggle')) return;
    const ref = e.target.dataset.ref;
    const next = new Set(state.excludedReferences);
    if (e.target.checked) next.add(ref); else next.delete(ref);
    await saveSettings(Array.from(next));
    recomputeAndRender();
  });

  // ---------------------------------------------------------------
  // Master render
  // ---------------------------------------------------------------
  function recomputeAndRender() {
    const selectedDate = el.filterDate.value || D.todayStr();
    const availableStateRows = D.buildStateForDate(state.liveStatus, state.events, selectedDate);
    const stateRows = D.applyReferenceExclusions(availableStateRows, state.excludedReferences);
    const visibleRefs = new Set(stateRows.map(r => String(r.reference || '').toUpperCase()));

    populateFilterOptions(stateRows);

    const filteredState = D.filterState(stateRows, state.filters.truck, state.filters.group, state.filters.status);
    const summary = D.statusSummary(filteredState);
    renderStats(summary);
    renderDonutFor(summary);
    renderIssueBanner();

    const selectedActivity = state.events.filter(e => e.localDateStr === selectedDate);

    renderExpectedTab(filteredState, selectedDate);
    renderTrucksTab(filteredState, selectedActivity, visibleRefs, selectedDate);
    renderScansTab(selectedActivity, filteredState, visibleRefs, selectedDate);
    renderCanceledTab(selectedActivity, filteredState, selectedDate);
    renderTimelineTab(selectedActivity, visibleRefs, selectedDate);
    renderSettingsTab(availableStateRows);
  }

  // ---------------------------------------------------------------
  // Wiring
  // ---------------------------------------------------------------
  el.filterDate.value = state.filters.date;
  el.filterDate.addEventListener('change', () => { state.filters.date = el.filterDate.value; recomputeAndRender(); });
  el.filterTruck.addEventListener('change', () => { state.filters.truck = el.filterTruck.value; recomputeAndRender(); });
  el.filterGroup.addEventListener('change', () => { state.filters.group = el.filterGroup.value; recomputeAndRender(); });
  el.filterStatus.addEventListener('change', () => { state.filters.status = el.filterStatus.value; recomputeAndRender(); });
  el.truckDetailSelect.addEventListener('change', () => recomputeAndRender());

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
    });
  });

  // ---------------------------------------------------------------
  // Boot
  // ---------------------------------------------------------------
  async function boot() {
    await Promise.all([loadConfig(), loadSettings(), loadStatus(), loadActivity()]);
    recomputeAndRender();
  }
  boot();

  setInterval(async () => { await loadStatus(); recomputeAndRender(); }, 5000);
  setInterval(async () => { await loadActivity(); recomputeAndRender(); }, 20000);
})();
