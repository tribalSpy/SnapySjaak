// Pure data-transform functions — a JS port of tools/scan_dashboard.py's pandas
// logic, kept deliberately literal (including its quirks, e.g. `x || fallback`
// treating 0 as "missing") so this dashboard's numbers match the Streamlit one
// it's meant to replace, not a reinterpretation of it.
window.WD = (function () {

  function parseTimestamp(value) {
    if (!value) return null;
    const d = new Date(value);
    return Number.isNaN(d.getTime()) ? null : d;
  }

  function localDateStr(d) {
    if (!d) return null;
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function todayStr() {
    return localDateStr(new Date());
  }

  // Attaches _localTs (Date) and localDateStr (YYYY-MM-DD, browser-local) to
  // every event once, up front — mirrors the parsed_timestamp/date columns
  // the Python version computes with pandas.
  function prepareEvents(events) {
    return events.map(e => {
      const ts = parseTimestamp(e.timestamp);
      return { ...e, _localTs: ts, localDateStr: ts ? localDateStr(ts) : null };
    });
  }

  function normalizeLocations(value) {
    if (Array.isArray(value)) return value.map(String).filter(v => v.trim());
    if (value === null || value === undefined || value === '') return [''];
    return [String(value)];
  }

  function statusFromCounts(scanned, expected) {
    if (scanned <= 0) return 'pending';
    if (expected <= 0) return 'extra';
    if (scanned < expected) return 'partial';
    if (scanned === expected) return 'scanned';
    return 'extra';
  }

  function rowsFromReferenceMap(byRef) {
    const rows = [];
    for (const item of Object.values(byRef)) {
      const locations = (item.locations && item.locations.length) ? item.locations : [''];
      for (const location of locations) {
        rows.push({
          location,
          reference: item.reference || '',
          truck: item.truck || '',
          group: item.group || '',
          status: item.status || '',
          trolleyCount: Number(item.trolleyCount) || 0,
          scannedCount: Number(item.scannedCount) || 0,
          lastScannedBy: item.lastScannedBy || '',
          lastScannedAt: item.lastScannedAt || ''
        });
      }
    }
    return rows;
  }

  // Reconstructs "what was expected/scanned as of this date" — for today, seeds
  // from the live backend state; for any other date, rebuilds purely from the
  // activity log (expected_reference / split_update / scan_complete events on
  // that date), same as the Python version.
  function buildStateForDate(liveStatus, events, selectedDateStr) {
    const useLiveFallback = selectedDateStr === todayStr();
    let byRef = {};

    if (liveStatus && useLiveFallback) {
      for (const [location, row] of Object.entries(liveStatus)) {
        const ref = String(row.reference || '').trim().toUpperCase();
        if (!ref) continue;
        byRef[ref] = {
          reference: ref, truck: '', group: row.group || '', status: 'pending',
          trolleyCount: Number(row.trolleyCount) || 0, scannedCount: 0,
          lastScannedBy: '', lastScannedAt: '',
          locations: location ? [location] : []
        };
      }
    }

    const byLocalTs = (a, b) => (a._localTs?.getTime() || 0) - (b._localTs?.getTime() || 0);

    const expectedEvents = events
      .filter(e => e.type === 'expected_reference' && e.localDateStr === selectedDateStr)
      .sort(byLocalTs);
    const hasHistoricalExpectedList = expectedEvents.length > 0;

    if (hasHistoricalExpectedList) {
      byRef = {};
      for (const e of expectedEvents) {
        const ref = String(e.reference || '').trim().toUpperCase();
        if (!ref) continue;
        byRef[ref] = {
          reference: ref, truck: '', group: e.group || '', status: 'pending',
          trolleyCount: Number(e.trolleyCount) || 0, scannedCount: 0,
          lastScannedBy: '', lastScannedAt: '',
          locations: normalizeLocations(e.locations)
        };
      }
    }

    const splitUpdates = events
      .filter(e => e.type === 'split_update' && e.localDateStr === selectedDateStr)
      .sort(byLocalTs);
    for (const e of splitUpdates) {
      const ref = String(e.reference || '').trim().toUpperCase();
      if (!ref) continue;
      const eventLocations = normalizeLocations(e.locations).filter(Boolean);
      if (!byRef[ref]) {
        byRef[ref] = {
          reference: ref, truck: '', group: '', status: 'pending',
          trolleyCount: Number(e.trolleyCount) || 0, scannedCount: 0,
          lastScannedBy: '', lastScannedAt: '', locations: eventLocations
        };
      }
      const item = byRef[ref];
      item.truck = e.truck || item.truck;
      item.group = e.group || item.group;
      if (e.trolleyCount) item.trolleyCount = Number(e.trolleyCount) || item.trolleyCount;
      if (eventLocations.length) item.locations = [eventLocations[eventLocations.length - 1]];
    }

    const scans = events
      .filter(e => e.type === 'scan_complete' && e.localDateStr === selectedDateStr)
      .sort(byLocalTs);
    for (const e of scans) {
      const ref = String(e.reference || '').trim().toUpperCase();
      if (!ref) continue;
      const eventLocations = normalizeLocations(e.locations).filter(Boolean);
      if (!byRef[ref]) {
        byRef[ref] = {
          reference: ref, truck: '', group: '', status: 'pending',
          trolleyCount: Number(e.trolleyCount) || 0, scannedCount: 0,
          lastScannedBy: '', lastScannedAt: '', locations: eventLocations
        };
      }
      const item = byRef[ref];
      item.scannedCount = Number(e.scannedCount) || item.scannedCount;
      item.trolleyCount = Number(e.trolleyCount) || item.trolleyCount;
      item.status = e.status || statusFromCounts(item.scannedCount, item.trolleyCount);
      item.lastScannedBy = e.scannerId || item.lastScannedBy;
      item.lastScannedAt = e.timestamp || item.lastScannedAt;
      item.group = e.group || item.group;
      if (eventLocations.length) item.locations = [eventLocations[eventLocations.length - 1]];
    }

    for (const item of Object.values(byRef)) {
      item.status = statusFromCounts(Number(item.scannedCount) || 0, Number(item.trolleyCount) || 0);
      if (!useLiveFallback && !hasHistoricalExpectedList && (Number(item.scannedCount) || 0) === 0) {
        item.status = 'scheduled';
      }
      item.locations = (item.locations && item.locations.length) ? item.locations.slice(-1) : [];
    }

    return rowsFromReferenceMap(byRef);
  }

  function dedupeByReference(rows) {
    const seen = new Set();
    const out = [];
    for (const r of rows) {
      const key = String(r.reference || '').toUpperCase();
      if (seen.has(key)) continue;
      seen.add(key);
      out.push(r);
    }
    return out;
  }

  function statusSummary(rows) {
    if (!rows.length) {
      return { expected_refs: 0, expected_trolleys: 0, scanned_refs: 0, scanned_trolleys: 0, effective_scanned_trolleys: 0, pending_refs: 0, pending_trolleys: 0, extra_refs: 0 };
    }
    const perRef = dedupeByReference(rows);
    const scannedRefs = perRef.filter(r => r.scannedCount >= r.trolleyCount);
    const pendingRefs = perRef.filter(r => r.scannedCount === 0);
    const pendingTrolleys = perRef.reduce((acc, r) => acc + Math.max(r.trolleyCount - r.scannedCount, 0), 0);
    const extraRefs = perRef.filter(r => String(r.status || '').toLowerCase() === 'extra');

    return {
      expected_refs: perRef.length,
      expected_trolleys: perRef.reduce((a, r) => a + r.trolleyCount, 0),
      scanned_refs: scannedRefs.length,
      scanned_trolleys: perRef.reduce((a, r) => a + r.scannedCount, 0),
      // Per-reference capped at its own trolleyCount, so extra scans on one
      // reference can never offset another reference that's still pending —
      // used for the completion percentage, not the raw "Scanned count" stat.
      effective_scanned_trolleys: perRef.reduce((a, r) => a + Math.min(r.scannedCount, r.trolleyCount), 0),
      pending_refs: pendingRefs.length,
      pending_trolleys: pendingTrolleys,
      extra_refs: extraRefs.length
    };
  }

  function filterState(rows, truck, group, status) {
    return rows.filter(r =>
      (truck === 'All' || (r.truck || '') === truck) &&
      (group === 'All' || (r.group || '') === group) &&
      (status === 'All' || (r.status || '') === status)
    );
  }

  function applyReferenceExclusions(rows, excludedSet) {
    if (!rows.length || !excludedSet || excludedSet.size === 0) return rows;
    return rows.filter(r => !excludedSet.has(String(r.reference || '').toUpperCase()));
  }

  function filterEventsToVisibleReferences(events, visibleRefs) {
    if (!events.length) return events;
    return events.filter(e => visibleRefs.has(String(e.reference || '').toUpperCase()));
  }

  function classifyCanceledScan(row) {
    const filtered = Number(row.filteredTagCount) || 0;
    const expected = Number(row.expectedCount) || 0;
    if (filtered <= 0) return 'no_tags';
    if (expected <= 0) return 'unknown_expected';
    if (filtered > expected) return 'too_many';
    if (filtered < expected) return 'missing';
    return 'count_matches';
  }

  function enrichScansWithState(events, stateRows) {
    if (!events.length || !stateRows.length) return events;
    const perRef = {};
    for (const r of stateRows) {
      const key = String(r.reference || '').toUpperCase();
      if (!key || perRef[key]) continue;
      perRef[key] = r;
    }
    return events.map(e => {
      const stateRow = perRef[String(e.reference || '').toUpperCase()];
      const truck = e.truck || (stateRow ? stateRow.truck : '') || '';
      const group = e.group || (stateRow ? stateRow.group : '') || '';
      let locations = e.locations;
      const hasLocations = Array.isArray(locations) ? locations.length > 0 : !!locations;
      if (!hasLocations) locations = stateRow ? stateRow.location : '';
      let expectedCount = e.expectedCount;
      if (!expectedCount) expectedCount = (stateRow ? stateRow.trolleyCount : 0) || 0;
      return { ...e, truck, group, locations, expectedCount };
    });
  }

  function formatHms(d) {
    return d.toLocaleTimeString([], { hour12: false });
  }

  function truckScanDurations(events) {
    const withTs = events.filter(e => e._localTs);
    const byTruck = {};
    for (const e of withTs) {
      const truck = e.truck || 'No truck';
      if (!byTruck[truck]) byTruck[truck] = [];
      byTruck[truck].push(e._localTs);
    }
    return Object.entries(byTruck).map(([truck, times]) => {
      const sorted = times.slice().sort((a, b) => a - b);
      const first = sorted[0], last = sorted[sorted.length - 1];
      return {
        truck,
        first_scan: formatHms(first),
        last_scan: formatHms(last),
        scan_events: times.length,
        load_minutes: Math.round(((last - first) / 60000) * 10) / 10
      };
    });
  }

  function scanTimeframe(events) {
    const counts = {};
    for (const e of events) {
      if (!e._localTs) continue;
      const hour = String(e._localTs.getHours()).padStart(2, '0') + ':00';
      const truck = e.truck || 'No truck';
      const key = hour + '|' + truck;
      counts[key] = (counts[key] || 0) + 1;
    }
    return Object.entries(counts).map(([key, scans]) => {
      const [hour, truck] = key.split('|');
      return { hour, truck, scans };
    });
  }

  function collapseMissingReferenceEvents(events) {
    const missing = events.filter(e => e.type === 'scan_missing_reference');
    const others = events.filter(e => e.type !== 'scan_missing_reference');
    if (!missing.length) return events;

    const groups = {};
    for (const e of missing) {
      const key = [e.localDateStr || '', e.reference || '', e.scannerId || ''].join('|');
      (groups[key] = groups[key] || []).push(e);
    }
    const collapsed = Object.values(groups).map(list => {
      const sorted = list.slice().sort((a, b) => (a._localTs?.getTime() || 0) - (b._localTs?.getTime() || 0));
      return { ...sorted[0], repeatCount: list.length };
    });
    return others.concat(collapsed);
  }

  function uniqueSorted(values) {
    return Array.from(new Set(values.filter(v => v))).sort();
  }

  return {
    parseTimestamp, localDateStr, todayStr, prepareEvents, normalizeLocations,
    statusFromCounts, rowsFromReferenceMap, buildStateForDate, dedupeByReference,
    statusSummary, filterState, applyReferenceExclusions, filterEventsToVisibleReferences,
    classifyCanceledScan, enrichScansWithState, truckScanDurations, scanTimeframe,
    collapseMissingReferenceEvents, uniqueSorted, formatHms
  };
})();
