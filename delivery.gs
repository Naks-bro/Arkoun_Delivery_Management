/*************************************************
 * delivery.gs — Arkoun Ops Console (CLEAN)
 * DELIVERY ROUTES (ORS ROAD-OPTIMIZED — ONE-WAY)
 *
 * ✅ IMPORTANT
 * - ONE-WAY ONLY: HUB → last stop (NO return to hub)
 * - Priority: routing + output tables + ROUTE BILLING TABLE
 * - Sidebar/Map REMOVED (as requested)
 *
 * ✅ NOTES
 * - Uses ORS via Cloudflare worker proxy: ROUTES_CFG.ORS_PROXY_URL
 * - If selected fileId can't open: fallback to latest ops_ (same folder → operational folder → anywhere)
 *
 * ⚠️ Requires Advanced Drive service enabled if you use:
 *   - getLatestOpsSpreadsheet_()
 *   - convertExcelToGoogleSheetSameFolder_()
 *************************************************/

/* =========================
   CONFIG  ✅ CHANGE HERE
   ========================= */

const ROUTES_CFG = {
  OPERATIONAL_FOLDER_ID: '1XO9LJ3DEZW3LqDPdpLa3dXx6WTyCYe4B',

  OPS_TEMPLATE_FILE_NAME: 'ops_TEMPLATE',
  PRODUCT_WEIGHTS_FILE_NAME: 'Product weights',
  PINCODES_FILE_NAME: 'Pincodes',

  ORDERS_SHEET_NAME: 'Orders List',
  OUTPUT_SHEET_NAME: 'Delivery Routes',

  // only constraint
  MAX_ROUTE_KG: 20,

  HUB_ADDRESS: 'Arkoun Farms, Jai Mahakali Society, Ambedkar Nagar, 8 Road, Chakala Industrial Area (MIDC), Andheri East, Mumbai, Maharashtra 400093',
  HUB_PINCODE: '400093',

  OPS_NAME_PREFIX: 'ops_',

  ORS_PROXY_URL: 'https://twilight-field-7f06.kokatenakul11.workers.dev/',
  ORS_TIMEOUT_MS: 30000,

  ROUTE_OPT_MODE: 'BEST',     // FAST / BEST
  TWO_OPT_MAX_PASSES: 8,
  TWO_OPT_MAX_SWAPS: 4000,

  MAX_ROUTE_MIN: 90,          // one-way duration cap per route
  MAX_DETOUR_RATIO: 1.45,
  MAX_PAIR_MIN: 35
};

/* Woo export columns (0-based index) */
const ORDERS_COL = {
  ORDER_REF: 1,
  PRODUCT: 2,
  QTY: 3,
  PRICE: 4,
  ADDRESS: 6
};

const SKIP_KEYWORDS = [/order value/i, /total value/i, /^total$/i];

/* =========================
   UI THEME (CLEAN)
   ========================= */
const UI = {
  TITLE_BG: '#e7f0fd',        // soft blue
  SECTION_BG: '#f1f5f9',      // light slate
  HEADER_BG: '#f8fafc',       // header grey
  ROUTE_BG: '#fff7ed',        // pale amber
  BORDER: '#cbd5e1',

  WARN_BG: '#ffecec',         // soft red
  WARN_HDR_BG: '#fecaca',     // stronger red header

  MISSING_BG: '#fff7ed',      // pale amber
  MISSING_HDR_BG: '#ffedd5'
};

/* =========================
   MENU (container-bound only)
   ========================= */

function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Arkoun Ops Console')
.addItem('Build Delivery Routes (ORS)', 'buildRoutesEntry')
      .addSeparator()
      .addItem('Create NEW ops_ from Template', 'createTodayOpsFromTemplate')
      .addToUi();
  } catch (err) {
    // ignore in non-UI contexts
  }
}

function safeAlert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/**
 * Creates new ops_ by copying ops_TEMPLATE (container-bound)
 */
function createTodayOpsFromTemplate() {
  const C = ROUTES_CFG;
  const folder = DriveApp.getFolderById(C.OPERATIONAL_FOLDER_ID);

  const template = getSingleSheetFile_(folder, C.OPS_TEMPLATE_FILE_NAME);
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
  const newName = C.OPS_NAME_PREFIX + ts;

  const copied = DriveApp.getFileById(template.getId()).makeCopy(newName, folder);
  safeAlert_('Created: ' + newName + '\nOpen it from the folder — menu will appear on open.');
  return copied.getId();
}

/* =========================
   MAIN ENTRY (ONE-WAY)
   ========================= */

function buildUserFriendlyRoutes(fileId) {
  Logger.log('Building routes (SMART VRP ONE-WAY via ORS Matrix)...');

  const C = ROUTES_CFG;

  // 1) Try selected fileId
  // 2) Fallback: same folder latest ops_ -> operational latest ops_ -> any latest ops_
  let ssOps;
  let usedMode = 'selected';
  let fallbackNote = '';

  try {
    if (!fileId) throw new Error('No fileId provided');
    ssOps = SpreadsheetApp.openById(fileId);
  } catch (e) {
    usedMode = 'fallback';
    const msg = (e && e.message) ? e.message : String(e);
    Logger.log('⚠️ Cannot open selected fileId: ' + fileId + ' | ' + msg);
    fallbackNote = 'Selected file not accessible; used latest ops_.';

    let sameFolderOps = null;
    try { sameFolderOps = getLatestOpsInSameFolder_(fileId, C.OPS_NAME_PREFIX); } catch (err) {}

    if (sameFolderOps) ssOps = sameFolderOps;
else ssOps = getLatestOpsSpreadsheetDriveApi_(C.OPS_NAME_PREFIX, C.OPERATIONAL_FOLDER_ID);
  }

  const ordersSheet = getOrdersSheetSmart_(ssOps, C.ORDERS_SHEET_NAME, new Set([]), C.OUTPUT_SHEET_NAME);
  if (!ordersSheet) {
    const names = ssOps.getSheets().map(s => s.getName()).join(', ');
    throw new Error('Orders sheet not found. Available sheets: ' + names);
  }

  const rows = ordersSheet.getDataRange().getValues();
  if (rows.length < 2) throw new Error('Orders sheet "' + ordersSheet.getName() + '" has no data rows.');

  // Reference files
  const folder = DriveApp.getFolderById(C.OPERATIONAL_FOLDER_ID);
  const fWeights = getSingleSheetFile_(folder, C.PRODUCT_WEIGHTS_FILE_NAME);
  const fPins = getSingleSheetFile_(folder, C.PINCODES_FILE_NAME);

  const ssWeights = SpreadsheetApp.openById(fWeights.getId());
  const weightIndex = buildFuzzyWeightIndex_(ssWeights);

  const ssPins = SpreadsheetApp.openById(fPins.getId());
  const geoIndex = buildPincodeGeoIndex_(ssPins);
  const pinToGeo = geoIndex.pinToGeo;
  const pinToCity = geoIndex.pinToCity;

  // Hub
  const hubGeo = pinToGeo[C.HUB_PINCODE];
  if (!hubGeo) throw new Error('Hub pincode ' + C.HUB_PINCODE + ' missing Lat/Lng in Pincodes sheet.');
  const hubLoc = [hubGeo.lng, hubGeo.lat]; // [lon,lat]

  // Parse orders + weights
  const parsed = parseOrders_(rows, weightIndex);
  const byOrder = parsed.byOrder;
  const missingWeights = parsed.missingWeights;

  // Validate geo + capacity
  const unassigned = [];
  const orders = [];

  Object.values(byOrder).forEach(o => {
    if (!o.pincode) {
      unassigned.push([o.orderId, o.name, o.address, '—', '', formatWeightHuman_(o.weightKg),
        'No pincode in address', 'Ensure 6-digit pincode in address']);
      return;
    }

    const g = pinToGeo[o.pincode];
    if (!g) {
      unassigned.push([o.orderId, o.name, o.address, o.pincode, (pinToCity[o.pincode] || ''),
        formatWeightHuman_(o.weightKg), 'No Lat/Lng for pincode', 'Add Latitude/Longitude in Pincodes sheet']);
      return;
    }

    if ((o.weightKg || 0) > C.MAX_ROUTE_KG) {
      unassigned.push([o.orderId, o.name, o.address, o.pincode, (pinToCity[o.pincode] || ''),
        formatWeightHuman_(o.weightKg), 'Order exceeds max kg', 'Split order or increase cap']);
      return;
    }

    o.city = pinToCity[o.pincode] || o.city || '';
    o._loc = [g.lng, g.lat]; // [lon,lat]
    orders.push(o);
  });

const ORDER_CHUNK_SIZE = 45; // SAFE LIMIT
const chunks = chunkArray_(orders, ORDER_CHUNK_SIZE);

let routes = [];

chunks.forEach((chunk, idx) => {
  const baseOffset = idx * ORDER_CHUNK_SIZE;
  const matrix = buildGlobalMatrix_(hubLoc, chunk);
  const chunkRoutes = buildRoutesBySavings_ONEWAY_(chunk, matrix, C);

  chunkRoutes.forEach(r => {
    r._matrix = matrix;
    r._baseOffset = baseOffset;
  });

  routes = routes.concat(chunkRoutes);
});




  // Optimize order + compute metrics
routes.forEach(r => {
  const opt = optimizeRouteWithinGlobalMatrix_ONEWAY_(r.stopIdx, r._matrix, C);

  r.stopIdx = opt.stopIdx;

  r.uniquePins = new Set(
    r.stopIdx.map(i => orders[i + r._baseOffset].pincode)
  ).size;

  r.estDurationMin = opt.totalDurS / 60;
  r.estDistanceKm = opt.totalDistM / 1000;

  r.totalKg = r3_(
    r.stopIdx.reduce(
      (a, i) => a + (orders[i + r._baseOffset].weightKg || 0),
      0
    )
  );

  r.totalItems = r.stopIdx.reduce(
    (a, i) => a + ((orders[i + r._baseOffset].items || []).length),
    0
  );
});

  // Assign IDs
  routes.forEach((r, idx) => {
    r.routeId = "R" + String(idx + 1).padStart(3, '0');
    r.zone = 'Auto (ORS)';
r.stops = r.stopIdx.map(i => orders[i + r._baseOffset]);
  });

  // Write output
  const out = ensureDeliveryRoutesAsThird_(ssOps, C.OUTPUT_SHEET_NAME);
  out.clear();

  writeRoutesOutput_ONEWAY_(out, routes, unassigned, pinToCity, missingWeights, ssOps);
  applyReadableLayout_(out);
  applyBandingToRouteTables_(out);

  Logger.log('DONE: SMART ONE-WAY routes written | MODE=' + C.ROUTE_OPT_MODE);

  const usedId = ssOps.getId();
  let usedName = '';
  try { usedName = ssOps.getName ? ssOps.getName() : ''; } catch (e) {}
  const usedMsg = usedName ? (usedName + ' (' + usedId + ')') : usedId;

  return (usedMode === 'selected')
    ? ('✅ Routes built using selected file: ' + usedMsg)
    : ('✅ Routes built using fallback file: ' + usedMsg + ' | ' + fallbackNote);
}

/* =========================
   FALLBACK: latest ops_ in same folder
   ========================= */

function getLatestOpsInSameFolder_(fileId, prefix) {
  if (!fileId) return null;
  prefix = (prefix || 'ops_').toLowerCase();

  const f = DriveApp.getFileById(fileId);
  const parents = f.getParents();
  if (!parents.hasNext()) return null;

  const folder = parents.next();
  const it = folder.getFiles();
  let best = null;

  while (it.hasNext()) {
    const x = it.next();
    if (x.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    const name = (x.getName() || '').toLowerCase();
    if (!name.startsWith(prefix)) continue;

    if (!best) best = x;
    else if (x.getLastUpdated() > best.getLastUpdated()) best = x;
  }

  return best ? SpreadsheetApp.openById(best.getId()) : null;
}

/* =========================
   GLOBAL MATRIX (hub + orders)
   ========================= */

function buildGlobalMatrix_(hubLoc, orders) {
  if (!orders.length) {
    return { locs: [hubLoc], dur: [[0]], dist: [[0]], hub: 0 };
  }

  const locations = [hubLoc].concat(orders.map(o => o._loc));
  const all = locations.map((_, i) => i);

  const m = callOrsMatrix_(locations, all, all);

  if (!m || !m.durations || !m.distances) {
    throw new Error('ORS matrix missing durations/distances');
  }

  return {
    locs: locations,
    dur: m.durations,
    dist: m.distances,
    hub: 0
  };
}


/* =========================
   CLARKE–WRIGHT SAVINGS (ONE-WAY)
   ========================= */

function buildRoutesBySavings_ONEWAY_(orders, matrix, C) {
  const n = orders.length;
  if (n === 0) return [];

  let routes = [];
  const routeOf = Array(n).fill(null);

  // solo routes hub->i
  for (let i = 0; i < n; i++) {
    const solo = soloOneWayCost_(i, matrix);
    const r = {
      stopIdx: [i],
      kg: orders[i].weightKg || 0,
      soloMin: solo.min,
      soloDistKm: solo.km
    };
    routes.push(r);
    routeOf[i] = r;
  }

  // savings pairs
  const pairs = [];
  for (let i = 0; i < n; i++) {
    for (let j = i + 1; j < n; j++) {
      const hij = roadMin_(i, j, matrix);
      if (hij > C.MAX_PAIR_MIN) continue;

      const si = roadMinHubTo_(i, matrix);
      const sj = roadMinHubTo_(j, matrix);

      const saving = (si + sj) - hij;
      pairs.push({ i, j, saving, hij });
    }
  }

  pairs.sort((a, b) => b.saving - a.saving);

  pairs.forEach(p => {
    const ra = routeOf[p.i];
    const rb = routeOf[p.j];
    if (!ra || !rb || ra === rb) return;

    const newKg = (ra.kg || 0) + (rb.kg || 0);
    if (newKg > C.MAX_ROUTE_KG) return;

    const mergedStops = ra.stopIdx.concat(rb.stopIdx);

    const evalRoute = evaluateStopSetCost_ONEWAY_(mergedStops, matrix, C);
    const mergedMin = evalRoute.min;

    if (mergedMin > C.MAX_ROUTE_MIN) return;

    const soloSum = (ra.soloMin || 0) + (rb.soloMin || 0);
    if (soloSum > 0 && mergedMin > soloSum * C.MAX_DETOUR_RATIO) return;

    const rNew = {
      stopIdx: mergedStops,
      kg: newKg,
      soloMin: soloSum,
      soloDistKm: (ra.soloDistKm || 0) + (rb.soloDistKm || 0)
    };

    routes = routes.filter(r => r !== ra && r !== rb);
    routes.push(rNew);
    rNew.stopIdx.forEach(idx => routeOf[idx] = rNew);
  });

  return routes.map(r => ({
    routeId: '',
    zone: '',
    stopIdx: r.stopIdx.slice(),
    uniquePins: 0,
    estDistanceKm: 0,
    estDurationMin: 0,
    totalKg: r3_(r.kg || 0),
    totalItems: 0,
    stops: []
  }));
}

function evaluateStopSetCost_ONEWAY_(stopIdx, matrix, C) {
  const opt = optimizeRouteWithinGlobalMatrix_ONEWAY_(stopIdx, matrix, C);
  return { min: opt.totalDurS / 60, km: opt.totalDistM / 1000 };
}

/* =========================
   ROUTE ORDER OPT (ONE-WAY)
   ========================= */

function optimizeRouteWithinGlobalMatrix_ONEWAY_(stopIdx, matrix, C) {
  if (!stopIdx || stopIdx.length === 0) return { stopIdx: [], totalDurS: 0, totalDistM: 0 };

  const visitNodes = nearestNeighborNodes_(stopIdx, matrix);

  let improved = visitNodes;
  if ((C.ROUTE_OPT_MODE || '').toUpperCase() === 'BEST' && improved.length >= 4) {
    improved = twoOptNodes_ONEWAY_(improved, matrix, C.TWO_OPT_MAX_PASSES, C.TWO_OPT_MAX_SWAPS);
  }

  const totals = computeNodeTourTotals_ONEWAY_(improved, matrix);

  return {
    stopIdx: improved.map(node => node - 1),
    totalDurS: totals.totalDurS,
    totalDistM: totals.totalDistM
  };
}

function nearestNeighborNodes_(stopIdx, matrix) {
  const nodes = stopIdx.map(i => i + 1);
  const remaining = new Set(nodes);
  let current = 0;
  const order = [];

  while (remaining.size) {
    let bestNext = null;
    let best = 9e15;
    remaining.forEach(n => {
      const d = safeCell_(matrix.dur, current, n);
      if (d < best) { best = d; bestNext = n; }
    });
    if (bestNext == null) bestNext = remaining.values().next().value;
    order.push(bestNext);
    remaining.delete(bestNext);
    current = bestNext;
  }
  return order;
}

/**
 * 2-Opt for ONE-WAY tours (no last->hub)
 */
function twoOptNodes_ONEWAY_(nodes, matrix, maxPasses, maxSwaps) {
  let best = nodes.slice();
  const passes = Math.max(1, Number(maxPasses) || 6);
  const swapLimit = Math.max(200, Number(maxSwaps) || 3000);
  let swaps = 0;

  for (let pass = 0; pass < passes && swaps < swapLimit; pass++) {
    let improved = false;

    for (let i = 0; i < best.length - 2 && swaps < swapLimit; i++) {
      for (let k = i + 1; k < best.length - 1 && swaps < swapLimit; k++) {
        const prevNode = (i === 0) ? 0 : best[i - 1];
        const A = best[i];
        const B = best[k];
        const nextNode = (k + 1 < best.length) ? best[k + 1] : null;

        const curr = safeCell_(matrix.dur, prevNode, A) + (nextNode ? safeCell_(matrix.dur, B, nextNode) : 0);
        const alt  = safeCell_(matrix.dur, prevNode, B) + (nextNode ? safeCell_(matrix.dur, A, nextNode) : 0);

        if (alt + 0.0001 < curr) {
          const rev = best.slice(i, k + 1).reverse();
          best.splice(i, k - i + 1, ...rev);
          improved = true;
          swaps++;
        }
      }
    }
    if (!improved) break;
  }
  return best;
}

/**
 * ONE-WAY totals: hub -> first + between stops only
 */
function computeNodeTourTotals_ONEWAY_(nodes, matrix) {
  let totalDurS = 0;
  let totalDistM = 0;

  let prev = 0;
  nodes.forEach(n => {
    totalDurS += safeCell_(matrix.dur, prev, n);
    totalDistM += safeCell_(matrix.dist, prev, n);
    prev = n;
  });

  return { totalDurS, totalDistM };
}

/* =========================
   COST HELPERS (ONE-WAY)
   ========================= */

function roadMinHubTo_(orderIdx, matrix) {
  return safeCell_(matrix.dur, 0, orderIdx + 1) / 60;
}

function roadMin_(i, j, matrix) {
  return safeCell_(matrix.dur, i + 1, j + 1) / 60;
}

function soloOneWayCost_(orderIdx, matrix) {
  const a = safeCell_(matrix.dur, 0, orderIdx + 1);
  const d1 = safeCell_(matrix.dist, 0, orderIdx + 1);
  return { min: a / 60, km: d1 / 1000 };
}

function safeCell_(grid, r, c) {
  try {
    const v = Number(grid[r][c]);
    return isFinite(v) ? v : 0;
  } catch (e) {
    return 0;
  }
}

/* =========================
   ORS MATRIX
   ========================= */

function orsUrl_(endpoint) {
  const base = (ROUTES_CFG.ORS_PROXY_URL || '').replace(/\/+$/, '/');
  return base + endpoint;
}

function callOrsMatrix_(locations, sources, destinations) {
  const payload = {
    locations: locations,
    sources: sources,
    destinations: destinations,
    metrics: ['duration', 'distance']
  };

  const url = orsUrl_('matrix');

  const resp = UrlFetchApp.fetch(url, {
  method: 'post',
  contentType: 'application/json',
  payload: JSON.stringify(payload),
  muteHttpExceptions: true,
  followRedirects: true
});


  const code = resp.getResponseCode();
  const text = resp.getContentText() || '';

  if (code < 200 || code >= 300) {
    throw new Error('ORS matrix failed (' + code + '): ' + text.slice(0, 400));
  }

  return JSON.parse(text);
}

/* =========================
   OUTPUT WRITER (ONE-WAY)
   + ROUTE BILLING TABLE
   + KM WITHOUT DECIMALS ✅
   + CLEAN COLORS ✅
   ========================= */

function writeRoutesOutput_ONEWAY_(out, routes, unassigned, pinToCity, missingWeights, ssOps) {
  let row = 1;

  mergeTitle_(out, row, 1, 12, 'DELIVERY ROUTES (ORS ROAD-OPTIMIZED — ONE-WAY)');
  out.getRange(row, 1, 1, 12).setBackground(UI.TITLE_BG);
  row++;

  out.getRange(row, 1, 1, 12).setValues([[
    'Dispatch Hub:', ROUTES_CFG.HUB_ADDRESS, '', '', '', '', '', '', '',
    'Built at:', new Date(), ''
  ]]);
  out.getRange(row, 1).setFontWeight('bold');
  out.getRange(row, 11).setNumberFormat('yyyy-mm-dd hh:mm');
  row += 2;

  // ===== Summary table (top)
  out.getRange(row, 1, 1, 7).setValues([[
    'Route ID', 'Zone', 'Total Weight (kg)', 'Stops',
    'Unique Pincodes', 'Est. Distance (km)', 'Est. Duration (min)'
  ]]).setFontWeight('bold').setBackground(UI.SECTION_BG);
  row++;

  const sumStart = row;
  routes.forEach(r => {
    out.getRange(row, 1, 1, 7).setValues([[
      r.routeId,
      r.zone,
      r3_(r.totalKg),
      (r.stops || []).length,
      r.uniquePins,
      Math.round(r.estDistanceKm),          // ✅ KM NO DECIMALS
      Math.round(r.estDurationMin)
    ]]);
    row++;
  });

  if (row === sumStart) {
    out.getRange(row, 1, 1, 7).setValues([['—', '—', 0, 0, 0, 0, 0]]);
    row++;
  }

  out.getRange(sumStart - 1, 1, (row - sumStart + 1), 7)
    .setBorder(true, true, true, true, true, true, UI.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  out.getRange(sumStart, 3, (row - sumStart), 1).setNumberFormat('#,##0.###');
  out.getRange(sumStart, 6, (row - sumStart), 1).setNumberFormat('0');      // ✅ KM NO DECIMALS
  out.getRange(sumStart, 7, (row - sumStart), 1).setNumberFormat('0');

  row += 2;

  // ===== Route blocks
  routes.forEach((r, idx) => {
    const hdr = out.getRange(row, 1, 1, 12).merge();
    hdr.setValue(
      'Route: ' + r.routeId +
      ' | Zone: ' + r.zone +
      ' | Total: ' + r3_(r.totalKg) + ' kg' +
      ' | Stops: ' + (r.stops || []).length +
      ' | Est (ONE-WAY): ' + Math.round(r.estDistanceKm) + ' km / ' + Math.round(r.estDurationMin) + ' min'
    );
    hdr.setFontWeight('bold').setBackground(UI.ROUTE_BG);
    row++;

    out.getRange(row, 1, 1, 9).setValues([[
      'Seq', 'Customer', 'Phone', 'Address', 'Pincode', 'City', 'Weight', 'Items', 'Notes'
    ]]).setFontWeight('bold').setBackground(UI.HEADER_BG);
    row++;

    (r.stops || []).forEach((o, sIdx) => {
      const phone = extractPhone_(o.address);
      const city = (pinToCity && o.pincode && pinToCity[o.pincode]) ? pinToCity[o.pincode] : (o.city || '');

      out.getRange(row, 1, 1, 9).setValues([[
        sIdx + 1,
        o.name || '',
        phone || '',
        o.address || '',
        o.pincode || '—',
        city,
        formatWeightHuman_(o.weightKg),
        (o.items || []).join(', '),
        sIdx === 0 ? 'Start →' : ''
      ]]);
      row++;
    });

    out.getRange(row - (r.stops || []).length - 1, 1, (r.stops || []).length + 2, 9)
      .setBorder(true, true, true, true, true, true, UI.BORDER, SpreadsheetApp.BorderStyle.SOLID);

    row++;
  });

  // ===== ROUTE BILLING TABLE ✅ (after routes, before unassigned)
  const bill = writeRouteBillingTable_(out, row, routes, pinToCity, ssOps);
  row = bill.nextRow;

  // ===== Unassigned
  out.getRange(row, 1, 1, 8).setValues([['UNASSIGNED ORDERS', '', '', '', '', '', '', '']])
    .merge().setFontWeight('bold').setBackground(UI.WARN_BG);
  row++;

  out.getRange(row, 1, 1, 8).setValues([[
    'Order ID', 'Customer', 'Address', 'Pincode', 'City', 'Weight', 'Reason', 'Hint'
  ]]).setFontWeight('bold').setBackground(UI.WARN_HDR_BG);
  row++;

  const uaStart = row;

  if (!unassigned.length) {
    out.getRange(row, 1, 1, 8).setValues([['—', '—', '—', '—', '—', '0 g', 'All assigned', '—']]);
    row++;
  } else {
    unassigned.forEach(u => { out.getRange(row, 1, 1, 8).setValues([u]); row++; });
  }

  out.getRange(uaStart - 1, 1, (row - uaStart + 1), 8)
    .setBorder(true, true, true, true, true, true, UI.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  row += 2;

  // ===== Missing weights
  out.getRange(row, 1, 1, 5).setValues([['MISSING PRODUCT WEIGHTS (with suggestions)', '', '', '', '']])
    .merge().setFontWeight('bold').setBackground(UI.MISSING_BG);
  row++;

  out.getRange(row, 1, 1, 4).setValues([['S no', 'Product (normalized)', 'Closest match', 'Note']])
    .setFontWeight('bold').setBackground(UI.MISSING_HDR_BG);
  row++;

  const mwStart = row;
  const uniqMissing = dedupeMissing_(missingWeights);

  if (!uniqMissing.length) {
    out.getRange(row, 1, 1, 4).setValues([[1, '—', '—', 'None missing']]);
    row++;
  } else {
    uniqMissing.forEach((m, i) => {
      const note = m.suggestion
        ? 'Maybe: "' + m.suggestion.name + '" (' + m.suggestion.grams + ' g)'
        : 'Add to Product weights sheet';

      out.getRange(row, 1, 1, 4).setValues([[i + 1, m.normalized, m.suggestion ? m.suggestion.name : '—', note]]);
      row++;
    });
  }

  out.getRange(mwStart - 1, 1, (row - mwStart + 1), 4)
    .setBorder(true, true, true, true, true, true, UI.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

/* =========================
   ROUTE BILLING TABLE ✅
   ========================= */
function writeRouteBillingTable_(out, startRow, routes, pinToCity, ssOps) {
  let row = startRow;

  // Title
  out.getRange(row, 1, 1, 6)
    .setValues([['ROUTE BILLING TABLE', '', '', '', '', '']])
    .merge()
    .setFontWeight('bold')
    .setBackground(UI.SECTION_BG);
  row++;

  // Header
  out.getRange(row, 1, 1, 6)
    .setValues([[
      'Sr no',
      'Name / Customer',
      'Area name / Pincode',
      'Delivery Cost (₹)',
      'Order Value (₹)',
      'Balance (₹)'
    ]])
    .setFontWeight('bold')
    .setBackground(UI.HEADER_BG);
  row++;

  // Rows
  let sr = 1;
  const rbStart = row;

  routes.forEach(r => {
    (r.stops || []).forEach(o => {
      const pin = o.pincode || '';
      const area =
        ((pinToCity && pin && pinToCity[pin]) ? (pinToCity[pin] + ' ') : '') +
        (pin ? ('(' + pin + ')') : '');

      out.getRange(row, 1, 1, 6).setValues([[
        sr,
        o.name || '',
        area,
        '',                      // manual entry
        Number(o.orderValue || 0),
        ''                       // formula below
      ]]);

      out.getRange(row, 6).setFormula('=IFERROR(E' + row + '-D' + row + ',0)');

      sr++;
      row++;
    });
  });

  if (row === rbStart) {
    out.getRange(row, 1, 1, 6).setValues([[1, '—', '—', '', 0, '']]);
    out.getRange(row, 6).setFormula('=IFERROR(E' + row + '-D' + row + ',0)');
    row++;
  }

  const rbEnd = row - 1;

  // Borders + formats
  out.getRange(rbStart - 1, 1, (row - rbStart + 1), 6)
    .setBorder(true, true, true, true, true, true, UI.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  out.getRange(rbStart, 4, (row - rbStart), 3).setNumberFormat('₹#,##0.00');

  // Totals row
  const totalsRow = row;

  out.getRange(totalsRow, 3)
    .setValue('TOTALS (₹)')
    .setFontWeight('bold')
    .setBackground(UI.TITLE_BG);

  const dCell = out.getRange(totalsRow, 4);
  const eCell = out.getRange(totalsRow, 5);
  const fCell = out.getRange(totalsRow, 6);

  dCell.setFormula('=IFERROR(SUM(D' + rbStart + ':D' + rbEnd + '),0)')
    .setNumberFormat('₹#,##0.00')
    .setFontWeight('bold')
    .setBackground(UI.TITLE_BG);

  eCell.setFormula('=IFERROR(SUM(E' + rbStart + ':E' + rbEnd + '),0)')
    .setNumberFormat('₹#,##0.00')
    .setFontWeight('bold')
    .setBackground(UI.TITLE_BG);

  fCell.setFormula('=IFERROR(SUM(F' + rbStart + ':F' + rbEnd + '),0)')
    .setNumberFormat('₹#,##0.00')
    .setFontWeight('bold')
    .setBackground(UI.TITLE_BG);

  out.getRange(totalsRow, 1, 1, 6)
    .setBorder(true, true, true, true, false, false, UI.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  row += 2;

  // Named ranges (best-effort)
  try {
    ssOps.setNamedRange('delivery_total', dCell);
    ssOps.setNamedRange('order_total', eCell);
    ssOps.setNamedRange('balance_total', fCell);
  } catch (e) {}

  return {
    nextRow: row,
    totalsRow,
    deliveryCell: dCell,
    orderCell: eCell,
    balanceCell: fCell
  };
}

/* =========================
   BANDINGS — stop tables only
   ========================= */
function applyBandingToRouteTables_(sheet) {
  (sheet.getBandings() || []).forEach(b => b.remove());

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  const hits = sheet.createTextFinder('Seq').matchCase(false).findAll() || [];
  hits.forEach(cell => {
    const headerRow = cell.getRow();
    let startData = headerRow + 1;
    let endData = startData - 1;

    for (let r = startData; r <= lastRow; r++) {
      const a = (sheet.getRange(r, 1).getDisplayValue() || '').toString().trim();
      if (!a) break;
      if (!/^\d+$/.test(a)) break;
      endData = r;
    }

    if (endData < startData) return;

    const bandRange = sheet.getRange(startData, 1, (endData - startData + 1), 9);
    bandRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  });
}

/* =========================
   LAYOUT / READABILITY
   ========================= */
function applyReadableLayout_(sheet) {
  sheet.setFrozenRows(3);

  const widths = { 1: 90, 2: 180, 3: 150, 4: 380, 5: 85, 6: 140, 7: 95, 8: 300, 9: 160, 10: 110, 11: 140, 12: 110 };
  Object.entries(widths).forEach(([col, w]) => sheet.setColumnWidth(Number(col), w));

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  sheet.getRange(1, 2, lastRow, 1).setWrap(true).setVerticalAlignment('top');
  sheet.getRange(1, 4, lastRow, 1).setWrap(true).setVerticalAlignment('top');
  sheet.getRange(1, 8, lastRow, 1).setWrap(true).setVerticalAlignment('top');
  sheet.getRange(1, 9, lastRow, 1).setWrap(true).setVerticalAlignment('top');

  sheet.getRange(1, 1, lastRow, 1).setHorizontalAlignment('center');
  sheet.getRange(1, 5, lastRow, 1).setHorizontalAlignment('center');
  sheet.getRange(1, 7, lastRow, 1).setHorizontalAlignment('right');
  sheet.getRange(1, 11, lastRow, 2).setHorizontalAlignment('right');

  try {
    sheet.setRowHeights(1, lastRow, 26);
    const tf2 = sheet.createTextFinder('Route: ').matchCase(false);
    (tf2.findAll() || []).forEach(r => sheet.setRowHeight(r.getRow(), 30));
  } catch (e) {}

  sheet.setRowHeight(2, 18);
}

/* =========================
   ORDER PARSING + WEIGHTS
   ========================= */

function parseOrders_(rows, weightIndex) {
  const missingWeights = [];
  const byOrder = {};

  let currentOrderRef = '';
  let currentAddress = '';
  let runningSum = 0;

  rows.slice(1).forEach(r => {
    const rawRef = (r[ORDERS_COL.ORDER_REF] || '').toString().trim();
    const rawProduct = (r[ORDERS_COL.PRODUCT] || '').toString().trim();
    const qty = Number((r[ORDERS_COL.QTY] || 0).toString().replace(/[^0-9.]/g, '')) || 0;
    const addr = (r[ORDERS_COL.ADDRESS] || '').toString().trim();

    const priceCell = (r[ORDERS_COL.PRICE] || '').toString().trim();
    const priceNum = Number(priceCell.replace(/[^0-9.]/g, '') || '0');

    if (rawRef) { currentOrderRef = rawRef; runningSum = 0; }
    if (addr) currentAddress = addr;
    if (!rawProduct) return;

    // "Order value" row
    if (/order value/i.test(rawProduct)) {
      const oFinal = byOrder[currentOrderRef] || baseOrder_(currentOrderRef, currentAddress);
      oFinal.orderValue = priceNum || runningSum || 0;
      byOrder[currentOrderRef] = oFinal;
      return;
    }

    if (SKIP_KEYWORDS.some(rx => rx.test(rawProduct))) return;

    const o = byOrder[currentOrderRef] || baseOrder_(currentOrderRef, currentAddress);
    o.items.push(rawProduct + ' x ' + qty);

    const normalized = normalizeName_(rawProduct);
    const decision = getBestWeightDecision_(normalized, weightIndex);

    if (decision) {
      const grams = decision.grams;
      o.weightKg = r3_(o.weightKg + (grams * qty) / 1000);
    } else {
      const inferred = inferPerGrams_(rawProduct);
      const inferred2 = (inferred != null) ? inferred : inferSpecialCases_(rawProduct);

      if (inferred2 != null) {
        o.weightKg = r3_(o.weightKg + (inferred2 * qty) / 1000);
      } else {
        const suggestion = suggestClosest_(normalized, weightIndex);
        missingWeights.push({ normalized, raw: rawProduct, suggestion });
      }
    }

    if (!isNaN(priceNum) && priceNum > 0) runningSum += priceNum;
    o.orderValue = o.orderValue || runningSum;

    byOrder[currentOrderRef] = o;
  });

  return { byOrder, missingWeights };
}

function baseOrder_(orderRef, address) {
  return {
    orderId: orderRef || ('NO-ID-' + Utilities.getUuid().slice(0, 8)),
    name: extractNameFromAddress_(address),
    address: address,
    pincode: extractPincode_(address),
    city: '',
    state: '',
    items: [],
    weightKg: 0,
    orderValue: 0,
    _loc: null
  };
}

/* =========================
   WEIGHTS INDEX + FUZZY
   ========================= */

function buildFuzzyWeightIndex_(ssWeights) {
  const tabs = ssWeights.getSheets();

  const pickTab = (sheet) => {
    const hdr = sheet.getRange(1, 1, 1, Math.min(12, sheet.getMaxColumns()))
      .getValues()[0].map(v => (v || '').toString().toLowerCase().trim());
    const joined = hdr.join('|');
    return joined.includes('product name') && (joined.includes('unit weight') || joined.includes('weight') || joined.includes('grams'));
  };

  const tab = tabs.find(pickTab) || tabs[0];
  const data = tab.getDataRange().getValues();
  if (data.length < 2) return { entries: [], exactMap: {} };

  const hdr = data[0].map(v => (v || '').toString().toLowerCase().trim());

  const colPN  = hdr.findIndex(h => h === 'product name');
  const colPM  = hdr.findIndex(h => h === 'product name mod');
  const colPOG = hdr.findIndex(h => h === 'product list og');
  const colUG  = hdr.findIndex(h => h.includes('unit weight') || h === 'grams' || h.includes('weight'));

  const entries = [];
  const exactMap = {};

  data.slice(1).forEach(r => {
    const pn  = colPN  >= 0 ? (r[colPN]  || '').toString().trim() : '';
    const pm  = colPM  >= 0 ? (r[colPM]  || '').toString().trim() : '';
    const pog = colPOG >= 0 ? (r[colPOG] || '').toString().trim() : '';

    const grams = Number((colUG >= 0 ? r[colUG] : '').toString().replace(/[^0-9.]/g, ''));
    if (!grams || grams <= 0) return;

    const variants = [pn, pm, pog].filter(Boolean);
    if (!variants.length) return;

    const normSet = new Set(variants.map(v => normalizeName_(v)).filter(Boolean));
    if (!normSet.size) return;

    const tokenSet = new Set(Array.from(normSet).flatMap(n => n.split(' ').filter(Boolean)));

    const entry = { name: variants[0], variants, grams, normSet, tokenSet };
    entries.push(entry);

    normSet.forEach(n => exactMap[n] = grams);
  });

  return { entries, exactMap };
}

function getBestWeightDecision_(normalized, weightIndex) {
  if (!normalized) return null;
  if (weightIndex.exactMap[normalized] != null) return { grams: weightIndex.exactMap[normalized], how: 'exact' };

  // substring
  let best = null, bestScore = -1;
  weightIndex.entries.forEach(e => {
    const s = bestSubstringScore_(normalized, e.normSet);
    if (s > bestScore) { bestScore = s; best = e; }
  });
  if (best && bestScore >= 0.9) return { grams: best.grams, how: 'substring' };

  // token jaccard
  best = null; bestScore = -1;
  const nTokens = new Set(normalized.split(' ').filter(Boolean));
  weightIndex.entries.forEach(e => {
    const s = jaccard_(nTokens, e.tokenSet);
    if (s > bestScore) { bestScore = s; best = e; }
  });
  if (best && bestScore >= 0.6) return { grams: best.grams, how: 'token' };

  // levenshtein
  best = null;
  let bestDist = Number.POSITIVE_INFINITY;
  weightIndex.entries.forEach(e => {
    const shortest = Array.from(e.normSet).reduce((a, b) => a.length <= b.length ? a : b);
    const d = levenshtein_(normalized, shortest);
    if (d < bestDist) { bestDist = d; best = e; }
  });
  if (best && bestDist <= Math.max(2, Math.floor(normalized.length * 0.25))) {
    return { grams: best.grams, how: 'lev' };
  }

  return null;
}

function suggestClosest_(normalized, weightIndex) {
  if (!normalized) return null;

  let best = null, bestScore = -1;
  weightIndex.entries.forEach(e => {
    const s = bestSubstringScore_(normalized, e.normSet);
    if (s > bestScore) { bestScore = s; best = e; }
  });
  if (best && bestScore >= 0.6) return { name: best.name, grams: best.grams, how: 'substring' };

  best = null; bestScore = -1;
  const nTokens = new Set(normalized.split(' ').filter(Boolean));
  weightIndex.entries.forEach(e => {
    const s = jaccard_(nTokens, e.tokenSet);
    if (s > bestScore) { bestScore = s; best = e; }
  });
  if (best && bestScore >= 0.4) return { name: best.name, grams: best.grams, how: 'token' };

  best = null;
  let bestDist = Number.POSITIVE_INFINITY;
  weightIndex.entries.forEach(e => {
    const shortest = Array.from(e.normSet).reduce((a, b) => a.length <= b.length ? a : b);
    const d = levenshtein_(normalized, shortest);
    if (d < bestDist) { bestDist = d; best = e; }
  });

  return best ? { name: best.name, grams: best.grams, how: 'lev' } : null;
}

/* =========================
   PINCODE GEO INDEX
   ========================= */
function buildPincodeGeoIndex_(ssPins) {
  const out = { pinToGeo: {}, pinToCity: {} };

  ssPins.getSheets().forEach(sh => {
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return;

    const hdr = values[0].map(v => (v || '').toString().trim().toLowerCase());

    const cPin = hdr.indexOf('pincode');
    const cCity = hdr.indexOf('city');
    const cLat = hdr.indexOf('latitude');
    const cLng = hdr.indexOf('longitude');
    const cStatus = hdr.indexOf('latlng status');

    if (cPin < 0 || cLat < 0 || cLng < 0) return;

    values.slice(1).forEach(r => {
      const pin = onlyDigits_(r[cPin]);
      if (!pin) return;

      if (cStatus >= 0) {
        const st = (r[cStatus] || '').toString().trim().toLowerCase();
        if (st && st !== 'ok') return;
      }

      const lat = Number(r[cLat]);
      const lng = Number(r[cLng]);
      if (!isFinite(lat) || !isFinite(lng)) return;

      out.pinToGeo[pin] = { lat, lng };

      if (cCity >= 0) {
        const city = (r[cCity] || '').toString().trim();
        if (city) out.pinToCity[pin] = city;
      }
    });
  });

  return out;
}

/* =========================
   FILE + SHEET HELPERS
   ========================= */

function ensureDeliveryRoutesAsThird_(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  try {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(3);
  } catch (e) {}
  return sh;
}

function getSingleSheetFile_(folder, fileName) {
  const it = folder.getFilesByName(fileName);
  if (!it.hasNext()) throw new Error('File "' + fileName + '" not found in folder');
  const f = it.next();
  if (it.hasNext()) Logger.log('Multiple "' + fileName + '" found; using first.');
  if (f.getMimeType() !== MimeType.GOOGLE_SHEETS) throw new Error('"' + fileName + '" is not a Google Sheet.');
  return f;
}

function getOrdersSheetSmart_(ss, preferredName, zoneNamesSet, outputName) {
  if (preferredName) {
    const exact = ss.getSheetByName(preferredName);
    if (exact && !exact.isSheetHidden()) return exact;
  }

  const candidates = ss.getSheets().filter(s =>
    !s.isSheetHidden() && !zoneNamesSet.has(s.getName()) && s.getName() !== outputName
  );

  const KEYWORDS = ['product', 'products', 'qty', 'quantity', 'price', 'amount', 'address'];
  let best = null;
  let bestScore = -1;

  candidates.forEach(sh => {
    const maxCols = Math.min(12, sh.getMaxColumns());
    const header = sh.getRange(1, 1, 1, maxCols).getValues()[0]
      .map(v => (v || '').toString().toLowerCase().trim());

    const score = KEYWORDS.reduce((acc, k) => acc + (header.some(h => h.includes(k)) ? 1 : 0), 0);
    if (score > bestScore) { bestScore = score; best = sh; }
  });

  if (best && bestScore > 0) return best;

  const fuzzy = candidates.find(s => /(order|orders|woo|export|daily)/i.test(s.getName()));
  if (fuzzy) return fuzzy;

  return candidates.length ? candidates[0] : null;
}

/**
 * Requires Advanced Drive service enabled (Drive API).
 */
function getLatestOpsSpreadsheetDriveApi_(prefix, preferredFolderId) {
  prefix = (prefix || 'ops_').toLowerCase();

  const res = Drive.Files.list({
    q: [
      "mimeType='application/vnd.google-apps.spreadsheet'",
      "trashed=false",
      "'" + preferredFolderId + "' in parents"
    ].join(' and '),
    orderBy: 'modifiedDate desc',
    maxResults: 10,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true
  });

  const items = res.items || [];
  const hit = items.find(f =>
    (f.title || '').toLowerCase().startsWith(prefix)
  );

  if (!hit) throw new Error("No ops_ sheet found in operational folder");

  return SpreadsheetApp.openById(hit.id);
}

/* =========================
   STRING + INFER HELPERS
   ========================= */

function normalizeName_(name) {
  return name.toString().toLowerCase()
    .replace(/\(.*?\)/g, ' ')
    .replace(/[\u2013\u2014_\/-]/g, ' ')
    .replace(/\s{2,}/g, ' ')
    .replace(/[^a-z0-9\s]/g, '')
    .trim();
}

function inferPerGrams_(productName) {
  const s = (productName || '').toString();

  let m = s.match(/\bper\s+(\d{2,4})\s*(g|gm|gms)\b/i);
  if (m && m[1]) return Number(m[1]);

  m = s.match(/\((\d{2,4})\s*(g|gm|gms)\)/i);
  if (m && m[1]) return Number(m[1]);

  m = s.match(/\b(\d{2,4})\s*(g|gm|gms)\b/i);
  if (m && m[1]) return Number(m[1]);

  return null;
}

function inferSpecialCases_(name) {
  const s = (name || '').toString().toLowerCase();

  if (s.includes('microgreen') || s.includes('microgreens')) {
    const g = inferPerGrams_(name);
    if (g) return g;
    if (s.includes('50')) return 50;
    if (s.includes('100')) return 100;
    return 100;
  }
  return null;
}

function extractNameFromAddress_(addr) {
  const lines = (addr || '').split(/\r?\n/).map(x => x.trim()).filter(Boolean);
  return lines.length ? lines[0] : '';
}

function extractPhone_(addr) {
  const m = (addr || '').match(/\b(\+?\d{10,13})\b/);
  return m ? m[1] : '';
}

function extractPincode_(addr) {
  const m = (addr || '').match(/\b(\d{6})\b/);
  return m ? m[1] : '';
}

function onlyDigits_(v) {
  return (v || '').toString().replace(/\D+/g, '');
}

function r3_(n) {
  return Math.round((Number(n) || 0) * 1000) / 1000;
}

function formatWeightHuman_(kg) {
  const n = Number(kg) || 0;
  if (n < 1) return Math.round(n * 1000) + ' g';
  return r3_(n) + ' kg';
}

/* =========================
   FUZZY HELPERS
   ========================= */

function bestSubstringScore_(normalized, normSet) {
  let sc = 0;
  normSet.forEach(n => {
    if (!n || !normalized) return;
    if (n === normalized) sc = Math.max(sc, 1);
    if (n.includes(normalized) || normalized.includes(n)) {
      const lenShort = Math.min(n.length, normalized.length);
      const lenLong = Math.max(n.length, normalized.length);
      sc = Math.max(sc, lenShort / lenLong);
    }
  });
  return sc;
}

function jaccard_(aSet, bSet) {
  if (!aSet || !bSet || aSet.size === 0 || bSet.size === 0) return 0;
  let inter = 0;
  aSet.forEach(t => { if (bSet.has(t)) inter++; });
  const uni = aSet.size + bSet.size - inter;
  return inter / (uni || 1);
}

function levenshtein_(a, b) {
  if (a === b) return 0;
  const m = a.length, n = b.length;
  if (!m) return n;
  if (!n) return m;

  const dp = Array(n + 1).fill(0);
  for (let j = 0; j <= n; j++) dp[j] = j;

  for (let i = 1; i <= m; i++) {
    let prev = dp[0];
    dp[0] = i;
    for (let j = 1; j <= n; j++) {
      const tmp = dp[j];
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[j] = Math.min(dp[j] + 1, dp[j - 1] + 1, prev + cost);
      prev = tmp;
    }
  }
  return dp[n];
}

function dedupeMissing_(arr) {
  const seen = new Set();
  const out = [];
  (arr || []).forEach(x => {
    if (x && x.normalized && !seen.has(x.normalized)) {
      seen.add(x.normalized);
      out.push(x);
    }
  });
  return out;
}

/* =========================
   TITLE HELPER
   ========================= */
function chunkArray_(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}


function mergeTitle_(sheet, row, colStart, colEnd, text) {
  const width = colEnd - colStart + 1;
  const r = sheet.getRange(row, colStart, 1, width).merge();
  r.setValue(text)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(UI.TITLE_BG);
}

/* =========================
   WEB APP SAFE: EXCEL → OPS
   ========================= */

/**
 * Web-app safe: if input is Excel, convert → Google Sheet in same folder,
 * then run routes on the Google Sheet ID (not Excel ID).
 */
function buildUserFriendlyRoutesFromExcel(fileId) {
  if (!fileId) throw new Error("Missing fileId");

  const f = DriveApp.getFileById(fileId);
  const mime = f.getMimeType();

  let targetSheetId = fileId;

  if (mime === MimeType.MICROSOFT_EXCEL) {
    const existingOps = findLatestOpsInSameFolderByDriveApp_(fileId, (ROUTES_CFG && ROUTES_CFG.OPS_NAME_PREFIX) || "ops_");
    if (existingOps) {
      targetSheetId = existingOps.getId();
    } else {
      targetSheetId = convertExcelToGoogleSheetSameFolder_(fileId);
    }
  }

const ss = SpreadsheetApp.openById(targetSheetId);
const msg = buildUserFriendlyRoutes(ss.getId());

  return {
    ok: true,
    message: msg,
    opsSheetId: targetSheetId,
    opsSheetUrl: "https://docs.google.com/spreadsheets/d/" + targetSheetId + "/edit"
  };
}

/**
 * DriveApp-based reuse: find latest ops_ in same folder
 */
function findLatestOpsInSameFolderByDriveApp_(fileId, prefix) {
  prefix = (prefix || "ops_").toLowerCase();

  const src = DriveApp.getFileById(fileId);
  const parents = src.getParents();
  if (!parents.hasNext()) return null;

  const folder = parents.next();
  const it = folder.getFiles();

  let best = null;
  while (it.hasNext()) {
    const x = it.next();
    if (x.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;

    const name = (x.getName() || "").toLowerCase();
    if (!name.startsWith(prefix)) continue;

    if (!best || x.getLastUpdated() > best.getLastUpdated()) best = x;
  }

  return best; // ✅ THIS WAS MISSING
}

/**
 * Requires Advanced Drive service enabled (Drive.Files.*)
 * Converts Excel → Google Sheet in the SAME folder
 */
function convertExcelToGoogleSheetSameFolder_(excelFileId) {
  // Requires Advanced Drive service enabled (Drive API)

  const excelMeta = Drive.Files.get(excelFileId, { supportsAllDrives: true });
  const parents = excelMeta.parents || [];
  if (!parents.length) throw new Error("Excel file has no parent folder.");

  const parentFolderId = parents[0].id;

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss");
  const newName = (ROUTES_CFG && ROUTES_CFG.OPS_NAME_PREFIX ? ROUTES_CFG.OPS_NAME_PREFIX : "ops_") + ts;

  const converted = Drive.Files.copy(
    {
      title: newName,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{ id: parentFolderId }]
    },
    excelFileId,
    { supportsAllDrives: true }
  );

  if (!converted || !converted.id) throw new Error("Excel → Google Sheet conversion failed.");
  return converted.id;
}
function buildRoutesEntry(fileId) {
  return buildUserFriendlyRoutesFromExcel(fileId);
}

