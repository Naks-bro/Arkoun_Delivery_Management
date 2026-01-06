/*************************************************
 * 🌿 ARKOUN OPS CONSOLE – MASTER SCRIPT (CLEAN)
 * - No duplicates
 * - Routes logic lives in delivery.gs (optional)
 * - Map backend reads:
 *    1) "_Route Map Data" (best)
 *    2) Delivery Routes sheet + Pincodes file (fallback)
 *************************************************/

// =======================
// 🧩 OPS CONFIGURATION
// =======================
const OPS_CONFIG = {
  ROOT_FOLDER_ID: '1EkUtycNhatLV_hU7AWfjrwayvNIna-eK',
  OPERATIONAL_FOLDER_ID: '1XO9LJ3DEZW3LqDPdpLa3dXx6WTyCYe4B',
  EXCEL_PREFIX: 'orders_',
  TARGET_SHEET_NAMES: ['Orders List', 'Vendor Order List', 'Delivery Routes', 'Summary']
};

// =======================
// 🚀 DAILY CONVERSION
// =======================
function dailyExcelConversion(fileId) {
  const now = new Date();
  const tz = Session.getScriptTimeZone();

  const currentDate = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const currentMonth = Utilities.formatDate(now, tz, 'MMMM');
  const currentYear = Utilities.formatDate(now, tz, 'yyyy');
  const timestamp = Utilities.formatDate(now, tz, 'yyyy-MM-dd_HH-mm-ss');

const root = DriveApp.getFolderById(OPS_CONFIG.ROOT_FOLDER_ID);
  const yearFolder = getOrCreateFolder_(root, currentYear);
  const monthFolder = getOrCreateFolder_(yearFolder, `Orders_${currentMonth}`);

  let excelFile = fileId
    ? DriveApp.getFileById(fileId)
    : null;

  if (!excelFile) {
    const it = monthFolder.getFilesByType(MimeType.MICROSOFT_EXCEL);
    while (it.hasNext()) {
      const f = it.next();
      if (f.getName().includes(currentDate)) {
        excelFile = f;
        break;
      }
    }
  }

  if (!excelFile) throw new Error("No Excel file found");

  const blob = excelFile.getBlob();

  const resource = {
    name: `ops_${timestamp}`,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [monthFolder.getId()]
  };

  const created = Drive.Files.create(resource, blob);

  const ss = SpreadsheetApp.openById(created.id);
  ss.getSheets()[0].setName("Orders List");

  ["Vendor Order List", "Delivery Routes", "Summary"].forEach(n => {
    if (!ss.getSheetByName(n)) ss.insertSheet(n);
  });

  return `✅ Converted: ${ss.getUrl()}`;
}


// =======================
// 🔧 HELPERS
// =======================
function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// =======================
// 🌐 WEB APP (CLEAN)
// =======================
function doGet(e) {
  const page = String(e?.parameter?.page || "index").toLowerCase();

  if (page === "map") {
    return HtmlService.createHtmlOutputFromFile("Map")
      .setTitle("Arkoun Routes Map")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Arkoun Ops Console")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// =======================
// 📄 FILE LISTING
// =======================
function listOpsFiles() {
  const out = [];
  scanFolder_(DriveApp.getFolderById(OPS_CONFIG.ROOT_FOLDER_ID), out);
  return out;
}

function scanFolder_(folder, out) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() === MimeType.MICROSOFT_EXCEL || f.getMimeType() === MimeType.GOOGLE_SHEETS) {
      out.push({ id: f.getId(), name: f.getName() });
    }
  }
  const sub = folder.getFolders();
  while (sub.hasNext()) scanFolder_(sub.next(), out);
}

// =======================
// 🔘 BUTTON ACTION FUNCTIONS
// =======================
function convertToOps(fileId) {
  return dailyExcelConversion(fileId);
}

function populateVendors(fileId) {
  // vendor.gs must have populateVendorOrdersWithSalads(fileId)
  return populateVendorOrdersWithSalads(fileId);
}

function buildRoutes(fileId) {
  return buildRoutesEntry(fileId);
}


function generateSummary(fileId) {
  // summary.gs must have buildSummaryLiveFormulas2(fileId)
  return buildSummaryLiveFormulas2(fileId);
}

// =======================
// 🔎 Latest ops_ sheet (optional helper)
// =======================
function getLatestOpsSpreadsheet_(prefix, rootFolderId) {
  prefix = (prefix || 'ops_').toLowerCase();
  const root = DriveApp.getFolderById(rootFolderId);

  let bestFile = null;
  let bestTime = 0;

  function walk_(folder) {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const name = (f.getName() || '').toLowerCase();
      if (name.startsWith(prefix) && f.getMimeType() === MimeType.GOOGLE_SHEETS) {
        const t = f.getLastUpdated().getTime();
        if (t > bestTime) { bestTime = t; bestFile = f; }
      }
    }
    const subs = folder.getFolders();
    while (subs.hasNext()) walk_(subs.next());
  }

  walk_(root);

  if (!bestFile) throw new Error(`No Google Sheet found starting with "${prefix}" inside root tree.`);
  return SpreadsheetApp.openById(bestFile.getId());
}


/*************************************************
 * ✅ MAP BACKEND (single source of truth)
 * Priority:
 *  1) _Route Map Data (generated by delivery.gs)
 *  2) Fallback: parse "Delivery Routes" blocks + pincode lookup from Pincodes file
 *************************************************/
function getMapDataForFile(fileId) {
  if (!fileId) throw new Error("fileId missing");

  const ss = SpreadsheetApp.openById(fileId);

  // fixed hub (you can replace later with geocoding if needed)
  const hub = { lat: 19.10885, lng: 72.8662, label: "Arkoun Dispatch Hub" };

  // 1) ✅ BEST: _Route Map Data
  const mapSheet = ss.getSheetByName("_Route Map Data");
  if (mapSheet) {
    const res = buildMapDataFromRouteMapSheet_(mapSheet, hub);
    if (res.routes.length) return res;
  }

  // 2) Fallback: Delivery Routes + Pincodes file
  const delivSheet = findDeliveryRoutesSheet_(ss);
  const grid = delivSheet.getDataRange().getValues();

  const parsed = parseDeliveryRoutesBlocks_(grid);
  const pinMap = loadPincodeLatLngMapByFileName_();

  const routes = [];
  (parsed.routes || []).forEach((r, idx) => {
    const color = pickColor_(idx);
    const stops = [];

    (r.stops || []).forEach(s => {
      const pin = String(s.pincode || "").trim();
      const ll = pinMap.get(pin);
      if (!ll) return;
      stops.push({
        name: s.customer || "",
        pincode: pin,
        lat: ll.lat,
        lng: ll.lng
      });
    });

    if (stops.length) routes.push({ id: r.id, color, stops });
  });

  if (!routes.length) {
    throw new Error(
      "No plotted stops found.\n" +
      "Reason: _Route Map Data missing AND Delivery Routes fallback couldn't plot.\n" +
      "Fix: Run 'Build Delivery Routes (ORS)' so _Route Map Data is created, OR ensure Pincodes file has lat/lng for all pins."
    );
  }

  return { hub, routes };
}

// ============ SOURCE A: _Route Map Data ============
function buildMapDataFromRouteMapSheet_(sh, hub) {
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return { hub, routes: [] };

  const hdr = vals[0].map(x => String(x || "").trim().toLowerCase());
  const cRoute = hdr.indexOf("route id");
  const cSeq   = hdr.indexOf("seq");
  const cName  = hdr.indexOf("customer");
  const cPin   = hdr.indexOf("pincode");
  const cLat   = hdr.indexOf("lat");
  const cLng   = hdr.indexOf("lng");

  if ([cRoute, cSeq, cLat, cLng].some(i => i < 0)) return { hub, routes: [] };

  const routeMap = new Map();

  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    const rid = String(r[cRoute] || "").trim();
    if (!rid) continue;

    const lat = Number(r[cLat]);
    const lng = Number(r[cLng]);
    if (!isFinite(lat) || !isFinite(lng)) continue;

    if (!routeMap.has(rid)) {
      routeMap.set(rid, { id: rid, color: pickColor_(routeMap.size), stops: [] });
    }

    routeMap.get(rid).stops.push({
      name: cName >= 0 ? String(r[cName] || "").trim() : "",
      pincode: cPin >= 0 ? String(r[cPin] || "").trim() : "",
      lat, lng,
      seq: Number(r[cSeq]) || 0
    });
  }

  const routes = Array.from(routeMap.values()).map(rt => {
    rt.stops.sort((a,b) => (a.seq||0) - (b.seq||0));
    rt.stops = rt.stops.map(s => ({ lat:s.lat, lng:s.lng, name:s.name, pincode:s.pincode }));
    return rt;
  });

  return { hub, routes };
}

// ============ DELIVERY ROUTES finder ============
function findDeliveryRoutesSheet_(ss) {
  const names = ss.getSheets().map(s => s.getName());
  const hit =
    names.find(n => n.toLowerCase() === "delivery routes") ||
    names.find(n => n.toLowerCase().includes("delivery routes")) ||
    names.find(n => n.toLowerCase().includes("delivery route"));

  if (!hit) throw new Error("No 'Delivery Routes' sheet found. Found: " + names.join(", "));
  return ss.getSheetByName(hit);
}

// ============ PINCODES lookup by FILE NAME ============
function loadPincodeLatLngMapByFileName_() {
  const PINCODES_FILE_NAME = "Pincodes";

  const folderId = OPS_CONFIG.OPERATIONAL_FOLDER_ID || "";
  let fileId = null;

  // Prefer Operational folder
  if (folderId) {
    const folder = DriveApp.getFolderById(folderId);
    const it = folder.getFilesByName(PINCODES_FILE_NAME);
    if (it.hasNext()) fileId = it.next().getId();
  }

  // Fallback: Drive search
  if (!fileId) {
    const it2 = DriveApp.searchFiles(
      `title = "${PINCODES_FILE_NAME.replace(/"/g, '\\"')}" and mimeType = "${MimeType.GOOGLE_SHEETS}" and trashed=false`
    );
    if (it2.hasNext()) fileId = it2.next().getId();
  }

  if (!fileId) throw new Error(`Pincodes file not found by name: ${PINCODES_FILE_NAME}`);

  const ss = SpreadsheetApp.openById(fileId);
  const sh = ss.getSheetByName("pincodes") || ss.getSheetByName("Pincodes") || ss.getSheets()[0];

  const v = sh.getDataRange().getValues();
  if (v.length < 2) throw new Error("Pincodes sheet has no data rows.");

  const headers = v[0].map(h => String(h || "").trim().toLowerCase());
  const iPin = headers.findIndex(h => h.includes("pincode") || h === "pin");
  const iLat = headers.findIndex(h => h === "lat" || h.includes("latitude"));
  const iLng = headers.findIndex(h => h === "lng" || h === "lon" || h.includes("longitude") || h.includes("long"));

  if (iPin < 0 || iLat < 0 || iLng < 0) {
    throw new Error("Pincodes file must contain columns: Pincode, Latitude, Longitude");
  }

  const map = new Map();
  for (let r = 1; r < v.length; r++) {
    const pin = String(v[r][iPin] || "").trim();
    const lat = Number(v[r][iLat]);
    const lng = Number(v[r][iLng]);
    if (!pin || !isFinite(lat) || !isFinite(lng)) continue;
    map.set(pin, { lat, lng });
  }
  return map;
}

// ============ Fallback parser: Delivery Routes blocks ============
function parseDeliveryRoutesBlocks_(grid) {
  let hub = null;
  const routes = [];

  const rowText_ = (row) => row.map(c => String(c || "").trim()).join(" | ");

  // hub (optional)
  for (let r = 0; r < Math.min(grid.length, 30); r++) {
    const t = rowText_(grid[r]).toLowerCase();
    if (t.includes("dispatch hub")) {
      hub = { lat: 19.10885, lng: 72.8662, label: "Arkoun Dispatch Hub" };
      break;
    }
  }

  // route blocks
  for (let r = 0; r < grid.length; r++) {
    const firstCell = String(grid[r][0] || "").trim();
    if (!firstCell) continue;

    if (firstCell.startsWith("Route:")) {
      const routeId = extractRouteId_(firstCell);
      if (!routeId) continue;

      let headerRow = -1;
      for (let k = r + 1; k < Math.min(r + 10, grid.length); k++) {
        const h0 = String(grid[k][0] || "").trim().toLowerCase();
        if (h0 === "seq") { headerRow = k; break; }
      }
      if (headerRow === -1) continue;

      const headers = grid[headerRow].map(x => String(x || "").trim().toLowerCase());
      const idxSeq  = headers.indexOf("seq");
      const idxCust = headers.indexOf("customer");
      const idxPin  = headers.indexOf("pincode");
      if (idxPin === -1) continue;

      const stops = [];
      for (let rr = headerRow + 1; rr < grid.length; rr++) {
        const a0 = String(grid[rr][0] || "").trim();
        if (a0.startsWith("Route:") || a0.toUpperCase().includes("UNASSIGNED ORDERS")) break;

        const rowHasAny = grid[rr].some(c => String(c || "").trim() !== "");
        if (!rowHasAny) break;

        const seqVal = idxSeq !== -1 ? String(grid[rr][idxSeq] || "").trim() : "";
        const pinVal = String(grid[rr][idxPin] || "").trim();
        if (!pinVal) continue;
        if (idxSeq !== -1 && !seqVal) continue;

        stops.push({
          customer: idxCust !== -1 ? String(grid[rr][idxCust] || "").trim() : "",
          pincode: pinVal
        });
      }

      routes.push({ id: routeId, stops });
    }
  }

  return { hub, routes };
}

function extractRouteId_(routeLine) {
  const m = String(routeLine).match(/Route:\s*([A-Za-z0-9_-]+)/);
  return m ? m[1].trim() : "";
}

function pickColor_(i) {
  const palette = ["#22c55e", "#3b82f6", "#f59e0b", "#ef4444", "#a855f7", "#14b8a6", "#f97316", "#e11d48"];
  return palette[i % palette.length];
}
