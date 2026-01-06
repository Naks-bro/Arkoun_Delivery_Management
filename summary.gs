/*************************************************
 * SUMMARY (4th TAB) — Active OR Recursive Search
 * - Uses ACTIVE spreadsheet if present
 * - Else recursively finds latest ops_ sheet
 * - Matches Arkoun folder rules exactly
 *************************************************/

const LIVE_SUMMARY_CFG = {
  ROOT_FOLDER_ID: '1EkUtycNhatLV_hU7AWfjrwayvNIna-eK', // Arkoun Farms Clientsheets
  OPS_PREFIX: 'ops_',
  SUMMARY_SHEET_NAME: 'Summary',
  ORDERS_SHEET_NAME: 'Orders List',
  ROUTES_SHEET_NAME: 'Delivery Routes'
};

function buildSummaryLiveFormulas2() {

  /* ================= HELPERS ================= */

  function getTargetSpreadsheet(cfg) {
    // 1️⃣ Active spreadsheet (normal ops flow)
    try {
      const active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) {
        Logger.log('✅ Using ACTIVE spreadsheet');
        return active;
      }
    } catch (_) {}

    // 2️⃣ Recursive search under ROOT
    Logger.log('🔍 Searching recursively for latest ops_ sheet...');
    const root = DriveApp.getFolderById(cfg.ROOT_FOLDER_ID);
    const result = findLatestOpsRecursive(root, cfg.OPS_PREFIX);

    if (!result) {
      throw new Error('No ops_ Google Sheet found anywhere under root folder');
    }

    Logger.log(`📄 Using ops sheet: ${result.file.getName()}`);
    return SpreadsheetApp.openById(result.file.getId());
  }

  function findLatestOpsRecursive(folder, prefix) {
    let latest = null;

    // Skip Operational_ARKOUN explicitly
    if (folder.getName() === 'Operational_ARKOUN') return null;

    // Check files
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      const f = files.next();
      if (!f.getName().toLowerCase().startsWith(prefix)) continue;

      const t = f.getLastUpdated().getTime();
      if (!latest || t > latest.time) {
        latest = { file: f, time: t };
      }
    }

    // Recurse folders
    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      const sub = subfolders.next();
      const candidate = findLatestOpsRecursive(sub, prefix);
      if (candidate && (!latest || candidate.time > latest.time)) {
        latest = candidate;
      }
    }

    return latest;
  }

  function getOrdersSheetSmart(ss, preferredName, outputName) {
    const exact = ss.getSheetByName(preferredName);
    if (exact && !exact.isSheetHidden()) return exact;

    const candidates = ss.getSheets().filter(s =>
      !s.isSheetHidden() && s.getName() !== outputName
    );

    const KEYS = ['product','qty','quantity','price','amount','address'];
    let best = null, scoreBest = -1;

    candidates.forEach(sh => {
      const header = sh.getRange(1,1,1,Math.min(12, sh.getMaxColumns()))
        .getValues()[0]
        .map(v => (v||'').toString().toLowerCase());
      const score = KEYS.filter(k => header.some(h => h.includes(k))).length;
      if (score > scoreBest) { scoreBest = score; best = sh; }
    });

    return best || candidates[0] || null;
  }

  function ensureSheetIndex(ss, name, index1Based) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    try {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(index1Based);
    } catch (_) {}
    return sh;
  }

  function escapeTab(name) {
    return name.replace(/'/g, "''");
  }

  function friendlyToday() {
    const d = new Date();
    const tz = Session.getScriptTimeZone();
    const s=["th","st","nd","rd"], v=d.getDate()%100;
    const ord=d.getDate()+(s[(v-20)%10]||s[v]||s[0]);
    return `${ord} ${Utilities.formatDate(d,tz,'MMMM yyyy')}`;
  }

  /* ================= MAIN ================= */

  const C = LIVE_SUMMARY_CFG;
  const ss = getTargetSpreadsheet(C);

  const ordersSheet = getOrdersSheetSmart(ss, C.ORDERS_SHEET_NAME, C.SUMMARY_SHEET_NAME);
  if (!ordersSheet) throw new Error('Orders sheet not found');

  const routesSheet = ss.getSheetByName(C.ROUTES_SHEET_NAME);
  if (!routesSheet) throw new Error('Delivery Routes sheet not found');

  const ordersTab = escapeTab(ordersSheet.getName());
  const routesTab = escapeTab(routesSheet.getName());

  const sh = ensureSheetIndex(ss, C.SUMMARY_SHEET_NAME, 4);
  sh.clear();

  /* ================= SUMMARY UI ================= */

  let r = 1;
  sh.getRange(r,1,1,3).merge()
    .setValue(`Deliveries ${friendlyToday()}`)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setBackground('#d9edf7');

  r += 2;

  const rows = {
    noOrders: r,
    orderValue: r+1,
    procurement: r+2,
    delivery: r+3,
    revenue: r+4
  };

  sh.getRange(rows.noOrders,1,5,1).setValues([
    ['No of Orders'],
    ['Order Value'],
    ['Procurement cost'],
    ['Delivery cost'],
    ['Revenue']
  ]).setFontWeight('bold');

  sh.getRange(rows.noOrders,2)
    .setFormula(`=COUNTIF('${ordersTab}'!C:C,"*Order Value*")`);

  sh.getRange(rows.orderValue,2)
    .setFormula(`=IFERROR(order_total,
      SUMIF('${ordersTab}'!C:C,"*Order Value*", '${ordersTab}'!E:E))`);

  sh.getRange(rows.procurement,2).setValue(0);

  sh.getRange(rows.delivery,2)
    .setFormula(`=IFERROR(delivery_total,
      SUMIF('${routesTab}'!C:C,"*TOTALS*", '${routesTab}'!D:D))`);

  sh.getRange(rows.revenue,2)
    .setFormula(`=B${rows.orderValue}-B${rows.procurement}-B${rows.delivery}`);

  ['orderValue','procurement','delivery','revenue'].forEach(k =>
    sh.getRange(rows[k],3).setValue('₹')
  );

  Logger.log('✅ Summary generated successfully.');
}
