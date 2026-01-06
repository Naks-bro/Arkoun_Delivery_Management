/*************************************************
 * POPULATE VENDOR ORDERS (AUTO-FIND LATEST ops_)
 * Arkoun Farms – Vendor POs + WhatsApp copy
 *
 * Sheets / Files:
 * - Latest ops_* file in folder: PREFERRED_FOLDER_ID
 *   - "Orders List" (input)
 *   - "Vendor Order List" (output; created/overwritten)
 * - "Vendor_Product_Mapping" (in Operational_ARKOUN folder)
 * - "salad breaker" (in Operational_ARKOUN folder)
 *
 * Vendor_Product_Mapping columns (0-based):
 *  0: Sr No
 *  1: Product Name mod
 *  2: Product List og
 *  3: Quantity / Units  (per ONE unit – special list for unit-only products)
 *  4: Vendor           (location / market; used for grouping – e.g. Andheri)
 *  5: Vendor_details   (contact name for WhatsApp, e.g. Adil / Andheri Market)
 *  6: Phone
 *  7: Lang (English / Hindi)
 *  8: Hindi Product Name
 *************************************************/

function populateVendorOrdersWithSalads() {
  try {
    Logger.log("🚀 Starting Vendor Order List population with salads and WhatsApp messages...");

    // ===== CONFIG =====
    const VENDOR_MAPPING_FOLDER_ID = '1XO9LJ3DEZW3LqDPdpLa3dXx6WTyCYe4B'; // Operational_ARKOUN
    const VENDOR_MAPPING_FILE_NAME = 'Vendor_Product_Mapping';
    const VENDOR_MAPPING_SHEET_NAME = 'Sheet1';

    const SALAD_BREAKER_FILE_NAME   = 'salad breaker';
    const SALAD_BREAKER_SHEET_NAME  = 'Sheet1';

    const ORDERS_LIST_SHEET        = 'Orders List';
    const VENDOR_ORDER_LIST_SHEET  = 'Vendor Order List';

    // Orders List (0-based)
    const PRODUCT_COLUMN_INDEX  = 2; // column C
    const QUANTITY_COLUMN_INDEX = 3; // column D

    const SKIP_KEYWORDS       = ['order value', 'total value', 'total'];
    const OPS_NAME_PREFIX     = 'ops_';
    const PREFERRED_FOLDER_ID = '1EkUtycNhatLV_hU7AWfjrwayvNIna-eK';

    const MESSAGE_COL       = 5;   // WhatsApp message column (E) in table 2
    const MESSAGE_COL_WIDTH = 320; // px

    // === 1) Find latest ops sheet automatically ===
    const ss = getLatestOpsSpreadsheet_(OPS_NAME_PREFIX, PREFERRED_FOLDER_ID);
    const ordersSheet = ss.getSheetByName(ORDERS_LIST_SHEET);
    if (!ordersSheet) throw new Error(`Sheet "${ORDERS_LIST_SHEET}" not found in ${ss.getName()}`);

    const vendorOrderSheet =
      ss.getSheetByName(VENDOR_ORDER_LIST_SHEET) || ss.insertSheet(VENDOR_ORDER_LIST_SHEET);
    vendorOrderSheet.clear();

    // === 1A) TIMESTAMP + BUFFER BLOCK ===
    const now = new Date();
    const stamp = Utilities.formatDate(
      now,
      Session.getScriptTimeZone() || 'Asia/Kolkata',
      'dd-MMM-yyyy HH:mm:ss'
    );
    vendorOrderSheet.getRange('F1').setValue('Last Run')
      .setFontWeight('bold')
      .setFontColor('#ffffff')
      .setBackground('#333333');
    vendorOrderSheet.getRange('G1').setValue(stamp)
      .setFontWeight('bold')
      .setBackground('#eeeeee');

    vendorOrderSheet.getRange('F2').setValue('Buffer %')
      .setFontWeight('bold')
      .setBackground('#d9d9d9');

    const bufferCell = vendorOrderSheet.getRange('G2');
    let bufferPercentRaw = bufferCell.getValue();
    let bufferPercent;

    // If empty → default 12%
    if (bufferPercentRaw === '' || bufferPercentRaw == null) {
      bufferPercent = 12;
      bufferCell.setValue(bufferPercent);
    } else {
      bufferPercent = Number(bufferPercentRaw);
      if (isNaN(bufferPercent)) {
        bufferPercent = 12;
      }
      // If user typed 12% and cell formatted as percent (stored as 0.12)
      if (bufferPercent > 0 && bufferPercent < 1) {
        bufferPercent = bufferPercent * 100;
      }
    }
    Logger.log(`📌 Using buffer % = ${bufferPercent}`);

    // === 2) Read vendor mapping + salad breaker ===
    const vendorFile = getFirstSpreadsheetByExactName_(VENDOR_MAPPING_FOLDER_ID, VENDOR_MAPPING_FILE_NAME);
    const vendorMapSheet = vendorFile.getSheetByName(VENDOR_MAPPING_SHEET_NAME);
    if (!vendorMapSheet) throw new Error(`Sheet "${VENDOR_MAPPING_SHEET_NAME}" not found in ${VENDOR_MAPPING_FILE_NAME}`);
    const vendorData = vendorMapSheet.getDataRange().getValues();

    const saladFile = getFirstSpreadsheetByExactName_(VENDOR_MAPPING_FOLDER_ID, SALAD_BREAKER_FILE_NAME);
    const saladSheet = saladFile.getSheetByName(SALAD_BREAKER_SHEET_NAME);
    if (!saladSheet) throw new Error(`Sheet "${SALAD_BREAKER_SHEET_NAME}" not found in ${SALAD_BREAKER_FILE_NAME}`);
    const saladData = saladSheet.getDataRange().getValues();

    const ordersData = ordersSheet.getDataRange().getValues();

    // === 3) Helpers ===

    const normalize = s => s
      ? s.toString().toLowerCase()
          .replace(/\(.*?\)/g, '')
          .replace(/[^a-z0-9\s]/g, ' ')
          .replace(/\s+/g, ' ')
          .trim()
      : '';

    const baseNormalize = s => {
      let n = normalize(s);
      if (!n) return '';

      n = n
        .replace(/\bper\s+\d+(\.\d+)?\s*(gms?|grams?|gm|kg|kgs?|kilo|kilos?)\b/g, ' ')
        .replace(/\b\d+(\.\d+)?\s*(gms?|grams?|gm|kg|kgs?|kilo|kilos?)\b/g, ' ')
        .replace(/\b\d+(\.\d+)?\s*(units?|unit|pcs?|pieces?|piece|bunch(?:es)?|head(?:s)?|pack|packs)\b/g, ' ')
        .replace(/\b(approx)\b/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

      return n;
    };

    const canonicalizeVendor = v => {
      let t = (v || '').toString().trim();
      if (!t) return 'Unknown';

      // strip "hindi" markers
      t = t.replace(/\bhindi\b/i, '').trim();

      // Aggressive normalization for Andheri misspellings
      if (/^anderi$/i.test(t)) return 'Andheri';
      if (/^andheri$/i.test(t)) return 'Andheri';
      if (/^ander/i.test(t))    return 'Andheri'; // "Anderi", "Anderhi", etc.

      return t || 'Unknown';
    };

    // Build total from a unit pattern for WhatsApp (e.g. 2 × "6 pcs" → "12 pcs")
    const makeSUnitString_ = (units, qtyPatternRaw) => {
      const u = Number(units) || 0;
      let p = (qtyPatternRaw || '').toString().trim();
      if (!p) return u ? String(u) : '';

      // Convert word-numbers at start to digits (one, two, etc.)
      const wordMap = {
        one: 1, two: 2, three: 3, four: 4, five: 5,
        six: 6, seven: 7, eight: 8, nine: 9, ten: 10
      };
      p = p.replace(/^(one|two|three|four|five|six|seven|eight|nine|ten)\b/i, m => String(wordMap[m.toLowerCase()]));

      const m = p.match(/(\d+(?:\.\d+)?)\s*([A-Za-z].*)/);
      if (!m) return p; // fallback if pattern is unexpected

      const per = Number(m[1]) || 0;
      if (!per || !u) return p;
      const word = m[2].trim();
      const total = per * u;
      return `${total} ${word}`;
    };

    /**
     * FORMAT WEIGHT (GRAMS → DISPLAY STRING)
     *  - input: grams (number)
     *  - if < 1000  → "XXX gms" (rounded integer)
     *  - if ≥ 1000 → "X kg" (up to 2 decimals, no .00)
     */
    const fmtWeight = g => {
      const n = Number(g) || 0;
      if (n === 0) return '0 gms';

      if (n < 1000) {
        const gInt = Math.round(n);
        return `${gInt} gms`;
      }

      const kg = n / 1000;
      const sKg = Number.isInteger(kg) ? kg.toString() : kg.toFixed(2).replace(/\.00$/, '');
      return `${sKg} kg`;
    };

    // parse grams embedded in product name → returns grams (number) or null
    function parseGramsFromName_(name) {
      if (!name) return null;
      const s = name.toString();

      const patterns = [
        /\(.*?\bper\s+(\d{1,4}(?:\.\d+)?)\s*g(?:rams|ms)?\b.*?\)/i,
        /\(\s*(\d{1,4}(?:\.\d+)?)\s*g(?:rams|ms)?\s*\)/i,
        /\b(\d{1,4}(?:\.\d+)?)\s*g(?:rams|ms)?\b/i
      ];
      for (const rx of patterns) {
        const m = s.match(rx);
        if (m && m[1]) return Number(m[1]);
      }

      const kgMatch = s.match(/\b(\d+(?:\.\d+)?)\s*kg\b/i);
      if (kgMatch && kgMatch[1]) {
        return Number(kgMatch[1]) * 1000;
      }

      return null;
    }

    function cleanDisplayName_(canonicalName, perUnitGMaybe) {
      let s = (canonicalName || '').toString();
      let g = perUnitGMaybe;
      if (!g || g <= 0) {
        g = parseGramsFromName_(s);
      }
      if (g && g > 0) {
        s = s.replace(/\([^)]*(kg|gms?|grams?)\)/gi, '').trim();
        s = s.replace(/[-–]\s*\d+(?:\.\d+)?\s*kg/gi, '').trim();
        s = s.replace(/\s+/g, ' ').replace(/\s*-\s*$/,'').trim();
        return `${s} (${fmtWeight(g)})`;
      }
      return s;
    }

    function parseQtyCellToGrams_(val) {
      if (!val) return null;
      const s = val.toString().toLowerCase().trim();
      if (!s || s === '—' || s === '-') return null;

      // if only units / pieces and no g/kg → treat as unit-only (no grams)
      if (/(unit|units|pcs?|pieces?|piece|bunch(?:es)?|head(?:s)?|pack|packs)/i.test(s) &&
          !/g\b|gms|gram|kg/i.test(s)) {
        return null;
      }

      const numMatches = s.match(/(\d+(?:\.\d+)?)/g);
      if (!numMatches || !numMatches.length) return null;
      const first = Number(numMatches[0]);

      if (/kg/i.test(s)) return first * 1000;
      if (/g\b|gms|gram/i.test(s)) return first;

      return null;
    }

    function isMicrogreensSubscription_(name) {
      const n = normalize(name);
      return n.includes('microgreen') && n.includes('subscription');
    }

    // === 4) Build vendor product map + Hindi name map ===
    const productVendorMapNorm = {}; // normalized → { vendorKey, phone, lang, canonical, perUnitG, unitPattern }
    const productVendorMapBase = {}; // baseNormalize → same
    const vendorMeta          = {};  // vendorKey OR contactName → { phone, lang, contactName }
    const hindiNameByBase     = {};  // baseNormalize(EnglishName) → Hindi name

    vendorData.slice(1).forEach(r => {
      const prodMod         = r[1];
      const prodOg          = r[2];
      const qtyCellRaw      = r[3];
      const vendorLocation  = r[4]; // physical vendor / market (Andheri, Dadar, Adil, etc.)
      const vendorDetailRaw = r[5]; // WhatsApp contact name ("Adil", "Andheri", etc.)
      const phone           = (r[6] || '').toString().trim();
      const langCellRaw     = (r[7] || '').toString().trim();
      const hindiNameRaw    = r[8];

      const vendorTextRaw = (vendorLocation || vendorDetailRaw || '').toString();

      // Lang logic:
      let lang = langCellRaw || 'English';
      if (langCellRaw.toLowerCase() === 'hindi' || /\bhindi\b/i.test(vendorTextRaw)) {
        lang = 'Hindi';
      }

      const vendorKey   = canonicalizeVendor(vendorLocation || vendorDetailRaw || 'Unknown');
      const contactName = vendorDetailRaw ? vendorDetailRaw.toString().trim() : vendorKey;
      const contactKey  = contactName || vendorKey; // used so phones also map to contact person

      const qtyStr  = (qtyCellRaw || '').toString();
      const qtyNorm = qtyStr.toLowerCase();
      const hasUnitWord   = /(unit|units|pcs?|pieces?|piece|bunch(?:es)?|head(?:s)?|pack|packs)/i.test(qtyNorm);
      const hasWeightWord = /(g\b|gms|gram|kg)/i.test(qtyNorm);

      let perUnitGFromQty = null;
      let unitPattern = '';

      if (hasUnitWord && !hasWeightWord) {
        // pure unit-only products (like "6 pcs", "1 head", "one bunch")
        unitPattern = qtyStr.trim();
        perUnitGFromQty = null;
      } else {
        // weight-based products
        perUnitGFromQty = parseQtyCellToGrams_(qtyStr);
      }

      // fill meta for vendor location
      if (!vendorMeta[vendorKey]) {
        vendorMeta[vendorKey] = { phone: '', lang, contactName };
      }
      if (phone) vendorMeta[vendorKey].phone = vendorMeta[vendorKey].phone || phone;
      if (lang.toLowerCase() === 'hindi') vendorMeta[vendorKey].lang = 'Hindi';
      if (contactName && !vendorMeta[vendorKey].contactName) {
        vendorMeta[vendorKey].contactName = contactName;
      }

      // ALSO map meta for contact person (Adil, etc.) so their row gets phone
      if (!vendorMeta[contactKey]) {
        vendorMeta[contactKey] = { phone: '', lang, contactName };
      }
      if (phone) vendorMeta[contactKey].phone = vendorMeta[contactKey].phone || phone;
      if (lang.toLowerCase() === 'hindi') vendorMeta[contactKey].lang = 'Hindi';
      if (contactName && !vendorMeta[contactKey].contactName) {
        vendorMeta[contactKey].contactName = contactName;
      }

      const prodCandidates = [prodMod, prodOg];
      prodCandidates.forEach(p => {
        if (!p) return;
        const pStr = p.toString().trim();
        const n = normalize(pStr);
        const b = baseNormalize(pStr);

        const info = {
          vendor: vendorKey,
          phone,
          lang,
          canonical: pStr,
          perUnitG: perUnitGFromQty,
          unitPattern
        };

        if (n && !productVendorMapNorm[n]) {
          productVendorMapNorm[n] = info;
        } else if (n && perUnitGFromQty && productVendorMapNorm[n] && !productVendorMapNorm[n].perUnitG) {
          productVendorMapNorm[n].perUnitG = perUnitGFromQty;
        }

        if (b && !productVendorMapBase[b]) {
          productVendorMapBase[b] = info;
        } else if (b && perUnitGFromQty && productVendorMapBase[b] && !productVendorMapBase[b].perUnitG) {
          productVendorMapBase[b].perUnitG = perUnitGFromQty;
        }

        const hindiName = hindiNameRaw ? hindiNameRaw.toString().trim() : '';
        if (hindiName && b && !hindiNameByBase[b]) {
          hindiNameByBase[b] = hindiName;
        }
      });
    });

    function getLocalizedName_(canonicalEngName, lang) {
      if (!lang || lang.toLowerCase() !== 'hindi') return canonicalEngName;
      const base = baseNormalize(canonicalEngName);
      if (!base) return canonicalEngName;
      const mapped = hindiNameByBase[base];
      return mapped || canonicalEngName;
    }

    // === 5) Parse salad breaker ===
    const saladMap = {};
    const saladMapNorm = {};

    let currentSalad = '';
    for (let i = 1; i < saladData.length; i++) {
      const row = saladData[i];

      const allEmpty = row.every(x => (x === '' || x == null));
      if (allEmpty) { currentSalad = ''; continue; }

      const saladNameCell = (row[1] || '').toString().trim();
      if (saladNameCell) currentSalad = saladNameCell;
      if (!currentSalad) continue;

      const ing  = (row[2] || '').toString().trim();
      const qStr = (row[3] || '').toString().trim();
      const vRaw = (row[4] || '').toString().trim();
      if (!ing) continue;

      const q = parseFloat(qStr.replace(/[^0-9.]/g, '')) || 0;
      const vend = canonicalizeVendor(vRaw);

      if (!saladMap[currentSalad]) saladMap[currentSalad] = [];
      saladMap[currentSalad].push({ ing, q, v: vend });

      const n = normalize(currentSalad);
      if (!saladMapNorm[n]) saladMapNorm[n] = saladMap[currentSalad];
    }

    // === 6) Walk Orders ===
    if (ordersData.length <= 1) throw new Error(`"${ORDERS_LIST_SHEET}" appears empty (no data rows).`);

    // vendorTotals[vendorKey][canonicalEngName] = { units (base), perUnitG, unitPattern }
    const vendorTotals = {};
    const unmatched = [];

    const add = (vendor, canonicalEngName, units, perUnitGOrNull, unitPatternOrNull) => {
      vendor = canonicalizeVendor(vendor);
      if (!vendorTotals[vendor]) vendorTotals[vendor] = {};
      if (!vendorTotals[vendor][canonicalEngName]) {
        vendorTotals[vendor][canonicalEngName] = {
          units: 0,
          perUnitG: perUnitGOrNull,
          unitPattern: unitPatternOrNull
        };
      }
      vendorTotals[vendor][canonicalEngName].units += Number(units || 0);
      if (perUnitGOrNull != null) vendorTotals[vendor][canonicalEngName].perUnitG = perUnitGOrNull;
      if (!vendorTotals[vendor][canonicalEngName].unitPattern && unitPatternOrNull) {
        vendorTotals[vendor][canonicalEngName].unitPattern = unitPatternOrNull;
      }
    };

    const matchVendorForProduct = (originalName) => {
      const n = normalize(originalName);
      if (!n) return null;

      if (n.includes('salad') || n.includes('box')) {
        Logger.log(`ℹ️ Skipping vendor fuzzy match for salad/box: "${originalName}"`);
        return null;
      }

      if (productVendorMapNorm[n]) {
        const { vendor, phone, lang, perUnitG, canonical, unitPattern } = productVendorMapNorm[n];
        Logger.log(`✅ Exact vendor match: "${originalName}" → "${vendor}"`);
        return { vendor: canonicalizeVendor(vendor), phone, lang, perUnitG, canonical, unitPattern };
      }

      const baseOrder = baseNormalize(originalName);
      if (!baseOrder) {
        Logger.log(`⚠️ Empty base name for "${originalName}"`);
        return null;
      }

      if (productVendorMapBase[baseOrder]) {
        const info = productVendorMapBase[baseOrder];
        const vendor = canonicalizeVendor(info.vendor);
        Logger.log(`✅ Base-name match: "${originalName}" → "${info.canonical}" | vendor="${vendor}"`);
        return {
          vendor,
          phone: info.phone,
          lang: info.lang,
          perUnitG: info.perUnitG,
          canonical: info.canonical,
          unitPattern: info.unitPattern
        };
      }

      const orderTokens = new Set(
        baseOrder.split(' ').filter(t => t && t.length > 2)
      );
      if (orderTokens.size === 0) {
        Logger.log(`⚠️ No meaningful tokens for "${originalName}"`);
        return null;
      }

      let strictMatchInfo = null;
      let strictMatchKey  = null;

      Object.keys(productVendorMapBase).forEach(baseMapKey => {
        const info = productVendorMapBase[baseMapKey];
        if (!info) return;

        const mapTokens = baseMapKey.split(' ').filter(t => t && t.length > 2);

        const orderCore = orderTokens;
        const mapCore   = new Set(mapTokens);

        let shared = 0;
        orderCore.forEach(t => { if (mapCore.has(t)) shared++; });

        if (shared === orderCore.size && shared > 0) {
          strictMatchInfo = info;
          strictMatchKey  = baseMapKey;
        }
      });

      if (strictMatchInfo) {
        const vendor = canonicalizeVendor(strictMatchInfo.vendor);
        Logger.log(
          `✅ Core-word base match: "${originalName}" → "${strictMatchKey}" | vendor="${vendor}"`
        );
        return {
          vendor,
          phone: strictMatchInfo.phone,
          lang:  strictMatchInfo.lang,
          perUnitG: strictMatchInfo.perUnitG,
          canonical: strictMatchInfo.canonical,
          unitPattern: strictMatchInfo.unitPattern
        };
      }

      let bestInfo  = null;
      let bestScore = -1;
      let bestKey   = null;

      Object.keys(productVendorMapBase).forEach(baseMapKey => {
        const info = productVendorMapBase[baseMapKey];
        if (!info) return;

        const baseMap = baseMapKey;
        const mapTokens = new Set(
          baseMap.split(' ').filter(t => t && t.length > 2)
        );
        if (mapTokens.size === 0) return;

        let shared = 0;
        orderTokens.forEach(t => { if (mapTokens.has(t)) shared++; });
        if (!shared) return;

        const sub  = _substringSim_(baseOrder, baseMap);
        const jac  = _jaccardSim_(orderTokens, mapTokens);
        const lev  = _levenshteinSim_(baseOrder, baseMap);
        const init = _initialsSim_(orderTokens, mapTokens);

        const score = (sub * 0.30) + (jac * 0.30) + (lev * 0.20) + (init * 0.20);

        if (score > bestScore) {
          bestScore = score;
          bestInfo  = info;
          bestKey   = baseMap;
        }
      });

      const MIN_SCORE = 0.30;

      if (bestInfo && bestScore >= MIN_SCORE) {
        const vendor = canonicalizeVendor(bestInfo.vendor);
        Logger.log(
          `🔍 Fuzzy vendor match: "${originalName}" → "${bestKey}" | vendor="${vendor}" | score=${bestScore.toFixed(3)}`
        );
        return {
          vendor,
          phone: bestInfo.phone,
          lang:  bestInfo.lang,
          perUnitG: bestInfo.perUnitG,
          canonical: bestInfo.canonical,
          unitPattern: bestInfo.unitPattern
        };
      }

      Logger.log(`⚠️ No reliable vendor match for: "${originalName}"`);
      return null;
    };

    // Walk through Orders
    ordersData.slice(1).forEach(r => {
      const prod = r[PRODUCT_COLUMN_INDEX];
      const q    = r[QUANTITY_COLUMN_INDEX];
      if (!prod || !q) return;

      const p = prod.toString().trim();
      const units = Number((q || '').toString().replace(/[^0-9.]/g, '')) || 0;
      if (!units) return;
      if (SKIP_KEYWORDS.some(k => p.toLowerCase().includes(k))) return;

      const pNorm = normalize(p);
      const saladLines = saladMap[p] || saladMapNorm[pNorm];

      if (saladLines) {
        // Salad → expand into ingredients (grams only, no S-units)
        saladLines.forEach(i => {
          const perUnitG = Number(i.q) || 0;
          add(i.v, i.ing, units, perUnitG, null);
        });
        return;
      }

      const info = matchVendorForProduct(p);
      if (!info) {
        unmatched.push([p, units, 'Order', 'No vendor match']);
        return;
      }

      let perUnitG = null;
      if (!isMicrogreensSubscription_(p)) {
        const fromName = parseGramsFromName_(p);
        if (fromName != null && fromName > 0) {
          perUnitG = fromName;
        } else if (info.perUnitG != null && info.perUnitG > 0) {
          perUnitG = info.perUnitG;
        }
      }

      const canonicalName = info.canonical || p;

      add(info.vendor, canonicalName, units, perUnitG, info.unitPattern || null);
    });

    // === 7) Output tables ===
    let row = 3; // row1 timestamp, row2 buffer row

    // ---- Table 1: VENDOR ORDER LIST (Units + S-Units + Weight) ----
    const titleRow1 = row++;

    vendorOrderSheet
      .getRange(titleRow1, 1, 1, 3)          // A:C
      .merge()
      .setValue('VENDOR ORDER LIST')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center');

    vendorOrderSheet.getRange(row, 1, 1, 7)
      .setValues([[
        'Sr No',
        'Product Name',
        'Vendor',
        'Units',
        'S-Units',
        'Weight (Base)',
        `Weight (+${bufferPercent}%)`
      ]])
      .setFontWeight('bold')
      .setBackground('#d9ead3');

    const t1Start = row; row++;

    let sr = 1;
    const vendors = Object.keys(vendorTotals).sort();

    vendors.forEach(vendor => {
      Object.keys(vendorTotals[vendor]).sort().forEach(canonicalEngName => {
        const obj = vendorTotals[vendor][canonicalEngName];
        const displayNameEng = cleanDisplayName_(canonicalEngName, obj.perUnitG);

        const units = obj.units || 0;
        const hasWeight =
          obj.perUnitG != null &&
          obj.perUnitG > 0 &&
          !isMicrogreensSubscription_(canonicalEngName);

        let sUnits = '';
        let baseWeightStr = '';
        let buffWeightStr = '';

        if (hasWeight) {
          // Weighted products: convert units → weight and apply buffer on WEIGHT
          const baseWeightG = units * obj.perUnitG;
          const buffWeightG = baseWeightG * (1 + bufferPercent / 100);

          baseWeightStr = fmtWeight(baseWeightG);
          buffWeightStr = fmtWeight(buffWeightG);
        } else {
          // Unit-only products: show TOTAL pattern in S-Units (e.g. 4 × "6 pcs" → "24 pcs")
          const totalPattern = obj.unitPattern
            ? makeSUnitString_(units, obj.unitPattern)
            : '';

          sUnits = totalPattern || '—';
          baseWeightStr = '—';
          buffWeightStr = '—';
        }

        vendorOrderSheet.getRange(row++, 1, 1, 7)
          .setValues([[sr++, displayNameEng, vendor, units, sUnits, baseWeightStr, buffWeightStr]]);
      });
    });

    if (row - 1 >= t1Start) {
      const lastRow = row - 1;
      vendorOrderSheet.getRange(`A${t1Start}:G${lastRow}`)
        .setBorder(true, true, true, true, true, true);

      // Highlight rows where both weight columns are effectively empty (unit-only)
      for (let r = t1Start + 1; r <= lastRow; r++) {
        const baseW = vendorOrderSheet.getRange(r, 6).getValue();
        const buffW = vendorOrderSheet.getRange(r, 7).getValue();
        const isEmptyBase   = baseW === '' || baseW == null || baseW === '—';
        const isEmptyBuffer = buffW === '' || buffW == null || buffW === '—';
        if (isEmptyBase && isEmptyBuffer) {
          vendorOrderSheet.getRange(r, 1, 1, 7).setBackground('#fff2cc');
        }
      }
    }

    vendorOrderSheet.appendRow(['']);
    row++;

    // ---- Table 2: WHATSAPP MESSAGES ----
    vendorOrderSheet.getRange(row++, 1).setValue('WHATSAPP MESSAGES')
      .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
    vendorOrderSheet.getRange(row, 1, 1, 5)
      .setValues([['S no', 'Vendor Name', 'Phone Number', 'Language', 'Message']])
      .setFontWeight('bold').setBackground('#fce5cd');
    const t2Start = row; row++;

    // Build message parts for each vendor
    const vendorMsgParts = {}; // vendorKey → { header, bodySimple, footer, lang }

    vendors.forEach(vendor => {
      const meta = vendorMeta[vendor] || { phone: '', lang: 'English', contactName: vendor };
      const lang = (meta.lang || 'English').trim();
      const isHindi = lang.toLowerCase() === 'hindi';

      const itemsInfo = Object.entries(vendorTotals[vendor])
        .sort((a, b) => a[0].localeCompare(b[0]))
        .map(([canonicalEngName, obj]) => {
          const baseUnits = obj.units || 0;

          const hasWeight =
            obj.perUnitG != null &&
            obj.perUnitG > 0 &&
            !isMicrogreensSubscription_(canonicalEngName);

          const localizedBaseName = getLocalizedName_(canonicalEngName, lang);

          let simpleLine;

          if (hasWeight) {
            // use BUFFERED units in weight for vendor reference
            const qtyBuffered = Math.ceil(baseUnits * (1 + bufferPercent / 100));
            const totalG = obj.perUnitG * qtyBuffered;
            const labelWeight = isHindi ? 'कुल' : 'Total';

            // Weighted items: show only total weight (vendor doesn’t care about units)
            simpleLine = `• ${localizedBaseName} – ${labelWeight} ${fmtWeight(totalG)}`;
          } else {
            // UNIT-ONLY ITEMS (like Elaichi Big Banana 6pcs)
            const sUnitsString = makeSUnitString_(baseUnits, obj.unitPattern || '');  // e.g. 2 × "6pcs" → "12 pcs"
            if (isHindi) {
              const unitsWordHi = 'यूनिट';
              const unitsPartHi = `${baseUnits} ${unitsWordHi}`;
              const sUnitsPartHi = (sUnitsString && sUnitsString !== '—') ? sUnitsString : unitsPartHi;
              simpleLine = `• ${localizedBaseName} – ${sUnitsPartHi}`;
            } else {
              const unitWordEn = baseUnits === 1 ? 'unit' : 'units';
              const unitsPartEn = `${baseUnits} ${unitWordEn}`;
              const sUnitsPartEn = (sUnitsString && sUnitsString !== '—') ? sUnitsString : unitsPartEn;
              simpleLine = `• ${localizedBaseName} – ${sUnitsPartEn}`;
            }

            // Microgreens subscription: override to just show units, no weight
            if (isMicrogreensSubscription_(canonicalEngName)) {
              if (isHindi) {
                const unitsWordHi = 'यूनिट';
                simpleLine = `• ${localizedBaseName} – ${baseUnits} ${unitsWordHi}`;
              } else {
                const unitWordEn = baseUnits === 1 ? 'unit' : 'units';
                simpleLine = `• ${localizedBaseName} – ${baseUnits} ${unitWordEn}`;
              }
            }
          }

          return { simpleLine };
        });

      const bodySimple = itemsInfo.map(i => i.simpleLine).join('\n');

      // 👉 Header now always uses VENDOR NAME (location) to avoid "प्रिय Adil" under Andheri row
      const headerName = vendor;
      const header = isHindi
        ? `प्रिय ${headerName},\nआपका कल का ऑर्डर (${vendor}):\n\n`
        : `Dear ${headerName},\nYour order for tomorrow (${vendor}):\n\n`;

      const footer = isHindi
        ? `\n\nकृपया कन्फर्म करें।\nधन्यवाद,\nArkoun Farms`
        : `\n\nPlease confirm.\nThanks,\nArkoun Farms`;

      vendorMsgParts[vendor] = { header, bodySimple, footer, lang };
    });

    // Write messages – 🔥 NO MORE SPECIAL COMBINED MESSAGE FOR ADIL
    let msgSr = 1;
    vendors.forEach(vendor => {
      const meta  = vendorMeta[vendor] || { phone: '', lang: 'English', contactName: vendor };
      const parts = vendorMsgParts[vendor];
      if (!parts) return;

      const lang  = (parts.lang || 'English').trim();
      const text  = parts.header + parts.bodySimple + parts.footer;

      vendorOrderSheet.getRange(row++, 1, 1, 5)
        .setValues([[msgSr++, vendor, meta.phone || '', lang || 'English', text]]);
    });

    if (row - 1 >= t2Start) vendorOrderSheet.getRange(`A${t2Start}:E${row - 1}`)
      .setBorder(true, true, true, true, true, true);
    vendorOrderSheet.appendRow(['']); row++;

// ---- Table 3: UNMATCHED ITEMS ----
const unmatchedTitleRow = row++;
vendorOrderSheet.getRange(unmatchedTitleRow, 1, 1, 3)  // A:C merged
  .merge()
  .setValue('UNMATCHED ITEMS')
  .setFontWeight('bold')
  .setFontSize(14)
  .setHorizontalAlignment('center');
vendorOrderSheet.getRange(row, 1, 1, 5)
  .setValues([['S no', 'Item Name', 'Quantity', 'Source', 'Note']])
  .setFontWeight('bold')
  .setBackground('#f4cccc');
const t3Start = row; 
row++;


    if (unmatched.length === 0) {
      vendorOrderSheet.getRange(row++, 1, 1, 5)
        .setValues([[1, '—', '—', '—', 'All items matched']]);
    } else {
      unmatched.forEach((u, i) => {
        vendorOrderSheet.getRange(row++, 1, 1, 5).setValues([[i + 1, ...u]]);
      });
    }
    if (row - 1 >= t3Start) vendorOrderSheet.getRange(`A${t3Start}:E${row - 1}`)
      .setBorder(true, true, true, true, true, true);

    // neat widths & readability
    vendorOrderSheet.setFrozenRows(3); // headers + timestamp + buffer visible
    vendorOrderSheet.autoResizeColumns(1, 7);
    try {
      vendorOrderSheet.setColumnWidth(1, 80);    // Sr no
      vendorOrderSheet.setColumnWidth(2, 320);   // Product Name
      vendorOrderSheet.setColumnWidth(3, 110);   // Vendor
      vendorOrderSheet.setColumnWidth(4, 80);    // Units
      vendorOrderSheet.setColumnWidth(5, 120);   // S-Units
      vendorOrderSheet.setColumnWidth(6, 110);   // Weight (Base)
      vendorOrderSheet.setColumnWidth(7, 110);   // Weight (+%)
      vendorOrderSheet.setColumnWidth(MESSAGE_COL, MESSAGE_COL_WIDTH);
      const last = vendorOrderSheet.getLastRow();
      if (last > 1) {
        vendorOrderSheet.getRange(1, 2, last, 1).setWrap(true); // product names wrap
      }
    } catch (e) {
      Logger.log('ℹ️ Layout tweak skipped: ' + e.message);
    }

    Logger.log(`✅ Wrote Vendor Order List for latest ops sheet: ${ss.getName()} (${ss.getId()})`);
  } catch (err) {
    Logger.log('❌ Error: ' + err.message);
    throw err;
  }
}

/**
 * Returns the latest Spreadsheet whose title contains 'ops_' (prefix).
 */
function getLatestOpsSpreadsheet_(prefix, rootFolderId) {
  prefix = (prefix || 'ops_').toLowerCase();

  const root = DriveApp.getFolderById(rootFolderId);

  let bestFile = null;
  let bestTime = 0;

  function walk_(folder) {
    // scan files in this folder
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      const name = (f.getName() || '').toLowerCase();
      if (name.startsWith(prefix) && f.getMimeType() === MimeType.GOOGLE_SHEETS) {
        const t = f.getLastUpdated().getTime();
        if (t > bestTime) {
          bestTime = t;
          bestFile = f;
        }
      }
    }

    // go into subfolders
    const subs = folder.getFolders();
    while (subs.hasNext()) {
      walk_(subs.next());
    }
  }

  walk_(root);

  if (!bestFile) throw new Error(`No Google Sheet found starting with "${prefix}" inside root tree.`);
  Logger.log(`📄 Latest ops sheet: ${bestFile.getName()} (${bestFile.getId()})`);
  return SpreadsheetApp.openById(bestFile.getId());
}


/**
 * Finds the first Google Sheet by exact file name inside a given folder.
 */
function getFirstSpreadsheetByExactName_(folderId, exactName) {
  const folder = DriveApp.getFolderById(folderId);
  const it = folder.getFilesByName(exactName);
  if (!it.hasNext()) {
    throw new Error(`File named "${exactName}" not found in folder ${folder.getName()} (${folderId})`);
  }
  const file = it.next();
  return SpreadsheetApp.openById(file.getId());
}

/*************************************************
 * FUZZY SIMILARITY HELPERS
 *************************************************/

function _substringSim_(a, b) {
  a = (a || '').toString();
  b = (b || '').toString();
  if (!a || !b) return 0;
  if (a === b) return 1;

  if (a.includes(b) || b.includes(a)) {
    const lenShort = Math.min(a.length, b.length);
    const lenLong  = Math.max(a.length, b.length);
    return lenShort / (lenLong || 1);
  }
  return 0;
}

function _jaccardSim_(aTokens, bTokens) {
  if (!aTokens || !bTokens || aTokens.size === 0 || bTokens.size === 0) return 0;

  let inter = 0;
  aTokens.forEach(t => { if (bTokens.has(t)) inter++; });
  const uni = aTokens.size + bTokens.size - inter;
  return uni ? inter / uni : 0;
}

function _levenshteinSim_(a, b) {
  a = (a || '').toString();
  b = (b || '').toString();
  if (a === b) return 1;
  const m = a.length;
  const n = b.length;
  if (!m || !n) return 0;

  const dp = new Array(n + 1);
  for (let j = 0; j <= n; j++) dp[j] = j;

  for (let i = 1; i <= m; i++) {
    let prev = dp[0];
    dp[0] = i;
    for (let j = 1; j <= n; j++) {
      const tmp = dp[j];
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[j] = Math.min(
        dp[j] + 1,
        dp[j - 1] + 1,
        prev + cost
      );
      prev = tmp;
    }
  }

  const dist = dp[n];
  const maxLen = Math.max(m, n);
  return maxLen ? 1 - dist / maxLen : 0;
}

function _initialsSim_(aTokens, bTokens) {
  if (!aTokens || !bTokens || aTokens.size === 0 || bTokens.size === 0) return 0;

  const aInit = new Set();
  const bInit = new Set();

  aTokens.forEach(t => {
    const c = t.charAt(0);
    if (c) aInit.add(c);
  });
  bTokens.forEach(t => {
    const c = t.charAt(0);
    if (c) bInit.add(c);
  });

  let inter = 0;
  aInit.forEach(c => { if (bInit.has(c)) inter++; });

  const uni = aInit.size + bInit.size - inter;
  return uni ? inter / uni : 0;
}
