function validateAndFixPincodeLatLng() {
  const PINCODES_SHEET_ID = '1uImutfd98vo7bYWJdVxnTDdXfr_LUM4UCv84g3_-M40'; // your Pincodes file
  const ZONE_TABS = ['Dahisar to Colaba','Mulund to CST','Airoli to Kharghar'];

  const ss = SpreadsheetApp.openById(PINCODES_SHEET_ID);

  ZONE_TABS.forEach(tabName => {
    const sh = ss.getSheetByName(tabName);
    if (!sh) throw new Error('Missing tab: ' + tabName);

    const rng = sh.getDataRange();
    const values = rng.getValues();
    if (values.length < 2) return;

    // Expected: A sr, B pincode, C city, D state, E lat, F lng
    // We'll write status in G (col 7)
    const outStatus = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      const pin = onlyDigits_(row[1]);
      let lat = parseFloat(row[4]);
      let lng = parseFloat(row[5]);

      let status = 'OK';

      if (!pin) status = 'NO_PIN';
      if (isNaN(lat) || isNaN(lng)) status = 'MISSING_LATLNG';
      else if (lat === 0 || lng === 0) status = 'ZERO_LATLNG';
      else {
        // India sanity range check
        const latOk = lat >= 6 && lat <= 38;
        const lngOk = lng >= 68 && lng <= 99;

        // If swapped: lat looks like longitude and lng looks like latitude
        const swappedLikely = (lat >= 68 && lat <= 99) && (lng >= 6 && lng <= 38);

        if (swappedLikely) {
          // AUTO-FIX: swap them
          const tmp = lat; lat = lng; lng = tmp;
          row[4] = lat;
          row[5] = lng;
          status = 'SWAPPED_FIXED';
        } else if (!latOk || !lngOk) {
          status = 'OUT_OF_RANGE';
        }
      }

      outStatus.push([status]);
    }

    // write back corrected lat/lng + status
    // update E:F (lat/lng) if any swaps happened (we updated values array)
    sh.getRange(2, 1, values.length - 1, values[0].length).setValues(values.slice(1));

    // write status in col G
    sh.getRange(1, 7).setValue('LatLng Status');
    sh.getRange(2, 7, outStatus.length, 1).setValues(outStatus);

    Logger.log('Validated tab: ' + tabName);
  });
}

/* needed helper */
function onlyDigits_(v) {
  return (v || '').toString().replace(/\D+/g, '');
}
