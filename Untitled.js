function testOrsProxyMatrix() {
  const base = "https://twilight-field-7f06.kokatenakul11.workers.dev/"; // your worker
  const url = base.replace(/\/+$/,'/') + "matrix";

  // Hub MIDC Andheri (approx) -> Juhu (approx)
  const payload = {
    locations: [
      [72.8693634, 19.1204828], // HUB (lng,lat)
      [72.8267098, 19.1048146]  // Juhu (lng,lat)
    ],
    metrics: ["duration", "distance"]
  };

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  Logger.log("HTTP: " + res.getResponseCode());
  Logger.log(res.getContentText().slice(0, 1000));
}
