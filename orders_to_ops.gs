// =======================
// 🧩 CONFIGURATION
// =======================
const CONFIG = {
  ROOT_FOLDER_ID: '1EkUtycNhatLV_hU7AWfjrwayvNIna-eK', // Arkoun Farms Clientsheets
  OPERATIONAL_FOLDER_ID: '1XO9LJ3DEZW3LqDPdpLa3dXx6WTyCYe4B', // Operational_ARKOUN
  EXCEL_PREFIX: 'orders_',
  TARGET_SHEET_NAMES: [
    'Orders List',
    'Vendor Order List',
    'Delivery Routes',
    'Summary'
  ]
};

// =======================
// 🚀 MAIN FUNCTION
// =======================
function dailyExcelConversion(fileId) {
  const now = new Date();
  const tz = Session.getScriptTimeZone();

  const currentDate  = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const currentMonth = Utilities.formatDate(now, tz, 'MMMM');
  const currentYear  = Utilities.formatDate(now, tz, 'yyyy');
  const timestamp    = Utilities.formatDate(now, tz, 'yyyy-MM-dd_HH-mm-ss');

  Logger.log(`🚀 Excel → Ops conversion started for ${currentDate}`);

  try {
    // =======================
    // 1️⃣ Folder resolution
    // =======================
    const rootFolder  = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    const yearFolder  = getOrCreateFolder_(rootFolder, currentYear);
    const monthFolder = getOrCreateFolder_(yearFolder, `Orders_${currentMonth}`);

    Logger.log(`📁 Target folder: ${monthFolder.getName()}`);

    // =======================
    // 2️⃣ Locate source Excel
    // =======================
    let excelFile = null;

    if (fileId) {
      excelFile = DriveApp.getFileById(fileId);
      Logger.log(`📎 Manual Excel selected: ${excelFile.getName()}`);
    } else {
      const it = monthFolder.getFilesByType(MimeType.MICROSOFT_EXCEL);
      while (it.hasNext()) {
        const f = it.next();
        if (
          f.getName().startsWith(CONFIG.EXCEL_PREFIX) &&
          f.getName().includes(currentDate)
        ) {
          excelFile = f;
          break;
        }
      }
    }

    if (!excelFile) {
      throw new Error(`No Excel file found for ${currentDate}`);
    }

    // =======================
    // 3️⃣ Convert Excel → Google Sheet (Drive v3)
    // =======================
    const blob = excelFile.getBlob();
    const newName = `ops_${timestamp}`;

    const resource = {
      name: newName,
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [monthFolder.getId()] // ✅ v3 format
    };

    const converted = Drive.Files.create(resource, blob, {
      supportsAllDrives: true
    });

    const ss = SpreadsheetApp.openById(converted.id);
    Logger.log(`✅ Converted → ${ss.getName()}`);

    // =======================
    // 4️⃣ Prepare sheets
    // =======================
    const sheets = ss.getSheets();
    sheets[0].setName(CONFIG.TARGET_SHEET_NAMES[0]);

    CONFIG.TARGET_SHEET_NAMES.slice(1).forEach(name => {
      if (!ss.getSheetByName(name)) {
        ss.insertSheet(name);
        Logger.log(`📝 Created sheet: ${name}`);
      }
    });

    // =======================
    // 5️⃣ Ownership (best-effort)
    // =======================
    try {
      const me = Session.getActiveUser().getEmail();
      DriveApp.getFileById(ss.getId()).setOwner(me);
      Logger.log(`👤 Ownership set to ${me}`);
    } catch (e) {
      Logger.log(`ℹ️ Ownership change skipped`);
    }

    // =======================
    // 6️⃣ Done
    // =======================
    const msg =
      `✅ Ops file created\n\n` +
      `📄 ${ss.getName()}\n` +
      `🔗 ${ss.getUrl()}`;

    Logger.log(msg);
    return msg;

  } catch (err) {
    const msg = `❌ dailyExcelConversion failed: ${err.message}`;
    Logger.log(msg + '\n' + err.stack);
    throw err;
  }
}

// =======================
// 🔧 HELPERS
// =======================
function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  Logger.log(`📂 Creating folder: ${name}`);
  return parent.createFolder(name);
}
