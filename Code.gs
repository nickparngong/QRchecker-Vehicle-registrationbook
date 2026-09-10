const SPREADSHEET_ID = '1rfd7D_NraAkv-0zhnKUfWYtGNg0OPsDH0B7QDcsUZ3A';
const VEHICLE_SHEET = 'VEHICLES';
const HISTORY_SHEET = 'SCAN_HISTORY';
const API_KEY = ''; // Optional. Leave empty.


/**
 * ===== AUTHORIZATION / SETUP =====
 * Run authorizeSheetAccess() ONCE from the Apps Script editor.
 * Google will show the permission dialog for the account that owns/deploys
 * this Web App. This prevents the Web App from failing because the script
 * has never been authorized to access the spreadsheet.
 */
function authorizeSheetAccess() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const vehicle = ss.getSheetByName(VEHICLE_SHEET);
  const history = ss.getSheetByName(HISTORY_SHEET);

  if (!vehicle) throw new Error('ไม่พบชีต VEHICLES');
  if (!history) throw new Error('ไม่พบชีต SCAN_HISTORY');

  // Force a real read/write authorization check.
  vehicle.getRange(1, 1).getDisplayValue();
  history.getRange(1, 1).getDisplayValue();

  return 'Authorization สำเร็จ: ' + ss.getName();
}

/**
 * Safe test function. Run after authorizeSheetAccess().
 */
function testGoogleSheetAccess() {
  const result = health_();
  if (!result.ok) throw new Error('Google Sheet access failed');
  return JSON.stringify(result, null, 2);
}

/**
 * Returns a user-friendly authorization status.
 */
function authorizationStatus_() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return {
      authorized: true,
      spreadsheetName: ss.getName(),
      spreadsheetId: ss.getId()
    };
  } catch (err) {
    return {
      authorized: false,
      error: String(err && err.message || err)
    };
  }
}

function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  try {
    const action = String(p.action || 'health').toLowerCase();
    let result;
    if (action === 'health') result = health_();
    else if (action === 'auth') result = authorizationStatus_();
    else if (action === 'getall') result = getAll_();
    else if (action === 'scan') result = scan_(p);
    else result = {ok:false, error:'Unknown action: '+action};
    return output_(result, p.callback);
  } catch (err) {
    return output_({ok:false,error:String(err && err.message || err)}, p.callback);
  }
}

function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const action = String(body.action || '').toLowerCase();
    const result = action === 'scan' ? scan_(body) :
      action === 'auth' ? authorizationStatus_() :
      {ok:false,error:'Unknown action'};
    return output_(result, body.callback);
  } catch (err) {
    return output_({ok:false,error:String(err && err.message || err)});
  }
}

function health_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const vehiclesSheet = ss.getSheetByName(VEHICLE_SHEET);
  const historySheet = ss.getSheetByName(HISTORY_SHEET);
  if (!vehiclesSheet) throw new Error('ไม่พบชีต VEHICLES');
  if (!historySheet) throw new Error('ไม่พบชีต SCAN_HISTORY');
  return {
    ok:true,
    authorized:true,
    spreadsheetId:ss.getId(),
    spreadsheetName:ss.getName(),
    vehiclesSheet:true,
    historySheet:true
  };
}

/* Backward-compatible public aliases */
function getAll() { return getAll_(); }
function health() { return health_(); }

function getAll_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const vs = ss.getSheetByName(VEHICLE_SHEET);
  const hs = ss.getSheetByName(HISTORY_SHEET);
  if (!vs) throw new Error('ไม่พบชีต VEHICLES');
  if (!hs) throw new Error('ไม่พบชีต SCAN_HISTORY');
  return {
    ok:true,
    vehicles: readSheet_(vs),
    history: readSheet_(hs)
  };
}

function readSheet_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length === 0) return [];
  const headers = values[0].map(String);
  return values.slice(1).map(function(row, i) {
    const obj = {_row:i+2};
    headers.forEach(function(h,j){ obj[h] = row[j] == null ? '' : String(row[j]); });
    return obj;
  });
}

function scan_(p) {
  checkKey_(p.key);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(VEHICLE_SHEET);
    const hist = ss.getSheetByName(HISTORY_SHEET);
    if (!sheet) throw new Error('ไม่พบชีต VEHICLES');
    if (!hist) throw new Error('ไม่พบชีต SCAN_HISTORY');

    const lookup = String(p.lookup || '').trim().toLowerCase();
    const result = String(p.result || '').trim();
    const inspector = String(p.inspector || '').trim();
    const area = String(p.area || '').trim();

    if (!lookup) throw new Error('ไม่มี Barcode / QR / Vehicle ID');
    if (!inspector) throw new Error('กรุณาระบุผู้ตรวจ');
    if (['พบเล่ม','ไม่พบเล่ม','ยังไม่ตรวจ'].indexOf(result) < 0)
      throw new Error('ผลการตรวจไม่ถูกต้อง');

    const values = sheet.getDataRange().getDisplayValues();
    if (values.length < 2) throw new Error('VEHICLES ไม่มีข้อมูล');
    const headers = values[0].map(String);
    const idx = {};
    headers.forEach(function(h,i){idx[h]=i;});

    const searchCols = ['Barcode','QRCode','VehicleID','ทะเบียน'];
    let foundRow = -1;
    for (let r=1; r<values.length; r++) {
      for (let k=0; k<searchCols.length; k++) {
        const c = idx[searchCols[k]];
        if (c !== undefined && String(values[r][c] || '').trim().toLowerCase() === lookup) {
          foundRow = r + 1;
          break;
        }
      }
      if (foundRow > 0) break;
    }

    if (foundRow < 0)
      return {ok:false,error:'ไม่พบรถจาก Barcode / QR / Vehicle ID นี้'};

    const vehicle = {};
    headers.forEach(function(h,i){ vehicle[h] = values[foundRow-1][i] || ''; });

    const oldStatus = String(vehicle['สถานะ'] || '').trim();
    if (oldStatus === result && result !== 'ยังไม่ตรวจ') {
      return {ok:true,duplicate:true,message:'รถคันนี้ถูกบันทึกผลนี้ไว้แล้ว',vehicle:vehicle};
    }

    const now = new Date();
    const tz = ss.getSpreadsheetTimeZone() || 'Asia/Bangkok';
    const date = Utilities.formatDate(now,tz,'dd/MM/yyyy');
    const time = Utilities.formatDate(now,tz,'HH:mm:ss');

    const statusCol = idx['สถานะ'];
    if (statusCol === undefined) throw new Error('ไม่พบคอลัมน์ สถานะ');
    const firstCol = statusCol + 1;

    // F:J in the current workbook:
    // สถานะ | ผู้ตรวจล่าสุด | วันที่ตรวจล่าสุด | เวลาตรวจล่าสุด | พื้นที่
    sheet.getRange(foundRow, firstCol, 1, 5)
      .setValues([[result, inspector, date, time, area]]);

    const scanId = 'SCAN-' + Utilities.getUuid();
    const qr = vehicle['QRCode'] || vehicle['Barcode'] || vehicle['VehicleID'] || '';
    const plate = vehicle['ทะเบียน'] || '';
    hist.appendRow([scanId,date,time,qr,plate,area,inspector,result]);

    vehicle['สถานะ']=result;
    vehicle['ผู้ตรวจล่าสุด']=inspector;
    vehicle['วันที่ตรวจล่าสุด']=date;
    vehicle['เวลาตรวจล่าสุด']=time;
    vehicle['พื้นที่']=area;

    return {
      ok:true,
      duplicate:false,
      vehicle:vehicle,
      scan:{
        Scan_ID:scanId, วันที่:date, เวลา:time, QR_Code:qr,
        ทะเบียน:plate, พื้นที่:area, ผู้ตรวจ:inspector, ผลการตรวจ:result
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function checkKey_(key) {
  if (API_KEY && String(key || '') !== API_KEY)
    throw new Error('Invalid API key');
}

function output_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return ContentService.createTextOutput(JSON.stringify({ok:false,error:'Invalid callback'}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
