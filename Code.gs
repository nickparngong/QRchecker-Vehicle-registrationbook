const SPREADSHEET_ID = '1rfd7D_NraAkv-0zhnKUfWYtGNg0OPsDH0B7QDcsUZ3A';
const VEHICLE_SHEET = 'VEHICLES';
const HISTORY_SHEET = 'SCAN_HISTORY';
const APP_KEY = 'CHANGE_THIS_KEY'; // optional shared key; change before deploy

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = p.callback || '';
  let result;
  try {
    if (APP_KEY && APP_KEY !== 'CHANGE_THIS_KEY' && p.key !== APP_KEY) {
      throw new Error('Unauthorized');
    }
    const action = p.action || 'getAll';
    if (action === 'getAll') result = getAll_();
    else if (action === 'scan') result = scan_(p);
    else if (action === 'health') result = { ok: true, spreadsheetId: SPREADSHEET_ID };
    else throw new Error('Unknown action: ' + action);
  } catch (err) {
    result = { ok: false, error: String(err && err.message || err) };
  }
  const body = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function normalize_(v) {
  return String(v == null ? '' : v).trim().toLowerCase();
}

function headerIndex_(headers, aliases) {
  const h = headers.map(x => String(x || '').trim());
  for (const a of aliases) {
    const i = h.indexOf(a);
    if (i >= 0) return i;
  }
  return -1;
}

function getMap_(headers) {
  return {
    id: headerIndex_(headers, ['VehicleID','Vehicle_ID','Vehicle ID']),
    qr: headerIndex_(headers, ['QRCode','QR_Code','QR Code']),
    plate: headerIndex_(headers, ['ทะเบียน','ทะเบียนรถ','License Plate']),
    month: headerIndex_(headers, ['เดือนจดทะเบียน','เดือน']),
    province: headerIndex_(headers, ['จังหวัด']),
    status: headerIndex_(headers, ['สถานะ']),
    inspector: headerIndex_(headers, ['ผู้ตรวจล่าสุด','ผู้ตรวจ']),
    date: headerIndex_(headers, ['วันที่ตรวจล่าสุด','วันที่']),
    time: headerIndex_(headers, ['เวลาตรวจล่าสุด','เวลา']),
    area: headerIndex_(headers, ['พื้นที่']),
    barcode: headerIndex_(headers, ['Barcode','BARCODE'])
  };
}

function serializable_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  return v == null ? '' : v;
}

function readSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) return { headers: [], rows: [] };
  const headers = values[0].map(serializable_);
  const rows = values.slice(1).map(r => r.map(serializable_));
  return { headers, rows };
}

function getAll_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const vs = ss.getSheetByName(VEHICLE_SHEET);
  const hs = ss.getSheetByName(HISTORY_SHEET);
  if (!vs) throw new Error('ไม่พบชีต ' + VEHICLE_SHEET);
  const v = readSheet_(vs);
  const h = hs ? readSheet_(hs) : { headers: [], rows: [] };
  return { ok: true, vehicles: { headers: v.headers, rows: v.rows }, history: { headers: h.headers, rows: h.rows } };
}

function scan_(p) {
  const code = String(p.code || '').trim();
  const inspector = String(p.inspector || '').trim();
  const requestedStatus = p.status === 'ไม่พบเล่ม' ? 'ไม่พบเล่ม' : 'พบเล่ม';
  if (!code) throw new Error('ไม่มี Barcode');
  if (!inspector) throw new Error('กรุณาระบุชื่อผู้ตรวจ');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const vs = ss.getSheetByName(VEHICLE_SHEET);
    const hs = ss.getSheetByName(HISTORY_SHEET);
    if (!vs) throw new Error('ไม่พบชีต ' + VEHICLE_SHEET);
    const data = vs.getDataRange().getValues();
    if (data.length < 2) throw new Error('VEHICLES ไม่มีข้อมูล');
    const headers = data[0].map(serializable_);
    const m = getMap_(headers);
    const searchable = [m.barcode, m.qr, m.id, m.plate].filter(i => i >= 0);
    let rowNumber = -1;
    for (let r = 1; r < data.length; r++) {
      const found = searchable.some(c => normalize_(data[r][c]) === normalize_(code));
      if (found) { rowNumber = r + 1; break; }
    }
    if (rowNumber < 0) throw new Error('ไม่พบ Barcode/Vehicle ใน Google Sheet: ' + code);

    const row = vs.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    const currentStatus = m.status >= 0 ? String(row[m.status] || '').trim() : '';
    if (currentStatus && currentStatus !== 'ยังไม่ตรวจ') {
      return { ok: false, duplicate: true, message: 'รถคันนี้ตรวจแล้ว', row: row.map(serializable_), headers };
    }

    const now = new Date();
    const dateText = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const timeText = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    if (m.status >= 0) vs.getRange(rowNumber, m.status + 1).setValue(requestedStatus);
    if (m.inspector >= 0) vs.getRange(rowNumber, m.inspector + 1).setValue(inspector);
    if (m.date >= 0) vs.getRange(rowNumber, m.date + 1).setValue(dateText);
    if (m.time >= 0) vs.getRange(rowNumber, m.time + 1).setValue(timeText);

    let scanId = 'SC' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMddHHmmss') + Math.floor(Math.random()*1000);
    if (hs) {
      const hv = hs.getDataRange().getValues();
      const hh = hv.length ? hv[0].map(serializable_) : [];
      const hrow = new Array(hh.length).fill('');
      const hm = getMapHistory_(hh);
      setIf_(hrow, hm.scanId, scanId);
      setIf_(hrow, hm.date, dateText);
      setIf_(hrow, hm.time, timeText);
      setIf_(hrow, hm.code, code);
      setIf_(hrow, hm.plate, m.plate >= 0 ? row[m.plate] : '');
      setIf_(hrow, hm.area, m.area >= 0 ? row[m.area] : '');
      setIf_(hrow, hm.inspector, inspector);
      setIf_(hrow, hm.result, requestedStatus);
      hs.appendRow(hrow);
    }

    const updated = vs.getRange(rowNumber, 1, 1, headers.length).getValues()[0].map(serializable_);
    return { ok: true, duplicate: false, message: 'บันทึกสำเร็จ', headers, row: updated, scanId, date: dateText, time: timeText };
  } finally {
    lock.releaseLock();
  }
}

function setIf_(row, index, value) { if (index >= 0) row[index] = serializable_(value); }
function getMapHistory_(headers) {
  return {
    scanId: headerIndex_(headers, ['Scan_ID','Scan ID','ScanID']),
    date: headerIndex_(headers, ['วันที่','Date']),
    time: headerIndex_(headers, ['เวลา','Time']),
    code: headerIndex_(headers, ['QR_Code','QRCode','Barcode','BARCODE','QR Code']),
    plate: headerIndex_(headers, ['ทะเบียน','ทะเบียนรถ','License Plate']),
    area: headerIndex_(headers, ['พื้นที่','Area']),
    inspector: headerIndex_(headers, ['ผู้ตรวจ','ผู้ตรวจล่าสุด','Inspector']),
    result: headerIndex_(headers, ['ผลการตรวจ','สถานะ','Result'])
  };
}
