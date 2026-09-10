const SPREADSHEET_ID = '1rfd7D_NraAkv-0zhnKUfWYtGNg0OPsDH0B7QDcsUZ3A';
const VEHICLE_SHEET = 'VEHICLES';
const HISTORY_SHEET = 'SCAN_HISTORY';

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    const action = String(p.action || 'getAll');
    let result;
    if (action === 'health') result = health_();
    else if (action === 'getAll') result = getAll_();
    else if (action === 'scan') result = scan_(p);
    else throw new Error('Unknown action: ' + action);
    return json_(result);
  } catch (err) {
    return json_({ ok:false, error:String(err && err.message ? err.message : err) });
  }
}

function health_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return {
    ok:true,
    message:'Google Sheets connection OK',
    spreadsheetId:SPREADSHEET_ID,
    vehicleSheet:VEHICLE_SHEET,
    historySheet:HISTORY_SHEET,
    vehicleSheetExists:!!ss.getSheetByName(VEHICLE_SHEET),
    historySheetExists:!!ss.getSheetByName(HISTORY_SHEET)
  };
}

function getAll_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const vs = ss.getSheetByName(VEHICLE_SHEET);
  const hs = ss.getSheetByName(HISTORY_SHEET);
  if (!vs) throw new Error('ไม่พบชีต ' + VEHICLE_SHEET);
  const v = readSheet_(vs);
  const h = hs ? readSheet_(hs) : {headers:[],rows:[]};
  return {ok:true, vehicles:v, history:h};
}

function readSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) return {headers:[],rows:[]};
  return {
    headers: values[0].map(clean_),
    rows: values.slice(1).map(r => r.map(clean_))
  };
}

function scan_(p) {
  const code = String(p.code || '').trim();
  const inspector = String(p.inspector || '').trim();
  const status = p.status === 'ไม่พบเล่ม' ? 'ไม่พบเล่ม' : 'พบเล่ม';
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
    const headers = data[0].map(clean_);
    const m = vehicleMap_(headers);
    const searchable = [m.barcode,m.qr,m.id,m.plate].filter(i => i >= 0);

    let rowNumber = -1;
    for (let r=1; r<data.length; r++) {
      if (searchable.some(c => normalize_(data[r][c]) === normalize_(code))) {
        rowNumber = r + 1;
        break;
      }
    }
    if (rowNumber < 0) throw new Error('ไม่พบ Barcode/Vehicle ใน Google Sheet: ' + code);

    const row = vs.getRange(rowNumber,1,1,headers.length).getValues()[0];
    const currentStatus = m.status >= 0 ? String(row[m.status] || '').trim() : '';
    if (currentStatus && currentStatus !== 'ยังไม่ตรวจ') {
      return {ok:false,duplicate:true,message:'รถคันนี้ตรวจแล้ว',headers:headers,row:row.map(clean_)};
    }

    const now = new Date();
    const tz = Session.getScriptTimeZone();
    const dateText = Utilities.formatDate(now,tz,'dd/MM/yyyy');
    const timeText = Utilities.formatDate(now,tz,'HH:mm:ss');

    setCell_(vs,rowNumber,m.status,status);
    setCell_(vs,rowNumber,m.inspector,inspector);
    setCell_(vs,rowNumber,m.date,dateText);
    setCell_(vs,rowNumber,m.time,timeText);

    const scanId = 'SC' + Utilities.formatDate(now,tz,'yyyyMMddHHmmss') + Math.floor(Math.random()*1000);
    if (hs) appendHistory_(hs,headers,row,m,code,inspector,status,dateText,timeText,scanId);

    const updated = vs.getRange(rowNumber,1,1,headers.length).getValues()[0].map(clean_);
    return {ok:true,duplicate:false,message:'บันทึกสำเร็จ',headers:headers,row:updated,scanId:scanId,date:dateText,time:timeText};
  } finally {
    lock.releaseLock();
  }
}

function appendHistory_(hs,headers,row,m,code,inspector,status,dateText,timeText,scanId) {
  const values = hs.getDataRange().getValues();
  const hh = values.length ? values[0].map(clean_) : [];
  const hrow = new Array(hh.length).fill('');
  const hm = historyMap_(hh);
  put_(hrow,hm.scanId,scanId); put_(hrow,hm.date,dateText); put_(hrow,hm.time,timeText);
  put_(hrow,hm.code,code); put_(hrow,hm.plate,m.plate>=0?row[m.plate]:'');
  put_(hrow,hm.area,m.area>=0?row[m.area]:''); put_(hrow,hm.inspector,inspector); put_(hrow,hm.result,status);
  hs.appendRow(hrow);
}

function vehicleMap_(h) {
  return {
    id:idx_(h,['VehicleID','Vehicle_ID','Vehicle ID']), qr:idx_(h,['QRCode','QR_Code','QR Code']),
    plate:idx_(h,['ทะเบียน','ทะเบียนรถ','License Plate']), month:idx_(h,['เดือนจดทะเบียน','เดือน']),
    province:idx_(h,['จังหวัด']), status:idx_(h,['สถานะ']), inspector:idx_(h,['ผู้ตรวจล่าสุด','ผู้ตรวจ']),
    date:idx_(h,['วันที่ตรวจล่าสุด','วันที่']), time:idx_(h,['เวลาตรวจล่าสุด','เวลา']),
    area:idx_(h,['พื้นที่']), barcode:idx_(h,['Barcode','BARCODE'])
  };
}

function historyMap_(h) {
  return {
    scanId:idx_(h,['Scan_ID','Scan ID','ScanID']), date:idx_(h,['วันที่','Date']), time:idx_(h,['เวลา','Time']),
    code:idx_(h,['QR_Code','QRCode','Barcode','BARCODE','QR Code']), plate:idx_(h,['ทะเบียน','ทะเบียนรถ','License Plate']),
    area:idx_(h,['พื้นที่','Area']), inspector:idx_(h,['ผู้ตรวจ','ผู้ตรวจล่าสุด','Inspector']), result:idx_(h,['ผลการตรวจ','สถานะ','Result'])
  };
}

function idx_(headers,aliases) {
  const h = headers.map(x => String(x || '').trim());
  for (const a of aliases) { const i=h.indexOf(a); if(i>=0) return i; }
  return -1;
}
function setCell_(sheet,row,col,value) { if(col>=0) sheet.getRange(row,col+1).setValue(value); }
function put_(row,i,value) { if(i>=0) row[i]=clean_(value); }
function normalize_(v) { return String(v == null ? '' : v).trim().toLowerCase(); }
function clean_(v) {
  if (v instanceof Date) return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
  return v == null ? '' : v;
}
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
