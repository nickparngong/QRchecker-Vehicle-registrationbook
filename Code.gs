const SPREADSHEET_ID='1rfd7D_NraAkv-0zhnKUfWYtGNg0OPsDH0B7QDcsUZ3A';
const VEHICLE_SHEET='VEHICLES', HISTORY_SHEET='SCAN_HISTORY', CYCLE_SHEET='SCAN_CYCLES';
function doGet(e){
 const p=(e&&e.parameter)||{},a=String(p.action||'health').toLowerCase();
 try{
  let r;
  if(a==='health')r=health_();
  else if(a==='getall')r=getAll_();
  else if(a==='scan')r=scan_(p);
  else if(a==='savevehicle')r=saveVehicle_(p);
  else if(a==='togglevehicle')r=toggleVehicle_(p);
  else if(a==='bulkimport')r=bulkImport_(p);
  else if(a==='resetmonth')r=resetMonth_(p);
  else r={ok:false,error:'Unknown action: '+a};
  return output_(r,p.callback);
 }catch(err){return output_({ok:false,error:String(err&&err.message||err)},p.callback);}
}
function ss_(){return SpreadsheetApp.openById(SPREADSHEET_ID);}
function health_(){const s=ss_();return {ok:true,spreadsheetId:s.getId(),spreadsheetName:s.getName(),time:new Date().toISOString()};}
function monthKey_(s){return Utilities.formatDate(new Date(),s.getSpreadsheetTimeZone()||'Asia/Bangkok','yyyy-MM');}
function ensureCycleSheet_(s){
 let sh=s.getSheetByName(CYCLE_SHEET);
 if(!sh){sh=s.insertSheet(CYCLE_SHEET);sh.appendRow(['MonthKey','Round','ResetAt','ResetBy']);}
 if(sh.getLastRow()===0)sh.appendRow(['MonthKey','Round','ResetAt','ResetBy']);
 return sh;
}
function currentCycle_(s){
 const key=monthKey_(s),sh=ensureCycleSheet_(s),v=sh.getDataRange().getDisplayValues();
 let round=1;
 for(let i=1;i<v.length;i++)if(String(v[i][0])===key)round=Math.max(round,Number(v[i][1]||1));
 return {monthKey:key,round:round};
}
function getAll_(){
 const s=ss_(),c=currentCycle_(s);
 return {ok:true,vehicles:readSheet_(s.getSheetByName(VEHICLE_SHEET)),history:readSheet_(s.getSheetByName(HISTORY_SHEET)),monthKey:c.monthKey,round:c.round};
}
function readSheet_(sh){
 if(!sh)throw new Error('ไม่พบชีตที่ต้องการอ่าน');
 const v=sh.getDataRange().getDisplayValues();if(!v.length)return [];
 const h=v[0].map(x=>String(x).trim());
 return v.slice(1).filter(r=>r.some(x=>String(x).trim()!=='')).map((r,i)=>{const o={_row:i+2};h.forEach((k,j)=>o[k]=r[j]===undefined?'':r[j]);return o;});
}
function ensureHistoryHeaders_(sh){
 if(sh.getLastRow()===0)sh.appendRow(['Scan_ID','วันที่','เวลา','QR_Code','ทะเบียน','พื้นที่','ผู้ตรวจ','ผลการตรวจ']);
 let h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(x=>String(x).trim());
 ['MonthKey','Round'].forEach(k=>{if(h.indexOf(k)<0){sh.getRange(1,h.length+1).setValue(k);h.push(k);}});
 return h;
}
function appendHistory_(sh,o){
 const h=ensureHistoryHeaders_(sh),row=h.map(k=>o[k]===undefined?'':o[k]);
 sh.appendRow(row);
}
function norm_(v){return String(v||'').replace(/[\r\n\t\uFEFF\u200B\u200C\u200D]/g,'').replace(/\s+/g,'').trim().toUpperCase();}
function scan_(p){
 const lock=LockService.getScriptLock();lock.waitLock(25000);
 try{
  const s=ss_(),vsh=s.getSheetByName(VEHICLE_SHEET),hsh=s.getSheetByName(HISTORY_SHEET);
  if(!vsh||!hsh)throw new Error('ไม่พบชีต VEHICLES หรือ SCAN_HISTORY');
  const code=norm_(p.lookup),inspector=String(p.inspector||'').trim(),area=String(p.area||'').trim();
  if(!code)throw new Error('ไม่พบ Barcode');if(!inspector)throw new Error('ไม่พบชื่อผู้ตรวจ');
  const c=currentCycle_(s),vh=vsh.getDataRange().getDisplayValues(),heads=vh[0].map(x=>String(x).trim()),ix={};heads.forEach((x,i)=>ix[x]=i);
  let row=-1;const searchCols=['Barcode','QRCode','VehicleID'];
  for(let r=1;r<vh.length;r++){if(searchCols.some(k=>ix[k]!==undefined&&norm_(vh[r][ix[k]])===code)){row=r+1;break;}}
  const tz=s.getSpreadsheetTimeZone()||'Asia/Bangkok',now=new Date(),date=Utilities.formatDate(now,tz,'dd/MM/yyyy'),time=Utilities.formatDate(now,tz,'HH:mm:ss');
  let plate='',qr=code,found=row>0;
  if(found){plate=ix['ทะเบียน']!==undefined?vh[row-1][ix['ทะเบียน']]||'':'';qr=(ix['QRCode']!==undefined&&vh[row-1][ix['QRCode']])|| (ix['Barcode']!==undefined&&vh[row-1][ix['Barcode']])||code;}
  const hheads=ensureHistoryHeaders_(hsh),hv=hsh.getDataRange().getDisplayValues(),hi={};hheads.forEach((x,i)=>hi[x]=i);
  let duplicate=false;
  if(found){
   for(let r=1;r<hv.length;r++){
    if(norm_(hv[r][hi['QR_Code']])===norm_(qr)&&String(hv[r][hi['MonthKey']]||'')===c.monthKey&&Number(hv[r][hi['Round']]||1)===c.round&&String(hv[r][hi['ผลการตรวจ']]||'')==='พบเล่ม'){duplicate=true;break;}
   }
  }
  const result=!found?'ไม่พบข้อมูล':duplicate?'ตรวจซ้ำ':'พบเล่ม';
  appendHistory_(hsh,{'Scan_ID':'SCAN-'+Utilities.getUuid(),'วันที่':date,'เวลา':time,'QR_Code':qr,'ทะเบียน':plate,'พื้นที่':area,'ผู้ตรวจ':inspector,'ผลการตรวจ':result,'MonthKey':c.monthKey,'Round':c.round});
  if(found&&!duplicate){
   const updates={'ผู้ตรวจล่าสุด':inspector,'วันที่ตรวจล่าสุด':date,'เวลาตรวจล่าสุด':time,'พื้นที่':area};
   Object.keys(updates).forEach(k=>{if(ix[k]!==undefined)vsh.getRange(row,ix[k]+1).setValue(updates[k]);});
  }
  return {ok:true,duplicate:duplicate,found:found,monthKey:c.monthKey,round:c.round,vehicle:found?{'ทะเบียน':plate,'Barcode':qr}:null,result:result};
 }finally{lock.releaseLock();}
}
function resetMonth_(p){
 const lock=LockService.getScriptLock();lock.waitLock(25000);
 try{
  const s=ss_(),sh=ensureCycleSheet_(s),key=monthKey_(s),v=sh.getDataRange().getDisplayValues();let round=0;
  for(let i=1;i<v.length;i++)if(String(v[i][0])===key)round=Math.max(round,Number(v[i][1]||1));
  round++;
  sh.appendRow([key,round,new Date(),String(p.inspector||'ADMIN')]);
  return {ok:true,monthKey:key,round:round};
 }finally{lock.releaseLock();}
}
function saveVehicle_(p){
 const barcode=String(p.barcode||'').trim(),plate=String(p.plate||'').trim(),month=String(p.month||'').trim();
 if(!barcode||!plate)throw new Error('กรุณาระบุ Barcode และทะเบียน');
 const lock=LockService.getScriptLock();lock.waitLock(25000);
 try{
  const s=ss_(),sh=s.getSheetByName(VEHICLE_SHEET);if(!sh)throw new Error('ไม่พบชีต VEHICLES');
  const vals=sh.getDataRange().getDisplayValues(),heads=vals[0].map(x=>String(x).trim()),ix={};heads.forEach((x,i)=>ix[x]=i);
  const bcCol=ix['Barcode']!==undefined?ix['Barcode']:ix['QRCode'];if(bcCol===undefined)throw new Error('ไม่พบคอลัมน์ Barcode/QRCode');
  let row=-1;for(let r=1;r<vals.length;r++)if(norm_(vals[r][bcCol])===norm_(barcode)){row=r+1;break;}
  let data=row>0?vals[row-1].slice():new Array(heads.length).fill('');
  const set=(k,v)=>{if(ix[k]!==undefined)data[ix[k]]=v;};
  set('Barcode',barcode);set('ทะเบียน',plate);set('เดือนจดทะเบียน',month);
  if(row<0){set('สถานะ','Active');sh.appendRow(data);}else sh.getRange(row,1,1,heads.length).setValues([data]);
  return {ok:true,updated:row>0};
 }finally{lock.releaseLock();}
}
function toggleVehicle_(p){
 const s=ss_(),sh=s.getSheetByName(VEHICLE_SHEET),v=sh.getDataRange().getDisplayValues(),h=v[0].map(x=>String(x).trim()),ix={};h.forEach((x,i)=>ix[x]=i);
 const bc=norm_(p.barcode),bcCol=ix.Barcode!==undefined?ix.Barcode:ix.QRCode;if(bcCol===undefined||ix['สถานะ']===undefined)throw new Error('ไม่พบคอลัมน์ที่จำเป็น');
 for(let r=1;r<v.length;r++)if(norm_(v[r][bcCol])===bc){sh.getRange(r+1,ix['สถานะ']+1).setValue(String(p.status)==='Active'?'Active':'Inactive');return {ok:true};}
 throw new Error('ไม่พบ Barcode');
}
function bulkImport_(p){
 const rows=JSON.parse(p.rows||'[]');if(!Array.isArray(rows))throw new Error('รูปแบบข้อมูลไม่ถูกต้อง');
 let added=0,updated=0;rows.forEach(x=>{const r=saveVehicle_({barcode:x.barcode,plate:x.plate,month:x.month});if(r.updated)updated++;else added++;});
 return {ok:true,added:added,updated:updated};
}
function output_(o,cb){const s=JSON.stringify(o);if(cb){if(!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(cb))return ContentService.createTextOutput('Invalid callback');return ContentService.createTextOutput(cb+'('+s+');').setMimeType(ContentService.MimeType.JAVASCRIPT);}return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);}
