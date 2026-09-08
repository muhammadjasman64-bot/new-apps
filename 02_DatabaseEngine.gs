/**
 * 02_DatabaseEngine.gs - engine database ringan.
 * Prinsip: header dinamis, batch read, cache hanya untuk data kecil/master.
 */
function dbGetSheet_(name){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sh=ss.getSheetByName(name);
  if(!sh) throw new Error('Sheet "'+name+'" tidak ditemukan.');
  return sh;
}
function dbNormalizeHeader_(v){
  return String(v==null?'':v).trim().toLowerCase().replace(/[\s\-\/]+/g,'_').replace(/[^\w]/g,'');
}
function dbGetHeaderMap_(sheet){
  const lastCol=sheet.getLastColumn(); if(!lastCol) return {};
  const headers=sheet.getRange(1,1,1,lastCol).getDisplayValues()[0], map={};
  headers.forEach((h,i)=>{const k=dbNormalizeHeader_(h);if(k)map[k]=i;});
  return map;
}
function dbFindColumn_(map,aliases){for(const a of aliases){const k=dbNormalizeHeader_(a);if(Object.prototype.hasOwnProperty.call(map,k))return map[k];}return -1;}

/** Membaca hanya area yang benar-benar berisi data. */
function dbRead_(sheetName){
  const sh=dbGetSheet_(sheetName), lastRow=sh.getLastRow(), lastCol=sh.getLastColumn();
  if(!lastRow||!lastCol) return {sheet:sh,headers:[],rows:[],map:{}};
  const values=sh.getRange(1,1,lastRow,lastCol).getValues();
  if(values.length<2) return {sheet:sh,headers:values[0]||[],rows:[],map:dbGetHeaderMap_(sh)};
  return {sheet:sh,headers:values[0],rows:values.slice(1),map:dbGetHeaderMap_(sh)};
}

function dbHeaders_(sheetName){
  const sh=dbGetSheet_(sheetName);
  const lastCol=sh.getLastColumn();
  return lastCol?sh.getRange(1,1,1,lastCol).getValues()[0].map(v=>String(v==null?'':v).trim()):[];
}
function dbReadObjects_(sheetName){const d=dbRead_(sheetName);return d.rows.map(row=>{const o={};d.headers.forEach((h,i)=>o[String(h||'').trim()]=row[i]);return o;});}
function dbColumnValue_(row,map,aliases){const i=dbFindColumn_(map,aliases);return i>=0?row[i]:'';}
function dbDate_(v){
  if(v instanceof Date&&!isNaN(v.getTime()))return new Date(v.getTime());
  const s=String(v||'').trim();if(!s)return null;
  let m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if(m){const d=Number(m[1]),mo=Number(m[2])-1,y=Number(m[3]),x=new Date(y,mo,d);if(x.getFullYear()===y&&x.getMonth()===mo&&x.getDate()===d)return x;}
  const x=new Date(s);return isNaN(x.getTime())?null:x;
}
function dbBetween_(v,start,end){const d=dbDate_(v);if(!d)return false;d.setHours(0,0,0,0);const a=new Date(start),b=new Date(end);a.setHours(0,0,0,0);b.setHours(23,59,59,999);return d>=a&&d<=b;}

function syncDatabase_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const wanted=['CONFIG','DATA_SISWA','ABSENSI','MASTER_PELANGGARAN','PELANGGARAN','MASTER_PENGHARGAAN','PENGHARGAAN','TINDAKAN','DOKUMEN'];
  const result=wanted.map(name=>{const sh=ss.getSheetByName(name);if(!sh)return{sheet:name,status:'MISSING',rows:0,columns:0,headers:[]};const lastRow=Math.max(sh.getLastRow()-1,0),lastCol=sh.getLastColumn();return{sheet:name,status:'OK',rows:lastRow,columns:lastCol,headers:lastCol?sh.getRange(1,1,1,lastCol).getDisplayValues()[0]:[]};});
  CacheService.getScriptCache().put('DB_LAST_SYNC',JSON.stringify({at:new Date().toISOString(),sheets:result}),21600);
  return {ok:true,at:new Date().toISOString(),sheets:result};
}
function getDatabaseStatus_(){const c=CacheService.getScriptCache().get('DB_LAST_SYNC');return c?JSON.parse(c):syncDatabase_();}
