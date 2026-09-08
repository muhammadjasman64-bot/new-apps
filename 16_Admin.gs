/** 16_Admin.gs - Administrasi Sistem V32
 * Backup, export CSV, audit log, validasi & status database.
 */
function getAdminSummary_(){
  const ss=getSS_();
  if(!ss) throw new Error('Spreadsheet aktif tidak ditemukan. Pastikan Web App terikat pada spreadsheet yang benar.');
  const sheets=ss.getSheets().map(function(sh){
    return {name:sh.getName(),rows:Math.max(0,sh.getLastRow()-1),columns:sh.getLastColumn()};
  });
  const name=String(ss.getName()||'Spreadsheet');
  const url=String(ss.getUrl()||'');
  const dbVersion=typeof getDbVersion_==='function'?String(getDbVersion_()||'1'):'1';
  return {success:true,spreadsheetId:ss.getId(),name:name,url:url,updated:new Date(),sheets:sheets,dbVersion:dbVersion};
}

function backupDatabase_(){
  const ss=getSS_();
  const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss');
  const folder=getOrCreateFolder_();
  const file=DriveApp.getFileById(ss.getId()).makeCopy('BACKUP_JURNAL_WALI_KELAS_'+stamp,folder);
  logActivity_('BACKUP_DATABASE','Backup dibuat: '+file.getName());
  return {success:true,name:file.getName(),id:file.getId(),url:file.getUrl()};
}
function exportDatabaseCsv_(){
  const ss=getSS_(), blobs=[];
  ss.getSheets().forEach(sh=>{
    const values=sh.getDataRange().getDisplayValues();
    const csv=values.map(row=>row.map(csvEscape_).join(',')).join('\r\n');
    blobs.push(Utilities.newBlob(csv,'text/csv',safeFileName_(sh.getName())+'.csv'));
  });
  const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss');
  const zip=Utilities.zip(blobs,'EXPORT_DATABASE_'+stamp+'.zip');
  const file=getOrCreateFolder_().createFile(zip);
  logActivity_('EXPORT_DATABASE','Export CSV ZIP: '+file.getName());
  return {success:true,name:file.getName(),id:file.getId(),url:file.getUrl(),downloadUrl:'https://drive.google.com/uc?export=download&id='+file.getId()};
}
function csvEscape_(v){const s=String(v==null?'':v);return '"'+s.replace(/"/g,'""')+'"';}
function safeFileName_(s){return String(s).replace(/[\\\/:*?"<>|]/g,'_').slice(0,80);}
function getAuditLog_(limit){
  const sh=getSheet_(APP.SHEETS.LOG), n=Math.max(1,Math.min(200,Number(limit)||50));
  const last=sh.getLastRow(); if(last<2)return [];
  const start=Math.max(2,last-n+1); const rows=sh.getRange(start,1,last-start+1,4).getDisplayValues();
  return rows.reverse().map(r=>({timestamp:r[0],user:r[1],aktivitas:r[2],detail:r[3]}));
}
function validateDatabaseDetailed_(){
  const base=validateDatabase_();
  const ss=getSS_(), required={
    DATA_SISWA:['ID_Siswa','NISN','Nama_Siswa','Kelas','Status'],
    ABSENSI:['ID_Absensi','Tanggal','NISN','Nama_Siswa','Kelas','Status'],
    PELANGGARAN:['ID_Pelanggaran','Tanggal','NISN','Poin'],
    PENGHARGAAN:['ID_Penghargaan','Tanggal','NISN','Poin']
  };
  const checks=[];
  Object.keys(required).forEach(name=>{
    const sh=ss.getSheetByName(name); const headers=sh?sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getDisplayValues()[0]:[];
    const missing=required[name].filter(h=>!headers.includes(h));
    checks.push({sheet:name,exists:!!sh,missing:missing,ok:!!sh&&!missing.length});
  });
  return {success:true,base:base,checks:checks,ok:checks.every(x=>x.ok)};
}
function clearApplicationCache_(){clearAppCache_();return {success:true,message:'Cache aplikasi dibersihkan.'};}
