/** V44.2 - Role Access, Auto User Generator & Database Protection */
const ROLE_ACCESS_ = {
  admin: {menus:['dash','siswa','abs','pel','rew','rekap','dok','bina','admin'], perms:['*']},
  wali_kelas: {menus:['dash','siswa','abs','pel','rew','rekap','dok','bina'], perms:['read_dashboard','read_students','write_students','write_attendance','read_attendance','read_rules','write_violation','read_violation','read_reward','write_reward','read_recap','read_document','read_coaching','resolve_warning','read_risk']},
  guru: {menus:['dash','siswa','abs','pel','rew','rekap'], perms:['read_dashboard','read_students','write_attendance','read_attendance','read_rules','write_violation','read_violation','read_reward','write_reward','read_recap','read_risk']},
  bk: {menus:['dash','siswa','pel','rekap','dok','bina'], perms:['read_dashboard','read_students','read_attendance','read_rules','write_violation','read_violation','read_recap','read_document','read_coaching','resolve_warning','read_risk']},
  kepala_sekolah: {menus:['dash','siswa','rekap','dok','bina'], perms:['read_dashboard','read_students','read_attendance','read_rules','read_recap','read_document','read_coaching','read_risk']}
};
const ROLE_LABELS_={admin:'Administrator',wali_kelas:'Wali Kelas',guru:'Guru',bk:'Guru BK',kepala_sekolah:'Kepala Sekolah'};
function getRoleAccess_(role){return ROLE_ACCESS_[String(role||'').toLowerCase()]||{menus:[],perms:[]};}
function requirePermission_(token,perm){const s=requireAuth_(token,AUTH.ROLES);if(s.role==='admin')return s;const a=getRoleAccess_(s.role);if(!a.perms.includes(perm))throw new Error('Akses ditolak. Role '+(ROLE_LABELS_[s.role]||s.role)+' tidak memiliki izin untuk tindakan ini.');return s;}
function getRoleProfile_(token){const s=requireAuth_(token,AUTH.ROLES),a=getRoleAccess_(s.role);return {success:true,role:s.role,label:ROLE_LABELS_[s.role]||s.role,menus:a.menus,permissions:a.perms};}
function slugUser_(name){let s=String(name||'user').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]+/g,'.').replace(/^\.+|\.+$/g,'');return s||'user';}
function randomPassword_(){const chars='ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#';let out='';for(let i=0;i<10;i++)out+=chars.charAt(Math.floor(Math.random()*chars.length));return out;}
function createAutoUser_(token,data){requireAuth_(token,['admin']);data=data||{};const role=String(data.role||'').toLowerCase();if(!['wali_kelas','guru','bk','kepala_sekolah'].includes(role))throw new Error('Role user otomatis harus Wali Kelas, Guru, BK, atau Kepala Sekolah.');const nama=String(data.nama||'').trim();if(!nama)throw new Error('Nama lengkap wajib diisi.');const kelas=String(data.kelas||'').trim();const nisn=String(data.nisn||'').trim();const prefix={wali_kelas:'wali',guru:'guru',bk:'bk',kepala_sekolah:'kepala'}[role];let username=prefix+'.'+slugUser_(nama),base=username,n=1;while(findUserByUsername_(username))username=base+'.'+(++n);const password=randomPassword_();const sh=getSheet_('USERS'),headers=authDbHeaders_('USERS');const now=new Date();const obj={ID_User:generateID_('USR'),Username:username,Password_Hash:hashPassword_(password),Nama_Lengkap:nama,Role:role,NISN:nisn,Kelas:kelas,Status:'aktif',Created_At:now,Updated_At:now};sh.getRange(sh.getLastRow()+1,1,1,headers.length).setValues([headers.map(h=>obj[h]!==undefined?obj[h]:'')]);clearAppCache_();auditSecurity_('AUTO_USER_DIBUAT','username='+username+';role='+role);return {success:true,username:username,password:password,role:role,label:ROLE_LABELS_[role],nama:nama};}
function getRoleAccessLinks_(token){requireAuth_(token,['admin']);const base=String(getWebAppUrl_()||'').replace(/\?.*$/,'');if(!base)throw new Error('URL Web App belum tersedia. Deploy sebagai Web App terlebih dahulu.');return Object.keys(ROLE_LABELS_).map(role=>({role:role,label:ROLE_LABELS_[role],url:base+'?access='+encodeURIComponent(role)}));}
function getWebAppUrl_(){const p=PropertiesService.getScriptProperties();return String(p.getProperty('WEB_APP_URL')||ScriptApp.getService().getUrl()||'');}
function protectDatabaseSheets_(){const ss=getSS_();if(!ss)throw new Error('Spreadsheet tidak ditemukan.');const names=ss.getSheets().map(sh=>sh.getName());let protectedCount=0;ss.getSheets().forEach(sh=>{try{let ps=sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);let p=ps.length?ps[0]:sh.protect();p.setDescription('JURNAL WALI KELAS — DATABASE PROTECTED. Akses melalui Web App.');p.setWarningOnly(false);const me=Session.getEffectiveUser().getEmail();const editors=p.getEditors();if(editors.length)p.removeEditors(editors);if(me)p.addEditor(me);protectedCount++;}catch(e){}});return {success:true,protectedSheets:protectedCount,sheets:names};}
function getDatabaseProtectionStatus_(token){requireAuth_(token,['admin']);const ss=getSS_();return {success:true,sheets:ss.getSheets().map(sh=>({name:sh.getName(),protected:sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).length>0}))};}
function ensureDatabaseSheetsBlank_(){const ss=getSS_();const exclude=[APP.SHEETS.CONFIG,APP.SHEETS.MASTER_PELANGGARAN,APP.SHEETS.MASTER_PENGHARGAAN,APP.SHEETS.MASTER_PEMBINAAN,'USERS'];ss.getSheets().forEach(sh=>{if(exclude.includes(sh.getName()))return;if(sh.getLastRow()<=1)return;/* intentionally preserve historical data; setup never wipes operational history */});return true;}

// V44.3 - Wali Kelas: import siswa terbatas pada kelas yang menjadi tanggung jawabnya.
function getScopedStudentRows_(token, includeInactive){
  const s=requireAuth_(token,AUTH.ROLES);
  const rows=includeInactive?getSiswaSemua_():getSiswaAktif_();
  if(s.role!=='wali_kelas') return rows;
  const kelas=String(s.kelas||'').trim();
  if(!kelas) throw new Error('Akun Wali Kelas belum memiliki Kelas pada Manajemen Pengguna. Admin harus mengisi Kelas terlebih dahulu.');
  return rows.filter(function(x){return String(x.kelas||'').trim()===kelas;});
}

function importSiswaRowsScoped_(token, rows){
  const s=requirePermission_(token,'write_students');if(s.role==='admin')return importSiswaRows_(rows);if(s.role!=='wali_kelas')throw new Error('Hanya Admin atau Wali Kelas yang dapat mengimpor data siswa.');
  const kelas=String(s.kelas||'').trim();if(!kelas)throw new Error('Akun Wali Kelas belum memiliki Kelas.');if(!Array.isArray(rows)||!rows.length)throw new Error('Tidak ada data untuk diimpor.');
  const TEMPLATE=['NISN','Nama_Siswa','Kelas','Jurusan','Tempat_Lahir','Tanggal_Lahir','Nama_Orang_Tua','Nomor_HP_Orang_Tua','Alamat','Status','Tahun_Pelajaran'],clean=v=>v==null?'':String(v).trim(),normalizeNisn=v=>clean(v).replace(/\.0$/,'').padStart(10,'0');
  const imported=[],seen={};rows.forEach(function(r,i){if(String(r&&r.Kelas||'').trim()!==kelas)throw new Error('Baris '+(i+2)+': Kelas harus '+kelas+'.');const x={};TEMPLATE.forEach(k=>x[k]=r&&Object.prototype.hasOwnProperty.call(r,k)?r[k]:'');const nisn=normalizeNisn(x.NISN);if(!/^\d{10}$/.test(nisn))throw new Error('Baris '+(i+2)+': NISN harus tepat 10 digit.');if(!clean(x.Nama_Siswa))throw new Error('Baris '+(i+2)+': Nama_Siswa wajib diisi.');if(seen[nisn])throw new Error('Duplikat NISN dalam file: '+nisn);seen[nisn]=true;imported.push([generateID_('SIS'),nisn,clean(x.Nama_Siswa),kelas,clean(x.Jurusan),clean(x.Tempat_Lahir),clean(x.Tanggal_Lahir),clean(x.Nama_Orang_Tua),clean(x.Nomor_HP_Orang_Tua),clean(x.Alamat),clean(x.Status)||'Aktif',clean(x.Tahun_Pelajaran)||getTahunPelajaran_()]);});
  const lock=LockService.getScriptLock();lock.waitLock(15000);try{
    const sh=getSheet_(APP.SHEETS.SISWA),lastCol=12,lastRow=sh.getLastRow(),old=lastRow>=2?sh.getRange(2,1,lastRow-1,lastCol).getValues():[],keep=old.filter(r=>String(r[3]||'').trim()!==kelas),merged=keep.concat(imported);
    if(old.length)sh.getRange(2,1,old.length,lastCol).clearContent();if(merged.length){sh.getRange(2,2,merged.length,1).setNumberFormat('@');sh.getRange(2,9,merged.length,1).setNumberFormat('@');sh.getRange(2,1,merged.length,lastCol).setValues(merged);}SpreadsheetApp.flush();
    if(merged.length&&!sh.getRange(2,2).getDisplayValue().trim())throw new Error('Import kelas gagal diverifikasi di DATA_SISWA.');
    invalidateStudentCaches_();appendActivityFast_('IMPORT SISWA WALI KELAS',kelas+'; '+imported.length+' siswa tersimpan.');return{success:true,berhasil:imported.length,diganti:imported.length,kelas,mode:'REPLACE_CLASS',message:'Data kelas '+kelas+' berhasil tersimpan. Kelas lain dan histori tetap dipertahankan.'};
  }finally{lock.releaseLock();}
}
