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

function importSiswaRowsCore_(rows, role, kelasWali){
  role=String(role||'admin'); kelasWali=String(kelasWali||'').trim();
  if(!Array.isArray(rows)||!rows.length)throw new Error('Tidak ada data untuk diimpor.');
  if(role==='wali_kelas'&&!kelasWali)throw new Error('Akun Wali Kelas belum memiliki Kelas.');
  const TEMPLATE=['NISN','Nama_Siswa','Kelas','Jurusan','Tempat_Lahir','Tanggal_Lahir','Nama_Orang_Tua','Nomor_HP_Orang_Tua','Alamat','Status','Tahun_Pelajaran'];
  const clean=v=>v==null?'':String(v).trim();
  const norm=v=>nisnText_(v);
  const normalized=[],seen={};
  rows.forEach((r,i)=>{
    const x={}; TEMPLATE.forEach(k=>x[k]=r&&Object.prototype.hasOwnProperty.call(r,k)?r[k]:'');
    const nisn=norm(x.NISN), kelas=clean(x.Kelas), nama=clean(x.Nama_Siswa);
    if(!/^\d{10}$/.test(nisn))throw new Error('Baris '+(i+2)+': NISN harus tepat 10 digit.');
    if(!nama)throw new Error('Baris '+(i+2)+': Nama_Siswa wajib diisi.');
    if(role==='wali_kelas'&&kelas!==kelasWali)throw new Error('Baris '+(i+2)+': Kelas harus '+kelasWali+'.');
    if(seen[nisn])throw new Error('Duplikat NISN dalam file: '+nisn); seen[nisn]=true;
    normalized.push({NISN:nisn,Nama_Siswa:nama,Kelas:kelas,Jurusan:clean(x.Jurusan),Tempat_Lahir:clean(x.Tempat_Lahir),Tanggal_Lahir:clean(x.Tanggal_Lahir),Nama_Orang_Tua:clean(x.Nama_Orang_Tua),Nomor_HP_Orang_Tua:clean(x.Nomor_HP_Orang_Tua),Alamat:clean(x.Alamat),Status:clean(x.Status)||'Aktif',Tahun_Pelajaran:clean(x.Tahun_Pelajaran)||getTahunPelajaran_()});
  });
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    const sh=getSheet_(APP.SHEETS.SISWA), lc=Math.max(sh.getLastColumn(),12);
    const h=sh.getRange(1,1,1,lc).getDisplayValues()[0].map(dbNormalizeHeader_);
    const col=k=>h.indexOf(dbNormalizeHeader_(k));
    const ni=col('nisn'); if(ni<0)throw new Error('Kolom NISN tidak ditemukan pada DATA_SISWA.');
    const idc=col('id_siswa')>=0?col('id_siswa'):col('id');
    const last=sh.getLastRow();
    const nisnValues=last>=2?sh.getRange(2,ni+1,last-1,1).getDisplayValues():[];
    const index={};
    nisnValues.forEach((r,i)=>{const n=nisnText_(r[0]);if(/^\d{10}$/.test(n)&&index[n]===undefined)index[n]=i+2;});
    const updates=[],appends=[];
    normalized.forEach(x=>{
      const rowNo=index[x.NISN];
      const vals=rowNo?sh.getRange(rowNo,1,1,lc).getValues()[0]:Array(lc).fill('');
      const set=(keys,v)=>{for(const k of (Array.isArray(keys)?keys:[keys])){const c=col(k);if(c>=0){vals[c]=v;return;}}};
      if(!rowNo&&idc>=0)vals[idc]=generateID_('SIS');
      set('nisn',x.NISN);set(['nama_siswa','nama'],x.Nama_Siswa);set('kelas',x.Kelas);set('jurusan',x.Jurusan);set('tempat_lahir',x.Tempat_Lahir);set('tanggal_lahir',x.Tanggal_Lahir);set('nama_orang_tua',x.Nama_Orang_Tua);set('nomor_hp_orang_tua',x.Nomor_HP_Orang_Tua);set('alamat',x.Alamat);set('status',x.Status);set('tahun_pelajaran',x.Tahun_Pelajaran);
      if(rowNo)updates.push({row:rowNo,vals}); else appends.push(vals);
    });
    // Existing rows: one write per changed row (safe for mixed data), new rows: one batch append.
    updates.forEach(u=>sh.getRange(u.row,1,1,lc).setValues([u.vals]));
    if(appends.length){const start=sh.getLastRow()+1;sh.getRange(start,ni+1,appends.length,1).setNumberFormat('@');sh.getRange(start,1,appends.length,lc).setValues(appends);}
    SpreadsheetApp.flush();
    // Verify every imported NISN. Existing rows are checked individually; new rows are one batch read.
    const verify={};
    Object.keys(index).forEach(n=>{ if(normalized.some(x=>x.NISN===n)) verify[n]=sh.getRange(index[n],ni+1).getDisplayValue().trim(); });
    if(appends.length){const start=sh.getLastRow()-appends.length+1;const nv=sh.getRange(start,ni+1,appends.length,1).getDisplayValues();appends.forEach((r,i)=>verify[nisnText_(r[ni])]=nv[i][0].trim());}
    const bad=normalized.findIndex(x=>nisnText_(verify[x.NISN]||'')!==x.NISN);
    if(bad>=0)throw new Error('Import gagal diverifikasi untuk NISN '+normalized[bad].NISN+'.');
    invalidateStudentCaches_(); appendActivityFast_('IMPORT SISWA', (role==='wali_kelas'?kelasWali+'; ':'')+normalized.length+' siswa (upsert).');
    return {success:true,berhasil:normalized.length,baru:appends.length,diperbarui:updates.length,dilewati:0,kelas:role==='wali_kelas'?kelasWali:'',mode:'UPSERT',message:'Data siswa berhasil disimpan ke DATA_SISWA tanpa menghapus data siswa lain.'};
  }finally{lock.releaseLock();}
}

function getDatabaseBindingInfo_(){
  const props=PropertiesService.getScriptProperties();
  const id=String(props.getProperty('SPREADSHEET_ID')||'').trim();
  if(!id)return {ok:false,message:'SPREADSHEET_ID belum diset.'};
  const ss=SpreadsheetApp.openById(id);
  return {ok:true,id:ss.getId(),name:ss.getName(),url:ss.getUrl(),activeId:(SpreadsheetApp.getActiveSpreadsheet()||{}).getId?SpreadsheetApp.getActiveSpreadsheet().getId():'',sheets:ss.getSheets().map(x=>x.getName())};
}

function importSiswaRowsScoped_(token, rows){ const s=requirePermission_(token,'write_students'); return importSiswaRowsCore_(rows,s.role,s.kelas); }
