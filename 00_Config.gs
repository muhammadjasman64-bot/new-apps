/** 00_Config.gs - WAJIB ditempatkan paling awal.
 * Berisi konfigurasi aplikasi + fungsi utilitas yang dipakai semua file.
 */
const APP = Object.freeze({
  NAME: 'Jurnal Wali Kelas Digital',
  SCHOOL: 'SMK NEGERI 1 BUAY BAHUGA',
  STORAGE: { FOLDER_NAME: 'Jurnal Wali Kelas Digital - Dokumen' },
  SHEETS: {
    CONFIG:'CONFIG', SISWA:'DATA_SISWA', MASTER_PELANGGARAN:'MASTER_PELANGGARAN', MASTER_PENGHARGAAN:'MASTER_PENGHARGAAN',
    ABSENSI:'ABSENSI', PELANGGARAN:'PELANGGARAN', PENGHARGAAN:'PENGHARGAAN', TINDAKAN:'TINDAKAN', PEMANGGILAN_ORANG_TUA:'PEMANGGILAN_ORANG_TUA', PEMANGGILAN:'PEMANGGILAN_ORANG_TUA',
    CATATAN_WALI_KELAS:'CATATAN_WALI_KELAS', CATATAN_BK:'CATATAN_BK', REKAP_HARIAN:'REKAP_HARIAN', REKAP_BULANAN:'REKAP_BULANAN',
    REKAP_SEMESTER:'REKAP_SEMESTER', REKAP_TAHUNAN:'REKAP_TAHUNAN', MASTER_PEMBINAAN:'MASTER_PEMBINAAN', PERINGATAN_DINI:'PERINGATAN_DINI', DOKUMEN:'DOKUMEN', LOG:'LOG_AKTIVITAS'
  },
  CONFIG_DEFAULTS: {
    Tahun_Pelajaran:'2026/2027', Semester:'1', Nama_Sekolah:'SMK NEGERI 1 BUAY BAHUGA',
    Alamat_Sekolah:'Jl. Simpang Empat Kebon Agung No. 08 Kec. Buay Bahuga Kab. Way Kanan',
    Kode_Pos:'34781', Email_Sekolah:'smkWANbahuga@yahoo.com', Nama_Kepala_Sekolah:'Jonlay Raden Supidin',
    NIP_Kepala_Sekolah:'197911302009021001', Nama_Waka_Kesiswaan:'', NIP_Waka_Kesiswaan:'', Nama_Guru_BK:'',
    NIP_Guru_BK:'', Nama_Wali_Kelas:'', NIP_Wali_Kelas:'', Kelas:'', Jurusan:'TKJ', ID_Kop_Surat:'', URL_Kop_Surat:'',
    ID_Logo:'', URL_Logo:'', ID_TTD_Wali:'', URL_TTD_Wali:'', ID_TTD_Waka:'', URL_TTD_Waka:'', ID_TTD_BK:'', URL_TTD_BK:'',
    ID_TTD_Kepala:'', URL_TTD_Kepala:'', Sumber_Aturan:'TATA TERTIB PESERTA DIDIK SMKN 01 BUAY BAHUGA fix.docx', Nomor_Keputusan_Aturan:'400.3.8.1/552/VII.III/SMKN1.BB/2026', Tanggal_Aturan:'14 Juli 2026', Versi_Aturan:'2026/2027', Min_Kehadiran_Semester:90, Max_Alpha_Semester:14, Max_Alpha_Tahunan:28
  }
});

function getSS_(){
  const active=SpreadsheetApp.getActiveSpreadsheet();
  if(active) return active;
  const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if(id){ try{return SpreadsheetApp.openById(id);}catch(e){} }
  return null;
}
function getSheet_(name){ const sh=getSS_().getSheetByName(name); if(!sh) throw new Error('Sheet '+name+' belum ada. Jalankan setupDatabase_ terlebih dahulu.'); return sh; }
function generateID_(prefix){ return prefix+'-'+Utilities.getUuid().split('-')[0].toUpperCase(); }
function formatTanggal_(v){ if(!v) return ''; return Utilities.formatDate(new Date(v),Session.getScriptTimeZone(),'dd/MM/yyyy'); }
function getConfigObject_(){
  const cached=typeof cacheGetJson_==='function'?cacheGetJson_('APP_CONFIG_CACHE'):null;
  if(cached) return cached;
  const sh=getSheet_(APP.SHEETS.CONFIG), lastRow=sh.getLastRow(), out={};
  if(lastRow>1){
    const values=sh.getRange(2,1,lastRow-1,2).getValues();
    values.forEach(r=>{if(r[0]!==''&&r[0]!=null)out[String(r[0])]=r[1];});
  }
  return typeof cachePutJson_==='function'?cachePutJson_('APP_CONFIG_CACHE',out,300):out;
}
function getTahunPelajaran_(){ return String(getConfigObject_().Tahun_Pelajaran||APP.CONFIG_DEFAULTS.Tahun_Pelajaran); }
function getSemesterAktif_(){ return String(getConfigObject_().Semester||APP.CONFIG_DEFAULTS.Semester); }
function escapeHtmlServer_(v){ return String(v==null?'':v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }

/** Canonical NISN: selalu mengikuti NISN pada DATA_SISWA. */
function canonicalNisn_(value){
  const raw=String(value==null?'':value).trim().replace(/\.0$/,'');
  if(!raw)return '';
  const siswa=typeof getSiswa_==='function'?getSiswa_():[];
  const exact=siswa.find(x=>String(x.nisn||'').trim()===raw);
  if(exact)return String(exact.nisn).trim();
  if(/^\d+$/.test(raw)){
    const padded=raw.padStart(10,'0');
    const hit=siswa.find(x=>String(x.nisn||'').trim()===padded);
    if(hit)return String(hit.nisn).trim();
    const numeric=String(Number(raw));
    const hit2=siswa.find(x=>String(Number(String(x.nisn||'').trim()))===numeric);
    if(hit2)return String(hit2.nisn).trim();
  }
  return raw;
}

function canonicalNisnForStudent_(value){
  const n=canonicalNisn_(value);
  if(/^\d{10}$/.test(n))return n;
  return n;
}

function repairConfig_(){
  const sh=getSheet_(APP.SHEETS.CONFIG);
  const existing=sh.getDataRange().getValues();
  const map={}; existing.slice(1).forEach(r=>{if(r[0]) map[String(r[0])]=r[1];});
  const rows=Object.keys(APP.CONFIG_DEFAULTS).map(k=>[k, Object.prototype.hasOwnProperty.call(map,k)?map[k]:APP.CONFIG_DEFAULTS[k]]);
  sh.clearContents(); sh.getRange(1,1,1,2).setValues([['Parameter','Nilai']]); sh.getRange(2,1,rows.length,2).setValues(rows); sh.setFrozenRows(1); sh.autoResizeColumns(1,2);
  return 'CONFIG berhasil diperbaiki dan dilengkapi.';
}
