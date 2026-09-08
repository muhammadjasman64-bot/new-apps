/**
 * 19_RulesEngine.gs - V37 RULE INTEGRITY ENGINE
 * Sumber aturan: TATA TERTIB PESERTA DIDIK SMKN 01 BB fix.docx
 * Dokumen: Keputusan Kepala Sekolah No. 400.3.8.1/552/VII.III/SMKN1.BB/2026
 * Tanggal: 14 Juli 2026 | TP 2026/2027
 *
 * Prinsip:
 * 1. Master aturan tidak diinput dari Web App.
 * 2. Jika master diedit langsung di Google Sheets, sistem mengembalikan
 *    master ke definisi resmi saat trigger onEdit berjalan.
 * 3. Hash integritas disimpan agar perubahan aturan dapat dideteksi.
 */

const RULE_ENGINE_ = Object.freeze({
  source: 'TATA TERTIB PESERTA DIDIK SMKN 01 BB fix.docx',
  number: '400.3.8.1/552/VII.III/SMKN1.BB/2026',
  date: '14 Juli 2026',
  version: '2026/2027'
});

function getOfficialRuleSheets_(){
  return [APP.SHEETS.MASTER_PELANGGARAN,APP.SHEETS.MASTER_PENGHARGAAN,APP.SHEETS.MASTER_PEMBINAAN];
}

function normalizeRuleValue_(v){
  if(v instanceof Date) return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss');
  if(typeof v==='boolean') return v?'TRUE':'FALSE';
  return String(v==null?'':v).trim();
}

function ruleSheetFingerprint_(sheetName){
  const sh=getSheet_(sheetName), values=sh.getDataRange().getValues();
  const payload=values.map(row=>row.map(normalizeRuleValue_).join('\u001F')).join('\u001E');
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,payload,Utilities.Charset.UTF_8);
  return bytes.map(b=>('0'+(b&255).toString(16)).slice(-2)).join('');
}

function saveRuleFingerprints_(){
  const props=PropertiesService.getScriptProperties();
  getOfficialRuleSheets_().forEach(name=>props.setProperty('RULE_HASH_'+name,ruleSheetFingerprint_(name)));
  props.setProperty('RULE_SOURCE',RULE_ENGINE_.source);
  props.setProperty('RULE_VERSION',RULE_ENGINE_.version);
  props.setProperty('RULE_NUMBER',RULE_ENGINE_.number);
  props.setProperty('RULE_DATE',RULE_ENGINE_.date);
}

function getRuleIntegrity_(){
  const props=PropertiesService.getScriptProperties(), items=[];
  getOfficialRuleSheets_().forEach(name=>{
    const actual=ruleSheetFingerprint_(name), expected=props.getProperty('RULE_HASH_'+name)||'';
    items.push({sheet:name,ok:!!expected&&actual===expected,hash:actual,expected:expected});
  });
  return {ok:items.every(x=>x.ok),source:RULE_ENGINE_.source,version:RULE_ENGINE_.version,number:RULE_ENGINE_.number,date:RULE_ENGINE_.date,items:items};
}

function restoreOfficialRuleMasters_(){
  // Seed selalu berasal dari definisi resmi di 01_Setup.gs.
  seedMasterPelanggaranOfficial_();
  seedMasterPenghargaanOfficial_();
  seedMasterPembinaanOfficial_();
  saveRuleFingerprints_();
  clearAppCache_();
  PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));
  return {success:true,message:'Master aturan dipulihkan ke sumber resmi.',integrity:getRuleIntegrity_()};
}

function lockOfficialRuleSheets_(){
  getOfficialRuleSheets_().forEach(name=>{
    const sh=getSheet_(name);
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p=>{
      if(String(p.getDescription()).indexOf('RULE_LOCK_V37')===0) p.remove();
    });
    const p=sh.protect().setDescription('RULE_LOCK_V37 | Sumber: '+RULE_ENGINE_.source+' | TP '+RULE_ENGINE_.version);
    // Owner/script tetap dapat mengelola sheet; pengguna lain tidak diberi editor.
    try{p.setWarningOnly(false);}catch(e){/* fallback bila domain policy membatasi */}
  });
  saveRuleFingerprints_();
  return {success:true,message:'Master pelanggaran, penghargaan, dan pembinaan dikunci.',integrity:getRuleIntegrity_()};
}

function enforceOfficialRules_(){
  const integrity=getRuleIntegrity_();
  if(!integrity.ok){
    restoreOfficialRuleMasters_();
    logActivity_('RESTORE MASTER ATURAN','Perubahan manual terdeteksi; master dikembalikan ke aturan resmi.');
    return {restored:true,integrity:getRuleIntegrity_()};
  }
  return {restored:false,integrity:integrity};
}

function getAturanInfoV37_(){
  const c=getConfigObject_(), integrity=getRuleIntegrity_();
  return {
    sumber:RULE_ENGINE_.source,
    nomor:RULE_ENGINE_.number,
    tanggal:RULE_ENGINE_.date,
    versi:RULE_ENGINE_.version,
    tahunPelajaran:String(c.Tahun_Pelajaran||RULE_ENGINE_.version),
    statusIntegritas:integrity.ok?'TERKUNCI & VALID':'PERLU PEMULIHAN',
    catatan:'Master aturan bersifat read-only di Web App dan dipulihkan otomatis jika diubah langsung pada spreadsheet.'
  };
}

function validateOfficialRules_(){return getRuleIntegrity_();}
