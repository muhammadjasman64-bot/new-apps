/** 14_AutoSync.gs - sinkronisasi ringan. Rekap hanya periode aktif. */
function onEdit(e){
  try{
    if(!e||!e.range)return;const name=e.range.getSheet().getName();
    const source=[APP.SHEETS.CONFIG,APP.SHEETS.SISWA,APP.SHEETS.ABSENSI,APP.SHEETS.PELANGGARAN,APP.SHEETS.PENGHARGAAN,APP.SHEETS.MASTER_PELANGGARAN,APP.SHEETS.MASTER_PENGHARGAAN,APP.SHEETS.MASTER_PEMBINAAN,APP.SHEETS.TINDAKAN,APP.SHEETS.PEMANGGILAN_ORANG_TUA];
    if(!source.includes(name))return;
    if(e.range.getRow()===1)return;
    if([APP.SHEETS.MASTER_PELANGGARAN,APP.SHEETS.MASTER_PENGHARGAAN,APP.SHEETS.MASTER_PEMBINAAN].includes(name)){
      enforceOfficialRules_();
      return;
    }
    clearAppCache_();
    if([APP.SHEETS.PELANGGARAN,APP.SHEETS.PENGHARGAAN].includes(name) && e.range.getRow()>1){
      const nisnCol=3;
      const nisn=String(e.range.getSheet().getRange(e.range.getRow(),nisnCol).getValue()||'').trim();
      if(nisn) applyDecisionForStudent_(nisn);
    }
    if(name===APP.SHEETS.ABSENSI && e.range.getRow()>1){
      clearAppCache_();
    }
    refreshDerivedRekap_();
  }catch(err){console.log('onEdit sync: '+err);}
}
function refreshDerivedRekap_(){
  const lock=LockService.getScriptLock();if(!lock.tryLock(1000))return;
  try{const cfg=getConfigObject_(),y=Number(String(cfg.Tahun_Pelajaran||new Date().getFullYear()).split('/')[0])||new Date().getFullYear(),sem=String(cfg.Semester||1),month=Number(cfg.Bulan_Aktif||new Date().getMonth()+1),kelas=String(cfg.Kelas||'');
    // Hanya membangun tiga snapshot aktif; tidak membaca/menulis seluruh histori secara terpisah.
    generateAndSaveRekap_('bulanan',{year:y,month:month,kelas:kelas,tahunPelajaran:cfg.Tahun_Pelajaran});
    generateAndSaveRekap_('semester',{year:y,semester:sem,kelas:kelas,tahunPelajaran:cfg.Tahun_Pelajaran});
    generateAndSaveRekap_('tahunan',{year:y,kelas:kelas,tahunPelajaran:cfg.Tahun_Pelajaran});
  }finally{lock.releaseLock();}
}
function installAutoSyncTrigger_(){const exists=ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='onEdit');if(!exists)ScriptApp.newTrigger('onEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();clearAppCache_();return{success:true,message:'Auto Sync trigger aktif.'};}
function removeAutoSyncTrigger_(){ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='onEdit').forEach(t=>ScriptApp.deleteTrigger(t));return{success:true,message:'Auto Sync trigger dihapus.'};}
function getDbVersion_(){return PropertiesService.getScriptProperties().getProperty('DB_VERSION')||'0';}
