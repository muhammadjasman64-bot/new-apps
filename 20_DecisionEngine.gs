/**
 * 20_DecisionEngine.gs - V38 AUTOMATIC DECISION ENGINE
 * Semua keputusan pembinaan mengikuti MASTER_PEMBINAAN resmi.
 * Tidak ada input manual untuk status/tahap pembinaan.
 */
function getDecisionState_(nisn){
  const p=getProfilPembinaan_(nisn), d=p.keputusan;
  const history=getTindakanHistory_({nisn:nisn});
  const calls=getPemanggilanHistory_({nisn:nisn});
  const samePoint=history.filter(x=>Number(x.totalPoin)===Number(p.poinBersih));
  const alreadyStage=samePoint.some(x=>String(x.status)===String(d.status));
  const callExists=Number(d.pemanggilan)>0 && calls.some(x=>Number(x.ke)===Number(d.pemanggilan));
  return {
    siswa:p.siswa,
    poinPelanggaran:p.poinPelanggaran,
    poinPenghargaan:p.poinPenghargaan,
    poinBersih:p.poinBersih,
    keputusan:d,
    tindakanSudahDicatat:alreadyStage,
    pemanggilanSudahDicatat:callExists,
    riwayatTindakan:history,
    riwayatPemanggilan:calls
  };
}

function applyDecisionForStudent_(nisn,options){
  options=options||{};
  const state=getDecisionState_(nisn), d=state.keputusan, cfg=getConfigObject_();
  const result={success:true,nisn:String(nisn),poinBersih:state.poinBersih,status:d.status,createdTindakan:false,createdPemanggilan:false,skipped:false};
  if(state.poinBersih<5){ result.skipped=true; result.reason='Belum mencapai 5 poin.'; return result; }

  // Satu record tindakan untuk setiap tahap yang benar-benar dicapai.
  if(!state.tindakanSudahDicatat && options.createTindakan!==false){
    const sh=getSheet_(APP.SHEETS.TINDAKAN);
    sh.appendRow([generateID_('TIN'),new Date(),state.siswa.nisn,state.siswa.nama,state.poinBersih,d.status,d.tindakan,d.dokumen||'',
      Session.getActiveUser().getEmail()||'WebApp','Dibuat otomatis oleh Decision Engine V38',cfg.Tahun_Pelajaran||getTahunPelajaran_(),cfg.Semester||getSemesterAktif_()]);
    result.createdTindakan=true;
  }

  // Pemanggilan hanya dibuat ketika aturan memang mensyaratkannya.
  if(Number(d.pemanggilan)>0 && !state.pemanggilanSudahDicatat && options.createPemanggilan!==false){
    const sh=getSheet_(APP.SHEETS.PEMANGGILAN_ORANG_TUA);
    sh.appendRow([generateID_('PGL'),new Date(),state.siswa.nisn,state.siswa.nama,state.poinBersih,Number(d.pemanggilan),d.pihak||'',
      '','', 'Belum Ditindaklanjuti',cfg.Tahun_Pelajaran||getTahunPelajaran_(),cfg.Semester||getSemesterAktif_()]);
    result.createdPemanggilan=true;
  }
  if(result.createdTindakan||result.createdPemanggilan){
    clearAppCache_();
    PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));
    logActivity_('DECISION ENGINE',state.siswa.nisn+' | '+state.poinBersih+' poin | '+d.status);
  }
  return result;
}

function evaluateAllStudentsDecisions_(){
  const students=getSiswaAktif_(), results=[];
  students.forEach(s=>{try{results.push(applyDecisionForStudent_(s.nisn));}catch(e){results.push({success:false,nisn:s.nisn,message:e.message});}});
  return {success:true,total:results.length,createdTindakan:results.filter(x=>x.createdTindakan).length,createdPemanggilan:results.filter(x=>x.createdPemanggilan).length,errors:results.filter(x=>x.success===false),results:results};
}

function getDecisionPreview_(nisn){
  const s=getDecisionState_(nisn);
  return {success:true,siswa:s.siswa,poinPelanggaran:s.poinPelanggaran,poinPenghargaan:s.poinPenghargaan,poinBersih:s.poinBersih,keputusan:s.keputusan,tindakanSudahDicatat:s.tindakanSudahDicatat,pemanggilanSudahDicatat:s.pemanggilanSudahDicatat};
}

function installDecisionEngineTrigger_(){
  const exists=ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='onEdit');
  if(!exists) ScriptApp.newTrigger('onEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  return {success:true,message:'Decision Engine menggunakan trigger onEdit yang sama dengan Auto Sync.'};
}
