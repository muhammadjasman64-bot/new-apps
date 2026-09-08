/**
 * V39 - Automatic Student Status & Alert
 * Status selalu dihitung dari poin bersih Tahun Pelajaran aktif + MASTER_PEMBINAAN resmi.
 * Tidak menyimpan status manual baru; alert dibuat sebagai hasil perhitungan real-time.
 */
function getStudentStatusAlert_(nisn){
  const s=getDecisionState_(nisn), d=s.keputusan;
  const urgent=['SP-1','SP-2','SP-3','SKORSING','KONFERENSI KASUS'].includes(String(d.status));
  const attention=d.status!=='NORMAL';
  return {
    nisn:s.siswa.nisn,nama:s.siswa.nama,kelas:s.siswa.kelas,
    poinPelanggaran:s.poinPelanggaran,poinPenghargaan:s.poinPenghargaan,poinBersih:s.poinBersih,
    status:d.status,min:d.min||0,max:d.max||4,pemanggilan:d.pemanggilan||0,pihak:d.pihak||'',
    tindakan:d.tindakan||'',dokumen:d.dokumen||d.surat||'',
    urgent:urgent,attention:attention,
    tindakanSudahDicatat:s.tindakanSudahDicatat,pemanggilanSudahDicatat:s.pemanggilanSudahDicatat
  };
}

function getStudentAlerts_(options){
  options=options||{};
  const students=getSiswaAktif_(), alerts=[];
  students.forEach(s=>{try{const a=getStudentStatusAlert_(s.nisn); if(options.onlyAttention!==false && !a.attention)return; alerts.push(a);}catch(e){}});
  alerts.sort((a,b)=>Number(b.poinBersih)-Number(a.poinBersih));
  return {success:true,total:alerts.length,urgent:alerts.filter(x=>x.urgent).length,attention:alerts.length,items:alerts};
}

function getStatusAlertSummary_(){
  const cached=cacheGetJson_('STATUS_ALERT_SUMMARY_CACHE');if(cached)return cached;
  const students=getSiswaAktif_(),cfg=getConfigObject_(),tp=String(cfg.Tahun_Pelajaran||getTahunPelajaran_()),points=getPointSummaryMap_(null,null,tp);
  const counts={},items=[];
  students.forEach(s=>{
    const p=points[String(s.nisn).trim()]||{pel:0,rew:0};
    const pel=toNumberPoin_(p.pel),rew=toNumberPoin_(p.rew),bersih=Math.max(0,pel-rew),d=getKeputusanPoin_(bersih);
    const a={nisn:s.nisn,nama:s.nama,kelas:s.kelas,poinPelanggaran:pel,poinPenghargaan:rew,poinBersih:bersih,status:d.status,min:d.min||0,max:d.max||4,pemanggilan:d.pemanggilan||0,pihak:d.pihak||'',tindakan:d.tindakan||'',dokumen:d.dokumen||d.surat||'',urgent:['SP-1','SP-2','SP-3','SKORSING','KONFERENSI KASUS'].includes(String(d.status)),attention:d.status!=='NORMAL'};
    counts[a.status]=(counts[a.status]||0)+1;items.push(a);
  });
  const urgent=items.filter(x=>x.urgent).sort((a,b)=>b.poinBersih-a.poinBersih),attention=items.filter(x=>x.attention).sort((a,b)=>b.poinBersih-a.poinBersih);
  return cachePutJson_('STATUS_ALERT_SUMMARY_CACHE',{success:true,counts:counts,totalSiswa:students.length,totalAttention:attention.length,totalUrgent:urgent.length,urgent:urgent.slice(0,20),attention:attention.slice(0,50)},CACHE_TTL.DASHBOARD);
}

function evaluateStudentStatusAlerts_(){
  const summary=getStatusAlertSummary_();
  clearAppCache_();
  PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));
  return summary;
}

function getDecisionThresholdTest_(){
  const points=[0,4,5,25,26,50,51,65,66,75,76,85,86,90,91,100,101];
  return points.map(p=>{const d=getKeputusanPoin_(p);return {poin:p,status:d.status,pemanggilan:d.pemanggilan||0,pihak:d.pihak||'',tindakan:d.tindakan||''};});
}
