/**
 * V41 - UNIFIED STUDENT RISK ENGINE
 *
 * Menggabungkan indikator absensi dan pembinaan dalam satu profil monitoring.
 * CATATAN PENTING: Indeks 0-100 di bawah adalah INDEKS MONITORING APLIKASI,
 * bukan sanksi baru dan tidak menggantikan MASTER_PEMBINAAN resmi.
 */
function getUnifiedRiskForStudent_(nisn){
  const s=findStudentByNisn_(nisn); if(!s) throw new Error('Siswa tidak ditemukan: '+nisn);
  const behavior=getStudentStatusAlert_(nisn);
  const attendance=getAttendanceRiskForStudent_(nisn);
  const calls=getPemanggilanHistory_({nisn:nisn});
  const actions=getTindakanHistory_({nisn:nisn});

  const net=Math.max(0,Number(behavior.poinPelanggaran)||0);
  const behaviorScore=Math.min(60,Number((net/100*60).toFixed(2)));
  let attendanceScore=0;
  if(attendance.total>0){
    if(attendance.kehadiran<90) attendanceScore=40;
    else if(attendance.kehadiran<93) attendanceScore=30;
    else if(attendance.kehadiran<95) attendanceScore=20;
    else if(attendance.kehadiran<97) attendanceScore=10;
  }
  if(attendance.annualLimitExceeded) attendanceScore=Math.max(attendanceScore,40);
  const index=Math.min(100,Number((behaviorScore+attendanceScore).toFixed(2)));

  let level='AMAN';
  if(index>=75) level='RISIKO TINGGI';
  else if(index>=50) level='RISIKO MENENGAH';
  else if(index>=25) level='PERLU PERHATIAN';

  const urgent=!!behavior.urgent||!!attendance.urgent||index>=75;
  const reasons=[];
  if(behavior.attention) reasons.push('Pembinaan: '+behavior.status);
  if(attendance.riskLevel!=='LOW'&&attendance.riskLevel!=='NO_DATA') reasons.push('Absensi: '+attendance.riskLabel);
  if(attendance.annualLimitExceeded) reasons.push('Alpa tahunan melewati batas');
  if(!reasons.length) reasons.push('Tidak ada indikator risiko utama.');

  const latestCall=calls.length?calls[0]:null;
  const latestAction=actions.length?actions[0]:null;
  return {
    nisn:s.nisn,nama:s.nama,kelas:s.kelas,tahunPelajaran:attendance.tahunPelajaran,semester:attendance.semester,
    indeksMonitoring:index,level:level,urgent:urgent,alasan:reasons.join(' • '),
    behavior:{poinPelanggaran:behavior.poinPelanggaran,poinPenghargaan:behavior.poinPenghargaan,poinBersih:behavior.poinBersih,status:behavior.status,pemanggilan:behavior.pemanggilan,pihak:behavior.pihak,tindakan:behavior.tindakan,dokumen:behavior.dokumen},
    attendance:{kehadiran:attendance.kehadiran,alphaSemester:attendance.A,alphaTahunan:attendance.alphaAnnual,riskLevel:attendance.riskLevel,riskLabel:attendance.riskLabel},
    components:{pembinaan:behaviorScore,absensi:attendanceScore},
    riwayat:{jumlahPemanggilan:calls.length,jumlahTindakan:actions.length,pemanggilanTerakhir:latestCall?tanggalSafe_(latestCall.tanggal):'',tindakanTerakhir:latestAction?String(latestAction.status||''):'',},
    disclaimer:'Indeks monitoring aplikasi; bukan sanksi tambahan. Keputusan pembinaan tetap mengikuti MASTER_PEMBINAAN.'
  };
}
function tanggalSafe_(v){return v||'';}
function getUnifiedRiskSummary_(){
  const cached=cacheGetJson_('UNIFIED_RISK_SUMMARY_CACHE');if(cached)return cached;
  const att=getAttendanceRiskSummary_(),behavior=getStatusAlertSummary_(),bm=Object.fromEntries((behavior.attention||[]).map(x=>[String(x.nisn),x]));
  const items=(att.items||[]).map(a=>{const b=bm[String(a.nisn)],net=b?Number(b.poinBersih)||0:0;let attendanceScore=0;if(Number(a.total)>0){if(Number(a.kehadiran)<90)attendanceScore=40;else if(Number(a.kehadiran)<93)attendanceScore=30;else if(Number(a.kehadiran)<95)attendanceScore=20;else if(Number(a.kehadiran)<97)attendanceScore=10;}if(a.annualLimitExceeded)attendanceScore=Math.max(attendanceScore,40);const behaviorScore=Math.min(60,Number((net/100*60).toFixed(2))),index=Math.min(100,Number((behaviorScore+attendanceScore).toFixed(2)));let level='AMAN';if(index>=75)level='RISIKO TINGGI';else if(index>=50)level='RISIKO MENENGAH';else if(index>=25)level='PERLU PERHATIAN';const urgent=!!a.urgent||!!(b&&b.urgent)||index>=75,reasons=[];if(b&&b.attention)reasons.push('Pembinaan: '+b.status);if(a.riskLevel!=='LOW'&&a.riskLevel!=='NO_DATA')reasons.push('Absensi: '+a.riskLabel);if(a.annualLimitExceeded)reasons.push('Alpa tahunan melewati batas');if(!reasons.length)reasons.push('Tidak ada indikator risiko utama.');return {nisn:a.nisn,nama:a.nama,kelas:a.kelas,tahunPelajaran:a.tahunPelajaran,semester:a.semester,indeksMonitoring:index,level:level,urgent:urgent,alasan:reasons.join(' • '),behavior:{poinPelanggaran:b?b.poinPelanggaran:0,poinPenghargaan:b?b.poinPenghargaan:0,poinBersih:net,status:b?b.status:'NORMAL',pemanggilan:b?b.pemanggilan:0,pihak:b?b.pihak:'',tindakan:b?b.tindakan:'',dokumen:b?b.dokumen:''},attendance:{kehadiran:a.kehadiran,alphaSemester:a.A,alphaTahunan:a.alphaAnnual,riskLevel:a.riskLevel,riskLabel:a.riskLabel},components:{pembinaan:behaviorScore,absensi:attendanceScore}};});
  items.sort((a,b)=>Number(b.indeksMonitoring)-Number(a.indeksMonitoring)||Number(b.urgent)-Number(a.urgent));
  const counts={'AMAN':0,'PERLU PERHATIAN':0,'RISIKO MENENGAH':0,'RISIKO TINGGI':0};items.forEach(x=>counts[x.level]=(counts[x.level]||0)+1);
  return cachePutJson_('UNIFIED_RISK_SUMMARY_CACHE',{success:true,totalSiswa:items.length,counts:counts,totalUrgent:items.filter(x=>x.urgent).length,items:items,top:items.slice(0,20)},CACHE_TTL.DASHBOARD);
}
function getUnifiedRiskAlerts_(){
  const r=getUnifiedRiskSummary_();
  return {success:true,total:r.totalUrgent,items:r.items.filter(x=>x.urgent).slice(0,50)};
}
function evaluateUnifiedRiskAll_(){
  const r=getUnifiedRiskSummary_();
  clearAppCache_();
  PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));
  logActivity_('UNIFIED RISK ENGINE','Evaluasi '+r.totalSiswa+' siswa');
  return r;
}
function getUnifiedRiskPolicy_(){
  return {
    name:'Indeks Monitoring Siswa V41',
    formula:'Indeks = komponen pembinaan (maks. 60) + komponen absensi (maks. 40)',
    behaviorMax:60,attendanceMax:40,
    thresholds:[
      {min:0,max:24.99,label:'AMAN'},
      {min:25,max:49.99,label:'PERLU PERHATIAN'},
      {min:50,max:74.99,label:'RISIKO MENENGAH'},
      {min:75,max:100,label:'RISIKO TINGGI'}
    ],
    note:'Indeks ini hanya untuk prioritas monitoring. Sanksi/tahap pembinaan tetap ditentukan oleh aturan resmi sekolah.'
  };
}
