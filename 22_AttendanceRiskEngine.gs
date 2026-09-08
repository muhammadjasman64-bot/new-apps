/**
 * V40 - AUTOMATIC ATTENDANCE RISK ENGINE
 * Risiko absensi dihitung real-time untuk semester aktif.
 * Dasar kebijakan: kehadiran minimum 90% per semester, maksimal Alpa 14 hari/semester,
 * dan maksimal Alpa 28 hari/tahun pelajaran.
 */
function getAttendancePolicy_(){
  const cfg=getConfigObject_();
  return {
    minAttendance:Number(cfg.Min_Kehadiran_Semester||90),
    maxAlphaSemester:Number(cfg.Max_Alpha_Semester||14),
    maxAlphaAnnual:Number(cfg.Max_Alpha_Tahunan||28)
  };
}

function getActiveSemesterBounds_(){
  const cfg=getConfigObject_();
  const year=Number(String(cfg.Tahun_Pelajaran||new Date().getFullYear()).split('/')[0])||new Date().getFullYear();
  const semester=String(cfg.Semester||1);
  return {year:year,semester:semester,bounds:getPeriodBounds_('semester',{year:year,semester:semester}),tahunPelajaran:String(cfg.Tahun_Pelajaran||getTahunPelajaran_())};
}

function classifyAttendanceRisk_(attendance,alpha,records,policy){
  attendance=Number(attendance||0); alpha=Number(alpha||0); records=Number(records||0);
  if(records<=0)return {level:'NO_DATA',label:'Belum Ada Data',reason:'Belum terdapat data absensi pada semester aktif.',urgent:false};
  if(alpha>policy.maxAlphaSemester || attendance<policy.minAttendance)
    return {level:'CRITICAL',label:'Risiko Tinggi',reason:alpha>policy.maxAlphaSemester?'Alpa telah melewati batas semester.':'Kehadiran berada di bawah minimum 90%.',urgent:true};
  if(alpha>=Math.max(1,policy.maxAlphaSemester-4) || attendance<93)
    return {level:'HIGH',label:'Perlu Tindakan',reason:alpha>=Math.max(1,policy.maxAlphaSemester-4)?'Jumlah Alpa mendekati batas semester.':'Kehadiran mendekati batas minimum.',urgent:true};
  if(alpha>=5 || attendance<95)
    return {level:'MEDIUM',label:'Perlu Perhatian',reason:'Tren kehadiran perlu dipantau.',urgent:false};
  return {level:'LOW',label:'Aman',reason:'Kehadiran masih dalam batas aman.',urgent:false};
}

function getAttendanceRiskForStudent_(nisn){
  const ctx=getActiveSemesterBounds_(),policy=getAttendancePolicy_(),s=findStudentByNisn_(nisn);
  if(!s)throw new Error('Siswa tidak ditemukan: '+nisn);
  const a=getRekapAbsensi_('semester',ctx.year,null,ctx.semester,'').find(x=>String(x.nisn).trim()===String(nisn).trim())||Object.assign({},s,{H:0,S:0,I:0,A:0,T:0,D:0,total:0,kehadiran:0});
  const annual=getRekapAbsensi_('tahunan',ctx.year,null,null,'').find(x=>String(x.nisn).trim()===String(nisn).trim())||Object.assign({},s,{H:0,S:0,I:0,A:0,T:0,D:0,total:0,kehadiran:0});
  const risk=classifyAttendanceRisk_(a.kehadiran,a.A,a.total,policy);
  const alphaAnnual=Number(annual.A||0);
  const annualLimitExceeded=alphaAnnual>policy.maxAlphaAnnual;
  return {
    nisn:s.nisn,nama:s.nama,kelas:s.kelas,tahunPelajaran:ctx.tahunPelajaran,semester:ctx.semester,
    H:Number(a.H||0),S:Number(a.S||0),I:Number(a.I||0),A:Number(a.A||0),T:Number(a.T||0),D:Number(a.D||0),
    total:Number(a.total||0),kehadiran:Number(a.kehadiran||0),alphaAnnual:alphaAnnual,annualAttendance:Number(annual.kehadiran||0),
    minAttendance:policy.minAttendance,maxAlphaSemester:policy.maxAlphaSemester,maxAlphaAnnual:policy.maxAlphaAnnual,
    annualLimitExceeded:annualLimitExceeded,riskLevel:risk.level,riskLabel:risk.label,riskReason:risk.reason,urgent:risk.urgent||annualLimitExceeded
  };
}

function getAttendanceRiskSummary_(){
  const cached=cacheGetJson_('ATTENDANCE_RISK_SUMMARY_CACHE');if(cached)return cached;
  const ctx=getActiveSemesterBounds_(),policy=getAttendancePolicy_(),students=getSiswaAktif_();
  const sem=getRekapAbsensi_('semester',ctx.year,null,ctx.semester,''),annual=getRekapAbsensi_('tahunan',ctx.year,null,null,'');
  const sm=Object.fromEntries(sem.map(x=>[String(x.nisn),x])),am=Object.fromEntries(annual.map(x=>[String(x.nisn),x])),items=[];
  students.forEach(s=>{const a=sm[String(s.nisn)]||{H:0,S:0,I:0,A:0,T:0,D:0,total:0,kehadiran:0},an=am[String(s.nisn)]||{A:0,kehadiran:0},risk=classifyAttendanceRisk_(a.kehadiran,a.A,a.total,policy),alphaAnnual=Number(an.A||0),annualLimitExceeded=alphaAnnual>policy.maxAlphaAnnual;items.push({nisn:s.nisn,nama:s.nama,kelas:s.kelas,tahunPelajaran:ctx.tahunPelajaran,semester:ctx.semester,H:Number(a.H||0),S:Number(a.S||0),I:Number(a.I||0),A:Number(a.A||0),T:Number(a.T||0),D:Number(a.D||0),total:Number(a.total||0),kehadiran:Number(a.kehadiran||0),alphaAnnual:alphaAnnual,annualAttendance:Number(an.kehadiran||0),minAttendance:policy.minAttendance,maxAlphaSemester:policy.maxAlphaSemester,maxAlphaAnnual:policy.maxAlphaAnnual,annualLimitExceeded:annualLimitExceeded,riskLevel:risk.level,riskLabel:risk.label,riskReason:risk.reason,urgent:risk.urgent||annualLimitExceeded});});
  const counts={NO_DATA:0,LOW:0,MEDIUM:0,HIGH:0,CRITICAL:0};items.forEach(x=>counts[x.riskLevel]=(counts[x.riskLevel]||0)+1);
  const urgent=items.filter(x=>x.urgent).sort((a,b)=>Number(b.A)-Number(a.A)||Number(a.kehadiran)-Number(b.kehadiran));
  return cachePutJson_('ATTENDANCE_RISK_SUMMARY_CACHE',{success:true,policy:policy,tahunPelajaran:ctx.tahunPelajaran,semester:ctx.semester,totalSiswa:items.length,counts:counts,totalUrgent:urgent.length,urgent:urgent.slice(0,50),items:items},CACHE_TTL.DASHBOARD);
}

function getCombinedRiskSummary_(){
  const att=getAttendanceRiskSummary_(), behavior=getStatusAlertSummary_();
  const byNisn={};att.items.forEach(x=>byNisn[String(x.nisn)]=x);
  const items=(behavior.attention||[]).map(b=>Object.assign({},b,{risikoAbsensi:byNisn[String(b.nisn)]||null}));
  const combined=att.items.map(a=>{
    const b=(behavior.attention||[]).find(x=>String(x.nisn)===String(a.nisn));
    const behaviorUrgent=!!(b&&b.urgent),overallUrgent=a.urgent||behaviorUrgent;
    return Object.assign({},a,{statusPembinaan:b?b.status:'NORMAL',poinBersih:b?b.poinBersih:0,risikoGabungan:overallUrgent?'TINGGI':(a.riskLevel==='MEDIUM'||b)?'PERLU PERHATIAN':'AMAN',overallUrgent:overallUrgent});
  }).sort((a,b)=>Number(b.overallUrgent)-Number(a.overallUrgent)||Number(b.poinBersih)-Number(a.poinBersih)||Number(a.kehadiran)-Number(b.kehadiran));
  return {success:true,attendance:att,behavior:behavior,totalUrgent:combined.filter(x=>x.overallUrgent).length,items:combined};
}

function evaluateAttendanceRiskAll_(){
  const r=getCombinedRiskSummary_();clearAppCache_();PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));
  logActivity_('ATTENDANCE RISK ENGINE','Evaluasi risiko absensi seluruh siswa');return r;
}
