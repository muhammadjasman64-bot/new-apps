/**
 * V42 - EARLY WARNING & INTERVENTION ENGINE
 * Peringatan dini adalah alat monitoring. Tidak menambah atau mengganti sanksi resmi.
 */
function earlyWarningKey_(nisn){return 'EW_SNAPSHOT_'+String(nisn).replace(/[^A-Za-z0-9_-]/g,'_')+'_'+String(getTahunPelajaran_()).replace(/[^A-Za-z0-9_-]/g,'_');}
function getPreviousEarlyWarningSnapshot_(nisn){try{const v=PropertiesService.getScriptProperties().getProperty(earlyWarningKey_(nisn));return v?JSON.parse(v):null;}catch(e){return null;}}
function saveEarlyWarningSnapshot_(x){PropertiesService.getScriptProperties().setProperty(earlyWarningKey_(x.nisn),JSON.stringify({ts:new Date().toISOString(),indeks:Number(x.indeksMonitoring)||0,level:x.level,kehadiran:Number(x.attendance.kehadiran)||0,alpha:Number(x.attendance.alphaSemester)||0,poin:Number(x.behavior.poinBersih)||0,status:x.behavior.status||''}));}
function buildEarlyWarning_(cur,prev){
  const reasons=[], actions=[]; let severity='INFO';
  const idx=Number(cur.indeksMonitoring)||0, pidx=prev?Number(prev.indeks)||0:null;
  const up=prev&&idx>pidx+0.01, levelUp=prev&&warningLevelRank_(cur.level)>warningLevelRank_(prev.level);
  if(levelUp){severity='HIGH';reasons.push('Level risiko meningkat dari '+prev.level+' menjadi '+cur.level);}
  if(up&&idx-(pidx||0)>=10){severity=severity==='HIGH'?'HIGH':'MEDIUM';reasons.push('Indeks monitoring meningkat '+(idx-(pidx||0)).toFixed(2)+' poin');}
  if(Number(cur.attendance.kehadiran)<90 && (!prev||Number(prev.kehadiran)>=90)){severity='HIGH';reasons.push('Kehadiran turun di bawah 90%');}
  if(Number(cur.attendance.alphaSemester)>=14 && (!prev||Number(prev.alpha)<14)){severity='HIGH';reasons.push('Alpa semester mencapai/melewati batas 14 hari');}
  if(cur.behavior.attention && (!prev||String(prev.status)!==String(cur.behavior.status))){severity=severity==='INFO'?'MEDIUM':severity;reasons.push('Tahap pembinaan berubah menjadi '+cur.behavior.status);}
  if(cur.urgent && !levelUp && !up && (!prev||!prev.urgent)){severity='HIGH';reasons.push('Siswa masuk indikator tindakan segera');}
  if(reasons.length){
    actions.push('Wali kelas melakukan verifikasi kondisi terbaru siswa.');
    if(Number(cur.attendance.kehadiran)<90 || Number(cur.attendance.alphaSemester)>0) actions.push('Periksa data dan riwayat absensi.');
    if(cur.behavior.attention) actions.push('Tinjau pembinaan sesuai Rule Engine dan aturan resmi.');
    if(cur.level==='RISIKO TINGGI') actions.push('Koordinasikan dengan Guru BK/pihak terkait sesuai kewenangan.');
  }
  return {trigger:reasons.length>0,severity:severity,reasons:reasons, rekomendasi:actions.join(' ')};
}
function warningLevelRank_(v){return {'AMAN':0,'PERLU PERHATIAN':1,'RISIKO MENENGAH':2,'RISIKO TINGGI':3}[String(v||'')]==null?0:{'AMAN':0,'PERLU PERHATIAN':1,'RISIKO MENENGAH':2,'RISIKO TINGGI':3}[String(v||'')];}
function createEarlyWarningAlert_(cur,ew){
  const sh=getSheet_(APP.SHEETS.PERINGATAN_DINI), rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,19).getValues():[];
  const fp=[cur.nisn,cur.indeksMonitoring.toFixed(2),cur.level,Number(cur.attendance.kehadiran).toFixed(2),cur.behavior.poinBersih,cur.behavior.status].join('|');
  const exists=rows.some(r=>[r[2],r[9],r[11],r[12],r[14],r[15]].join('|')===fp && String(r[16])!=='SELESAI');
  if(exists)return false;
  sh.appendRow([generateID_('EW'),new Date(),cur.nisn,cur.nama,cur.tahunPelajaran,cur.semester,ew.severity,ew.reasons.join(' • '),null,cur.indeksMonitoring, '',cur.level,cur.attendance.kehadiran,cur.attendance.alphaSemester,cur.behavior.poinBersih,ew.rekomendasi,'OPEN','','']);
  return true;
}
function evaluateEarlyWarningForStudent_(nisn){
  const cur=getUnifiedRiskForStudent_(nisn), prev=getPreviousEarlyWarningSnapshot_(nisn), ew=buildEarlyWarning_(cur,prev); let created=false;
  if(ew.trigger)created=createEarlyWarningAlert_(cur,ew);
  saveEarlyWarningSnapshot_(cur);
  return {success:true,student:cur,previous:prev,warning:ew,created:created};
}
function evaluateEarlyWarningAll_(){
  const students=getSiswaAktif_(),out=[]; students.forEach(s=>{try{out.push(evaluateEarlyWarningForStudent_(s.nisn));}catch(e){out.push({success:false,nisn:s.nisn,message:e.message});}});
  const created=out.filter(x=>x.created).length; clearAppCache_(); PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));
  if(created)logActivity_('EARLY WARNING','Membuat '+created+' peringatan dini.');
  return {success:true,total:out.length,created:created,items:out};
}
function getEarlyWarningSummary_(){
  const cached=cacheGetJson_('EARLY_WARNING_SUMMARY_CACHE');if(cached)return cached;
  const sh=getSheet_(APP.SHEETS.PERINGATAN_DINI), rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,19).getValues():[];
  const open=rows.filter(r=>String(r[16]||'OPEN')!=='SELESAI').map(r=>({id:r[0],timestamp:r[1],nisn:canonicalNisn_(r[2]),nama:r[3],tahunPelajaran:r[4],semester:r[5],severity:r[6],jenis:r[7],indeksSebelum:r[8],indeksSesudah:r[9],levelSebelum:r[10],levelSesudah:r[11],kehadiran:r[12],alpha:r[13],poin:r[14],rekomendasi:r[15],status:r[16]}));
  const counts={HIGH:0,MEDIUM:0,INFO:0};open.forEach(x=>counts[x.severity]=(counts[x.severity]||0)+1);
  return cachePutJson_('EARLY_WARNING_SUMMARY_CACHE',{success:true,totalOpen:open.length,counts:counts,items:open.slice(-50).reverse()},CACHE_TTL.DASHBOARD);
}

function getEarlyWarningForStudent_(nisn){
  const sh=getSheet_(APP.SHEETS.PERINGATAN_DINI), rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,19).getValues():[];
  return rows.filter(r=>String(r[2])===String(nisn)).map(r=>({id:r[0],timestamp:r[1],severity:r[6],jenis:r[7],indeksSebelum:r[8],indeksSesudah:r[9],levelSebelum:r[10],levelSesudah:r[11],kehadiran:r[12],alpha:r[13],poin:r[14],rekomendasi:r[15],status:r[16]})).reverse();
}
function resolveEarlyWarning_(id,status){
  const sh=getSheet_(APP.SHEETS.PERINGATAN_DINI), vals=sh.getDataRange().getValues(); let found=false;
  for(let i=1;i<vals.length;i++)if(String(vals[i][0])===String(id)){sh.getRange(i+1,17,1,3).setValues([[String(status||'SELESAI'),new Date(),Session.getActiveUser().getEmail()||'WebApp']]);found=true;break;}
  if(!found)throw new Error('Peringatan tidak ditemukan.'); clearAppCache_(); return {success:true};
}
function getEarlyWarningPolicy_(){return {name:'Early Warning & Intervention V42',nonPunitive:true,triggers:['Level risiko meningkat','Indeks naik >=10 poin','Kehadiran turun <90%','Alpa semester >=14','Tahap pembinaan berubah','Masuk indikator tindakan segera'],note:'Peringatan dini hanya untuk monitoring dan rekomendasi intervensi. Sanksi tetap mengikuti aturan resmi sekolah.'};}
