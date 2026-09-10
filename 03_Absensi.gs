/** 03_Absensi.gs - absensi dengan agregasi satu lintasan. */

/**
 * NISN ABSENSI MASTER: DATA_SISWA adalah sumber tunggal NISN.
 * Sengaja ditempatkan di file ABSENSI saja agar patch ini tidak mengubah file lain.
 * Menggunakan displayValues agar leading zero dari DATA_SISWA dipertahankan.
 */
function absensiMasterNisnMap_(){
  const sh=getSheet_(APP.SHEETS.SISWA), lr=sh.getLastRow(), lc=sh.getLastColumn();
  const out={};
  if(lr<2||lc<1)return out;
  const h=sh.getRange(1,1,1,lc).getDisplayValues()[0].map(dbNormalizeHeader_);
  const map=Object.fromEntries(h.map((x,i)=>[x,i]));
  const c=dbFindColumn_(map,['NISN','NIS','Nomor Induk Siswa Nasional']);
  if(c<0)return out;
  const rows=sh.getRange(2,1,lr-1,lc).getDisplayValues();
  rows.forEach(r=>{
    const exact=String(r[c]??'').trim(); if(!exact)return;
    const variants=[exact, exact.replace(/\.0$/,'')];
    if(/^\d+$/.test(exact))variants.push(exact.padStart(10,'0'));
    variants.forEach(v=>{ if(v)out[v]=exact; });
  });
  return out;
}
function absensiCanonicalNisn_(value){ return canonicalNisn_(value); }
function absensiSyncAllNisnToMaster_(){
  const master=absensiMasterNisnMap_();
  const names=[APP.SHEETS.ABSENSI,APP.SHEETS.PELANGGARAN,APP.SHEETS.PENGHARGAAN,APP.SHEETS.TINDAKAN,APP.SHEETS.PEMANGGILAN_ORANG_TUA,APP.SHEETS.CATATAN_WALI_KELAS,APP.SHEETS.CATATAN_BK,APP.SHEETS.REKAP_HARIAN,APP.SHEETS.REKAP_BULANAN,APP.SHEETS.REKAP_SEMESTER,APP.SHEETS.REKAP_TAHUNAN,APP.SHEETS.PERINGATAN_DINI];
  let changed=0;
  names.forEach(name=>{
    let sh; try{sh=getSheet_(name);}catch(e){return;}
    const lr=sh.getLastRow(),lc=sh.getLastColumn(); if(lr<2||lc<1)return;
    const h=sh.getRange(1,1,1,lc).getDisplayValues()[0].map(dbNormalizeHeader_);
    const map=Object.fromEntries(h.map((x,i)=>[x,i]));
    const c=dbFindColumn_(map,['NISN','NIS','Nomor Induk Siswa Nasional']); if(c<0)return;
    const vals=sh.getRange(2,1,lr-1,lc).getValues();
    const disp=sh.getRange(2,c+1,lr-1,1).getDisplayValues();
    let dirty=false;
    vals.forEach((r,i)=>{
      const old=String(disp[i][0]??'').trim();
      if(!old)return;
      const n=master[old]||master[old.replace(/\.0$/,'')]||master[old.padStart(10,'0')];
      if(n && String(r[c]??'').trim()!==n){r[c]=n;dirty=true;changed++;}
    });
    if(dirty){
      sh.getRange(2,c+1,lr-1,1).setNumberFormat('@');
      sh.getRange(2,1,lr-1,lc).setValues(vals);
    }
  });
  return changed;
}

function saveAbsensi_(records){
  if(typeof ensureAttendanceMonthCycle_==='function')ensureAttendanceMonthCycle_();
  if(!Array.isArray(records)||!records.length)throw new Error('Data absensi kosong.');
  const sh=getSheet_(APP.SHEETS.ABSENSI),cfg=getConfigObject_(),user=Session.getActiveUser().getEmail()||'WebApp',now=new Date();
  const lastRow=sh.getLastRow(),lastCol=sh.getLastColumn();
  const headers=lastCol?sh.getRange(1,1,1,lastCol).getValues()[0].map(dbNormalizeHeader_):[];
  const map=Object.fromEntries(headers.map((x,i)=>[x,i]));
  const cId=dbFindColumn_(map,['ID','ID_Absensi']),cT=dbFindColumn_(map,['Tanggal','Tanggal_Absensi','Tanggal Absensi','Tgl']),cN=dbFindColumn_(map,['NISN','NIS','Nomor Induk Siswa Nasional']),cNm=dbFindColumn_(map,['Nama_Siswa','Nama Siswa','Nama']),cK=dbFindColumn_(map,['Kelas','Rombel','Kelas/Rombel']),cS=dbFindColumn_(map,['Status','Status Kehadiran','Kehadiran']),cKet=dbFindColumn_(map,['Keterangan','Catatan']),cU=dbFindColumn_(map,['User','Penginput','Dibuat_Oleh']),cY=dbFindColumn_(map,['Tahun_Pelajaran','Tahun Pelajaran']),cSem=dbFindColumn_(map,['Semester']);
  if(cT<0||cN<0||cS<0)throw new Error('ABSENSI: header wajib tidak ditemukan. Diperlukan Tanggal, NISN, dan Status.');
  const all=lastRow>=2?sh.getRange(2,1,lastRow-1,lastCol).getValues():[];
  const rowByKey={}; all.forEach((r,i)=>{const n=absensiCanonicalNisn_(cN>=0?r[cN]:''),d=cT>=0?formatDateKey_(r[cT]):'';if(n&&d)rowByKey[n+'|'+d]=i+2;});
  const updates=[],adds=[]; let updated=0,inserted=0;
  records.forEach(x=>{
    const s=findStudentByNisn_(absensiCanonicalNisn_(x.nisn));if(!s)throw new Error('NISN tidak ditemukan: '+x.nisn);
    const status=String(x.status||'H').toUpperCase();if(!['H','S','I','A','T','D'].includes(status))throw new Error('Status absensi tidak valid: '+status);
    const dt=parseDateInput_(x.tanggal)||now, key=absensiCanonicalNisn_(s.nisn)+'|'+formatDateKey_(dt), existingRow=rowByKey[key];
    const vals={};
    if(cId>=0)vals[cId]=existingRow?sh.getRange(existingRow,cId+1).getValue():generateID_('ABS');
    if(cT>=0)vals[cT]=dt;if(cN>=0)vals[cN]=absensiCanonicalNisn_(s.nisn);if(cNm>=0)vals[cNm]=s.nama;if(cK>=0)vals[cK]=x.kelas||s.kelas;if(cS>=0)vals[cS]=status;if(cKet>=0)vals[cKet]=x.keterangan||'';if(cU>=0)vals[cU]=user;if(cY>=0)vals[cY]=cfg.Tahun_Pelajaran||getTahunPelajaran_();if(cSem>=0)vals[cSem]=cfg.Semester||getSemesterAktif_();
    const row=existingRow?all[existingRow-2].slice():Array(lastCol).fill('');Object.keys(vals).forEach(k=>row[Number(k)]=vals[k]);
    if(existingRow){updates.push({row:existingRow,values:row});updated++;}else{adds.push(row);inserted++;}
  });
  updates.forEach(u=>sh.getRange(u.row,1,1,lastCol).setValues([u.values]));
  if(adds.length)sh.getRange(sh.getLastRow()+1,1,adds.length,lastCol).setValues(adds);
  clearAppCache_();PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));logActivity_('INPUT ABSENSI',inserted+' baru, '+updated+' diperbarui');
  return{success:true,count:inserted+updated,inserted:inserted,updated:updated,date:formatDateKey_(parseDateInput_(records[0].tanggal)||now)};
}

function getAbsensiByDate_(dateStr,kelas){
  const d=formatDateKey_(parseDateInput_(dateStr)),sh=getSheet_(APP.SHEETS.ABSENSI),lastRow=sh.getLastRow(),lastCol=sh.getLastColumn(),out={};if(lastRow<2)return out;
  const all=sh.getRange(1,1,lastRow,lastCol).getValues(),h=all[0].map(dbNormalizeHeader_),ct=dbFindColumn_(Object.fromEntries(h.map((x,i)=>[x,i])),['Tanggal','Tanggal_Absensi','Tanggal Absensi','Tgl']),cn=dbFindColumn_(Object.fromEntries(h.map((x,i)=>[x,i])),['NISN','NIS']),ck=dbFindColumn_(Object.fromEntries(h.map((x,i)=>[x,i])),['Kelas','Rombel']),cs=dbFindColumn_(Object.fromEntries(h.map((x,i)=>[x,i])),['Status','Status Kehadiran','Kehadiran']),cket=dbFindColumn_(Object.fromEntries(h.map((x,i)=>[x,i])),['Keterangan','Catatan']);
  if(ct<0||cn<0||cs<0)return out;all.slice(1).forEach(r=>{if(formatDateKey_(r[ct])===d&&(!kelas||ck<0||String(r[ck])===String(kelas)))out[absensiCanonicalNisn_(r[cn])]={status:r[cs],keterangan:cket>=0?r[cket]:''};});return out;
}
function getRekapAbsensi_(period,year,month,semester,kelas,weekStart){
  period=String(period||'bulanan').toLowerCase();
  const ck=cacheKey_('ABS_REKAP',[period,year||'',month||'',semester||'',kelas||'',weekStart||'']);
  const cached=cacheGetJson_(ck);if(cached)return cached;const y=Number(year||new Date().getFullYear()),m=Number(month||1),sem=String(semester||'1');
  const sh=getSheet_(APP.SHEETS.ABSENSI),lastRow=sh.getLastRow(),lastCol=sh.getLastColumn(),siswa=getSiswa_().filter(s=>String(s.status||'Aktif').trim().toLowerCase()!=='nonaktif');
  const resultMap={};siswa.forEach(s=>{resultMap[absensiCanonicalNisn_(s.nisn).trim()]={H:0,S:0,I:0,A:0,T:0,D:0};});
  if(lastRow>=2&&lastCol){
    const all=sh.getRange(1,1,lastRow,lastCol).getValues(),headers=all[0].map(dbNormalizeHeader_),map=Object.fromEntries(headers.map((x,i)=>[x,i]));
    const ct=dbFindColumn_(map,['Tanggal','Tanggal_Absensi','Tanggal Absensi','Tgl']),cn=dbFindColumn_(map,['NISN','NIS','Nomor Induk Siswa Nasional']),ck=dbFindColumn_(map,['Kelas','Rombel','Kelas/Rombel']),cs=dbFindColumn_(map,['Status','Status Kehadiran','Kehadiran']);
    if(ct<0||cn<0||cs<0)throw new Error('ABSENSI: header wajib tidak ditemukan. Diperlukan Tanggal, NISN, dan Status.');
    const kelasFilter=String(kelas==null?'':kelas).trim();
    let wa=null,wb=null;if(period==='mingguan'){wa=parseDateInput_(weekStart||new Date());if(!wa)throw new Error('Tanggal awal minggu tidak valid.');wa.setHours(0,0,0,0);wb=new Date(wa);wb.setDate(wb.getDate()+6);wb.setHours(23,59,59,999);}
    all.slice(1).forEach(r=>{
      const d=parseDateInput_(r[ct]);if(!d)return;const yy=d.getFullYear(),mm=d.getMonth()+1;
      const inPeriod=period==='mingguan'?(d>=wa&&d<=wb):period==='bulanan'?yy===y&&mm===m:period==='semester'?(sem==='2'?yy===y+1&&mm>=1&&mm<=6:yy===y&&mm>=7&&mm<=12):period==='tahunan'?((yy===y&&mm>=7)||(yy===y+1&&mm<=6)):false;if(!inPeriod)return;
      if(kelasFilter&&ck>=0&&String(r[ck]||'').trim()!==kelasFilter)return;const nisn=absensiCanonicalNisn_(r[cn]),st=String(r[cs]??'').trim().toUpperCase();if(!resultMap[nisn]||!Object.prototype.hasOwnProperty.call(resultMap[nisn],st))return;resultMap[nisn][st]++;
    });
  }
  const result=siswa.map(s=>{const z=resultMap[absensiCanonicalNisn_(s.nisn).trim()]||{H:0,S:0,I:0,A:0,T:0,D:0},total=z.H+z.S+z.I+z.A+z.T+z.D,had=z.H+z.T;return Object.assign({},s,z,{total:total,kehadiran:total?Number((had/total*100).toFixed(2)):0});});
  return cachePutJson_(ck,result,CACHE_TTL.DERIVED);
}

function getAbsensiTable_(dateStr,kelas){
  const key=formatDateKey_(parseDateInput_(dateStr)||new Date()), sh=getSheet_(APP.SHEETS.ABSENSI), lastRow=sh.getLastRow(), lastCol=sh.getLastColumn();
  if(lastRow<2||!lastCol)return [];
  const all=sh.getRange(1,1,lastRow,lastCol).getValues(), headers=all[0].map(dbNormalizeHeader_), map=Object.fromEntries(headers.map((x,i)=>[x,i]));
  const cId=dbFindColumn_(map,['ID','ID_Absensi']),cT=dbFindColumn_(map,['Tanggal','Tanggal_Absensi','Tanggal Absensi','Tgl']),cN=dbFindColumn_(map,['NISN','NIS','Nomor Induk Siswa Nasional']),cNm=dbFindColumn_(map,['Nama_Siswa','Nama Siswa','Nama']),cK=dbFindColumn_(map,['Kelas','Rombel','Kelas/Rombel']),cS=dbFindColumn_(map,['Status','Status Kehadiran','Kehadiran']),cKet=dbFindColumn_(map,['Keterangan','Catatan']),cU=dbFindColumn_(map,['User','Penginput','Dibuat_Oleh']),cY=dbFindColumn_(map,['Tahun_Pelajaran','Tahun Pelajaran']),cSem=dbFindColumn_(map,['Semester']);
  if(cT<0||cN<0||cS<0)return [];
  return all.slice(1).filter(r=>formatDateKey_(r[cT])===key&&(!kelas||cK<0||String(r[cK]).trim()===String(kelas).trim())).map(r=>({id:cId>=0?r[cId]:'',tanggal:formatTanggal_(r[cT]),nisn:cN>=0?absensiCanonicalNisn_(r[cN]):'',nama:cNm>=0?r[cNm]:'',kelas:cK>=0?r[cK]:'',status:cS>=0?r[cS]:'',keterangan:cKet>=0?r[cKet]:'',user:cU>=0?r[cU]:'',tahunPelajaran:cY>=0?r[cY]:'',semester:cSem>=0?r[cSem]:''}));
}


function getAbsensiHistory_(filters){
  filters=filters||{}; const sh=getSheet_(APP.SHEETS.ABSENSI),lr=sh.getLastRow(),lc=sh.getLastColumn();
  if(lr<2||!lc)return [];
  const all=sh.getRange(1,1,lr,lc).getValues(),h=all[0].map(dbNormalizeHeader_),map=Object.fromEntries(h.map((x,i)=>[x,i]));
  const ct=dbFindColumn_(map,['Tanggal','Tanggal_Absensi','Tanggal Absensi','Tgl']),cn=dbFindColumn_(map,['NISN','NIS']),cnm=dbFindColumn_(map,['Nama_Siswa','Nama Siswa','Nama']),ck=dbFindColumn_(map,['Kelas','Rombel','Kelas/Rombel']),cs=dbFindColumn_(map,['Status','Status Kehadiran','Kehadiran']),cket=dbFindColumn_(map,['Keterangan','Catatan']);
  if(ct<0||cn<0||cs<0)return [];
  const from=filters.from?formatDateKey_(parseDateInput_(filters.from)):''; const to=filters.to?formatDateKey_(parseDateInput_(filters.to)):''; const kelas=String(filters.kelas||'').trim(); const nisn=absensiCanonicalNisn_(filters.nisn||'');
  return all.slice(1).map(r=>{const d=parseDateInput_(r[ct]);return {tanggal:d?formatTanggal_(d):'',tanggalKey:d?formatDateKey_(d):'',nisn:canonicalNisn_(r[cn]),nama:cnm>=0?String(r[cnm]||''):'',kelas:ck>=0?String(r[ck]||''):'',status:String(r[cs]||''),keterangan:cket>=0?String(r[cket]||''):''};}).filter(x=>(!from||x.tanggalKey>=from)&&(!to||x.tanggalKey<=to)&&(!kelas||x.kelas===kelas)&&(!nisn||x.nisn===nisn)).sort((a,b)=>b.tanggalKey.localeCompare(a.tanggalKey)||a.nama.localeCompare(b.nama));
}
function getAbsensiSummary_(filters){
  filters=filters||{}; const rows=getAbsensiHistory_(filters), m={}; rows.forEach(x=>{if(!m[x.nisn])m[x.nisn]={nisn:x.nisn,nama:x.nama,kelas:x.kelas,H:0,S:0,I:0,A:0,T:0,D:0,total:0}; const st=x.status.toUpperCase();if(Object.prototype.hasOwnProperty.call(m[x.nisn],st))m[x.nisn][st]++;m[x.nisn].total++;});
  return Object.values(m).map(x=>Object.assign(x,{kehadiran:x.total?+(x.H+x.T)/x.total*100:0}));
}
function deleteAbsensi_(tanggal,nisn){
  const sh=getSheet_(APP.SHEETS.ABSENSI),lr=sh.getLastRow(),lc=sh.getLastColumn();if(lr<2)throw new Error('Data absensi kosong.');
  const all=sh.getRange(1,1,lr,lc).getValues(),h=all[0].map(dbNormalizeHeader_),map=Object.fromEntries(h.map((x,i)=>[x,i])),ct=dbFindColumn_(map,['Tanggal','Tanggal_Absensi','Tanggal Absensi','Tgl']),cn=dbFindColumn_(map,['NISN','NIS']);if(ct<0||cn<0)throw new Error('Header ABSENSI tidak lengkap.');
  const key=formatDateKey_(parseDateInput_(tanggal)); const row=all.findIndex((r,i)=>i>0&&formatDateKey_(r[ct])===key&&canonicalNisn_(r[cn]).trim()===canonicalNisn_(nisn));if(row<1)throw new Error('Data absensi tidak ditemukan.');sh.deleteRow(row+1);clearAppCache_();PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));logActivity_('HAPUS ABSENSI',String(nisn)+' - '+key);return{success:true};
}
