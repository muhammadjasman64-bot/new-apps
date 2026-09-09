/** 27_NisnIntegrity.gs - satu sumber kebenaran NISN dari DATA_SISWA. */
const NISN_SYNC_SHEETS_=[
  'USERS','ABSENSI','PELANGGARAN','PENGHARGAAN','TINDAKAN','PEMANGGILAN_ORANG_TUA',
  'CATATAN_WALI_KELAS','CATATAN_BK','REKAP_HARIAN','REKAP_BULANAN','REKAP_SEMESTER',
  'REKAP_TAHUNAN','PERINGATAN_DINI'
];
function buildCanonicalNisnMap_(){
  const map={}, bySignature={};
  getSiswa_().forEach(s=>{
    const n=String(s.nisn||'').trim(); if(!n)return;
    map[n]=n;
    if(/^\d{10}$/.test(n)){ const sig=n.replace(/^0+/,'')||'0'; if(!bySignature[sig])bySignature[sig]=[]; bySignature[sig].push(n); }
  });
  return {exact:map,signature:bySignature};
}
function canonicalFromMap_(value,map){
  const raw=String(value==null?'':value).trim().replace(/\.0$/,''); if(!raw)return '';
  if(map.exact[raw])return raw;
  if(/^\d+$/.test(raw)){
    const sig=raw.replace(/^0+/,'')||'0', hits=map.signature[sig]||[];
    if(hits.length===1)return hits[0];
    if(raw.length<10){const pad=raw.padStart(10,'0');if(map.exact[pad])return pad;}
  }
  return raw;
}
function syncNisnDatabase_(){
  const lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    const map=buildCanonicalNisnMap_();
    const result=[];
    NISN_SYNC_SHEETS_.forEach(name=>{
      const sh=getSS_().getSheetByName(name); if(!sh||sh.getLastRow()<2)return;
      const lastCol=sh.getLastColumn(), headers=sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
      const hmap={}; headers.forEach((h,i)=>hmap[dbNormalizeHeader_(h)]=i);
      const cn=dbFindColumn_(hmap,['NISN','NIS','Nomor Induk Siswa Nasional']); if(cn<0)return;
      const range=sh.getRange(2,cn+1,sh.getLastRow()-1,1), vals=range.getDisplayValues();
      let changed=0; const out=vals.map(r=>{const old=String(r[0]||'').trim(), neu=canonicalFromMap_(old,map); if(old!==neu)changed++; return [neu];});
      range.setNumberFormat('@'); range.setValues(out);
      if(changed)result.push({sheet:name,changed:changed});
    });
    const ds=getSS_().getSheetByName(APP.SHEETS.SISWA);
    if(ds&&ds.getLastRow()>=2){
      const h=ds.getRange(1,1,1,ds.getLastColumn()).getDisplayValues()[0], hm={};h.forEach((x,i)=>hm[dbNormalizeHeader_(x)]=i);
      const cn=dbFindColumn_(hm,['NISN']); if(cn>=0)ds.getRange(2,cn+1,ds.getLastRow()-1,1).setNumberFormat('@');
    }
    clearAppCache_(); PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));
    logActivity_('SINKRONISASI NISN','Semua sheet terkait diseragamkan mengikuti NISN master DATA_SISWA.');
    return {success:true,sheets:result,totalChanged:result.reduce((a,x)=>a+x.changed,0),message:'NISN seluruh database diseragamkan mengikuti DATA_SISWA.'};
  } finally {lock.releaseLock();}
}
function validateNisnIntegrity_(){
  const map=buildCanonicalNisnMap_(), issues=[];
  NISN_SYNC_SHEETS_.forEach(name=>{
    const sh=getSS_().getSheetByName(name);if(!sh||sh.getLastRow()<2)return;
    const h=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0],hm={};h.forEach((x,i)=>hm[dbNormalizeHeader_(x)]=i);
    const cn=dbFindColumn_(hm,['NISN','NIS','Nomor Induk Siswa Nasional']);if(cn<0)return;
    sh.getRange(2,cn+1,sh.getLastRow()-1,1).getDisplayValues().forEach((r,i)=>{const old=String(r[0]||'').trim(),neu=canonicalFromMap_(old,map);if(old!==neu)issues.push({sheet:name,row:i+2,nisn:old,canonical:neu});});
  });
  return {ok:issues.length===0,issues:issues,total:issues.length};
}
