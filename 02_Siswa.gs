/** 02_Siswa.gs - siswa dengan cache ringan. */
function getSiswa_(){
  const cached=typeof cacheGetJson_==='function'?cacheGetJson_('STUDENTS_CACHE'):null;
  if(cached) return cached;
  const sh=getSheet_(APP.SHEETS.SISWA), lastRow=sh.getLastRow(), lastCol=sh.getLastColumn();
  if(lastRow<2||lastCol<2)return [];
  const v=sh.getRange(1,1,lastRow,lastCol).getDisplayValues(), h=v[0].map(x=>String(x||'').trim().toLowerCase());
  const nisnCol=h.indexOf('nisn'),namaCol=h.indexOf('nama_siswa')>=0?h.indexOf('nama_siswa'):h.indexOf('nama'),kelasCol=h.indexOf('kelas'),jurusanCol=h.indexOf('jurusan'),statusCol=h.indexOf('status');
  let out=[];
  if(nisnCol>=0&&namaCol>=0){
    const idx=k=>h.indexOf(k);
    out=v.slice(1).filter(r=>String(r[nisnCol]||'').trim()).map(r=>({
      id:r[0]||generateID_('SIS'),nisn:normalizeNisnValue_(r[nisnCol]),nama:String(r[namaCol]||'').trim(),kelas:kelasCol>=0?r[kelasCol]:'',jurusan:jurusanCol>=0?r[jurusanCol]:'',
      tempatLahir:idx('tempat_lahir')>=0?r[idx('tempat_lahir') ]:'',tanggalLahir:idx('tanggal_lahir')>=0?r[idx('tanggal_lahir') ]:'',orangTua:idx('nama_orang_tua')>=0?r[idx('nama_orang_tua') ]:'',hp:idx('nomor_hp_orang_tua')>=0?r[idx('nomor_hp_orang_tua') ]:'',alamat:idx('alamat')>=0?r[idx('alamat') ]:'',status:statusCol>=0?(r[statusCol]||'Aktif'):'Aktif',tahun:idx('tahun_pelajaran')>=0?r[idx('tahun_pelajaran') ]:getTahunPelajaran_()
    }));
  }else{
    out=v.slice(1).filter(r=>r[1]).map(r=>({id:r[0],nisn:normalizeNisnValue_(r[1]),nama:r[2],kelas:r[3],jurusan:r[4],tempatLahir:r[5],tanggalLahir:r[6],orangTua:r[7],hp:r[8],alamat:r[9],status:r[10]||'Aktif',tahun:r[11]||getTahunPelajaran_()}));
  }
  return typeof cachePutJson_==='function'?cachePutJson_('STUDENTS_CACHE',out,300):out;
}
function getSiswaAktif_(){return getSiswa_().filter(s=>String(s.status||'Aktif').toLowerCase()!=='nonaktif');}
function getSiswaSemua_(){return getSiswa_();}
function saveSiswa_(obj){
  if(!obj||!obj.nisn||!obj.nama)throw new Error('NISN dan Nama wajib diisi.');
  const nisn=String(obj.nisn).trim().replace(/\.0$/,'').padStart(10,'0');
  if(!/^\d{10}$/.test(nisn))throw new Error('NISN harus 10 digit angka.');
  const sh=getSheet_(APP.SHEETS.SISWA),lastCol=Math.max(sh.getLastColumn(),12),lastRow=sh.getLastRow();
  const fallback=['ID_Siswa','NISN','Nama_Siswa','Kelas','Jurusan','Tempat_Lahir','Tanggal_Lahir','Nama_Orang_Tua','Nomor_HP_Orang_Tua','Alamat','Status','Tahun_Pelajaran'];
  const headers=(lastCol?sh.getRange(1,1,1,lastCol).getDisplayValues()[0]:fallback).map(dbNormalizeHeader_);
  const col=k=>headers.indexOf(dbNormalizeHeader_(k));
  const ni=col('nisn')>=0?col('nisn'):1;
  const rowNo=findRowByNisnFast_(sh,ni+1,nisn,2);
  const cfg=getConfigObject_();
  const target=rowNo>0?sh.getRange(rowNo,1,1,lastCol).getValues()[0]:Array(lastCol).fill('');
  const set=(keys,val)=>{const list=Array.isArray(keys)?keys:[keys];for(const k of list){const i=col(k);if(i>=0){target[i]=val;return true;}}return false;};
  if(rowNo<1){const id=col('id_siswa')>=0?col('id_siswa'):col('id');if(id>=0&&!target[id])target[id]=generateID_('SIS');}
  set(['nisn'],nisn);
  set(['nama_siswa','nama'],String(obj.nama).trim());
  set(['kelas'],obj.kelas||cfg.Kelas||'');set(['jurusan'],obj.jurusan||cfg.Jurusan||'');
  set(['tempat_lahir'],obj.tempatLahir||'');set(['tanggal_lahir'],obj.tanggalLahir||'');set(['nama_orang_tua'],obj.orangTua||'');set(['nomor_hp_orang_tua'],obj.hp||'');set(['alamat'],obj.alamat||'');set(['status'],obj.status||'Aktif');set(['tahun_pelajaran'],obj.tahun||getTahunPelajaran_());
  if(rowNo>0){sh.getRange(rowNo,1,1,lastCol).setValues([target]);}else{const nr=sh.getLastRow()+1;sh.getRange(nr,1,1,lastCol).setValues([target]);}
  // Invalidate only the student cache; do not flush every derived dashboard cache on a simple student CRUD.
  try{if(typeof cacheRemove_==='function')cacheRemove_('STUDENTS_CACHE');}catch(e){}
  try{PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));}catch(e){}
  appendActivityFast_('SIMPAN SISWA',nisn+' - '+obj.nama);
  return{success:true,updated:rowNo>0,message:rowNo>0?'Data siswa diperbarui.':'Data siswa ditambahkan.',student:{nisn:nisn,nama:String(obj.nama).trim(),kelas:obj.kelas||cfg.Kelas||'',jurusan:obj.jurusan||cfg.Jurusan||'',status:obj.status||'Aktif',tahun:obj.tahun||getTahunPelajaran_()}};
}
function setSiswaStatus_(nisn,status){
  const sh=getSheet_(APP.SHEETS.SISWA),lastRow=sh.getLastRow(),lastCol=sh.getLastColumn();if(lastRow<2)throw new Error('Data siswa kosong.');
  const data=sh.getRange(1,1,lastRow,lastCol).getDisplayValues(),h=data[0].map(x=>String(x||'').trim().toLowerCase()),ni=h.indexOf('nisn'),si=h.indexOf('status');
  const idx=data.findIndex((r,i)=>i>0&&String(r[ni>=0?ni:1]||'').trim()===String(nisn).trim());if(idx<1)throw new Error('Siswa tidak ditemukan.');if(si<0)throw new Error('Kolom Status tidak ditemukan.');
  sh.getRange(idx+1,si+1).setValue(status);clearAppCache_();PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));logActivity_(status==='Aktif'?'AKTIFKAN SISWA':'NONAKTIFKAN SISWA',String(nisn));return{success:true,status};
}
function deleteSiswa_(nisn){return setSiswaStatus_(nisn,'Nonaktif');}
function restoreSiswa_(nisn){return setSiswaStatus_(nisn,'Aktif');}
function importSiswaRows_(rows){ return importSiswaRowsCore_(rows,'admin',''); }
