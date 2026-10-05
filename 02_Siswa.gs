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
      id:r[0]||generateID_('SIS'),nisn:String(r[nisnCol]).trim(),nama:String(r[namaCol]||'').trim(),kelas:kelasCol>=0?r[kelasCol]:'',jurusan:jurusanCol>=0?r[jurusanCol]:'',
      tempatLahir:idx('tempat_lahir')>=0?r[idx('tempat_lahir') ]:'',tanggalLahir:idx('tanggal_lahir')>=0?r[idx('tanggal_lahir') ]:'',orangTua:idx('nama_orang_tua')>=0?r[idx('nama_orang_tua') ]:'',hp:idx('nomor_hp_orang_tua')>=0?r[idx('nomor_hp_orang_tua') ]:'',alamat:idx('alamat')>=0?r[idx('alamat') ]:'',status:statusCol>=0?(r[statusCol]||'Aktif'):'Aktif',tahun:idx('tahun_pelajaran')>=0?r[idx('tahun_pelajaran') ]:getTahunPelajaran_()
    }));
  }else{
    out=v.slice(1).filter(r=>r[1]).map(r=>({id:r[0],nisn:String(r[1]),nama:r[2],kelas:r[3],jurusan:r[4],tempatLahir:r[5],tanggalLahir:r[6],orangTua:r[7],hp:r[8],alamat:r[9],status:r[10]||'Aktif',tahun:r[11]||getTahunPelajaran_()}));
  }
  return typeof cachePutJson_==='function'?cachePutJson_('STUDENTS_CACHE',out,300):out;
}
function getSiswaAktif_(){return getSiswa_().filter(s=>String(s.status||'Aktif').toLowerCase()!=='nonaktif');}
function getSiswaSemua_(){return getSiswa_();}
function saveSiswa_(obj){
  if(!obj||!obj.nisn||!obj.nama)throw new Error('NISN dan Nama wajib diisi.');
  const nisn=String(obj.nisn).trim();
  if(!/^\d{10}$/.test(nisn))throw new Error('NISN harus 10 digit angka.');
  const sh=getSheet_(APP.SHEETS.SISWA),lastCol=Math.max(sh.getLastColumn(),12),lastRow=sh.getLastRow();
  const fallback=['ID_Siswa','NISN','Nama_Siswa','Kelas','Jurusan','Tempat_Lahir','Tanggal_Lahir','Nama_Orang_Tua','Nomor_HP_Orang_Tua','Alamat','Status','Tahun_Pelajaran'];
  const headers=(lastCol?sh.getRange(1,1,1,lastCol).getDisplayValues()[0]:fallback).map(dbNormalizeHeader_);
  const col=k=>headers.indexOf(dbNormalizeHeader_(k));
  const ni=col('nisn')>=0?col('nisn'):1;
  const rowNo=findRowByColumnFast_(sh,ni+1,nisn,2);
  const cfg=getConfigObject_();
  const target=rowNo>0?sh.getRange(rowNo,1,1,lastCol).getValues()[0]:Array(lastCol).fill('');
  const set=(keys,val)=>{const list=Array.isArray(keys)?keys:[keys];for(const k of list){const i=col(k);if(i>=0){target[i]=val;return true;}}return false;};
  if(rowNo<1){const id=col('id_siswa')>=0?col('id_siswa'):col('id');if(id>=0&&!target[id])target[id]=generateID_('SIS');}
  set(['nisn'],nisn);
  set(['nama_siswa','nama'],String(obj.nama).trim());
  set(['kelas'],obj.kelas||cfg.Kelas||'');set(['jurusan'],obj.jurusan||cfg.Jurusan||'');
  set(['tempat_lahir'],obj.tempatLahir||'');set(['tanggal_lahir'],obj.tanggalLahir||'');set(['nama_orang_tua'],obj.orangTua||'');set(['nomor_hp_orang_tua'],obj.hp||'');set(['alamat'],obj.alamat||'');set(['status'],obj.status||'Aktif');set(['tahun_pelajaran'],obj.tahun||getTahunPelajaran_());
  if(rowNo>0)sh.getRange(rowNo,1,1,lastCol).setValues([target]);else sh.getRange(sh.getLastRow()+1,1,1,lastCol).setValues([target]);
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
function importSiswaRows_(rows){
  if(!Array.isArray(rows)||!rows.length)throw new Error('Tidak ada data untuk diimpor.');
  const TEMPLATE=['NISN','Nama_Siswa','Kelas','Jurusan','Tempat_Lahir','Tanggal_Lahir','Nama_Orang_Tua','Nomor_HP_Orang_Tua','Alamat','Status','Tahun_Pelajaran'];
  const DB_HEADERS=['ID_Siswa'].concat(TEMPLATE);
  const clean=v=>v==null?'':String(v).trim();
  const normalizeNisn=v=>clean(v).replace(/\.0$/,'');
  const normalizeDate=v=>{if(v instanceof Date&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');return clean(v);};

  // Validasi seluruh file SEBELUM data lama disentuh.
  const normalized=[];const seen={};
  rows.forEach((r,i)=>{
    const x={};TEMPLATE.forEach(k=>x[k]=r&&Object.prototype.hasOwnProperty.call(r,k)?r[k]:'');
    const nisn=normalizeNisn(x.NISN),nama=clean(x.Nama_Siswa);
    if(!/^\d{10}$/.test(nisn))throw new Error('Baris '+(i+2)+': NISN harus tepat 10 digit angka.');
    if(!nama)throw new Error('Baris '+(i+2)+': Nama_Siswa wajib diisi.');
    if(seen[nisn])throw new Error('Duplikat NISN dalam file pada baris '+(i+2)+': '+nisn);
    seen[nisn]=true;
    normalized.push({NISN:nisn,Nama_Siswa:nama,Kelas:clean(x.Kelas),Jurusan:clean(x.Jurusan),Tempat_Lahir:clean(x.Tempat_Lahir),Tanggal_Lahir:normalizeDate(x.Tanggal_Lahir),Nama_Orang_Tua:clean(x.Nama_Orang_Tua),Nomor_HP_Orang_Tua:clean(x.Nomor_HP_Orang_Tua),Alamat:clean(x.Alamat),Status:clean(x.Status)||'Aktif',Tahun_Pelajaran:clean(x.Tahun_Pelajaran)||getTahunPelajaran_()});
  });

  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try{
    const sh=getSheet_(APP.SHEETS.SISWA),lastCol=12;
    // Master DATA_SISWA bersifat CURRENT SNAPSHOT: upload baru menggantikan seluruh daftar siswa lama.
    // Data ABSENSI/PELANGGARAN/PENGHARGAAN tidak disentuh sehingga histori tetap dapat di-recall.
    if(sh.getMaxColumns()<lastCol)sh.insertColumnsAfter(sh.getMaxColumns(),lastCol-sh.getMaxColumns());
    sh.getRange(1,1,1,lastCol).setValues([DB_HEADERS]);
    const values=normalized.map(x=>[generateID_('SIS'),x.NISN,x.Nama_Siswa,x.Kelas,x.Jurusan,x.Tempat_Lahir,x.Tanggal_Lahir,x.Nama_Orang_Tua,x.Nomor_HP_Orang_Tua,x.Alamat,x.Status,x.Tahun_Pelajaran]);
    const oldRows=sh.getLastRow();
    if(oldRows>1)sh.getRange(2,1,oldRows-1,lastCol).clearContent();
    // NISN dan nomor HP HARUS diperlakukan sebagai teks sebelum data ditulis.
    // Ini mencegah Google Sheets mengubah 10 digit berawalan 0 menjadi angka.
    if(values.length){
      sh.getRange(2,2,values.length,1).setNumberFormat('@');
      sh.getRange(2,9,values.length,1).setNumberFormat('@');
      sh.getRange(2,1,values.length,lastCol).setValues(values);
    }
    SpreadsheetApp.flush();
    if(values.length){
      const check=sh.getRange(2,2,1,2).getDisplayValues()[0];
      if(String(check[0]||'').trim()!==String(values[0][1]||'').trim()) throw new Error('Import gagal diverifikasi: data pertama tidak tersimpan pada DATA_SISWA.');
    }
    // Hapus sisa baris lama secara fisik agar tidak ada data siswa lama yang masih terbaca.
    const totalRows=sh.getMaxRows(),needed=values.length+1;
    if(totalRows>needed)sh.deleteRows(needed+1,totalRows-needed);
    clearAppCache_();PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));
    logActivity_('IMPORT SISWA - REPLACE',normalized.length+' siswa aktif pada master terbaru. Data master lama digantikan; histori absensi/pelanggaran/penghargaan dipertahankan.');
    return{success:true,berhasil:normalized.length,dilewati:0,baru:normalized.length,diperbarui:0,diganti:oldRows>1?oldRows-1:0,mode:'REPLACE_ALL',template:TEMPLATE,message:'Data siswa lama pada DATA_SISWA telah diganti seluruhnya dengan data dari template. Histori absensi dan catatan lain tidak dihapus.'};
  }finally{lock.releaseLock();}
}
