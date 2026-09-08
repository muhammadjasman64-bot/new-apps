/** 02_Siswa.gs - siswa dengan cache ringan. */
function getSiswa_(){
  const cached=typeof cacheGetJson_==='function'?cacheGetJson_('STUDENTS_CACHE'):null;
  if(cached) return cached;
  const sh=getSheet_(APP.SHEETS.SISWA), lastRow=sh.getLastRow(), lastCol=sh.getLastColumn();
  if(lastRow<2||lastCol<2)return [];
  const v=sh.getRange(1,1,lastRow,lastCol).getValues(), h=v[0].map(x=>String(x||'').trim().toLowerCase());
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
  const data=lastRow>=2?sh.getRange(1,1,lastRow,lastCol).getValues():[['ID','NISN','Nama_Siswa','Kelas','Jurusan','Tempat_Lahir','Tanggal_Lahir','Nama_Orang_Tua','Nomor_HP_Orang_Tua','Alamat','Status','Tahun_Pelajaran']];
  const h=data[0].map(x=>String(x||'').trim().toLowerCase());
  const col=k=>h.indexOf(k.toLowerCase());
  const ni=col('nisn'), idx=data.findIndex((r,i)=>i>0&&String(r[ni>=0?ni:1]||'').trim()===nisn);
  const cfg=getConfigObject_();
  const row=idx>0?data[idx].slice(0,lastCol):Array(lastCol).fill('');
  const set=(key,val)=>{const i=col(key);if(i>=0)row[i]=val;};
  if(idx<1){const id=col('id');if(id>=0)row[id]=generateID_('SIS');}
  set('nisn',nisn);set('nama_siswa',String(obj.nama).trim());
  if(col('nama_siswa')<0)set('nama',String(obj.nama).trim());
  set('kelas',obj.kelas||cfg.Kelas||'');set('jurusan',obj.jurusan||cfg.Jurusan||'');
  set('tempat_lahir',obj.tempatLahir||'');set('tanggal_lahir',obj.tanggalLahir||'');set('nama_orang_tua',obj.orangTua||'');set('nomor_hp_orang_tua',obj.hp||'');set('alamat',obj.alamat||'');set('status',obj.status||'Aktif');set('tahun_pelajaran',obj.tahun||getTahunPelajaran_());
  if(idx>0)sh.getRange(idx+1,1,1,lastCol).setValues([row]);else sh.getRange(sh.getLastRow()+1,1,1,lastCol).setValues([row]);
  clearAppCache_();PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));logActivity_('SIMPAN SISWA',nisn+' - '+obj.nama);return{success:true,message:idx>0?'Data siswa diperbarui.':'Data siswa ditambahkan.'};
}
function setSiswaStatus_(nisn,status){
  const sh=getSheet_(APP.SHEETS.SISWA),lastRow=sh.getLastRow(),lastCol=sh.getLastColumn();if(lastRow<2)throw new Error('Data siswa kosong.');
  const data=sh.getRange(1,1,lastRow,lastCol).getValues(),h=data[0].map(x=>String(x||'').trim().toLowerCase()),ni=h.indexOf('nisn'),si=h.indexOf('status');
  const idx=data.findIndex((r,i)=>i>0&&String(r[ni>=0?ni:1]||'').trim()===String(nisn).trim());if(idx<1)throw new Error('Siswa tidak ditemukan.');if(si<0)throw new Error('Kolom Status tidak ditemukan.');
  sh.getRange(idx+1,si+1).setValue(status);clearAppCache_();PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));logActivity_(status==='Aktif'?'AKTIFKAN SISWA':'NONAKTIFKAN SISWA',String(nisn));return{success:true,status};
}
function deleteSiswa_(nisn){return setSiswaStatus_(nisn,'Nonaktif');}
function restoreSiswa_(nisn){return setSiswaStatus_(nisn,'Aktif');}
function importSiswaRows_(rows){
  if(!Array.isArray(rows)||!rows.length)throw new Error('Tidak ada data untuk diimpor.');
  const sh=getSheet_(APP.SHEETS.SISWA),cfg=getConfigObject_(),lastCol=Math.max(sh.getLastColumn(),12),lastRow=sh.getLastRow();
  const data=lastRow>=2?sh.getRange(1,1,lastRow,lastCol).getValues():[['ID','NISN','Nama_Siswa','Kelas','Jurusan','Tempat_Lahir','Tanggal_Lahir','Nama_Orang_Tua','Nomor_HP_Orang_Tua','Alamat','Status','Tahun_Pelajaran']],h=data[0].map(x=>String(x||'').trim().toLowerCase()),ni=h.indexOf('nisn');
  const dataRows=data.slice(1),byNisn={};dataRows.forEach((r,i)=>{const n=String(r[ni>=0?ni:1]||'').trim();if(n)byNisn[n]=i+2;});
  const clean=v=>v==null?'':String(v).trim(),normDate=v=>v instanceof Date?v:(clean(v)||'');
  const updates=[],adds=[];let ok=0,skip=0,updated=0,inserted=0;
  rows.forEach(r=>{const nisn=clean(r.NISN??r.nisn??r[0]),nama=clean(r.Nama_Siswa??r.Nama??r.nama??r[1]);if(!nisn||!nama){skip++;return;}const row=Array(lastCol).fill('');
    const set=(key,val)=>{const i=h.indexOf(key.toLowerCase());if(i>=0)row[i]=val;};const existingRow=byNisn[nisn]?data[byNisn[nisn]-2].slice(0,lastCol):null;if(existingRow)existingRow.forEach((v,i)=>row[i]=v);
    set('id',existingRow?row[h.indexOf('id')]:generateID_('SIS'));set('nisn',nisn);set('nama_siswa',nama);if(h.indexOf('nama_siswa')<0)set('nama',nama);set('kelas',clean(r.Kelas??r.kelas??r[2])||clean(cfg.Kelas));set('jurusan',clean(r.Jurusan??r.jurusan??r[3])||clean(cfg.Jurusan));set('tempat_lahir',clean(r.Tempat_Lahir??r.tempatLahir??r[4]));set('tanggal_lahir',normDate(r.Tanggal_Lahir??r.tanggalLahir??r[5]));set('nama_orang_tua',clean(r.Nama_Orang_Tua??r.orangTua??r[6]));set('nomor_hp_orang_tua',clean(r.Nomor_HP_Orang_Tua??r.hp??r[7]));set('alamat',clean(r.Alamat??r.alamat??r[8]));set('status',clean(r.Status??r.status??r[9])||'Aktif');set('tahun_pelajaran',clean(r.Tahun_Pelajaran??r.tahun??r[10])||getTahunPelajaran_());
    const rr=byNisn[nisn];if(rr){updates.push({row:rr,values:row});updated++;}else{adds.push(row);inserted++;}ok++;
  });updates.forEach(u=>sh.getRange(u.row,1,1,lastCol).setValues([u.values]));if(adds.length)sh.getRange(sh.getLastRow()+1,1,adds.length,lastCol).setValues(adds);clearAppCache_();PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));logActivity_('IMPORT SISWA',ok+' diproses, '+inserted+' baru, '+updated+' diperbarui, '+skip+' dilewati');return{success:true,berhasil:ok,dilewati:skip,baru:inserted,diperbarui:updated};
}
