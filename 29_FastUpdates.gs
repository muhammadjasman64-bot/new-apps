/** V46 - Fast update/delete helpers. */
function fastRowById_(sh, id, aliases) {
  const lc=sh.getLastColumn(); if(lc<1) return -1;
  const h=sh.getRange(1,1,1,lc).getDisplayValues()[0].map(dbNormalizeHeader_);
  const map=Object.fromEntries(h.map((x,i)=>[x,i]));
  const c=dbFindColumn_(map,aliases); if(c<0) return -1;
  return findRowByColumnFast_(sh,c+1,id,2);
}

function deletePelanggaranFast_(id){
  if(!id) throw new Error('ID pelanggaran wajib diisi.');
  const sh=getSheet_(APP.SHEETS.PELANGGARAN), row=fastRowById_(sh,id,['ID_Pelanggaran','ID']);
  if(row<2) throw new Error('Data pelanggaran tidak ditemukan.');
  const lc=sh.getLastColumn(), vals=sh.getRange(row,1,1,lc).getValues()[0], map={};
  sh.getRange(1,1,1,lc).getDisplayValues()[0].forEach((x,i)=>map[dbNormalizeHeader_(x)]=i);
  const nisn=map.nisn>=0?vals[map.nisn]:'';
  sh.deleteRow(row); dbVersionTouchFast_(); appendActivityFast_('HAPUS PELANGGARAN',String(id));
  return {success:true,nisn:nisn};
}

function updatePelanggaranFast_(obj){
  if(!obj||!obj.id) throw new Error('ID pelanggaran wajib diisi.');
  const sh=getSheet_(APP.SHEETS.PELANGGARAN),row=fastRowById_(sh,obj.id,['ID_Pelanggaran','ID']);
  if(row<2) throw new Error('Data pelanggaran tidak ditemukan.');
  const lc=sh.getLastColumn(), h=sh.getRange(1,1,1,lc).getDisplayValues()[0], map={};h.forEach((x,i)=>map[dbNormalizeHeader_(x)]=i);
  const vals=sh.getRange(row,1,1,lc).getValues()[0], s=findStudentByNisn_(obj.nisn||vals[map.nisn]); if(!s) throw new Error('Siswa tidak ditemukan.');
  const master=getMasterPelanggaran_().find(x=>String(x.kode)===String(obj.kode)); if(!master) throw new Error('Kode pelanggaran tidak ditemukan.');
  const set=(names,v)=>{const c=dbFindColumn_(map,names);if(c>=0)vals[c]=v;};
  set(['Tanggal','Tanggal Pelanggaran','Tgl'],parseDateInput_(obj.tanggal)||new Date());
  set(['NISN','NIS','Nomor Induk Siswa Nasional'],s.nisn);set(['Nama_Siswa','Nama Siswa','Nama'],s.nama);set(['Kelas'],obj.kelas||s.kelas);
  set(['Kategori'],master.kategori);set(['Kode_Pelanggaran','Kode','Kode Pelanggaran'],master.kode);set(['Jenis_Pelanggaran','Jenis Pelanggaran','Jenis'],master.jenis);set(['Poin','Point'],Number(master.poin));set(['Keterangan','Catatan'],obj.keterangan||'');
  sh.getRange(row,1,1,lc).setValues([vals]); dbVersionTouchFast_(); appendActivityFast_('EDIT PELANGGARAN',String(obj.id));
  return {success:true,poin:Number(master.poin)};
}

function deleteAbsensiFast_(tanggal,nisn){
  const sh=getSheet_(APP.SHEETS.ABSENSI),lc=sh.getLastColumn(),h=sh.getRange(1,1,1,lc).getDisplayValues()[0].map(dbNormalizeHeader_),map=Object.fromEntries(h.map((x,i)=>[x,i]));
  const ct=dbFindColumn_(map,['Tanggal','Tanggal_Absensi','Tanggal Absensi','Tgl']),cn=dbFindColumn_(map,['NISN','NIS']);if(ct<0||cn<0)throw new Error('Kolom absensi tidak ditemukan.');
  const targetDate=formatDateKey_(parseDateInput_(tanggal)), last=sh.getLastRow(); if(last<2)throw new Error('Data absensi kosong.');
  const dates=sh.getRange(2,ct+1,last-1,1).getValues(), nisns=sh.getRange(2,cn+1,last-1,1).getDisplayValues();let row=-1;
  const key=canonicalNisn_(nisn).trim(); for(let i=0;i<dates.length;i++){if(formatDateKey_(dates[i][0])===targetDate&&canonicalNisn_(nisns[i][0]).trim()===key){row=i+2;break;}}
  if(row<2)throw new Error('Data absensi tidak ditemukan.'); sh.deleteRow(row); dbVersionTouchFast_(); appendActivityFast_('HAPUS ABSENSI',String(nisn)+' - '+targetDate); return {success:true};
}

function updateTindakanStatusFast_(id,status,keterangan){
  if(!id)throw new Error('ID tindakan wajib diisi.');const sh=getSheet_(APP.SHEETS.TINDAKAN),row=fastRowById_(sh,id,['ID_Tindakan','ID']);if(row<2)throw new Error('Tindakan tidak ditemukan.');
  const lc=sh.getLastColumn(),h=sh.getRange(1,1,1,lc).getDisplayValues()[0].map(dbNormalizeHeader_),vals=sh.getRange(row,1,1,lc).getValues()[0];
  const cs=dbFindColumn_(Object.fromEntries(h.map((x,i)=>[x,i])),['Status']),ck=dbFindColumn_(Object.fromEntries(h.map((x,i)=>[x,i])),['Keterangan']);if(cs>=0)vals[cs]=status||'Selesai';if(ck>=0)vals[ck]=keterangan||'';
  sh.getRange(row,1,1,lc).setValues([vals]);dbVersionTouchFast_();appendActivityFast_('UBAH STATUS TINDAKAN',String(id));return {success:true};
}

function updatePemanggilanStatusFast_(id,status,hasil,tindakLanjut){
  if(!id)throw new Error('ID pemanggilan wajib diisi.');const sh=getSheet_(APP.SHEETS.PEMANGGILAN_ORANG_TUA),row=fastRowById_(sh,id,['ID_Pemanggilan','ID']);if(row<2)throw new Error('Pemanggilan tidak ditemukan.');
  const lc=sh.getLastColumn(),h=sh.getRange(1,1,1,lc).getDisplayValues()[0].map(dbNormalizeHeader_),map=Object.fromEntries(h.map((x,i)=>[x,i])),vals=sh.getRange(row,1,1,lc).getValues()[0];
  const set=(names,v)=>{const c=dbFindColumn_(map,names);if(c>=0)vals[c]=v;};set(['Status'],status||'Selesai');set(['Hasil'],hasil||'');set(['Tindak_Lanjut'],tindakLanjut||'');sh.getRange(row,1,1,lc).setValues([vals]);dbVersionTouchFast_();appendActivityFast_('UBAH STATUS PEMANGGILAN',String(id));return {success:true};
}
