/** 28_PerformanceCRUD.gs - CRUD stabil V49
 * Prinsip: write langsung ke Spreadsheet database yang dikunci, NISN selalu teks,
 * satu batch write bila memungkinkan, dan tidak ada reload database penuh setelah CRUD.
 */
function findRowByColumnFast_(sh,col,value,startRow){
  col=Number(col); if(col<1)return -1;
  const needle=String(value==null?'':value).trim(); if(!needle)return -1;
  const start=Number(startRow||2),last=sh.getLastRow(); if(last<start)return -1;
  const hit=sh.getRange(start,col,last-start+1,1).createTextFinder(needle).matchEntireCell(true).matchCase(false).findNext();
  return hit?hit.getRow():-1;
}
/** Reliable NISN lookup: compares normalized display values, so text/number storage is equivalent. */
function findRowByNisnFast_(sh,col,nisn,startRow){
  col=Number(col); if(col<1)return -1;
  const key=nisnText_(nisn); if(!/^\d{10}$/.test(key))return -1;
  const start=Number(startRow||2),last=sh.getLastRow(); if(last<start)return -1;
  const vals=sh.getRange(start,col,last-start+1,1).getDisplayValues();
  for(let i=0;i<vals.length;i++) if(nisnText_(vals[i][0])===key) return start+i;
  return -1;
}
function dbVersionTouchFast_(){
  try{if(typeof MASTER_PEMBINAAN_MEM_!=='undefined')MASTER_PEMBINAAN_MEM_=null;if(typeof DECISION_RULES_MEM_!=='undefined')DECISION_RULES_MEM_=null;}catch(e){}
  try{PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));}catch(e){}
  try{cacheRemove_('STUDENTS_CACHE');cacheRemove_('MASTER_NISN_MAP_V1');}catch(e){}
}
function appendActivityFast_(a,d){
  // Audit tetap ada, tetapi kegagalan logging tidak boleh membatalkan transaksi utama.
  try{const sh=getSheet_(APP.SHEETS.LOG);sh.getRange(sh.getLastRow()+1,1,1,4).setValues([[new Date(),Session.getActiveUser().getEmail()||'WebApp',a,d||'']]);}catch(e){}
}
function nisnText_(v){
  const s=String(v==null?'':v).trim().replace(/\.0$/,'');
  if(!/^\d+$/.test(s))return s;
  return s.padStart(10,'0');
}
function setNisnColumnText_(sh,col){if(col>=1&&sh.getMaxRows()>=2)sh.getRange(2,col,Math.max(sh.getMaxRows()-1,1),1).setNumberFormat('@');}
function invalidateStudentCaches_(){try{cacheRemove_('STUDENTS_CACHE');cacheRemove_('MASTER_NISN_MAP_V1');}catch(e){}try{PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));}catch(e){}}

function saveSiswaFast_(obj){
  if(!obj||!obj.nisn||!obj.nama)throw new Error('NISN dan Nama wajib diisi.');
  const nisn=nisnText_(obj.nisn); if(!/^\d{10}$/.test(nisn))throw new Error('NISN harus 10 digit angka.');
  const sh=getSheet_(APP.SHEETS.SISWA),lastCol=Math.max(sh.getLastColumn(),12),headers=sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(dbNormalizeHeader_);
  const col=k=>headers.indexOf(dbNormalizeHeader_(k)),ni=col('nisn')>=0?col('nisn'):1;
  const rowNo=findRowByNisnFast_(sh,ni+1,nisn,2),wasUpdate=rowNo>0,cfg=getConfigObject_();
  const target=wasUpdate?sh.getRange(rowNo,1,1,lastCol).getValues()[0]:Array(lastCol).fill('');
  const set=(keys,val)=>{for(const k of (Array.isArray(keys)?keys:[keys])){const i=col(k);if(i>=0){target[i]=val;return;}}};
  if(!wasUpdate){const id=col('id_siswa')>=0?col('id_siswa'):col('id');if(id>=0)target[id]=generateID_('SIS');}
  const nama=String(obj.nama).trim();
  set('nisn',nisn);set(['nama_siswa','nama'],nama);set('kelas',obj.kelas||cfg.Kelas||'');set('jurusan',obj.jurusan||cfg.Jurusan||'');
  set('tempat_lahir',obj.tempatLahir||'');set('tanggal_lahir',obj.tanggalLahir||'');set('nama_orang_tua',obj.orangTua||'');set('nomor_hp_orang_tua',obj.hp||'');set('alamat',obj.alamat||'');set('status',obj.status||'Aktif');set('tahun_pelajaran',obj.tahun||getTahunPelajaran_());
  const writeRow=wasUpdate?rowNo:sh.getLastRow()+1;
  // NISN ditulis sebagai string; format kolom ditetapkan saat setup, bukan setiap CRUD.
  sh.getRange(writeRow,1,1,lastCol).setValues([target]);
  SpreadsheetApp.flush();
  const verify=sh.getRange(writeRow,ni+1).getDisplayValue().trim();
  if(verify!==nisn)throw new Error('Data siswa tidak tersimpan. NISN setelah write tidak sesuai: '+verify);
  invalidateStudentCaches_();appendActivityFast_('SIMPAN SISWA',nisn+' - '+nama);
  return{success:true,updated:wasUpdate,message:wasUpdate?'Data siswa diperbarui.':'Data siswa ditambahkan.',student:{id:target[0],nisn,nama,kelas:obj.kelas||cfg.Kelas||'',jurusan:obj.jurusan||cfg.Jurusan||'',status:obj.status||'Aktif',tahun:obj.tahun||getTahunPelajaran_()}};
}
function savePelanggaranFast_(obj){
  if(!obj||!obj.nisn||!obj.kode)throw new Error('Siswa dan pelanggaran wajib dipilih.');
  const s=findStudentByNisn_(nisnText_(obj.nisn));if(!s)throw new Error('Siswa tidak ditemukan: '+obj.nisn);
  const master=getMasterPelanggaran_().find(x=>String(x.kode)===String(obj.kode));if(!master)throw new Error('Kode pelanggaran tidak ditemukan.');
  const cfg=getConfigObject_(),sh=getSheet_(APP.SHEETS.PELANGGARAN),oldTotal=getTotalPoin_(s.nisn),total=oldTotal+Number(master.poin||0),keputusan=getKeputusanPoin_(total),id=generateID_('PEL'),now=parseDateInput_(obj.tanggal)||new Date();
  const row=[id,now,s.nisn,s.nama,obj.kelas||s.kelas,master.kategori,master.kode,master.jenis,Number(master.poin),keputusan.tindakan,'Total poin setelah pelanggaran: '+total+'. Status pembinaan: '+keputusan.status+'. '+keputusan.tindakan,Session.getActiveUser().getEmail()||'WebApp',cfg.Tahun_Pelajaran||getTahunPelajaran_(),cfg.Semester||getSemesterAktif_()];
  const rowNo=sh.getLastRow()+1;const ncol=sh.getLastColumn();sh.getRange(2,3,Math.max(sh.getLastRow()-1,1),1).setNumberFormat('@');sh.getRange(rowNo,1,1,Math.min(row.length,ncol)).setValues([row.slice(0,ncol)]);SpreadsheetApp.flush();
  if(sh.getRange(rowNo,3).getDisplayValue().trim()!==s.nisn)throw new Error('Pelanggaran tidak tersimpan ke Spreadsheet.');
  // Tindakan adalah satu write tambahan yang memang diperlukan oleh fitur otomatis.
  const tin=getSheet_(APP.SHEETS.TINDAKAN),trow=[generateID_('TIN'),now,s.nisn,s.nama,total,keputusan.status,keputusan.tindakan,keputusan.surat||'',Session.getActiveUser().getEmail()||'WebApp','Otomatis berdasarkan sistem poin',cfg.Tahun_Pelajaran||getTahunPelajaran_(),cfg.Semester||getSemesterAktif_()];tin.getRange(2,3,Math.max(tin.getMaxRows()-1,1),1).setNumberFormat('@');tin.getRange(tin.getLastRow()+1,1,1,Math.min(trow.length,tin.getLastColumn())).setValues([trow.slice(0,tin.getLastColumn())]);
  dbVersionTouchFast_();appendActivityFast_('INPUT PELANGGARAN',s.nisn+' / '+master.kode);
  return{success:true,id,poin:Number(master.poin),totalPoin:total,keputusan,record:{id,tanggal:formatTanggal_(now),nisn:s.nisn,nama:s.nama,kelas:obj.kelas||s.kelas,kategori:master.kategori,kode:master.kode,jenis:master.jenis,poin:Number(master.poin),tindakan:keputusan.tindakan}};
}
function savePenghargaanFast_(obj){
  if(!obj||!obj.nisn||!obj.kode)throw new Error('Siswa dan penghargaan wajib dipilih.');
  const s=findStudentByNisn_(nisnText_(obj.nisn));if(!s)throw new Error('Siswa tidak ditemukan: '+obj.nisn);
  const m=getMasterPenghargaan_().find(x=>String(x.kode)===String(obj.kode));if(!m)throw new Error('Kode penghargaan tidak ditemukan.');
  const cfg=getConfigObject_(),sh=getSheet_(APP.SHEETS.PENGHARGAAN),now=parseDateInput_(obj.tanggal)||new Date(),row=[generateID_('REW'),now,s.nisn,s.nama,obj.kelas||s.kelas,m.jenis,m.prestasi,Number(m.poin),obj.keterangan||'',Session.getActiveUser().getEmail()||'WebApp',cfg.Tahun_Pelajaran||getTahunPelajaran_(),cfg.Semester||getSemesterAktif_()],rowNo=sh.getLastRow()+1;
  sh.getRange(2,3,Math.max(sh.getLastRow()-1,1),1).setNumberFormat('@');sh.getRange(rowNo,1,1,Math.min(row.length,sh.getLastColumn())).setValues([row.slice(0,sh.getLastColumn())]);SpreadsheetApp.flush();
  if(sh.getRange(rowNo,3).getDisplayValue().trim()!==s.nisn)throw new Error('Penghargaan tidak tersimpan ke Spreadsheet.');
  dbVersionTouchFast_();appendActivityFast_('INPUT PENGHARGAAN',s.nisn+' / '+m.kode);return{success:true,poin:Number(m.poin),record:{nisn:s.nisn,nama:s.nama,kode:m.kode,prestasi:m.prestasi,poin:Number(m.poin)}};
}
