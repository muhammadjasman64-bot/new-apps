/** 04_Pelanggaran.gs - master dropdown, input pelanggaran, poin otomatis */

function getMasterPelanggaran_() {
  const cached=cacheGetJson_('MASTER_PEL_CACHE');if(cached)return cached;
  const d=dbRead_(APP.SHEETS.MASTER_PELANGGARAN);
  const cKode=dbFindColumn_(d.map,['Kode','Kode Pelanggaran']);
  const cKat=dbFindColumn_(d.map,['Kategori']);
  const cNo=dbFindColumn_(d.map,['No','Nomor']);
  const cJenis=dbFindColumn_(d.map,['Jenis_Pelanggaran','Jenis Pelanggaran','Jenis']);
  const cPoin=dbFindColumn_(d.map,['Poin','Point']);
  const cAktif=dbFindColumn_(d.map,['Aktif','Status']);
  if(cKode<0||cJenis<0||cPoin<0) return [];
  const out=d.rows.filter(r=>{
    const active=cAktif<0?true:(r[cAktif]===true||String(r[cAktif]).toUpperCase()==='TRUE'||String(r[cAktif]).trim()==='');
    return active && String(r[cKode]||'').trim();
  }).map(r=>({kode:r[cKode],kategori:cKat>=0?r[cKat]:'',no:cNo>=0?r[cNo]:'',jenis:r[cJenis],poin:toNumberPoin_(r[cPoin])}));
  return cachePutJson_('MASTER_PEL_CACHE',out,CACHE_TTL.MASTER);
}
function getPelanggaranOptions_() {
  // Fungsi publik untuk dropdown frontend. Jangan diganti namanya karena index.html
  // memanggil fungsi ini langsung melalui google.script.run.
  const data = getMasterPelanggaran_();
  const categories = [...new Set(data.map(x => String(x.kategori || '').trim()).filter(Boolean))];
  return { success:true, data:data, categories:categories };
}

// Tes cepat untuk memastikan deployment mengenali fungsi publik ini.
function testGetPelanggaranOptions_() {
  const result = getPelanggaranOptions_();
  return { success:true, count:Array.isArray(result.data) ? result.data.length : 0, result:result };
}
function savePelanggaran_(obj) {
  if(!obj || !obj.nisn || !obj.kode) throw new Error('Siswa dan pelanggaran wajib dipilih.');
  const s=findStudentByNisn_(obj.nisn); if(!s) throw new Error('Siswa tidak ditemukan.');
  const master=getMasterPelanggaran_().find(x=>String(x.kode)===String(obj.kode));
  if(!master) throw new Error('Kode pelanggaran tidak ditemukan.');
  const cfg=getConfigObject_();
  const sh=getSheet_(APP.SHEETS.PELANGGARAN);
  // Simpan dahulu, lalu hitung total poin kumulatif siswa agar tindakan otomatis mengikuti tahapan pembinaan.
  const row=[generateID_('PEL'),parseDateInput_(obj.tanggal)||new Date(),s.nisn,s.nama,obj.kelas||s.kelas,
    master.kategori,master.kode,master.jenis,Number(master.poin),'','',
    Session.getActiveUser().getEmail()||'WebApp',cfg.Tahun_Pelajaran||getTahunPelajaran_(),cfg.Semester||getSemesterAktif_()];
  sh.getRange(sh.getLastRow()+1,1,1,row.length).setValues([row]);
  clearAppCache_();
  const total=getTotalPoin_(s.nisn);
  const keputusan=getKeputusanPoin_(total);
  const nr=sh.getLastRow();
  sh.getRange(nr,10,1,3).setValues([[Number(master.poin),keputusan.tindakan,'Total poin setelah pelanggaran: '+total+'. Status pembinaan: '+keputusan.status+'. '+keputusan.tindakan]]);
    syncTindakanUntukSiswa_(s.nisn);
  logActivity_('INPUT PELANGGARAN',s.nisn+' / '+master.kode);
  return {success:true,poin:master.poin};
}
function getPelanggaranSiswa_(nisn,start,end){
  const d=dbRead_(APP.SHEETS.PELANGGARAN),range=dateRangeInclusive_(start,end);
  const cn=dbFindColumn_(d.map,['NISN','NIS','Nomor Induk Siswa Nasional']),cd=dbFindColumn_(d.map,['Tanggal','Tanggal Pelanggaran','Tgl']),cid=dbFindColumn_(d.map,['ID_Pelanggaran','ID']),cc=dbFindColumn_(d.map,['Kategori']),ck=dbFindColumn_(d.map,['Kode','Kode Pelanggaran']),cj=dbFindColumn_(d.map,['Jenis_Pelanggaran','Jenis Pelanggaran','Jenis']),cp=dbFindColumn_(d.map,['Poin','Point']),ct=dbFindColumn_(d.map,['Tindakan']),cket=dbFindColumn_(d.map,['Keterangan','Catatan']);
  if(cn<0||cd<0||cp<0)return [];
  return d.rows.filter(r=>{const rv=canonicalNisn_(r[cn]),dt=dbDate_(r[cd]);return rv===canonicalNisn_(nisn)&&dt&&(!range.a||dt>=range.a)&&(!range.b||dt<=range.b);}).map(r=>({id:cid>=0?r[cid]:'',tanggal:formatTanggal_(r[cd]),kategori:cc>=0?r[cc]:'',kode:ck>=0?r[ck]:'',jenis:cj>=0?r[cj]:'',poin:toNumberPoin_(r[cp]),tindakan:ct>=0?r[ct]:'',keterangan:cket>=0?r[cket]:''}));
}

function getTotalPoin_(nisn,start,end){
  const d=dbRead_(APP.SHEETS.PELANGGARAN),cfg=getConfigObject_(),range=dateRangeInclusive_(start,end);
  const cn=dbFindColumn_(d.map,['NISN','NIS','Nomor Induk Siswa Nasional']),cd=dbFindColumn_(d.map,['Tanggal','Tanggal Pelanggaran','Tgl']),cp=dbFindColumn_(d.map,['Poin','Point']),ctp=dbFindColumn_(d.map,['Tahun_Pelajaran','Tahun Pelajaran','TP']);
  if(cn<0||cp<0)return 0;
  const activeTP=String(cfg.Tahun_Pelajaran||getTahunPelajaran_()).trim();
  let period=range;
  if(!period.a&&!period.b){
    const first=Number(activeTP.split('/')[0])||new Date().getFullYear();
    period={a:new Date(first,6,1,0,0,0,0),b:new Date(first+1,5,30,23,59,59,999)};
  }
  return d.rows.filter(r=>{
    const rv=canonicalNisn_(r[cn]); if(rv!==canonicalNisn_(nisn))return false;
    if(ctp>=0 && activeTP && String(r[ctp]??'').trim()!==activeTP)return false;
    const dt=cd>=0?dbDate_(r[cd]):null;
    return !period.a&&!period.b || (dt&&(!period.a||dt>=period.a)&&(!period.b||dt<=period.b));
  }).reduce((n,r)=>n+toNumberPoin_(r[cp]),0);
}


/** V27 - riwayat, edit, hapus, dan profil pelanggaran siswa. */
function getPelanggaranHistory_(filters){
  filters=filters||{};
  const d=dbRead_(APP.SHEETS.PELANGGARAN);
  const cn=dbFindColumn_(d.map,['NISN','NIS','Nomor Induk Siswa Nasional']), cd=dbFindColumn_(d.map,['Tanggal','Tanggal Pelanggaran','Tgl']), cns=dbFindColumn_(d.map,['Nama_Siswa','Nama Siswa','Nama']), ck=dbFindColumn_(d.map,['Kode_Pelanggaran','Kode','Kode Pelanggaran']), cj=dbFindColumn_(d.map,['Jenis_Pelanggaran','Jenis Pelanggaran','Jenis']), cp=dbFindColumn_(d.map,['Poin','Point']), ct=dbFindColumn_(d.map,['Tindakan']), ccat=dbFindColumn_(d.map,['Kategori']), cket=dbFindColumn_(d.map,['Keterangan','Catatan']), cid=dbFindColumn_(d.map,['ID_Pelanggaran','ID']), ckelas=dbFindColumn_(d.map,['Kelas']);
  if(cn<0||cd<0) return [];
  const range=dateRangeInclusive_(filters.from,filters.to), nisn=String(filters.nisn||'').trim(), kelas=String(filters.kelas||'').trim();
  return d.rows.map((r,i)=>({row:i+2,id:cid>=0?r[cid]:'',nisn:canonicalNisn_(r[cn]),nama:cns>=0?r[cns]:'',kelas:ckelas>=0?r[ckelas]:'',tanggal:formatTanggal_(r[cd]),tanggalKey:dbDate_(r[cd])?Utilities.formatDate(dbDate_(r[cd]),Session.getScriptTimeZone(),'yyyy-MM-dd'):'',kategori:ccat>=0?r[ccat]:'',kode:ck>=0?r[ck]:'',jenis:cj>=0?r[cj]:'',poin:cp>=0?toNumberPoin_(r[cp]):0,tindakan:ct>=0?r[ct]:'',keterangan:cket>=0?r[cket]:''}))
    .filter(x=>x.nisn.trim() && (!nisn||x.nisn.trim()===nisn) && (!kelas||String(x.kelas).trim()===kelas) && (!range.a||parseDateInput_(x.tanggalKey)>=range.a) && (!range.b||parseDateInput_(x.tanggalKey)<=range.b)).reverse();
}
function deletePelanggaran_(id){
  if(!id) throw new Error('ID pelanggaran wajib diisi.');
  const sh=getSheet_(APP.SHEETS.PELANGGARAN), d=dbRead_(APP.SHEETS.PELANGGARAN), ci=dbFindColumn_(d.map,['ID_Pelanggaran','ID']);
  if(ci<0) throw new Error('Kolom ID pelanggaran tidak ditemukan.');
  const idx=d.rows.findIndex(r=>String(r[ci]||'').trim()===String(id).trim());
  if(idx<0) throw new Error('Data pelanggaran tidak ditemukan.');
  const row=idx+2; const nisn=String(d.rows[idx][dbFindColumn_(d.map,['NISN','NIS','Nomor Induk Siswa Nasional'])]||'');
  sh.deleteRow(row); clearAppCache_(); PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now())); logActivity_('HAPUS PELANGGARAN',String(id)); return {success:true,nisn:nisn};
}
function updatePelanggaran_(obj){
  if(!obj||!obj.id||!obj.nisn||!obj.kode) throw new Error('ID, siswa, dan pelanggaran wajib diisi.');
  const s=findStudentByNisn_(obj.nisn); if(!s) throw new Error('Siswa tidak ditemukan.');
  const master=getMasterPelanggaran_().find(x=>String(x.kode)===String(obj.kode)); if(!master) throw new Error('Kode pelanggaran tidak ditemukan.');
  const sh=getSheet_(APP.SHEETS.PELANGGARAN), d=dbRead_(APP.SHEETS.PELANGGARAN), ci=dbFindColumn_(d.map,['ID_Pelanggaran','ID']);
  const idx=d.rows.findIndex(r=>String(r[ci]||'').trim()===String(obj.id).trim()); if(idx<0) throw new Error('Data pelanggaran tidak ditemukan.');
  const row=idx+2, h=d.map; const vals=sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];
  const set=(names,val)=>{const c=dbFindColumn_(h,names);if(c>=0)vals[c]=val;};
  set(['Tanggal','Tanggal Pelanggaran','Tgl'],parseDateInput_(obj.tanggal)||new Date()); set(['NISN','NIS','Nomor Induk Siswa Nasional'],s.nisn); set(['Nama_Siswa','Nama Siswa','Nama'],s.nama); set(['Kelas'],obj.kelas||s.kelas); set(['Kategori'],master.kategori); set(['Kode_Pelanggaran','Kode','Kode Pelanggaran'],master.kode); set(['Jenis_Pelanggaran','Jenis Pelanggaran','Jenis'],master.jenis); set(['Poin','Point'],Number(master.poin)); set(['Keterangan','Catatan'],obj.keterangan||'');
  sh.getRange(row,1,1,vals.length).setValues([vals]); clearAppCache_(); PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now())); logActivity_('EDIT PELANGGARAN',String(obj.id)); return {success:true,poin:Number(master.poin)};
}
function getProfilPelanggaran_(nisn){
  const s=findStudentByNisn_(nisn); if(!s) throw new Error('Siswa tidak ditemukan.');
  const total=getTotalPoin_(nisn), keputusan=getKeputusanPoin_(total), rows=getPelanggaranSiswa_(nisn);
  return {siswa:s,totalPoin:total,keputusan:keputusan,riwayat:rows};
}
