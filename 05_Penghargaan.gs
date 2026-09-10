/** 05_Penghargaan.gs - master penghargaan + input penghargaan */

function getMasterPenghargaan_() {
  const cached=cacheGetJson_('MASTER_REWARD_CACHE');if(cached)return cached;
  const d=dbRead_(APP.SHEETS.MASTER_PENGHARGAAN);
  const cKode=dbFindColumn_(d.map,['Kode']);
  const cJenis=dbFindColumn_(d.map,['Jenis','Kategori']);
  const cPrest=dbFindColumn_(d.map,['Prestasi','Penghargaan','Jenis Penghargaan']);
  const cTingkat=dbFindColumn_(d.map,['Tingkat']);
  const cPoin=dbFindColumn_(d.map,['Poin','Point']);
  const cAktif=dbFindColumn_(d.map,['Aktif','Status']);
  if(cKode<0||cPoin<0) return [];
  const out=d.rows.filter(r=>{
    const active=cAktif<0?true:(r[cAktif]===true||String(r[cAktif]).toUpperCase()==='TRUE'||String(r[cAktif]).trim()==='');
    return active && String(r[cKode]||'').trim();
  }).map(r=>({kode:String(r[cKode]),jenis:cJenis>=0?String(r[cJenis]||''):'',prestasi:cPrest>=0?String(r[cPrest]||''):'',tingkat:cTingkat>=0?String(r[cTingkat]||''):'',poin:toNumberPoin_(r[cPoin])}));
  return cachePutJson_('MASTER_REWARD_CACHE',out,CACHE_TTL.MASTER);
}
function seedMasterPenghargaanOfficial_(){
  const sh=getSheet_(APP.SHEETS.MASTER_PENGHARGAAN);
  const data=[
['R01','TINDAKAN POSITIF','Berprestasi dalam kegiatan perlombaan dan pertandingan diluar sekolah','',10,true],['R02','TINDAKAN POSITIF','Menjadi wakil tingkat kabupaten Way kanan dalam hal kegiatan sekolah atau perlombaan','Kabupaten',25,true],['R03','TINDAKAN POSITIF','Menjadi wakil tingkat provinsi Lampung dalam hal kegiatan sekolah atau perlombaan','Provinsi',60,true],['R04','TINDAKAN POSITIF','Menjadi wakil tingkat Nasional dalam hal kegiatan sekolah atau perlombaan','Nasional',100,true],['R05','TINDAKAN POSITIF','Menjadi wakil tingkat Internasional dalam hal kegiatan sekolah atau perlombaan','Internasional',200,true],['R06','TINDAKAN POSITIF','Peringkat 1 sampai dengan 3 dikelas','Kelas',20,true],['R07','TINDAKAN POSITIF','Peringkat 4 sampai dengan 5 dikelas','Kelas',15,true],['R08','TINDAKAN POSITIF','Peringkat 6 sampai dengan 10 dikelas','Kelas',8,true],['R09','TINDAKAN POSITIF','Juara umum 1 sampai dengan 3 Jurusan','Jurusan',45,true],['R10','TINDAKAN POSITIF','Menjadi duta sekolah','Sekolah',5,true],
['P01','PRESTASI','Pemenang juara 1','Kabupaten',50,true],['P02','PRESTASI','Pemenang juara 2','Kabupaten',40,true],['P03','PRESTASI','Pemenang juara 3','Kabupaten',30,true],['P04','PRESTASI','Harapan','Kabupaten',20,true],['P05','PRESTASI','Pemenang juara 1','Provinsi',70,true],['P06','PRESTASI','Pemenang juara 2','Provinsi',60,true],['P07','PRESTASI','Pemenang juara 3','Provinsi',50,true],['P08','PRESTASI','Harapan','Provinsi',40,true],['P09','PRESTASI','Pemenang juara 1','Nasional',150,true],['P10','PRESTASI','Pemenang juara 2','Nasional',130,true],['P11','PRESTASI','Pemenang juara 3','Nasional',120,true],['P12','PRESTASI','Harapan','Nasional',100,true]];
  sh.clearContents();sh.getRange(1,1,1,6).setValues([['Kode','Jenis','Prestasi','Tingkat','Poin','Aktif']]);sh.getRange(2,1,data.length,6).setValues(data);sh.setFrozenRows(1);sh.autoResizeColumns(1,6);return data.length;
}

function getMasterPembinaan_(){
  const d=dbRead_(APP.SHEETS.MASTER_PEMBINAAN);
  return d.rows.map(r=>({kode:String(r.Kode||''),min:Number(r.Min_Poin)||0,max:Number(r.Max_Poin)||0,status:String(r.Status||''),pemanggilan:Number(r.Pemanggilan_Ke)||0,pihak:String(r.Pihak||''),tindakan:String(r.Tindakan||''),dokumen:String(r.Dokumen||'')})).filter(x=>x.kode);
}

function resetMasterPenghargaan_(){ return seedMasterPenghargaanOfficial_(); }
function savePenghargaan_(obj) {
  if(!obj||!obj.nisn||!obj.kode) throw new Error('Siswa dan penghargaan wajib dipilih.');
  const s=findStudentByNisn_(obj.nisn); if(!s) throw new Error('Siswa tidak ditemukan.');
  const m=getMasterPenghargaan_().find(x=>String(x.kode)===String(obj.kode)); if(!m) throw new Error('Kode penghargaan tidak ditemukan.');
  const cfg=getConfigObject_();
  const sh=getSheet_(APP.SHEETS.PENGHARGAAN);
  sh.getRange(sh.getLastRow()+1,1,1,12).setValues([[generateID_('REW'),parseDateInput_(obj.tanggal)||new Date(),s.nisn,s.nama,obj.kelas||s.kelas,m.jenis,m.prestasi,Number(m.poin),obj.keterangan||'',Session.getActiveUser().getEmail()||'WebApp',cfg.Tahun_Pelajaran||getTahunPelajaran_(),cfg.Semester||getSemesterAktif_()]]);
  commitDatabaseMutation_();
  syncTindakanUntukSiswa_(s.nisn); logActivity_('INPUT PENGHARGAAN',s.nisn+' / '+m.kode);
  return {success:true,poin:m.poin};
}

function getTotalPenghargaan_(nisn,start,end) {
  const d=dbRead_(APP.SHEETS.PENGHARGAAN),range=dateRangeInclusive_(start,end);
  const cNisn=dbFindColumn_(d.map,['NISN','NIS','Nomor Induk Siswa Nasional']);
  const cDate=dbFindColumn_(d.map,['Tanggal','Tanggal Penghargaan','Tgl']);
  const cPoin=dbFindColumn_(d.map,['Poin','Point']);
  if(cNisn<0||cDate<0||cPoin<0) return 0;
  return d.rows.filter(r=>{
    const rv=String(r[cNisn]??'').trim(),dt=dbDate_(r[cDate]);
    return rv===canonicalNisn_(nisn).trim() && dt && (!range.a||dt>=range.a) && (!range.b||dt<=range.b);
  }).reduce((sum,r)=>sum+toNumberPoin_(r[cPoin]),0);
}

function getAturanInfo_(){
  const c=getConfigObject_();
  return {sumber:c.Sumber_Aturan||'TATA TERTIB PESERTA DIDIK SMKN 01 BUAY BAHUGA fix.docx',nomor:c.Nomor_Keputusan_Aturan||'',tanggal:c.Tanggal_Aturan||'',versi:c.Versi_Aturan||'2026/2027',catatan:'Master pelanggaran, penghargaan dan pembinaan bersumber dari dokumen aturan sekolah; tidak disediakan input master manual pada Web App.'};
}
