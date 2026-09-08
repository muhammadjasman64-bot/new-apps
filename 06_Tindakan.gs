/** 06_Tindakan.gs - mesin pembinaan sesuai dokumen sekolah */
function getKeputusanPoin_(totalPoin) {
  const p=Number(totalPoin)||0;
  if(p<5) return {status:'NORMAL',pemanggilan:0,surat:'',tindakan:'Belum masuk tahapan pembinaan berdasarkan sistem poin.'};
  const rules=getMasterPembinaan_();
  const r=rules.find(x=>p>=x.min&&p<=x.max);
  if(r) return {status:r.status,pemanggilan:r.pemanggilan,surat:r.dokumen,tindakan:r.tindakan,pihak:r.pihak,min:r.min,max:r.max};
  return {status:'MELEBIHI 100',pemanggilan:6,surat:'',tindakan:'Poin melebihi batas 100 pada tabel prosedur. Wajib diverifikasi oleh pihak sekolah.',pihak:'Kepala Sekolah'};
}
function getProfilPembinaan_(nisn) {
  const s=findStudentByNisn_(nisn); if(!s) throw new Error('Siswa tidak ditemukan.');
  const x=getPointSummaryMap_(null,null,getTahunPelajaran_())[String(nisn).trim()]||{pel:0,rew:0};
  const pel=toNumberPoin_(x.pel),rew=toNumberPoin_(x.rew),bersih=Math.max(0,pel-rew);
  return {siswa:s,poinPelanggaran:pel,poinPenghargaan:rew,poinBersih:bersih,keputusan:getKeputusanPoin_(bersih)};
}

function syncTindakanUntukSiswa_(nisn) {
  const p=getProfilPembinaan_(nisn), d=p.keputusan, cfg=getConfigObject_();
  const sh=getSheet_(APP.SHEETS.TINDAKAN);
  sh.appendRow([generateID_('TIN'),new Date(),p.siswa.nisn,p.siswa.nama,p.poinBersih,d.status,d.tindakan,'',
    Session.getActiveUser().getEmail()||'WebApp','Otomatis berdasarkan sistem poin',
    cfg.Tahun_Pelajaran||getTahunPelajaran_(),cfg.Semester||getSemesterAktif_()]);
  return p;
}
function saveTindakanManual_(){ throw new Error('Tindakan pembinaan ditentukan otomatis berdasarkan MASTER_PEMBINAAN; tidak dapat diinput manual.'); }
function savePemanggilan_(obj) {
  const s=findStudentByNisn_(obj.nisn); if(!s) throw new Error('Siswa tidak ditemukan.');
  const cfg=getConfigObject_(), p=getProfilPembinaan_(s.nisn);
  getSheet_(APP.SHEETS.PEMANGGILAN).appendRow([generateID_('PGL'),parseDateInput_(obj.tanggal)||new Date(),s.nisn,s.nama,
    p.poinBersih,Number(obj.ke)||p.keputusan.pemanggilan,obj.pihak||'',obj.hasil||'',obj.tindakLanjut||'',obj.status||'Selesai',
    cfg.Tahun_Pelajaran||getTahunPelajaran_(),cfg.Semester||getSemesterAktif_()]);
  return {success:true};
}


/** V28 - riwayat pembinaan, pemanggilan orang tua, status dan surat. */
function getTindakanHistory_(filters){
  filters=filters||{}; const d=dbRead_(APP.SHEETS.TINDAKAN);
  const h=d.map, cn=dbFindColumn_(h,['NISN','NIS']), cdt=dbFindColumn_(h,['Tanggal']), cnama=dbFindColumn_(h,['Nama_Siswa','Nama']), cp=dbFindColumn_(h,['Total_Poin','Poin']), cs=dbFindColumn_(h,['Status']), ct=dbFindColumn_(h,['Tindakan']), cd=dbFindColumn_(h,['Detail']), cpet=dbFindColumn_(h,['Petugas']), cket=dbFindColumn_(h,['Keterangan']), cid=dbFindColumn_(h,['ID_Tindakan','ID']);
  const range=dateRangeInclusive_(filters.from,filters.to), nisn=String(filters.nisn||'').trim();
  return d.rows.map((r,i)=>({row:i+2,id:cid>=0?r[cid]:'',tanggal:cdt>=0?formatTanggal_(r[cdt]):'',nisn:cn>=0?String(r[cn]||''):'',nama:cnama>=0?r[cnama]:'',totalPoin:cp>=0?toNumberPoin_(r[cp]):0,status:cs>=0?r[cs]:'',tindakan:ct>=0?r[ct]:'',detail:cd>=0?r[cd]:'',petugas:cpet>=0?r[cpet]:'',keterangan:cket>=0?r[cket]:''}))
   .filter(x=>(!nisn||x.nisn.trim()===nisn)&&(!range.a||parseDateInput_(x.tanggal)>=range.a)&&(!range.b||parseDateInput_(x.tanggal)<=range.b)).reverse();
}
function updateTindakanStatus_(id,status,keterangan){
  if(!id)throw new Error('ID tindakan wajib diisi.');
  const sh=getSheet_(APP.SHEETS.TINDAKAN),d=dbRead_(APP.SHEETS.TINDAKAN),ci=dbFindColumn_(d.map,['ID_Tindakan','ID']);
  const i=d.rows.findIndex(r=>String(r[ci]||'').trim()===String(id).trim()); if(i<0)throw new Error('Tindakan tidak ditemukan.');
  const row=i+2,h=d.map,vals=sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];
  const cs=dbFindColumn_(h,['Status']),ck=dbFindColumn_(h,['Keterangan']); if(cs>=0)vals[cs]=status||'Selesai'; if(ck>=0)vals[ck]=keterangan||'';
  sh.getRange(row,1,1,vals.length).setValues([vals]);clearAppCache_();PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));logActivity_('UBAH STATUS TINDAKAN',String(id));return {success:true};
}
function getPemanggilanHistory_(filters){
  filters=filters||{};const d=dbRead_(APP.SHEETS.PEMANGGILAN_ORANG_TUA),h=d.map;
  const cn=dbFindColumn_(h,['NISN','NIS']),cd=dbFindColumn_(h,['Tanggal']),cna=dbFindColumn_(h,['Nama_Siswa','Nama']),cp=dbFindColumn_(h,['Poin','Total_Poin']),cke=dbFindColumn_(h,['Pemanggilan_Ke']),cpp=dbFindColumn_(h,['Pihak_Pemanggil']),ch=dbFindColumn_(h,['Hasil']),ctl=dbFindColumn_(h,['Tindak_Lanjut']),cs=dbFindColumn_(h,['Status']),cid=dbFindColumn_(h,['ID_Pemanggilan','ID']);
  const range=dateRangeInclusive_(filters.from,filters.to),nisn=String(filters.nisn||'').trim();
  return d.rows.map((r,i)=>({row:i+2,id:cid>=0?r[cid]:'',tanggal:cd>=0?formatTanggal_(r[cd]):'',nisn:cn>=0?String(r[cn]||''):'',nama:cna>=0?r[cna]:'',poin:cp>=0?toNumberPoin_(r[cp]):0,ke:cke>=0?r[cke]:'',pihak:cpp>=0?r[cpp]:'',hasil:ch>=0?r[ch]:'',tindakLanjut:ctl>=0?r[ctl]:'',status:cs>=0?r[cs]:'Selesai'}))
   .filter(x=>(!nisn||x.nisn.trim()===nisn)&&(!range.a||parseDateInput_(x.tanggal)>=range.a)&&(!range.b||parseDateInput_(x.tanggal)<=range.b)).reverse();
}
function updatePemanggilanStatus_(id,status,hasil,tindakLanjut){
  if(!id)throw new Error('ID pemanggilan wajib diisi.');const sh=getSheet_(APP.SHEETS.PEMANGGILAN_ORANG_TUA),d=dbRead_(APP.SHEETS.PEMANGGILAN_ORANG_TUA),ci=dbFindColumn_(d.map,['ID_Pemanggilan','ID']);const i=d.rows.findIndex(r=>String(r[ci]||'').trim()===String(id).trim());if(i<0)throw new Error('Pemanggilan tidak ditemukan.');
  const row=i+2,h=d.map,vals=sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];const set=(names,v)=>{const c=dbFindColumn_(h,names);if(c>=0)vals[c]=v;};set(['Status'],status||'Selesai');set(['Hasil'],hasil||'');set(['Tindak_Lanjut'],tindakLanjut||'');sh.getRange(row,1,1,vals.length).setValues([vals]);clearAppCache_();PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));logActivity_('UBAH STATUS PEMANGGILAN',String(id));return {success:true};
}
function renderSuratPemanggilanHtml_(nisn,ke){
  const s=findStudentByNisn_(nisn);if(!s)throw new Error('Siswa tidak ditemukan.');const cfg=getConfigObject_(),docs=getDocumentConfig_(),kop=fileImageDataUrl_(docs.kop),wali=fileImageDataUrl_(docs.wali),kep=fileImageDataUrl_(docs.kepsek);
  const p=getProfilPembinaan_(nisn), nomor='421.5/'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy')+'/PGL/'+(ke||p.keputusan.pemanggilan||1);
  const head=kop?'<img style="width:100%;border-bottom:2px solid #111;padding-bottom:5px" src="'+kop+'">':'<div style="text-align:center;border-bottom:3px double #111;padding:8px"><b style="font-size:15px">'+escapeHtmlServer_(cfg.Nama_Sekolah||APP.SCHOOL)+'</b><br>'+escapeHtmlServer_(cfg.Alamat_Sekolah||'')+'</div>';
  return '<!doctype html><html><head><meta charset="UTF-8"><style>@page{size:A4 portrait;margin:20mm}body{font-family:Arial;font-size:11pt;line-height:1.6}h2{text-align:center;text-decoration:underline;font-size:14pt}.right{text-align:right}.sign{margin-top:45px;margin-left:60%;text-align:center}.ttd{height:55px;max-width:140px;object-fit:contain;display:block;margin:10px auto}</style></head><body>'+head+'<div style="margin-top:20px"><div class="right">Nomor: '+escapeHtmlServer_(nomor)+'<br>Lampiran: -<br>Perihal: <b>Pemanggilan Orang Tua/Wali</b></div><h2>SURAT PEMANGGILAN ORANG TUA/WALI</h2><p>Yth. Bapak/Ibu Orang Tua/Wali dari:</p><table><tr><td width="130">Nama</td><td>: '+escapeHtmlServer_(s.nama)+'</td></tr><tr><td>NISN</td><td>: '+escapeHtmlServer_(s.nisn)+'</td></tr><tr><td>Kelas</td><td>: '+escapeHtmlServer_(s.kelas)+'</td></tr></table><p>Dengan hormat, sehubungan dengan hasil pemantauan pembinaan peserta didik, kami mengharapkan kehadiran Bapak/Ibu untuk melakukan pembinaan bersama di sekolah. Saat ini akumulasi poin peserta didik adalah <b>'+escapeHtmlServer_(p.poinBersih)+'</b> poin dan berada pada tahap <b>'+escapeHtmlServer_(p.keputusan.status)+'</b>.</p><p>Pemanggilan ini merupakan pemanggilan ke-'+escapeHtmlServer_(ke||p.keputusan.pemanggilan||1)+'. Mohon hadir sesuai jadwal yang akan disampaikan oleh sekolah.</p><p>Demikian surat ini disampaikan. Atas perhatian dan kerja sama Bapak/Ibu, kami ucapkan terima kasih.</p><div class="sign">'+(cfg.Kota_Sekolah||'Buay Bahuga')+', '+formatTanggal_(new Date())+'<br>Wali Kelas'+(wali?'<img class="ttd" src="'+wali+'">':'<div style="height:65px"></div>')+'<b><u>'+escapeHtmlServer_(cfg.Nama_Wali_Kelas||'')+'</u></b><br>NIP. '+escapeHtmlServer_(cfg.NIP_Wali_Kelas||'')+'</div></div></body></html>';
}
function createPdfSuratPemanggilan_(nisn,ke){
  const key=pdfRequestKey_('surat',{nisn:String(nisn||''),ke:Number(ke||0)}),cached=cacheGetJson_(key);
  if(cached&&cached.id){try{const existing=DriveApp.getFileById(cached.id);if(existing.getMimeType()===MimeType.PDF&&existing.getSize()>0)return buildPdfClientResult_(existing);}catch(e){}}
  const html=renderSuratPemanggilanHtml_(nisn,ke),stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss'),blob=Utilities.newBlob(html,'text/html','surat.html').getAs(MimeType.PDF).setName('Surat_Pemanggilan_'+nisn+'_'+stamp+'.pdf');
  if(blob.getContentType()!=='application/pdf')throw new Error('Gagal membuat surat PDF. File hasil konversi bukan PDF.');
  const bytes=blob.getBytes();if(!bytes.length)throw new Error('Gagal membuat surat PDF. File PDF kosong.');
  const f=getOrCreateFolder_().createFile(blob),result=buildPdfClientResult_(f);cachePutJson_(key,buildPdfResult_(f),300);return result;
}
