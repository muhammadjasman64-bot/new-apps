/** 09_Laporan.gs - preview/PDF dengan periode eksplisit, grafik, detail */
function getReportData_(type,params){
  params=params||{};
  const cfg=getConfigObject_(), period=params.period||'bulanan';
  const data=getRekapForReport_(period,params);
  return {type:type,config:cfg,docs:getDocumentConfig_(),period:period,periodLabel:getPeriodLabel_(period,params),params:params,data:data,stats:getReportStatistics_(data),selected:params.nisn?getStudentReportDetail_(params.nisn,period,params):null};
}
function renderReportHtml_(type,params){const r=getReportData_(type,params); return type==='detail'?buildDetailReportHtml_(r):buildReportHtml_(r);}
function pdfRequestKey_(type,params){const raw=JSON.stringify({type:type||'rekap',params:params||{}});const digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,raw);return cacheKey_('PDF_RESULT',[Utilities.base64EncodeWebSafe(digest).replace(/=+$/,'')]);}
function buildPdfResult_(file){
  const id=file.getId(),url='https://drive.google.com/uc?export=download&id='+encodeURIComponent(id);
  return {success:true,name:file.getName(),url:file.getUrl(),downloadUrl:url,id:id,mimeType:'application/pdf',size:file.getSize()};
}
function buildPdfClientResult_(file){
  const meta=buildPdfResult_(file),blob=file.getBlob();
  if(blob.getContentType()!==MimeType.PDF)throw new Error('File PDF tidak valid.');
  const bytes=blob.getBytes();
  if(!bytes.length)throw new Error('File PDF kosong.');
  return {success:true,name:meta.name,id:meta.id,mimeType:MimeType.PDF,size:bytes.length,downloadUrl:meta.downloadUrl,dataBase64:Utilities.base64Encode(bytes)};
}
function createPdfFromReport_(type,params){
  const key=pdfRequestKey_(type,params),cached=cacheGetJson_(key);
  if(cached&&cached.id){try{const existing=DriveApp.getFileById(cached.id);if(existing.getMimeType()===MimeType.PDF&&existing.getSize()>0)return buildPdfClientResult_(existing);}catch(e){}}
  const r=getReportData_(type,params),html=type==='detail'?buildDetailReportHtml_(r):buildReportHtml_(r);
  const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss'),prefix=type==='detail'?'Laporan_Detail_Siswa_':'Laporan_Rekap_Kelas_';
  const blob=Utilities.newBlob(html,'text/html','laporan.html').getAs(MimeType.PDF).setName(prefix+stamp+'.pdf');
  if(blob.getContentType()!=='application/pdf')throw new Error('Gagal membuat PDF. File hasil konversi bukan PDF.');
  const bytes=blob.getBytes();if(!bytes.length)throw new Error('Gagal membuat PDF. File PDF kosong.');
  const file=getOrCreateFolder_().createFile(blob),result=buildPdfClientResult_(file);cachePutJson_(key,buildPdfResult_(file),300);return result;
}


function buildDetailReportHtml_(r){
  const c=r.config||{},d=r.docs||{},q=r.selected||{};
  const kop=fileImageDataUrl_(d.kop),logo=fileImageDataUrl_(d.logo),wali=fileImageDataUrl_(d.wali),waka=fileImageDataUrl_(d.waka),bk=fileImageDataUrl_(d.bk),kepsek=fileImageDataUrl_(d.kepsek);
  if(!q.siswa) return '<!doctype html><html><body><h3>Data siswa tidak ditemukan.</h3></body></html>';

  const head=kop
    ? '<div class="kopimg"><img src="'+kop+'" alt="Kop Surat"></div>'
    : '<div class="kopfallback">'+(logo?'<img class="logo" src="'+logo+'">':'')+
      '<div><b>'+escapeHtmlServer_(c.Nama_Sekolah||APP.SCHOOL)+'</b><br>'+
      escapeHtmlServer_(c.Alamat_Sekolah||'')+'<br>'+
      escapeHtmlServer_(c.Kode_Pos||'')+' | '+escapeHtmlServer_(c.Email_Sekolah||'')+'</div></div>';

  const a=q.absensi||{}, k=q.keputusan||{};
  const pelRows=(q.pelanggaran||[]).map(x=>
    '<tr><td>'+escapeHtmlServer_(x.tanggal)+'</td>'+
    '<td>'+escapeHtmlServer_(x.kode)+'</td>'+
    '<td class="left">'+escapeHtmlServer_(x.jenis)+'</td>'+
    '<td>'+escapeHtmlServer_(x.poin)+'</td>'+
    '<td class="left">'+escapeHtmlServer_(x.tindakan)+'</td>'+
    '<td class="left">'+escapeHtmlServer_(x.keterangan)+'</td></tr>').join('');
  const pel=pelRows || '<tr><td colspan="6" class="empty">Tidak ada pelanggaran pada periode ini.</td></tr>';

  const rewardRows=(q.penghargaan||[]).filter(x=>Number(x.poin||0)>0);
  const showRewardTable=rewardRows.length>0;
  const rew=rewardRows.map(x=>
    '<tr><td>'+escapeHtmlServer_(x.tanggal)+'</td>'+
    '<td class="left">'+escapeHtmlServer_(x.prestasi)+'</td>'+
    '<td>'+escapeHtmlServer_(x.poin)+'</td>'+
    '<td class="left">'+escapeHtmlServer_(x.keterangan)+'</td></tr>').join('');

  const sign='<div class="sign">'+
    '<div>Mengetahui,<br><b>Kepala Sekolah</b>'+
    (kepsek?'<img src="'+kepsek+'" class="ttd">':'<div class="ttdspace"></div>')+
    '<strong>'+escapeHtmlServer_(c.Nama_Kepala_Sekolah||'')+'</strong><br>NIP. '+escapeHtmlServer_(c.NIP_Kepala_Sekolah||'')+'</div>'+
    '<div><b>Waka Kesiswaan</b>'+
    (waka?'<img src="'+waka+'" class="ttd">':'<div class="ttdspace"></div>')+
    '<strong>'+escapeHtmlServer_(c.Nama_Waka_Kesiswaan||'')+'</strong><br>NIP. '+escapeHtmlServer_(c.NIP_Waka_Kesiswaan||'')+'</div>'+
    '<div><b>Guru BK</b>'+
    (bk?'<img src="'+bk+'" class="ttd">':'<div class="ttdspace"></div>')+
    '<strong>'+escapeHtmlServer_(c.Nama_Guru_BK||'')+'</strong></div>'+
    '<div><b>Wali Kelas</b>'+
    (wali?'<img src="'+wali+'" class="ttd">':'<div class="ttdspace"></div>')+
    '<strong>'+escapeHtmlServer_(c.Nama_Wali_Kelas||'')+'</strong><br>NIP. '+escapeHtmlServer_(c.NIP_Wali_Kelas||'')+'</div>'+
    '</div>';

  return '<!doctype html><html><head><meta charset="UTF-8"><style>'+
  '@page{size:A4 portrait;margin:9mm}'+
  '*{box-sizing:border-box}'+
  'html,body{margin:0;padding:0;background:#fff;color:#111;font-family:Arial,Helvetica,sans-serif;font-size:9pt;line-height:1.28}'+
  'body{width:192mm;min-height:279mm}'+
  '.kopimg{width:100%;text-align:center;margin:0 0 3mm 0;padding:0 0 2mm 0;border-bottom:2px solid #111}'+
  '.kopimg img{display:block;width:100%;height:auto;max-width:100%;margin:0 auto}'+
  '.kopfallback{min-height:24mm;border-bottom:2px double #111;padding:2mm 0;text-align:center;font-size:8.5pt;line-height:1.25}'+
  '.kopfallback .logo{width:15mm;height:15mm;object-fit:contain;float:left;margin-right:3mm}'+
  'h2{text-align:center;font-size:12pt;letter-spacing:.2px;margin:2mm 0 1mm}'+
  '.period{text-align:center;font-weight:700;font-size:9pt;margin-bottom:2.5mm}'+
  '.info{display:grid;grid-template-columns:1fr 1fr;gap:3mm;margin:0 0 2.5mm}'+
  '.box{border:1px solid #777;border-radius:1mm;padding:2mm;min-height:15mm}'+
  '.box b{display:inline-block;min-width:31mm}'+
  '.summary{display:grid;grid-template-columns:repeat(5,1fr);gap:2mm;margin:0 0 2.5mm}'+
  '.summary div{border:1px solid #777;text-align:center;padding:1.6mm 1mm;min-height:14mm;display:flex;flex-direction:column;justify-content:center}'+
  '.summary b{display:block;font-size:10pt;margin-bottom:.7mm}'+
  'table{width:100%;border-collapse:collapse;table-layout:fixed;margin:0 0 2.5mm}'+
  'th,td{border:1px solid #555;padding:1.15mm .9mm;vertical-align:middle;font-size:8pt;line-height:1.22;overflow-wrap:anywhere;word-break:normal}'+
  'th{background:#f1f1f1;font-weight:700;text-align:center}'+
  'td{text-align:center}'+
  '.left{text-align:left}'+
  '.empty{text-align:center;font-style:italic;padding:2.5mm}'+
  'h3{font-size:8.8pt;margin:2mm 0 1mm}'+
  '.follow{border:1px solid #777;padding:2mm;margin-top:1mm;line-height:1.25}'+
  '.sign{display:grid;grid-template-columns:repeat(4,1fr);gap:4mm;text-align:center;margin-top:5mm;font-size:7pt;page-break-inside:avoid}'+
  '.sign>div{min-height:31mm}'+
  '.ttd{display:block;width:30mm;height:17mm;object-fit:contain;margin:1.5mm auto}'+
  '.ttdspace{height:20mm}'+
  '.sign strong{font-size:7.2pt}'+
  '.pel-table col.date{width:12%}.pel-table col.code{width:7%}.pel-table col.type{width:27%}.pel-table col.point{width:6%}.pel-table col.action{width:23%}.pel-table col.note{width:25%}'+
  '.reward-table col.date{width:16%}.reward-table col.achievement{width:38%}.reward-table col.point{width:12%}.reward-table col.note{width:34%}'+
  '.att-table col{width:16.666%}'+
  '@media print{html,body{background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact}.no-print{display:none!important}img{break-inside:avoid}}'+
  '</style></head><body>'+head+
  '<h2>LAPORAN DETAIL SISWA</h2><div class="period">'+escapeHtmlServer_(q.periode)+'</div>'+
  '<div class="info"><div class="box"><b>Nama Siswa</b>: '+escapeHtmlServer_(q.siswa.nama)+'<br><b>NISN</b>: '+escapeHtmlServer_(q.siswa.nisn)+'</div>'+
  '<div class="box"><b>Kelas</b>: '+escapeHtmlServer_(q.siswa.kelas||c.Kelas||'-')+'<br><b>Jurusan</b>: '+escapeHtmlServer_(q.siswa.jurusan||c.Jurusan||'-')+'</div></div>'+
  '<div class="summary"><div><b>'+escapeHtmlServer_(a.kehadiran)+'%</b>Kehadiran</div>'+
  '<div><b>'+escapeHtmlServer_(q.poinPelanggaran)+'</b>Poin Pelanggaran</div>'+
  '<div><b>'+escapeHtmlServer_(q.poinPenghargaan)+'</b>Poin Penghargaan</div>'+
  '<div><b>'+escapeHtmlServer_(q.poinBersih)+'</b>Poin Bersih</div>'+
  '<div><b>'+escapeHtmlServer_(k.status||'NORMAL')+'</b>Status</div></div>'+
  '<table class="att-table"><colgroup><col><col><col><col><col><col></colgroup><tr><th>Hadir</th><th>Sakit</th><th>Izin</th><th>Alpa</th><th>Terlambat</th><th>Dispensasi</th></tr>'+
  '<tr><td>'+a.H+'</td><td>'+a.S+'</td><td>'+a.I+'</td><td>'+a.A+'</td><td>'+a.T+'</td><td>'+a.D+'</td></tr></table>'+
  '<h3>Riwayat Pelanggaran</h3>'+
  '<table class="pel-table"><colgroup><col class="date"><col class="code"><col class="type"><col class="point"><col class="action"><col class="note"></colgroup>'+
  '<thead><tr><th>Tanggal</th><th>Kode</th><th>Jenis Pelanggaran</th><th>Poin</th><th>Tindakan</th><th>Keterangan</th></tr></thead><tbody>'+pel+'</tbody></table>'+
  (showRewardTable?'<h3>Riwayat Penghargaan</h3><table class="reward-table"><colgroup><col class="date"><col class="achievement"><col class="point"><col class="note"></colgroup><thead><tr><th>Tanggal</th><th>Prestasi/Penghargaan</th><th>Poin</th><th>Keterangan</th></tr></thead><tbody>'+rew+'</tbody></table>':'')+
  '<div class="follow"><b>Tindak Lanjut:</b> '+escapeHtmlServer_(k.tindakan||'-')+'</div>'+sign+
  '</body></html>';
}


function buildReportHtml_(r){
  const c=r.config||{},d=r.docs||{},kop=fileImageDataUrl_(d.kop),logo=fileImageDataUrl_(d.logo),wali=fileImageDataUrl_(d.wali),waka=fileImageDataUrl_(d.waka),bk=fileImageDataUrl_(d.bk),kepsek=fileImageDataUrl_(d.kepsek);
  const head=kop
    ? '<div class="kopimg"><img src="'+kop+'" alt="Kop Surat"></div>'
    : '<div class="kopfallback">'+(logo?'<img class="logo" src="'+logo+'">':'')+
      '<div><b>'+escapeHtmlServer_(c.Nama_Sekolah||APP.SCHOOL)+'</b><br>'+
      escapeHtmlServer_(c.Alamat_Sekolah||'')+'<br>'+
      escapeHtmlServer_(c.Kode_Pos||'')+' | '+escapeHtmlServer_(c.Email_Sekolah||'')+'</div></div>';

  const s=r.stats||{};
  let body='<h2>REKAP JURNAL WALI KELAS</h2>'+
    '<div class="period">'+escapeHtmlServer_(r.periodLabel)+'</div>'+
    '<div class="subinfo"><b>Kelas:</b> '+escapeHtmlServer_(c.Kelas||'-')+
    ' &nbsp;&nbsp; <b>Jurusan:</b> '+escapeHtmlServer_(c.Jurusan||'-')+
    ' &nbsp;&nbsp; <b>Tahun Pelajaran:</b> '+escapeHtmlServer_(r.params&&r.params.tahunPelajaran||'-')+'</div>'+
    '<div class="summary"><div><b>'+escapeHtmlServer_(s.jumlahSiswa)+'</b><span>Jumlah Siswa</span></div>'+
    '<div><b>'+escapeHtmlServer_(s.rataKehadiran)+'%</b><span>Rata-rata Kehadiran</span></div>'+
    '<div><b>'+escapeHtmlServer_(s.totalPelanggaran)+'</b><span>Total Poin Pelanggaran</span></div>'+
    '<div><b>'+escapeHtmlServer_(s.totalPenghargaan)+'</b><span>Total Poin Penghargaan</span></div></div>';

  body+='<table class="rekap-table"><colgroup>'+
    '<col class="no"><col class="nisn"><col class="name"><col class="att"><col class="pel"><col class="rew"><col class="follow">'+
    '</colgroup><thead><tr><th>No</th><th>NISN</th><th>Nama Siswa</th><th>Persentase Kehadiran</th><th>Total Poin Pelanggaran</th><th>Poin Penghargaan</th><th>Tindak Lanjut</th></tr></thead><tbody>';

  (r.data||[]).forEach((x,i)=>{
    const lanjut=getKeputusanPoin_(Math.max(0,Number(x.poinBersih||0)));
    body+='<tr><td>'+ (i+1) +'</td>'+
      '<td>'+escapeHtmlServer_(x.nisn)+'</td>'+
      '<td class="left namecell">'+escapeHtmlServer_(x.nama)+'</td>'+
      '<td><b>'+escapeHtmlServer_(x.kehadiran)+'%</b></td>'+
      '<td>'+escapeHtmlServer_(x.poinPelanggaran)+'</td>'+
      '<td>'+escapeHtmlServer_(x.poinPenghargaan)+'</td>'+
      '<td class="left followcell"><b>'+escapeHtmlServer_(lanjut.status)+'</b><br>'+escapeHtmlServer_(lanjut.tindakan)+'</td></tr>';
  });
  if(!(r.data||[]).length) body+='<tr><td colspan="7" class="empty">Tidak ada data siswa pada periode yang dipilih.</td></tr>';
  body+='</tbody></table>';

  const sign='<div class="sign">'+
    '<div>Mengetahui,<br><b>Kepala Sekolah</b>'+
    (kepsek?'<img src="'+kepsek+'" class="ttd">':'<div class="ttdspace"></div>')+
    '<strong>'+escapeHtmlServer_(c.Nama_Kepala_Sekolah||'')+'</strong><br>NIP. '+escapeHtmlServer_(c.NIP_Kepala_Sekolah||'')+'</div>'+
    '<div><b>Waka Kesiswaan</b>'+
    (waka?'<img src="'+waka+'" class="ttd">':'<div class="ttdspace"></div>')+
    '<strong>'+escapeHtmlServer_(c.Nama_Waka_Kesiswaan||'')+'</strong><br>NIP. '+escapeHtmlServer_(c.NIP_Waka_Kesiswaan||'')+'</div>'+
    '<div><b>Guru BK</b>'+
    (bk?'<img src="'+bk+'" class="ttd">':'<div class="ttdspace"></div>')+
    '<strong>'+escapeHtmlServer_(c.Nama_Guru_BK||'')+'</strong></div>'+
    '<div><b>Wali Kelas</b>'+
    (wali?'<img src="'+wali+'" class="ttd">':'<div class="ttdspace"></div>')+
    '<strong>'+escapeHtmlServer_(c.Nama_Wali_Kelas||'')+'</strong><br>NIP. '+escapeHtmlServer_(c.NIP_Wali_Kelas||'')+'</div>'+
    '</div>';

  return '<!doctype html><html><head><meta charset="UTF-8"><style>'+
  '@page{size:A4 landscape;margin:7mm}'+
  '*{box-sizing:border-box}'+
  'html,body{margin:0;padding:0;background:#fff;color:#111;font-family:Arial,Helvetica,sans-serif;font-size:8pt;line-height:1.18}'+
  'body{width:283mm;min-height:192mm}'+
  '.kopimg{width:100%;text-align:center;margin:0 0 2.5mm;padding:0 0 1.5mm;border-bottom:2px solid #111}'+
  '.kopimg img{display:block;width:100%;height:auto;max-width:100%;margin:0 auto}'+
  '.kopfallback{min-height:21mm;border-bottom:2px double #111;padding:1.5mm 0;text-align:center;font-size:8pt;line-height:1.2}'+
  '.kopfallback .logo{width:14mm;height:14mm;object-fit:contain;float:left;margin-right:3mm}'+
  'h2{text-align:center;font-size:11pt;letter-spacing:.2px;margin:1.5mm 0 .5mm}'+
  '.period{text-align:center;font-weight:700;font-size:8pt;margin-bottom:1.2mm}'+
  '.subinfo{text-align:center;font-size:6.7pt;margin-bottom:1.8mm}'+
  '.summary{display:grid;grid-template-columns:repeat(4,1fr);gap:2mm;margin-bottom:2mm}'+
  '.summary div{border:1px solid #777;text-align:center;padding:1mm;min-height:10mm;display:flex;flex-direction:column;justify-content:center}'+
  '.summary b{display:block;font-size:9pt;margin-bottom:.5mm}.summary span{font-size:6pt}'+
  'table{width:100%;border-collapse:collapse;table-layout:fixed;margin:0}'+
  'th,td{border:1px solid #555;padding:.9mm .8mm;vertical-align:middle;overflow-wrap:anywhere;word-break:normal;font-size:7.1pt;line-height:1.18}'+
  'th{background:#f1f1f1;text-align:center;font-weight:700;white-space:normal}'+
  'td{text-align:center}'+
  '.left{text-align:left}.namecell{font-weight:600}.followcell{font-size:6.8pt}.followcell b{font-size:7.1pt}'+
  '.empty{text-align:center;padding:3mm;font-style:italic}'+
  '.rekap-table col.no{width:3.5%}.rekap-table col.nisn{width:14%}.rekap-table col.name{width:24%}.rekap-table col.att{width:14%}.rekap-table col.pel{width:9%}.rekap-table col.rew{width:9%}.rekap-table col.follow{width:26.5%}'+
  '.sign{display:grid;grid-template-columns:repeat(4,1fr);gap:5mm;text-align:center;margin-top:3.5mm;font-size:7pt;page-break-inside:avoid}'+
  '.sign>div{min-height:23mm}.ttd{display:block;width:28mm;height:13mm;object-fit:contain;margin:1mm auto}.ttdspace{height:15mm}.sign strong{font-size:7.1pt}'+
  '@media print{html,body{background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact} img{break-inside:avoid}}'+
  '</style></head><body>'+head+body+sign+'</body></html>';
}

function bar_(value,max,label){const pct=max?Math.max(0,Math.min(100,value/max*100)):0;return '<div class="barrow"><span>'+escapeHtmlServer_(label)+'</span><div class="bar_"><i style="width:'+pct+'%"></i></div><b>'+value+'</b></div>';}
