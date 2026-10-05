/**
 * 28_PerformanceCRUD.gs - Fast CRUD V46
 *
 * Prinsip:
 * 1. Cari baris berdasarkan TextFinder, bukan membaca seluruh sheet.
 * 2. Batch write untuk transaksi multi-baris (terutama ABSENSI).
 * 3. Hindari read-after-write hanya untuk menghitung nilai yang sebenarnya sudah diketahui.
 * 4. Jalur CRUD tidak melakukan sinkronisasi/rekap berat secara synchronous.
 */
function findRowByColumnFast_(sh, col, value, startRow) {
  col = Number(col); if (col < 1) return -1;
  const needle = String(value == null ? '' : value).trim();
  if (!needle) return -1;
  const last = sh.getLastRow(); if (last < (startRow || 2)) return -1;
  const range = sh.getRange(startRow || 2, col, last - (startRow || 2) + 1, 1);
  const hit = range.createTextFinder(needle).matchEntireCell(true).matchCase(false).findNext();
  return hit ? hit.getRow() : -1;
}

function dbVersionTouchFast_() {
  // Hanya bump versi. cacheKey_ memakai DB_VERSION sehingga cache lama otomatis tidak terbaca.
  try { if (typeof MASTER_PEMBINAAN_MEM_ !== 'undefined') MASTER_PEMBINAAN_MEM_ = null; if (typeof DECISION_RULES_MEM_ !== 'undefined') DECISION_RULES_MEM_ = null; } catch(e) {}
  try { PropertiesService.getScriptProperties().setProperty('DB_VERSION', String(Date.now())); } catch(e) {}
}
function appendActivityFast_(a, d) {
  // Logging tidak boleh menggagalkan CRUD utama.
  try {
    const sh = getSheet_(APP.SHEETS.LOG);
    const row = [new Date(), Session.getActiveUser().getEmail() || 'WebApp', a, d || ''];
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  } catch (e) {}
}

function saveSiswaFast_(obj) {
  if (!obj || !obj.nisn || !obj.nama) throw new Error('NISN dan Nama wajib diisi.');
  const nisn=String(obj.nisn).trim(); if(!/^\d{10}$/.test(nisn)) throw new Error('NISN harus 10 digit angka.');
  const sh=getSheet_(APP.SHEETS.SISWA), lastCol=Math.max(sh.getLastColumn(),12), lastRow=sh.getLastRow();
  const headers=(lastCol?sh.getRange(1,1,1,lastCol).getDisplayValues()[0]:[]).map(dbNormalizeHeader_);
  const col=k=>headers.indexOf(dbNormalizeHeader_(k));
  const ni=col('nisn')>=0?col('nisn'):1;
  const rowNo=findRowByColumnFast_(sh,ni+1,nisn,2), cfg=getConfigObject_();
  const target=rowNo>0?sh.getRange(rowNo,1,1,lastCol).getValues()[0]:Array(lastCol).fill('');
  const set=(keys,val)=>{for(const k of (Array.isArray(keys)?keys:[keys])){const i=col(k);if(i>=0){target[i]=val;return;}}};
  if(rowNo<1){const id=col('id_siswa')>=0?col('id_siswa'):col('id');if(id>=0)target[id]=generateID_('SIS');}
  const nama=String(obj.nama).trim();
  set('nisn',nisn);set(['nama_siswa','nama'],nama);set('kelas',obj.kelas||cfg.Kelas||'');set('jurusan',obj.jurusan||cfg.Jurusan||'');
  set('tempat_lahir',obj.tempatLahir||'');set('tanggal_lahir',obj.tanggalLahir||'');set('nama_orang_tua',obj.orangTua||'');set('nomor_hp_orang_tua',obj.hp||'');set('alamat',obj.alamat||'');set('status',obj.status||'Aktif');set('tahun_pelajaran',obj.tahun||getTahunPelajaran_());
  if(rowNo>0){
    sh.getRange(rowNo,1,1,lastCol).setValues([target]);
  }else{
    rowNo=sh.getLastRow()+1;
    sh.getRange(rowNo,1,1,lastCol).setValues([target]);
  }
  SpreadsheetApp.flush();
  const verify=sh.getRange(rowNo,1,1,lastCol).getDisplayValues()[0];
  if(String(verify[ni]||'').trim()!==nisn) throw new Error('Data siswa gagal diverifikasi setelah penyimpanan ke Spreadsheet.');
  dbVersionTouchFast_(); appendActivityFast_('SIMPAN SISWA',nisn+' - '+nama);
  return {success:true,updated:rowNo>0,message:rowNo>0?'Data siswa diperbarui.':'Data siswa ditambahkan.',student:{nisn,nama,kelas:obj.kelas||cfg.Kelas||'',jurusan:obj.jurusan||cfg.Jurusan||'',status:obj.status||'Aktif',tahun:obj.tahun||getTahunPelajaran_()}};
}
function savePelanggaranFast_(obj) {
  if (!obj || !obj.nisn || !obj.kode) throw new Error('Siswa dan pelanggaran wajib dipilih.');
  const s = findStudentByNisn_(obj.nisn); if (!s) throw new Error('Siswa tidak ditemukan.');
  const master = getMasterPelanggaran_().find(x => String(x.kode) === String(obj.kode));
  if (!master) throw new Error('Kode pelanggaran tidak ditemukan.');
  const cfg = getConfigObject_(), sh = getSheet_(APP.SHEETS.PELANGGARAN);
  // Hitung total lama SEBELUM menulis. Setelah itu total baru cukup old + poin transaksi.
  const oldTotal = getTotalPoin_(s.nisn);
  const total = oldTotal + Number(master.poin || 0);
  const keputusan = getKeputusanPoin_(total);
  const id = generateID_('PEL');
  const row = [id, parseDateInput_(obj.tanggal) || new Date(), s.nisn, s.nama, obj.kelas || s.kelas,
    master.kategori, master.kode, master.jenis, Number(master.poin), keputusan.tindakan,
    'Total poin setelah pelanggaran: ' + total + '. Status pembinaan: ' + keputusan.status + '. ' + keputusan.tindakan,
    Session.getActiveUser().getEmail() || 'WebApp', cfg.Tahun_Pelajaran || getTahunPelajaran_(), cfg.Semester || getSemesterAktif_()];
  const pelRow=sh.getLastRow()+1;
  sh.getRange(pelRow,1,1,row.length).setValues([row]);
  SpreadsheetApp.flush();
  const pelCheck=sh.getRange(pelRow,1,1,3).getDisplayValues()[0];
  if(String(pelCheck[2]||'').trim()!==String(s.nisn).trim()) throw new Error('Pelanggaran gagal diverifikasi setelah penyimpanan.');

  // Tindakan otomatis ditulis langsung dari keputusan yang sudah tersedia.
  const tin = getSheet_(APP.SHEETS.TINDAKAN);
  tin.getRange(tin.getLastRow() + 1, 1, 1, 12).setValues([[generateID_('TIN'), new Date(), s.nisn, s.nama, total,
    keputusan.status, keputusan.tindakan, keputusan.surat || '', Session.getActiveUser().getEmail() || 'WebApp',
    'Otomatis berdasarkan sistem poin', cfg.Tahun_Pelajaran || getTahunPelajaran_(), cfg.Semester || getSemesterAktif_()]]);

  dbVersionTouchFast_();
  appendActivityFast_('INPUT PELANGGARAN', s.nisn + ' / ' + master.kode);
  return { success: true, id: id, poin: Number(master.poin), totalPoin: total, keputusan: keputusan,
    record: { id: id, tanggal: formatTanggal_(parseDateInput_(obj.tanggal) || new Date()), nisn: s.nisn, nama: s.nama, kelas: obj.kelas || s.kelas, kategori: master.kategori, kode: master.kode, jenis: master.jenis, poin: Number(master.poin), tindakan: keputusan.tindakan } };
}

function savePenghargaanFast_(obj) {
  if (!obj || !obj.nisn || !obj.kode) throw new Error('Siswa dan penghargaan wajib dipilih.');
  const s = findStudentByNisn_(obj.nisn); if (!s) throw new Error('Siswa tidak ditemukan.');
  const m = getMasterPenghargaan_().find(x => String(x.kode) === String(obj.kode)); if (!m) throw new Error('Kode penghargaan tidak ditemukan.');
  const cfg = getConfigObject_(), sh = getSheet_(APP.SHEETS.PENGHARGAAN), now = parseDateInput_(obj.tanggal) || new Date();
  const rewRow=sh.getLastRow()+1;
  sh.getRange(rewRow,1,1,12).setValues([[generateID_('REW'), now, s.nisn, s.nama, obj.kelas || s.kelas, m.jenis, m.prestasi, Number(m.poin), obj.keterangan || '', Session.getActiveUser().getEmail() || 'WebApp', cfg.Tahun_Pelajaran || getTahunPelajaran_(), cfg.Semester || getSemesterAktif_()]]);
  SpreadsheetApp.flush();
  const rewCheck=sh.getRange(rewRow,1,1,4).getDisplayValues()[0];
  if(String(rewCheck[2]||'').trim()!==String(s.nisn).trim()) throw new Error('Penghargaan gagal diverifikasi setelah penyimpanan.');
  dbVersionTouchFast_();
  appendActivityFast_('INPUT PENGHARGAAN', s.nisn + ' / ' + m.kode);
  return { success: true, poin: Number(m.poin), record: { nisn: s.nisn, nama: s.nama, kode: m.kode, prestasi: m.prestasi, poin: Number(m.poin) } };
}
