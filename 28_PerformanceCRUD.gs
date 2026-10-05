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
  // Satu titik invalidasi; jangan melakukan read database setelah write.
  try {
    if (typeof MASTER_PEMBINAAN_MEM_ !== 'undefined') MASTER_PEMBINAAN_MEM_ = null;
    if (typeof DECISION_RULES_MEM_ !== 'undefined') DECISION_RULES_MEM_ = null;
  } catch (e) {}
  try {
    CacheService.getScriptCache().removeAll([
      'APP_CONFIG_CACHE','MASTER_PEL_CACHE','MASTER_REWARD_CACHE','MASTER_PEMBINAAN_CACHE',
      'STUDENTS_CACHE','DASHBOARD_CACHE','DASHBOARD_DETAIL_CACHE','STATUS_ALERT_SUMMARY_CACHE',
      'ATTENDANCE_RISK_SUMMARY_CACHE','UNIFIED_RISK_SUMMARY_CACHE','EARLY_WARNING_SUMMARY_CACHE'
    ]);
  } catch (e) {}
  PropertiesService.getScriptProperties().setProperty('DB_VERSION', String(Date.now()));
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
  const nisn = String(obj.nisn).trim();
  if (!/^\d{10}$/.test(nisn)) throw new Error('NISN harus 10 digit angka.');
  const sh = getSheet_(APP.SHEETS.SISWA);
  const lastCol = Math.max(sh.getLastColumn(), 12);
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(x => String(x || '').trim().toLowerCase());
  const col = k => headers.indexOf(String(k).toLowerCase());
  const ni = col('nisn') >= 0 ? col('nisn') + 1 : 2;
  const rowNo = findRowByColumnFast_(sh, ni, nisn, 2);
  const cfg = getConfigObject_();
  const row = Array(lastCol).fill('');
  const set = (key, val) => { const i = col(key); if (i >= 0) row[i] = val; };
  const current = rowNo > 0 ? sh.getRange(rowNo, 1, 1, lastCol).getValues()[0] : row;
  const target = rowNo > 0 ? current : row;
  if (rowNo < 1) { const id = col('id_siswa') >= 0 ? col('id_siswa') : col('id'); if (id >= 0) target[id] = generateID_('SIS'); }
  set.call(null, 'nisn', nisn);
  const nama = String(obj.nama).trim();
  set.call(null, 'nama_siswa', nama); if (col('nama_siswa') < 0) set.call(null, 'nama', nama);
  set.call(null, 'kelas', obj.kelas || cfg.Kelas || '');
  set.call(null, 'jurusan', obj.jurusan || cfg.Jurusan || '');
  set.call(null, 'tempat_lahir', obj.tempatLahir || '');
  set.call(null, 'tanggal_lahir', obj.tanggalLahir || '');
  set.call(null, 'nama_orang_tua', obj.orangTua || '');
  set.call(null, 'nomor_hp_orang_tua', obj.hp || '');
  set.call(null, 'alamat', obj.alamat || '');
  set.call(null, 'status', obj.status || 'Aktif');
  set.call(null, 'tahun_pelajaran', obj.tahun || getTahunPelajaran_());
  if (rowNo > 0) sh.getRange(rowNo, 1, 1, lastCol).setValues([target]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([target]);
  dbVersionTouchFast_();
  appendActivityFast_('SIMPAN SISWA', nisn + ' - ' + nama);
  return { success: true, updated: rowNo > 0, message: rowNo > 0 ? 'Data siswa diperbarui.' : 'Data siswa ditambahkan.', student: { nisn: nisn, nama: nama, kelas: obj.kelas || cfg.Kelas || '', jurusan: obj.jurusan || cfg.Jurusan || '', status: obj.status || 'Aktif', tahun: obj.tahun || getTahunPelajaran_() } };
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
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);

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
  sh.getRange(sh.getLastRow() + 1, 1, 1, 12).setValues([[generateID_('REW'), now, s.nisn, s.nama, obj.kelas || s.kelas, m.jenis, m.prestasi, Number(m.poin), obj.keterangan || '', Session.getActiveUser().getEmail() || 'WebApp', cfg.Tahun_Pelajaran || getTahunPelajaran_(), cfg.Semester || getSemesterAktif_()]]);
  dbVersionTouchFast_();
  appendActivityFast_('INPUT PENGHARGAAN', s.nisn + ' / ' + m.kode);
  return { success: true, poin: Number(m.poin), record: { nisn: s.nisn, nama: s.nama, kode: m.kode, prestasi: m.prestasi, poin: Number(m.poin) } };
}
