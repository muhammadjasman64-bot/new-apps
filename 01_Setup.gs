/** 01_Setup.gs - DATABASE SETUP V3
 * Sumber master aturan: TATA TERTIB PESERTA DIDIK SMKN 01 BB fix.docx
 * Tahun Pelajaran 2026/2027.
 */
function setupDatabase_() {
  const ss=getSS_();
  const defs={
    USERS:['ID_User','Username','Password_Hash','Nama_Lengkap','Role','NISN','Kelas','Status','Created_At','Updated_At'],
    CONFIG:['Parameter','Nilai'],
    DATA_SISWA:['ID_Siswa','NISN','Nama_Siswa','Kelas','Jurusan','Tempat_Lahir','Tanggal_Lahir','Nama_Orang_Tua','Nomor_HP_Orang_Tua','Alamat','Status','Tahun_Pelajaran'],
    MASTER_PELANGGARAN:['Kode','Kategori','No','Jenis_Pelanggaran','Poin','Aktif'],
    MASTER_PENGHARGAAN:['Kode','Jenis','Prestasi','Tingkat','Poin','Aktif'],
    MASTER_PEMBINAAN:['Kode','Min_Poin','Max_Poin','Status','Pemanggilan_Ke','Pihak','Tindakan','Dokumen','Aktif'],
    PERINGATAN_DINI:['ID_Alert','Timestamp','NISN','Nama_Siswa','Tahun_Pelajaran','Semester','Severity','Jenis','Indeks_Sebelum','Indeks_Sesudah','Level_Sebelum','Level_Sesudah','Kehadiran_%','Alpa_Semester','Poin_Bersih','Rekomendasi','Status','Resolved_At','Resolved_By'],
    ABSENSI:['ID_Absensi','Tanggal','NISN','Nama_Siswa','Kelas','Status','Keterangan','Input_Oleh','Tahun_Pelajaran','Semester'],
    PELANGGARAN:['ID_Pelanggaran','Tanggal','NISN','Nama_Siswa','Kelas','Kategori','Kode_Pelanggaran','Jenis_Pelanggaran','Poin','Tindakan','Keterangan','Input_Oleh','Tahun_Pelajaran','Semester'],
    PENGHARGAAN:['ID_Penghargaan','Tanggal','NISN','Nama_Siswa','Kelas','Kategori','Prestasi','Poin','Keterangan','Input_Oleh','Tahun_Pelajaran','Semester'],
    TINDAKAN:['ID_Tindakan','Tanggal','NISN','Nama_Siswa','Total_Poin','Status','Tindakan','Detail','Petugas','Keterangan','Tahun_Pelajaran','Semester'],
    PEMANGGILAN_ORANG_TUA:['ID_Pemanggilan','Tanggal','NISN','Nama_Siswa','Poin','Pemanggilan_Ke','Pihak_Pemanggil','Hasil','Tindak_Lanjut','Status','Tahun_Pelajaran','Semester'],
    CATATAN_WALI_KELAS:['ID','Tanggal','NISN','Nama_Siswa','Catatan','Tindak_Lanjut','Input_Oleh','Tahun_Pelajaran','Semester'],
    CATATAN_BK:['ID','Tanggal','NISN','Nama_Siswa','Catatan','Tindak_Lanjut','Input_Oleh','Tahun_Pelajaran','Semester'],
    REKAP_HARIAN:['Tanggal','NISN','Nama_Siswa','Hadir','Sakit','Izin','Alpa','Terlambat','Dispensasi','Kehadiran_%','Poin_Pelanggaran','Poin_Penghargaan','Poin_Bersih'],
    REKAP_BULANAN:['Bulan','Tahun','NISN','Nama_Siswa','Hadir','Sakit','Izin','Alpa','Terlambat','Dispensasi','Kehadiran_%','Jml_Pelanggaran','Poin_Pelanggaran','Poin_Penghargaan','Poin_Bersih','Status_Pembinaan'],
    REKAP_SEMESTER:['Tahun_Pelajaran','Semester','NISN','Nama_Siswa','Hadir','Sakit','Izin','Alpa','Terlambat','Dispensasi','Kehadiran_%','Jml_Pelanggaran','Poin_Pelanggaran','Poin_Penghargaan','Poin_Bersih','Status_Pembinaan'],
    REKAP_TAHUNAN:['Tahun_Pelajaran','NISN','Nama_Siswa','Hadir','Sakit','Izin','Alpa','Terlambat','Dispensasi','Kehadiran_%','Jml_Pelanggaran','Poin_Pelanggaran','Poin_Penghargaan','Poin_Bersih','Status_Pembinaan'],
    DOKUMEN:['ID','Jenis_Dokumen','Nama_File','File_ID','URL','Tanggal_Upload','Aktif'],
    LOG_AKTIVITAS:['Timestamp','User','Aktivitas','Detail']
  };
  Object.keys(defs).forEach(k=>ensureSheetSchema_(ss,APP.SHEETS[k]||k,defs[k]));
  seedConfig_();
  seedMasterPelanggaranOfficial_();
  seedMasterPenghargaanOfficial_();
  seedMasterPembinaanOfficial_();
  if(typeof seedDefaultAdmin_==='function') seedDefaultAdmin_();
  if(typeof lockOfficialRuleSheets_==='function') lockOfficialRuleSheets_();
  return 'OK: database, CONFIG, dan master aturan resmi berhasil disiapkan dan dikunci.';
}

function ensureSheetSchema_(ss,name,headers){
  let sh=ss.getSheetByName(name);
  if(!sh) sh=ss.insertSheet(name);
  if(sh.getLastRow()===0){ sh.getRange(1,1,1,headers.length).setValues([headers]); }
  else if(sh.getLastRow()===1){
    const old=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0].slice(0,headers.length);
    if(old.join('|')!==headers.join('|')) sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  return sh;
}

function seedConfig_(){
  const sh=getSheet_(APP.SHEETS.CONFIG);
  const current=sh.getDataRange().getValues(), map={};
  current.slice(1).forEach(r=>{if(r[0])map[String(r[0])]=r[1];});
  const rows=Object.keys(APP.CONFIG_DEFAULTS).map(k=>[k,Object.prototype.hasOwnProperty.call(map,k)?map[k]:APP.CONFIG_DEFAULTS[k]]);
  sh.clearContents();
  sh.getRange(1,1,1,2).setValues([['Parameter','Nilai']]);
  sh.getRange(2,1,rows.length,2).setValues(rows);
  sh.setFrozenRows(1); sh.autoResizeColumns(1,2);
}

function seedMasterPelanggaranOfficial_(){
  const sh=getSheet_(APP.SHEETS.MASTER_PELANGGARAN);
  const d=[
['A01','A. KEHADIRAN DAN KEGIATAN BELAJAR',1,'Terlambat masuk sekolah lebih dari pukul 07.15',2,true],['A02','A. KEHADIRAN DAN KEGIATAN BELAJAR',2,'Terlambat masuk kelas setelah 15 menit pergantian jam pelajaran',2,true],['A03','A. KEHADIRAN DAN KEGIATAN BELAJAR',3,'Tidak masuk tanpa keterangan',4,true],['A04','A. KEHADIRAN DAN KEGIATAN BELAJAR',4,'Memalsukan surat untuk keperluan izin sekolah',6,true],['A05','A. KEHADIRAN DAN KEGIATAN BELAJAR',5,'Izin keluar kelas dan tidak Kembali lagi kekelas semula',4,true],['A06','A. KEHADIRAN DAN KEGIATAN BELAJAR',6,'Tidak mengumpulkan handphone pada tempatnya',5,true],['A07','A. KEHADIRAN DAN KEGIATAN BELAJAR',7,'Keluar kelas tanpa seizin guru mata pelajaran',3,true],['A08','A. KEHADIRAN DAN KEGIATAN BELAJAR',8,'Lompat pagar',5,true],['A09','A. KEHADIRAN DAN KEGIATAN BELAJAR',9,'Lompat jendela',4,true],['A10','A. KEHADIRAN DAN KEGIATAN BELAJAR',10,'Tidak mengikuti upacara dan senam setiap hari jumat',5,true],['A11','A. KEHADIRAN DAN KEGIATAN BELAJAR',11,'Mencontek pada pembuatan tugas, ulangan harian, ujian tengah semester dan ujian semester',10,true],
['B01','B. PAKAIAN/SERAGAM SEKOLAH',1,'Memakai baju seragam tidak sesuai ketentuan sekolah',5,true],['B02','B. PAKAIAN/SERAGAM SEKOLAH',2,'Memakai sepatu tidak sesuai ketentuan sekolah',2,true],['B03','B. PAKAIAN/SERAGAM SEKOLAH',3,'Memakai kaos kaki tidak sesuai ketentuan sekolah',2,true],['B04','B. PAKAIAN/SERAGAM SEKOLAH',4,'Tidak memakai kaos kaki',3,true],['B05','B. PAKAIAN/SERAGAM SEKOLAH',5,'Memakai topi dan jaket bukan ketentuan sekolah dilingkungan sekolah',5,true],['B06','B. PAKAIAN/SERAGAM SEKOLAH',6,'Tidak memakai bed atau identitas sekolah',5,true],['B07','B. PAKAIAN/SERAGAM SEKOLAH',7,'Tidak memakai dasi atau topi sekolah',5,true],
['C01','C. KEPRIBADIAN',1,'Berhias dan atau mengenakan perhiasan berlebihan (LIPSTIK,LIP GLOSS,PEMERAH BIBIR,ALIS UKIR,JARI DIKUTEK/CAT DAN MAKE UP PEMUTIH MUKA.',5,true],['C02','C. KEPRIBADIAN',2,'Rambut Panjang melebihi ketentuan sekolah bagi laki-laki',5,true],['C03','C. KEPRIBADIAN',3,'Menggunakan perhiasan gelang,kalung,anting,bertindik,dan cincin',5,true],['C04','C. KEPRIBADIAN',4,'Kuku panjang',2,true],['C05','C. KEPRIBADIAN',5,'Potongan rambut tidak sesuai ketentuan sekolah (MOHAX)',3,true],['C06','C. KEPRIBADIAN',6,'Peserta didik mewarnai rambut/pirang',4,true],['C07','C. KEPRIBADIAN',7,'Mengeluarkan kata-kata yang bersifat merendahkan,menghina,memprovokasi,atau menyebarkan kebencian serta berita bohong (Hoax) melalui media sosial yang mengatas namakan sekolah secara lisan ataupun tulisan',30,true],['C08','C. KEPRIBADIAN',8,'Mengeluarkan kata-kata tidak sopan,mengejek atau menghina kepada KEPALA SEKOLAH,GURU,KARYAWAN dan TAMU SEKOLAH yang dilakukan baik secara langsung maupun tidak langsung.',35,true],['C09','C. KEPRIBADIAN',9,'Melawan,mengancam,dan membahayakn kepala sekolah,guru,dan tenaga kependidikan.',75,true],['C10','C. KEPRIBADIAN',10,'Mengeluarkan kata kata kotor atau umpatan tidak sopan kepada teman',5,true],['C11','C. KEPRIBADIAN',11,'Melakukan pemerasan atau pemalakan terhadap peserta didik atau pihak warga sekolah lainnya',25,true],['C12','C. KEPRIBADIAN',12,'Melakukan pencurian milik sekolah ataupun teman',75,true],['C13','C. KEPRIBADIAN',13,'Sengaja memainkan knalpot sepeda motor dengan suara memekakkan telinga dilingkunga sekolah',5,true],['C14A','C. KEPRIBADIAN',14,'Terbukti berbuat asusila: Peserta didik bermesraan dengan lawan jenis atau sesama jenis dilingkungan sekolah (berciuman dan atau berpelukan.)',40,true],['C14B','C. KEPRIBADIAN',14,'Terbukti berbuat asusila: Berbuat zina (berhubungan badan)',100,true],['C15','C. KEPRIBADIAN',15,'Menikah',100,true],
['D01','D. ROKOK KONVENSIONAL (HISAP) DAN ROKOK ELEKTRIK',1,'Membawa rokok hisap dan rokok elektrik kesekolah',25,true],['D02','D. ROKOK KONVENSIONAL (HISAP) DAN ROKOK ELEKTRIK',2,'Menghisap rokok dan rokok elektrik didalam lingkungan sekolah',30,true],['D03','D. ROKOK KONVENSIONAL (HISAP) DAN ROKOK ELEKTRIK',3,'Menghisap rokok dan rokok elektrik diluar lingkungan sekolah dengan menggunakan atribut SMKN 1 BUAY BAHUGA',10,true],['D04','D. ROKOK KONVENSIONAL (HISAP) DAN ROKOK ELEKTRIK',4,'Membagikan atau menjual rokok hisap atau rokok elektrik dilingkungan sekolah',30,true],['D05','D. ROKOK KONVENSIONAL (HISAP) DAN ROKOK ELEKTRIK',5,'Berkumpul Bersama-sama dengan peserta didik yang menghisap rokok konvensional atau elektrik dilingkungan sekolah.',8,true],
['E01','E. BERJUDI',1,'Membawa alat-alat judi dan sejenisnya (Kartu remi,uno,dll) dilingkungan sekolah',20,true],['E02','E. BERJUDI',2,'Melakukan permainan menggunakan alat judi dilingkungan sekolah',30,true],['E03','E. BERJUDI',3,'Melakukan perjudian/taruhan menggunakan uang atau barang dan sejenisnya',40,true],
['F01','F. BACAAN,GAMBAR DAN FILM PORNO',1,'Membawa buku,surat kabar,Dvd/Vcd,flasdisk,kartu memori yang bersifat porno',20,true],['F02','F. BACAAN,GAMBAR DAN FILM PORNO',2,'Memperjualbelikan,meminjam buku,surat kabar,DVD,VCD,flashdisk,kartu memori atau gawai yang bersifat porno',30,true],['F03','F. BACAAN,GAMBAR DAN FILM PORNO',3,'Menonton/melihat buku/majalah/surat kabar/tabloid/kaset vcd,dvd,flashdisk atau gawai yang bersifat porno',25,true],['F04','F. BACAAN,GAMBAR DAN FILM PORNO',4,'Merekam,mengambil gambar dan mempublikasikan kegiatan bersifat porno',75,true],
['G01','G. NARKOBA DAN MINUMAN KERAS',1,'Membawa dan mengkonsumsi minuman beralkohol,narkotika,psikotropika dan zat adiktif lainnya yang dilarang undang-undang.',100,true],['G02','G. NARKOBA DAN MINUMAN KERAS',2,'Menjual MIRAS,atau beralkohol,narkotika,psikotropika,dan zat adiktif lainya yang dilarang undang-undang.',100,true],['G03','G. NARKOBA DAN MINUMAN KERAS',3,'Berkumpul Bersama teman yang sedang mengkonsumsi minuman beralkohol,narkotika,psikotropika dan zat adiktif lainnya yang dilarang undang-undang.',45,true],['G04','G. NARKOBA DAN MINUMAN KERAS',4,'Datang kesekolah dalam keadaan mabuk atau sakaw (ngefly)',75,true],
['H01','H. SENJATA TAJAM (SAJAM) DAN SENJATA API (SENPI)',1,'Membawa senjata tajam dan senjata api dilingkungan sekolah',25,true],['H02','H. SENJATA TAJAM (SAJAM) DAN SENJATA API (SENPI)',2,'Memperjualbelikan senjata tajam dan senjata api dilingkungan sekolah',50,true],['H03','H. SENJATA TAJAM (SAJAM) DAN SENJATA API (SENPI)',3,'Menggunakan senjata tajam dan senjata api dilingkungan sekolah untuk mengancam/melukai orang',75,true],['H04','H. SENJATA TAJAM (SAJAM) DAN SENJATA API (SENPI)',4,'Membawa atau memakai ikat pinggang berkepala Gear',20,true],
['I01','I. BERKELAHI/TAWURAN',1,'Menjadi pemicu perkelahian atau keributan (PROVOKATOR)',40,true],['I02','I. BERKELAHI/TAWURAN',2,'Perkelahian antar peserta didik/kelas/kelompok sehingga menimbulkan keresahan disekolah dan dimasyarakat baik perkelahian fisik maupun non fisik.',50,true],['I03','I. BERKELAHI/TAWURAN',3,'Perkelahian dengan guru dan pegawai sekolah',100,true],['I04','I. BERKELAHI/TAWURAN',4,'Perkelahian antar sekolah/ kelompok luar sekolah ( TAWURAN )',75,true],
['J01','J. INTIMIDASI/ANCAMAN/PENGANIYAAN',1,'Mengintimidasi/mengancam sesame peserta didik',40,true],['J02','J. INTIMIDASI/ANCAMAN/PENGANIYAAN',2,'Menganiay,mengeroyok,peserta didik',60,true],['J03','J. INTIMIDASI/ANCAMAN/PENGANIYAAN',3,'Mengintimidasi kepala sekolah,guru dan karyawan',60,true],['J04','J. INTIMIDASI/ANCAMAN/PENGANIYAAN',4,'Mengeroyok kepala sekolah,guru dan karyawan',100,true],
['K01','K. KEJAHATAN/PEMBANGKANGAN',1,'Merusak,mengotori,dan mencoret coret fasilitas milik sekolah',45,true],['K02','K. KEJAHATAN/PEMBANGKANGAN',2,'Membuang sampah tidak pada tempatnya',5,true],['K03','K. KEJAHATAN/PEMBANGKANGAN',3,'Tidak melaksanakan piket kelas',5,true],['K04','K. KEJAHATAN/PEMBANGKANGAN',4,'Buang air kecil/besar sembarangan tempat/tidak pada tempatnya',15,true],
['L01','L. KEJAHATAN PENIPUAN',1,'Peserta didik memalsukan tanda tangan',50,true],['L02','L. KEJAHATAN PENIPUAN',2,'Peserta didik, Memalsukan nilai dari guru',20,true],['L03','L. KEJAHATAN PENIPUAN',3,'Peserta didik melakukan penipuan didalam lingkungan sekolah.',20,true],['L04','L. KEJAHATAN PENIPUAN',4,'Peserta didik memalsukan dokumen sekolah/dokumen lainnya',50,true],['L05','L. KEJAHATAN PENIPUAN',5,'Peserta didik melakukan penipuan atas nama sekolah/organisasi didalam sekolah',50,true],['L06','L. KEJAHATAN PENIPUAN',6,'Peserta didik memalsukan surat izin,surat keterangan sakit atau lainnya',30,true]
  ];
  sh.clearContents();sh.getRange(1,1,1,6).setValues([['Kode','Kategori','No','Jenis_Pelanggaran','Poin','Aktif']]);sh.getRange(2,1,d.length,6).setValues(d);sh.setFrozenRows(1);sh.autoResizeColumns(1,6);return d.length;
}

function seedMasterPembinaanOfficial_(){
  const sh=getSheet_(APP.SHEETS.MASTER_PEMBINAAN);
  const d=[
['B01',5,25,'PEMBINAAN I',0,'Wali Kelas + Guru BK','Teguran dan peringatan oleh Wali Kelas dan Guru BK. Simpulan pembinaan dicatat dalam buku pembinaan peserta didik dan orang tua.','',true],
['B02',26,50,'PEMBINAAN II',1,'Wali Kelas','Orang tua diundang ke-1 oleh Wali Kelas. Peserta didik dinasihati Wali Kelas dan Guru BK di hadapan orang tua. Wali Kelas menandai konsultasi ke-1.','',true],
['B03',51,65,'SP-1',2,'Wali Kelas','Orang tua diundang ke-2 oleh Wali Kelas. Peserta didik membuat SP-1 bermaterai 10.000 diketahui orang tua, Guru BK dan Waka Kesiswaan.','SP-1',true],
['B04',66,75,'SP-2',3,'Wali Kelas','Orang tua diundang ke-3 oleh Wali Kelas. Peserta didik membuat SP-2 bermaterai 10.000 diketahui orang tua, Guru BK dan Waka Kesiswaan.','SP-2',true],
['B05',76,85,'SP-3',4,'Waka Kesiswaan','Orang tua diundang ke-4 oleh Waka Kesiswaan. Peserta didik membuat SP-3 bermaterai 10.000 diketahui orang tua, Wali Kelas, Guru BK dan Kesiswaan.','SP-3',true],
['B06',86,90,'SKORSING',5,'Kepala Sekolah','Orang tua diundang ke-5 oleh Kepala Sekolah. Peserta didik diskorsing selama 5 hari kerja dan berada di rumah di bawah pengawasan orang tua.','',true],
['B07',91,100,'KONFERENSI KASUS',6,'Kepala Sekolah','Orang tua diundang ke-6 oleh Kepala Sekolah. Dilakukan konferensi kasus bersama peserta didik, orang tua, Wali Kelas, Guru BK, Kesiswaan dan seluruh manajemen sekolah. Peserta didik dikembalikan sepenuhnya kepada orang tua.','',true]
  ];
  sh.clearContents();sh.getRange(1,1,1,9).setValues([['Kode','Min_Poin','Max_Poin','Status','Pemanggilan_Ke','Pihak','Tindakan','Dokumen','Aktif']]);sh.getRange(2,1,d.length,9).setValues(d);sh.setFrozenRows(1);sh.autoResizeColumns(1,9);return d.length;
}
