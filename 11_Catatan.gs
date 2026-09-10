/** 11_Catatan.gs - jurnal/catatan wali kelas dan BK */
function saveCatatanWali_(obj){
  return saveCatatanGeneric_(APP.SHEETS.CATATAN_WALI_KELAS,obj);
}
function saveCatatanBK_(obj){
  return saveCatatanGeneric_(APP.SHEETS.CATATAN_BK,obj);
}
function saveCatatanGeneric_(sheetName,obj){
  const s=findStudentByNisn_(obj.nisn); if(!s)throw new Error('Siswa tidak ditemukan.');
  const c=getConfigObject_();
  getSheet_(sheetName).appendRow([generateID_('CAT'),parseDateInput_(obj.tanggal)||new Date(),s.nisn,s.nama,obj.catatan||'',obj.tindakLanjut||'',
    Session.getActiveUser().getEmail()||'WebApp',c.Tahun_Pelajaran||getTahunPelajaran_(),c.Semester||getSemesterAktif_()]);
  commitDatabaseMutation_();
  return {success:true,nisn:s.nisn,nama:s.nama};
}
function getCatatan_(sheetName,nisn){
  return getSheet_(sheetName).getDataRange().getValues().slice(1).filter(r=>canonicalNisn_(r[2])===canonicalNisn_(nisn))
    .map(r=>({id:r[0],tanggal:formatTanggal_(r[1]),nisn:canonicalNisn_(r[2]),nama:r[3],catatan:r[4],tindakLanjut:r[5]}));
}
