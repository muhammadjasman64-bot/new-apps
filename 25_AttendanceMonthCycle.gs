/**
 * 25_AttendanceMonthCycle.gs
 * Siklus bulan absensi: bulan baru otomatis menjadi periode input aktif,
 * tetapi histori ABSENSI tidak pernah dihapus.
 */
function ensureAttendanceMonthCycle_(){
  const tz=Session.getScriptTimeZone()||'Asia/Jakarta',now=new Date();
  const current=Utilities.formatDate(now,tz,'yyyy-MM');
  const props=PropertiesService.getScriptProperties();
  const previous=props.getProperty('ATTENDANCE_ACTIVE_MONTH')||'';
  const changed=previous!==current;
  if(changed){
    props.setProperty('ATTENDANCE_ACTIVE_MONTH',current);
    clearAppCache_();
    if(previous)logActivity_('ROLLOVER ABSENSI BULAN',previous+' -> '+current+'; histori ABSENSI dipertahankan.');
  }
  return {success:true,currentMonth:current,previousMonth:previous||null,changed:changed,historyRetained:true,message:changed?'Bulan absensi baru aktif. Data bulan sebelumnya tetap tersimpan dan dapat dipanggil kembali melalui Riwayat Absensi.':'Bulan absensi aktif.'};
}
function getAttendanceMonthState_(){
  return ensureAttendanceMonthCycle_();
}
