/** 10_WebApp.gs - API WebApp V6 */
function doGet(){return HtmlService.createTemplateFromFile('index').evaluate().setTitle(APP.NAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}
function include_(name){return HtmlService.createHtmlOutputFromFile(name).getContent();}
function api(name,args,token){
  try{
    const fn=String(name||''); const a=Array.isArray(args)?args:[];
    if(fn==='login') return login_.apply(null,a);
    if(fn==='logout') return logout_.apply(null,a);
    if(fn==='getSession') return getSessionForClient_.apply(null,a);
    if(fn==='listUsers') return listUsers_.apply(null,a);
    if(fn==='saveUser') return saveUser_.apply(null,a);
    if(fn==='resetUserPassword') return resetUserPassword_.apply(null,a);
    if(fn==='deleteUser') return deleteUser_.apply(null,a);
    switch(fn){
      case'getSiswaAktif':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getSiswaAktif_.apply(null,args);
      case'getSiswaSemua':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getSiswaSemua_.apply(null,args);
      case'setSiswaStatus':requireAuth_(token,['admin']);return setSiswaStatus_.apply(null,args);
      case'restoreSiswa':requireAuth_(token,['admin']);return restoreSiswa_.apply(null,args);
      case'importSiswaRows':requireAuth_(token,['admin']);return importSiswaRows_.apply(null,args);
      case'getPelanggaranOptions':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getPelanggaranOptions_.apply(null,args);
      case'getMasterPenghargaan':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getMasterPenghargaan_.apply(null,args);
      case'getMasterPembinaan':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getMasterPembinaan_.apply(null,args);
      case'getAturanInfo':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getAturanInfoV37_.apply(null,args);
      case'validateOfficialRules':requireAuth_(token,['admin']);return validateOfficialRules_();
      case'getDashboard':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getDashboard_.apply(null,args);
      case'getDashboardDetail':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getDashboardDetail_.apply(null,args);
      case'saveSiswa':requireAuth_(token,['admin']);return saveSiswa_.apply(null,args);
      case'saveAbsensi':requireAuth_(token,['admin']);return saveAbsensi_.apply(null,args);
      case'getAbsensiTable':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getAbsensiTable_.apply(null,args);
      case'getAbsensiHistory':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getAbsensiHistory_.apply(null,args);
      case'getAbsensiSummary':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getAbsensiSummary_.apply(null,args);
      case'deleteAbsensi':requireAuth_(token,['admin']);return deleteAbsensi_.apply(null,args);
      case'savePelanggaran':requireAuth_(token,['admin']);return savePelanggaran_.apply(null,args);
      case'getPelanggaranHistory':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getPelanggaranHistory_.apply(null,args);
      case'getProfilPelanggaran':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getProfilPelanggaran_.apply(null,args);
      case'getProfilPembinaan':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getProfilPembinaan_.apply(null,args);
      case'getDecisionPreview':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getDecisionPreview_.apply(null,args);
      case'evaluateAllStudentsDecisions':requireAuth_(token,['admin']);return evaluateAllStudentsDecisions_.apply(null,args);
      case'getStudentStatusAlert':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getStudentStatusAlert_.apply(null,args);
      case'getStudentAlerts':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getStudentAlerts_.apply(null,args);
      case'getStatusAlertSummary':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getStatusAlertSummary_.apply(null,args);
      case'evaluateStudentStatusAlerts':requireAuth_(token,['admin']);return evaluateStudentStatusAlerts_.apply(null,args);
      case'getDecisionThresholdTest':requireAuth_(token,['admin']);return getDecisionThresholdTest_.apply(null,args);
      case'getAttendancePolicy':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getAttendancePolicy_();
      case'getAttendanceRiskForStudent':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getAttendanceRiskForStudent_.apply(null,args);
      case'getAttendanceRiskSummary':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getAttendanceRiskSummary_();
      case'getCombinedRiskSummary':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getCombinedRiskSummary_();
      case'evaluateAttendanceRiskAll':requireAuth_(token,['admin']);return evaluateAttendanceRiskAll_();
      case'getUnifiedRiskForStudent':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getUnifiedRiskForStudent_.apply(null,args);
      case'getUnifiedRiskSummary':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getUnifiedRiskSummary_();
      case'getUnifiedRiskAlerts':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getUnifiedRiskAlerts_();
      case'getUnifiedRiskPolicy':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getUnifiedRiskPolicy_();
      case'evaluateUnifiedRiskAll':requireAuth_(token,['admin']);return evaluateUnifiedRiskAll_();
      case'getEarlyWarningSummary':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getEarlyWarningSummary_();
      case'getEarlyWarningForStudent':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getEarlyWarningForStudent_.apply(null,args);
      case'getEarlyWarningPolicy':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getEarlyWarningPolicy_();
      case'evaluateEarlyWarningAll':requireAuth_(token,['admin']);return evaluateEarlyWarningAll_();
      case'resolveEarlyWarning':requireAuth_(token,['admin', 'wali_kelas', 'bk']);return resolveEarlyWarning_.apply(null,args);
      case'updatePelanggaran':requireAuth_(token,['admin']);return updatePelanggaran_.apply(null,args);
      case'deletePelanggaran':requireAuth_(token,['admin']);return deletePelanggaran_.apply(null,args);
      case'savePenghargaan':requireAuth_(token,['admin']);return savePenghargaan_.apply(null,args);
      case'buildRekap':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return buildRekap_.apply(null,args);
      case'generateAndSaveRekap':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return generateAndSaveRekap_.apply(null,args);
      case'testRekap':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return testRekap_.apply(null,args);
      case'testGetPelanggaranOptions':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return testGetPelanggaranOptions_.apply(null,args);
      case'diagnoseRekap':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return diagnoseRekap_.apply(null,args);
      case'getRekapSnapshot':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getRekapSnapshot_.apply(null,args);
      case'getStudentReportDetail':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getStudentReportDetail_.apply(null,args);
      case'renderReportHtml':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return renderReportHtml_.apply(null,args);
      case'createPdfFromReport':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return createPdfFromReport_.apply(null,args);
      case'updateRekapRow':requireAuth_(token,['admin']);return updateRekapRow_.apply(null,args);
      case'deleteRekapRow':requireAuth_(token,['admin']);return deleteRekapRow_.apply(null,args);
      case'uploadDocument':requireAuth_(token,['admin']);return uploadDocument_.apply(null,args);
      case'apiSyncDatabase':requireAuth_(token,['admin']);return apiSyncDatabase_.apply(null,args);
      case'apiValidateDatabase':requireAuth_(token,['admin']);return apiValidateDatabase_.apply(null,args);
      case'getDbVersion':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getDbVersion_.apply(null,args);
      case'getAdminSummary':requireAuth_(token,['admin']);return getAdminSummary_.apply(null,args);
      case'backupDatabase':requireAuth_(token,['admin']);return backupDatabase_.apply(null,args);
      case'exportDatabaseCsv':requireAuth_(token,['admin']);return exportDatabaseCsv_.apply(null,args);
      case'getAuditLog':requireAuth_(token,['admin']);return getAuditLog_.apply(null,args);
      case'validateDatabaseDetailed':requireAuth_(token,['admin']);return validateDatabaseDetailed_.apply(null,args);
      case'clearApplicationCache':requireAuth_(token,['admin']);return clearApplicationCache_.apply(null,args);
      case'securityValidateUsers':requireAuth_(token,['admin']);return securityValidateUsers_();
      case'getTindakanHistory':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getTindakanHistory_.apply(null,args);
      case'updateTindakanStatus':requireAuth_(token,['admin']);return updateTindakanStatus_.apply(null,args);
      case'getPemanggilanHistory':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return getPemanggilanHistory_.apply(null,args);
      case'savePemanggilan':requireAuth_(token,['admin']);return savePemanggilan_.apply(null,args);
      case'updatePemanggilanStatus':requireAuth_(token,['admin']);return updatePemanggilanStatus_.apply(null,args);
      case'renderSuratPemanggilanHtml':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return renderSuratPemanggilanHtml_.apply(null,args);
      case'createPdfSuratPemanggilan':requireAuth_(token,['admin', 'wali_kelas', 'guru', 'bk', 'kepala_sekolah']);return createPdfSuratPemanggilan_.apply(null,args);
      default:throw new Error('API tidak ditemukan: '+fn);
    }
  }catch(e){ return {success:false,message:e&&e.message?e.message:String(e)}; }
}
function findStudentByNisn_(nisn){return getSiswa_().find(s=>String(s.nisn).trim()===String(nisn).trim());}
function parseDateInput_(v){if(!v)return null;if(v instanceof Date&&!isNaN(v.getTime()))return new Date(v.getTime());const s=String(v).trim();let m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);if(m)return new Date(+m[1],+m[2]-1,+m[3]);m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);if(m)return new Date(+m[3],+m[2]-1,+m[1]);const dt=new Date(s);return isNaN(dt.getTime())?null:dt;}
function formatDateKey_(v){const d=new Date(v);return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd');}
function formatInputDate_(v){return Utilities.formatDate(new Date(v),Session.getScriptTimeZone(),'yyyy-MM-dd');}
function logActivity_(a,d){try{getSheet_(APP.SHEETS.LOG).appendRow([new Date(),Session.getActiveUser().getEmail()||'WebApp',a,d||'']);}catch(e){}}


/** API aman untuk tombol Sinkronisasi Database pada frontend. */
function apiSyncDatabase_(){
  clearAppCache_();
  return syncDatabase_();
}
function apiValidateDatabase_(){
  return validateDatabase_();
}
