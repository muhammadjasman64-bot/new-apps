/**
 * ================================================================
 * JURNAL WALI KELAS DIGITAL - V43.5
 * PUBLIC RUN ENTRY POINTS
 * ================================================================
 *
 * Fungsi di file ini sengaja PUBLIC agar muncul pada dropdown Run
 * Google Apps Script setelah project di-pull dari GitHub/Clasp.
 *
 * Fungsi inti tetap PRIVATE (akhiran _) untuk mempertahankan
 * arsitektur API Gateway dan keamanan V35+.
 *
 * Jangan menghapus file ini saat sinkronisasi GitHub.
 * ================================================================
 */

function setupDatabase() {
  return setupDatabase_();
}

function validateDatabase() {
  return apiValidateDatabase_();
}

function setupAdmin() {
  return setupAdmin_('Admin123!');
}

function validateAuthentication() {
  return validateAuthentication_();
}

function validateOfficialRules() {
  return validateOfficialRules_();
}

function restoreOfficialRules() {
  return restoreOfficialRuleMasters_();
}

function lockOfficialRules() {
  return lockOfficialRuleSheets_();
}

function syncDatabase() {
  return apiSyncDatabase_();
}

function rolloverAttendanceMonth() {
  return ensureAttendanceMonthCycle_();
}

function evaluateDecisionEngine() {
  return evaluateAllStudentsDecisions_();
}

function evaluateAttendanceRisk() {
  return evaluateAttendanceRiskAll_();
}

function evaluateUnifiedRisk() {
  return evaluateUnifiedRiskAll_();
}

function evaluateEarlyWarning() {
  return evaluateEarlyWarningAll_();
}

/**
 * Pemeriksaan cepat: memastikan fungsi inti tersedia setelah pull.
 */
function systemHealthCheck() {
  var required = [
    'setupDatabase_',
    'apiValidateDatabase_',
    'setupAdmin_',
    'validateAuthentication_',
    'validateOfficialRules_',
    'restoreOfficialRuleMasters_',
    'lockOfficialRuleSheets_',
    'apiSyncDatabase_',
    'evaluateAllStudentsDecisions_',
    'evaluateAttendanceRiskAll_',
    'evaluateUnifiedRiskAll_',
    'evaluateEarlyWarningAll_'
  ];

  var missing = required.filter(function(name) {
    try {
      return typeof this[name] !== 'function';
    } catch (e) {
      return true;
    }
  }, this);

  return {
    ok: missing.length === 0,
    version: 'V43.5',
    message: missing.length ? 'Ada fungsi inti yang tidak ditemukan.' : 'Public Run entry points OK. Auth/DB wrappers tersedia.',
    missing: missing
  };
}
