/** 12_Cache.gs - cache kecil dan aman. Jangan cache transaksi besar. */
const CACHE_TTL={CONFIG:300,MASTER:600,STUDENTS:300,DASHBOARD:30,DERIVED:30};
function cacheDateKey_(v){return v?formatInputDate_(v):'ALL';}
function cacheKey_(prefix,parts){const v=PropertiesService.getScriptProperties().getProperty('DB_VERSION')||'0';return String(prefix)+'_'+v+'_'+parts.map(function(x){return String(x==null?'':x).replace(/[^A-Za-z0-9_.-]/g,'_');}).join('_');}
function cacheGetJson_(key){try{const v=CacheService.getScriptCache().get(key);return v?JSON.parse(v):null;}catch(e){return null;}}
function cachePutJson_(key,value,ttl){try{CacheService.getScriptCache().put(key,JSON.stringify(value),ttl||300);}catch(e){}return value;}
function clearAppCache_(){
  CacheService.getScriptCache().removeAll([
    'DB_LAST_SYNC','APP_CONFIG_CACHE','MASTER_PEL_CACHE','MASTER_REWARD_CACHE','STUDENTS_CACHE','DASHBOARD_CACHE','DASHBOARD_DETAIL_CACHE','STATUS_ALERT_SUMMARY_CACHE','ATTENDANCE_RISK_SUMMARY_CACHE','UNIFIED_RISK_SUMMARY_CACHE','EARLY_WARNING_SUMMARY_CACHE'
  ]);
  PropertiesService.getScriptProperties().setProperty('DB_VERSION',String(Date.now()));
  return {ok:true,message:'Cache aplikasi berhasil dibersihkan.'};
}
