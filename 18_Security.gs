/** V34 - Security hardening helpers. */
const SECURITY={LOGIN_MAX_FAILS:5,LOGIN_WINDOW_SEC:900,SESSION_REFRESH_SEC:21600,MIN_PASSWORD:8};
function securityKey_(username){return 'AUTH_FAIL_'+Utilities.base64EncodeWebSafe(String(username||'').trim().toLowerCase()).slice(0,80);}
function recordLoginFailure_(username){const c=CacheService.getScriptCache(),k=securityKey_(username);let n=Number(c.get(k)||0)+1;c.put(k,String(n),SECURITY.LOGIN_WINDOW_SEC);return n;}
function clearLoginFailures_(username){CacheService.getScriptCache().remove(securityKey_(username));}
function validatePasswordPolicy_(password){const p=String(password||'');if(p.length<SECURITY.MIN_PASSWORD)throw new Error('Password minimal 8 karakter.');if(!/[A-Za-z]/.test(p)||!/[0-9]/.test(p))throw new Error('Password harus mengandung huruf dan angka.');return true;}
function auditSecurity_(action,detail){try{logActivity_('SECURITY:'+action,detail||'')}catch(e){}}
function refreshSession_(token,session){if(!token||!session)return session;const refreshed=Object.assign({},session,{lastSeenAt:new Date().toISOString()});CacheService.getScriptCache().put(AUTH.CACHE_PREFIX+String(token),JSON.stringify(refreshed),SECURITY.SESSION_REFRESH_SEC);return refreshed;}
function requireAuth_(token,allowed){const s=getSession_(token);if(!s)throw new Error('Sesi tidak valid atau telah berakhir. Silakan login_ kembali.');if(allowed&&allowed.length&&!allowed.includes(s.role))throw new Error('Anda tidak memiliki hak akses untuk tindakan ini.');return refreshSession_(token,s);}
function securityValidateUsers_(){const rows=dbReadObjects_('USERS');const seen={};let activeAdmins=0;rows.forEach(r=>{const u=String(r.Username||'').trim().toLowerCase();if(u)seen[u]=(seen[u]||0)+1;if(String(r.Status||'').toLowerCase()==='aktif'&&String(r.Role||'').toLowerCase()==='admin')activeAdmins++;});const duplicateUsernames=Object.keys(seen).filter(k=>seen[k]>1);return {success:true,activeAdmins,duplicateUsernames};}
