/** V42.3 - Authentication, admin recovery & server-side authorization. */
const AUTH={TOKEN_TTL:21600, CACHE_PREFIX:'AUTH_SESSION_', ROLES:['admin','wali_kelas','guru','bk','kepala_sekolah']};
function hashPassword_(password){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(password),Utilities.Charset.UTF_8).map(b=>{const n=b<0?b+256:b;return ('0'+n.toString(16)).slice(-2)}).join('');}

function authDbHeaders_(sheetName){
  const sh=getSheet_(sheetName);
  const lastCol=sh.getLastColumn();
  if(!lastCol) return [];
  return sh.getRange(1,1,1,lastCol).getValues()[0].map(function(v){
    return String(v==null?'':v).trim();
  });
}

function findUserByUsername_(username){
  const users=dbReadObjects_('USERS');
  const target=String(username||'').trim().toLowerCase();
  for(let i=0;i<users.length;i++){
    const u=users[i];
    if(String(u.Username||'').trim().toLowerCase()===target)return u;
  }
  return null;
}
function seedDefaultAdmin_(){
  const sh=getSheet_('USERS');
  const rows=dbReadObjects_('USERS');
  const existing=findUserByUsername_('admin');
  if(existing){
    const role=String(existing.Role||'').trim().toLowerCase();
    const status=String(existing.Status||'').trim().toLowerCase();
    if(role==='admin' && status==='aktif' && String(existing.Password_Hash||'').trim())
      return {success:true,created:false,message:'Admin sudah ada dan aktif.'};
    const rowIndex=rows.indexOf(existing)+2;
    const headers=authDbHeaders_('USERS');
    const obj={}; headers.forEach(h=>obj[h]=existing[h]!==undefined?existing[h]:'');
    obj.Role='admin'; obj.Status='aktif'; obj.Password_Hash=hashPassword_('Admin123!'); obj.Updated_At=new Date();
    sh.getRange(rowIndex,1,1,headers.length).setValues([headers.map(h=>obj[h]!==undefined?obj[h]:'')]);
    return {success:true,created:false,repaired:true,message:'Admin diperbaiki dan diaktifkan. Username: admin | Password: Admin123!'};
  }
  sh.appendRow([generateID_('USR'),'admin',hashPassword_('Admin123!'),'Administrator','admin','','','aktif',new Date(),new Date()]);
  return {success:true,created:true,message:'Admin awal dibuat. Username: admin | Password: Admin123! Segera ganti password.'};
}
function setupAdmin_(newPassword){
  const password=String(newPassword||'Admin123!');
  validatePasswordPolicy_(password);
  const sh=getSheet_('USERS'), rows=dbReadObjects_('USERS');
  let idx=-1; for(let i=0;i<rows.length;i++){if(String(rows[i].Username||'').trim().toLowerCase()==='admin'){idx=i;break;}}
  if(idx<0){
    sh.appendRow([generateID_('USR'),'admin',hashPassword_(password),'Administrator','admin','','','aktif',new Date(),new Date()]);
  }else{
    const headers=authDbHeaders_('USERS'), row=rows[idx], values=headers.map(h=>row[h]!==undefined?row[h]:'');
    const set=(h,v)=>{const i=headers.indexOf(h);if(i>=0)values[i]=v;};
    set('Password_Hash',hashPassword_(password)); set('Role','admin'); set('Status','aktif'); set('Nama_Lengkap',row.Nama_Lengkap||'Administrator'); set('Updated_At',new Date());
    sh.getRange(idx+2,1,1,headers.length).setValues([values]);
  }
  clearLoginFailures_('admin'); clearAppCache_(); auditSecurity_('ADMIN_BOOTSTRAP','username=admin');
  return {success:true,username:'admin',message:'Admin aktif. Password berhasil diatur ulang.'};
}
function validateAuthentication_(){
  const rows=dbReadObjects_('USERS');
  const admins=rows.filter(r=>String(r.Role||'').toLowerCase()==='admin');
  const activeAdmins=admins.filter(r=>String(r.Status||'').toLowerCase()==='aktif');
  const validAdmins=activeAdmins.filter(r=>String(r.Username||'').trim() && String(r.Password_Hash||'').trim());
  return {success:true,totalUsers:rows.length,totalAdmins:admins.length,activeAdmins:activeAdmins.length,validAdmins:validAdmins.length,adminUsernames:activeAdmins.map(r=>String(r.Username||''))};
}

function login_(username,password){
  const u=String(username||'').trim(), p=String(password||'');
  if(!u||!p)throw new Error('Username dan password wajib diisi.');
  const failKey=securityKey_(u), failCount=Number(CacheService.getScriptCache().get(failKey)||0);
  if(failCount>=SECURITY.LOGIN_MAX_FAILS)throw new Error('Terlalu banyak percobaan login_ gagal. Coba lagi beberapa menit kemudian.');
  const row=findUserByUsername_(u);
  const h=hashPassword_(p);
  const validRow=row && String(row.Password_Hash||'')===h && String(row.Status||'aktif').trim().toLowerCase()==='aktif';
  if(!validRow){const n=recordLoginFailure_(u);auditSecurity_('LOGIN_GAGAL','username='+u+';percobaan='+n);throw new Error(n>=SECURITY.LOGIN_MAX_FAILS?'Terlalu banyak percobaan login_ gagal. Coba lagi beberapa menit kemudian.':'Username/password salah atau akun tidak aktif.');}
  clearLoginFailures_(u);
  const role=String(row.Role||'').toLowerCase();
  if(!AUTH.ROLES.includes(role))throw new Error('Role akun tidak valid. Hubungi administrator.');
  const token=Utilities.getUuid()+'-'+Utilities.getUuid();
  const now=new Date().toISOString();
  const payload={uid:String(row.ID_User||''),username:u,nama:String(row.Nama_Lengkap||u),role:role,nisn:String(row.NISN||''),kelas:String(row.Kelas||''),loginAt:now,lastSeenAt:now};
  CacheService.getScriptCache().put(AUTH.CACHE_PREFIX+token,JSON.stringify(payload),SECURITY.SESSION_REFRESH_SEC);
  auditSecurity_('LOGIN_BERHASIL','username='+u+';role='+role);
  return {success:true,token:token,user:payload};
}
function logout_(token){const s=getSession_(token);if(token)CacheService.getScriptCache().remove(AUTH.CACHE_PREFIX+String(token));if(s)auditSecurity_('LOGOUT','username='+s.username);return {success:true};}
function getSessionForClient_(token){const s=getSession_(token);if(!s)return {success:false,user:null};const r=refreshSession_(token,s);return {success:true,user:r};}
function getSession_(token){if(!token)return null;const raw=CacheService.getScriptCache().get(AUTH.CACHE_PREFIX+String(token));if(!raw)return null;try{return JSON.parse(raw)}catch(e){return null;}}
function listUsers_(token){requireAuth_(token,['admin']);return dbReadObjects_('USERS').map(r=>({id:r.ID_User,username:r.Username,nama:r.Nama_Lengkap,role:r.Role,nisn:r.NISN,kelas:r.Kelas,status:r.Status}));}
function saveUser_(token,data){requireAuth_(token,['admin']);data=data||{};const sh=getSheet_('USERS');const headers=authDbHeaders_('USERS');const rows=dbReadObjects_('USERS');const username=String(data.username||'').trim();if(!username)throw new Error('Username wajib diisi.');if(!AUTH.ROLES.includes(String(data.role||'').toLowerCase()))throw new Error('Role tidak valid.');if(rows.some(r=>String(r.Username||'').toLowerCase()===username.toLowerCase()&&String(r.ID_User)!==String(data.id||'')))throw new Error('Username sudah digunakan.');const obj={ID_User:data.id||generateID_('USR'),Username:username,Password_Hash:data.password?hashPassword_(data.password):(data.password_hash||''),Nama_Lengkap:String(data.nama||''),Role:String(data.role).toLowerCase(),NISN:String(data.nisn||''),Kelas:String(data.kelas||''),Status:String(data.status||'aktif'),Created_At:new Date(),Updated_At:new Date()};if(!obj.Password_Hash&& !data.id)throw new Error('Password wajib diisi untuk user baru.');
 if(data.password)validatePasswordPolicy_(data.password);if(!obj.Password_Hash){const old=rows.find(r=>String(r.ID_User)===String(data.id));obj.Password_Hash=old?old.Password_Hash:'';obj.Created_At=old?old.Created_At:new Date();}
 const idx=rows.findIndex(r=>String(r.ID_User)===String(data.id));const vals=headers.map(h=>obj[h]!==undefined?obj[h]:'');if(idx>=0){const row=idx+2;sh.getRange(row,1,1,headers.length).setValues([vals]);}else sh.getRange(sh.getLastRow()+1,1,1,headers.length).setValues([vals]);clearAppCache_();return {success:true,id:obj.ID_User};}
function resetUserPassword_(token,id,newPassword){requireAuth_(token,['admin']);validatePasswordPolicy_(newPassword);const sh=getSheet_('USERS'), rows=dbReadObjects_('USERS'), idx=rows.findIndex(r=>String(r.ID_User)===String(id));if(idx<0)throw new Error('User tidak ditemukan.');const col=authDbHeaders_('USERS').indexOf('Password_Hash')+1;sh.getRange(idx+2,col).setValue(hashPassword_(newPassword));clearAppCache_();return {success:true};}
function deleteUser_(token,id){
  const actor=requireAuth_(token,['admin']);
  const rows=dbReadObjects_('USERS'),idx=rows.findIndex(r=>String(r.ID_User)===String(id));
  if(idx<0)throw new Error('User tidak ditemukan.');
  if(String(rows[idx].ID_User)===String(actor.uid))throw new Error('Admin yang sedang login_ tidak dapat menghapus akunnya sendiri.');
  const targetRole=String(rows[idx].Role||'').toLowerCase(),targetStatus=String(rows[idx].Status||'').toLowerCase();
  if(targetRole==='admin'&&targetStatus==='aktif'){const admins=rows.filter(r=>String(r.Role||'').toLowerCase()==='admin'&&String(r.Status||'').toLowerCase()==='aktif');if(admins.length<=1)throw new Error('Admin terakhir tidak boleh dihapus.');}
  getSheet_('USERS').deleteRow(idx+2);
  clearAppCache_();auditSecurity_('USER_DIHPUS','id='+id);
  return {success:true};
}
