/** 08_Dokumen.gs - upload kop/logo/tanda tangan */
function uploadDocument_(obj) {
  if(!obj||!obj.data||!obj.name) throw new Error('File tidak valid.');
  const folder=getOrCreateFolder_();
  const bytes=Utilities.base64Decode(obj.data.split(',').pop());
  const blob=Utilities.newBlob(bytes,obj.mimeType||MimeType.PNG,obj.name);
  const file=folder.createFile(blob);
  const type=obj.jenis||'LAINNYA';
  getSheet_(APP.SHEETS.DOKUMEN).appendRow([generateID_('DOC'),type,obj.name,file.getId(),file.getUrl(),new Date(),true]);
  saveConfigDocument_(type,file.getId(),file.getUrl());
  return {success:true,fileId:file.getId(),url:file.getUrl()};
}
function getOrCreateFolder_(){const props=PropertiesService.getScriptProperties(),key='PDF_FOLDER_ID',saved=props.getProperty(key);if(saved){try{return DriveApp.getFolderById(saved);}catch(e){props.deleteProperty(key);}}const it=DriveApp.getFoldersByName(APP.STORAGE.FOLDER_NAME);const folder=it.hasNext()?it.next():DriveApp.createFolder(APP.STORAGE.FOLDER_NAME);try{props.setProperty(key,folder.getId());}catch(e){}return folder;}
function fileImageDataUrl_(fileId){if(!fileId)return '';const key='DOC_IMG_'+String(fileId).replace(/[^A-Za-z0-9_.-]/g,'_');const cached=cacheGetJson_(key);if(cached&&cached.dataUrl)return cached.dataUrl;try{const b=DriveApp.getFileById(fileId).getBlob(),dataUrl='data:'+b.getContentType()+';base64,'+Utilities.base64Encode(b.getBytes());cachePutJson_(key,{dataUrl:dataUrl},600);return dataUrl;}catch(e){return '';}}

function saveConfigDocument_(type,id,url){
  const map={KOP_SURAT:'ID_Kop_Surat',LOGO:'ID_Logo',TTD_WALI:'ID_TTD_Wali',TTD_WAKA:'ID_TTD_Waka',TTD_BK:'ID_TTD_BK',TTD_KEPSEK:'ID_TTD_Kepala'};
  if(!map[type])return;
  const sh=getSheet_(APP.SHEETS.CONFIG), data=sh.getDataRange().getValues();
  let idx=data.findIndex((r,i)=>i>0&&r[0]===map[type]);
  if(idx<1)sh.appendRow([map[type],id]);else sh.getRange(idx+1,2).setValue(id);
  const urlMap={KOP_SURAT:'URL_Kop_Surat',LOGO:'URL_Logo',TTD_WALI:'URL_TTD_Wali',TTD_WAKA:'URL_TTD_Waka',TTD_BK:'URL_TTD_BK',TTD_KEPSEK:'URL_TTD_Kepala'};
  idx=data.findIndex((r,i)=>i>0&&r[0]===urlMap[type]);
  if(idx<1)sh.appendRow([urlMap[type],url]);else sh.getRange(idx+1,2).setValue(url);
}
function getDocumentConfig_(){
  const c=getConfigObject_();
  return {kop:c.ID_Kop_Surat||'',logo:c.ID_Logo||'',wali:c.ID_TTD_Wali||'',waka:c.ID_TTD_Waka||'',bk:c.ID_TTD_BK||'',kepsek:c.ID_TTD_Kepala||''};
}
