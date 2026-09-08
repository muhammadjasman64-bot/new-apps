
function validateDatabase_(){
  const required={
    'DATA_SISWA':[
      ['NISN','NIS','nomor_induk_siswa_nasional'],
      ['Nama','Nama_Siswa','Nama Siswa']
    ],
    'ABSENSI':[
      ['Tanggal','Tanggal Absensi','Tgl','Tanggal_Kehadiran'],
      ['NISN','NIS','nomor_induk_siswa_nasional'],
      ['Status','Status Kehadiran','Kehadiran']
    ],
    'MASTER_PELANGGARAN':[
      ['Kode','Kode Pelanggaran'],
      ['Jenis Pelanggaran','Jenis','Nama Pelanggaran'],
      ['Poin','Point']
    ]
  };
  const out=[];
  Object.keys(required).forEach(name=>{
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if(!sh){out.push({sheet:name,ok:false,error:'Sheet tidak ditemukan'});return;}
    const map=dbGetHeaderMap_(sh);
    const missing=[];
    required[name].forEach(group=>{
      if(dbFindColumn_(map,group)<0) missing.push(group[0]);
    });
    out.push({sheet:name,ok:missing.length===0,missing:missing,headers:sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0]});
  });
  return {ok:out.every(x=>x.ok),items:out};
}
