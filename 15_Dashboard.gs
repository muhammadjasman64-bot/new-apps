/** V30 - Dashboard & Rekap Terpadu */
function getDashboardDetail_(){
  const cached=cacheGetJson_('DASHBOARD_DETAIL_CACHE');if(cached)return cached;
  const cfg=getConfigObject_();
  const year=Number(String(cfg.Tahun_Pelajaran||new Date().getFullYear()).split('/')[0])||new Date().getFullYear();
  const sem=String(cfg.Semester||1);
  const students=getSiswaAktif_();
  const abs=getRekapAbsensi_('semester',year,null,sem,cfg.Kelas||'');
  const b=getPeriodBounds_('semester',{year,semester:sem});
  const agg=aggregatePointPeriod_(b.start,b.end);
  const byClass={};
  const absMap=Object.fromEntries(abs.map(function(a){return [String(a.nisn),a];}));
  students.forEach(s=>{
    const kelas=String(s.kelas||'Tanpa Kelas');
    if(!byClass[kelas])byClass[kelas]={kelas,jumlah:0,kehadiran:0,poinPelanggaran:0,poinPenghargaan:0,perluPembinaan:0};
    const c=byClass[kelas]; c.jumlah++;
    const a=absMap[String(s.nisn)];
    c.kehadiran+=a?Number(a.kehadiran||0):0;
    const p=agg[String(s.nisn).trim()]||{pel:0,rew:0};
    c.poinPelanggaran+=Number(p.pel||0); c.poinPenghargaan+=Number(p.rew||0);
    if(getKeputusanPoin_(Math.max(0,(p.pel||0)-(p.rew||0))).status!=='NORMAL')c.perluPembinaan++;
  });
  const kelas=Object.values(byClass).map(x=>Object.assign(x,{rataKehadiran:x.jumlah?Number((x.kehadiran/x.jumlah).toFixed(2)):0})).sort((a,b)=>a.kelas.localeCompare(b.kelas));
  const top=students.map(s=>{const p=agg[String(s.nisn).trim()]||{pel:0,rew:0};return{nisn:s.nisn,nama:s.nama,kelas:s.kelas,poinBersih:Math.max(0,(p.pel||0)-(p.rew||0)),status:getKeputusanPoin_(Math.max(0,(p.pel||0)-(p.rew||0))).status};}).filter(x=>x.poinBersih>0).sort((a,b)=>b.poinBersih-a.poinBersih).slice(0,10);
  return cachePutJson_('DASHBOARD_DETAIL_CACHE',{periode:'Semester '+sem+' — Tahun Pelajaran '+(cfg.Tahun_Pelajaran||''),kelas,top},CACHE_TTL.DASHBOARD);
}
