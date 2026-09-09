// Pure, read-only planning. A suggested match is never permission to overwrite.
export function planLegacyImport(snapshot, projects) {
  if (snapshot?.format !== 'ffe-ose-legacy-export-v1' || !Array.isArray(snapshot.projects)) throw new Error('지원하지 않는 백업 형식입니다.');
  const seen=new Set();
  const plan=snapshot.projects.map(entry=>{
    if (!entry.id || seen.has(entry.id)) throw new Error('백업에 중복 또는 누락된 프로젝트 ID가 있습니다.');
    seen.add(entry.id);
    let candidates=projects.filter(p=>p.client_id===entry.id || p.id===entry.id || `ov-${p.id}`===entry.id);
    let matchBy='id';
    if (!candidates.length) {candidates=projects.filter(p=>p.name===entry.name);matchBy='exact-name-review';}
    const payload=entry.payload;
    if(payload!==null && (!payload || typeof payload!=='object' || Array.isArray(payload)))throw new Error('프로젝트 payload 형식이 올바르지 않습니다.');
    const errors=[];
    if(payload){
      if(payload.roomTypes!==undefined && !Array.isArray(payload.roomTypes))errors.push('roomTypes 형식 오류');
      if(payload.oseItems!==undefined && !Array.isArray(payload.oseItems))errors.push('oseItems 형식 오류');
      if(payload.ffeItems!==undefined && (typeof payload.ffeItems!=='object'||Array.isArray(payload.ffeItems)||payload.ffeItems===null))errors.push('ffeItems 형식 오류');
      const roomIds=new Set((Array.isArray(payload.roomTypes)?payload.roomTypes:[]).map(r=>String(r.id)));
      if(payload.ffeItems && typeof payload.ffeItems==='object')for(const [id,items] of Object.entries(payload.ffeItems)){
        if(!Array.isArray(items))errors.push('FF&E 품목 목록 형식 오류');
        else if(items.length && !roomIds.has(id))errors.push('객실 유형이 없는 FF&E 품목');
      }
    }
    return {legacyId:entry.id,name:entry.name,projectId:candidates.length===1?candidates[0].id:null,
      matchBy:candidates.length===1?matchBy:null,
      status:errors.length?'invalid':candidates.length>1?'ambiguous':!payload?'no-saved-payload':candidates.length===0?'unmatched':matchBy==='exact-name-review'?'review-match':'ready',
      errors,roomTypes:Array.isArray(payload?.roomTypes)?payload.roomTypes.length:0,
      orderItems:(Array.isArray(payload?.oseItems)?payload.oseItems.length:0)+Object.values(payload?.ffeItems||{}).reduce((n,a)=>n+(Array.isArray(a)?a.length:0),0),
      payload};
  });
  const targets=new Map();
  for(const entry of plan)if(entry.projectId){const a=targets.get(entry.projectId)||[];a.push(entry);targets.set(entry.projectId,a);}
  for(const entries of targets.values())if(entries.length>1)for(const entry of entries){entry.status='duplicate-target';entry.errors.push('여러 이전 데이터가 같은 프로젝트에 연결됨');}
  return plan;
}
