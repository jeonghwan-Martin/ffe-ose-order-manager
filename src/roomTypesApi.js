// 발주관리 탭 룸타입 빌더 — Supabase room_types(+room_type_beds) 테이블 데이터 레이어
// App.jsx는 로컬 nextId()로 룸타입 id를 관리하므로, client_id 컬럼으로 Supabase UUID와 매칭한다.
// byFloor(층별 배치)는 이번 단계에서 보류 — Supabase엔 아직 반영하지 않고 기존 Apps Script 블롭에만 유지.
const SUPABASE_URL = "https://fsjyzehxovazlmuihxxd.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZzanl6ZWh4b3ZhemxtdWloeHhkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5MTQ5NzcsImV4cCI6MjEwMTQ5MDk3N30.SayUMy8ajeMKGYmzek0H152dKwCLEzTP38yYm8u0a-g";

const sbHeaders = {
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  "Content-Type": "application/json",
};

// App.jsx의 codeFor(rt)와 동일한 규칙 — room_types.code는 NOT NULL이라 반드시 채워야 함
function codeFor(rt) {
  const bed = rt.bed === "싱글" ? "S" : "Q";
  const bath = rt.bathtub === "유" ? "B" : "NB";
  const cat = (rt.category || "").slice(0, 2) || "RT";
  return `${cat}-${bed}-${bath}`;
}

// App.jsx의 roomTypeTotal(rt)와 동일한 규칙(byFloor 합산) — room_count는 참고용으로 같이 저장
function roomCountOf(rt) {
  return Object.values(rt.byFloor || {}).reduce((a, b) => a + b, 0);
}

// App.jsx의 effectiveBedComposition(rt)와 동일한 규칙 — bedComposition 미지정 시 bed/mattressQty로 단일구성 유추
function effectiveBedComposition(rt) {
  if (rt.bedComposition && rt.bedComposition.length > 0) return rt.bedComposition;
  const size = rt.bed === "싱글" ? "S" : "Q";
  return [{ size, qty: rt.mattressQty || 1 }];
}

function toRow(rt, projectId) {
  return {
    project_id: projectId,
    client_id: String(rt.id),
    code: codeFor(rt),
    bed_type: rt.bed ?? null,
    has_bathtub: rt.bathtub === "유", // 기존엔 !!rt.bathtub라 "무" 문자열도 true로 잘못 저장되던 버그 수정
    category: rt.category ?? null,
    mattress_qty: rt.mattressQty ?? null,
    room_count: roomCountOf(rt),
    capacity: rt.capacity ?? null,
    irregular_options: rt.irregular ?? [],
    grade: rt.grade ?? null,
    features: rt.features ?? [],
    view_type: rt.view ?? null,
    include_bed_in_name: !!rt.includeBedInName,
    include_view_in_name: !!rt.includeViewInName,
    custom_name: rt.customName ?? null,
    ota_bed_count: rt.otaBedCount ?? null,
    ota_bed_size: rt.otaBedSize ?? null,
    ota_max_occupancy: rt.otaMaxOccupancy ?? null,
    ota_facilities: rt.otaFacilities ?? null,
    room_numbers: rt.roomNumbers ?? [],
  };
}

// idMap({localId: supabaseUuid})과 roomTypes를 받아 room_type_beds를 전체교체 방식으로 동기화
async function syncRoomTypeBeds(idMap, roomTypes) {
  const roomTypeUuids = Object.values(idMap);
  if (roomTypeUuids.length > 0) {
    const idList = roomTypeUuids.join(",");
    const delRes = await fetch(
      `${SUPABASE_URL}/rest/v1/room_type_beds?room_type_id=in.(${idList})`,
      { method: "DELETE", headers: sbHeaders }
    );
    if (!delRes.ok) throw new Error(`room_type_beds 삭제 실패 (${delRes.status})`);
  }

  const rows = [];
  roomTypes.forEach((rt) => {
    const uuid = idMap[String(rt.id)];
    if (!uuid) return;
    effectiveBedComposition(rt).forEach((b) => {
      rows.push({ room_type_id: uuid, mattress_size: b.size, bed_qty: b.qty });
    });
  });
  if (rows.length === 0) return;

  const insRes = await fetch(`${SUPABASE_URL}/rest/v1/room_type_beds`, {
    method: "POST",
    headers: sbHeaders,
    body: JSON.stringify(rows),
  });
  if (!insRes.ok) throw new Error(`room_type_beds 저장 실패 (${insRes.status})`);
}

// projectUuid: projectIdApi.resolveProjectUuid()로 확보한 Supabase projects.id(uuid)
// roomTypes(App.jsx 로컬 상태)를 Supabase room_types+room_type_beds와 동기화(client_id 기준 upsert, 삭제된 항목 제거)
// 반환값: { [localId]: supabaseUuid } — order_items 저장 시 room_type_id FK 채우는 데 사용
export async function saveRoomTypes(projectUuid, roomTypes) {
  // 1) 이 프로젝트에서 사라진 룸타입(더 이상 로컬에 없는 client_id) 삭제
  const existingRes = await fetch(
    `${SUPABASE_URL}/rest/v1/room_types?project_id=eq.${projectUuid}&select=id,client_id`,
    { headers: sbHeaders }
  );
  if (!existingRes.ok) throw new Error(`room_types 조회 실패 (${existingRes.status})`);
  const existing = await existingRes.json();
  const currentIds = new Set(roomTypes.map((rt) => String(rt.id)));
  const toDelete = existing.filter((e) => e.client_id && !currentIds.has(e.client_id));
  if (toDelete.length > 0) {
    const idList = toDelete.map((e) => e.id).join(",");
    const delRes = await fetch(
      `${SUPABASE_URL}/rest/v1/room_types?id=in.(${idList})`,
      { method: "DELETE", headers: sbHeaders }
    );
    if (!delRes.ok) throw new Error(`room_types 삭제 실패 (${delRes.status})`);
    // room_type_beds는 room_type_id FK CASCADE 여부가 불확실하므로 명시적으로도 정리
    const bedsDelRes = await fetch(
      `${SUPABASE_URL}/rest/v1/room_type_beds?room_type_id=in.(${idList})`,
      { method: "DELETE", headers: sbHeaders }
    );
    if (!bedsDelRes.ok) throw new Error(`room_type_beds(삭제된 룸타입) 정리 실패 (${bedsDelRes.status})`);
  }

  if (roomTypes.length === 0) return {};

  // 2) upsert (client_id, project_id 유니크 인덱스 기준)
  const rows = roomTypes.map((rt) => toRow(rt, projectUuid));
  const upsertRes = await fetch(
    `${SUPABASE_URL}/rest/v1/room_types?on_conflict=project_id,client_id`,
    {
      method: "POST",
      headers: { ...sbHeaders, Prefer: "resolution=merge-duplicates,return=representation" },
      body: JSON.stringify(rows),
    }
  );
  if (!upsertRes.ok) throw new Error(`room_types 저장 실패 (${upsertRes.status})`);
  const saved = await upsertRes.json();

  const idMap = {};
  saved.forEach((row) => {
    idMap[row.client_id] = row.id;
  });

  // 3) room_type_beds 동기화(전체교체) — 침대구성은 room_types upsert 이후 uuid가 확정된 다음에 처리
  await syncRoomTypeBeds(idMap, roomTypes);

  return idMap;
}

// Supabase room_types(+room_type_beds)를 읽어 App.jsx 로컬 형태로 변환
// (byFloor/room_count는 포함하지 않음 — Apps Script 블롭 값으로 보완 필요, byFloor 합산이 곧 room_count이므로)
export async function loadRoomTypes(projectId) {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/room_types?project_id=eq.${projectId}&select=*`,
    { headers: sbHeaders }
  );
  if (!res.ok) throw new Error(`room_types 조회 실패 (${res.status})`);
  const rows = await res.json();

  const roomTypeIds = rows.map((r) => r.id);
  let bedsByRoomType = {};
  if (roomTypeIds.length > 0) {
    const bedsRes = await fetch(
      `${SUPABASE_URL}/rest/v1/room_type_beds?room_type_id=in.(${roomTypeIds.join(",")})&select=*`,
      { headers: sbHeaders }
    );
    if (bedsRes.ok) {
      const bedsRows = await bedsRes.json();
      bedsRows.forEach((b) => {
        if (!bedsByRoomType[b.room_type_id]) bedsByRoomType[b.room_type_id] = [];
        bedsByRoomType[b.room_type_id].push({ size: b.mattress_size, qty: b.bed_qty });
      });
    }
  }

  const idMap = {};
  const roomTypes = rows.map((row) => {
    idMap[row.client_id] = row.id;
    return {
      id: row.client_id,
      bed: row.bed_type,
      bathtub: row.has_bathtub ? "유" : "무", // 저장 시 문자열→boolean 변환의 역변환
      category: row.category,
      mattressQty: row.mattress_qty,
      capacity: row.capacity ?? 2,
      bedComposition: bedsByRoomType[row.id] || [],
      irregular: row.irregular_options || [],
      grade: row.grade,
      features: row.features || [],
      view: row.view_type,
      includeBedInName: row.include_bed_in_name,
      includeViewInName: row.include_view_in_name,
      customName: row.custom_name,
      otaBedCount: row.ota_bed_count,
      otaBedSize: row.ota_bed_size,
      otaMaxOccupancy: row.ota_max_occupancy,
      otaFacilities: row.ota_facilities,
      roomNumbers: row.room_numbers || [],
    };
  });
  return { roomTypes, idMap };
}
