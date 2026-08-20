// 발주관리 탭 룸타입 빌더 — Supabase room_types 테이블 데이터 레이어
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

function toRow(rt, projectId) {
  return {
    project_id: projectId,
    client_id: String(rt.id),
    bed_type: rt.bed ?? null,
    has_bathtub: !!rt.bathtub,
    category: rt.category ?? null,
    mattress_qty: rt.mattressQty ?? null,
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

// projectUuid: projectIdApi.resolveProjectUuid()로 확보한 Supabase projects.id(uuid)
// roomTypes(App.jsx 로컬 상태)를 Supabase room_types와 동기화(client_id 기준 upsert, 삭제된 항목 제거)
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
  return idMap;
}

// Supabase room_types를 읽어 App.jsx 로컬 형태로 변환 (byFloor는 포함하지 않음 — Apps Script 블롭 값으로 보완 필요)
export async function loadRoomTypes(projectId) {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/room_types?project_id=eq.${projectId}&select=*`,
    { headers: sbHeaders }
  );
  if (!res.ok) throw new Error(`room_types 조회 실패 (${res.status})`);
  const rows = await res.json();
  const idMap = {};
  const roomTypes = rows.map((row) => {
    idMap[row.client_id] = row.id;
    return {
      id: row.client_id,
      bed: row.bed_type,
      bathtub: row.has_bathtub,
      category: row.category,
      mattressQty: row.mattress_qty,
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
