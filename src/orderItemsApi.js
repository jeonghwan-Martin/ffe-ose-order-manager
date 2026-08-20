// 발주관리 탭(FF&E/OS&E) — Supabase order_items 테이블 데이터 레이어
// 공정표 탭(TestScheduleDashboard.jsx)과 동일하게 raw fetch + PostgREST 방식 사용
const SUPABASE_URL = "https://fsjyzehxovazlmuihxxd.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZzanl6ZWh4b3ZhemxtdWloeHhkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5MTQ5NzcsImV4cCI6MjEwMTQ5MDk3N30.SayUMy8ajeMKGYmzek0H152dKwCLEzTP38yYm8u0a-g";

const sbHeaders = {
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  "Content-Type": "application/json",
};

// ---- App.jsx 내부 형태(ffeItems/oseItems) <-> Supabase order_items 로우 매핑 ----
// App.jsx 내부: { id, name, unitPrice, actualUnitPrice, installUnitPrice, installActualUnitPrice, qtyPerRoom }
// order_items 로우: item_name, supply_budget_unit_price, supply_actual_unit_price,
//                   install_budget_unit_price, install_actual_unit_price, quantity,
//                   room_type_id(FF&E) / null(OS&E는 common_area_id도 null, 프로젝트 전체 공통),
//                   category_group, sub_category, order_owner (분류값, 로컬 상태엔 없으므로 기존 값 보존)

function toRow(it, { projectId, roomTypeId }) {
  return {
    project_id: projectId,
    room_type_id: roomTypeId || null,
    common_area_id: null,
    item_name: it.name || "",
    quantity: it.qtyPerRoom || 0,
    supply_budget_unit_price: it.unitPrice || 0,
    supply_actual_unit_price: it.actualUnitPrice || 0,
    install_budget_unit_price: it.installUnitPrice || 0,
    install_actual_unit_price: it.installActualUnitPrice || 0,
    category_group: it.categoryGroup ?? null,
    sub_category: it.subCategory ?? null,
    order_owner: it.orderOwner ?? null,
  };
}

function fromRow(row) {
  return {
    id: row.id, // Supabase UUID를 그대로 프론트 id로 사용
    name: row.item_name || "",
    unitPrice: Number(row.supply_budget_unit_price) || 0,
    actualUnitPrice: Number(row.supply_actual_unit_price) || 0,
    installUnitPrice: Number(row.install_budget_unit_price) || 0,
    installActualUnitPrice: Number(row.install_actual_unit_price) || 0,
    qtyPerRoom: Number(row.quantity) || 0,
    categoryGroup: row.category_group,
    subCategory: row.sub_category,
    orderOwner: row.order_owner,
    roomTypeId: row.room_type_id,
  };
}

// 프로젝트의 order_items 전체를 읽어와 App.jsx 상태 형태(ffeItems/oseItems)로 변환
// roomTypeIdMap: { [App.jsx 로컬 roomTypeId]: Supabase room_types.id(uuid) } — ffeItems 키를 로컬 id로 되돌리는 데 사용
export async function loadOrderItems(projectId, roomTypeIdMap = {}) {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/order_items?project_id=eq.${projectId}&select=*`,
    { headers: sbHeaders }
  );
  if (!res.ok) throw new Error(`order_items 조회 실패 (${res.status})`);
  const rows = await res.json();

  const reverseMap = {};
  Object.entries(roomTypeIdMap).forEach(([localId, uuid]) => {
    reverseMap[uuid] = localId;
  });

  const ffeItems = {};
  const oseItems = [];
  rows.forEach((row) => {
    const item = fromRow(row);
    if (row.room_type_id) {
      const localRoomTypeId = reverseMap[row.room_type_id] || row.room_type_id;
      if (!ffeItems[localRoomTypeId]) ffeItems[localRoomTypeId] = [];
      ffeItems[localRoomTypeId].push(item);
    } else {
      oseItems.push(item);
    }
  });
  return { ffeItems, oseItems };
}

// 프로젝트의 ffeItems/oseItems를 order_items 테이블과 동기화
// 전체 교체 방식: 기존 프로젝트 order_items를 삭제 후 현재 상태를 다시 insert
// (품목 수가 많지 않은 스프레드시트형 데이터라 diff 대신 단순 전체 교체가 안전하고 구현이 명확함)
// roomTypeIdMap: { [App.jsx 로컬 roomTypeId]: Supabase room_types.id(uuid) } — roomTypesApi.saveRoomTypes()의 반환값을 그대로 전달
export async function saveOrderItems(projectId, ffeItems, oseItems, roomTypeIdMap) {
  const rows = [];
  Object.entries(ffeItems).forEach(([localRoomTypeId, items]) => {
    const supabaseRoomTypeId = roomTypeIdMap[localRoomTypeId];
    if (!supabaseRoomTypeId) return; // 매칭되는 room_types row가 없으면(동기화 누락) 건너뜀
    items.forEach((it) => rows.push(toRow(it, { projectId, roomTypeId: supabaseRoomTypeId })));
  });
  oseItems.forEach((it) => rows.push(toRow(it, { projectId, roomTypeId: null })));

  const delRes = await fetch(
    `${SUPABASE_URL}/rest/v1/order_items?project_id=eq.${projectId}`,
    { method: "DELETE", headers: sbHeaders }
  );
  if (!delRes.ok) throw new Error(`order_items 삭제 실패 (${delRes.status})`);

  if (rows.length === 0) return { count: 0 };

  const insRes = await fetch(`${SUPABASE_URL}/rest/v1/order_items`, {
    method: "POST",
    headers: { ...sbHeaders, Prefer: "return=representation" },
    body: JSON.stringify(rows),
  });
  if (!insRes.ok) throw new Error(`order_items 저장 실패 (${insRes.status})`);
  const inserted = await insRes.json();
  return { count: inserted.length };
}
