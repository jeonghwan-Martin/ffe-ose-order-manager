// 발주관리 탭(FF&E/OS&E) — Supabase order_items 테이블 데이터 레이어
// 공정표 탭(TestScheduleDashboard.jsx)과 동일하게 raw fetch + PostgREST 방식 사용
import { syncRows, isUuid } from "./supabaseSync";

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
//                   category_group, sub_category, order_owner (분류값, 로컬 상태엔 없으므로 기존 값 보존),
//                   carton_size (박스/팩당 개수 — 발주수량 올림 기준)

function toRow(it, { projectId, roomTypeId }) {
  return {
    id: it.id, // assignOrderItemIds()로 미리 UUID를 채워둔 값 — upsert 기준
    project_id: projectId,
    room_type_id: roomTypeId || null,
    common_area_id: null,
    item_name: it.name || "",
    quantity: it.qtyPerRoom || 0, // 기준수량(룸당/인당/침대당) — 총 발주수량은 계산근거(base_calc_basis)와 함께 화면에서 산출
    supply_budget_unit_price: it.unitPrice || 0,
    supply_actual_unit_price: it.actualUnitPrice || 0,
    install_budget_unit_price: it.installUnitPrice || 0,
    install_actual_unit_price: it.installActualUnitPrice || 0,
    category_group: it.categoryGroup ?? null,
    sub_category: it.subCategory ?? null,
    order_owner: it.orderOwner ?? null,
    catalog_item_id: it.catalogItemId ?? null,
    applied_multiplier: it.multiplier != null ? it.multiplier : null,
    base_calc_basis: it.calcBasis ?? null,
    mattress_size: it.mattressSize ?? null,
    vendor_id: it.vendorId || null, // 발주 업체(FK) — 발주서 생성 시 업체별 그룹핑 기준
    brand: it.brand || null, // 브랜드/제조사 — 발주서 열용, 수동 입력
    spec: it.spec || null, // 규격 — 카탈로그 default_spec에서 초기값, 수정 가능
    // 박스/팩당 개수 — 발주수량 올림(roundToCarton)의 기준. 저장 대상에서 빠져 있어
    // 새로고침하면 카톤 올림이 풀리던 버그가 있었음(2026-09-11 수정)
    carton_size: it.cartonSize != null && it.cartonSize !== "" ? Number(it.cartonSize) : null,
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
    catalogItemId: row.catalog_item_id,
    multiplier: row.applied_multiplier != null ? Number(row.applied_multiplier) : 1,
    calcBasis: row.base_calc_basis || "room",
    mattressSize: row.mattress_size,
    vendorId: row.vendor_id || "",
    brand: row.brand || "",
    spec: row.spec || "",
    cartonSize: row.carton_size != null ? Number(row.carton_size) : null,
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

// 저장 전에 모든 품목이 고유한 UUID id를 갖도록 정리한다.
// 새로 추가한 품목(로컬 숫자 id)이나 복사로 id가 겹친 품목에는 새 UUID를 붙인다.
// 반환된 ffeItems/oseItems를 화면 상태에도 반영해야 다음 저장 때 같은 행으로 인식된다.
export function assignOrderItemIds(ffeItems, oseItems) {
  const seen = new Set();
  let changed = false;
  const fix = (it) => {
    if (isUuid(it.id) && !seen.has(it.id)) {
      seen.add(it.id);
      return it;
    }
    changed = true;
    const id = crypto.randomUUID();
    seen.add(id);
    return { ...it, id };
  };
  const nextFfe = {};
  Object.entries(ffeItems).forEach(([rtId, items]) => {
    nextFfe[rtId] = items.map(fix);
  });
  const nextOse = oseItems.map(fix);
  return { ffeItems: nextFfe, oseItems: nextOse, changed };
}

// 프로젝트의 ffeItems/oseItems를 order_items 테이블과 동기화
// id 기준 upsert 후 화면에서 사라진 품목만 삭제(supabaseSync.syncRows) — 저장 도중 실패해도 기존 품목이 지워지지 않는다.
// ffeItems/oseItems는 assignOrderItemIds()를 거친 값이어야 한다.
// roomTypeIdMap: { [App.jsx 로컬 roomTypeId]: Supabase room_types.id(uuid) } — roomTypesApi.saveRoomTypes()의 반환값을 그대로 전달
export async function saveOrderItems(projectId, ffeItems, oseItems, roomTypeIdMap) {
  const rows = [];
  Object.entries(ffeItems).forEach(([localRoomTypeId, items]) => {
    const supabaseRoomTypeId = roomTypeIdMap[localRoomTypeId];
    if (!supabaseRoomTypeId) return; // 매칭되는 room_types row가 없으면(동기화 누락) 건너뜀
    items.forEach((it) => rows.push(toRow(it, { projectId, roomTypeId: supabaseRoomTypeId })));
  });
  oseItems.forEach((it) => rows.push(toRow(it, { projectId, roomTypeId: null })));

  return syncRows("order_items", `project_id=eq.${projectId}`, rows);
}
