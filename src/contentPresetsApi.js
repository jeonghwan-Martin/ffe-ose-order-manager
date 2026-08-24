// 룸타입 "카탈로그 기본세트 불러오기" — content_presets(+catalog_items) 템플릿 조회
// room/capacity/bed 3축 계산의 기준값(qtyPerRoom, calcBasis, multiplier, mattressSize)을 그대로 가져와
// ffeItems[roomTypeId]에 채워넣는다. 이후 값은 프로젝트별로 자유롭게 편집 가능(템플릿은 원본 그대로 유지됨).
const SUPABASE_URL = "https://fsjyzehxovazlmuihxxd.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZzanl6ZWh4b3ZhemxtdWloeHhkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5MTQ5NzcsImV4cCI6MjEwMTQ5MDk3N30.SayUMy8ajeMKGYmzek0H152dKwCLEzTP38yYm8u0a-g";

const sbHeaders = {
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  "Content-Type": "application/json",
};

// category: 룸타입의 category 값(예: "MASSAGE", 체인형은 보통 "STANDARD" 등 — content_presets에 없는 값이면
// 공통베이스만 매칭되고 전용콘텐츠는 자연스럽게 매칭 안 됨, 이게 일반형/체인형을 나누는 별도 플래그가 필요 없는 이유)
// 반환: [{ name, catalogItemId, unitPrice, calcBasis, qtyPerRoom, multiplier, mattressSize }]
export async function fetchContentPresets(category) {
  // PostgREST or= 조건 조합이 복잡해지므로 공통베이스/전용콘텐츠를 두 번 나눠 조회 후 합친다
  const commonRes = await fetch(
    `${SUPABASE_URL}/rest/v1/content_presets?category=is.null&select=*,catalog_items(item_name,calc_basis,reference_supply_price)`,
    { headers: sbHeaders }
  );
  if (!commonRes.ok) throw new Error(`content_presets(공통) 조회 실패 (${commonRes.status})`);
  const common = await commonRes.json();

  let categoryRows = [];
  if (category) {
    const catRes = await fetch(
      `${SUPABASE_URL}/rest/v1/content_presets?category=eq.${encodeURIComponent(category)}&select=*,catalog_items(item_name,calc_basis,reference_supply_price)`,
      { headers: sbHeaders }
    );
    if (!catRes.ok) throw new Error(`content_presets(전용) 조회 실패 (${catRes.status})`);
    categoryRows = await catRes.json();
  }

  const rows = [...common, ...categoryRows];
  return rows.map((row) => {
    const ci = row.catalog_items || {};
    return {
      name: row.item_name ?? ci.item_name ?? "",
      catalogItemId: row.catalog_item_id ?? null,
      unitPrice: Number(ci.reference_supply_price) || 0,
      calcBasis: ci.calc_basis || "room",
      qtyPerRoom: Number(row.default_qty) || 0,
      multiplier: row.default_multiplier != null ? Number(row.default_multiplier) : 1,
      mattressSize: row.mattress_size ?? null,
    };
  });
}

// OS&E(공통 품목) 전용 — content_presets 중 room 기준(calc_basis='room') 공통베이스만 조회.
// OS&E는 특정 룸타입에 묶이지 않는 프로젝트 전체 공통 품목이라 capacity/bed축(룸타입별 인원·침대구성 필요)은
// 여기서 계산할 수 없음 — 그 축의 품목(린넨류 등)은 기존처럼 FF&E 쪽(룸타입별 카드)에서 불러오는 게 맞음.
// OS&E 화면엔 별도 multiplier 컬럼이 없으므로 default_qty×default_multiplier를 미리 곱해 qtyPerRoom 하나로 반환.
export async function fetchOseContentPresets() {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/content_presets?category=is.null&select=*,catalog_items(item_name,calc_basis,reference_supply_price)`,
    { headers: sbHeaders }
  );
  if (!res.ok) throw new Error(`content_presets(OS&E) 조회 실패 (${res.status})`);
  const rows = await res.json();
  return rows
    .filter((row) => (row.catalog_items?.calc_basis || "room") === "room")
    .map((row) => {
      const ci = row.catalog_items || {};
      const mult = row.default_multiplier != null ? Number(row.default_multiplier) : 1;
      return {
        name: row.item_name ?? ci.item_name ?? "",
        catalogItemId: row.catalog_item_id ?? null,
        unitPrice: Number(ci.reference_supply_price) || 0,
        qtyPerRoom: (Number(row.default_qty) || 0) * mult,
      };
    });
}
