// 업체 마스터(vendors) — Supabase vendors 테이블 데이터 레이어
// 전역 공통 테이블(특정 프로젝트에 속하지 않음)이라, order_items/expenses처럼 "전체 삭제 후 재삽입"하면
// 이미 order_items.vendor_id가 참조 중인 UUID가 깨질 수 있음 — 대신 프로젝트 추가/삭제(TestScheduleDashboard.jsx)와
// 동일하게 행 단위 직접 CRUD(POST/PATCH/DELETE) 방식을 쓴다. 저장 버튼 없이 각 동작이 즉시 반영된다.
const SUPABASE_URL = "https://fsjyzehxovazlmuihxxd.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZzanl6ZWh4b3ZhemxtdWloeHhkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5MTQ5NzcsImV4cCI6MjEwMTQ5MDk3N30.SayUMy8ajeMKGYmzek0H152dKwCLEzTP38yYm8u0a-g";

const sbHeaders = {
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  "Content-Type": "application/json",
};

// Supabase row -> App.jsx 로컬 형태(camelCase)
function fromRow(row) {
  return {
    id: row.id,
    name: row.name || "",
    itemGroup: row.item_group || "",
    category: row.category || "",
    bankName: row.bank_name || "",
    accountNumber: row.account_number || "",
    accountHolder: row.account_holder || "",
    contactName: row.contact_name || "",
    phone: row.phone || "",
    email: row.email || "",
    remark: row.remark || "",
  };
}

// 전체 업체 목록 조회 (이름순 정렬)
export async function loadVendors() {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/vendors?select=*&order=name.asc`, {
    headers: sbHeaders,
  });
  if (!res.ok) throw new Error(`vendors 조회 실패 (${res.status})`);
  const rows = await res.json();
  return rows.map(fromRow);
}

// 업체 신규 생성 — name만 필수, 나머지는 비워서 만든 뒤 인라인으로 채울 수 있음
export async function createVendor({
  name,
  itemGroup = "",
  category = "",
  bankName = "",
  accountNumber = "",
  accountHolder = "",
  contactName = "",
  phone = "",
  email = "",
  remark = "",
}) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/vendors`, {
    method: "POST",
    headers: { ...sbHeaders, Prefer: "return=representation" },
    body: JSON.stringify([
      {
        name: name || "",
        item_group: itemGroup || null,
        category: category || null,
        bank_name: bankName || null,
        account_number: accountNumber || null,
        account_holder: accountHolder || null,
        contact_name: contactName || null,
        phone: phone || null,
        email: email || null,
        remark: remark || null,
      },
    ]),
  });
  if (!res.ok) throw new Error(`업체 생성 실패 (${res.status})`);
  const [row] = await res.json();
  return fromRow(row);
}

// 필드 하나(또는 여러 개) 업데이트 — field는 camelCase 키, patch는 { [field]: value } 형태
const FIELD_TO_COLUMN = {
  name: "name",
  itemGroup: "item_group",
  category: "category",
  bankName: "bank_name",
  accountNumber: "account_number",
  accountHolder: "account_holder",
  contactName: "contact_name",
  phone: "phone",
  email: "email",
  remark: "remark",
};

export async function updateVendor(id, patch) {
  const row = {};
  Object.entries(patch).forEach(([field, value]) => {
    const column = FIELD_TO_COLUMN[field];
    if (!column) return;
    row[column] = value === "" ? null : value;
  });
  if (Object.keys(row).length === 0) return;

  const res = await fetch(`${SUPABASE_URL}/rest/v1/vendors?id=eq.${id}`, {
    method: "PATCH",
    headers: sbHeaders,
    body: JSON.stringify(row),
  });
  if (!res.ok) throw new Error(`업체 수정 실패 (${res.status})`);
}

// 업체 삭제 — order_items.vendor_id FK가 ON DELETE NO ACTION이라, 이미 발주 품목에 배정된 업체는
// DB가 삭제를 거부한다(409 conflict). 그 경우 사용자가 이해할 수 있는 메시지로 바꿔서 던진다.
export async function deleteVendor(id) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/vendors?id=eq.${id}`, {
    method: "DELETE",
    headers: sbHeaders,
  });
  if (!res.ok) {
    if (res.status === 409) {
      throw new Error("이미 발주 품목에 배정된 업체라 삭제할 수 없습니다. 해당 품목의 업체 배정을 먼저 해제해주세요.");
    }
    throw new Error(`업체 삭제 실패 (${res.status})`);
  }
}
