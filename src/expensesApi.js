// 현장지출/인건비지출/예산외지출 — Supabase project_expenses 테이블 데이터 레이어
// 현장지출/예산외지출은 {id,name,budgetAmount,actualAmount} 직접입력 방식.
// 인건비지출은 인원수가 유동적인 항목(알바 등)이 있어 {..., unitPrice,actualUnitPrice,quantity,actualQuantity}도
// 함께 저장 — budgetAmount/actualAmount는 App.jsx에서 단가×인원수로 미리 계산해 채워 넣은 값(하위호환 유지).
// 세 지출 종류는 expense_type 컬럼으로 한 테이블에서 구분한다.
// room_types와 동일하게 client_id로 App.jsx 로컬 nextId()와 Supabase UUID를 매칭한다.
const SUPABASE_URL = "https://fsjyzehxovazlmuihxxd.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZzanl6ZWh4b3ZhemxtdWloeHhkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5MTQ5NzcsImV4cCI6MjEwMTQ5MDk3N30.SayUMy8ajeMKGYmzek0H152dKwCLEzTP38yYm8u0a-g";

const sbHeaders = {
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  "Content-Type": "application/json",
};

const TYPE_MAP = {
  site: "현장지출",
  labor: "인건비지출",
  extra: "예산외지출",
};
const TYPE_KEYS = Object.keys(TYPE_MAP); // ["site", "labor", "extra"]

function toRow(it, projectId, expenseType) {
  return {
    project_id: projectId,
    client_id: String(it.id),
    expense_type: expenseType,
    name: it.name || "",
    budget_amount: it.budgetAmount || 0,
    actual_amount: it.actualAmount || 0,
    unit_price: it.unitPrice != null ? it.unitPrice : null,
    actual_unit_price: it.actualUnitPrice != null ? it.actualUnitPrice : null,
    quantity: it.quantity != null ? it.quantity : null,
    actual_quantity: it.actualQuantity != null ? it.actualQuantity : null,
    days: it.days != null ? it.days : null,
    actual_days: it.actualDays != null ? it.actualDays : null,
  };
}

// siteExpenses/laborExpenses/extraExpenses를 project_expenses와 동기화
// (전체 삭제 후 재삽입 — order_items와 동일한 전체교체 방식, 품목 수가 적어 안전하고 단순함)
export async function saveExpenses(projectId, { siteExpenses, laborExpenses, extraExpenses }) {
  const rows = [
    ...siteExpenses.map((it) => toRow(it, projectId, TYPE_MAP.site)),
    ...laborExpenses.map((it) => toRow(it, projectId, TYPE_MAP.labor)),
    ...extraExpenses.map((it) => toRow(it, projectId, TYPE_MAP.extra)),
  ];

  const delRes = await fetch(
    `${SUPABASE_URL}/rest/v1/project_expenses?project_id=eq.${projectId}`,
    { method: "DELETE", headers: sbHeaders }
  );
  if (!delRes.ok) throw new Error(`project_expenses 삭제 실패 (${delRes.status})`);

  if (rows.length === 0) return { count: 0 };

  const insRes = await fetch(`${SUPABASE_URL}/rest/v1/project_expenses`, {
    method: "POST",
    headers: { ...sbHeaders, Prefer: "return=representation" },
    body: JSON.stringify(rows),
  });
  if (!insRes.ok) throw new Error(`project_expenses 저장 실패 (${insRes.status})`);
  const inserted = await insRes.json();
  return { count: inserted.length };
}

// Supabase project_expenses를 읽어 App.jsx 로컬 형태(siteExpenses/laborExpenses/extraExpenses)로 변환
export async function loadExpenses(projectId) {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/project_expenses?project_id=eq.${projectId}&select=*`,
    { headers: sbHeaders }
  );
  if (!res.ok) throw new Error(`project_expenses 조회 실패 (${res.status})`);
  const rows = await res.json();

  const result = { siteExpenses: [], laborExpenses: [], extraExpenses: [] };
  const reverseTypeMap = {
    [TYPE_MAP.site]: "siteExpenses",
    [TYPE_MAP.labor]: "laborExpenses",
    [TYPE_MAP.extra]: "extraExpenses",
  };
  rows.forEach((row) => {
    const key = reverseTypeMap[row.expense_type];
    if (!key) return;
    result[key].push({
      id: row.client_id,
      name: row.name || "",
      budgetAmount: Number(row.budget_amount) || 0,
      actualAmount: Number(row.actual_amount) || 0,
      ...(row.unit_price != null ? { unitPrice: Number(row.unit_price) } : {}),
      ...(row.actual_unit_price != null ? { actualUnitPrice: Number(row.actual_unit_price) } : {}),
      ...(row.quantity != null ? { quantity: Number(row.quantity) } : {}),
      ...(row.actual_quantity != null ? { actualQuantity: Number(row.actual_quantity) } : {}),
      ...(row.days != null ? { days: Number(row.days) } : {}),
      ...(row.actual_days != null ? { actualDays: Number(row.actual_days) } : {}),
    });
  });
  return result;
}
