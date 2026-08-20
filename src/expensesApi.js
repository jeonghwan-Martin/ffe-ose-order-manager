// 현장지출/인건비지출/예산외지출 — Supabase project_expenses 테이블 데이터 레이어
// 세 지출 종류는 App.jsx에서 동일한 형태({id,name,budgetAmount,actualAmount})라
// 테이블 하나 + expense_type 컬럼으로 구분한다.
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
    });
  });
  return result;
}
