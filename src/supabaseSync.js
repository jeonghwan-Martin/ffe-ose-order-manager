// 여러 행을 한 번에 "안전하게" 동기화하는 공용 헬퍼 (2026-09-28)
// 예전 방식(전체 삭제 → 재삽입)은 삭제 후 삽입이 실패하면 프로젝트 데이터가 통째로 비어버렸다.
// 이제는 ① 기존 id 조회 → ② upsert(추가·수정) → ③ 화면에서 사라진 행만 삭제 순서로 진행한다.
// ②가 실패하면 아무것도 지워지지 않고, ③이 실패해도 지워야 할 행이 남을 뿐 데이터는 사라지지 않는다.
const SUPABASE_URL = "https://fsjyzehxovazlmuihxxd.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZzanl6ZWh4b3ZhemxtdWloeHhkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5MTQ5NzcsImV4cCI6MjEwMTQ5MDk3N30.SayUMy8ajeMKGYmzek0H152dKwCLEzTP38yYm8u0a-g";

const sbHeaders = {
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  "Content-Type": "application/json",
};

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export const isUuid = (v) => typeof v === "string" && UUID_RE.test(v);

// URL 길이 제한을 피하려고 id 목록을 나눠서 삭제
const DELETE_CHUNK = 100;

// table: 테이블명
// filter: 이 동기화 범위를 정하는 PostgREST 필터 문자열 (예: "project_id=eq.<uuid>")
// rows: 저장할 행 배열 — 모든 행에 keyField 값이 채워져 있어야 한다
// keyField: 행을 식별하는 컬럼(기본 "id"). onConflict: upsert 기준 컬럼(기본 keyField)
export async function syncRows(table, filter, rows, { keyField = "id", onConflict } = {}) {
  // 1) 현재 DB에 있는 행 조회
  const select = keyField === "id" ? "id" : `id,${keyField}`;
  const existingRes = await fetch(`${SUPABASE_URL}/rest/v1/${table}?${filter}&select=${select}`, {
    headers: sbHeaders,
  });
  if (!existingRes.ok) throw new Error(`${table} 조회 실패 (${existingRes.status})`);
  const existing = await existingRes.json();

  // 2) upsert — 실패하면 여기서 멈추므로 기존 데이터는 그대로 남는다
  if (rows.length > 0) {
    const upsertRes = await fetch(
      `${SUPABASE_URL}/rest/v1/${table}?on_conflict=${onConflict || keyField}`,
      {
        method: "POST",
        headers: { ...sbHeaders, Prefer: "resolution=merge-duplicates,return=minimal" },
        body: JSON.stringify(rows),
      }
    );
    if (!upsertRes.ok) throw new Error(`${table} 저장 실패 (${upsertRes.status})`);
  }

  // 3) 화면에서 사라진 행만 삭제
  const kept = new Set(rows.map((r) => String(r[keyField])));
  const toDelete = existing.filter((e) => !kept.has(String(e[keyField]))).map((e) => e.id);
  for (let i = 0; i < toDelete.length; i += DELETE_CHUNK) {
    const idList = toDelete.slice(i, i + DELETE_CHUNK).join(",");
    const delRes = await fetch(`${SUPABASE_URL}/rest/v1/${table}?id=in.(${idList})`, {
      method: "DELETE",
      headers: sbHeaders,
    });
    if (!delRes.ok) throw new Error(`${table} 삭제 실패 (${delRes.status})`);
  }
  return { count: rows.length, deleted: toDelete.length };
}
