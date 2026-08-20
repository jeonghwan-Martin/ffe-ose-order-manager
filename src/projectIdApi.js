// 발주관리 탭(App.jsx)은 로컬 문자열 id(예: proj_1234567_abcd)로 프로젝트를 관리하지만,
// Supabase projects.id는 uuid다. room_types.client_id와 동일한 패턴으로
// projects.client_id에 로컬 id를 보관하고, 이 함수로 실제 uuid를 조회/생성한다.
// room_types/order_items/project_expenses/projects.settings 저장은 전부 이 uuid를 써야 한다.
const SUPABASE_URL = "https://fsjyzehxovazlmuihxxd.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZzanl6ZWh4b3ZhemxtdWloeHhkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5MTQ5NzcsImV4cCI6MjEwMTQ5MDk3N30.SayUMy8ajeMKGYmzek0H152dKwCLEzTP38yYm8u0a-g";

const sbHeaders = {
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  "Content-Type": "application/json",
};

// clientId(발주관리 탭 로컬 프로젝트 id)에 대응하는 Supabase projects.id(uuid)를 반환.
// 없으면 최소 정보로 새로 생성한다.
export async function resolveProjectUuid(clientId, projectName) {
  const findRes = await fetch(
    `${SUPABASE_URL}/rest/v1/projects?client_id=eq.${encodeURIComponent(clientId)}&select=id`,
    { headers: sbHeaders }
  );
  if (!findRes.ok) throw new Error(`projects 조회 실패 (${findRes.status})`);
  const found = await findRes.json();
  if (found.length > 0) return found[0].id;

  const insRes = await fetch(`${SUPABASE_URL}/rest/v1/projects`, {
    method: "POST",
    headers: { ...sbHeaders, Prefer: "return=representation" },
    body: JSON.stringify([{ client_id: clientId, name: projectName || "" }]),
  });
  if (!insRes.ok) throw new Error(`projects 생성 실패 (${insRes.status})`);
  const inserted = await insRes.json();
  return inserted[0].id;
}
