// 프로젝트 설정값(카테고리/이레귤러옵션/브랜드룸네임/roomFeatures/viewTypes/floors/기본세트)
// — Supabase projects.settings(jsonb) 컬럼 데이터 레이어
// 서로 관계없는 옵션 리스트라 개별 테이블 대신 jsonb 한 덩어리로 저장한다.
const SUPABASE_URL = "https://fsjyzehxovazlmuihxxd.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZzanl6ZWh4b3ZhemxtdWloeHhkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5MTQ5NzcsImV4cCI6MjEwMTQ5MDk3N30.SayUMy8ajeMKGYmzek0H152dKwCLEzTP38yYm8u0a-g";

const sbHeaders = {
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  "Content-Type": "application/json",
};

// App.jsx 로컬 상태 -> settings jsonb 페이로드
function toSettings({ categories, irregularOptions, brandRoomName, roomFeatures, viewTypes, floors, basicPreset }) {
  return { categories, irregularOptions, brandRoomName, roomFeatures, viewTypes, floors, basicPreset };
}

// projectUuid: projectIdApi.resolveProjectUuid()로 확보한 Supabase projects.id(uuid)
export async function saveProjectSettings(projectUuid, state) {
  const settings = toSettings(state);
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/projects?id=eq.${projectUuid}`,
    {
      method: "PATCH",
      headers: { ...sbHeaders, Prefer: "return=representation" },
      body: JSON.stringify({ settings }),
    }
  );
  if (!res.ok) throw new Error(`projects.settings 저장 실패 (${res.status})`);
  const updated = await res.json();
  return updated[0];
}

// Supabase projects.settings를 읽어 App.jsx 로컬 형태로 변환
export async function loadProjectSettings(projectUuid) {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/projects?id=eq.${projectUuid}&select=settings`,
    { headers: sbHeaders }
  );
  if (!res.ok) throw new Error(`projects.settings 조회 실패 (${res.status})`);
  const rows = await res.json();
  if (rows.length === 0) return null;
  return rows[0].settings || null;
}
