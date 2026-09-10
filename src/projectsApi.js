// 발주관리 탭 프로젝트 목록 — Supabase projects 테이블 데이터 레이어 (2026-09-10 신규)
// 예전엔 발주관리 탭이 Apps Script(PropertiesService)에 별도 프로젝트 목록({id, name})을 갖고 있어서
// 공정표(Supabase projects)와 이름이 어긋나면 새 프로젝트가 중복 생성되는 문제가 있었다.
// 이제 발주관리 탭 드롭다운도 이 파일을 통해 공정표와 같은 projects 테이블을 직접 읽고 쓴다.
const SUPABASE_URL = "https://fsjyzehxovazlmuihxxd.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZzanl6ZWh4b3ZhemxtdWloeHhkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5MTQ5NzcsImV4cCI6MjEwMTQ5MDk3N30.SayUMy8ajeMKGYmzek0H152dKwCLEzTP38yYm8u0a-g";

const sbHeaders = {
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  "Content-Type": "application/json",
};

// 드롭다운용 전체 목록. 공정표 탭과 같은 우선순위로 정렬한다:
// 드래그로 수동 순서(manual_sort_order)를 지정한 프로젝트가 앞, 나머지는 이름순.
export async function listProjects() {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/projects?select=id,name,client_id,assigned_budget,has_content,manual_sort_order,settings`,
    { headers: sbHeaders }
  );
  if (!res.ok) throw new Error(`projects 목록 조회 실패 (${res.status})`);
  const rows = await res.json();
  return rows.sort((a, b) => {
    const av = a.manual_sort_order ?? Infinity;
    const bv = b.manual_sort_order ?? Infinity;
    if (av !== bv) return av - bv;
    return (a.name || "").localeCompare(b.name || "", "ko");
  });
}

// 프로젝트 1건 조회(이름/배정예산/settings/client_id)
export async function getProject(id) {
  const res = await fetch(
    `${SUPABASE_URL}/rest/v1/projects?id=eq.${id}&select=id,name,client_id,assigned_budget,has_content,settings`,
    { headers: sbHeaders }
  );
  if (!res.ok) throw new Error(`projects 조회 실패 (${res.status})`);
  const rows = await res.json();
  return rows[0] || null;
}

// 이름·배정예산 등 컬럼 값 갱신 (settings jsonb는 projectSettingsApi.saveProjectSettings에서 따로 처리)
export async function updateProject(id, patch) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/projects?id=eq.${id}`, {
    method: "PATCH",
    headers: { ...sbHeaders, Prefer: "return=representation" },
    body: JSON.stringify(patch),
  });
  if (!res.ok) throw new Error(`projects 갱신 실패 (${res.status})`);
  const rows = await res.json();
  return rows[0];
}

// 새 프로젝트 생성 — 공정표 탭의 "+ 새 프로젝트"와 동일하게 기본 마일스톤 5단계(milestone_templates)도 같이 채운다.
// (발주관리 탭에서 만든 프로젝트가 공정표에도 바로 정상적으로 뜨게 하기 위함)
export async function createProject(name) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/projects`, {
    method: "POST",
    headers: { ...sbHeaders, Prefer: "return=representation" },
    body: JSON.stringify([{ name }]),
  });
  if (!res.ok) throw new Error(`projects 생성 실패 (${res.status})`);
  const [project] = await res.json();

  try {
    const tplRes = await fetch(
      `${SUPABASE_URL}/rest/v1/milestone_templates?select=*&order=sort_order.asc`,
      { headers: sbHeaders }
    );
    if (tplRes.ok) {
      const templates = await tplRes.json();
      if (templates.length > 0) {
        await fetch(`${SUPABASE_URL}/rest/v1/project_milestones`, {
          method: "POST",
          headers: { ...sbHeaders, Prefer: "return=minimal" },
          body: JSON.stringify(
            templates.map((t) => ({
              project_id: project.id,
              template_id: t.id,
              name: t.name,
              sort_order: t.sort_order,
              weight: t.weight,
            }))
          ),
        });
      }
    }
  } catch (e) {
    // 마일스톤 시딩 실패는 프로젝트 생성 자체를 되돌릴 사유가 아님 — 공정표 탭에서 '+ 직접입력'으로 보완 가능
  }
  return project;
}
