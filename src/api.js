// 구글 Apps Script를 "웹 앱"으로 배포한 뒤 나오는 URL을 여기에 붙여넣으세요.
// (배포 방법은 프로젝트 루트의 apps-script/README.md 참고)
export const APPS_SCRIPT_URL = "https://script.google.com/a/macros/spaceplanning.co.kr/s/AKfycbxJ_aERi7IWbCteRno_wHXeNhrUxFq-JrWJPKL4ltssmcKjBVst-vtgETdbyMLX11bTQw/exec";

async function apiGet(action, params = {}) {
  const qs = new URLSearchParams({ action, ...params }).toString();
  const res = await fetch(`${APPS_SCRIPT_URL}?${qs}`);
  if (!res.ok) throw new Error(`API GET 실패 (${res.status})`);
  return res.json();
}

// Apps Script 웹앱은 JSON Content-Type의 POST에서 브라우저 preflight(OPTIONS)를 제대로
// 처리하지 못하는 경우가 많아, text/plain으로 보내고 서버(doPost)에서 JSON.parse 합니다.
async function apiPost(payload) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify(payload),
  });
  if (!res.ok) throw new Error(`API POST 실패 (${res.status})`);
  return res.json();
}

export async function getProjectIndex() {
  const data = await apiGet("listProjects");
  return data.list || [];
}

export async function getProjectData(id) {
  const data = await apiGet("getProject", { id });
  return data.value ? JSON.parse(data.value) : null;
}

export async function saveProjectData(id, payload) {
  return apiPost({ action: "saveProject", id, value: JSON.stringify(payload) });
}

export async function saveProjectIndex(list) {
  return apiPost({ action: "saveIndex", value: JSON.stringify(list) });
}

export async function getCatalog() {
  const data = await apiGet("getCatalog");
  return data.value ? JSON.parse(data.value) : [];
}

export async function saveCatalog(items) {
  return apiPost({ action: "saveCatalog", value: JSON.stringify(items) });
}
