// Supabase cutover adapter. Not enabled in App.jsx until legacy import is verified.
// Complete JSON documents preserve floor placement, price precision and fields
// which the old normalized synchronization omitted. Existing tables stay intact.
export function createProjectDocumentStore({url, key, fetchImpl = fetch}) {
  if (!url || !key) throw new Error('Supabase connection settings are required');
  async function request(resource, options = {}) {
    const response = await fetchImpl(`${url.replace(/\/$/, '')}/rest/v1/${resource}`, {
      ...options,
      headers:{apikey:key, Authorization:`Bearer ${key}`, 'Content-Type':'application/json', ...options.headers},
    });
    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      if (body.code === '40001') throw new Error('다른 팀원이 먼저 저장했습니다. 현재 내용을 보관한 뒤 최신 데이터를 다시 불러와 주세요.');
      throw new Error(body.message || `데이터 요청 실패 (${response.status})`);
    }
    return response.status === 204 ? null : response.json();
  }
  return {
    async load(projectId) {
      const rows = await request(`project_documents?project_id=eq.${encodeURIComponent(projectId)}&select=payload,version`);
      if (!rows.length) throw new Error('이 프로젝트의 데이터 이전이 아직 완료되지 않았습니다.');
      return rows[0];
    },
    async save(projectId, payload, expectedVersion) {
      if (!Number.isInteger(expectedVersion) || expectedVersion < 1) throw new Error('최신 데이터를 불러온 뒤 저장해 주세요.');
      if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw new Error('프로젝트 데이터가 올바르지 않습니다.');
      return request('rpc/save_project_document', {
        method:'POST',body:JSON.stringify({p_project_id:projectId,p_payload:payload,p_expected_version:expectedVersion}),
      });
    },
  };
}
