import { useState, useEffect, useMemo, useRef } from "react";

// 실배포용: Supabase에서 실시간으로 데이터를 읽고, 체크박스 클릭 시 즉시 저장합니다.
const SUPABASE_URL = "https://fsjyzehxovazlmuihxxd.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZzanl6ZWh4b3ZhemxtdWloeHhkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU5MTQ5NzcsImV4cCI6MjEwMTQ5MDk3N30.SayUMy8ajeMKGYmzek0H152dKwCLEzTP38yYm8u0a-g";

const sbHeaders = {
  apikey: SUPABASE_ANON_KEY,
  Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
  "Content-Type": "application/json",
};

// 구글 캘린더 → Supabase 동기화용 Apps Script 웹앱(doGet?action=syncNow). 매일 오전 7시 자동 실행되는
// 것과 같은 함수(syncOpenBuyingCalendar)를 수동으로도 바로 돌릴 수 있게 웹앱으로 배포해둔 URL.
const CALENDAR_SYNC_URL =
  "https://script.google.com/macros/s/AKfycbwe8Ol1nvPHrDEHbtYwamgZgFQ_NtBet32g9vD1PNXhg4q0ipSUFjRW2zYp0uYl_WwFjA/exec";

async function syncCalendarNow() {
  const res = await fetch(`${CALENDAR_SYNC_URL}?action=syncNow`);
  if (!res.ok) throw new Error(`동기화 요청 실패 (${res.status})`);
  const data = await res.json();
  if (!data.ok) throw new Error(data.message || "동기화에 실패했어요.");
  return data;
}

async function fetchAll() {
  const [pRes, mRes] = await Promise.all([
    fetch(`${SUPABASE_URL}/rest/v1/projects?select=*&order=target_open_date.asc`, {
      headers: sbHeaders,
    }),
    fetch(`${SUPABASE_URL}/rest/v1/project_milestones?select=*&order=sort_order.asc`, {
      headers: sbHeaders,
    }),
  ]);
  if (!pRes.ok || !mRes.ok) throw new Error("데이터를 불러오지 못했습니다");
  const [projects, milestones] = await Promise.all([pRes.json(), mRes.json()]);
  return { projects, milestones };
}

async function persistMilestone(id, patch) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/project_milestones?id=eq.${id}`, {
    method: "PATCH",
    headers: { ...sbHeaders, Prefer: "return=minimal" },
    body: JSON.stringify(patch),
  });
  if (!res.ok) throw new Error("저장 실패");
}

async function persistProject(id, patch) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/projects?id=eq.${id}`, {
    method: "PATCH",
    headers: { ...sbHeaders, Prefer: "return=minimal" },
    body: JSON.stringify(patch),
  });
  if (!res.ok) throw new Error("저장 실패");
}

async function createMilestone(patch) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/project_milestones`, {
    method: "POST",
    headers: { ...sbHeaders, Prefer: "return=representation" },
    body: JSON.stringify(patch),
  });
  if (!res.ok) throw new Error("추가 실패");
  const [row] = await res.json();
  return row;
}

async function deleteMilestoneRemote(id) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/project_milestones?id=eq.${id}`, {
    method: "DELETE",
    headers: sbHeaders,
  });
  if (!res.ok) throw new Error("삭제 실패");
}

// 프로젝트 신규 생성 시 기본 마일스톤 5단계를 채우기 위한 템플릿 조회
async function fetchMilestoneTemplates() {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/milestone_templates?select=*&order=sort_order.asc`, {
    headers: sbHeaders,
  });
  if (!res.ok) throw new Error("마일스톤 템플릿 조회 실패");
  return res.json();
}

async function createProjectRemote(name) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/projects`, {
    method: "POST",
    headers: { ...sbHeaders, Prefer: "return=representation" },
    body: JSON.stringify([{ name }]),
  });
  if (!res.ok) throw new Error("프로젝트 생성 실패");
  const rows = await res.json();
  return rows[0];
}

// 프로젝트 삭제 — room_types/order_items/project_expenses/project_milestones가 전부
// FK CASCADE로 함께 삭제됨(Free 플랜, 백업/PITR 없음 — 되돌릴 수 없으니 호출부에서 반드시 이중 확인할 것)
async function deleteProjectRemote(id) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/projects?id=eq.${id}`, {
    method: "DELETE",
    headers: sbHeaders,
  });
  if (!res.ok) throw new Error("프로젝트 삭제 실패");
}

function toDateStr(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(
    d.getDate()
  ).padStart(2, "0")}`;
}

const ZOOM = {
  month: { pxPerDay: 3.2, label: "월" },
  week: { pxPerDay: 9, label: "주" },
  day: { pxPerDay: 36, label: "일" },
};

const WEEKDAY_KO = ["일", "월", "화", "수", "목", "금", "토"];

const STATUS_STYLE = {
  완료: { bg: "#E1F5EE", border: "#0F6E56", text: "#085041" },
  진행중: { bg: "#E6F1FB", border: "#185FA5", text: "#0C447C" },
  지연: { bg: "#FCEBEB", border: "#A32D2D", text: "#791F1F" },
  예정: { bg: "#F1EFE8", border: "#888780", text: "#5F5E5A" },
};

function parseDate(s) {
  if (!s) return null;
  const d = new Date(s + "T00:00:00");
  return isNaN(d.getTime()) ? null : d;
}

function fmt(d) {
  if (!d) return "-";
  const date = d instanceof Date ? d : new Date(d);
  if (isNaN(date.getTime())) return "-";
  return `${date.getFullYear()}.${String(date.getMonth() + 1).padStart(2, "0")}.${String(
    date.getDate()
  ).padStart(2, "0")}`;
}

function daysBetween(a, b) {
  return Math.round((b - a) / 86400000);
}

// 기간이 겹치는 마일스톤들에 자동으로 서로 다른 레인(줄)을 배정 — 겹치지 않으면 전부 0번 레인,
// 겹치면 그리디하게 빈 레인을 찾아 배정(고전적인 구간 스케줄링 레인 배정)
function assignLanes(dated) {
  const sorted = [...dated].sort((a, b) => a.s - b.s || a.e - b.e);
  const laneEnds = []; // 각 레인에 마지막으로 배정된 항목의 종료시각(ms)
  const laneOf = new Map();
  for (const item of sorted) {
    const startMs = item.s.getTime();
    let lane = laneEnds.findIndex((end) => end <= startMs);
    if (lane === -1) {
      lane = laneEnds.length;
      laneEnds.push(item.e.getTime());
    } else {
      laneEnds[lane] = item.e.getTime();
    }
    laneOf.set(item.m.id, lane);
  }
  return { laneOf, laneCount: Math.max(laneEnds.length, 1) };
}

function computeStatus(m, today) {
  const s = parseDate(m.actual_start_date) || parseDate(m.planned_start_date);
  const e = parseDate(m.actual_end_date) || parseDate(m.planned_end_date);
  if (!s || !e) return "예정";
  if (e < today) return "완료";
  if (s <= today && today <= e) return "진행중";
  return "예정";
}

// 최종 상태 — 완료 토글이 켜져 있으면 무조건 완료, 아니면 날짜 기반 자동 판정
function effectiveStatus(m, today) {
  if (m.is_completed) return "완료";
  return computeStatus(m, today);
}

// 마일스톤 진행률 — 완료 토글이면 100%, 아니면 상태 기반 추정치
function milestoneProgress(m, today) {
  if (m.is_completed) return 100;
  const status = computeStatus(m, today);
  if (status === "진행중" || status === "지연") return 50;
  return 0;
}

// "현장 설치·오픈바이징 마감" 마일스톤의 종료일(실제 종료일 우선, 없으면 계획 종료일)
// — "지난 프로젝트" 판정 기준(오픈예정일 대신 실제 오픈바이징 일정 사용)
function getOpenMilestoneEnd(pMilestones) {
  const m = pMilestones.find((x) => x.name === "현장 설치·오픈바이징 마감");
  if (!m) return null;
  return parseDate(m.actual_end_date) || parseDate(m.planned_end_date);
}

// 프로젝트 전체 공정률 — 마일스톤 weight로 가중평균
function projectProgress(pMilestones, today) {
  let totalWeight = 0;
  let sum = 0;
  for (const m of pMilestones) {
    const p = milestoneProgress(m, today);
    const w = Number(m.weight) || 1;
    sum += p * w;
    totalWeight += w;
  }
  return totalWeight ? Math.round(sum / totalWeight) : 0;
}

export default function ScheduleDashboard({ onOpenInOrderManager } = {}) {
  const [projects, setProjects] = useState([]);
  const [milestones, setMilestones] = useState([]);
  const [loading, setLoading] = useState(true);
  const [refreshing, setRefreshing] = useState(false);
  const [syncingCalendar, setSyncingCalendar] = useState(false);
  const [calendarSyncMessage, setCalendarSyncMessage] = useState(null); // {ok, text}
  const [error, setError] = useState(null);
  const [zoom, setZoom] = useState("day");
  const [expanded, setExpanded] = useState(new Set());
  const [editingProjectId, setEditingProjectId] = useState(null); // 프로젝트명 인라인 수정 중인 항목
  const [editingOpenDateId, setEditingOpenDateId] = useState(null); // 오픈예정일 인라인 수정 중인 항목
  const [hovered, setHovered] = useState(null);
  const [dragState, setDragState] = useState(null); // 마일스톤 바 드래그(날짜 변경)
  const [draggedProjectId, setDraggedProjectId] = useState(null); // 프로젝트 행 순서 드래그
  const [rangeSelect, setRangeSelect] = useState(null); // 빈 타임라인 영역 드래그 선택(신규 일정 입력)
  const [rangePopover, setRangePopover] = useState(null); // 드래그 선택 완료 후 마일스톤 선택 팝업
  const [showPastProjects, setShowPastProjects] = useState(false); // 오픈예정일이 지난 프로젝트는 기본적으로 접어서 가시성을 높임
  const rangeMovedRef = useRef(false); // 방금 실제로 드래그했는지(단순 클릭과 구분)
  const scrollRef = useRef(null);
  const bottomScrollRef = useRef(null); // 하단 고정 가로 스크롤바 — 트랙패드 없이 마우스만 쓰는 사람도 좌우 이동 가능하게
  const syncingScrollRef = useRef(false); // 두 스크롤바가 서로 onScroll을 트리거해 무한루프 도는 것 방지

  useEffect(() => {
    fetchAll()
      .then(({ projects, milestones }) => {
        setProjects(projects);
        setMilestones(milestones);
      })
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, []);

  // 새로고침 버튼 — 페이지 전체를 새로고침하지 않고 Supabase 데이터만 다시 불러옴
  // (다른 사람이 새 프로젝트를 추가했거나, 캘린더 자동 동기화로 일정이 바뀐 경우 등)
  function refresh() {
    setRefreshing(true);
    fetchAll()
      .then(({ projects, milestones }) => {
        setProjects(projects);
        setMilestones(milestones);
        setError(null);
      })
      .catch((e) => alert("새로고침에 실패했어요: " + e.message))
      .finally(() => setRefreshing(false));
  }

  // "지금 동기화" — 매일 오전 7시 자동으로 도는 캘린더 동기화를 수동으로 즉시 실행.
  // 동기화 자체는 Apps Script(구글 캘린더 → Supabase)가 처리하고, 끝나면 여기서 최신 마일스톤을 다시 불러온다.
  function handleSyncCalendar() {
    setSyncingCalendar(true);
    setCalendarSyncMessage(null);
    syncCalendarNow()
      .then(() => {
        setCalendarSyncMessage({ ok: true, text: "캘린더 동기화를 완료했어요." });
        return fetchAll().then(({ projects, milestones }) => {
          setProjects(projects);
          setMilestones(milestones);
        });
      })
      .catch((e) => setCalendarSyncMessage({ ok: false, text: "동기화 실패: " + e.message }))
      .finally(() => setSyncingCalendar(false));
  }

  function updateMilestoneField(milestoneId, field, value) {
    setMilestones((prev) =>
      prev.map((m) => (m.id === milestoneId ? { ...m, [field]: value } : m))
    );
  }

  function saveMilestoneField(milestoneId, field, value) {
    persistMilestone(milestoneId, { [field]: value }).catch(() => {
      alert("저장에 실패했어요. 네트워크 상태를 확인해주세요.");
    });
  }

  // "+ 직접입력" — 템플릿에 없는 커스텀 마일스톤을 프로젝트에 하나 추가(맨 아래, weight 1)
  function addCustomMilestone(projectId, pMilestones) {
    const nextSort =
      (pMilestones.length ? Math.max(...pMilestones.map((m) => m.sort_order)) : 0) + 1;
    createMilestone({
      project_id: projectId,
      template_id: null,
      name: "새 항목",
      sort_order: nextSort,
      weight: 1,
    })
      .then((row) => setMilestones((prev) => [...prev, row]))
      .catch(() => alert("마일스톤 추가에 실패했어요. 네트워크 상태를 확인해주세요."));
  }

  // 직접 추가한 마일스톤만 삭제 가능(템플릿에서 온 기본 6개는 보호)
  function deleteCustomMilestone(milestoneId) {
    if (!window.confirm("이 마일스톤을 삭제할까요? 되돌릴 수 없습니다.")) return;
    const prevMilestones = milestones;
    setMilestones((prev) => prev.filter((m) => m.id !== milestoneId));
    deleteMilestoneRemote(milestoneId).catch(() => {
      setMilestones(prevMilestones);
      alert("삭제에 실패했어요. 네트워크 상태를 확인해주세요.");
    });
  }

  const [projectActionBusy, setProjectActionBusy] = useState(false); // 생성/삭제 진행 중 버튼 비활성화용

  // "+ 새 프로젝트" — Supabase에 프로젝트 행을 만들고 기본 마일스톤 5단계(milestone_templates)를 자동으로 채운다.
  async function handleCreateProject() {
    const name = window.prompt("새 프로젝트 이름을 입력하세요");
    if (!name || !name.trim()) return;
    setProjectActionBusy(true);
    try {
      const project = await createProjectRemote(name.trim());
      setProjects((prev) => [...prev, project]);
      try {
        const templates = await fetchMilestoneTemplates();
        const created = await Promise.all(
          templates.map((t) =>
            createMilestone({
              project_id: project.id,
              template_id: t.id,
              name: t.name,
              sort_order: t.sort_order,
              weight: t.weight,
            })
          )
        );
        setMilestones((prev) => [...prev, ...created]);
      } catch (tplErr) {
        alert(
          "프로젝트는 생성됐지만 기본 마일스톤 5단계를 자동으로 채우지 못했어요. '+ 직접입력'으로 추가해주세요."
        );
      }
    } catch (err) {
      alert("프로젝트 생성에 실패했어요: " + err.message);
    } finally {
      setProjectActionBusy(false);
    }
  }

  // 프로젝트 삭제 — 발주품목/룸타입/지출까지 전부 영구 삭제되므로 이름을 그대로 입력해야만 진행되는
  // 이중 확인을 둔다(Free 플랜은 백업이 없어 되돌릴 수 없음, 과거 실수로 데이터 소실된 전례가 있음)
  function handleDeleteProject(project) {
    if (
      !window.confirm(
        `"${project.name}" 프로젝트를 삭제할까요?\n\n연결된 발주 품목·룸타입·지출·마일스톤이 전부 영구 삭제되며 되돌릴 수 없습니다(백업 없음).`
      )
    )
      return;
    const typed = window.prompt(
      `정말 삭제하려면 프로젝트명을 정확히 입력하세요: "${project.name}"`
    );
    if (typed !== project.name) {
      if (typed !== null) alert("입력한 이름이 일치하지 않아 취소했어요.");
      return;
    }
    const prevProjects = projects;
    const prevMilestones = milestones;
    setProjects((prev) => prev.filter((p) => p.id !== project.id));
    setMilestones((prev) => prev.filter((m) => m.project_id !== project.id));
    deleteProjectRemote(project.id).catch((err) => {
      setProjects(prevProjects);
      setMilestones(prevMilestones);
      alert("삭제에 실패했어요: " + err.message);
    });
  }

  // 프로젝트 행을 드래그해서 목록 순서를 수동으로 바꿈 — 처음 드래그하는 순간부터
  // 그 시점의 전체 순서를 manual_sort_order로 고정하고, 이후엔 자동(일정순) 정렬 대신 이 값을 우선함
  function handleDropReorder(targetProjectId) {
    if (!draggedProjectId || draggedProjectId === targetProjectId) {
      setDraggedProjectId(null);
      return;
    }
    const currentOrder = sortedProjects.map((p) => p.id);
    const fromIdx = currentOrder.indexOf(draggedProjectId);
    const toIdx = currentOrder.indexOf(targetProjectId);
    setDraggedProjectId(null);
    if (fromIdx === -1 || toIdx === -1) return;
    const reordered = [...currentOrder];
    reordered.splice(fromIdx, 1);
    reordered.splice(toIdx, 0, draggedProjectId);
    setProjects((prev) =>
      prev.map((p) => {
        const idx = reordered.indexOf(p.id);
        return idx === -1 ? p : { ...p, manual_sort_order: idx };
      })
    );
    reordered.forEach((id, idx) => {
      persistProject(id, { manual_sort_order: idx }).catch(() => {
        alert("순서 저장에 실패했어요. 네트워크 상태를 확인해주세요.");
      });
    });
  }

  function toggleMilestoneComplete(milestoneId) {
    const target = milestones.find((m) => m.id === milestoneId);
    const nextCompleted = !target?.is_completed;
    setMilestones((prev) =>
      prev.map((m) =>
        m.id !== milestoneId
          ? m
          : {
              ...m,
              is_completed: nextCompleted,
              completed_at: nextCompleted ? new Date().toISOString() : null,
            }
      )
    );
    persistMilestone(milestoneId, {
      is_completed: nextCompleted,
      completed_at: nextCompleted ? new Date().toISOString() : null,
    }).catch(() => {
      // 저장 실패 시 화면 상태를 원래대로 되돌림
      setMilestones((prev) =>
        prev.map((m) =>
          m.id === milestoneId
            ? { ...m, is_completed: !nextCompleted, completed_at: target?.completed_at ?? null }
            : m
        )
      );
      alert("저장에 실패했어요. 네트워크 상태를 확인해주세요.");
    });
  }

  // 간트 바를 지워서 일정을 미입력 상태로 되돌림 (드래그로 만든 일정을 빠르게 삭제할 때 사용)
  function clearMilestoneDates(milestoneId) {
    const target = milestones.find((m) => m.id === milestoneId);
    if (!target) return;
    const prevStart = target.actual_start_date;
    const prevEnd = target.actual_end_date;
    setMilestones((prev) =>
      prev.map((m) =>
        m.id === milestoneId ? { ...m, actual_start_date: null, actual_end_date: null } : m
      )
    );
    persistMilestone(milestoneId, { actual_start_date: null, actual_end_date: null }).catch(() => {
      setMilestones((prev) =>
        prev.map((m) =>
          m.id === milestoneId
            ? { ...m, actual_start_date: prevStart, actual_end_date: prevEnd }
            : m
        )
      );
      alert("삭제에 실패했어요. 네트워크 상태를 확인해주세요.");
    });
  }

  const today = useMemo(() => new Date(), []);

  // 이번주(월요일~일요일) 범위 — "이번주 진행되는 현장" 강조 판정에 사용
  const { weekStart, weekEnd } = useMemo(() => {
    const d = new Date(today);
    const day = d.getDay(); // 0=일 ... 6=토
    const diffToMonday = day === 0 ? -6 : 1 - day;
    const start = new Date(d);
    start.setDate(d.getDate() + diffToMonday);
    start.setHours(0, 0, 0, 0);
    const end = new Date(start);
    end.setDate(start.getDate() + 6);
    end.setHours(23, 59, 59, 999);
    return { weekStart: start, weekEnd: end };
  }, [today]);

  const milestonesByProject = useMemo(() => {
    const map = {};
    for (const m of milestones) {
      if (!map[m.project_id]) map[m.project_id] = [];
      map[m.project_id].push(m);
    }
    return map;
  }, [milestones]);

  // 오픈바이징 마감 일정이 "이번주 시작(월요일)"보다 이전인 프로젝트 수 — 지난 프로젝트 토글 버튼 표시에 사용
  // (오늘 기준으로 하면 이번주 안에 끝난 일정까지 "지난 프로젝트"로 숨겨져 "이번주 강조"와 충돌하므로 주 단위로 판정)
  const pastProjectCount = useMemo(() => {
    let count = 0;
    for (const p of projects) {
      const end = getOpenMilestoneEnd(milestonesByProject[p.id] || []);
      if (end && end < weekStart) count++;
    }
    return count;
  }, [projects, milestonesByProject, weekStart]);

  const { rangeStart, rangeEnd } = useMemo(() => {
    let allDates = [today];
    for (const m of milestones) {
      const s = parseDate(m.actual_start_date) || parseDate(m.planned_start_date);
      const e = parseDate(m.actual_end_date) || parseDate(m.planned_end_date);
      if (s) allDates.push(s);
      if (e) allDates.push(e);
    }
    for (const p of projects) {
      const t = parseDate(p.target_open_date);
      if (t) allDates.push(t);
    }
    const min = new Date(Math.min(...allDates.map((d) => d.getTime())));
    const max = new Date(Math.max(...allDates.map((d) => d.getTime())));
    min.setDate(min.getDate() - 10);
    max.setDate(max.getDate() + 10);
    return { rangeStart: min, rangeEnd: max };
  }, [milestones, projects, today]);

  const pxPerDay = ZOOM[zoom].pxPerDay;
  const totalDays = daysBetween(rangeStart, rangeEnd);
  const totalWidth = totalDays * pxPerDay;

  const monthTicks = useMemo(() => {
    const ticks = [];
    const cur = new Date(rangeStart.getFullYear(), rangeStart.getMonth(), 1);
    while (cur <= rangeEnd) {
      ticks.push(new Date(cur));
      cur.setMonth(cur.getMonth() + 1);
    }
    return ticks;
  }, [rangeStart, rangeEnd]);

  const dayTicks = useMemo(() => {
    if (zoom !== "day") return [];
    const ticks = [];
    const cur = new Date(rangeStart);
    while (cur <= rangeEnd) {
      ticks.push(new Date(cur));
      cur.setDate(cur.getDate() + 1);
    }
    return ticks;
  }, [zoom, rangeStart, rangeEnd]);

  // 프로젝트 전체 담당자 — 마일스톤마다 다른 담당자가 붙어있을 수 있어 중복 제거 후 전부 나열
  function managersOf(pMilestones) {
    const names = [...new Set(pMilestones.map((m) => m.manager).filter(Boolean))];
    return names.length ? names.join(', ') : null;
  }

  // 프로젝트명 인라인 수정
  function updateProjectField(projectId, field, value) {
    setProjects((prev) => prev.map((p) => (p.id === projectId ? { ...p, [field]: value } : p)));
  }
  function saveProjectField(projectId, field, value) {
    persistProject(projectId, { [field]: value }).catch(() => {
      alert("저장에 실패했어요. 네트워크 상태를 확인해주세요.");
    });
  }

  // 프로젝트 목록 정렬 — 드래그로 수동 순서를 한 번이라도 지정했으면 그 순서(manual_sort_order)를
  // 우선 적용하고, 아직 아무도 안 건드렸으면 기존처럼 실제 일정(마일스톤 시작일) 기준 자동 정렬
  // (일정이 아예 없는 프로젝트는 맨 아래로)
  const getEarliestStart = useMemo(() => {
    return (id) => {
      const ms = milestonesByProject[id] || [];
      const dates = ms
        .map((m) => parseDate(m.actual_start_date) || parseDate(m.planned_start_date))
        .filter(Boolean);
      return dates.length ? Math.min(...dates.map((d) => d.getTime())) : Infinity;
    };
  }, [milestonesByProject]);

  const sortedProjects = useMemo(() => {
    const hasManualOrder = projects.some((p) => p.manual_sort_order != null);
    return [...projects].sort((a, b) => {
      if (hasManualOrder) {
        const av = a.manual_sort_order ?? Infinity;
        const bv = b.manual_sort_order ?? Infinity;
        if (av !== bv) return av - bv;
      }
      return getEarliestStart(a.id) - getEarliestStart(b.id);
    });
  }, [projects, getEarliestStart]);

  function xFor(date) {
    return daysBetween(rangeStart, date) * pxPerDay;
  }

  function dateFromX(x) {
    const d = new Date(rangeStart);
    d.setDate(d.getDate() + Math.round(x / pxPerDay));
    return d;
  }

  function toggleExpand(id) {
    setExpanded((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }


  // 마일스톤 바를 좌우로 끌면 픽셀 이동량을 날짜로 환산 — 놓는 순간 실제 시작/종료일로 저장
  useEffect(() => {
    if (!dragState) return;
    function handleMove(evt) {
      const deltaX = evt.clientX - dragState.startX;
      const deltaDays = Math.round(deltaX / pxPerDay);
      setDragState((prev) => (prev ? { ...prev, deltaDays } : prev));
    }
    function handleUp() {
      setDragState((prev) => {
        if (prev && prev.deltaDays) {
          const newStart = new Date(prev.origStart);
          newStart.setDate(newStart.getDate() + prev.deltaDays);
          const newEnd = new Date(prev.origEnd);
          newEnd.setDate(newEnd.getDate() + prev.deltaDays);
          const newStartStr = toDateStr(newStart);
          const newEndStr = toDateStr(newEnd);
          updateMilestoneField(prev.milestoneId, "actual_start_date", newStartStr);
          updateMilestoneField(prev.milestoneId, "actual_end_date", newEndStr);
          saveMilestoneField(prev.milestoneId, "actual_start_date", newStartStr);
          saveMilestoneField(prev.milestoneId, "actual_end_date", newEndStr);
        }
        return null;
      });
    }
    window.addEventListener("mousemove", handleMove);
    window.addEventListener("mouseup", handleUp);
    return () => {
      window.removeEventListener("mousemove", handleMove);
      window.removeEventListener("mouseup", handleUp);
    };
  }, [dragState, pxPerDay]);

  // 빈 타임라인 영역을 드래그로 선택 — 놓는 순간 어느 마일스톤에 적용할지 팝업으로 확인
  useEffect(() => {
    if (!rangeSelect) return;
    function handleMove(evt) {
      const x = evt.clientX - rangeSelect.containerLeft;
      setRangeSelect((prev) =>
        prev
          ? { ...prev, curX: x, moved: prev.moved || Math.abs(x - prev.startX) > 4 }
          : prev
      );
    }
    function handleUp() {
      setRangeSelect((prev) => {
        if (prev && prev.moved) {
          rangeMovedRef.current = true;
          const x1 = Math.min(prev.startX, prev.curX);
          const x2 = Math.max(prev.startX, prev.curX);
          setRangePopover({
            projectId: prev.projectId,
            startDate: toDateStr(dateFromX(x1)),
            endDate: toDateStr(dateFromX(x2)),
          });
        }
        return null;
      });
    }
    window.addEventListener("mousemove", handleMove);
    window.addEventListener("mouseup", handleUp);
    return () => {
      window.removeEventListener("mousemove", handleMove);
      window.removeEventListener("mouseup", handleUp);
    };
  }, [rangeSelect, pxPerDay]);

  // 마일스톤 바 하나를 그림 — top(px)을 인자로 받아 기본 트랙/서브트랙 레인 어디에도 재사용
  function renderMilestoneBar(m, s, e, top) {
    const status = effectiveStatus(m, today);
    const style = STATUS_STYLE[status];
    const isDragging = dragState && dragState.milestoneId === m.id;
    const dragDays = isDragging ? dragState.deltaDays || 0 : 0;
    const dispS = dragDays ? new Date(s.getTime() + dragDays * 86400000) : s;
    const dispE = dragDays ? new Date(e.getTime() + dragDays * 86400000) : e;
    const left = xFor(dispS);
    const width = Math.max(xFor(dispE) - xFor(dispS), 4);
    const shortDate = (d) =>
      `${String(d.getMonth() + 1).padStart(2, "0")}.${String(d.getDate()).padStart(2, "0")}`;
    return (
      <div
        key={m.id}
        className="absolute group/bar"
        style={{ left, top, width: Math.max(width, 46) }}
      >
        <div
          className={`relative rounded-md border text-[11px] pl-5 pr-5 flex items-center overflow-visible whitespace-nowrap select-none ${
            isDragging ? "cursor-grabbing shadow-md" : "cursor-grab"
          }`}
          style={{
            width,
            height: 24,
            background: style.bg,
            borderColor: style.border,
            color: style.text,
          }}
          onMouseEnter={(evt) =>
            !dragState &&
            setHovered({ m, s, e, status, x: evt.clientX, y: evt.clientY })
          }
          onMouseMove={(evt) =>
            !dragState &&
            setHovered((prev) =>
              prev && prev.m.id === m.id
                ? { ...prev, x: evt.clientX, y: evt.clientY }
                : prev
            )
          }
          onMouseLeave={() => setHovered(null)}
          onMouseDown={(evt) => {
            evt.stopPropagation();
            evt.preventDefault();
            setDragState({
              milestoneId: m.id,
              startX: evt.clientX,
              origStart: s,
              origEnd: e,
              deltaDays: 0,
            });
          }}
          onClick={(evt) => evt.stopPropagation()}
          title="끌어서 일정 변경"
        >
          {/* 완료 체크 — 클릭 한 번으로 이 마일스톤을 완료/미완료 전환 */}
          <button
            type="button"
            className="absolute left-1 top-1/2 -translate-y-1/2 w-3.5 h-3.5 rounded-sm border flex items-center justify-center flex-shrink-0"
            style={{
              background: m.is_completed ? "#10b981" : "rgba(255,255,255,0.6)",
              borderColor: m.is_completed ? "#059669" : style.border,
            }}
            onMouseDown={(evt) => evt.stopPropagation()}
            onClick={(evt) => {
              evt.stopPropagation();
              toggleMilestoneComplete(m.id);
            }}
            title={m.is_completed ? "완료됨 — 클릭해서 미완료로" : "클릭해서 완료 처리"}
          >
            {m.is_completed && (
              <svg viewBox="0 0 12 12" className="w-3 h-3 text-white">
                <path
                  d="M2.5 6.5L5 9L9.5 3.5"
                  stroke="currentColor"
                  strokeWidth="1.5"
                  fill="none"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            )}
          </button>
          {width > 80 ? m.name : ""}
          {/* 삭제 — 평소엔 숨겨져있다가 바에 마우스 올리면 나타남(드래그로 만든 일정 빠르게 지우기) */}
          <button
            type="button"
            className="absolute right-0.5 top-1/2 -translate-y-1/2 w-4 h-4 rounded-full bg-white border border-slate-300 text-slate-500 hover:bg-rose-50 hover:border-rose-300 hover:text-rose-600 items-center justify-center text-[10px] leading-none opacity-0 group-hover/bar:opacity-100 transition-opacity hidden sm:flex"
            onMouseDown={(evt) => evt.stopPropagation()}
            onClick={(evt) => {
              evt.stopPropagation();
              if (window.confirm(`"${m.name}" 일정을 지울까요? (날짜만 지워지고 마일스톤 자체는 남습니다)`)) {
                clearMilestoneDates(m.id);
              }
            }}
            title="일정 삭제"
          >
            ✕
          </button>
        </div>
        <div className="text-[10px] text-slate-400 mt-0.5 whitespace-nowrap">
          {shortDate(dispS)}~{shortDate(dispE)}
        </div>
      </div>
    );
  }

  const todayX = xFor(today);

  // 줌 레벨이 바뀌거나 데이터가 로드되면 오늘 날짜가 화면 왼쪽 근처에 오도록 자동 스크롤
  useEffect(() => {
    if (scrollRef.current && !loading) {
      scrollRef.current.scrollLeft = Math.max(todayX - 120, 0);
      if (bottomScrollRef.current) bottomScrollRef.current.scrollLeft = scrollRef.current.scrollLeft;
    }
  }, [zoom, loading, todayX]);

  // 본문 타임라인과 하단 고정 스크롤바의 scrollLeft를 서로 동기화
  function handleTimelineScroll() {
    if (syncingScrollRef.current) return;
    syncingScrollRef.current = true;
    if (bottomScrollRef.current && scrollRef.current) {
      bottomScrollRef.current.scrollLeft = scrollRef.current.scrollLeft;
    }
    syncingScrollRef.current = false;
  }
  function handleBottomScrollbarScroll() {
    if (syncingScrollRef.current) return;
    syncingScrollRef.current = true;
    if (bottomScrollRef.current && scrollRef.current) {
      scrollRef.current.scrollLeft = bottomScrollRef.current.scrollLeft;
    }
    syncingScrollRef.current = false;
  }


  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50">
        <div className="text-slate-400 text-sm">공정 데이터를 불러오는 중...</div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50">
        <div className="text-rose-600 text-sm">오류: {error}</div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900">
      <div className="border-b border-slate-200 bg-white">
        <div className="px-6 py-4 flex items-center justify-between">
          <div>
            <h1 className="text-lg font-medium tracking-tight text-slate-900">
              오픈바이징 공정표
            </h1>
            <p className="text-xs text-slate-500 mt-0.5">
              프로젝트 {projects.length}개 · 오늘 {fmt(today)}
              {pastProjectCount > 0 && !showPastProjects && (
                <> · 지난 프로젝트 {pastProjectCount}개 숨김</>
              )}
            </p>
          </div>
          <div className="flex items-center gap-3">
            {pastProjectCount > 0 && (
              <button
                onClick={() => setShowPastProjects((v) => !v)}
                className={`flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md border ${
                  showPastProjects
                    ? "border-slate-300 text-slate-700 bg-slate-100"
                    : "border-slate-200 text-slate-500 bg-white hover:bg-slate-50"
                }`}
                title="현장 설치·오픈바이징 마감 일정이 오늘보다 지난 프로젝트를 목록에서 접거나 펼칩니다"
              >
                {showPastProjects ? "지난 프로젝트 숨기기" : `지난 프로젝트 보기 (${pastProjectCount})`}
              </button>
            )}
            <button
              onClick={handleCreateProject}
              disabled={projectActionBusy}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md border border-amber-300 text-amber-700 bg-amber-50 hover:bg-amber-100 disabled:opacity-50 disabled:cursor-not-allowed"
              title="새 프로젝트를 만들고 기본 마일스톤 5단계를 자동으로 채웁니다"
            >
              + 새 프로젝트
            </button>
            <button
              onClick={handleSyncCalendar}
              disabled={syncingCalendar}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md border border-indigo-200 text-indigo-700 bg-indigo-50 hover:bg-indigo-100 disabled:opacity-50 disabled:cursor-not-allowed"
              title='구글 캘린더의 [OV] 일정을 "현장 설치·오픈바이징 마감" 마일스톤에 지금 바로 반영합니다 (평소엔 매일 오전 7시 자동 실행)'
            >
              <span className={syncingCalendar ? "inline-block animate-spin" : "inline-block"}>📅</span>
              {syncingCalendar ? "동기화 중..." : "캘린더 지금 동기화"}
            </button>
            <button
              onClick={refresh}
              disabled={refreshing}
              className="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md border border-slate-200 text-slate-600 hover:bg-slate-50 disabled:opacity-50 disabled:cursor-not-allowed"
              title="Supabase에서 최신 데이터를 다시 불러옵니다"
            >
              <span className={refreshing ? "inline-block animate-spin" : "inline-block"}>
                ⟳
              </span>
              새로고침
            </button>
            <div className="flex items-center gap-1 bg-slate-100 rounded-lg p-1">
              {Object.keys(ZOOM).map((z) => (
                <button
                  key={z}
                  onClick={() => setZoom(z)}
                  className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors ${
                    zoom === z
                      ? "bg-white text-slate-900 shadow-sm"
                      : "text-slate-500 hover:text-slate-700"
                  }`}
                >
                  {ZOOM[z].label}
                </button>
              ))}
            </div>
          </div>
        </div>
        {calendarSyncMessage && (
          <div
            className={`px-6 pb-3 text-xs ${
              calendarSyncMessage.ok ? "text-emerald-700" : "text-rose-600"
            }`}
          >
            {calendarSyncMessage.text}
          </div>
        )}
        <div className="px-6 pb-3 flex items-center gap-4 text-xs text-slate-500">
          {Object.entries(STATUS_STYLE).map(([label, s]) => (
            <div key={label} className="flex items-center gap-1.5">
              <span
                className="inline-block w-2.5 h-2.5 rounded-sm"
                style={{ background: s.border }}
              />
              {label}
            </div>
          ))}
          <div className="flex items-center gap-1.5 ml-2 border-l border-slate-200 pl-4">
            <span className="inline-block w-2.5 h-2.5 rounded-full bg-rose-500" />
            오늘
          </div>
          <div className="flex items-center gap-1.5">
            <span className="inline-block w-2 h-2 rotate-45 bg-amber-400 border border-amber-600" />
            오픈예정일
          </div>
          <div className="flex items-center gap-1.5 ml-2 border-l border-slate-200 pl-4">
            <span className="inline-block w-2.5 h-2.5 rounded-sm bg-indigo-100 border border-indigo-300" />
            이번주 진행중
          </div>
        </div>
      </div>

      <div className="overflow-x-auto" ref={scrollRef} onScroll={handleTimelineScroll}>
        <div style={{ minWidth: totalWidth + 280, paddingBottom: 16 }}>
          <div
            className="flex border-b border-slate-200 bg-white sticky top-0 z-30"
          >
            <div className="w-[280px] flex-shrink-0 border-r border-slate-200 sticky left-0 z-10 bg-white" />
            <div
              className="relative flex-1"
              style={{ width: totalWidth, height: zoom === "day" ? 40 : 32 }}
            >
              {zoom === "day"
                ? dayTicks.map((t, i) => {
                    const dow = t.getDay();
                    const isSun = dow === 0;
                    const isSat = dow === 6;
                    return (
                      <div
                        key={i}
                        className={`absolute top-0 bottom-0 border-l text-[10px] text-center leading-tight pt-1 ${
                          isSun || isSat
                            ? "border-slate-100 bg-slate-50/80"
                            : "border-slate-100"
                        }`}
                        style={{ left: xFor(t), width: pxPerDay }}
                      >
                        <div
                          className={
                            isSun
                              ? "text-rose-400 font-medium"
                              : isSat
                              ? "text-blue-400 font-medium"
                              : "text-slate-500 font-medium"
                          }
                        >
                          {t.getMonth() + 1}/{t.getDate()}
                        </div>
                        <div
                          className={
                            isSun
                              ? "text-rose-300"
                              : isSat
                              ? "text-blue-300"
                              : "text-slate-400"
                          }
                        >
                          {WEEKDAY_KO[dow]}
                        </div>
                      </div>
                    );
                  })
                : monthTicks.map((t, i) => (
                    <div
                      key={i}
                      className="absolute top-0 bottom-0 border-l border-slate-100 text-[11px] text-slate-400 pl-1.5 pt-1.5"
                      style={{ left: xFor(t) }}
                    >
                      {t.getFullYear()}.{String(t.getMonth() + 1).padStart(2, "0")}
                    </div>
                  ))}
            </div>
          </div>

          {sortedProjects.map((p) => {
            const pMilestones = (milestonesByProject[p.id] || []).sort(
              (a, b) => a.sort_order - b.sort_order
            );
            const isOpen = expanded.has(p.id);
            const targetOpen = parseDate(p.target_open_date);

            const dated = pMilestones
              .map((m) => ({
                m,
                s: parseDate(m.actual_start_date) || parseDate(m.planned_start_date),
                e: parseDate(m.actual_end_date) || parseDate(m.planned_end_date),
              }))
              .filter((x) => x.s && x.e);

            let isDelayed = false;
            if (dated.length && targetOpen) {
              const lastEnd = dated.reduce((max, x) => (x.e > max ? x.e : max), dated[0].e);
              const totalSpan = daysBetween(rangeStart, targetOpen) || 1;
              const overrun = daysBetween(targetOpen, lastEnd);
              if (overrun > 0 && overrun / totalSpan > 0.2) isDelayed = true;
              if (lastEnd > targetOpen) isDelayed = true;
            }

            // 지난 프로젝트: "현장 설치·오픈바이징 마감" 마일스톤 종료일이 "이번주 시작(월요일)"보다 이전
            // (오늘 기준으로 하면 이번주 안에 끝난 일정까지 지난 프로젝트로 숨겨져 아래 "이번주 강조"와 충돌함)
            const openMilestoneEnd = getOpenMilestoneEnd(pMilestones);
            const isPast = !!(openMilestoneEnd && openMilestoneEnd < weekStart);
            if (isPast && !showPastProjects) return null;

            // 이번주(월~일) 진행중: 마일스톤 일정이 이번주와 겹치는 경우 (지난 프로젝트 여부와 무관하게 판정)
            const isThisWeek = dated.some(({ s, e }) => s <= weekEnd && e >= weekStart);

            const progress = projectProgress(pMilestones, today);

            // 겹치는 마일스톤은 자동으로 레인(줄)을 나눠서 전부 항상 보이게 함(클릭 불필요)
            const { laneOf, laneCount } = assignLanes(dated);
            const laneHeight = 34;
            const rowHeight = 62 + (laneCount - 1) * laneHeight;

            return (
              <div
                key={p.id}
                className={`border-b border-slate-100 ${
                  isPast ? "opacity-50 grayscale-[50%]" : ""
                } ${isThisWeek ? "bg-indigo-50/50" : ""}`}
              >
                <div
                  className={`flex hover:bg-slate-50/60 cursor-pointer ${
                    draggedProjectId && draggedProjectId !== p.id
                      ? "border-t-2 border-t-transparent hover:border-t-indigo-400"
                      : ""
                  } ${isThisWeek ? "border-l-2 border-l-indigo-400" : ""}`}
                  onClick={() => toggleExpand(p.id)}
                  onDragOver={(evt) => evt.preventDefault()}
                  onDrop={(evt) => {
                    evt.preventDefault();
                    handleDropReorder(p.id);
                  }}
                >
                  <div className="relative w-[280px] flex-shrink-0 border-r border-slate-200 px-4 py-3 flex items-start gap-2 sticky left-0 z-10 bg-white">
                    <span
                      draggable
                      onDragStart={(evt) => {
                        evt.stopPropagation();
                        setDraggedProjectId(p.id);
                      }}
                      onDragEnd={() => setDraggedProjectId(null)}
                      onClick={(evt) => evt.stopPropagation()}
                      title="드래그해서 순서 변경"
                      className="text-slate-300 hover:text-slate-500 text-xs mt-0.5 cursor-grab active:cursor-grabbing select-none"
                    >
                      ⠿
                    </span>
                    <span
                      className={`text-slate-300 text-xs mt-0.5 transition-transform ${
                        isOpen ? "rotate-90" : ""
                      }`}
                    >
                      ▶
                    </span>
                    <div className="min-w-0">
                      {dated.length > 0 ? (
                        <div className="text-[11px] font-medium text-indigo-600">
                          {fmt(
                            dated.reduce((min, x) => (x.s < min ? x.s : min), dated[0].s)
                          )}
                          {" ~ "}
                          {fmt(dated.reduce((max, x) => (x.e > max ? x.e : max), dated[0].e))}
                        </div>
                      ) : (
                        <div className="text-[11px] font-medium text-slate-400">
                          일정 미입력
                        </div>
                      )}
                      <div className="text-sm font-medium text-slate-800 flex items-center gap-1 group/name">
                        {editingProjectId === p.id ? (
                          <input
                            type="text"
                            autoFocus
                            value={p.name}
                            onClick={(evt) => evt.stopPropagation()}
                            onChange={(e) => updateProjectField(p.id, "name", e.target.value)}
                            onBlur={(e) => {
                              const v = e.target.value.trim() || "이름 없음";
                              updateProjectField(p.id, "name", v);
                              saveProjectField(p.id, "name", v);
                              setEditingProjectId(null);
                            }}
                            onKeyDown={(e) => {
                              if (e.key === "Enter") e.currentTarget.blur();
                              if (e.key === "Escape") setEditingProjectId(null);
                            }}
                            className="bg-white border border-indigo-300 focus:outline-none focus:ring-1 focus:ring-indigo-300 rounded px-1 py-0.5 text-sm font-medium w-full"
                          />
                        ) : (
                          <>
                            <span className="truncate">{p.name}</span>
                            {isThisWeek && (
                              <span className="flex-shrink-0 text-[10px] font-semibold text-indigo-600 bg-indigo-100 rounded px-1.5 py-0.5">
                                이번주
                              </span>
                            )}
                            <button
                              type="button"
                              onClick={(evt) => {
                                evt.stopPropagation();
                                setEditingProjectId(p.id);
                              }}
                              title="프로젝트명 수정"
                              className="text-slate-300 hover:text-indigo-500 opacity-0 group-hover/name:opacity-100 transition-opacity flex-shrink-0"
                            >
                              ✎
                            </button>
                            {onOpenInOrderManager && (
                              <button
                                type="button"
                                onClick={(evt) => {
                                  evt.stopPropagation();
                                  onOpenInOrderManager(p);
                                }}
                                title="발주 관리 탭에서 열기"
                                className="text-slate-300 hover:text-amber-600 opacity-0 group-hover/name:opacity-100 transition-opacity flex-shrink-0"
                              >
                                ↗
                              </button>
                            )}
                            <button
                              type="button"
                              onClick={(evt) => {
                                evt.stopPropagation();
                                handleDeleteProject(p);
                              }}
                              title="프로젝트 삭제"
                              className="text-slate-300 hover:text-rose-600 opacity-0 group-hover/name:opacity-100 transition-opacity flex-shrink-0"
                            >
                              🗑
                            </button>
                            {managersOf(pMilestones) && (
                              <span className="text-indigo-500 font-normal ml-1 truncate">
                                · {managersOf(pMilestones)}
                              </span>
                            )}
                          </>
                        )}
                      </div>
                      <div
                        className="text-[11px] text-slate-400 mt-0.5 flex items-center gap-1"
                        onClick={(evt) => evt.stopPropagation()}
                      >
                        오픈예정{" "}
                        {editingOpenDateId === p.id ? (
                          <input
                            type="date"
                            autoFocus
                            value={p.target_open_date || ""}
                            onChange={(e) => {
                              const v = e.target.value || null;
                              updateProjectField(p.id, "target_open_date", v);
                              saveProjectField(p.id, "target_open_date", v);
                            }}
                            onBlur={() => setEditingOpenDateId(null)}
                            onKeyDown={(e) => {
                              if (e.key === "Enter" || e.key === "Escape")
                                setEditingOpenDateId(null);
                            }}
                            className="bg-white border border-indigo-300 focus:outline-none focus:ring-1 focus:ring-indigo-300 rounded px-1 py-0 text-[11px] leading-tight"
                          />
                        ) : (
                          <button
                            type="button"
                            onClick={() => setEditingOpenDateId(p.id)}
                            title="클릭해서 오픈예정일 수정"
                            className="hover:text-indigo-500 underline decoration-dotted underline-offset-2"
                          >
                            {fmt(targetOpen)}
                          </button>
                        )}
                        {isDelayed && (
                          <span className="text-rose-600 font-medium ml-1">지연</span>
                        )}
                      </div>
                      <div className="flex items-center gap-1.5 mt-1">
                        <div className="flex-1 h-1.5 bg-slate-100 rounded-full overflow-hidden">
                          <div
                            className="h-full bg-indigo-400 rounded-full"
                            style={{ width: `${progress}%` }}
                          />
                        </div>
                        <span className="text-[10px] text-slate-400 tabular-nums">
                          {progress}%
                        </span>
                      </div>
                    </div>
                  </div>
                  <div
                    className="relative flex-1 cursor-crosshair"
                    style={{ width: totalWidth, height: rowHeight }}
                    onMouseDown={(evt) => {
                      if (evt.button !== 0) return;
                      const rect = evt.currentTarget.getBoundingClientRect();
                      const x = evt.clientX - rect.left;
                      setRangeSelect({
                        projectId: p.id,
                        containerLeft: rect.left,
                        startX: x,
                        curX: x,
                        moved: false,
                      });
                    }}
                    onClick={(evt) => {
                      // 드래그로 영역을 선택한 직후엔 행 펼침 토글이 같이 발동하지 않도록 막음
                      if (rangeMovedRef.current) {
                        rangeMovedRef.current = false;
                        evt.stopPropagation();
                      }
                    }}
                  >
                    {rangeSelect && rangeSelect.projectId === p.id && (
                      <div
                        className="absolute top-2 bg-indigo-200/50 border border-indigo-400 border-dashed rounded pointer-events-none"
                        style={{
                          left: Math.min(rangeSelect.startX, rangeSelect.curX),
                          width: Math.abs(rangeSelect.curX - rangeSelect.startX),
                          height: 24,
                        }}
                      />
                    )}
                    {(zoom === "day" ? dayTicks : monthTicks).map((t, i) => (
                      <div
                        key={`grid-${i}`}
                        className="absolute top-0 bottom-0 border-l border-slate-200"
                        style={{ left: xFor(t) }}
                      />
                    ))}
                    <div
                      className="absolute top-0 bottom-0 w-px bg-rose-300"
                      style={{ left: todayX }}
                    />
                    {targetOpen && (
                      <div
                        className="absolute top-1 flex flex-col items-center"
                        style={{ left: xFor(targetOpen) - 4 }}
                        title={`오픈예정일 ${fmt(targetOpen)}`}
                      >
                        <div className="w-2 h-2 rotate-45 bg-amber-400 border border-amber-600" />
                      </div>
                    )}
                    {dated.map(({ m, s, e }) =>
                      renderMilestoneBar(m, s, e, 10 + laneOf.get(m.id) * laneHeight)
                    )}
                    {dated.length === 0 && (
                      <div className="absolute left-3 top-4 text-[11px] text-slate-300">
                        일정 미입력
                      </div>
                    )}
                  </div>
                </div>

                {isOpen && (
                  <div
                    className="sticky left-0 bg-slate-50/70 border-t border-slate-100 overflow-x-auto"
                    style={{ width: "100vw", maxWidth: "100vw" }}
                  >
                    <table className="text-xs" style={{ minWidth: 900 }}>
                      <thead>
                        <tr className="text-slate-400 border-b border-slate-200">
                          <th className="text-left font-medium px-4 py-2 w-[180px] sticky left-0 z-10 bg-slate-100 border-r border-slate-200">
                            마일스톤
                          </th>
                          <th className="text-left font-medium px-2 py-2 w-[100px]">상태</th>
                          <th className="text-left font-medium px-2 py-2 w-[100px]">
                            진행률
                          </th>
                          <th className="text-left font-medium px-2 py-2 w-[110px]">
                            시작일
                          </th>
                          <th className="text-left font-medium px-2 py-2 w-[110px]">
                            종료일
                          </th>
                          <th className="text-left font-medium px-2 py-2 w-[90px]">
                            카테고리
                          </th>
                          <th className="text-left font-medium px-2 py-2 w-[90px]">
                            담당자
                          </th>
                          <th className="text-left font-medium px-4 py-2 w-[160px]">완료</th>
                          <th className="text-left font-medium px-2 py-2 w-[40px]"></th>
                        </tr>
                      </thead>
                      <tbody>
                        {pMilestones.map((m) => {
                          const status = effectiveStatus(m, today);
                          const style = STATUS_STYLE[status];
                          const mProgress = milestoneProgress(m, today);
                          const isCustom = !m.template_id;
                          return (
                            <tr key={m.id} className="border-b border-slate-100 last:border-0">
                              <td className="px-4 py-2.5 text-slate-700 sticky left-0 z-10 bg-slate-50 border-r border-slate-200">
                                <input
                                  type="text"
                                  value={m.name}
                                  onChange={(e) => updateMilestoneField(m.id, "name", e.target.value)}
                                  onBlur={(e) => {
                                    const v = e.target.value.trim() || "이름 없음";
                                    updateMilestoneField(m.id, "name", v);
                                    saveMilestoneField(m.id, "name", v);
                                  }}
                                  className="bg-transparent border border-transparent hover:border-slate-200 focus:border-indigo-300 focus:outline-none rounded px-1 py-0.5 text-xs font-medium w-full"
                                />
                              </td>
                              <td className="px-2 py-2.5">
                                <span
                                  className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-medium"
                                  style={{ background: style.bg, color: style.text }}
                                >
                                  {status}
                                </span>
                              </td>
                              <td className="px-2 py-2.5">
                                <div className="flex items-center gap-1.5">
                                  <div className="w-10 h-1.5 bg-slate-200 rounded-full overflow-hidden">
                                    <div
                                      className="h-full bg-indigo-400 rounded-full"
                                      style={{ width: `${mProgress}%` }}
                                    />
                                  </div>
                                  <span className="text-slate-400 tabular-nums">
                                    {mProgress}%
                                  </span>
                                </div>
                              </td>
                              <td className="px-2 py-2.5 text-slate-500">
                                <input
                                  type="date"
                                  value={m.actual_start_date || m.planned_start_date || ""}
                                  onChange={(e) => {
                                    const v = e.target.value || null;
                                    updateMilestoneField(m.id, "actual_start_date", v);
                                    saveMilestoneField(m.id, "actual_start_date", v);
                                  }}
                                  className="bg-transparent border border-transparent hover:border-slate-200 focus:border-indigo-300 focus:outline-none rounded px-1 py-0.5 text-[11px] w-[108px]"
                                />
                              </td>
                              <td className="px-2 py-2.5 text-slate-500">
                                <input
                                  type="date"
                                  value={m.actual_end_date || m.planned_end_date || ""}
                                  onChange={(e) => {
                                    const v = e.target.value || null;
                                    updateMilestoneField(m.id, "actual_end_date", v);
                                    saveMilestoneField(m.id, "actual_end_date", v);
                                  }}
                                  className="bg-transparent border border-transparent hover:border-slate-200 focus:border-indigo-300 focus:outline-none rounded px-1 py-0.5 text-[11px] w-[108px]"
                                />
                              </td>
                              <td className="px-2 py-2.5 text-slate-500">
                                <input
                                  type="text"
                                  value={m.category || ""}
                                  placeholder="-"
                                  onChange={(e) => updateMilestoneField(m.id, "category", e.target.value)}
                                  onBlur={(e) =>
                                    saveMilestoneField(m.id, "category", e.target.value || null)
                                  }
                                  className="bg-transparent border border-transparent hover:border-slate-200 focus:border-indigo-300 focus:outline-none rounded px-1 py-0.5 text-[11px] w-full"
                                />
                              </td>
                              <td className="px-2 py-2.5 text-slate-500">
                                <input
                                  type="text"
                                  value={m.manager || ""}
                                  placeholder="-"
                                  onChange={(e) => updateMilestoneField(m.id, "manager", e.target.value)}
                                  onBlur={(e) =>
                                    saveMilestoneField(m.id, "manager", e.target.value || null)
                                  }
                                  className="bg-transparent border border-transparent hover:border-slate-200 focus:border-indigo-300 focus:outline-none rounded px-1 py-0.5 text-[11px] w-full"
                                />
                              </td>
                              <td className="px-4 py-2.5">
                                <button
                                  type="button"
                                  onClick={() => toggleMilestoneComplete(m.id)}
                                  className={`inline-flex items-center gap-1.5 cursor-pointer hover:opacity-70 transition-opacity ${
                                    m.is_completed ? "text-emerald-700" : "text-slate-600"
                                  }`}
                                >
                                  <span
                                    className={`inline-block w-3.5 h-3.5 rounded-sm border flex-shrink-0 ${
                                      m.is_completed
                                        ? "bg-emerald-500 border-emerald-600"
                                        : "border-slate-300"
                                    }`}
                                  >
                                    {m.is_completed && (
                                      <svg viewBox="0 0 12 12" className="w-3.5 h-3.5 text-white">
                                        <path
                                          d="M2.5 6.5L5 9L9.5 3.5"
                                          stroke="currentColor"
                                          strokeWidth="1.5"
                                          fill="none"
                                          strokeLinecap="round"
                                          strokeLinejoin="round"
                                        />
                                      </svg>
                                    )}
                                  </span>
                                  {m.is_completed ? (
                                    <span>
                                      완료
                                      {m.completed_at && ` (${fmt(m.completed_at)})`}
                                    </span>
                                  ) : (
                                    <span>미완료</span>
                                  )}
                                </button>
                              </td>
                              <td className="px-2 py-2.5">
                                {isCustom && (
                                  <button
                                    type="button"
                                    onClick={() => deleteCustomMilestone(m.id)}
                                    title="이 마일스톤 삭제"
                                    className="text-slate-300 hover:text-rose-500 transition-colors"
                                  >
                                    ✕
                                  </button>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                    <button
                      type="button"
                      onClick={() => addCustomMilestone(p.id, pMilestones)}
                      className="w-full text-left px-4 py-2 text-xs text-indigo-500 hover:bg-indigo-50 hover:text-indigo-600 transition-colors border-t border-slate-100"
                    >
                      + 직접입력
                    </button>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      </div>

      {hovered && (
        <div
          className="fixed bg-slate-900 text-white text-xs rounded-lg px-3 py-2 shadow-lg max-w-xs z-40 pointer-events-none"
          style={{ left: hovered.x + 14, top: hovered.y + 14 }}
        >
          <div className="font-medium">{hovered.m.name}</div>
          <div className="text-slate-300 mt-0.5">
            {fmt(hovered.s)} ~ {fmt(hovered.e)}
          </div>
          {hovered.m.manager && <div className="text-slate-300">담당: {hovered.m.manager}</div>}
          {hovered.m.memo && <div className="text-slate-400 mt-1">{hovered.m.memo}</div>}
        </div>
      )}

      {rangePopover && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/20"
          onClick={() => setRangePopover(null)}
        >
          <div
            className="bg-white rounded-lg shadow-xl p-4 w-[320px]"
            onClick={(evt) => evt.stopPropagation()}
          >
            <div className="text-sm font-medium text-slate-800">
              {fmt(parseDate(rangePopover.startDate))} ~ {fmt(parseDate(rangePopover.endDate))}
            </div>
            <div className="text-xs text-slate-500 mt-0.5 mb-3">
              이 구간을 적용할 마일스톤을 선택하세요
            </div>
            <div className="space-y-1 max-h-64 overflow-y-auto">
              {(milestonesByProject[rangePopover.projectId] || []).map((m) => {
                const hasDate = m.actual_start_date || m.planned_start_date;
                return (
                  <button
                    key={m.id}
                    className="w-full text-left px-2.5 py-2 rounded-md hover:bg-indigo-50 text-sm text-slate-700 flex items-center justify-between"
                    onClick={() => {
                      updateMilestoneField(m.id, "actual_start_date", rangePopover.startDate);
                      updateMilestoneField(m.id, "actual_end_date", rangePopover.endDate);
                      saveMilestoneField(m.id, "actual_start_date", rangePopover.startDate);
                      saveMilestoneField(m.id, "actual_end_date", rangePopover.endDate);
                      setRangePopover(null);
                    }}
                  >
                    <span>{m.name}</span>
                    {hasDate && (
                      <span className="text-[10px] text-amber-600 flex-shrink-0 ml-2">
                        기존 일정 덮어씀
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
            <button
              className="mt-3 text-xs text-slate-400 hover:text-slate-600"
              onClick={() => setRangePopover(null)}
            >
              취소
            </button>
          </div>
        </div>
      )}
      {/* 하단 고정 가로 스크롤바 — 트랙패드 없이 마우스만 쓰는 사람도 좌우 스크롤 가능하게(위 타임라인과 위치 동기화) */}
      {totalWidth > 0 && (
        <div
          className="fixed bottom-0 left-0 right-0 h-4 overflow-x-auto overflow-y-hidden bg-white border-t border-slate-200 z-40"
          ref={bottomScrollRef}
          onScroll={handleBottomScrollbarScroll}
        >
          <div style={{ width: totalWidth + 280, height: 1 }} />
        </div>
      )}
    </div>
  );
}
