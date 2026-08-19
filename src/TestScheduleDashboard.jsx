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

// 기간이 겹치는 마일스톤들을 클러스터로 묶음(스윕: 시작일순 정렬 후 이어붙이기)
// 클러스터 안 items가 2개 이상이면 실제로 겹치는 그룹, 1개면 단독 일정
function buildOverlapClusters(dated) {
  const sorted = [...dated].sort((a, b) => a.s - b.s);
  const clusters = [];
  let current = null;
  for (const item of sorted) {
    if (!current) {
      current = { items: [item], maxEnd: item.e };
    } else if (item.s <= current.maxEnd) {
      current.items.push(item);
      if (item.e > current.maxEnd) current.maxEnd = item.e;
    } else {
      clusters.push(current);
      current = { items: [item], maxEnd: item.e };
    }
  }
  if (current) clusters.push(current);
  return clusters;
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

export default function ScheduleDashboard() {
  const [projects, setProjects] = useState([]);
  const [milestones, setMilestones] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [zoom, setZoom] = useState("day");
  const [expanded, setExpanded] = useState(new Set());
  const [subtrackExpanded, setSubtrackExpanded] = useState(new Set()); // 겹치는 마일스톤 카테고리별 서브트랙 펼침
  const [hovered, setHovered] = useState(null);
  const [dragState, setDragState] = useState(null); // 마일스톤 바 드래그(날짜 변경)
  const [draggedProjectId, setDraggedProjectId] = useState(null); // 프로젝트 행 순서 드래그
  const [rangeSelect, setRangeSelect] = useState(null); // 빈 타임라인 영역 드래그 선택(신규 일정 입력)
  const [rangePopover, setRangePopover] = useState(null); // 드래그 선택 완료 후 마일스톤 선택 팝업
  const rangeMovedRef = useRef(false); // 방금 실제로 드래그했는지(단순 클릭과 구분)
  const scrollRef = useRef(null);

  useEffect(() => {
    fetchAll()
      .then(({ projects, milestones }) => {
        setProjects(projects);
        setMilestones(milestones);
      })
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, []);

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

  const today = useMemo(() => new Date(), []);

  const milestonesByProject = useMemo(() => {
    const map = {};
    for (const m of milestones) {
      if (!map[m.project_id]) map[m.project_id] = [];
      map[m.project_id].push(m);
    }
    return map;
  }, [milestones]);

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

  function managerOf(pMilestones) {
    const withManager = pMilestones.find((m) => m.manager);
    return withManager ? withManager.manager : null;
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

  // 겹치는 마일스톤 클러스터의 카테고리별 서브트랙 펼침/접힘
  function toggleSubtrack(id) {
    setSubtrackExpanded((prev) => {
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
      <div key={m.id} className="absolute" style={{ left, top, width: Math.max(width, 46) }}>
        <div
          className={`rounded-md border text-[11px] px-1.5 flex items-center overflow-hidden whitespace-nowrap select-none ${
            isDragging ? "cursor-grabbing shadow-md" : "cursor-grab"
          }`}
          style={{
            width,
            height: 24,
            background: style.bg,
            borderColor: style.border,
            color: style.text,
          }}
          onMouseEnter={() => !dragState && setHovered({ m, s, e, status })}
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
          {width > 60 ? m.name : ""}
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
    }
  }, [zoom, loading, todayX]);


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
              프로젝트 {projects.length}개 · 오늘 {fmt(today)} · 테스트 배포
            </p>
          </div>
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
          <div className="flex items-center gap-1.5">
            <span className="inline-flex items-center justify-center w-3.5 h-3.5 rounded-full bg-indigo-500 text-white text-[8px] font-bold">
              n
            </span>
            겹치는 일정(클릭시 카테고리별로 펼치기)
          </div>
        </div>
      </div>

      <div className="overflow-x-auto" ref={scrollRef}>
        <div style={{ minWidth: totalWidth + 280 }}>
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

            const progress = projectProgress(pMilestones, today);

            // 겹치는 마일스톤 클러스터 계산 — 클러스터가 2건 이상 겹치면 기본 트랙에선 점 마커로만 표시
            const isSubtrackOpen = subtrackExpanded.has(p.id);
            const clusters = buildOverlapClusters(dated);
            const overlapClusters = clusters.filter((c) => c.items.length > 1);
            const clusteredIds = new Set(
              overlapClusters.flatMap((c) => c.items.map((x) => x.m.id))
            );
            const soloItems = dated.filter((x) => !clusteredIds.has(x.m.id));
            // 서브트랙을 펼쳤을 때 카테고리별 레인 순서(처음 등장 순, 미분류는 맨 뒤)
            const categoryLanes = [];
            if (isSubtrackOpen) {
              for (const c of overlapClusters) {
                for (const item of c.items) {
                  const cat = item.m.category || "미분류";
                  if (!categoryLanes.includes(cat)) categoryLanes.push(cat);
                }
              }
              categoryLanes.sort((a, b) =>
                a === "미분류" ? 1 : b === "미분류" ? -1 : a.localeCompare(b, "ko")
              );
            }
            const laneHeight = 34;
            const rowHeight = 62 + categoryLanes.length * laneHeight;

            return (
              <div key={p.id} className="border-b border-slate-100">
                <div
                  className={`flex hover:bg-slate-50/60 cursor-pointer ${
                    draggedProjectId && draggedProjectId !== p.id
                      ? "border-t-2 border-t-transparent hover:border-t-indigo-400"
                      : ""
                  }`}
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
                      <div className="text-sm font-medium text-slate-800 truncate">
                        {p.name}
                        {managerOf(pMilestones) && (
                          <span className="text-indigo-500 font-normal ml-1.5">
                            · {managerOf(pMilestones)}
                          </span>
                        )}
                      </div>
                      <div className="text-[11px] text-slate-400 mt-0.5 flex items-center gap-1">
                        오픈예정 {fmt(targetOpen)}
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
                    {isSubtrackOpen &&
                      categoryLanes.map((cat, i) => (
                        <div
                          key={cat}
                          className="absolute left-4 text-[10px] font-medium text-slate-400 truncate max-w-[220px]"
                          style={{ top: 10 + (i + 1) * laneHeight + 4 }}
                        >
                          {cat}
                        </div>
                      ))}
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
                    {soloItems.map(({ m, s, e }) => renderMilestoneBar(m, s, e, 10))}
                    {isSubtrackOpen
                      ? overlapClusters.flatMap((c) =>
                          c.items.map(({ m, s, e }) => {
                            const laneIdx = categoryLanes.indexOf(m.category || "미분류");
                            return renderMilestoneBar(m, s, e, 10 + (laneIdx + 1) * laneHeight);
                          })
                        )
                      : overlapClusters.map((c) => {
                          const clusterStart = c.items.reduce(
                            (min, x) => (x.s < min ? x.s : min),
                            c.items[0].s
                          );
                          const names = c.items.map((x) => x.m.name).join(", ");
                          return (
                            <div
                              key={`cluster-${clusterStart.getTime()}`}
                              className="absolute flex flex-col items-center cursor-pointer"
                              style={{ left: xFor(clusterStart) - 3, top: 14 }}
                              onClick={(evt) => {
                                evt.stopPropagation();
                                toggleSubtrack(p.id);
                              }}
                              title={`겹치는 일정 ${c.items.length}건: ${names} — 클릭해서 카테고리별로 펼치기`}
                            >
                              <span className="w-4 h-4 rounded-full bg-indigo-500 border-2 border-white shadow flex items-center justify-center text-[8px] text-white font-bold leading-none">
                                {c.items.length}
                              </span>
                            </div>
                          );
                        })}
                    {isSubtrackOpen && overlapClusters.length > 0 && (
                      <button
                        type="button"
                        className="absolute right-2 top-1 text-[10px] text-indigo-400 hover:text-indigo-600 bg-white/90 px-1.5 py-0.5 rounded border border-indigo-100"
                        onClick={(evt) => {
                          evt.stopPropagation();
                          toggleSubtrack(p.id);
                        }}
                      >
                        ▲ 서브트랙 접기
                      </button>
                    )}
                    {dated.length === 0 && (
                      <div className="absolute left-3 top-4 text-[11px] text-slate-300">
                        일정 미입력
                      </div>
                    )}
                  </div>
                </div>

                {isOpen && (
                  <div className="bg-slate-50/70 border-t border-slate-100">
                    <table className="w-full text-xs">
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
        <div className="fixed bottom-6 right-6 bg-slate-900 text-white text-xs rounded-lg px-3 py-2 shadow-lg max-w-xs z-40">
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
    </div>
  );
}
