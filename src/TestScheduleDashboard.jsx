import { useState, useEffect, useMemo } from "react";

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
  const [zoom, setZoom] = useState("month");
  const [expanded, setExpanded] = useState(new Set());
  const [hovered, setHovered] = useState(null);

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

  // 실제 일정(마일스톤 시작일) 기준 정렬 — 일정이 아예 없는 프로젝트는 맨 아래로
  const sortedProjects = useMemo(() => {
    return [...projects].sort((a, b) => {
      const getEarliest = (id) => {
        const ms = milestonesByProject[id] || [];
        const dates = ms
          .map((m) => parseDate(m.actual_start_date) || parseDate(m.planned_start_date))
          .filter(Boolean);
        return dates.length ? Math.min(...dates.map((d) => d.getTime())) : Infinity;
      };
      return getEarliest(a.id) - getEarliest(b.id);
    });
  }, [projects, milestonesByProject]);

  function xFor(date) {
    return daysBetween(rangeStart, date) * pxPerDay;
  }

  function toggleExpand(id) {
    setExpanded((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  const todayX = xFor(today);

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
        </div>
      </div>

      <div className="overflow-x-auto">
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

            return (
              <div key={p.id} className="border-b border-slate-100">
                <div
                  className="flex hover:bg-slate-50/60 cursor-pointer"
                  onClick={() => toggleExpand(p.id)}
                >
                  <div className="w-[280px] flex-shrink-0 border-r border-slate-200 px-4 py-3 flex items-start gap-2 sticky left-0 z-10 bg-white">
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
                  </div>
                  <div className="relative flex-1" style={{ width: totalWidth, height: 62 }}>
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
                    {dated.map(({ m, s, e }) => {
                      const status = effectiveStatus(m, today);
                      const style = STATUS_STYLE[status];
                      const left = xFor(s);
                      const width = Math.max(xFor(e) - xFor(s), 4);
                      const shortDate = (d) =>
                        `${String(d.getMonth() + 1).padStart(2, "0")}.${String(
                          d.getDate()
                        ).padStart(2, "0")}`;
                      return (
                        <div key={m.id} className="absolute" style={{ left, top: 10, width: Math.max(width, 46) }}>
                          <div
                            className="rounded-md border text-[11px] px-1.5 flex items-center overflow-hidden whitespace-nowrap"
                            style={{
                              width,
                              height: 24,
                              background: style.bg,
                              borderColor: style.border,
                              color: style.text,
                            }}
                            onMouseEnter={() => setHovered({ m, s, e, status })}
                            onMouseLeave={() => setHovered(null)}
                          >
                            {width > 60 ? m.name : ""}
                          </div>
                          <div className="text-[10px] text-slate-400 mt-0.5 whitespace-nowrap">
                            {shortDate(s)}~{shortDate(e)}
                          </div>
                        </div>
                      );
                    })}
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
                        </tr>
                      </thead>
                      <tbody>
                        {pMilestones.map((m) => {
                          const status = effectiveStatus(m, today);
                          const style = STATUS_STYLE[status];
                          const mProgress = milestoneProgress(m, today);
                          return (
                            <tr key={m.id} className="border-b border-slate-100 last:border-0">
                              <td className="px-4 py-2.5 text-slate-700 sticky left-0 z-10 bg-slate-50 border-r border-slate-200">
                                {m.name}
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
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
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
    </div>
  );
}
