import { useState, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { Plus, X, Building2, LayoutGrid, Table2, Trash2, Upload, Save, Users, Loader2, LayoutDashboard } from "lucide-react";
import {
  Tooltip, Legend, ResponsiveContainer,
  BarChart, Bar, XAxis, YAxis, CartesianGrid,
} from "recharts";

import { getProjectIndex, getProjectData, saveProjectData, saveProjectIndex, getCatalog, saveCatalog } from "./api";
import { saveRoomTypes, loadRoomTypes } from "./roomTypesApi";
import { saveOrderItems, loadOrderItems } from "./orderItemsApi";
import TestScheduleDashboard from "./TestScheduleDashboard";

const TIERS = ["Flagship", "Premium", "Upper Select", "Select", "Essential"];
const BED_TYPES = ["싱글", "퀸"];
const BATHTUB = ["유", "무"];
// 브랜드 룸 네이밍 가이드라인 v1.0 — Core Room Grade 고정 6단계
const ROOM_GRADES = ["Standard", "Superior", "Deluxe", "Premier", "Suite", "Signature Suite"];

const CATEGORY_COLORS = [
  { bg: "bg-indigo-100", text: "text-indigo-800", bar: "bg-indigo-400", border: "border-indigo-300" },
  { bg: "bg-teal-100", text: "text-teal-800", bar: "bg-teal-400", border: "border-teal-300" },
  { bg: "bg-rose-100", text: "text-rose-800", bar: "bg-rose-400", border: "border-rose-300" },
  { bg: "bg-amber-100", text: "text-amber-800", bar: "bg-amber-400", border: "border-amber-300" },
  { bg: "bg-sky-100", text: "text-sky-800", bar: "bg-sky-400", border: "border-sky-300" },
  { bg: "bg-violet-100", text: "text-violet-800", bar: "bg-violet-400", border: "border-violet-300" },
];
const CHART_COLORS = ["#818cf8", "#2dd4bf", "#fb7185", "#fbbf24", "#38bdf8", "#a78bfa"];

// 공정표 — 공정별 레인 (발주관리 탭 대분류와 맞춤)
const SCHED_CATS = ["기획", "가구", "가전", "린넨", "현장"];
const SCHED_CAT_DOT = { 기획: "#A79CEE", 가구: "#7F77DD", 가전: "#5DCAA5", 린넨: "#E0A857", 현장: "#0F6E56" };
const SCHED_STATUSES = ["시작전", "진행중", "확인필요", "완료"];
const DEFAULT_SCHEDULE_ITEMS = [
  { name: "룸타입 확정", cat: "기획" },
  { name: "예산안 확정", cat: "기획" },
  { name: "FF&E 품목 확정", cat: "가구" },
  { name: "OS&E 품목 확정", cat: "린넨" },
  { name: "가구 발주", cat: "가구" },
  { name: "가전 발주", cat: "가전" },
  { name: "린넨 발주", cat: "린넨" },
  { name: "가구 입고/설치", cat: "가구" },
  { name: "가전 입고/설치", cat: "가전" },
  { name: "스타일링", cat: "현장" },
  { name: "촬영", cat: "현장" },
  { name: "운영사 인계", cat: "현장" },
];
function makeDefaultScheduleItems() {
  return DEFAULT_SCHEDULE_ITEMS.map((it) => ({
    id: nextId(),
    name: it.name,
    cat: it.cat,
    status: "시작전",
    delay: false,
    dueDate: null,
    workStart: null,
    workEnd: null,
    assignee: "",
    memo: "",
  }));
}

function toDate(s) {
  return s ? new Date(s + "T00:00:00") : null;
}
function daysBetween(a, b) {
  return (b - a) / (1000 * 60 * 60 * 24);
}

const DEFAULT_BASIC_PRESET = [
  "헤드보드", "침대바디", "협탁", "옷장", "붙박이 책장", "콘센트", "슬리퍼걸이",
  "구둣주걱", "리모컨거치대", "입식 테이블", "입식 의자", "객실 슬리퍼", "욕실 슬리퍼",
  "객실 휴지통", "욕실 휴지통", "객실 휴지케이스", "옷걸이", "바지걸이", "TV",
  "소형 무소음 냉장고", "거울", "조명",
];

let idCounter = 1;
const nextId = () => idCounter++;

function CatalogPickerPanel({ items, selected, onToggle, onConfirm, onCancel, categoryFilter, setCategoryFilter }) {
  if (items.length === 0) {
    return (
      <div className="mb-4 bg-slate-50 border border-slate-200 rounded-lg p-3 text-xs text-slate-400">
        카탈로그가 비어있어요. 위쪽 "품목 마스터 카탈로그 업로드"에서 엑셀을 먼저 올려주세요.
        <div className="mt-2 text-right">
          <button onClick={onCancel} className="text-xs border border-slate-300 rounded-md px-3 py-1 hover:bg-white">
            닫기
          </button>
        </div>
      </div>
    );
  }
  const categories = ["전체", ...Array.from(new Set(items.map((it) => it.category)))];
  const filtered = categoryFilter === "전체" ? items : items.filter((it) => it.category === categoryFilter);
  return (
    <div className="mb-4 bg-slate-50 border border-slate-200 rounded-lg p-3">
      <div className="flex items-center justify-between mb-2">
        <select
          value={categoryFilter}
          onChange={(e) => setCategoryFilter(e.target.value)}
          className="text-xs border border-slate-300 rounded-lg px-2 py-1"
        >
          {categories.map((c) => (
            <option key={c} value={c}>
              {c}
            </option>
          ))}
        </select>
        <span className="text-[11px] text-slate-500">{selected.size}개 선택됨</span>
      </div>
      <div className="max-h-48 overflow-y-auto grid grid-cols-2 gap-1 mb-2">
        {filtered.map((it) => (
          <label
            key={it.id}
            className="flex items-center gap-1.5 text-xs px-2 py-1 rounded hover:bg-white cursor-pointer"
          >
            <input type="checkbox" checked={selected.has(it.id)} onChange={() => onToggle(it.id)} />
            <span className="truncate">{it.name}</span>
            <span className="text-slate-400 ml-auto whitespace-nowrap">
              {it.unitPrice.toLocaleString("ko-KR")}원
            </span>
          </label>
        ))}
      </div>
      <div className="flex justify-end gap-2">
        <button onClick={onCancel} className="text-xs border border-slate-300 rounded-md px-3 py-1 hover:bg-white">
          취소
        </button>
        <button
          onClick={onConfirm}
          className="text-xs bg-amber-700 text-white px-3 py-1 rounded-md hover:bg-amber-800"
        >
          선택 항목 추가
        </button>
      </div>
    </div>
  );
}

function ExpenseSection({ title, items, onAdd, onUpdate, onRemove }) {
  const budgetTotal = items.reduce((s, it) => s + (it.budgetAmount || 0), 0);
  const actualTotal = items.reduce((s, it) => s + (it.actualAmount || 0), 0);
  return (
    <div className="bg-white border border-slate-200 rounded-xl p-6">
      <div className="flex items-center justify-between mb-4">
        <span className="text-sm font-medium tracking-wide text-slate-500">{title}</span>
        <button
          onClick={onAdd}
          className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50 flex items-center gap-1"
        >
          <Plus size={13} /> 항목 추가
        </button>
      </div>
      {items.length > 0 ? (
        <>
          <div className="max-h-80 overflow-y-auto">
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-white z-10">
                <tr className="text-left text-xs text-slate-500 border-b border-slate-200">
                  <th className="py-1.5 font-normal">항목명</th>
                  <th className="py-1.5 font-normal text-right">예산금액</th>
                  <th className="py-1.5 font-normal text-right">집행금액</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {items.map((it) => (
                  <tr key={it.id} className="border-b border-slate-100">
                    <td className="py-1.5">
                      <input
                        value={it.name}
                        onChange={(e) => onUpdate(it.id, "name", e.target.value)}
                        placeholder="예: 폐기물 처리비"
                        className="w-full border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        value={it.budgetAmount}
                        onChange={(e) => onUpdate(it.id, "budgetAmount", e.target.value)}
                        className="w-32 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        value={it.actualAmount}
                        onChange={(e) => onUpdate(it.id, "actualAmount", e.target.value)}
                        className="w-32 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 bg-teal-50/40"
                      />
                    </td>
                    <td className="py-1.5 pl-2">
                      <button onClick={() => onRemove(it.id)} className="text-slate-400 hover:text-rose-600">
                        <Trash2 size={14} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div className="text-right text-xs text-slate-500 mt-2">
            소계 — 예산 <span className="font-semibold text-slate-800">{budgetTotal.toLocaleString("ko-KR")}원</span>
            {" / "}집행 <span className="font-semibold text-teal-700">{actualTotal.toLocaleString("ko-KR")}원</span>
          </div>
        </>
      ) : (
        <p className="text-xs text-slate-400">항목을 추가해 예산/집행 금액을 입력하세요.</p>
      )}
    </div>
  );
}

function fmtLocalDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}
const SCHED_LEFT_COL = 220;
const SCHED_PX_PER_DAY = { month: 3.6, week: 11, day: 30 };

function schedPx(d, rangeStart, pxPerDay) {
  return daysBetween(rangeStart, d) * pxPerDay;
}
function schedItemRange(it, today) {
  if (!it.workStart) return null;
  const start = toDate(it.workStart);
  const end = it.workEnd ? toDate(it.workEnd) : today;
  return { start, end, inprog: !it.workEnd };
}
function schedFindOverlaps(items, today) {
  const ranges = items.map((it) => ({ it, r: schedItemRange(it, today) })).filter((x) => x.r);
  const points = [];
  for (let i = 0; i < ranges.length; i++) {
    for (let j = i + 1; j < ranges.length; j++) {
      const a = ranges[i], b = ranges[j];
      if (a.it.cat === b.it.cat) continue;
      const s = a.r.start > b.r.start ? a.r.start : b.r.start;
      const e = a.r.end < b.r.end ? a.r.end : b.r.end;
      if (s <= e) points.push(s);
    }
  }
  return points;
}
function schedProjectProgress(items) {
  const total = items.length, done = items.filter((i) => i.status === "완료").length;
  return total ? (done / total) * 100 : 0;
}
function schedExpectedProgress(start, target, today) {
  const s = toDate(start), t = toDate(target);
  if (!s || !t || t <= s) return 100;
  const el = Math.max(0, Math.min(t - s, today - s));
  return (el / (t - s)) * 100;
}
function schedIsDelayed(items, start, target, today) {
  return schedExpectedProgress(start, target, today) - schedProjectProgress(items) >= 20 || items.some((i) => i.delay);
}
function schedTeamWorkDone(items) {
  const shoot = items.find((it) => it.name === "촬영");
  const handover = items.find((it) => it.name === "운영사 인계");
  return !!(shoot && shoot.status === "완료") || !!(handover && handover.status === "완료");
}
function schedDdayLabel(target, today) {
  const t = toDate(target);
  if (!t) return "";
  const d = Math.round(daysBetween(today, t));
  return d >= 0 ? `D-${d}` : `D+${-d}`;
}
function schedGetFullRange(projects, today) {
  const years = [today.getFullYear()];
  projects.forEach((p) => {
    const s = toDate(p.start), t = toDate(p.target);
    if (s) years.push(s.getFullYear());
    if (t) years.push(t.getFullYear());
  });
  const minY = Math.min(...years) - 1;
  const maxY = Math.max(...years) + 1;
  return [new Date(minY, 0, 1), new Date(maxY + 1, 0, 1)];
}
function schedTooltipLines(it, today) {
  const r = schedItemRange(it, today);
  const lines = [`${it.name} (${it.cat})`, `${it.workStart || "?"} ~ ${it.workEnd || (r && r.inprog ? "진행중" : "?")}`, `상태: ${it.status}`];
  if (it.dueDate) lines.push(`계획 종료: ${it.dueDate}`);
  if (it.assignee) lines.push(`담당자: ${it.assignee}`);
  const isDelayedItem = it.delay || it.status === "확인필요";
  if (isDelayedItem) lines.push(`🚩 ${it.memo || "지연"}`);
  else if (it.memo) lines.push(`메모: ${it.memo}`);
  return lines;
}
function schedSegClass(status) {
  if (status === "확인필요") return "sp-segst-check";
  if (status === "완료") return "sp-segst-done";
  return "sp-segst-progress";
}

const SP_STYLE = `
.sp-root{--sp-bg:#F7F7F5;--sp-surface:#FFFFFF;--sp-surface-2:#FBFBFA;--sp-border:#E3E1DC;--sp-border-strong:#CFCCC5;--sp-text:#1F1E1C;--sp-text-secondary:#5B5952;--sp-text-muted:#8B8880;--sp-accent:#0F6E56;--sp-accent-bg:#E7F3EF;--sp-accent-text:#0F6E56;--sp-danger:#E5484D;--sp-danger-bg:#FBEAEA;--sp-progress:#3B82F6;--sp-check:#E5484D;--sp-done:#12805C;font-size:14px;color:var(--sp-text)}
.sp-root *{box-sizing:border-box}
.sp-toolbar{display:flex;align-items:center;gap:10px;margin-bottom:10px;flex-wrap:wrap}
.sp-zoombtns{display:flex;border:1px solid var(--sp-border);border-radius:8px;overflow:hidden}
.sp-zoombtns button{padding:6px 14px;font-size:12.5px;border:none;background:transparent;color:var(--sp-text-secondary);cursor:pointer;font-family:inherit;border-left:1px solid var(--sp-border)}
.sp-zoombtns button:first-child{border-left:none}
.sp-zoombtns button.active{background:var(--sp-accent-bg);color:var(--sp-accent-text);font-weight:600}
.sp-todaybtn,.sp-chk,.sp-modebtn{background:var(--sp-surface);border:1px solid var(--sp-border);color:var(--sp-text-secondary);padding:6px 12px;border-radius:8px;font-size:12px;cursor:pointer;font-family:inherit}
.sp-modebtn{color:var(--sp-text);font-weight:600;border-color:var(--sp-border-strong)}
.sp-chk.active{background:var(--sp-text);color:#fff;border-color:var(--sp-text)}
.sp-addbtn{margin-left:auto;background:var(--sp-text);color:#fff;border:none;padding:7px 14px;border-radius:8px;font-size:12.5px;cursor:pointer;font-family:inherit}
.sp-addbtn:hover{opacity:.85}
.sp-filterbar{display:flex;align-items:center;gap:8px;margin-bottom:14px;flex-wrap:wrap;font-size:11.5px;color:var(--sp-text-secondary)}
.sp-catchip{padding:4px 10px;border-radius:99px;border:1px solid var(--sp-border);cursor:pointer;background:#fff;font-size:11.5px;display:inline-flex;align-items:center;gap:5px}
.sp-catchip.sp-off{opacity:.35}
.sp-dot{width:8px;height:8px;border-radius:2px;display:inline-block}
.sp-legend{display:flex;align-items:center;gap:12px;font-size:11.5px;color:var(--sp-text-secondary);flex-wrap:wrap;margin-bottom:12px}
.sp-legend span{display:flex;align-items:center;gap:4px}
.sp-hatch{background:var(--sp-danger);background-image:repeating-linear-gradient(45deg,rgba(255,255,255,.4) 0 3px,transparent 3px 6px)}
.sp-diamond-legend{display:inline-block;width:9px;height:9px;background:#B8862E;transform:rotate(45deg)}
.sp-todaydot{width:2px;height:10px;background:var(--sp-danger);display:inline-block}
.sp-board{border:1px solid var(--sp-border);border-radius:12px;overflow:hidden;background:var(--sp-surface-2);margin-bottom:16px}
.sp-boardscroll{overflow-x:auto;overflow-y:hidden}
.sp-headerrow{display:grid}
.sp-headerrow>div:first-child{border-right:1px solid var(--sp-border);background:var(--sp-surface);position:sticky;left:0;z-index:5}
.sp-tlheader{position:relative;background:var(--sp-surface)}
.sp-monthrow{display:flex;height:20px;border-bottom:1px solid var(--sp-border)}
.sp-monthcell{flex-shrink:0;border-right:1px solid var(--sp-border-strong);font-size:11px;font-weight:600;color:var(--sp-text-secondary);padding:3px 0 0 6px;box-sizing:border-box}
.sp-subrow{display:flex;height:16px;border-bottom:1px solid var(--sp-border)}
.sp-subcell{flex-shrink:0;border-right:1px solid var(--sp-border);font-size:9px;color:var(--sp-text-muted);text-align:center;box-sizing:border-box}
.sp-row{display:grid;border-top:1px solid var(--sp-border)}
.sp-rowmeta{padding:12px;border-right:1px solid var(--sp-border);background:var(--sp-surface);position:sticky;left:0;z-index:5;display:flex;flex-direction:column;justify-content:center}
.sp-rowmeta-top{display:flex;align-items:center;gap:6px}
.sp-collapsebtn{background:none;border:none;cursor:pointer;color:var(--sp-text-muted);font-size:10px;padding:0}
.sp-name{font-size:13.5px;font-weight:700;display:flex;align-items:center;gap:6px}
.sp-nameinput{font-size:13.5px;font-weight:700;border:1px solid var(--sp-border-strong);border-radius:5px;padding:1px 5px;width:100%;font-family:inherit;background:#fff}
.sp-pencil{cursor:pointer;color:var(--sp-text-muted);font-size:12px}
.sp-pctrow{display:flex;align-items:baseline;gap:6px;margin-top:4px}
.sp-pct{font-size:19px;font-weight:800}
.sp-meta{font-size:10.5px;color:var(--sp-text-muted);margin-top:2px;display:flex;align-items:center;gap:5px;flex-wrap:wrap}
.sp-badge{font-size:10px;padding:1px 7px;border-radius:99px;font-weight:700}
.sp-badge-delay{background:var(--sp-danger-bg);color:var(--sp-danger)}
.sp-badge-ontrack{background:var(--sp-accent-bg);color:var(--sp-accent-text)}
.sp-tlwrap{position:relative;padding:8px 0}
.sp-lane{position:relative;height:28px;margin-bottom:4px}
.sp-lane:last-child{margin-bottom:0}
.sp-lanelabel{position:absolute;left:-2px;top:0;font-size:8.5px;color:var(--sp-text-muted);transform:translateY(-11px)}
.sp-seg{position:absolute;top:0;height:24px;border-radius:5px;color:#fff;font-size:10.5px;font-weight:600;display:flex;align-items:center;padding:0 6px;overflow:hidden;white-space:nowrap;cursor:default;border:1.5px solid transparent}
.sp-icotxt{display:flex;align-items:center;gap:3px;overflow:hidden;text-overflow:ellipsis}
.sp-segst-progress{background:var(--sp-progress)}
.sp-segst-check{background:var(--sp-check)}
.sp-segst-done{background:var(--sp-done)}
.sp-seg-delayed{border-color:#7A1518;box-shadow:0 0 0 1px #7A1518}
.sp-seg-pattern{background-image:repeating-linear-gradient(45deg,rgba(255,255,255,.28) 0 5px,transparent 5px 10px)}
.sp-seg-planned{opacity:.55;border:1.5px dashed rgba(255,255,255,.6)}
.sp-seg-overrun{background:var(--sp-danger);background-image:repeating-linear-gradient(45deg,rgba(255,255,255,.35) 0 5px,transparent 5px 10px);border:1.5px solid #7A1518}
.sp-seg-collapsed{position:absolute;height:14px;border-radius:5px;background:var(--sp-border)}
.sp-seg-collapsed-fill{position:absolute;left:0;top:0;bottom:0;border-radius:5px}
.sp-overlapband{position:absolute;top:-4px;bottom:-4px;width:6px;background:rgba(229,72,77,.18);border-left:1.5px dashed var(--sp-danger);border-right:1.5px dashed var(--sp-danger);z-index:2}
.sp-overlapicon{position:absolute;top:-16px;font-size:11px;transform:translateX(-50%)}
.sp-todayline{position:absolute;top:0;bottom:0;width:1.5px;background:var(--sp-danger);z-index:4}
.sp-milestone{position:absolute;top:-4px;bottom:-4px;display:flex;flex-direction:column;align-items:center;transform:translateX(-50%);z-index:4}
.sp-diamond{width:11px;height:11px;background:#B8862E;transform:rotate(45deg);border:1.5px solid #fff;box-shadow:0 0 0 1px #B8862E}
.sp-mlabel{font-size:8.5px;color:#8A6416;font-weight:700;margin-top:2px;white-space:nowrap}
.sp-addrow{padding:10px 12px;color:var(--sp-text-muted);font-size:12.5px;cursor:pointer;display:flex;align-items:center;gap:6px;border-top:1px solid var(--sp-border);background:var(--sp-surface);position:sticky;left:0;width:220px}
.sp-addrow:hover{color:var(--sp-text)}
.sp-detail{background:var(--sp-surface);border-top:1px solid var(--sp-border);position:sticky;left:0;overflow-x:auto;box-shadow:2px 0 10px rgba(0,0,0,.05);grid-column:1/-1}
.sp-progresswrap{padding:12px 16px 0}
.sp-pbar{height:5px;background:var(--sp-border);border-radius:99px;overflow:hidden}
.sp-pfill{height:100%;background:var(--sp-accent)}
.sp-ptext{font-size:11.5px;color:var(--sp-text-secondary);margin-bottom:6px}
.sp-tablewrap{overflow-x:auto}
table.sp-mstable{width:100%;border-collapse:collapse;font-size:12px;margin-top:10px}
table.sp-mstable th{background:#123B32;color:#fff;font-weight:600;text-align:left;padding:8px 10px;font-size:11px;white-space:nowrap}
table.sp-mstable td{padding:6px 10px;border-bottom:1px solid var(--sp-border);vertical-align:middle}
table.sp-mstable tr:nth-child(even) td{background:#F4F3F0}
.sp-itemname{font-weight:500;white-space:nowrap}
.sp-catsel{font-size:11px;padding:3px 6px;border-radius:6px;border:1px solid var(--sp-border);font-family:inherit;background:#fff}
.sp-pill{font-size:10.5px;padding:2px 9px;border-radius:99px;font-weight:600;border:none;cursor:pointer;font-family:inherit}
.sp-pillst-시작전{background:#EAE8E3;color:#7A776E}
.sp-pillst-진행중{background:#DCEAFB;color:#1D5FCC}
.sp-pillst-확인필요{background:var(--sp-danger-bg);color:var(--sp-danger)}
.sp-pillst-완료{background:#E3F5E9;color:#15803D}
.sp-datein{font-size:11px;padding:3px 5px;border:1px solid var(--sp-border);border-radius:5px;font-family:inherit;width:120px}
.sp-w70{width:70px}
.sp-w130{width:130px}
.sp-addmsrow{display:flex;gap:8px;padding:10px 16px 14px;flex-wrap:wrap;align-items:center}
.sp-addmsrow input,.sp-addmsrow select{padding:5px 8px;border:1px solid var(--sp-border);border-radius:6px;font-size:11.5px;font-family:inherit}
.sp-addmsrow button{padding:6px 12px;border-radius:6px;border:none;background:var(--sp-accent);color:#fff;font-size:11.5px;cursor:pointer;font-family:inherit}
.sp-formrow{display:flex;gap:10px;flex-wrap:wrap;align-items:flex-end;padding:14px 16px;background:var(--sp-surface);border:1px solid var(--sp-border);border-radius:10px;margin-bottom:12px}
.sp-formrow label{display:flex;flex-direction:column;gap:4px;font-size:11px;color:var(--sp-text-secondary)}
.sp-formrow input,.sp-formrow select{padding:6px 8px;border:1px solid var(--sp-border);border-radius:6px;font-size:12.5px;font-family:inherit}
.sp-formrow button{padding:7px 14px;border-radius:6px;border:none;background:var(--sp-text);color:#fff;font-size:12.5px;cursor:pointer;font-family:inherit}
.sp-empty{padding:40px;text-align:center;color:var(--sp-text-muted);font-size:13px}
.sp-delcell{text-align:center}
.sp-rmbtn{background:none;border:none;color:var(--sp-text-muted);cursor:pointer;font-size:12px}
.sp-rmbtn:hover{color:var(--sp-danger)}
.sp-tooltip{position:fixed;background:#1F1E1C;color:#fff;font-size:11.5px;padding:6px 9px;border-radius:6px;pointer-events:none;z-index:50;line-height:1.5;max-width:230px;box-shadow:0 4px 14px rgba(0,0,0,.2)}
.sp-editentry{margin-top:6px;font-size:10.5px;padding:3px 9px;border-radius:6px;border:1px solid var(--sp-border-strong);background:#fff;color:var(--sp-text-secondary);cursor:pointer;font-family:inherit;align-self:flex-start}
.sp-editentry:hover{background:var(--sp-accent-bg);color:var(--sp-accent-text);border-color:var(--sp-accent)}
`;

function ScheduleTab({ projects, loading, onUpdateItem, onAddItem, onRemoveItem, onRenameProject, onAddProject }) {
  const [zoom, setZoom] = useState("month");
  const [delayFilter, setDelayFilter] = useState(false);
  const [patternMode, setPatternMode] = useState(false);
  const [viewMode, setViewMode] = useState("view"); // 'view' | 'edit'
  const [visibleCats, setVisibleCats] = useState(() => new Set(SCHED_CATS));
  const [expanded, setExpanded] = useState({});
  const [laneCollapsed, setLaneCollapsed] = useState({});
  const [editingNameId, setEditingNameId] = useState(null);
  const [nameDraft, setNameDraft] = useState("");
  const [showAddForm, setShowAddForm] = useState(false);
  const [newProj, setNewProj] = useState({ name: "", tier: TIERS[0], start: "", target: "" });
  const [newItemDraft, setNewItemDraft] = useState({}); // { [projectId]: {name, cat} }
  const [tooltip, setTooltip] = useState(null);
  const [scrollTick, setScrollTick] = useState(0);
  const scrollRef = useRef(null);

  const today = useMemo(() => {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    return d;
  }, []);

  const visibleProjects = projects.filter((p) => !delayFilter || schedIsDelayed(p.items, p.start, p.target, today));
  const [rangeStart, rangeEnd] = schedGetFullRange(projects, today);
  const pxPerDay = SCHED_PX_PER_DAY[zoom];
  const totalWidth = Math.max(0, Math.round(daysBetween(rangeStart, rangeEnd) * pxPerDay));
  const todayPx = schedPx(today, rangeStart, pxPerDay);

  useEffect(() => {
    if (scrollTick === 0) return;
    const el = scrollRef.current;
    if (!el) return;
    el.scrollLeft = Math.max(0, todayPx - el.clientWidth / 2 + SCHED_LEFT_COL);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [scrollTick]);

  function jumpToday() {
    setScrollTick((t) => t + 1);
  }
  function changeZoom(z) {
    setZoom(z);
    setScrollTick((t) => t + 1);
  }
  function toggleCat(c) {
    setVisibleCats((prev) => {
      const next = new Set(prev);
      if (next.has(c)) next.delete(c);
      else next.add(c);
      return next;
    });
  }
  function toggleDetail(id) {
    setExpanded((prev) => ({ ...prev, [id]: !prev[id] }));
  }
  function toggleLane(id) {
    setLaneCollapsed((prev) => ({ ...prev, [id]: !prev[id] }));
  }
  function startEditName(p) {
    setEditingNameId(p.id);
    setNameDraft(p.name);
  }
  function commitEditName(id) {
    const v = nameDraft.trim();
    if (v) onRenameProject(id, v);
    setEditingNameId(null);
  }
  function submitNewProject() {
    if (!newProj.name.trim() || !newProj.start || !newProj.target) {
      alert("프로젝트명, 시작일, 오픈예정일을 입력해주세요.");
      return;
    }
    onAddProject({ name: newProj.name.trim(), tier: newProj.tier, start: newProj.start, target: newProj.target });
    setNewProj({ name: "", tier: TIERS[0], start: "", target: "" });
    setShowAddForm(false);
  }
  function submitNewItem(projectId) {
    const draft = newItemDraft[projectId] || { name: "", cat: SCHED_CATS[0] };
    if (!draft.name.trim()) return;
    onAddItem(projectId, draft.name.trim(), draft.cat || SCHED_CATS[0]);
    setNewItemDraft((prev) => ({ ...prev, [projectId]: { name: "", cat: draft.cat || SCHED_CATS[0] } }));
  }
  function showTooltip(e, lines) {
    setTooltip({ x: e.clientX + 12, y: e.clientY + 12, lines });
  }
  function moveTooltip(e) {
    setTooltip((prev) => (prev ? { ...prev, x: e.clientX + 12, y: e.clientY + 12 } : prev));
  }
  function hideTooltip() {
    setTooltip(null);
  }

  function renderTimelineHeader() {
    const months = [];
    let cur = new Date(rangeStart.getFullYear(), rangeStart.getMonth(), 1);
    while (cur < rangeEnd) {
      const next = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
      months.push({
        key: `${cur.getFullYear()}-${cur.getMonth()}`,
        label: `${cur.getFullYear()}.${String(cur.getMonth() + 1).padStart(2, "0")}`,
        w: daysBetween(cur, next) * pxPerDay,
      });
      cur = next;
    }
    const subTicks = [];
    if (zoom === "week") {
      let wc = new Date(rangeStart);
      while (wc < rangeEnd) {
        subTicks.push({ key: fmtLocalDate(wc), label: `${wc.getMonth() + 1}/${wc.getDate()}`, w: 7 * pxPerDay });
        wc = new Date(wc.getTime() + 7 * 86400000);
      }
    } else if (zoom === "day") {
      let dc = new Date(rangeStart);
      while (dc < rangeEnd) {
        subTicks.push({ key: fmtLocalDate(dc), label: `${dc.getDate()}`, w: pxPerDay });
        dc = new Date(dc.getTime() + 86400000);
      }
    }
    return (
      <div className="sp-tlheader" style={{ width: totalWidth }}>
        <div className="sp-monthrow">
          {months.map((m) => (
            <div key={m.key} className="sp-monthcell" style={{ width: m.w }}>{m.label}</div>
          ))}
        </div>
        {subTicks.length > 0 && (
          <div className="sp-subrow">
            {subTicks.map((t) => (
              <div key={t.key} className="sp-subcell" style={{ width: t.w }}>{t.label}</div>
            ))}
          </div>
        )}
        <div className="sp-todayline" style={{ left: todayPx }} />
      </div>
    );
  }

  function renderSegment(it) {
    const r = schedItemRange(it, today);
    if (!r) return null;
    const actLeft = schedPx(r.start, rangeStart, pxPerDay);
    const actRight = schedPx(r.end, rangeStart, pxPerDay);
    const isDelayedItem = it.delay || it.status === "확인필요";
    const planned = it.dueDate ? toDate(it.dueDate) : null;
    const lines = schedTooltipLines(it, today);
    const handlers = {
      onMouseEnter: (e) => showTooltip(e, lines),
      onMouseMove: moveTooltip,
      onMouseLeave: hideTooltip,
    };

    if (planned && planned > r.start) {
      const planLeft = actLeft, planRight = schedPx(planned, rangeStart, pxPerDay);
      const showOverrun = actRight > planRight + 1;
      return [
        <div key={it.id} className={`sp-seg sp-seg-planned ${schedSegClass(it.status)}`} style={{ left: planLeft, width: Math.max(planRight - planLeft, 4) }} {...handlers}>
          <span className="sp-icotxt">{it.name}</span>
        </div>,
        showOverrun && (
          <div key={it.id + "_o"} className="sp-seg sp-seg-overrun" style={{ left: planRight, width: Math.max(actRight - planRight, 4) }} {...handlers} />
        ),
      ];
    }
    const width = Math.max(actRight - actLeft, 4);
    const showText = width > 46;
    return (
      <div
        key={it.id}
        className={`sp-seg ${schedSegClass(it.status)} ${isDelayedItem ? "sp-seg-delayed" : ""} ${patternMode ? "sp-seg-pattern" : ""}`}
        style={{ left: actLeft, width }}
        {...handlers}
      >
        <span className="sp-icotxt">{isDelayedItem ? "🚩 " : ""}{showText ? it.name : ""}</span>
      </div>
    );
  }

  function renderDetail(p) {
    const items = p.items || [];
    const prog = Math.round(schedProjectProgress(items));
    const draft = newItemDraft[p.id] || { name: "", cat: SCHED_CATS[0] };
    return (
      <div className="sp-detail">
        <div className="sp-progresswrap">
          <div className="sp-ptext">전체 진행률 {prog}% (기대치 {Math.round(schedExpectedProgress(p.start, p.target, today))}%)</div>
          <div className="sp-pbar"><div className="sp-pfill" style={{ width: `${prog}%` }} /></div>
        </div>
        <div className="sp-tablewrap">
          <table className="sp-mstable">
            <thead>
              <tr>
                <th>항목</th><th>카테고리</th><th>Status</th><th>Delay</th>
                <th>계획 종료일</th><th>실제 시작일</th><th>실제 종료일</th>
                <th>담당자</th><th>메모</th><th></th>
              </tr>
            </thead>
            <tbody>
              {items.map((it) => (
                <tr key={it.id}>
                  <td className="sp-itemname">{it.name}</td>
                  <td>
                    <select className="sp-catsel" value={it.cat} onChange={(e) => onUpdateItem(p.id, it.id, "cat", e.target.value)}>
                      {SCHED_CATS.map((c) => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </td>
                  <td>
                    <select className={`sp-pill sp-pillst-${it.status}`} value={it.status} onChange={(e) => onUpdateItem(p.id, it.id, "status", e.target.value)}>
                      {SCHED_STATUSES.map((st) => <option key={st} value={st}>{st}</option>)}
                    </select>
                  </td>
                  <td className="sp-delcell">
                    <input type="checkbox" checked={!!it.delay} onChange={(e) => onUpdateItem(p.id, it.id, "delay", e.target.checked)} />
                  </td>
                  <td><input type="date" className="sp-datein" value={it.dueDate || ""} onChange={(e) => onUpdateItem(p.id, it.id, "dueDate", e.target.value)} /></td>
                  <td><input type="date" className="sp-datein" value={it.workStart || ""} onChange={(e) => onUpdateItem(p.id, it.id, "workStart", e.target.value)} /></td>
                  <td><input type="date" className="sp-datein" value={it.workEnd || ""} onChange={(e) => onUpdateItem(p.id, it.id, "workEnd", e.target.value)} /></td>
                  <td><input type="text" className="sp-datein sp-w70" placeholder="담당자" value={it.assignee || ""} onChange={(e) => onUpdateItem(p.id, it.id, "assignee", e.target.value)} /></td>
                  <td><input type="text" className="sp-datein sp-w130" placeholder="메모" value={it.memo || ""} onChange={(e) => onUpdateItem(p.id, it.id, "memo", e.target.value)} /></td>
                  <td className="sp-delcell">
                    <button className="sp-rmbtn" onClick={() => onRemoveItem(p.id, it.id)}>삭제</button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="sp-addmsrow">
          <input
            placeholder="새 항목명"
            style={{ width: 160 }}
            value={draft.name}
            onChange={(e) => setNewItemDraft((prev) => ({ ...prev, [p.id]: { ...draft, name: e.target.value } }))}
            onKeyDown={(e) => e.key === "Enter" && submitNewItem(p.id)}
          />
          <select value={draft.cat} onChange={(e) => setNewItemDraft((prev) => ({ ...prev, [p.id]: { ...draft, cat: e.target.value } }))}>
            {SCHED_CATS.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
          <button onClick={() => submitNewItem(p.id)}>+ 항목 추가</button>
        </div>
      </div>
    );
  }

  function renderProjectRow(p) {
    const items = p.items || [];
    const delayed = schedIsDelayed(items, p.start, p.target, today);
    const collapsed = !!laneCollapsed[p.id];
    const lanes = SCHED_CATS.filter((c) => visibleCats.has(c) && items.some((it) => it.cat === c));
    const laneH = 32;
    const targetDate = p.target ? toDate(p.target) : null;
    const targetPx = targetDate ? schedPx(targetDate, rangeStart, pxPerDay) : null;
    const progress = schedProjectProgress(items);
    const overlaps = collapsed ? [] : schedFindOverlaps(items, today);
    const startDate = p.start ? toDate(p.start) : null;
    const collapsedLeft = startDate ? schedPx(startDate, rangeStart, pxPerDay) : 0;
    const collapsedRight = targetPx !== null ? targetPx : collapsedLeft;

    return (
      <div key={p.id} className="sp-row" style={{ gridTemplateColumns: `${SCHED_LEFT_COL}px ${totalWidth}px` }}>
        <div className="sp-rowmeta">
          <div className="sp-rowmeta-top">
            <button className="sp-collapsebtn" onClick={() => toggleLane(p.id)}>{collapsed ? "▶" : "▼"}</button>
            {editingNameId === p.id ? (
              <input
                autoFocus
                className="sp-nameinput"
                value={nameDraft}
                onChange={(e) => setNameDraft(e.target.value)}
                onBlur={() => commitEditName(p.id)}
                onKeyDown={(e) => e.key === "Enter" && commitEditName(p.id)}
              />
            ) : (
              <div className="sp-name">
                {p.name || "(이름 없음)"}
                {viewMode === "edit" && <span className="sp-pencil" onClick={() => startEditName(p)}>✎</span>}
              </div>
            )}
          </div>
          <div className="sp-pctrow">
            <span className="sp-pct" style={{ color: delayed ? "var(--sp-check)" : "var(--sp-done)" }}>{Math.round(progress)}%</span>
            <span className={`sp-badge ${delayed ? "sp-badge-delay" : "sp-badge-ontrack"}`}>{delayed ? "🚩 지연" : "정상"}</span>
          </div>
          <div className="sp-meta">
            {p.tier || ""} · {schedDdayLabel(p.target, today)}
            {schedTeamWorkDone(items) && <span className="sp-badge sp-badge-ontrack">✅ 현장 업무 완료</span>}
          </div>
          {viewMode === "edit" && (
            <button className="sp-editentry" onClick={() => toggleDetail(p.id)}>
              {expanded[p.id] ? "▲ 접기" : "📝 상세 일정 입력"}
            </button>
          )}
        </div>
        <div
          className="sp-tlwrap"
          style={{ width: totalWidth, minHeight: collapsed ? 40 : Math.max(lanes.length * laneH, 40) + 16 }}
        >
          {!collapsed
            ? lanes.map((cat) => (
                <div className="sp-lane" key={cat}>
                  <div className="sp-lanelabel">{cat}</div>
                  {items.filter((it) => it.cat === cat).map((it) => renderSegment(it))}
                </div>
              ))
            : (
                <div className="sp-lane" style={{ height: 14 }}>
                  <div className="sp-seg-collapsed" style={{ left: collapsedLeft, width: Math.max(collapsedRight - collapsedLeft, 4) }}>
                    <div className="sp-seg-collapsed-fill" style={{ width: `${progress}%`, background: delayed ? "var(--sp-check)" : "var(--sp-done)" }} />
                  </div>
                </div>
              )}
          {overlaps.map((d, i) => (
            <div
              key={`ob_${i}`}
              className="sp-overlapband"
              style={{ left: schedPx(d, rangeStart, pxPerDay) - 3 }}
              onMouseEnter={(e) => showTooltip(e, [`${fmtLocalDate(d)} 공정 겹침`])}
              onMouseMove={moveTooltip}
              onMouseLeave={hideTooltip}
            />
          ))}
          {overlaps.map((d, i) => (
            <div key={`oi_${i}`} className="sp-overlapicon" style={{ left: schedPx(d, rangeStart, pxPerDay) }}>⚡</div>
          ))}
          <div className="sp-todayline" style={{ left: todayPx, opacity: 0.4 }} />
          {targetPx !== null && (
            <div
              className="sp-milestone"
              style={{ left: targetPx }}
              onMouseEnter={(e) => showTooltip(e, [`오픈예정일 ${p.target} (${schedDdayLabel(p.target, today)})`])}
              onMouseMove={moveTooltip}
              onMouseLeave={hideTooltip}
            >
              <div className="sp-diamond" />
              <div className="sp-mlabel">오픈</div>
            </div>
          )}
        </div>
        {viewMode === "edit" && expanded[p.id] && renderDetail(p)}
      </div>
    );
  }

  return (
    <div className="sp-root">
      <style>{SP_STYLE}</style>
      <div className="sp-toolbar">
        <div className="sp-zoombtns">
          {[["month", "월"], ["week", "주"], ["day", "일"]].map(([z, label]) => (
            <button key={z} className={zoom === z ? "active" : ""} onClick={() => changeZoom(z)}>{label}</button>
          ))}
        </div>
        <button className="sp-todaybtn" onClick={jumpToday}>오늘로 이동</button>
        <button className={`sp-chk ${delayFilter ? "active" : ""}`} onClick={() => setDelayFilter((v) => !v)}>🚩 지연만 보기</button>
        <button className={`sp-chk ${patternMode ? "active" : ""}`} onClick={() => setPatternMode((v) => !v)}>🎨 패턴 표시(색약 모드)</button>
        <button
          className="sp-modebtn"
          onClick={() => {
            setViewMode((v) => (v === "view" ? "edit" : "view"));
            if (viewMode === "edit") setExpanded({});
          }}
        >
          {viewMode === "view" ? "✏️ 관리자 뷰로 전환" : "👁️ 대시보드 뷰로 전환"}
        </button>
        {viewMode === "edit" && (
          <button className="sp-addbtn" onClick={() => setShowAddForm((v) => !v)}>+ 새 프로젝트</button>
        )}
      </div>

      <div className="sp-filterbar">
        공정 필터:
        {SCHED_CATS.map((c) => (
          <span key={c} className={`sp-catchip ${visibleCats.has(c) ? "" : "sp-off"}`} onClick={() => toggleCat(c)}>
            <span className="sp-dot" style={{ background: SCHED_CAT_DOT[c] }} />{c}
          </span>
        ))}
      </div>

      <div className="sp-legend">
        <span><span className="sp-dot" style={{ background: "var(--sp-progress)" }} />진행중</span>
        <span><span className="sp-dot" style={{ background: "var(--sp-check)" }} />확인필요</span>
        <span><span className="sp-dot" style={{ background: "var(--sp-done)" }} />완료</span>
        <span><span className="sp-dot" style={{ background: "var(--sp-done)", opacity: 0.55, border: "1px dashed #fff" }} />계획 구간</span>
        <span><span className="sp-dot sp-hatch" />지연폭</span>
        <span>🚩 지연 항목</span>
        <span>⚡ 공정 겹침</span>
        <span><span className="sp-diamond-legend" />오픈예정일</span>
        <span><span className="sp-todaydot" />오늘</span>
      </div>

      {viewMode === "edit" && showAddForm && (
        <div className="sp-formrow">
          <label>프로젝트명<input value={newProj.name} onChange={(e) => setNewProj((p) => ({ ...p, name: e.target.value }))} /></label>
          <label>
            브랜드티어
            <select value={newProj.tier} onChange={(e) => setNewProj((p) => ({ ...p, tier: e.target.value }))}>
              {TIERS.map((t) => <option key={t} value={t}>{t}</option>)}
            </select>
          </label>
          <label>시작일<input type="date" value={newProj.start} onChange={(e) => setNewProj((p) => ({ ...p, start: e.target.value }))} /></label>
          <label>오픈예정일<input type="date" value={newProj.target} onChange={(e) => setNewProj((p) => ({ ...p, target: e.target.value }))} /></label>
          <button onClick={submitNewProject}>추가</button>
        </div>
      )}

      <div className="sp-board">
        {loading ? (
          <div className="sp-empty">불러오는 중...</div>
        ) : (
          <div className="sp-boardscroll" ref={scrollRef}>
            <div className="sp-boardinner" style={{ width: SCHED_LEFT_COL + totalWidth }}>
              {visibleProjects.length === 0 ? (
                <div className="sp-empty">
                  {projects.length === 0 ? '아직 프로젝트가 없어요. "+ 새 프로젝트"로 추가해보세요.' : "조건에 맞는 프로젝트가 없어요."}
                </div>
              ) : (
                <>
                  <div className="sp-headerrow" style={{ gridTemplateColumns: `${SCHED_LEFT_COL}px ${totalWidth}px` }}>
                    <div />
                    {renderTimelineHeader()}
                  </div>
                  {visibleProjects.map((p) => renderProjectRow(p))}
                  {viewMode === "edit" && (
                    <div className="sp-addrow" onClick={() => setShowAddForm(true)}>+ 새 프로젝트 추가</div>
                  )}
                </>
              )}
            </div>
          </div>
        )}
      </div>

      {tooltip && (
        <div className="sp-tooltip" style={{ left: tooltip.x, top: tooltip.y }}>
          {tooltip.lines.map((line, i) => <div key={i}>{line}</div>)}
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [projectName, setProjectName] = useState("");
  const [tier, setTier] = useState(TIERS[2]);
  const [totalBudget, setTotalBudget] = useState(0); // 전체 공사비 중 오픈바이징팀에 배정된 예산

  const [categories, setCategories] = useState(["트윈", "더블", "패밀리", "스위트"]);
  const [newCategory, setNewCategory] = useState("");

  const [irregularOptions, setIrregularOptions] = useState(["마사지체어", "발코니", "복층"]);
  const [newIrregular, setNewIrregular] = useState("");

  // 브랜드 룸 네이밍 가이드라인 v1.0 — 라이프스타일/콘셉트 호텔용 브랜드 고유 룸 네임 (예: '시즈쿠')
  const [brandRoomName, setBrandRoomName] = useState("");
  // 객실 콘텐츠 (Room Feature) — 프로젝트별 커스텀, 룸 네이밍에 반영
  const [roomFeatures, setRoomFeatures] = useState(["스파", "테라스"]);
  const [newRoomFeature, setNewRoomFeature] = useState("");
  // 전망 (View Type) — 프로젝트별 커스텀, 필요한 경우에만 룸 네이밍에 반영
  const [viewTypes, setViewTypes] = useState(["오션뷰", "시티뷰"]);
  const [newViewType, setNewViewType] = useState("");

  const [floors, setFloors] = useState(["3F", "4F", "5F"]);
  const [newFloor, setNewFloor] = useState("");
  const [bulkFrom, setBulkFrom] = useState("");
  const [bulkTo, setBulkTo] = useState("");

  function floorOrderKey(label) {
    const m = label.match(/^B(\d+)/i);
    if (m) return -parseInt(m[1], 10);
    const n = label.match(/^(\d+)/);
    if (n) return parseInt(n[1], 10);
    return Infinity;
  }
  function sortFloors(list) {
    return [...list].sort((a, b) => floorOrderKey(a) - floorOrderKey(b));
  }

  const [roomTypes, setRoomTypes] = useState([]);
  const [activeTab, setActiveTab] = useState("main"); // "main" | "schedule"
  const [scheduleStart, setScheduleStart] = useState(() => new Date().toISOString().slice(0, 10));
  const [scheduleTargetDate, setScheduleTargetDate] = useState("");
  const [scheduleItems, setScheduleItems] = useState(() => makeDefaultScheduleItems());
  // 공정표 탭(포트폴리오 뷰)에서 쓰는, 현재 편집 중인 프로젝트 외 나머지 프로젝트들의 캐시
  const [scheduleProjectsCache, setScheduleProjectsCache] = useState({}); // { [projectId]: fullSavedPayload }
  const [scheduleTabLoading, setScheduleTabLoading] = useState(false);
  const [showFloorPlan, setShowFloorPlan] = useState(true);
  const [floorPlanAutoOffApplied, setFloorPlanAutoOffApplied] = useState(false);

  // 발주 품목: FF&E는 룸타입별, OS&E는 공통 리스트 (각 항목은 예산단가 unitPrice + 집행단가 actualUnitPrice)
  const [ffeItems, setFfeItems] = useState({}); // { [roomTypeId]: [{id,name,unitPrice,actualUnitPrice,qtyPerRoom}] }
  const [oseItems, setOseItems] = useState([]); // [{id,name,unitPrice,actualUnitPrice,qtyPerRoom}]

  // 현장지출 / 인건비 지출 / 예산외 지출 (품목형이 아닌 금액 직접 입력 방식, 예산금액/집행금액)
  const [siteExpenses, setSiteExpenses] = useState([]); // [{id,name,budgetAmount,actualAmount}]
  const [laborExpenses, setLaborExpenses] = useState([]);
  const [extraExpenses, setExtraExpenses] = useState([]);

  function makeExpenseHandlers(setList) {
    const add = () =>
      setList((prev) => [...prev, { id: nextId(), name: "", budgetAmount: 0, actualAmount: 0 }]);
    const update = (id, field, value) =>
      setList((prev) =>
        prev.map((it) =>
          it.id === id
            ? { ...it, [field]: field === "name" ? value : Math.max(0, parseFloat(value || "0") || 0) }
            : it
        )
      );
    const remove = (id) => setList((prev) => prev.filter((it) => it.id !== id));
    return { add, update, remove };
  }
  const siteExpenseHandlers = makeExpenseHandlers(setSiteExpenses);
  const laborExpenseHandlers = makeExpenseHandlers(setLaborExpenses);
  const extraExpenseHandlers = makeExpenseHandlers(setExtraExpenses);
  const [pasteOpenFor, setPasteOpenFor] = useState(null); // roomTypeId | "OSE" | null
  const [pasteText, setPasteText] = useState("");
  const [copySourceFor, setCopySourceFor] = useState({}); // { [roomTypeId]: sourceRoomTypeId }
  const [basicPreset, setBasicPreset] = useState(
    DEFAULT_BASIC_PRESET.map((name) => ({ id: nextId(), name }))
  );
  const [newPresetItem, setNewPresetItem] = useState("");
  const [presetEditorOpen, setPresetEditorOpen] = useState(false);

  function addPresetItem() {
    const v = newPresetItem.trim();
    if (!v) return;
    setBasicPreset((prev) => [...prev, { id: nextId(), name: v }]);
    setNewPresetItem("");
  }
  function removePresetItem(id) {
    setBasicPreset((prev) => prev.filter((it) => it.id !== id));
  }
  function updatePresetItemName(id, name) {
    setBasicPreset((prev) => prev.map((it) => (it.id === id ? { ...it, name } : it)));
  }

  function addFfeItem(roomTypeId) {
    setFfeItems((prev) => ({
      ...prev,
      [roomTypeId]: [
        ...(prev[roomTypeId] || []),
        { id: nextId(), name: "", unitPrice: 0, actualUnitPrice: 0, installUnitPrice: 0, installActualUnitPrice: 0, qtyPerRoom: 1 },
      ],
    }));
  }
  function bulkAddFfeItems(roomTypeId, rows) {
    const parsed = rows
      .split("\n")
      .map((line) => line.split(/\t|,/).map((s) => s.trim()))
      .filter((cols) => cols[0])
      .map((cols) => ({
        id: nextId(),
        name: cols[0] || "",
        unitPrice: Math.max(0, parseFloat(cols[1] || "0") || 0),
        actualUnitPrice: 0,
        installUnitPrice: 0,
        installActualUnitPrice: 0,
        qtyPerRoom: cols[2] ? Math.max(0, parseFloat(cols[2]) || 0) : 1,
      }));
    if (parsed.length === 0) return;
    setFfeItems((prev) => ({
      ...prev,
      [roomTypeId]: [...(prev[roomTypeId] || []), ...parsed],
    }));
  }
  function copyFfeItems(fromRoomTypeId, toRoomTypeId) {
    const source = ffeItems[fromRoomTypeId] || [];
    if (source.length === 0) return;
    setFfeItems((prev) => ({
      ...prev,
      [toRoomTypeId]: [
        ...(prev[toRoomTypeId] || []),
        ...source.map((it) => ({ ...it, id: nextId() })),
      ],
    }));
  }
  function addPresetFfeItems(roomTypeId) {
    setFfeItems((prev) => ({
      ...prev,
      [roomTypeId]: [
        ...(prev[roomTypeId] || []),
        ...basicPreset.map((p) => ({ id: nextId(), name: p.name, unitPrice: 0, actualUnitPrice: 0, installUnitPrice: 0, installActualUnitPrice: 0, qtyPerRoom: 1 })),
      ],
    }));
  }
  function updateFfeItem(roomTypeId, itemId, field, value) {
    setFfeItems((prev) => ({
      ...prev,
      [roomTypeId]: (prev[roomTypeId] || []).map((it) =>
        it.id === itemId
          ? { ...it, [field]: field === "name" ? value : Math.max(0, parseFloat(value || "0") || 0) }
          : it
      ),
    }));
  }
  function removeFfeItem(roomTypeId, itemId) {
    setFfeItems((prev) => ({
      ...prev,
      [roomTypeId]: (prev[roomTypeId] || []).filter((it) => it.id !== itemId),
    }));
  }

  function addOseItem() {
    setOseItems((prev) => [...prev, { id: nextId(), name: "", unitPrice: 0, actualUnitPrice: 0, installUnitPrice: 0, installActualUnitPrice: 0, qtyPerRoom: 1 }]);
  }
  function bulkAddOseItems(rows) {
    const parsed = rows
      .split("\n")
      .map((line) => line.split(/\t|,/).map((s) => s.trim()))
      .filter((cols) => cols[0])
      .map((cols) => ({
        id: nextId(),
        name: cols[0] || "",
        unitPrice: Math.max(0, parseFloat(cols[1] || "0") || 0),
        actualUnitPrice: 0,
        installUnitPrice: 0,
        installActualUnitPrice: 0,
        qtyPerRoom: cols[2] ? Math.max(0, parseFloat(cols[2]) || 0) : 1,
      }));
    if (parsed.length === 0) return;
    setOseItems((prev) => [...prev, ...parsed]);
  }
  function updateOseItem(itemId, field, value) {
    setOseItems((prev) =>
      prev.map((it) =>
        it.id === itemId
          ? { ...it, [field]: field === "name" ? value : Math.max(0, parseFloat(value || "0") || 0) }
          : it
      )
    );
  }
  function removeOseItem(itemId) {
    setOseItems((prev) => prev.filter((it) => it.id !== itemId));
  }

  const won = (n) => `${Math.round(n).toLocaleString("ko-KR")}원`;

  const [draft, setDraft] = useState({
    bed: BED_TYPES[0],
    bathtub: BATHTUB[1],
    category: categories[0] || "",
    irregular: [],
    mattressQty: 1,
    grade: "Superior",
    features: [],
    view: "",
    includeBedInName: false,
    includeViewInName: false,
    customName: "",
  });

  function toggleDraftFeature(opt) {
    setDraft((d) => ({
      ...d,
      features: d.features.includes(opt)
        ? d.features.filter((o) => o !== opt)
        : [...d.features, opt],
    }));
  }

  // 브랜드 룸 네이밍 가이드라인 v1.0 — 객실등급 + 객실콘텐츠 + (필요시)침대구성 + (필요시)전망 조합으로 자동 생성
  function generateRoomName(rt) {
    if (rt.customName && rt.customName.trim()) return rt.customName.trim();
    const parts = [];
    if (brandRoomName.trim()) parts.push(brandRoomName.trim());
    parts.push(rt.grade || "Superior");
    if (rt.features && rt.features.length > 0) parts.push(rt.features.join(" "));
    if (rt.includeBedInName) parts.push(rt.category);
    if (rt.includeViewInName && rt.view) parts.push(rt.view);
    return parts.join(" ");
  }

  const categoryColor = (category) => {
    const idx = categories.indexOf(category);
    return CATEGORY_COLORS[idx % CATEGORY_COLORS.length] || CATEGORY_COLORS[0];
  };

  function addCategory() {
    const v = newCategory.trim();
    if (!v || categories.includes(v)) return;
    setCategories([...categories, v]);
    setNewCategory("");
  }
  function removeCategory(cat) {
    setCategories(categories.filter((c) => c !== cat));
  }
  function addIrregular() {
    const v = newIrregular.trim();
    if (!v || irregularOptions.includes(v)) return;
    setIrregularOptions([...irregularOptions, v]);
    setNewIrregular("");
  }
  function removeIrregular(opt) {
    setIrregularOptions(irregularOptions.filter((o) => o !== opt));
  }
  function addRoomFeature() {
    const v = newRoomFeature.trim();
    if (!v || roomFeatures.includes(v)) return;
    setRoomFeatures([...roomFeatures, v]);
    setNewRoomFeature("");
  }
  function removeRoomFeature(opt) {
    setRoomFeatures(roomFeatures.filter((o) => o !== opt));
  }
  function addViewType() {
    const v = newViewType.trim();
    if (!v || viewTypes.includes(v)) return;
    setViewTypes([...viewTypes, v]);
    setNewViewType("");
  }
  function removeViewType(opt) {
    setViewTypes(viewTypes.filter((o) => o !== opt));
  }
  function addFloor() {
    const v = newFloor.trim();
    if (!v || floors.includes(v)) return;
    setFloors(sortFloors([...floors, v]));
    setNewFloor("");
  }
  function removeFloor(f) {
    setFloors(floors.filter((x) => x !== f));
    setRoomTypes(roomTypes.map((rt) => {
      const copy = { ...rt.byFloor };
      delete copy[f];
      return { ...rt, byFloor: copy };
    }));
  }
  function addFloorRange() {
    const from = parseInt(bulkFrom, 10);
    const to = parseInt(bulkTo, 10);
    if (Number.isNaN(from) || Number.isNaN(to) || from > to) return;
    const generated = [];
    for (let n = from; n <= to; n++) generated.push(`${n}F`);
    const merged = [...floors];
    generated.forEach((f) => {
      if (!merged.includes(f)) merged.push(f);
    });
    setFloors(sortFloors(merged));
    setRoomTypes(
      roomTypes.map((rt) => ({
        ...rt,
        byFloor: { ...Object.fromEntries(generated.map((f) => [f, 0])), ...rt.byFloor },
      }))
    );
    setBulkFrom("");
    setBulkTo("");
  }

  function toggleDraftIrregular(opt) {
    setDraft((d) => ({
      ...d,
      irregular: d.irregular.includes(opt)
        ? d.irregular.filter((o) => o !== opt)
        : [...d.irregular, opt],
    }));
  }

  function addRoomType() {
    if (!draft.category) return;
    const exists = roomTypes.some(
      (rt) =>
        rt.bed === draft.bed &&
        rt.bathtub === draft.bathtub &&
        rt.category === draft.category &&
        rt.mattressQty === draft.mattressQty &&
        rt.grade === draft.grade &&
        rt.view === draft.view &&
        JSON.stringify([...rt.irregular].sort()) === JSON.stringify([...draft.irregular].sort()) &&
        JSON.stringify([...rt.features].sort()) === JSON.stringify([...draft.features].sort())
    );
    if (exists) return;
    setRoomTypes([
      ...roomTypes,
      {
        id: nextId(),
        bed: draft.bed,
        bathtub: draft.bathtub,
        category: draft.category,
        irregular: draft.irregular,
        mattressQty: draft.mattressQty,
        grade: draft.grade,
        features: draft.features,
        view: draft.view,
        includeBedInName: draft.includeBedInName,
        includeViewInName: draft.includeViewInName,
        customName: draft.customName,
        otaBedCount: "",
        otaBedSize: "",
        otaMaxOccupancy: "",
        otaFacilities: "",
        roomNumbers: [],
        byFloor: Object.fromEntries(floors.map((f) => [f, 0])),
      },
    ]);
    setDraft((d) => ({ ...d, irregular: [], features: [], mattressQty: 1, customName: "" }));
  }

  function updateRoomTypeField(id, field, value) {
    setRoomTypes((prev) => prev.map((rt) => (rt.id === id ? { ...rt, [field]: value } : rt)));
  }

  function removeRoomType(id) {
    setRoomTypes(roomTypes.filter((rt) => rt.id !== id));
  }

  function setQty(id, floor, value) {
    const n = Math.max(0, parseInt(value || "0", 10) || 0);
    setRoomTypes(
      roomTypes.map((rt) =>
        rt.id === id ? { ...rt, byFloor: { ...rt.byFloor, [floor]: n } } : rt
      )
    );
  }
  function setMattressQty(id, value) {
    const n = Math.max(1, parseInt(value || "1", 10) || 1);
    setRoomTypes(roomTypes.map((rt) => (rt.id === id ? { ...rt, mattressQty: n } : rt)));
  }

  // ---- 엑셀 업로드로 룸타입 + 호수 일괄 생성 ----
  const [importSummary, setImportSummary] = useState(null);
  const [importError, setImportError] = useState("");
  const [overwriteOnImport, setOverwriteOnImport] = useState(true);

  // ---- 품목 마스터 카탈로그 (엑셀 업로드 후 골라서 삽입) ----
  const [itemCatalog, setItemCatalog] = useState([]); // [{id,name,unitPrice,category}]
  const [catalogSummary, setCatalogSummary] = useState(null);
  const [catalogError, setCatalogError] = useState("");
  const [pickerOpenFor, setPickerOpenFor] = useState(null); // roomTypeId | "OSE" | null
  const [pickerSelected, setPickerSelected] = useState(new Set());
  const [pickerCategoryFilter, setPickerCategoryFilter] = useState("전체");

  // ---- 공유 저장소(팀원과 데이터 공유) + 다중 프로젝트 ----
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [lastSaved, setLastSaved] = useState(null);
  const [saveError, setSaveError] = useState("");
  const [supabaseSyncError, setSupabaseSyncError] = useState("");
  const [loadNotice, setLoadNotice] = useState("");
  const [projectList, setProjectList] = useState([]); // [{id, name}]
  const [currentProjectId, setCurrentProjectId] = useState(null);

  function maxIdIn(data) {
    let max = 0;
    const scan = (obj) => {
      if (Array.isArray(obj)) obj.forEach(scan);
      else if (obj && typeof obj === "object") {
        if (typeof obj.id === "number") max = Math.max(max, obj.id);
        Object.values(obj).forEach(scan);
      }
    };
    scan(data);
    return max;
  }

  function resetProjectState() {
    setProjectName("");
    setTier(TIERS[2]);
    setTotalBudget(0);
    setCategories(["트윈", "더블", "패밀리", "스위트"]);
    setIrregularOptions(["마사지체어", "발코니", "복층"]);
    setBrandRoomName("");
    setRoomFeatures(["스파", "테라스"]);
    setViewTypes(["오션뷰", "시티뷰"]);
    setFloors(["3F", "4F", "5F"]);
    setRoomTypes([]);
    setFfeItems({});
    setOseItems([]);
    setSiteExpenses([]);
    setLaborExpenses([]);
    setExtraExpenses([]);
    setBasicPreset(DEFAULT_BASIC_PRESET.map((name) => ({ id: nextId(), name })));
    setScheduleStart(new Date().toISOString().slice(0, 10));
    setScheduleTargetDate("");
    setScheduleItems(makeDefaultScheduleItems());
    setLastSaved(null);
  }

  function applyLoadedData(data) {
    if (data.projectName !== undefined) setProjectName(data.projectName);
    if (data.tier) setTier(data.tier);
    if (data.totalBudget !== undefined) setTotalBudget(data.totalBudget);
    if (data.categories) setCategories(data.categories);
    if (data.irregularOptions) setIrregularOptions(data.irregularOptions);
    if (data.brandRoomName !== undefined) setBrandRoomName(data.brandRoomName);
    if (data.roomFeatures) setRoomFeatures(data.roomFeatures);
    if (data.viewTypes) setViewTypes(data.viewTypes);
    if (data.floors) setFloors(data.floors);
    if (data.roomTypes) setRoomTypes(data.roomTypes);
    if (data.ffeItems) setFfeItems(data.ffeItems);
    if (data.oseItems) setOseItems(data.oseItems);
    if (data.siteExpenses) setSiteExpenses(data.siteExpenses);
    if (data.laborExpenses) setLaborExpenses(data.laborExpenses);
    if (data.extraExpenses) setExtraExpenses(data.extraExpenses);
    if (data.basicPreset) setBasicPreset(data.basicPreset);
    if (data.scheduleStart) setScheduleStart(data.scheduleStart);
    if (data.scheduleTargetDate !== undefined) setScheduleTargetDate(data.scheduleTargetDate);
    if (data.scheduleItems) setScheduleItems(data.scheduleItems);
    if (data.savedAt) setLastSaved(data.savedAt);
    idCounter = Math.max(idCounter, maxIdIn(data) + 1);
  }

  async function switchToProject(id) {
    setIsLoading(true);
    resetProjectState();
    try {
      const data = await getProjectData(id);
      if (data) {
        applyLoadedData(data);
        setLoadNotice("팀원들과 공유된 이전 데이터를 불러왔어요.");
      }
    } catch (err) {
      // 아직 저장 이력 없는 새 프로젝트이거나 네트워크 오류
    } finally {
      setCurrentProjectId(id);
      setIsLoading(false);
    }
  }

  async function createNewProject() {
    const name = window.prompt("새 프로젝트 이름을 입력하세요", "새 프로젝트");
    if (!name) return;
    const id = `proj_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
    const newList = [...projectList, { id, name }];
    setProjectList(newList);
    await saveProjectIndex(newList);
    resetProjectState();
    setProjectName(name);
    setCurrentProjectId(id);
    setIsLoading(false);
    setLoadNotice("");
  }

  useEffect(() => {
    (async () => {
      try {
        const catalog = await getCatalog();
        if (catalog && catalog.length) setItemCatalog(catalog);
      } catch (err) {
        // 카탈로그 없음 (최초 사용) 또는 네트워크 오류 — 조용히 넘어감
      }
      try {
        const list = await getProjectIndex();
        if (list.length === 0) {
          const id = `proj_${Date.now()}_init`;
          const initialList = [{ id, name: "기본 프로젝트" }];
          setProjectList(initialList);
          await saveProjectIndex(initialList);
          setCurrentProjectId(id);
          setIsLoading(false);
          return;
        }
        setProjectList(list);
        await switchToProject(list[0].id);
      } catch (err) {
        setSaveError("데이터를 불러오지 못했어요. APPS_SCRIPT_URL 설정을 확인해주세요.");
        setIsLoading(false);
      }
    })();
  }, []);

  function buildCurrentPayload(overrides = {}) {
    return {
      projectName,
      tier,
      totalBudget,
      categories,
      irregularOptions,
      brandRoomName,
      roomFeatures,
      viewTypes,
      floors,
      roomTypes,
      ffeItems,
      oseItems,
      siteExpenses,
      laborExpenses,
      extraExpenses,
      basicPreset,
      scheduleStart,
      scheduleTargetDate,
      scheduleItems,
      savedAt: new Date().toISOString(),
      ...overrides,
    };
  }

  async function saveProject() {
    if (!currentProjectId) return;
    setIsSaving(true);
    setSaveError("");
    try {
      const payload = buildCurrentPayload();
      const result = await saveProjectData(currentProjectId, payload);
      if (!result || result.error) throw new Error(result && result.error);
      setLastSaved(payload.savedAt);
      const updatedList = projectList.map((p) => (p.id === currentProjectId ? { ...p, name: projectName || p.name } : p));
      setProjectList(updatedList);
      await saveProjectIndex(updatedList);
      // Supabase room_types/order_items 동기화 — 기존 Apps Script 저장(위 로직)이 여전히 주 저장소이므로
      // 여기서 실패해도 전체 저장 실패로 취급하지 않고 별도 경고만 표시한다 (단계적 이전 중)
      try {
        const roomTypeIdMap = await saveRoomTypes(currentProjectId, roomTypes);
        await saveOrderItems(currentProjectId, ffeItems, oseItems, roomTypeIdMap);
        setSupabaseSyncError("");
      } catch (syncErr) {
        setSupabaseSyncError("Supabase 동기화에 실패했어요(기본 저장은 정상 완료됨).");
      }
    } catch (err) {
      setSaveError("저장에 실패했어요. 잠시 후 다시 시도해주세요.");
    } finally {
      setIsSaving(false);
    }
  }

  // ---- 공정표 탭(포트폴리오 뷰): 발주관리와 프로젝트 배열을 공유하되,
  // 현재 편집 중이 아닌 프로젝트들의 데이터는 별도 캐시로 불러온다 ----
  async function loadScheduleProjects() {
    const missing = projectList.filter((p) => p.id !== currentProjectId && !(p.id in scheduleProjectsCache));
    if (missing.length === 0) return;
    setScheduleTabLoading(true);
    try {
      const entries = await Promise.all(
        missing.map(async (p) => {
          try {
            const data = await getProjectData(p.id);
            return [p.id, data || {}];
          } catch (err) {
            return [p.id, {}];
          }
        })
      );
      setScheduleProjectsCache((prev) => {
        const next = { ...prev };
        entries.forEach(([id, data]) => {
          next[id] = data;
        });
        return next;
      });
    } finally {
      setScheduleTabLoading(false);
    }
  }

  useEffect(() => {
    if (activeTab === "schedule" && !isLoading) {
      loadScheduleProjects();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeTab, projectList, currentProjectId, isLoading]);

  async function persistScheduleForCurrent(patch) {
    const payload = buildCurrentPayload(patch);
    try {
      await saveProjectData(currentProjectId, payload);
      setLastSaved(payload.savedAt);
    } catch (err) {
      setSaveError("공정표 저장에 실패했어요. 잠시 후 다시 시도해주세요.");
    }
  }

  async function persistScheduleForProject(projectId, patch) {
    const base = scheduleProjectsCache[projectId] || {};
    const payload = { ...base, ...patch, savedAt: new Date().toISOString() };
    setScheduleProjectsCache((prev) => ({ ...prev, [projectId]: payload }));
    try {
      await saveProjectData(projectId, payload);
    } catch (err) {
      setSaveError("공정표 저장에 실패했어요. 잠시 후 다시 시도해주세요.");
    }
  }

  function updateScheduleItem(projectId, itemId, field, value) {
    if (projectId === currentProjectId) {
      const next = scheduleItems.map((it) => (it.id === itemId ? { ...it, [field]: value } : it));
      setScheduleItems(next);
      persistScheduleForCurrent({ scheduleItems: next });
    } else {
      const base = scheduleProjectsCache[projectId] || {};
      const next = (base.scheduleItems || []).map((it) => (it.id === itemId ? { ...it, [field]: value } : it));
      persistScheduleForProject(projectId, { scheduleItems: next });
    }
  }

  function addScheduleItem(projectId, name, cat) {
    const newItem = {
      id: nextId(),
      name,
      cat,
      status: "시작전",
      delay: false,
      dueDate: null,
      workStart: null,
      workEnd: null,
      assignee: "",
      memo: "",
    };
    if (projectId === currentProjectId) {
      const next = [...scheduleItems, newItem];
      setScheduleItems(next);
      persistScheduleForCurrent({ scheduleItems: next });
    } else {
      const base = scheduleProjectsCache[projectId] || {};
      const next = [...(base.scheduleItems || []), newItem];
      persistScheduleForProject(projectId, { scheduleItems: next });
    }
  }

  function removeScheduleItem(projectId, itemId) {
    if (projectId === currentProjectId) {
      const next = scheduleItems.filter((it) => it.id !== itemId);
      setScheduleItems(next);
      persistScheduleForCurrent({ scheduleItems: next });
    } else {
      const base = scheduleProjectsCache[projectId] || {};
      const next = (base.scheduleItems || []).filter((it) => it.id !== itemId);
      persistScheduleForProject(projectId, { scheduleItems: next });
    }
  }

  async function renameScheduleProject(projectId, name) {
    if (projectId === currentProjectId) {
      setProjectName(name);
      persistScheduleForCurrent({ projectName: name });
    } else {
      persistScheduleForProject(projectId, { projectName: name });
    }
    const updatedList = projectList.map((p) => (p.id === projectId ? { ...p, name } : p));
    setProjectList(updatedList);
    await saveProjectIndex(updatedList);
  }

  async function addScheduleProject({ name, tier: projTier, start, target }) {
    const id = `proj_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
    const newList = [...projectList, { id, name }];
    const payload = {
      projectName: name,
      tier: projTier,
      scheduleStart: start,
      scheduleTargetDate: target,
      scheduleItems: makeDefaultScheduleItems(),
      savedAt: new Date().toISOString(),
    };
    // 로컬 상태(프로젝트 목록 + 캐시)를 먼저 동기적으로 갱신해야
    // loadScheduleProjects의 재조회가 빈 데이터로 캐시를 덮어쓰는 경쟁 상태를 피할 수 있다.
    setProjectList(newList);
    setScheduleProjectsCache((prev) => ({ ...prev, [id]: payload }));
    try {
      await saveProjectIndex(newList);
      await saveProjectData(id, payload);
    } catch (err) {
      setSaveError("새 프로젝트 저장에 실패했어요. 잠시 후 다시 시도해주세요.");
    }
  }

  const scheduleProjects = projectList.map((p) => {
    const isCurrent = p.id === currentProjectId;
    const src = isCurrent
      ? { projectName, tier, scheduleStart, scheduleTargetDate, scheduleItems }
      : scheduleProjectsCache[p.id] || {};
    return {
      id: p.id,
      name: src.projectName || p.name || "(이름 없음)",
      tier: src.tier || "",
      start: src.scheduleStart || "",
      target: src.scheduleTargetDate || "",
      items: src.scheduleItems || [],
    };
  });

  function togglePickerItem(id) {
    setPickerSelected((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }
  function openPicker(target) {
    setPickerOpenFor(target);
    setPickerSelected(new Set());
    setPickerCategoryFilter("전체");
  }
  function confirmPickerAdd() {
    const chosen = itemCatalog.filter((it) => pickerSelected.has(it.id));
    if (chosen.length === 0 || !pickerOpenFor) {
      setPickerOpenFor(null);
      return;
    }
    if (pickerOpenFor === "OSE") {
      setOseItems((prev) => [
        ...prev,
        ...chosen.map((it) => ({ id: nextId(), name: it.name, unitPrice: it.unitPrice, actualUnitPrice: 0, installUnitPrice: 0, installActualUnitPrice: 0, qtyPerRoom: 1 })),
      ]);
    } else {
      setFfeItems((prev) => ({
        ...prev,
        [pickerOpenFor]: [
          ...(prev[pickerOpenFor] || []),
          ...chosen.map((it) => ({ id: nextId(), name: it.name, unitPrice: it.unitPrice, actualUnitPrice: 0, installUnitPrice: 0, installActualUnitPrice: 0, qtyPerRoom: 1 })),
        ],
      }));
    }
    setPickerOpenFor(null);
  }

  async function handleCatalogUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    setCatalogError("");
    setCatalogSummary(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const norm = (s) => String(s).replace(/\s+/g, "").trim();

      let headerIdx = -1;
      let colIdx = {};
      let rows = null;
      const scanned = [];

      for (const sheetName of wb.SheetNames) {
        const candidateRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "" });
        scanned.push({ sheetName, preview: candidateRows.slice(0, 3) });
        for (let i = 0; i < Math.min(candidateRows.length, 20); i++) {
          const normedRow = candidateRows[i].map((c) => norm(c));
          const nameIdx = normedRow.findIndex((c) => c.includes("품목명") || c === "품목");
          const priceIdx = normedRow.findIndex((c) => c.includes("단가"));
          if (nameIdx === -1 || priceIdx === -1) continue;
          const catIdx = normedRow.findIndex((c) => c.includes("카테고리") || c.includes("분류"));
          rows = candidateRows;
          headerIdx = i;
          colIdx = { name: nameIdx, price: priceIdx, category: catIdx };
          break;
        }
        if (headerIdx !== -1) break;
      }

      if (headerIdx === -1) {
        const detail = scanned
          .map((s) => `[${s.sheetName}] ` + s.preview.map((r) => r.filter(Boolean).join(" | ")).join(" / "))
          .join("\n");
        setCatalogError(
          `헤더 행("품목명", "단가" 등)을 찾지 못했어요.\n\n실제로 읽은 내용:\n${detail}`
        );
        return;
      }

      const newItems = [];
      for (let i = headerIdx + 1; i < rows.length; i++) {
        const row = rows[i];
        const name = String(row[colIdx.name] || "").trim();
        if (!name) continue;
        const priceRaw = row[colIdx.price];
        const unitPrice = Math.max(0, parseFloat(String(priceRaw).replace(/[^0-9.]/g, "")) || 0);
        const category = colIdx.category !== -1 ? String(row[colIdx.category] || "").trim() || "미분류" : "미분류";
        newItems.push({ id: nextId(), name, unitPrice, category });
      }

      if (newItems.length === 0) {
        setCatalogError("품목 데이터를 찾지 못했어요.");
        return;
      }
      setItemCatalog(newItems);
      saveCatalog(newItems).catch(() => {});
      setCatalogSummary(`품목 ${newItems.length}개를 카탈로그로 불러왔어요. (전 프로젝트 공통으로 공유됩니다)`);
    } catch (err) {
      setCatalogError("파일을 읽는 중 문제가 발생했어요. xlsx 형식인지 확인해주세요.");
    }
    e.target.value = "";
  }

  function floorFromRoomNumber(numStr) {
    const digits = numStr.replace(/\D/g, "");
    if (!digits) return null;
    if (digits.length <= 2) return "1F";
    return `${parseInt(digits.slice(0, digits.length - 2), 10)}F`;
  }

  function parseRoomTypeLabel(label) {
    const afterDash = label.split("-").slice(1).join("-").trim();
    const parenMatch = afterDash.match(/\(([^)]+)\)\s*$/);
    const paren = parenMatch ? parenMatch[1].trim() : "";
    let base = afterDash.replace(/\([^)]+\)\s*$/, "").trim();
    base = base.replace(/^스탠다드\s*/, "").trim();
    return { category: base || "미분류", paren };
  }

  function parseComposition(comp, paren) {
    const text = comp || "";
    const hasQueen = /퀸베드/.test(text);
    const singleMatch = text.match(/싱글베드\s*(\d+)?/);
    let bed = hasQueen ? "퀸" : singleMatch ? "싱글" : "퀸";
    let mattressQty = 1;
    const extraIrregular = [];
    if (singleMatch && !hasQueen) {
      mattressQty = singleMatch[1] ? parseInt(singleMatch[1], 10) : 1;
    } else if (singleMatch && hasQueen) {
      extraIrregular.push("싱글베드 추가");
    }
    let bathtub = "무";
    if (paren === "욕조" || /아크릴\s*욕조/.test(text)) bathtub = "유";
    if (paren === "장애인") {
      bathtub = "무";
      extraIrregular.push("장애인객실");
    }
    return { bed, mattressQty, bathtub, extraIrregular };
  }

  async function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    setImportError("");
    setImportSummary(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });

      const norm = (s) => String(s).replace(/\s+/g, "").trim();
      const HEADER_KEYS = ["룸타입", "수량", "인원", "구성", "호수"];

      let rows = null;
      let headerIdx = -1;
      let colIdx = {};
      const scanned = []; // for diagnostics if nothing matches

      for (const sheetName of wb.SheetNames) {
        const candidateRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "" });
        scanned.push({ sheetName, preview: candidateRows.slice(0, 3) });
        for (let i = 0; i < Math.min(candidateRows.length, 20); i++) {
          const rawRow = candidateRows[i];
          const normedRow = rawRow.map((c) => norm(c));
          // "호수" 컬럼 존재 여부를 기준으로 헤더 행을 판단 (가장 구분력 높은 컬럼)
          const hoIdx = normedRow.findIndex((c) => c.includes("호수"));
          if (hoIdx === -1) continue;
          const guessedCols = {};
          HEADER_KEYS.forEach((key) => {
            const idx = normedRow.findIndex((c) => c.includes(key));
            if (idx !== -1) guessedCols[key] = idx;
          });
          if (guessedCols["룸타입"] === undefined) {
            // "룸타입" 정확히 없으면 "타입"이 들어간 첫 컬럼으로 대체 추정
            const altIdx = normedRow.findIndex((c) => c.includes("타입"));
            if (altIdx !== -1) guessedCols["룸타입"] = altIdx;
          }
          if (guessedCols["룸타입"] === undefined || guessedCols["구성"] === undefined) continue;
          rows = candidateRows;
          headerIdx = i;
          colIdx = guessedCols;
          break;
        }
        if (headerIdx !== -1) break;
      }

      if (headerIdx === -1) {
        const detail = scanned
          .map(
            (s) =>
              `[${s.sheetName}] ` +
              s.preview.map((r) => r.filter(Boolean).join(" | ")).join(" / ")
          )
          .join("\n");
        setImportError(
          `헤더 행("룸타입", "구성", "호수" 등)을 찾지 못했어요. 열 이름에 "호수"와 "구성"이 포함된 헤더 행이 있는지 확인해주세요.\n\n실제로 읽은 내용:\n${detail}`
        );
        return;
      }

      const newRoomTypes = [];
      const newCategories = new Set(categories);
      const newIrregular = new Set(irregularOptions);
      const floorSet = new Set(floors);
      let facilityCount = 0;

      for (let i = headerIdx + 1; i < rows.length; i++) {
        const row = rows[i];
        const label = String(row[colIdx["룸타입"]] || "").trim();
        if (!label) continue;
        if (label === "총계" || label.includes("총 계")) break;
        if (!label.startsWith("ROOM(")) {
          facilityCount++;
          continue;
        }
        const composition = String(row[colIdx["구성"]] || "");
        const hoStr = String(row[colIdx["호수"]] || "");
        const roomNumbers = hoStr
          .split(",")
          .map((s) => s.trim())
          .filter(Boolean);

        const { category, paren } = parseRoomTypeLabel(label);
        const { bed, mattressQty, bathtub, extraIrregular } = parseComposition(composition, paren);

        newCategories.add(category);
        extraIrregular.forEach((o) => newIrregular.add(o));

        const byFloor = {};
        roomNumbers.forEach((num) => {
          const f = floorFromRoomNumber(num);
          if (!f) return;
          floorSet.add(f);
          byFloor[f] = (byFloor[f] || 0) + 1;
        });

        newRoomTypes.push({
          id: nextId(),
          bed,
          bathtub,
          category,
          irregular: extraIrregular,
          mattressQty,
          grade: "Superior",
          features: [],
          view: "",
          includeBedInName: false,
          includeViewInName: false,
          customName: "",
          otaBedCount: "",
          otaBedSize: "",
          otaMaxOccupancy: "",
          otaFacilities: "",
          roomNumbers,
          sourceLabel: label.split("-")[0].trim(),
          byFloor,
        });
      }

      if (newRoomTypes.length === 0) {
        setImportError('"ROOM(...)" 형식의 룸타입 행을 찾지 못했어요.');
        return;
      }

      const finalFloors = sortFloors([...floorSet]);
      setCategories([...newCategories]);
      setIrregularOptions([...newIrregular]);
      setFloors(finalFloors);

      const matchKey = (rt) =>
        [rt.category, rt.bed, rt.bathtub, [...rt.irregular].sort().join(",")].join("|");

      setRoomTypes((prev) => {
        const byKey = new Map(prev.map((rt) => [matchKey(rt), rt]));
        const untouchedIds = new Set(prev.map((rt) => rt.id));
        const merged = [];
        const appended = [];

        newRoomTypes.forEach((incoming) => {
          const normalized = {
            ...incoming,
            byFloor: Object.fromEntries(finalFloors.map((f) => [f, incoming.byFloor[f] || 0])),
          };
          const key = matchKey(incoming);
          const existing = overwriteOnImport ? byKey.get(key) : null;
          if (existing) {
            // 같은 룸타입(카테고리·침대·욕조·이레귤러 동일)이 이미 있으면 새로 덮어씀(중복 누적 방지)
            merged.push({ ...existing, ...normalized, id: existing.id });
            untouchedIds.delete(existing.id);
          } else {
            appended.push(normalized);
          }
        });

        const mergedIds = new Set(merged.map((rt) => rt.id));
        const kept = prev
          .filter((rt) => untouchedIds.has(rt.id) && !mergedIds.has(rt.id))
          .map((rt) => ({ ...rt, byFloor: Object.fromEntries(finalFloors.map((f) => [f, rt.byFloor[f] || 0])) }));

        return [...kept, ...merged, ...appended];
      });

      const totalRooms = newRoomTypes.reduce((s, rt) => s + rt.roomNumbers.length, 0);
      setImportSummary(
        `룸타입 ${newRoomTypes.length}개, 호수 ${totalRooms}개를 가져왔어요.` +
          (facilityCount > 0 ? ` (부대시설성 항목 ${facilityCount}개는 제외됨)` : "") +
          (overwriteOnImport ? " 기존에 같은 룸타입이 있으면 덮어썼어요." : "")
      );
    } catch (err) {
      setImportError("파일을 읽는 중 문제가 발생했어요. xlsx 형식인지 확인해주세요.");
    }
    e.target.value = "";
  }

  const roomTypeTotal = (rt) => Object.values(rt.byFloor).reduce((a, b) => a + b, 0);
  const grandTotal = useMemo(
    () => roomTypes.reduce((sum, rt) => sum + roomTypeTotal(rt), 0),
    [roomTypes]
  );
  const floorTotal = (floor) =>
    roomTypes.reduce((sum, rt) => sum + (rt.byFloor[floor] || 0), 0);

  useEffect(() => {
    if (!floorPlanAutoOffApplied && grandTotal > 100) {
      setShowFloorPlan(false);
      setFloorPlanAutoOffApplied(true);
    }
  }, [grandTotal, floorPlanAutoOffApplied]);

  function codeFor(rt) {
    const bed = rt.bed === "싱글" ? "S" : "Q";
    const bath = rt.bathtub === "유" ? "B" : "NB";
    const cat = rt.category.slice(0, 2);
    return `${cat}-${bed}-${bath}`;
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans">
      <style>{`
        @media print {
          body * { visibility: hidden; }
          #report-print-area, #report-print-area * { visibility: visible; }
          #report-print-area { position: absolute; left: 0; top: 0; width: 100%; box-shadow: none; border: none; }
          .no-print { display: none !important; }
        }
      `}</style>
      <div className="max-w-6xl mx-auto px-6 py-8 space-y-8">
        {/* Save bar */}
        <div className="bg-white border border-slate-200 rounded-xl px-5 py-3 flex items-center justify-between flex-wrap gap-3">
          <div className="flex items-center gap-2 text-xs text-slate-500 flex-wrap">
            <Users size={15} />
            <select
              value={currentProjectId || ""}
              onChange={(e) => switchToProject(e.target.value)}
              disabled={isLoading}
              className="border border-slate-300 rounded-lg px-2 py-1 text-xs bg-white"
            >
              {projectList.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.name || "(이름 없음)"}
                </option>
              ))}
            </select>
            <button
              onClick={createNewProject}
              className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50"
            >
              + 새 프로젝트
            </button>
            <span>
              {isLoading
                ? "불러오는 중..."
                : lastSaved
                ? `마지막 저장: ${new Date(lastSaved).toLocaleString("ko-KR")}`
                : "아직 저장된 데이터가 없어요."}
            </span>
          </div>
          <div className="flex items-center gap-2">
            {saveError && (
              <span className="text-xs text-rose-600 flex items-center gap-1">
                {saveError}
                <button
                  onClick={() => setSaveError("")}
                  className="text-rose-400 hover:text-rose-700"
                  title="닫기"
                >
                  <X size={12} />
                </button>
              </span>
            )}
            <button
              onClick={saveProject}
              disabled={isSaving || isLoading}
              className="text-sm bg-amber-700 text-white px-4 py-1.5 rounded-lg hover:bg-amber-800 disabled:opacity-50 flex items-center gap-1.5"
            >
              {isSaving ? <Loader2 size={14} className="animate-spin" /> : <Save size={14} />}
              {isSaving ? "저장 중..." : "팀 저장소에 저장"}
            </button>
          </div>
        </div>
        {loadNotice && (
          <p className="text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-md px-3 py-2 -mt-4">
            {loadNotice}
          </p>
        )}
        {supabaseSyncError && (
          <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-md px-3 py-2 -mt-4 flex items-center justify-between">
            <span>{supabaseSyncError}</span>
            <button onClick={() => setSupabaseSyncError("")} className="text-amber-500 hover:text-amber-800" title="닫기">
              <X size={12} />
            </button>
          </p>
        )}

        {/* Tab navigation */}
        <div className="flex gap-2 no-print">
          {[
            { key: "main", label: "발주 관리" },
            { key: "schedule", label: "공정표" },
            { key: "test-schedule", label: "테스트 공정표" },
          ].map((t) => (
            <button
              key={t.key}
              onClick={() => setActiveTab(t.key)}
              className={`text-sm px-4 py-2 rounded-lg border ${
                activeTab === t.key
                  ? "bg-amber-700 text-white border-amber-700"
                  : "bg-white text-slate-600 border-slate-300 hover:bg-slate-50"
              }`}
            >
              {t.label}
            </button>
          ))}
        </div>

        {activeTab === "schedule" && (
          <ScheduleTab
            projects={scheduleProjects}
            loading={scheduleTabLoading}
            onUpdateItem={updateScheduleItem}
            onAddItem={addScheduleItem}
            onRemoveItem={removeScheduleItem}
            onRenameProject={renameScheduleProject}
            onAddProject={addScheduleProject}
          />
        )}

        {activeTab === "test-schedule" && <TestScheduleDashboard />}

        {activeTab === "main" && (
          <>
        {/* Header / project */}
        <div className="bg-white border border-slate-200 rounded-xl p-6">
          <div className="flex items-center gap-2 mb-4 text-slate-500">
            <Building2 size={18} />
            <span className="text-sm font-medium tracking-wide">프로젝트 설정</span>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div>
              <label className="block text-xs text-slate-500 mb-1">프로젝트명</label>
              <input
                value={projectName}
                onChange={(e) => setProjectName(e.target.value)}
                placeholder="예: 청담 리브랜딩 프로젝트"
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
              />
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">브랜드 티어</label>
              <select
                value={tier}
                onChange={(e) => setTier(e.target.value)}
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
              >
                {TIERS.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">오픈바이징 배정 예산 (전체 공사비 중)</label>
              <input
                type="number"
                min="0"
                value={totalBudget}
                onChange={(e) => setTotalBudget(Math.max(0, parseFloat(e.target.value || "0") || 0))}
                placeholder="예: 500000000"
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
              />
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">
                브랜드 룸 네임 <span className="text-slate-400">(라이프스타일·콘셉트 호텔만 해당)</span>
              </label>
              <input
                value={brandRoomName}
                onChange={(e) => setBrandRoomName(e.target.value)}
                placeholder="예: 시즈쿠"
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
              />
              <p className="text-[11px] text-slate-400 mt-1">
                입력하면 객실 등급명 앞에 자동으로 붙어요. (브랜드 룸 네이밍 가이드라인 v1.0)
              </p>
            </div>
          </div>
        </div>

        {/* Excel bulk import */}
        <div className="bg-white border border-slate-200 rounded-xl p-6">
          <div className="flex items-center justify-between flex-wrap gap-3">
            <div className="flex items-center gap-2 text-slate-500">
              <Upload size={18} />
              <span className="text-sm font-medium tracking-wide">엑셀로 룸타입 + 호수 일괄 생성</span>
            </div>
            <label className="text-sm bg-amber-700 text-white px-4 py-2 rounded-lg hover:bg-amber-800 cursor-pointer flex items-center gap-1.5">
              <Upload size={15} />
              엑셀 파일 선택
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleExcelUpload} className="hidden" />
            </label>
          </div>
          <p className="text-xs text-slate-400 mt-2">
            "룸타입 / 수량 / 인원 / 구성 / 호수" 열이 있는 시트를 올리면, 룸타입·속성·층별 배치·호수까지 한 번에 채워드려요.
          </p>
          <label className="flex items-center gap-1.5 text-xs text-slate-500 mt-2 cursor-pointer w-fit">
            <input
              type="checkbox"
              checked={overwriteOnImport}
              onChange={(e) => setOverwriteOnImport(e.target.checked)}
            />
            같은 룸타입(카테고리·침대·욕조·이레귤러 동일)을 다시 올리면 새로 추가하지 않고 덮어쓰기
          </label>
          {importSummary && (
            <p className="text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-md px-3 py-2 mt-3">
              {importSummary}
            </p>
          )}
          {importError && (
            <p className="text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-md px-3 py-2 mt-3 whitespace-pre-wrap">
              {importError}
            </p>
          )}
        </div>

        {/* Item catalog upload */}
        <div className="bg-white border border-slate-200 rounded-xl p-6">
          <div className="flex items-center justify-between flex-wrap gap-3">
            <div className="flex items-center gap-2 text-slate-500">
              <Upload size={18} />
              <span className="text-sm font-medium tracking-wide">품목 마스터 카탈로그 업로드</span>
            </div>
            <label className="text-sm bg-amber-700 text-white px-4 py-2 rounded-lg hover:bg-amber-800 cursor-pointer flex items-center gap-1.5">
              <Upload size={15} />
              엑셀 파일 선택
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleCatalogUpload} className="hidden" />
            </label>
          </div>
          <p className="text-xs text-slate-400 mt-2">
            "품목명 / 단가 / 카테고리" 열이 있는 시트를 올리면, FF&E·OS&E 등록할 때 카탈로그에서 골라서 바로 넣을 수 있어요. (이 카탈로그는 프로젝트 구분 없이 전체 공통으로 공유돼요)
          </p>
          {itemCatalog.length > 0 && (
            <p className="text-xs text-slate-500 mt-2">
              현재 카탈로그: <span className="font-medium text-slate-700">{itemCatalog.length}개 품목</span>
            </p>
          )}
          {catalogSummary && (
            <p className="text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-md px-3 py-2 mt-3">
              {catalogSummary}
            </p>
          )}
          {catalogError && (
            <p className="text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-md px-3 py-2 mt-3 whitespace-pre-wrap">
              {catalogError}
            </p>
          )}
        </div>

        {/* Attribute definitions */}
        <div className="bg-white border border-slate-200 rounded-xl p-6">
          <div className="flex items-center gap-2 mb-4 text-slate-500">
            <LayoutGrid size={18} />
            <span className="text-sm font-medium tracking-wide">룸타입 속성 정의</span>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <p className="text-xs text-slate-500 mb-2">룸카테고리 (프로젝트별 커스텀)</p>
              <div className="flex flex-wrap gap-2 mb-2">
                {categories.map((c) => (
                  <span
                    key={c}
                    className={`inline-flex items-center gap-1 ${categoryColor(c).bg} ${categoryColor(c).text} text-xs px-2.5 py-1 rounded-full`}
                  >
                    {c}
                    <button onClick={() => removeCategory(c)} className="hover:opacity-60">
                      <X size={12} />
                    </button>
                  </span>
                ))}
              </div>
              <div className="flex gap-2">
                <input
                  value={newCategory}
                  onChange={(e) => setNewCategory(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && addCategory()}
                  placeholder="새 카테고리 추가"
                  className="flex-1 border border-slate-300 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                />
                <button
                  onClick={addCategory}
                  className="border border-slate-300 rounded-lg px-2.5 hover:bg-slate-50"
                >
                  <Plus size={16} />
                </button>
              </div>
            </div>

            <div>
              <p className="text-xs text-slate-500 mb-2">이레귤러 옵션 (다중 선택 가능)</p>
              <div className="flex flex-wrap gap-2 mb-2">
                {irregularOptions.map((o) => (
                  <span
                    key={o}
                    className="inline-flex items-center gap-1 bg-slate-100 text-slate-700 text-xs px-2.5 py-1 rounded-full"
                  >
                    {o}
                    <button onClick={() => removeIrregular(o)} className="hover:opacity-60">
                      <X size={12} />
                    </button>
                  </span>
                ))}
              </div>
              <div className="flex gap-2">
                <input
                  value={newIrregular}
                  onChange={(e) => setNewIrregular(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && addIrregular()}
                  placeholder="새 옵션 추가"
                  className="flex-1 border border-slate-300 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                />
                <button
                  onClick={addIrregular}
                  className="border border-slate-300 rounded-lg px-2.5 hover:bg-slate-50"
                >
                  <Plus size={16} />
                </button>
              </div>
            </div>

            <div>
              <p className="text-xs text-slate-500 mb-2">객실 콘텐츠 Room Feature (룸 네이밍에 반영)</p>
              <div className="flex flex-wrap gap-2 mb-2">
                {roomFeatures.map((o) => (
                  <span
                    key={o}
                    className="inline-flex items-center gap-1 bg-amber-50 text-amber-800 text-xs px-2.5 py-1 rounded-full"
                  >
                    {o}
                    <button onClick={() => removeRoomFeature(o)} className="hover:opacity-60">
                      <X size={12} />
                    </button>
                  </span>
                ))}
              </div>
              <div className="flex gap-2">
                <input
                  value={newRoomFeature}
                  onChange={(e) => setNewRoomFeature(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && addRoomFeature()}
                  placeholder="예: 스파, 복층"
                  className="flex-1 border border-slate-300 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                />
                <button
                  onClick={addRoomFeature}
                  className="border border-slate-300 rounded-lg px-2.5 hover:bg-slate-50"
                >
                  <Plus size={16} />
                </button>
              </div>
            </div>

            <div>
              <p className="text-xs text-slate-500 mb-2">전망 View Type (선택 시에만 룸 네이밍에 반영)</p>
              <div className="flex flex-wrap gap-2 mb-2">
                {viewTypes.map((o) => (
                  <span
                    key={o}
                    className="inline-flex items-center gap-1 bg-sky-50 text-sky-800 text-xs px-2.5 py-1 rounded-full"
                  >
                    {o}
                    <button onClick={() => removeViewType(o)} className="hover:opacity-60">
                      <X size={12} />
                    </button>
                  </span>
                ))}
              </div>
              <div className="flex gap-2">
                <input
                  value={newViewType}
                  onChange={(e) => setNewViewType(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && addViewType()}
                  placeholder="예: 오션뷰"
                  className="flex-1 border border-slate-300 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                />
                <button
                  onClick={addViewType}
                  className="border border-slate-300 rounded-lg px-2.5 hover:bg-slate-50"
                >
                  <Plus size={16} />
                </button>
              </div>
            </div>
          </div>

          <div className="h-px bg-slate-200 my-5" />

          {/* Room type composer */}
          <p className="text-xs text-slate-500 mb-3">룸타입 조합 만들기</p>
          <div className="flex flex-wrap items-end gap-4">
            <div>
              <label className="block text-xs text-slate-500 mb-1">침대타입</label>
              <select
                value={draft.bed}
                onChange={(e) => setDraft({ ...draft, bed: e.target.value })}
                className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm"
              >
                {BED_TYPES.map((b) => (
                  <option key={b}>{b}</option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">욕조유무</label>
              <select
                value={draft.bathtub}
                onChange={(e) => setDraft({ ...draft, bathtub: e.target.value })}
                className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm"
              >
                {BATHTUB.map((b) => (
                  <option key={b}>{b}</option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">룸카테고리</label>
              <select
                value={draft.category}
                onChange={(e) => setDraft({ ...draft, category: e.target.value })}
                className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm"
              >
                {categories.map((c) => (
                  <option key={c}>{c}</option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">객실 등급 Core Room Grade</label>
              <select
                value={draft.grade}
                onChange={(e) => setDraft({ ...draft, grade: e.target.value })}
                className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm"
              >
                {ROOM_GRADES.map((g) => (
                  <option key={g}>{g}</option>
                ))}
              </select>
              {draft.grade === "Standard" && (
                <p className="text-[11px] text-amber-600 mt-1 max-w-[160px]">
                  면적이 더 작은 타입에 한해서만 사용
                </p>
              )}
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">매트리스 수량</label>
              <input
                type="number"
                min="1"
                value={draft.mattressQty}
                onChange={(e) =>
                  setDraft({ ...draft, mattressQty: Math.max(1, parseInt(e.target.value || "1", 10)) })
                }
                className="w-20 border border-slate-300 rounded-lg px-3 py-1.5 text-sm"
              />
            </div>
            <div className="flex-1 min-w-[180px]">
              <label className="block text-xs text-slate-500 mb-1">이레귤러 옵션</label>
              <div className="flex flex-wrap gap-1.5">
                {irregularOptions.map((o) => (
                  <button
                    key={o}
                    onClick={() => toggleDraftIrregular(o)}
                    className={`text-xs px-2.5 py-1 rounded-full border ${
                      draft.irregular.includes(o)
                        ? "bg-amber-600 text-white border-amber-600"
                        : "border-slate-300 text-slate-600 hover:bg-slate-50"
                    }`}
                  >
                    {o}
                  </button>
                ))}
              </div>
            </div>
          </div>

          <div className="h-px bg-slate-200 my-4" />

          {/* Room naming — 브랜드 룸 네이밍 가이드라인 v1.0 */}
          <p className="text-xs text-slate-500 mb-3">룸 네이밍 (브랜드 룸 네이밍 가이드라인 v1.0)</p>
          <div className="flex flex-wrap items-end gap-4">
            <div className="min-w-[180px]">
              <label className="block text-xs text-slate-500 mb-1">객실 콘텐츠 Room Feature</label>
              <div className="flex flex-wrap gap-1.5">
                {roomFeatures.map((o) => (
                  <button
                    key={o}
                    onClick={() => toggleDraftFeature(o)}
                    className={`text-xs px-2.5 py-1 rounded-full border ${
                      draft.features.includes(o)
                        ? "bg-amber-600 text-white border-amber-600"
                        : "border-slate-300 text-slate-600 hover:bg-slate-50"
                    }`}
                  >
                    {o}
                  </button>
                ))}
              </div>
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">전망 View Type</label>
              <select
                value={draft.view}
                onChange={(e) => setDraft({ ...draft, view: e.target.value })}
                className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm"
              >
                <option value="">미지정</option>
                {viewTypes.map((v) => (
                  <option key={v}>{v}</option>
                ))}
              </select>
            </div>
            <div className="flex items-center gap-3 pb-1.5">
              <label className="flex items-center gap-1.5 text-xs text-slate-500 cursor-pointer">
                <input
                  type="checkbox"
                  checked={draft.includeBedInName}
                  onChange={(e) => setDraft({ ...draft, includeBedInName: e.target.checked })}
                />
                침대구성을 이름에 포함
              </label>
              <label className="flex items-center gap-1.5 text-xs text-slate-500 cursor-pointer">
                <input
                  type="checkbox"
                  checked={draft.includeViewInName}
                  onChange={(e) => setDraft({ ...draft, includeViewInName: e.target.checked })}
                  disabled={!draft.view}
                />
                전망을 이름에 포함
              </label>
            </div>
            <div className="flex-1 min-w-[220px]">
              <label className="block text-xs text-slate-500 mb-1">
                커스텀 명칭 (선택 — 비워두면 자동 생성)
              </label>
              <input
                value={draft.customName}
                onChange={(e) => setDraft({ ...draft, customName: e.target.value })}
                placeholder={generateRoomName(draft)}
                className="w-full border border-slate-300 rounded-lg px-3 py-1.5 text-sm"
              />
            </div>
          </div>
          <p className="text-xs text-slate-500 mt-3">
            생성될 객실명: <span className="font-medium text-amber-700">{generateRoomName(draft)}</span>
          </p>

          <div className="mt-4 text-right">
            <button
              onClick={addRoomType}
              className="bg-amber-700 text-white text-sm px-4 py-2 rounded-lg hover:bg-amber-800 inline-flex items-center gap-1.5"
            >
              <Plus size={16} /> 룸타입 추가
            </button>
          </div>
        </div>

        {/* Floor allocation + floor plan */}
        {roomTypes.length > 0 && (
          <div className="bg-white border border-slate-200 rounded-xl p-6">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2 text-slate-500">
                <Table2 size={18} />
                <span className="text-sm font-medium tracking-wide">층별 배치</span>
              </div>
              <div className="flex flex-wrap gap-2 items-center">
                <input
                  value={newFloor}
                  onChange={(e) => setNewFloor(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && addFloor()}
                  placeholder="층 추가 (예: B1, PH)"
                  className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm w-32 focus:outline-none focus:ring-2 focus:ring-amber-500"
                />
                <button
                  onClick={addFloor}
                  className="border border-slate-300 rounded-lg px-2.5 py-1.5 hover:bg-slate-50"
                >
                  <Plus size={16} />
                </button>
                <span className="text-xs text-slate-300">|</span>
                <input
                  type="number"
                  value={bulkFrom}
                  onChange={(e) => setBulkFrom(e.target.value)}
                  placeholder="시작층"
                  className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm w-20 focus:outline-none focus:ring-2 focus:ring-amber-500"
                />
                <span className="text-xs text-slate-400">~</span>
                <input
                  type="number"
                  value={bulkTo}
                  onChange={(e) => setBulkTo(e.target.value)}
                  placeholder="끝층"
                  className="border border-slate-300 rounded-lg px-3 py-1.5 text-sm w-20 focus:outline-none focus:ring-2 focus:ring-amber-500"
                />
                <button
                  onClick={addFloorRange}
                  className="border border-slate-300 rounded-lg px-3 py-1.5 text-xs hover:bg-slate-50 whitespace-nowrap"
                >
                  일괄 추가 (nF)
                </button>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="min-w-max text-sm border-collapse">
                <thead>
                  <tr>
                    <th className="text-left text-xs text-slate-500 font-normal pb-2 pr-3 sticky left-0 bg-white min-w-[260px]">
                      룸타입
                    </th>
                    {floors.map((f) => (
                      <th key={f} className="text-center text-xs text-slate-500 font-normal pb-2 px-2 min-w-[64px]">
                        <div className="flex items-center justify-center gap-1 whitespace-nowrap">
                          {f}
                          <button onClick={() => removeFloor(f)} className="hover:text-slate-700">
                            <X size={11} />
                          </button>
                        </div>
                      </th>
                    ))}
                    <th className="text-center text-xs text-slate-500 font-normal pb-2 pl-2 min-w-[56px]">합계</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {roomTypes.map((rt) => {
                    const c = categoryColor(rt.category);
                    return (
                      <tr key={rt.id} className="border-t border-slate-100">
                        <td className="py-2 pr-3 sticky left-0 bg-white min-w-[260px]">
                          <div className="mb-1">
                            <input
                              value={rt.customName || ""}
                              onChange={(e) => updateRoomTypeField(rt.id, "customName", e.target.value)}
                              placeholder={generateRoomName(rt)}
                              title="객실명 (비워두면 자동 생성명 사용)"
                              className="text-sm font-medium text-slate-800 border-b border-dashed border-slate-300 focus:outline-none focus:border-amber-500 bg-transparent w-full max-w-[240px]"
                            />
                          </div>
                          <div className="flex items-center gap-2 flex-wrap">
                            <span className={`text-[11px] font-medium px-2 py-0.5 rounded whitespace-nowrap ${c.bg} ${c.text}`}>
                              {codeFor(rt)}
                            </span>
                            <span className="text-xs text-slate-500 whitespace-nowrap">
                              {rt.category} · {rt.bed} · 욕조{rt.bathtub}
                              {rt.irregular.length > 0 ? ` · ${rt.irregular.join(", ")}` : ""}
                            </span>
                            <span className="flex items-center gap-1 text-xs text-slate-500 whitespace-nowrap">
                              매트리스×
                              <input
                                type="number"
                                min="1"
                                value={rt.mattressQty}
                                onChange={(e) => setMattressQty(rt.id, e.target.value)}
                                className="w-12 text-center border border-slate-200 rounded-md py-0.5 text-xs focus:outline-none focus:ring-2 focus:ring-amber-500"
                              />
                            </span>
                          </div>
                          {rt.roomNumbers && rt.roomNumbers.length > 0 && (
                            <details className="mt-1">
                              <summary className="text-[11px] text-amber-700 cursor-pointer select-none">
                                호수 보기 ({rt.roomNumbers.length}개)
                              </summary>
                              <p className="text-[11px] text-slate-500 mt-1 leading-relaxed max-w-[240px]">
                                {rt.roomNumbers.join(", ")}
                              </p>
                            </details>
                          )}
                        </td>
                        {floors.map((f) => (
                          <td key={f} className="text-center px-2 min-w-[64px]">
                            <input
                              type="number"
                              min="0"
                              value={rt.byFloor[f] || 0}
                              onChange={(e) => setQty(rt.id, f, e.target.value)}
                              className="w-14 text-center border border-slate-200 rounded-md py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                            />
                          </td>
                        ))}
                        <td className="text-center font-medium pl-2 min-w-[56px]">{roomTypeTotal(rt)}</td>
                        <td className="pl-2">
                          <button
                            onClick={() => removeRoomType(rt.id)}
                            className="text-slate-400 hover:text-rose-600"
                          >
                            <Trash2 size={14} />
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                  <tr className="border-t border-slate-200">
                    <td className="py-2 pr-3 text-xs text-slate-500 sticky left-0 bg-white">층별 합계</td>
                    {floors.map((f) => (
                      <td key={f} className="text-center text-xs text-slate-500">
                        {floorTotal(f)}
                      </td>
                    ))}
                    <td className="text-center text-sm font-semibold">{grandTotal}</td>
                    <td></td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Floor plan visual: cell grid per floor */}
            <div className="mt-6">
              <div className="flex items-center justify-between mb-2">
                <p className="text-xs text-slate-500">배치도 (층별 객실 셀)</p>
                <label className="flex items-center gap-1.5 text-xs text-slate-500 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={showFloorPlan}
                    onChange={(e) => setShowFloorPlan(e.target.checked)}
                  />
                  배치도 표시
                  {grandTotal > 100 && <span className="text-amber-600">(100실 초과 — 무거울 수 있어요)</span>}
                </label>
              </div>
              {showFloorPlan ? (
                <div className="space-y-2">
                  {floors.map((f) => {
                    const total = floorTotal(f);
                    const cells = [];
                    roomTypes.forEach((rt) => {
                      const qty = rt.byFloor[f] || 0;
                      for (let i = 0; i < qty; i++) cells.push(rt);
                    });
                    return (
                      <div key={f} className="flex items-start gap-3">
                        <span className="w-10 pt-1 text-xs text-slate-500 shrink-0">{f}</span>
                        <div className="flex-1 flex flex-wrap gap-1 bg-slate-50 border border-slate-200 rounded-md p-1.5 min-h-[34px]">
                          {total === 0 ? (
                            <span className="text-[11px] text-slate-400 px-1">미배치</span>
                          ) : (
                            cells.map((rt, i) => {
                              const c = categoryColor(rt.category);
                              return (
                                <div
                                  key={`${rt.id}-${i}`}
                                  title={`${codeFor(rt)} · ${generateRoomName(rt)}`}
                                  className={`w-6 h-6 rounded ${c.bar} flex items-center justify-center`}
                                />
                              );
                            })
                          )}
                        </div>
                        <span className="w-10 pt-1 text-xs text-slate-500 text-right shrink-0">{total}실</span>
                      </div>
                    );
                  })}
                  <div className="flex flex-wrap gap-3 pt-1">
                    {roomTypes.map((rt) => {
                      const c = categoryColor(rt.category);
                      return (
                        <span key={rt.id} className="flex items-center gap-1.5 text-[11px] text-slate-500">
                          <span className={`w-3 h-3 rounded ${c.bar}`} />
                          {codeFor(rt)}
                        </span>
                      );
                    })}
                  </div>
                </div>
              ) : (
                <p className="text-xs text-slate-400 bg-slate-50 border border-slate-200 rounded-md p-3">
                  배치도가 꺼져 있어요. 위 체크박스로 다시 켤 수 있어요.
                </p>
              )}
            </div>

          </div>
        )}

        {/* Summary table */}
        {roomTypes.length > 0 && (
          <div className="bg-white border border-slate-200 rounded-xl p-6">
            <div className="flex items-center justify-between mb-4">
              <span className="text-sm font-medium tracking-wide text-slate-500">
                요약 — {projectName || "프로젝트명 미입력"} ({tier})
              </span>
              <span className="text-sm text-slate-500">
                총 <span className="font-semibold text-slate-800">{grandTotal}</span>실
              </span>
            </div>
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-xs text-slate-500 border-b border-slate-200">
                  <th className="py-2 font-normal">코드</th>
                  <th className="py-2 font-normal">객실명</th>
                  <th className="py-2 font-normal">등급</th>
                  <th className="py-2 font-normal">룸카테고리</th>
                  <th className="py-2 font-normal">침대타입</th>
                  <th className="py-2 font-normal">욕조</th>
                  <th className="py-2 font-normal">이레귤러 옵션</th>
                  <th className="py-2 font-normal text-right">객실 수</th>
                  <th className="py-2 font-normal text-right">매트리스/실</th>
                  <th className="py-2 font-normal text-right">총 매트리스 수</th>
                  <th className="py-2 font-normal text-right">비중</th>
                </tr>
              </thead>
              <tbody>
                {roomTypes.map((rt) => {
                  const c = categoryColor(rt.category);
                  const total = roomTypeTotal(rt);
                  return (
                    <tr key={rt.id} className="border-b border-slate-100">
                      <td className="py-2">
                        <span className={`text-[11px] font-medium px-2 py-0.5 rounded ${c.bg} ${c.text}`}>
                          {codeFor(rt)}
                        </span>
                      </td>
                      <td className="py-2 font-medium text-slate-800">{generateRoomName(rt)}</td>
                      <td className="py-2 text-slate-500">{rt.grade || "Superior"}</td>
                      <td className="py-2">{rt.category}</td>
                      <td className="py-2">{rt.bed}</td>
                      <td className="py-2">{rt.bathtub}</td>
                      <td className="py-2 text-slate-500">
                        {rt.irregular.length > 0 ? rt.irregular.join(", ") : "—"}
                      </td>
                      <td className="py-2 text-right">{total}</td>
                      <td className="py-2 text-right text-slate-500">{rt.mattressQty}</td>
                      <td className="py-2 text-right">{total * rt.mattressQty}</td>
                      <td className="py-2 text-right text-slate-500">
                        {grandTotal > 0 ? `${Math.round((total / grandTotal) * 100)}%` : "0%"}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {/* OTA 상세정보 — 객실명과 분리된 구조화 정보 (브랜드 룸 네이밍 가이드라인 v1.0) */}
        {roomTypes.length > 0 && (
          <div className="bg-white border border-slate-200 rounded-xl p-6">
            <div className="flex items-center gap-2 mb-1 text-slate-500">
              <Table2 size={18} />
              <span className="text-sm font-medium tracking-wide">OTA 상세정보</span>
            </div>
            <p className="text-xs text-slate-400 mb-4">
              객실명에는 넣지 않고, OTA·홈페이지 상세 정보란에 별도로 표기할 항목이에요.
            </p>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-xs text-slate-500 border-b border-slate-200">
                    <th className="py-2 font-normal">객실명</th>
                    <th className="py-2 font-normal">침대 수량</th>
                    <th className="py-2 font-normal">침대 규격</th>
                    <th className="py-2 font-normal">최대 투숙인원</th>
                    <th className="py-2 font-normal">주요 시설</th>
                  </tr>
                </thead>
                <tbody>
                  {roomTypes.map((rt) => (
                    <tr key={rt.id} className="border-b border-slate-100">
                      <td className="py-2 font-medium text-slate-800 whitespace-nowrap">{generateRoomName(rt)}</td>
                      <td className="py-2">
                        <input
                          value={rt.otaBedCount || ""}
                          onChange={(e) => updateRoomTypeField(rt.id, "otaBedCount", e.target.value)}
                          placeholder="예: 퀸 1개"
                          className="w-24 border border-slate-200 rounded-md px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-amber-500"
                        />
                      </td>
                      <td className="py-2">
                        <input
                          value={rt.otaBedSize || ""}
                          onChange={(e) => updateRoomTypeField(rt.id, "otaBedSize", e.target.value)}
                          placeholder="예: 160x200cm"
                          className="w-28 border border-slate-200 rounded-md px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-amber-500"
                        />
                      </td>
                      <td className="py-2">
                        <input
                          value={rt.otaMaxOccupancy || ""}
                          onChange={(e) => updateRoomTypeField(rt.id, "otaMaxOccupancy", e.target.value)}
                          placeholder="예: 기준2/최대3"
                          className="w-28 border border-slate-200 rounded-md px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-amber-500"
                        />
                      </td>
                      <td className="py-2">
                        <input
                          value={rt.otaFacilities || ""}
                          onChange={(e) => updateRoomTypeField(rt.id, "otaFacilities", e.target.value)}
                          placeholder="예: 스파욕조, 테라스, 반신욕조"
                          className="w-full min-w-[220px] border border-slate-200 rounded-md px-2 py-1 text-xs focus:outline-none focus:ring-2 focus:ring-amber-500"
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* FF&E 발주 품목 (룸타입별) */}
        {roomTypes.length > 0 && (
          <div className="bg-white border border-slate-200 rounded-xl p-6">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2 text-slate-500">
                <LayoutGrid size={18} />
                <span className="text-sm font-medium tracking-wide">FF&E 발주 품목 (룸타입별)</span>
              </div>
              <button
                onClick={() => setPresetEditorOpen((v) => !v)}
                className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50"
              >
                기본 세트 편집 ({basicPreset.length}개)
              </button>
            </div>
            {presetEditorOpen && (
              <div className="mb-5 bg-slate-50 border border-slate-200 rounded-lg p-4">
                <p className="text-[11px] text-slate-500 mb-2">
                  "기본 세트 추가" 버튼을 눌렀을 때 룸타입에 깔리는 품목 목록이에요. 여기서 편집하면 다음부터 바로 반영됩니다.
                </p>
                <div className="flex flex-wrap gap-1.5 mb-2">
                  {basicPreset.map((it) => (
                    <span
                      key={it.id}
                      className="inline-flex items-center gap-1 bg-white border border-slate-300 text-xs px-2 py-1 rounded-full"
                    >
                      <input
                        value={it.name}
                        onChange={(e) => updatePresetItemName(it.id, e.target.value)}
                        className="w-24 text-xs focus:outline-none"
                      />
                      <button onClick={() => removePresetItem(it.id)} className="text-slate-400 hover:text-rose-600">
                        <X size={11} />
                      </button>
                    </span>
                  ))}
                </div>
                <div className="flex gap-2">
                  <input
                    value={newPresetItem}
                    onChange={(e) => setNewPresetItem(e.target.value)}
                    onKeyDown={(e) => e.key === "Enter" && addPresetItem()}
                    placeholder="새 품목 이름 추가"
                    className="flex-1 border border-slate-300 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                  />
                  <button
                    onClick={addPresetItem}
                    className="border border-slate-300 rounded-lg px-2.5 hover:bg-white"
                  >
                    <Plus size={16} />
                  </button>
                </div>
              </div>
            )}
            <div className="space-y-6">
              {roomTypes.map((rt) => {
                const c = categoryColor(rt.category);
                const items = ffeItems[rt.id] || [];
                const roomCount = roomTypeTotal(rt);
                const typeTotal = items.reduce(
                  (sum, it) => sum + (it.unitPrice + (it.installUnitPrice || 0)) * it.qtyPerRoom * roomCount,
                  0
                );
                const typeActualTotal = items.reduce(
                  (sum, it) => sum + ((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * it.qtyPerRoom * roomCount,
                  0
                );
                return (
                  <div key={rt.id}>
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-2">
                        <span className={`text-[11px] font-medium px-2 py-0.5 rounded ${c.bg} ${c.text}`}>
                          {codeFor(rt)}
                        </span>
                        <span className="text-xs text-slate-500">
                          {generateRoomName(rt)} · 객실 {roomCount}실
                        </span>
                      </div>
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <button
                          onClick={() => addPresetFfeItems(rt.id)}
                          className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50"
                        >
                          기본 세트 추가
                        </button>
                        {roomTypes.length > 1 && (
                          <>
                            <select
                              value={copySourceFor[rt.id] || ""}
                              onChange={(e) =>
                                setCopySourceFor((prev) => ({ ...prev, [rt.id]: e.target.value }))
                              }
                              className="text-xs border border-slate-300 rounded-lg px-2 py-1"
                            >
                              <option value="">다른 룸타입에서 복사...</option>
                              {roomTypes
                                .filter((r) => r.id !== rt.id)
                                .map((r) => (
                                  <option key={r.id} value={r.id}>
                                    {codeFor(r)} ({(ffeItems[r.id] || []).length}개)
                                  </option>
                                ))}
                            </select>
                            <button
                              onClick={() => copySourceFor[rt.id] && copyFfeItems(copySourceFor[rt.id], rt.id)}
                              disabled={!copySourceFor[rt.id]}
                              className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50 disabled:opacity-40"
                            >
                              복사
                            </button>
                          </>
                        )}
                        <button
                          onClick={() => setPasteOpenFor(pasteOpenFor === rt.id ? null : rt.id)}
                          className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50"
                        >
                          붙여넣기로 추가
                        </button>
                        <button
                          onClick={() => openPicker(rt.id)}
                          className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50"
                        >
                          카탈로그에서 선택
                        </button>
                        <button
                          onClick={() => addFfeItem(rt.id)}
                          className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50 flex items-center gap-1"
                        >
                          <Plus size={13} /> 품목 추가
                        </button>
                      </div>
                    </div>
                    {pickerOpenFor === rt.id && (
                      <CatalogPickerPanel
                        items={itemCatalog}
                        selected={pickerSelected}
                        onToggle={togglePickerItem}
                        onConfirm={confirmPickerAdd}
                        onCancel={() => setPickerOpenFor(null)}
                        categoryFilter={pickerCategoryFilter}
                        setCategoryFilter={setPickerCategoryFilter}
                      />
                    )}
                    {pasteOpenFor === rt.id && (
                      <div className="mb-3 bg-slate-50 border border-slate-200 rounded-lg p-3">
                        <p className="text-[11px] text-slate-500 mb-1.5">
                          엑셀에서 "품목명 [탭] 단가 [탭] 실당수량" 형태로 복사해 붙여넣으세요 (단가·수량 생략 가능, 한 줄에 하나씩)
                        </p>
                        <textarea
                          value={pasteText}
                          onChange={(e) => setPasteText(e.target.value)}
                          rows={4}
                          placeholder={"헤드보드\t120000\t1\n협탁\t80000\t2"}
                          className="w-full border border-slate-300 rounded-md px-2 py-1.5 text-xs font-mono focus:outline-none focus:ring-2 focus:ring-amber-500"
                        />
                        <div className="flex justify-end gap-2 mt-1.5">
                          <button
                            onClick={() => {
                              bulkAddFfeItems(rt.id, pasteText);
                              setPasteText("");
                              setPasteOpenFor(null);
                            }}
                            className="text-xs bg-amber-700 text-white px-3 py-1 rounded-md hover:bg-amber-800"
                          >
                            일괄 추가
                          </button>
                        </div>
                      </div>
                    )}
                    {items.length > 0 && (
                      <table className="w-full text-sm mb-1">
                        <thead>
                          <tr className="text-left text-xs text-slate-500 border-b border-slate-200">
                            <th className="py-1.5 font-normal">품목명</th>
                            <th className="py-1.5 font-normal text-right">공급예산단가</th>
                            <th className="py-1.5 font-normal text-right">공급집행단가</th>
                            <th className="py-1.5 font-normal text-right">설치예산단가</th>
                            <th className="py-1.5 font-normal text-right">설치집행단가</th>
                            <th className="py-1.5 font-normal text-right">실당 수량</th>
                            <th className="py-1.5 font-normal text-right">필요 수량</th>
                            <th className="py-1.5 font-normal text-right">예산금액</th>
                            <th className="py-1.5 font-normal text-right">집행금액</th>
                            <th></th>
                          </tr>
                        </thead>
                        <tbody>
                          {items.map((it) => (
                            <tr key={it.id} className="border-b border-slate-100">
                              <td className="py-1.5">
                                <input
                                  value={it.name}
                                  onChange={(e) => updateFfeItem(rt.id, it.id, "name", e.target.value)}
                                  placeholder="예: 퀸 매트리스"
                                  className="w-full border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                                />
                              </td>
                              <td className="py-1.5">
                                <input
                                  type="number"
                                  min="0"
                                  value={it.unitPrice}
                                  onChange={(e) => updateFfeItem(rt.id, it.id, "unitPrice", e.target.value)}
                                  className="w-24 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                                />
                              </td>
                              <td className="py-1.5">
                                <input
                                  type="number"
                                  min="0"
                                  value={it.actualUnitPrice || 0}
                                  onChange={(e) => updateFfeItem(rt.id, it.id, "actualUnitPrice", e.target.value)}
                                  className="w-24 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 bg-teal-50/40"
                                />
                              </td>
                              <td className="py-1.5">
                                <input
                                  type="number"
                                  min="0"
                                  value={it.installUnitPrice || 0}
                                  onChange={(e) => updateFfeItem(rt.id, it.id, "installUnitPrice", e.target.value)}
                                  className="w-24 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500 bg-amber-50/30"
                                />
                              </td>
                              <td className="py-1.5">
                                <input
                                  type="number"
                                  min="0"
                                  value={it.installActualUnitPrice || 0}
                                  onChange={(e) => updateFfeItem(rt.id, it.id, "installActualUnitPrice", e.target.value)}
                                  className="w-24 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 bg-teal-50/40"
                                />
                              </td>
                              <td className="py-1.5">
                                <input
                                  type="number"
                                  min="0"
                                  step="0.1"
                                  value={it.qtyPerRoom}
                                  onChange={(e) => updateFfeItem(rt.id, it.id, "qtyPerRoom", e.target.value)}
                                  className="w-16 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                                />
                              </td>
                              <td className="py-1.5 text-right text-slate-500">
                                {(it.qtyPerRoom * roomCount).toLocaleString("ko-KR")}
                              </td>
                              <td className="py-1.5 text-right font-medium">
                                {won((it.unitPrice + (it.installUnitPrice || 0)) * it.qtyPerRoom * roomCount)}
                              </td>
                              <td className="py-1.5 text-right font-medium text-teal-700">
                                {won(((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * it.qtyPerRoom * roomCount)}
                              </td>
                              <td className="py-1.5 pl-2">
                                <button
                                  onClick={() => removeFfeItem(rt.id, it.id)}
                                  className="text-slate-400 hover:text-rose-600"
                                >
                                  <Trash2 size={14} />
                                </button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                    {items.length > 0 && (
                      <div className="text-right text-xs text-slate-500">
                        {codeFor(rt)} 소계 — 예산 <span className="font-semibold text-slate-800">{won(typeTotal)}</span>
                        {" / "}집행 <span className="font-semibold text-teal-700">{won(typeActualTotal)}</span>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* OS&E 발주 품목 (공통 리스트) */}
        <div className="bg-white border border-slate-200 rounded-xl p-6">
          <div className="flex items-center justify-between mb-4">
            <div className="flex items-center gap-2 text-slate-500">
              <LayoutGrid size={18} />
              <span className="text-sm font-medium tracking-wide">OS&E 발주 품목 (공통, 전체 {grandTotal}실 기준)</span>
            </div>
            <div className="flex items-center gap-1.5">
              <button
                onClick={() => openPicker("OSE")}
                className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50"
              >
                카탈로그에서 선택
              </button>
              <button
                onClick={() => setPasteOpenFor(pasteOpenFor === "OSE" ? null : "OSE")}
                className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50"
              >
                붙여넣기로 추가
              </button>
              <button
                onClick={addOseItem}
                className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50 flex items-center gap-1"
              >
                <Plus size={13} /> 품목 추가
              </button>
            </div>
          </div>
          {pickerOpenFor === "OSE" && (
            <CatalogPickerPanel
              items={itemCatalog}
              selected={pickerSelected}
              onToggle={togglePickerItem}
              onConfirm={confirmPickerAdd}
              onCancel={() => setPickerOpenFor(null)}
              categoryFilter={pickerCategoryFilter}
              setCategoryFilter={setPickerCategoryFilter}
            />
          )}
          {pasteOpenFor === "OSE" && (
            <div className="mb-3 bg-slate-50 border border-slate-200 rounded-lg p-3">
              <p className="text-[11px] text-slate-500 mb-1.5">
                엑셀에서 "품목명 [탭] 단가 [탭] 실당수량" 형태로 복사해 붙여넣으세요 (단가·수량 생략 가능, 한 줄에 하나씩)
              </p>
              <textarea
                value={pasteText}
                onChange={(e) => setPasteText(e.target.value)}
                rows={4}
                placeholder={"타월 세트\t18000\t2\n어메니티\t3000\t1"}
                className="w-full border border-slate-300 rounded-md px-2 py-1.5 text-xs font-mono focus:outline-none focus:ring-2 focus:ring-amber-500"
              />
              <div className="flex justify-end gap-2 mt-1.5">
                <button
                  onClick={() => {
                    bulkAddOseItems(pasteText);
                    setPasteText("");
                    setPasteOpenFor(null);
                  }}
                  className="text-xs bg-amber-700 text-white px-3 py-1 rounded-md hover:bg-amber-800"
                >
                  일괄 추가
                </button>
              </div>
            </div>
          )}
          {oseItems.length > 0 ? (
            <table className="w-full text-sm">
              <thead>
                <tr className="text-left text-xs text-slate-500 border-b border-slate-200">
                  <th className="py-1.5 font-normal">품목명</th>
                  <th className="py-1.5 font-normal text-right">공급예산단가</th>
                  <th className="py-1.5 font-normal text-right">공급집행단가</th>
                  <th className="py-1.5 font-normal text-right">설치예산단가</th>
                  <th className="py-1.5 font-normal text-right">설치집행단가</th>
                  <th className="py-1.5 font-normal text-right">실당 수량</th>
                  <th className="py-1.5 font-normal text-right">필요 수량</th>
                  <th className="py-1.5 font-normal text-right">예산금액</th>
                  <th className="py-1.5 font-normal text-right">집행금액</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {oseItems.map((it) => (
                  <tr key={it.id} className="border-b border-slate-100">
                    <td className="py-1.5">
                      <input
                        value={it.name}
                        onChange={(e) => updateOseItem(it.id, "name", e.target.value)}
                        placeholder="예: 객실 타월 세트"
                        className="w-full border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        value={it.unitPrice}
                        onChange={(e) => updateOseItem(it.id, "unitPrice", e.target.value)}
                        className="w-24 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        value={it.actualUnitPrice || 0}
                        onChange={(e) => updateOseItem(it.id, "actualUnitPrice", e.target.value)}
                        className="w-24 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 bg-teal-50/40"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        value={it.installUnitPrice || 0}
                        onChange={(e) => updateOseItem(it.id, "installUnitPrice", e.target.value)}
                        className="w-24 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500 bg-amber-50/30"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        value={it.installActualUnitPrice || 0}
                        onChange={(e) => updateOseItem(it.id, "installActualUnitPrice", e.target.value)}
                        className="w-24 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 bg-teal-50/40"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        step="0.1"
                        value={it.qtyPerRoom}
                        onChange={(e) => updateOseItem(it.id, "qtyPerRoom", e.target.value)}
                        className="w-16 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                      />
                    </td>
                    <td className="py-1.5 text-right text-slate-500">
                      {(it.qtyPerRoom * grandTotal).toLocaleString("ko-KR")}
                    </td>
                    <td className="py-1.5 text-right font-medium">
                      {won((it.unitPrice + (it.installUnitPrice || 0)) * it.qtyPerRoom * grandTotal)}
                    </td>
                    <td className="py-1.5 text-right font-medium text-teal-700">
                      {won(((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * it.qtyPerRoom * grandTotal)}
                    </td>
                    <td className="py-1.5 pl-2">
                      <button onClick={() => removeOseItem(it.id)} className="text-slate-400 hover:text-rose-600">
                        <Trash2 size={14} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          ) : (
            <p className="text-xs text-slate-400">품목을 추가하면 전체 객실 수 기준으로 자동 계산됩니다.</p>
          )}
        </div>

        {/* 현장지출 / 인건비 / 예산외 지출 */}
        <ExpenseSection
          title="현장지출"
          items={siteExpenses}
          onAdd={siteExpenseHandlers.add}
          onUpdate={siteExpenseHandlers.update}
          onRemove={siteExpenseHandlers.remove}
        />
        <ExpenseSection
          title="인건비 지출"
          items={laborExpenses}
          onAdd={laborExpenseHandlers.add}
          onUpdate={laborExpenseHandlers.update}
          onRemove={laborExpenseHandlers.remove}
        />
        <ExpenseSection
          title="예산외 지출"
          items={extraExpenses}
          onAdd={extraExpenseHandlers.add}
          onUpdate={extraExpenseHandlers.update}
          onRemove={extraExpenseHandlers.remove}
        />

        {/* 발주 총액 요약 */}
        {(Object.values(ffeItems).some((arr) => arr.length > 0) || oseItems.length > 0) && (
          <div className="bg-white border border-slate-200 rounded-xl p-6">
            <div className="flex items-center justify-between">
              <span className="text-sm font-medium tracking-wide text-slate-500">발주 총액</span>
              <div className="text-right">
                {(() => {
                  const ffeTotal = roomTypes.reduce((sum, rt) => {
                    const items = ffeItems[rt.id] || [];
                    const roomCount = roomTypeTotal(rt);
                    return sum + items.reduce((s, it) => s + (it.unitPrice + (it.installUnitPrice || 0)) * it.qtyPerRoom * roomCount, 0);
                  }, 0);
                  const oseTotal = oseItems.reduce(
                    (sum, it) => sum + (it.unitPrice + (it.installUnitPrice || 0)) * it.qtyPerRoom * grandTotal,
                    0
                  );
                  return (
                    <div className="space-y-1">
                      <div className="text-xs text-slate-500">
                        FF&E {won(ffeTotal)} + OS&E {won(oseTotal)}
                      </div>
                      <div className="text-lg font-semibold text-slate-800">{won(ffeTotal + oseTotal)}</div>
                    </div>
                  );
                })()}
              </div>
            </div>
          </div>
        )}

        {/* 프로젝트 대시보드 */}
        {roomTypes.length > 0 &&
          (() => {
            const sumBudget = (items, roomCountFn) =>
              items.reduce((s, it) => s + (it.unitPrice + (it.installUnitPrice || 0)) * it.qtyPerRoom * roomCountFn(it), 0);
            const sumActual = (items, roomCountFn) =>
              items.reduce((s, it) => s + ((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * it.qtyPerRoom * roomCountFn(it), 0);

            let ffeBudget = 0;
            let ffeActual = 0;
            roomTypes.forEach((rt) => {
              const items = ffeItems[rt.id] || [];
              const roomCount = roomTypeTotal(rt);
              ffeBudget += sumBudget(items, () => roomCount);
              ffeActual += sumActual(items, () => roomCount);
            });
            const oseBudget = sumBudget(oseItems, () => grandTotal);
            const oseActual = sumActual(oseItems, () => grandTotal);
            const siteBudget = siteExpenses.reduce((s, it) => s + (it.budgetAmount || 0), 0);
            const siteActual = siteExpenses.reduce((s, it) => s + (it.actualAmount || 0), 0);
            const laborBudget = laborExpenses.reduce((s, it) => s + (it.budgetAmount || 0), 0);
            const laborActual = laborExpenses.reduce((s, it) => s + (it.actualAmount || 0), 0);
            const extraBudget = extraExpenses.reduce((s, it) => s + (it.budgetAmount || 0), 0);
            const extraActual = extraExpenses.reduce((s, it) => s + (it.actualAmount || 0), 0);

            const plannedTotal = ffeBudget + oseBudget + siteBudget + laborBudget + extraBudget;
            const actualTotal = ffeActual + oseActual + siteActual + laborActual + extraActual;
            const perRoomBudget = grandTotal > 0 ? totalBudget / grandTotal : 0;
            const perRoomPlanned = grandTotal > 0 ? plannedTotal / grandTotal : 0;
            const remaining = totalBudget - actualTotal; // 잔여비 (총예산 - 실사용비)
            const planVsActual = plannedTotal - actualTotal; // 계획 대비 집행 차이

            const categoryRows = [
              { name: "FF&E", budget: ffeBudget, actual: ffeActual },
              { name: "OS&E", budget: oseBudget, actual: oseActual },
              { name: "현장지출", budget: siteBudget, actual: siteActual },
              { name: "인건비", budget: laborBudget, actual: laborActual },
              { name: "예산외 지출", budget: extraBudget, actual: extraActual },
            ];

            const ffeByRoomType = roomTypes
              .map((rt) => {
                const items = ffeItems[rt.id] || [];
                const roomCount = roomTypeTotal(rt);
                return {
                  name: codeFor(rt),
                  예산: sumBudget(items, () => roomCount),
                  집행: sumActual(items, () => roomCount),
                };
              })
              .filter((d) => d.예산 > 0 || d.집행 > 0)
              .sort((a, b) => b.예산 - a.예산);

            // 품목명 기준 전체 집계 (룸타입 구분 없이 뭉뚱그려서)
            const itemAgg = new Map();
            roomTypes.forEach((rt) => {
              const roomCount = roomTypeTotal(rt);
              (ffeItems[rt.id] || []).forEach((it) => {
                if (!it.name) return;
                const cur = itemAgg.get(it.name) || { name: it.name, budget: 0, actual: 0 };
                cur.budget += (it.unitPrice + (it.installUnitPrice || 0)) * it.qtyPerRoom * roomCount;
                cur.actual += ((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * it.qtyPerRoom * roomCount;
                itemAgg.set(it.name, cur);
              });
            });
            oseItems.forEach((it) => {
              if (!it.name) return;
              const cur = itemAgg.get(it.name) || { name: it.name, budget: 0, actual: 0 };
              cur.budget += (it.unitPrice + (it.installUnitPrice || 0)) * it.qtyPerRoom * grandTotal;
              cur.actual += ((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * it.qtyPerRoom * grandTotal;
              itemAgg.set(it.name, cur);
            });
            const itemAggRows = [...itemAgg.values()].sort((a, b) => b.budget - a.budget);

            const remainColor = remaining < 0 ? "text-rose-700" : "text-emerald-700";

            function exportToExcel() {
              const wb = XLSX.utils.book_new();

              const summarySheet = XLSX.utils.aoa_to_sheet([
                [`${projectName || "프로젝트"} 예산 보고`],
                [],
                ["오픈바이징 배정 예산", totalBudget],
                ["총 객실수", grandTotal],
                ["객실당 배정예산", Math.round(perRoomBudget)],
                ["예산 예상 사용비(계획)", plannedTotal],
                ["객실당 예상 사용비", Math.round(perRoomPlanned)],
                ["실제 집행 금액", actualTotal],
                ["잔여비(예산-실사용)", remaining],
                ["계획 대비 집행차이", planVsActual],
                [],
                ["대분류", "예산", "집행", "잔여", "집행률"],
                ...categoryRows.map((c) => [
                  c.name,
                  c.budget,
                  c.actual,
                  c.budget - c.actual,
                  c.budget > 0 ? `${((c.actual / c.budget) * 100).toFixed(0)}%` : "0%",
                ]),
                ["합계", plannedTotal, actualTotal, plannedTotal - actualTotal],
              ]);
              XLSX.utils.book_append_sheet(wb, summarySheet, "요약");

              const ffeSheet = XLSX.utils.aoa_to_sheet([
                ["룸타입", "예산", "집행"],
                ...ffeByRoomType.map((d) => [d.name, d.예산, d.집행]),
              ]);
              XLSX.utils.book_append_sheet(wb, ffeSheet, "룸타입별FFE");

              const itemSheet = XLSX.utils.aoa_to_sheet([
                ["품목명", "예산금액", "집행금액", "차액"],
                ...itemAggRows.map((r) => [r.name, r.budget, r.actual, r.budget - r.actual]),
              ]);
              XLSX.utils.book_append_sheet(wb, itemSheet, "품목별집계");

              const floorSheet = XLSX.utils.aoa_to_sheet([
                ["층", ...roomTypes.map((rt) => codeFor(rt)), "합계"],
                ...floors.map((f) => [
                  f,
                  ...roomTypes.map((rt) => rt.byFloor[f] || 0),
                  floorTotal(f),
                ]),
                ["합계", ...roomTypes.map((rt) => roomTypeTotal(rt)), grandTotal],
              ]);
              XLSX.utils.book_append_sheet(wb, floorSheet, "층별배치");

              XLSX.writeFile(wb, `${projectName || "프로젝트"}_보고서.xlsx`);
            }

            return (
              <div className="bg-white border border-slate-200 rounded-xl p-6" id="report-print-area">
                <div className="flex items-center justify-between mb-5 flex-wrap gap-2 no-print">
                  <div className="flex items-center gap-2 text-slate-500">
                    <LayoutDashboard size={18} />
                    <span className="text-sm font-medium tracking-wide">
                      프로젝트 대시보드 — {projectName || "프로젝트명 미입력"}
                    </span>
                  </div>
                  <div className="flex items-center gap-2">
                    <button
                      onClick={exportToExcel}
                      className="text-xs border border-slate-300 rounded-lg px-3 py-1.5 hover:bg-slate-50"
                    >
                      엑셀로 내보내기
                    </button>
                    <button
                      onClick={() => window.print()}
                      className="text-xs border border-slate-300 rounded-lg px-3 py-1.5 hover:bg-slate-50"
                    >
                      PDF로 저장 (인쇄)
                    </button>
                  </div>
                </div>
                <p className="text-sm font-medium tracking-wide text-slate-500 mb-5 hidden print:block">
                  프로젝트 대시보드 — {projectName || "프로젝트명 미입력"}
                </p>

                {/* 핵심 지표 카드 */}
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-3">
                  {[
                    { label: "오픈바이징 배정 예산", value: won(totalBudget) },
                    { label: "총 객실수", value: `${grandTotal}실` },
                    { label: "객실당 배정예산", value: won(perRoomBudget) },
                    { label: "예산 예상 사용비 (계획)", value: won(plannedTotal) },
                  ].map((m) => (
                    <div key={m.label} className="bg-slate-50 border border-slate-200 rounded-lg p-3 text-center">
                      <p className="text-[11px] text-slate-500 mb-1">{m.label}</p>
                      <p className="text-base font-semibold text-slate-800">{m.value}</p>
                    </div>
                  ))}
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-6">
                  {[
                    { label: "객실당 예상 사용비", value: won(perRoomPlanned) },
                    { label: "실제 집행 금액", value: won(actualTotal), color: "text-teal-700" },
                    { label: "잔여비 (예산-실사용)", value: won(remaining), color: remainColor },
                    { label: "계획 대비 집행차이", value: won(planVsActual), color: planVsActual < 0 ? "text-rose-700" : "text-emerald-700" },
                  ].map((m) => (
                    <div key={m.label} className="bg-slate-50 border border-slate-200 rounded-lg p-3 text-center">
                      <p className="text-[11px] text-slate-500 mb-1">{m.label}</p>
                      <p className={`text-base font-semibold ${m.color || "text-slate-800"}`}>{m.value}</p>
                    </div>
                  ))}
                </div>

                {/* 대분류별 예산 대비 집행 */}
                <div className="mb-6">
                  <p className="text-xs text-slate-500 mb-2">대분류별 예산 대비 집행</p>
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-left text-xs text-slate-500 border-b border-slate-200">
                        <th className="py-1.5 font-normal">대분류</th>
                        <th className="py-1.5 font-normal text-right">예산</th>
                        <th className="py-1.5 font-normal text-right">집행</th>
                        <th className="py-1.5 font-normal text-right">잔여</th>
                        <th className="py-1.5 font-normal text-right">집행률</th>
                      </tr>
                    </thead>
                    <tbody>
                      {categoryRows.map((c) => {
                        const rem = c.budget - c.actual;
                        const rate = c.budget > 0 ? (c.actual / c.budget) * 100 : 0;
                        return (
                          <tr key={c.name} className="border-b border-slate-100">
                            <td className="py-1.5">{c.name}</td>
                            <td className="py-1.5 text-right">{won(c.budget)}</td>
                            <td className="py-1.5 text-right text-teal-700">{won(c.actual)}</td>
                            <td className={`py-1.5 text-right ${rem < 0 ? "text-rose-700" : "text-slate-600"}`}>
                              {won(rem)}
                            </td>
                            <td className="py-1.5 text-right text-slate-500">{rate.toFixed(0)}%</td>
                          </tr>
                        );
                      })}
                      <tr className="font-semibold">
                        <td className="py-1.5">합계</td>
                        <td className="py-1.5 text-right">{won(plannedTotal)}</td>
                        <td className="py-1.5 text-right text-teal-700">{won(actualTotal)}</td>
                        <td className={`py-1.5 text-right ${plannedTotal - actualTotal < 0 ? "text-rose-700" : "text-slate-800"}`}>
                          {won(plannedTotal - actualTotal)}
                        </td>
                        <td className="py-1.5 text-right text-slate-500">
                          {plannedTotal > 0 ? ((actualTotal / plannedTotal) * 100).toFixed(0) : 0}%
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* 룸타입별 FF&E 예산 vs 집행 */}
                {ffeByRoomType.length > 0 && (
                  <div className="mb-6">
                    <p className="text-xs text-slate-500 mb-2">룸타입별 FF&E 예산 대비 집행</p>
                    <div style={{ width: "100%", height: 240 }}>
                      <ResponsiveContainer>
                        <BarChart data={ffeByRoomType} layout="vertical" margin={{ left: 20 }}>
                          <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                          <XAxis type="number" tick={{ fontSize: 10 }} tickFormatter={(v) => v.toLocaleString("ko-KR")} />
                          <YAxis type="category" dataKey="name" tick={{ fontSize: 11 }} width={70} />
                          <Tooltip formatter={(v) => won(v)} />
                          <Legend wrapperStyle={{ fontSize: 11 }} />
                          <Bar dataKey="예산" fill="#c7d2fe" radius={[0, 4, 4, 0]} />
                          <Bar dataKey="집행" fill="#2dd4bf" radius={[0, 4, 4, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </div>
                )}

                {/* 품목별 집계 (룸타입 구분 없이 뭉뚱그려서) */}
                {itemAggRows.length > 0 && (
                  <div>
                    <p className="text-xs text-slate-500 mb-2">품목별 지출 (전체 룸타입 통합 집계)</p>
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="text-left text-xs text-slate-500 border-b border-slate-200">
                          <th className="py-1.5 font-normal">품목명</th>
                          <th className="py-1.5 font-normal text-right">예산금액</th>
                          <th className="py-1.5 font-normal text-right">집행금액</th>
                          <th className="py-1.5 font-normal text-right">차액</th>
                        </tr>
                      </thead>
                      <tbody>
                        {itemAggRows.map((row) => (
                          <tr key={row.name} className="border-b border-slate-100">
                            <td className="py-1.5">{row.name}</td>
                            <td className="py-1.5 text-right">{won(row.budget)}</td>
                            <td className="py-1.5 text-right text-teal-700">{won(row.actual)}</td>
                            <td
                              className={`py-1.5 text-right ${
                                row.budget - row.actual < 0 ? "text-rose-700" : "text-slate-500"
                              }`}
                            >
                              {won(row.budget - row.actual)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            );
          })()}
          </>
        )}
      </div>
    </div>
  );
}
