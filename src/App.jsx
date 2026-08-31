import { useState, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { Plus, X, Building2, LayoutGrid, Table2, Trash2, Upload, Save, Users, Loader2, LayoutDashboard, ChevronDown, ChevronRight } from "lucide-react";
import {
  Tooltip, Legend, ResponsiveContainer,
  BarChart, Bar, XAxis, YAxis, CartesianGrid,
} from "recharts";

import { getProjectIndex, getProjectData, saveProjectData, saveProjectIndex } from "./api";
import { saveRoomTypes, loadRoomTypes } from "./roomTypesApi";
import { saveOrderItems, loadOrderItems } from "./orderItemsApi";
import { saveExpenses, loadExpenses } from "./expensesApi";
import { resolveProjectUuid } from "./projectIdApi";
import { saveProjectSettings, loadProjectSettings } from "./projectSettingsApi";
import { fetchContentPresets, fetchOseContentPresets } from "./contentPresetsApi";
import { loadVendors, createVendor, updateVendor, deleteVendor } from "./vendorsApi";
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


const DEFAULT_BASIC_PRESET = [
  "헤드보드", "침대바디", "협탁", "옷장", "붙박이 책장", "콘센트", "슬리퍼걸이",
  "구둣주걱", "리모컨거치대", "입식 테이블", "입식 의자", "객실 슬리퍼", "욕실 슬리퍼",
  "객실 휴지통", "욕실 휴지통", "객실 휴지케이스", "옷걸이", "바지걸이", "TV",
  "소형 무소음 냉장고", "거울", "조명",
];

let idCounter = 1;
const nextId = () => idCounter++;

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

// 인건비 지출 전용 — 현장 알바처럼 인원수가 유동적인(3~20명 등) 항목을 단가×인원수로 관리.
// 예산/집행 각각 단가·인원수를 따로 둬서(집행 시점에 실제 투입 인원이 달라질 수 있으므로) 금액은 자동 계산됨.
function LaborExpenseSection({ items, onAdd, onAddDefaults, onUpdate, onRemove }) {
  const budgetTotal = items.reduce((s, it) => s + (it.budgetAmount || 0), 0);
  const actualTotal = items.reduce((s, it) => s + (it.actualAmount || 0), 0);
  return (
    <div className="bg-white border border-slate-200 rounded-xl p-6">
      <div className="flex items-center justify-between mb-4">
        <span className="text-sm font-medium tracking-wide text-slate-500">인건비 지출</span>
        <div className="flex items-center gap-1.5">
          <button
            onClick={onAddDefaults}
            className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50"
            title="현장 운용 인원(알바)/사진 작가/모델 3개 항목을 기본값(인원 1명)으로 추가"
          >
            기본 인원 추가
          </button>
          <button
            onClick={onAdd}
            className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50 flex items-center gap-1"
          >
            <Plus size={13} /> 항목 추가
          </button>
        </div>
      </div>
      {items.length > 0 ? (
        <>
          <div className="max-h-96 overflow-y-auto">
            <table className="w-full text-sm">
              <thead className="sticky top-0 bg-white z-10">
                <tr className="text-left text-xs text-slate-500 border-b border-slate-200">
                  <th className="py-1.5 font-normal">항목명</th>
                  <th className="py-1.5 font-normal text-right">예산단가</th>
                  <th className="py-1.5 font-normal text-right">인원수(예산)</th>
                  <th className="py-1.5 font-normal text-right" title="며칠 운용하는지(예: 알바 2~3일)">운용일수(예산)</th>
                  <th className="py-1.5 font-normal text-right">집행단가</th>
                  <th className="py-1.5 font-normal text-right">인원수(집행)</th>
                  <th className="py-1.5 font-normal text-right">운용일수(집행)</th>
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
                        placeholder="예: 현장 운용 인원(알바)"
                        className="w-full border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        value={it.unitPrice || 0}
                        onChange={(e) => onUpdate(it.id, "unitPrice", e.target.value)}
                        className="w-24 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        value={it.quantity != null ? it.quantity : 1}
                        onChange={(e) => onUpdate(it.id, "quantity", e.target.value)}
                        title="예: 현장 알바는 3~20명까지 조정 가능"
                        className="w-16 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        step="0.5"
                        value={it.days != null ? it.days : 1}
                        onChange={(e) => onUpdate(it.id, "days", e.target.value)}
                        title="예: 2~3일씩 고용하는 경우"
                        className="w-14 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        value={it.actualUnitPrice || 0}
                        onChange={(e) => onUpdate(it.id, "actualUnitPrice", e.target.value)}
                        className="w-24 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 bg-teal-50/40"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        value={it.actualQuantity != null ? it.actualQuantity : 1}
                        onChange={(e) => onUpdate(it.id, "actualQuantity", e.target.value)}
                        className="w-16 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 bg-teal-50/40"
                      />
                    </td>
                    <td className="py-1.5">
                      <input
                        type="number"
                        min="0"
                        step="0.5"
                        value={it.actualDays != null ? it.actualDays : 1}
                        onChange={(e) => onUpdate(it.id, "actualDays", e.target.value)}
                        className="w-14 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 bg-teal-50/40"
                      />
                    </td>
                    <td className="py-1.5 text-right font-medium">
                      {(it.budgetAmount || 0).toLocaleString("ko-KR")}원
                    </td>
                    <td className="py-1.5 text-right font-medium text-teal-700">
                      {(it.actualAmount || 0).toLocaleString("ko-KR")}원
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
        <p className="text-xs text-slate-400">항목을 추가하면 단가×인원수×운용일수로 예산/집행 금액이 자동 계산돼요.</p>
      )}
    </div>
  );
}

// 업체 마스터 관리 — 전역 공통(vendors 테이블), 특정 프로젝트에 속하지 않음.
// 저장 버튼 없이 각 입력/추가/삭제가 즉시 Supabase에 반영됨(vendorsApi.js, 프로젝트 추가/삭제와 동일한 방식).
function VendorManagement({ vendors, loading, error, onReload, onAdd, onFieldChange, onFieldCommit, onRemove }) {
  const [newName, setNewName] = useState("");
  const [adding, setAdding] = useState(false);

  async function handleAdd() {
    if (!newName.trim()) return;
    setAdding(true);
    try {
      await onAdd(newName.trim());
      setNewName("");
    } finally {
      setAdding(false);
    }
  }

  const COLUMNS = [
    { field: "name", label: "업체명", width: "w-32", placeholder: "예: 대주상사" },
    { field: "itemGroup", label: "항목", width: "w-24", placeholder: "예: 초도비품" },
    { field: "category", label: "분류", width: "w-28", placeholder: "예: 린넨류" },
    { field: "bankName", label: "은행", width: "w-20", placeholder: "" },
    { field: "accountNumber", label: "계좌번호", width: "w-32", placeholder: "" },
    { field: "accountHolder", label: "예금주", width: "w-32", placeholder: "" },
    { field: "contactName", label: "담당자", width: "w-24", placeholder: "" },
    { field: "phone", label: "연락처", width: "w-28", placeholder: "010-0000-0000" },
    { field: "email", label: "이메일", width: "w-36", placeholder: "" },
    { field: "remark", label: "비고", width: "w-32", placeholder: "" },
  ];

  return (
    <div className="bg-white border border-slate-200 rounded-xl p-6 space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-base font-semibold text-slate-800">업체 관리</h2>
          <p className="text-xs text-slate-500 mt-0.5">
            전역 공통 업체 마스터 — 여기서 등록한 업체를 발주 품목에 배정하고, 발주서 생성 시 업체별로 그룹핑합니다.
          </p>
        </div>
        <button
          onClick={onReload}
          className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50 flex items-center gap-1"
        >
          <Loader2 size={13} className={loading ? "animate-spin" : ""} /> 새로고침
        </button>
      </div>

      {error && (
        <p className="text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-md px-3 py-2">{error}</p>
      )}

      <div className="flex items-end gap-2">
        <div className="flex-1 max-w-xs">
          <label className="block text-xs text-slate-500 mb-1">새 업체명</label>
          <input
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleAdd()}
            placeholder="업체명 입력 후 추가"
            className="w-full border border-slate-300 rounded-md px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
          />
        </div>
        <button
          onClick={handleAdd}
          disabled={!newName.trim() || adding}
          className="text-sm bg-amber-700 text-white px-3 py-1.5 rounded-lg hover:bg-amber-800 disabled:opacity-50 flex items-center gap-1.5"
        >
          <Plus size={14} /> 업체 추가
        </button>
      </div>

      {loading ? (
        <p className="text-xs text-slate-400 flex items-center gap-1.5">
          <Loader2 size={13} className="animate-spin" /> 불러오는 중...
        </p>
      ) : vendors.length === 0 ? (
        <p className="text-xs text-slate-400">등록된 업체가 없습니다. 위에서 업체명을 입력해 추가하세요.</p>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-xs text-slate-500 border-b border-slate-200">
                {COLUMNS.map((c) => (
                  <th key={c.field} className="py-1.5 font-normal pr-2">{c.label}</th>
                ))}
                <th></th>
              </tr>
            </thead>
            <tbody>
              {vendors.map((v) => (
                <tr key={v.id} className="border-b border-slate-100">
                  {COLUMNS.map((c) => (
                    <td key={c.field} className="py-1.5 pr-2">
                      <input
                        value={v[c.field] || ""}
                        placeholder={c.placeholder}
                        onChange={(e) => onFieldChange(v.id, c.field, e.target.value)}
                        onBlur={(e) => onFieldCommit(v.id, c.field, e.target.value)}
                        className={`${c.width} border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500`}
                      />
                    </td>
                  ))}
                  <td className="py-1.5 pl-2">
                    <button onClick={() => onRemove(v)} className="text-slate-400 hover:text-rose-600" title="삭제">
                      <Trash2 size={14} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}


export default function App() {
  const [projectName, setProjectName] = useState("");
  const [tier, setTier] = useState(TIERS[2]);
  const [totalBudget, setTotalBudget] = useState(0); // 전체 공사비 중 오픈바이징팀에 배정된 예산
  // 목표 객실수 + 실당 예산: 둘 다 입력하면 오픈바이징 배정 예산을 자동 계산해준다.
  // (자동계산은 이 두 필드를 직접 건드릴 때만 일어남 — totalBudget을 직접 수정하면 그 값 그대로 유지됨)
  const [targetRoomCount, setTargetRoomCount] = useState(0);
  const [budgetPerRoom, setBudgetPerRoom] = useState(0);

  function handleTargetRoomCountChange(value) {
    const n = Math.max(0, parseInt(value || "0", 10) || 0);
    setTargetRoomCount(n);
    if (budgetPerRoom > 0) setTotalBudget(n * budgetPerRoom);
  }
  function handleBudgetPerRoomChange(value) {
    const n = Math.max(0, parseFloat(value || "0") || 0);
    setBudgetPerRoom(n);
    if (targetRoomCount > 0) setTotalBudget(targetRoomCount * n);
  }

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
  const [activeTab, setActiveTab] = useState("main"); // "main" | "test-schedule" | "vendors"

  // 업체 마스터(전역 공통, vendors 테이블) — "업체 관리" 탭을 처음 열 때 로드
  const [vendors, setVendors] = useState([]);
  const [vendorsLoaded, setVendorsLoaded] = useState(false);
  const [vendorsLoading, setVendorsLoading] = useState(false);
  const [vendorsError, setVendorsError] = useState("");

  async function reloadVendors() {
    setVendorsLoading(true);
    setVendorsError("");
    try {
      const list = await loadVendors();
      setVendors(list);
      setVendorsLoaded(true);
    } catch (err) {
      setVendorsError(err.message || "업체 목록을 불러오지 못했습니다.");
    } finally {
      setVendorsLoading(false);
    }
  }

  useEffect(() => {
    if (activeTab === "vendors" && !vendorsLoaded) {
      reloadVendors();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeTab]);

  async function handleAddVendor(name) {
    setVendorsError("");
    try {
      const created = await createVendor({ name });
      setVendors((prev) => [...prev, created].sort((a, b) => a.name.localeCompare(b.name, "ko")));
    } catch (err) {
      setVendorsError(err.message || "업체 추가에 실패했습니다.");
    }
  }

  // 입력 중엔 로컬 상태만 갱신(타이핑마다 요청 보내지 않음), blur 시점에 실제 저장
  function handleVendorFieldChange(id, field, value) {
    setVendors((prev) => prev.map((v) => (v.id === id ? { ...v, [field]: value } : v)));
  }

  async function handleVendorFieldCommit(id, field, value) {
    try {
      await updateVendor(id, { [field]: value });
    } catch (err) {
      setVendorsError(err.message || "업체 정보 저장에 실패했습니다.");
    }
  }

  async function handleRemoveVendor(vendor) {
    if (!window.confirm(`"${vendor.name}" 업체를 삭제할까요?`)) return;
    setVendorsError("");
    try {
      await deleteVendor(vendor.id);
      setVendors((prev) => prev.filter((v) => v.id !== vendor.id));
    } catch (err) {
      setVendorsError(err.message || "업체 삭제에 실패했습니다.");
    }
  }
  const [showFloorPlan, setShowFloorPlan] = useState(true);
  const [floorPlanAutoOffApplied, setFloorPlanAutoOffApplied] = useState(false);

  // 발주 품목: FF&E는 룸타입별, OS&E는 공통 리스트 (각 항목은 예산단가 unitPrice + 집행단가 actualUnitPrice)
  const [ffeItems, setFfeItems] = useState({}); // { [roomTypeId]: [{id,name,unitPrice,actualUnitPrice,qtyPerRoom}] }
  const [oseItems, setOseItems] = useState([]); // [{id,name,unitPrice,actualUnitPrice,qtyPerRoom}]
  // 룸타입별 품목 카드 접기/펼치기 — 품목이 많아지면 스크롤이 매우 길어져서 필요한 카드만 펴놓고 볼 수 있게
  const [collapsedRoomTypeIds, setCollapsedRoomTypeIds] = useState(new Set());
  function toggleRoomTypeCollapsed(roomTypeId) {
    setCollapsedRoomTypeIds((prev) => {
      const next = new Set(prev);
      if (next.has(roomTypeId)) next.delete(roomTypeId);
      else next.add(roomTypeId);
      return next;
    });
  }

  // 현장지출 / 인건비 지출 / 예산외 지출 (현장지출·예산외지출은 직접입력, 인건비지출만 단가×인원수 자동계산)
  const [siteExpenses, setSiteExpenses] = useState([]); // [{id,name,budgetAmount,actualAmount}]
  const [laborExpenses, setLaborExpenses] = useState([]); // [{id,name,unitPrice,actualUnitPrice,quantity,actualQuantity,budgetAmount,actualAmount}]
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
  const DEFAULT_LABOR_ROLES = ["현장 운용 인원(알바)", "사진 작가", "모델"];
  // 인건비 지출 전용 핸들러 — budgetAmount/actualAmount를 단가×인원수로 항상 재계산해서 저장
  // (다른 집계 로직(laborBudget 등)은 그대로 budgetAmount/actualAmount만 읽으므로 이 계산만 맞으면 나머지는 자동으로 맞음)
  function recalcLaborItem(it) {
    return {
      ...it,
      budgetAmount:
        (it.unitPrice || 0) * (it.quantity != null ? it.quantity : 1) * (it.days != null ? it.days : 1),
      actualAmount:
        (it.actualUnitPrice || 0) *
        (it.actualQuantity != null ? it.actualQuantity : 1) *
        (it.actualDays != null ? it.actualDays : 1),
    };
  }
  function addLaborExpense() {
    setLaborExpenses((prev) => [
      ...prev,
      recalcLaborItem({
        id: nextId(),
        name: "",
        unitPrice: 0,
        actualUnitPrice: 0,
        quantity: 1,
        actualQuantity: 1,
        days: 1,
        actualDays: 1,
      }),
    ]);
  }
  function addDefaultLaborRoles() {
    setLaborExpenses((prev) => [
      ...prev,
      ...DEFAULT_LABOR_ROLES.map((name) =>
        recalcLaborItem({
          id: nextId(),
          name,
          unitPrice: 0,
          actualUnitPrice: 0,
          quantity: 1,
          actualQuantity: 1,
          days: 1,
          actualDays: 1,
        })
      ),
    ]);
  }
  function updateLaborExpense(id, field, value) {
    setLaborExpenses((prev) =>
      prev.map((it) => {
        if (it.id !== id) return it;
        const updated = { ...it, [field]: field === "name" ? value : Math.max(0, parseFloat(value || "0") || 0) };
        return recalcLaborItem(updated);
      })
    );
  }
  function removeLaborExpense(id) {
    setLaborExpenses((prev) => prev.filter((it) => it.id !== id));
  }
  const siteExpenseHandlers = makeExpenseHandlers(setSiteExpenses);
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

  // Supabase content_presets(전사 표준 초도발주 템플릿)를 이 룸타입에 불러와 채워넣음
  // 룸타입의 category와 일치하는 전용 콘텐츠 + 모든 룸타입 공통 베이스를 함께 가져옴
  const [loadingPresetFor, setLoadingPresetFor] = useState(null);
  const [presetError, setPresetError] = useState("");
  // 카탈로그 기본세트를 즉시 다 쏟아붓지 않고, 먼저 불러온 뒤 세부 카테고리(린넨류/타올류 등)를
  // 골라서 선택한 것만 실제로 추가하는 중간 선택 단계 — presetPickerFor: roomTypeId | "OSE" | null
  const [presetPickerFor, setPresetPickerFor] = useState(null);
  const [presetPickerItems, setPresetPickerItems] = useState([]);
  const [presetPickerSelectedCats, setPresetPickerSelectedCats] = useState(new Set());
  const presetPickerCatCounts = useMemo(() => {
    const map = new Map();
    presetPickerItems.forEach((p) => {
      const cat = p.subCategory || "기타";
      map.set(cat, (map.get(cat) || 0) + 1);
    });
    return [...map.entries()];
  }, [presetPickerItems]);
  function togglePresetPickerCat(cat) {
    setPresetPickerSelectedCats((prev) => {
      const next = new Set(prev);
      next.has(cat) ? next.delete(cat) : next.add(cat);
      return next;
    });
  }
  function cancelPresetPicker() {
    setPresetPickerFor(null);
    setPresetPickerItems([]);
  }
  function presetItemsToObjects(presets) {
    return presets.map((p) => ({
      id: nextId(),
      name: p.name,
      unitPrice: p.unitPrice,
      actualUnitPrice: 0,
      installUnitPrice: 0,
      installActualUnitPrice: 0,
      qtyPerRoom: p.qtyPerRoom,
      calcBasis: p.calcBasis,
      multiplier: p.multiplier,
      mattressSize: p.mattressSize,
      catalogItemId: p.catalogItemId,
      categoryGroup: p.categoryGroup, // 'FF&E' | 'OS&E' — 카탈로그 실제 회계분류, 카드 위치와 무관
      subCategory: p.subCategory, // 세부 품목군(객실비품/린넨류/타올류/매트리스/기기류)
      cartonSize: p.cartonSize, // 박스/팩당 개수(카탈로그 값) — 있으면 필요수량을 이 배수로 올림해서 발주수량 산출
    }));
  }
  function confirmPresetPicker() {
    const chosen = presetPickerItems.filter((p) =>
      presetPickerSelectedCats.has(p.subCategory || "기타")
    );
    const newItems = presetItemsToObjects(chosen);
    if (presetPickerFor === "OSE") {
      setOseItems((prev) => [...prev, ...newItems]);
    } else if (presetPickerFor) {
      setFfeItems((prev) => ({
        ...prev,
        [presetPickerFor]: [...(prev[presetPickerFor] || []), ...newItems],
      }));
    }
    cancelPresetPicker();
  }
  async function openPresetPicker(rt) {
    setLoadingPresetFor(rt.id);
    setPresetError("");
    try {
      const presets = await fetchContentPresets(rt.category);
      setPresetPickerItems(presets);
      setPresetPickerSelectedCats(new Set(presets.map((p) => p.subCategory || "기타")));
      setPresetPickerFor(rt.id);
    } catch (err) {
      setPresetError(`기본세트를 불러오지 못했어요: ${err.message}`);
    } finally {
      setLoadingPresetFor(null);
    }
  }
  function updateFfeItem(roomTypeId, itemId, field, value) {
    const stringFields = ["name", "calcBasis", "mattressSize", "categoryGroup", "subCategory"];
    setFfeItems((prev) => ({
      ...prev,
      [roomTypeId]: (prev[roomTypeId] || []).map((it) =>
        it.id === itemId
          ? { ...it, [field]: stringFields.includes(field) ? value : Math.max(0, parseFloat(value || "0") || 0) }
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
    const stringFields = ["name", "categoryGroup", "subCategory"];
    setOseItems((prev) =>
      prev.map((it) =>
        it.id === itemId
          ? { ...it, [field]: stringFields.includes(field) ? value : Math.max(0, parseFloat(value || "0") || 0) }
          : it
      )
    );
  }
  function removeOseItem(itemId) {
    setOseItems((prev) => prev.filter((it) => it.id !== itemId));
  }

  // Supabase content_presets(room 기준 공통베이스)를 OS&E 공통 품목 리스트에 불러오기 위한 선택 단계 진입
  const [loadingOsePreset, setLoadingOsePreset] = useState(false);
  const [osePresetError, setOsePresetError] = useState("");
  async function openOsePresetPicker() {
    setLoadingOsePreset(true);
    setOsePresetError("");
    try {
      const presets = await fetchOseContentPresets();
      setPresetPickerItems(presets);
      setPresetPickerSelectedCats(new Set(presets.map((p) => p.subCategory || "기타")));
      setPresetPickerFor("OSE");
    } catch (err) {
      setOsePresetError(`기본세트를 불러오지 못했어요: ${err.message}`);
    } finally {
      setLoadingOsePreset(false);
    }
  }

  const won = (n) => `${Math.round(n).toLocaleString("ko-KR")}원`;

  const [draft, setDraft] = useState({
    bed: BED_TYPES[0],
    bathtub: BATHTUB[1],
    category: categories[0] || "",
    irregular: [],
    mattressQty: 1,
    capacity: 2, // 수용인원 — 베개류 등 인당(capacity) 기준 발주 계산에 사용
    bedComposition: [], // 복합 침대구성, 예: [{size:'Q',qty:1},{size:'S',qty:1}] — 비어있으면 bed/mattressQty로 단일구성 간주
    grade: "Superior",
    features: [],
    view: "",
    includeBedInName: false,
    includeViewInName: false,
    customName: "",
  });

  function addDraftBedCompRow() {
    setDraft((d) => ({ ...d, bedComposition: [...d.bedComposition, { size: "Q", qty: 1 }] }));
  }
  function updateDraftBedCompRow(idx, field, value) {
    setDraft((d) => ({
      ...d,
      bedComposition: d.bedComposition.map((row, i) =>
        i === idx ? { ...row, [field]: field === "qty" ? Math.max(1, parseInt(value || "1", 10)) : value } : row
      ),
    }));
  }
  function removeDraftBedCompRow(idx) {
    setDraft((d) => ({ ...d, bedComposition: d.bedComposition.filter((_, i) => i !== idx) }));
  }

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
        rt.capacity === draft.capacity &&
        rt.grade === draft.grade &&
        rt.view === draft.view &&
        JSON.stringify([...rt.irregular].sort()) === JSON.stringify([...draft.irregular].sort()) &&
        JSON.stringify([...rt.features].sort()) === JSON.stringify([...draft.features].sort()) &&
        JSON.stringify(rt.bedComposition || []) === JSON.stringify(draft.bedComposition || [])
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
        capacity: draft.capacity,
        bedComposition: draft.bedComposition,
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
    setDraft((d) => ({ ...d, irregular: [], features: [], mattressQty: 1, capacity: 2, bedComposition: [], customName: "" }));
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
  function setCapacity(id, value) {
    const n = Math.max(1, parseInt(value || "1", 10) || 1);
    setRoomTypes(roomTypes.map((rt) => (rt.id === id ? { ...rt, capacity: n } : rt)));
  }
  function addBedCompRow(id) {
    setRoomTypes((prev) =>
      prev.map((rt) =>
        rt.id === id ? { ...rt, bedComposition: [...(rt.bedComposition || []), { size: "Q", qty: 1 }] } : rt
      )
    );
  }
  function updateBedCompRow(id, idx, field, value) {
    setRoomTypes((prev) =>
      prev.map((rt) =>
        rt.id === id
          ? {
              ...rt,
              bedComposition: (rt.bedComposition || []).map((row, i) =>
                i === idx ? { ...row, [field]: field === "qty" ? Math.max(1, parseInt(value || "1", 10)) : value } : row
              ),
            }
          : rt
      )
    );
  }
  function removeBedCompRow(id, idx) {
    setRoomTypes((prev) =>
      prev.map((rt) =>
        rt.id === id ? { ...rt, bedComposition: (rt.bedComposition || []).filter((_, i) => i !== idx) } : rt
      )
    );
  }
  // bedComposition이 비어있으면(수동 미입력) bed/mattressQty로부터 단일구성을 유추 — 기존 데이터와 호환용
  function effectiveBedComposition(rt) {
    if (rt.bedComposition && rt.bedComposition.length > 0) return rt.bedComposition;
    const size = rt.bed === "싱글" ? "S" : "Q";
    return [{ size, qty: rt.mattressQty || 1 }];
  }

  // FF&E 품목의 실제 발주수량 계산 — calcBasis(room/capacity/bed)에 따라 분기
  // calcBasis 없는 레거시 품목은 room 기준(qtyPerRoom×roomCount)으로 그대로 동작(하위호환)
  // 카톤/팩 단위로만 살 수 있는 소모품용 올림 처리 — cartonSize가 없으면(낱개 구매 가능) 그대로 반환
  function roundToCarton(qty, cartonSize) {
    if (!cartonSize || cartonSize <= 0) return qty;
    return Math.ceil(qty / cartonSize) * cartonSize;
  }
  function ffeItemQty(it, rt) {
    const roomCount = roomTypeTotal(rt);
    const mult = it.multiplier != null && it.multiplier !== "" ? Number(it.multiplier) : 1;
    const base = Number(it.qtyPerRoom) || 0;
    let raw;
    if (it.calcBasis === "capacity") {
      raw = base * mult * (rt.capacity || 1) * roomCount;
    } else if (it.calcBasis === "bed") {
      const beds = effectiveBedComposition(rt);
      const bedQty = it.mattressSize
        ? beds.filter((b) => b.size === it.mattressSize).reduce((s, b) => s + (b.qty || 0), 0)
        : beds.reduce((s, b) => s + (b.qty || 0), 0);
      raw = base * mult * bedQty * roomCount;
    } else {
      raw = base * mult * roomCount;
    }
    return roundToCarton(raw, it.cartonSize);
  }
  // OS&E 공통 리스트 품목의 필요수량 — 프로젝트 전체 객실수(grandTotal) 비례, 카톤 단위 올림 동일 적용
  function oseItemQty(it) {
    return roundToCarton(it.qtyPerRoom * grandTotal, it.cartonSize);
  }

  // ---- 엑셀 업로드로 룸타입 + 호수 일괄 생성 ----
  const [importSummary, setImportSummary] = useState(null);
  const [importError, setImportError] = useState("");
  const [overwriteOnImport, setOverwriteOnImport] = useState(true);



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
    setTargetRoomCount(0);
    setBudgetPerRoom(0);
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
    setLastSaved(null);
  }

  function applyLoadedData(data) {
    if (data.projectName !== undefined) setProjectName(data.projectName);
    if (data.tier) setTier(data.tier);
    if (data.totalBudget !== undefined) setTotalBudget(data.totalBudget);
    if (data.targetRoomCount !== undefined) setTargetRoomCount(data.targetRoomCount);
    if (data.budgetPerRoom !== undefined) setBudgetPerRoom(data.budgetPerRoom);
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
    if (data.savedAt) setLastSaved(data.savedAt);
    idCounter = Math.max(idCounter, maxIdIn(data) + 1);
  }

  async function switchToProject(id) {
    setIsLoading(true);
    resetProjectState();
    try {
      const data = await getProjectData(id);
      if (data) {
        let merged = { ...data };
        try {
          const projectUuid = await resolveProjectUuid(id, data.projectName);
          const { roomTypes: sbRoomTypes, idMap } = await loadRoomTypes(projectUuid);
          const { ffeItems: sbFfeItems, oseItems: sbOseItems } = await loadOrderItems(projectUuid, idMap);
          const sbExpenses = await loadExpenses(projectUuid);
          const sbSettings = await loadProjectSettings(projectUuid);

          // 이 프로젝트가 실제로 한 번이라도 Supabase 동기화 저장을 거쳤는지 판단.
          // 룸타입/발주품목/지출이 전부 비어있으면 아직 Apps Script 레거시 데이터만
          // 있는 상태로 보고 그대로 둔다 (섣불리 덮어써서 데이터가 사라진 것처럼 보이는 것 방지 —
          // 프로젝트별로 최초 저장이 일어나는 순간 자동으로 Supabase 쪽으로 전환됨).
          const hasSupabaseData =
            sbRoomTypes.length > 0 ||
            sbOseItems.length > 0 ||
            sbExpenses.siteExpenses.length > 0 ||
            sbExpenses.laborExpenses.length > 0 ||
            sbExpenses.extraExpenses.length > 0;

          if (hasSupabaseData) {
            // byFloor(층별 배치)는 아직 Supabase에 없으므로 기존 Apps Script 데이터에서 id 매칭으로 보완
            const legacyById = {};
            (data.roomTypes || []).forEach((rt) => {
              legacyById[rt.id] = rt;
            });
            merged.roomTypes = sbRoomTypes.map((rt) => ({
              ...rt,
              byFloor: (legacyById[rt.id] && legacyById[rt.id].byFloor) || rt.byFloor || {},
            }));
            merged.ffeItems = sbFfeItems;
            merged.oseItems = sbOseItems;
            merged.siteExpenses = sbExpenses.siteExpenses;
            merged.laborExpenses = sbExpenses.laborExpenses;
            merged.extraExpenses = sbExpenses.extraExpenses;
            if (sbSettings) merged = { ...merged, ...sbSettings };
          }
          setSupabaseSyncError("");
        } catch (sbErr) {
          // Supabase 읽기 실패 — Apps Script 데이터만으로 계속 진행(기존 동작과 동일), 경고만 표시
          setSupabaseSyncError("Supabase에서 최신 데이터를 불러오지 못했어요(기존 저장 데이터로 표시 중).");
        }
        applyLoadedData(merged);
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

  // 테스트 공정표 탭에서 프로젝트명을 클릭해 발주 관리 탭으로 바로 넘어올 때 사용.
  // supaProject: Supabase projects 테이블의 행({ id: uuid, client_id, name, ... })
  async function handleOpenInOrderManager(supaProject) {
    if (!supaProject) return;
    let localId = supaProject.client_id;
    let list = projectList;
    if (localId) {
      // client_id는 있는데 로컬 프로젝트 목록에 없는 경우(드물게 인덱스가 어긋난 경우) 목록에 채워 넣는다.
      if (!list.some((p) => p.id === localId)) {
        list = [...list, { id: localId, name: supaProject.name || "(이름 없음)" }];
        setProjectList(list);
        await saveProjectIndex(list);
      }
    } else {
      // 아직 발주 관리 탭에서 한 번도 열어본 적 없는 프로젝트(전사시트로만 시딩됨) — 새 로컬 id를 만들어
      // 목록에 추가하면, switchToProject 내부의 resolveProjectUuid가 이름 매칭으로 이 Supabase 행을 찾아
      // client_id를 자동으로 채워준다.
      localId = `proj_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
      list = [...list, { id: localId, name: supaProject.name || "(이름 없음)" }];
      setProjectList(list);
      await saveProjectIndex(list);
    }
    setActiveTab("main");
    await switchToProject(localId);
  }

  useEffect(() => {
    (async () => {
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
      targetRoomCount,
      budgetPerRoom,
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
        const projectUuid = await resolveProjectUuid(currentProjectId, projectName);
        const roomTypeIdMap = await saveRoomTypes(projectUuid, roomTypes);
        await saveOrderItems(projectUuid, ffeItems, oseItems, roomTypeIdMap);
        await saveExpenses(projectUuid, { siteExpenses, laborExpenses, extraExpenses });
        await saveProjectSettings(projectUuid, {
          categories, irregularOptions, brandRoomName, roomFeatures, viewTypes, floors, basicPreset,
        });
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

  function floorFromRoomNumber(numStr) {
    const digits = numStr.replace(/\D/g, "");
    if (!digits) return null;
    if (digits.length <= 2) return "1F";
    return `${parseInt(digits.slice(0, digits.length - 2), 10)}F`;
  }

  // 호수 셀은 팀마다 표기가 제각각(콤마 구분, 마침표 구분, 공백 구분)이고
  // 마침표로 구분한 값을 엑셀이 숫자로 잘못 인식해 "208.308.408.508" → 208308408508 처럼
  // 구분자가 통째로 소실되는 경우도 있어, 이런 케이스까지 최대한 복구한다.
  function parseRoomNumberTokens(raw) {
    const str = String(raw || "").trim();
    if (!str) return [];
    let tokens = str
      .split(/[,.\s]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    // 구분자가 소실되어 통짜 숫자 하나로 붙어버린 경우, 3자리 단위(호수 표기 관례)로 복구 시도
    if (
      tokens.length === 1 &&
      /^\d+$/.test(tokens[0]) &&
      tokens[0].length > 3 &&
      tokens[0].length % 3 === 0
    ) {
      const digits = tokens[0];
      const chunks = [];
      for (let i = 0; i < digits.length; i += 3) chunks.push(digits.slice(i, i + 3));
      tokens = chunks;
    }
    return tokens;
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
    const bedComposition = [];
    if (singleMatch && !hasQueen) {
      mattressQty = singleMatch[1] ? parseInt(singleMatch[1], 10) : 1;
      bedComposition.push({ size: "S", qty: mattressQty });
    } else if (singleMatch && hasQueen) {
      // 퀸+싱글 복합구성(예: 패밀리룸) — 침대구성 배열로 정확히 기록, 태그는 참고용으로 유지
      extraIrregular.push("싱글베드 추가");
      const singleQty = singleMatch[1] ? parseInt(singleMatch[1], 10) : 1;
      bedComposition.push({ size: "Q", qty: 1 }, { size: "S", qty: singleQty });
    } else if (hasQueen) {
      bedComposition.push({ size: "Q", qty: 1 });
    }
    let bathtub = "무";
    if (paren === "욕조" || /아크릴\s*욕조/.test(text)) bathtub = "유";
    if (paren === "장애인") {
      bathtub = "무";
      extraIrregular.push("장애인객실");
    }
    return { bed, mattressQty, bathtub, extraIrregular, bedComposition };
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
      const normKey = (s) => norm(s).toLowerCase().replace(/['’]/g, "");
      const HEADER_KEYS = ["룸타입", "수량", "인원", "구성", "호수"];

      let rows = null;
      let headerIdx = -1;
      let colIdx = {};
      let format = null; // "A"(구성+호수 상세형) | "B"(설계팀 룸믹스: Room type/Q'ty/Room number)
      const scanned = []; // for diagnostics if nothing matches

      for (const sheetName of wb.SheetNames) {
        const candidateRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "" });
        scanned.push({ sheetName, preview: candidateRows.slice(0, 3) });
        for (let i = 0; i < Math.min(candidateRows.length, 200); i++) {
          const rawRow = candidateRows[i];
          const normedRow = rawRow.map((c) => norm(c));
          const normedKeyRow = rawRow.map((c) => normKey(c));

          // ---- Format A: 회사 관리시트("구성"+"호수" 상세형) ----
          const hoIdxA = normedRow.findIndex((c) => c.includes("호수"));
          if (hoIdxA !== -1) {
            const guessedCols = {};
            ["룸타입", "구성", "인원"].forEach((key) => {
              const idx = normedRow.findIndex((c) => c.includes(key));
              if (idx !== -1) guessedCols[key] = idx;
            });
            if (guessedCols["룸타입"] === undefined) {
              const altIdx = normedRow.findIndex((c) => c.includes("타입"));
              if (altIdx !== -1) guessedCols["룸타입"] = altIdx;
            }
            if (guessedCols["룸타입"] !== undefined && guessedCols["구성"] !== undefined) {
              rows = candidateRows;
              headerIdx = i;
              colIdx = { ...guessedCols, 호수: hoIdxA };
              format = "A";
              break;
            }
          }

          // ---- Format B: 설계팀 룸믹스("객실타입별 수량" 표 — Room type/Q'ty/Room number) ----
          const roomNumIdx =
            hoIdxA !== -1 ? hoIdxA : normedKeyRow.findIndex((c) => c.includes("roomnumber"));
          const qtyIdx =
            normedKeyRow.findIndex((c) => c.includes("qty")) !== -1
              ? normedKeyRow.findIndex((c) => c.includes("qty"))
              : normedRow.findIndex((c) => c.includes("수량"));
          const typeIdx =
            normedRow.findIndex((c) => c.includes("룸타입")) !== -1
              ? normedRow.findIndex((c) => c.includes("룸타입"))
              : normedKeyRow.findIndex((c) => c.includes("roomtype")) !== -1
              ? normedKeyRow.findIndex((c) => c.includes("roomtype"))
              : normedRow.findIndex((c) => c.includes("타입"));
          if (roomNumIdx !== -1 && qtyIdx !== -1 && typeIdx !== -1) {
            const maxOccIdx =
              normedRow.findIndex((c) => c.includes("최대인원")) !== -1
                ? normedRow.findIndex((c) => c.includes("최대인원"))
                : normedKeyRow.findIndex((c) => c.includes("maxocc"));
            rows = candidateRows;
            headerIdx = i;
            colIdx = { type: typeIdx, qty: qtyIdx, roomNumber: roomNumIdx, maxOcc: maxOccIdx };
            format = "B";
            break;
          }
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
          `헤더 행("룸타입"+"구성"+"호수", 또는 "Room type"+"Q'ty"+"Room number")을 찾지 못했어요.\n\n실제로 읽은 내용:\n${detail}`
        );
        return;
      }

      const newRoomTypes = [];
      const newCategories = new Set(categories);
      const newIrregular = new Set(irregularOptions);
      const floorSet = new Set(floors);
      let facilityCount = 0;
      const roomNumberShortfalls = []; // "수량"보다 "호수" 목록이 모자란 룸타입 경고용

      if (format === "B") {
        // 설계팀 룸믹스: 행 하나 = 룸타입 하나. 침대타입/욕조유무/객실등급은 라벨에서 추측하지 않고
        // 기본값(퀸/무/Superior)으로 두고 사용자가 직접 보완하는 방식으로 확정함.
        for (let i = headerIdx + 1; i < rows.length; i++) {
          const row = rows[i];
          const label = String(row[colIdx.type] || "").trim();
          if (!label) continue;
          if (norm(label).includes("합계") || norm(label).includes("총계")) break;
          const qtyRaw = String(row[colIdx.qty] ?? "").trim();
          const qtyNum = qtyRaw === "" ? null : Number(qtyRaw.replace(/[^0-9.-]/g, ""));
          if (qtyNum !== null && !Number.isNaN(qtyNum) && qtyNum === 0) continue; // 수량 0(운용 종료 등)이면 호수가 남아있어도 스킵
          const roomNumStr = String(row[colIdx.roomNumber] || "").trim();
          if (!roomNumStr) continue; // 호수 없는 행(부대시설 등)은 건너뜀
          const roomNumbers = parseRoomNumberTokens(roomNumStr);
          if (roomNumbers.length === 0) continue;
          // "수량"이 진짜 기준값 — 호수 목록에 오타/누락으로 개수가 모자라면(원본 파일 자체의 흔한 실수)
          // 조용히 객실수가 줄어들지 않도록 부족한 만큼 "미기재" 자리를 채워 총 객실수를 수량과 맞춘다
          if (qtyNum && qtyNum > 0 && roomNumbers.length < qtyNum) {
            const shortfall = qtyNum - roomNumbers.length;
            for (let m = 1; m <= shortfall; m++) roomNumbers.push(`미기재-${m}`);
            roomNumberShortfalls.push(`${label}(${shortfall}개)`);
          }
          const maxOcc = colIdx.maxOcc !== -1 ? String(row[colIdx.maxOcc] || "").trim() : "";
          const maxOccNum = maxOcc === "" ? null : parseInt(maxOcc.replace(/[^0-9]/g, ""), 10);

          newCategories.add(label);

          const byFloor = {};
          roomNumbers.forEach((num) => {
            const isPlaceholder = String(num).startsWith("미기재");
            const f = isPlaceholder ? "미배치" : floorFromRoomNumber(num);
            if (!f) return;
            floorSet.add(f);
            byFloor[f] = (byFloor[f] || 0) + 1;
          });

          newRoomTypes.push({
            id: nextId(),
            bed: "퀸",
            bathtub: "무",
            category: label,
            irregular: [],
            mattressQty: 1,
            capacity: maxOccNum && maxOccNum > 0 ? maxOccNum : 2,
            bedComposition: [], // Format B는 침대구성 정보가 없어 추후 수동 보완 필요(기본값 퀸1로 계산됨)
            grade: "Superior",
            features: [],
            view: "",
            includeBedInName: false,
            includeViewInName: false,
            customName: label,
            otaBedCount: "",
            otaBedSize: "",
            otaMaxOccupancy: maxOcc,
            otaFacilities: "",
            roomNumbers,
            sourceLabel: label,
            byFloor,
          });
        }

        if (newRoomTypes.length === 0) {
          setImportError('"Room number" 열에 호수가 채워진 행을 찾지 못했어요.');
          return;
        }
      } else {
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
          const roomNumbers = parseRoomNumberTokens(hoStr);
          const capRaw = colIdx["인원"] !== undefined ? String(row[colIdx["인원"]] ?? "").trim() : "";
          const capNum = capRaw === "" ? null : parseInt(capRaw.replace(/[^0-9]/g, ""), 10);

          const { category, paren } = parseRoomTypeLabel(label);
          const { bed, mattressQty, bathtub, extraIrregular, bedComposition } = parseComposition(composition, paren);

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
            capacity: capNum && capNum > 0 ? capNum : 2,
            bedComposition,
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
        (format === "B"
          ? `설계팀 룸믹스 형식으로 인식했어요. 룸타입 ${newRoomTypes.length}개, 호수 ${totalRooms}개를 가져왔어요. 침대타입·욕조유무·객실등급은 기본값으로 채워졌으니 필요하면 직접 수정해주세요.`
          : `룸타입 ${newRoomTypes.length}개, 호수 ${totalRooms}개를 가져왔어요.` +
            (facilityCount > 0 ? ` (부대시설성 항목 ${facilityCount}개는 제외됨)` : "")) +
          (overwriteOnImport ? " 기존에 같은 룸타입이 있으면 덮어썼어요." : "") +
          (roomNumberShortfalls.length > 0
            ? ` ⚠️ 다음 룸타입은 "수량"보다 "호수" 목록이 모자라 부족한 만큼 "미기재-N"으로 채워 넣었어요(총 객실수는 정확함, 호수만 비어있음) — 층별 배치표에서 실제 호수로 직접 채워주세요: ${roomNumberShortfalls.join(", ")}`
            : "")
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
            { key: "test-schedule", label: "테스트 공정표" },
            { key: "vendors", label: "업체 관리" },
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

        {activeTab === "test-schedule" && (
          <TestScheduleDashboard onOpenInOrderManager={handleOpenInOrderManager} />
        )}

        {activeTab === "vendors" && (
          <VendorManagement
            vendors={vendors}
            loading={vendorsLoading}
            error={vendorsError}
            onReload={reloadVendors}
            onAdd={handleAddVendor}
            onFieldChange={handleVendorFieldChange}
            onFieldCommit={handleVendorFieldCommit}
            onRemove={handleRemoveVendor}
          />
        )}

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
              <label className="block text-xs text-slate-500 mb-1">목표 객실수</label>
              <input
                type="number"
                min="0"
                value={targetRoomCount}
                onChange={(e) => handleTargetRoomCountChange(e.target.value)}
                placeholder="예: 42"
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
              />
              <p className="text-[11px] text-slate-400 mt-1">
                아래 룸타입·층별 배치가 이 숫자를 목표로 채워지는지 실시간으로 보여줘요.
              </p>
            </div>
            <div>
              <label className="block text-xs text-slate-500 mb-1">실당 예산</label>
              <input
                type="number"
                min="0"
                value={budgetPerRoom}
                onChange={(e) => handleBudgetPerRoomChange(e.target.value)}
                placeholder="예: 3500000"
                className="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
              />
              <p className="text-[11px] text-slate-400 mt-1">
                목표 객실수와 함께 입력하면 아래 배정 예산이 자동 계산돼요.
              </p>
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
              {targetRoomCount > 0 && budgetPerRoom > 0 && (
                <p className="text-[11px] text-slate-400 mt-1">
                  목표 객실수 × 실당 예산으로 자동 계산됨 (직접 수정하면 그 값 그대로 유지돼요)
                </p>
              )}
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
          {targetRoomCount > 0 && (
            <div className="mt-5 pt-4 border-t border-slate-100">
              <div className="flex items-center justify-between mb-1.5">
                <span className="text-xs text-slate-500">
                  객실 배치 현황 — {grandTotal} / {targetRoomCount}실
                  {grandTotal < targetRoomCount && (
                    <span className="text-amber-600"> ({targetRoomCount - grandTotal}실 남음)</span>
                  )}
                  {grandTotal === targetRoomCount && (
                    <span className="text-emerald-600 font-medium"> 배치 완료</span>
                  )}
                  {grandTotal > targetRoomCount && (
                    <span className="text-rose-600 font-medium"> ({grandTotal - targetRoomCount}실 초과)</span>
                  )}
                </span>
                <span className="text-xs text-slate-400">
                  {targetRoomCount > 0 ? Math.round((grandTotal / targetRoomCount) * 100) : 0}%
                </span>
              </div>
              <div className="h-1.5 bg-slate-100 rounded-full overflow-hidden">
                <div
                  className={`h-full rounded-full ${
                    grandTotal > targetRoomCount
                      ? "bg-rose-500"
                      : grandTotal === targetRoomCount
                      ? "bg-emerald-500"
                      : "bg-amber-500"
                  }`}
                  style={{ width: `${Math.min(100, (grandTotal / targetRoomCount) * 100)}%` }}
                />
              </div>
            </div>
          )}
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
            <br />
            설계팀에서 넘어오는 "Room type / Q'ty / 최대인원 / Room number" 형식의 룸믹스 시트도 자동으로 인식해요 —
            이 경우 침대타입·욕조유무·객실등급은 기본값(퀸·무·Superior)으로 채워지니 필요하면 직접 수정해주세요.
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
            <div>
              <label className="block text-xs text-slate-500 mb-1">수용인원</label>
              <input
                type="number"
                min="1"
                value={draft.capacity}
                onChange={(e) =>
                  setDraft({ ...draft, capacity: Math.max(1, parseInt(e.target.value || "1", 10)) })
                }
                title="베개류 등 인당(capacity) 기준 발주 계산에 사용"
                className="w-20 border border-slate-300 rounded-lg px-3 py-1.5 text-sm"
              />
            </div>
            <div className="min-w-[200px]">
              <label className="block text-xs text-slate-500 mb-1">
                복합 침대구성 <span className="text-slate-400">(비워두면 위 침대타입/매트리스수량으로 자동 계산)</span>
              </label>
              <div className="flex flex-col gap-1">
                {draft.bedComposition.map((row, idx) => (
                  <div key={idx} className="flex items-center gap-1.5">
                    <select
                      value={row.size}
                      onChange={(e) => updateDraftBedCompRow(idx, "size", e.target.value)}
                      className="border border-slate-300 rounded-lg px-2 py-1 text-xs"
                    >
                      <option value="Q">퀸(Q)</option>
                      <option value="S">싱글(S)</option>
                    </select>
                    <input
                      type="number"
                      min="1"
                      value={row.qty}
                      onChange={(e) => updateDraftBedCompRow(idx, "qty", e.target.value)}
                      className="w-14 border border-slate-300 rounded-lg px-2 py-1 text-xs"
                    />
                    <button
                      onClick={() => removeDraftBedCompRow(idx)}
                      className="text-xs text-slate-400 hover:text-red-500 px-1"
                    >
                      ✕
                    </button>
                  </div>
                ))}
                <button
                  onClick={addDraftBedCompRow}
                  className="text-xs text-amber-700 hover:text-amber-800 self-start"
                >
                  + 침대 구성 추가 (예: 퀸+싱글)
                </button>
              </div>
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
                            <span className="flex items-center gap-1 text-xs text-slate-500 whitespace-nowrap">
                              인원
                              <input
                                type="number"
                                min="1"
                                value={rt.capacity ?? 2}
                                onChange={(e) => setCapacity(rt.id, e.target.value)}
                                title="수용인원 — 베개류 등 인당 기준 발주 계산에 사용"
                                className="w-12 text-center border border-slate-200 rounded-md py-0.5 text-xs focus:outline-none focus:ring-2 focus:ring-amber-500"
                              />
                            </span>
                          </div>
                          <details className="mt-1">
                            <summary className="text-[11px] text-amber-700 cursor-pointer select-none">
                              침대구성 ({effectiveBedComposition(rt).map((b) => `${b.size}×${b.qty}`).join(" + ")})
                            </summary>
                            <div className="mt-1 flex flex-col gap-1 max-w-[240px]">
                              {(rt.bedComposition || []).map((row, idx) => (
                                <div key={idx} className="flex items-center gap-1.5">
                                  <select
                                    value={row.size}
                                    onChange={(e) => updateBedCompRow(rt.id, idx, "size", e.target.value)}
                                    className="border border-slate-200 rounded-md px-1.5 py-0.5 text-[11px]"
                                  >
                                    <option value="Q">퀸(Q)</option>
                                    <option value="S">싱글(S)</option>
                                  </select>
                                  <input
                                    type="number"
                                    min="1"
                                    value={row.qty}
                                    onChange={(e) => updateBedCompRow(rt.id, idx, "qty", e.target.value)}
                                    className="w-12 border border-slate-200 rounded-md px-1.5 py-0.5 text-[11px]"
                                  />
                                  <button
                                    onClick={() => removeBedCompRow(rt.id, idx)}
                                    className="text-[11px] text-slate-400 hover:text-red-500 px-1"
                                  >
                                    ✕
                                  </button>
                                </div>
                              ))}
                              <button
                                onClick={() => addBedCompRow(rt.id)}
                                className="text-[11px] text-amber-700 hover:text-amber-800 self-start"
                              >
                                + 구성 직접 지정 (미지정 시 위 침대타입/매트리스수량 자동 사용)
                              </button>
                            </div>
                          </details>
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
                <span className="text-sm font-medium tracking-wide" title="이 카드 이름은 계산 방식(룸타입별)을 뜻할 뿐, 실제 FF&E/OS&E 구분은 각 품목의 '구분' 칸을 따름">룸타입별 품목</span>
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
            {presetError && (
              <p className="text-xs text-red-600 mb-2">{presetError}</p>
            )}
            <div className="space-y-6">
              {roomTypes.map((rt) => {
                const c = categoryColor(rt.category);
                const items = ffeItems[rt.id] || [];
                const roomCount = roomTypeTotal(rt);
                const typeTotal = items.reduce(
                  (sum, it) => sum + (it.unitPrice + (it.installUnitPrice || 0)) * ffeItemQty(it, rt),
                  0
                );
                const typeActualTotal = items.reduce(
                  (sum, it) => sum + ((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * ffeItemQty(it, rt),
                  0
                );
                return (
                  <div key={rt.id}>
                    <div className="flex items-center justify-between mb-2">
                      <div className="flex items-center gap-2">
                        <button
                          onClick={() => toggleRoomTypeCollapsed(rt.id)}
                          className="text-slate-400 hover:text-slate-700"
                          title={collapsedRoomTypeIds.has(rt.id) ? "펼치기" : "접기"}
                        >
                          {collapsedRoomTypeIds.has(rt.id) ? <ChevronRight size={16} /> : <ChevronDown size={16} />}
                        </button>
                        <span className={`text-[11px] font-medium px-2 py-0.5 rounded ${c.bg} ${c.text}`}>
                          {codeFor(rt)}
                        </span>
                        <span className="text-xs text-slate-500">
                          {generateRoomName(rt)} · 객실 {roomCount}실
                        </span>
                        {collapsedRoomTypeIds.has(rt.id) && items.length > 0 && (
                          <span className="text-xs text-slate-400">
                            — 예산 <span className="font-medium text-slate-600">{won(typeTotal)}</span>
                            {" / "}집행 <span className="font-medium text-teal-700">{won(typeActualTotal)}</span>
                            {" ("}{items.length}개 품목{")"}
                          </span>
                        )}
                      </div>
                      {!collapsedRoomTypeIds.has(rt.id) && (
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <button
                          onClick={() => addPresetFfeItems(rt.id)}
                          className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50"
                        >
                          기본 세트 추가
                        </button>
                        <button
                          onClick={() => openPresetPicker(rt)}
                          disabled={loadingPresetFor === rt.id}
                          title="Supabase 전사 표준 템플릿(공통베이스+룸타입 전용콘텐츠)을 room/capacity/bed 기준수량 그대로 불러옴"
                          className="text-xs border border-amber-300 text-amber-700 rounded-lg px-2.5 py-1 hover:bg-amber-50 disabled:opacity-50"
                        >
                          {loadingPresetFor === rt.id ? "불러오는 중..." : "카탈로그 기본세트 불러오기"}
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
                          onClick={() => addFfeItem(rt.id)}
                          className="text-xs border border-slate-300 rounded-lg px-2.5 py-1 hover:bg-slate-50 flex items-center gap-1"
                        >
                          <Plus size={13} /> 품목 추가
                        </button>
                      </div>
                      )}
                    </div>
                    {!collapsedRoomTypeIds.has(rt.id) && (
                    <>
                    {presetPickerFor === rt.id && (
                      <div className="mb-3 bg-amber-50 border border-amber-200 rounded-lg p-3">
                        <p className="text-[11px] text-slate-600 mb-2">
                          불러올 카테고리를 선택하세요 (기본은 전체 선택됨)
                        </p>
                        <div className="flex flex-wrap gap-2 mb-3">
                          {presetPickerCatCounts.map(([cat, count]) => (
                            <label
                              key={cat}
                              className="flex items-center gap-1.5 text-xs bg-white border border-slate-200 rounded-md px-2 py-1 cursor-pointer"
                            >
                              <input
                                type="checkbox"
                                checked={presetPickerSelectedCats.has(cat)}
                                onChange={() => togglePresetPickerCat(cat)}
                              />
                              {cat} ({count})
                            </label>
                          ))}
                        </div>
                        <div className="flex justify-end gap-2">
                          <button
                            onClick={cancelPresetPicker}
                            className="text-xs border border-slate-300 rounded-md px-3 py-1 hover:bg-white"
                          >
                            취소
                          </button>
                          <button
                            onClick={confirmPresetPicker}
                            disabled={presetPickerSelectedCats.size === 0}
                            className="text-xs bg-amber-700 text-white px-3 py-1 rounded-md hover:bg-amber-800 disabled:opacity-50"
                          >
                            선택한 카테고리 불러오기
                          </button>
                        </div>
                      </div>
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
                            <th className="py-1.5 font-normal" title="회계 대분류 — 룸타입 카드 안에 있어도 실제로는 OS&E(린넨/타올 등)일 수 있음">구분</th>
                            <th className="py-1.5 font-normal text-right">공급예산단가</th>
                            <th className="py-1.5 font-normal text-right">공급집행단가</th>
                            <th className="py-1.5 font-normal text-right">설치예산단가</th>
                            <th className="py-1.5 font-normal text-right">설치집행단가</th>
                            <th className="py-1.5 font-normal text-right">기준수량</th>
                            <th className="py-1.5 font-normal text-right">배수</th>
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
                                <select
                                  value={it.categoryGroup || ""}
                                  onChange={(e) => updateFfeItem(rt.id, it.id, "categoryGroup", e.target.value)}
                                  title="대시보드 FF&E/OS&E 집계 기준. 카탈로그에서 불러온 품목은 자동 지정됨"
                                  className={`text-xs border rounded-md px-1.5 py-1 ${
                                    !it.categoryGroup
                                      ? "border-rose-300 text-rose-500 bg-rose-50"
                                      : "border-slate-200 text-slate-600"
                                  }`}
                                >
                                  <option value="">미지정</option>
                                  <option value="FF&E">FF&E</option>
                                  <option value="OS&E">OS&E</option>
                                </select>
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
                                <div className="flex flex-col items-end gap-0.5">
                                  <input
                                    type="number"
                                    min="0"
                                    step="0.1"
                                    value={it.qtyPerRoom}
                                    onChange={(e) => updateFfeItem(rt.id, it.id, "qtyPerRoom", e.target.value)}
                                    className="w-16 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                                  />
                                  <select
                                    value={it.calcBasis || "room"}
                                    onChange={(e) => updateFfeItem(rt.id, it.id, "calcBasis", e.target.value)}
                                    title="기준수량 계산 기준"
                                    className="w-16 border border-slate-200 rounded-md text-[10px] text-slate-500 focus:outline-none"
                                  >
                                    <option value="room">룸당</option>
                                    <option value="capacity">인당</option>
                                    <option value="bed">침대당</option>
                                  </select>
                                </div>
                              </td>
                              <td className="py-1.5">
                                <input
                                  type="number"
                                  min="0"
                                  step="0.1"
                                  value={it.multiplier != null ? it.multiplier : 1}
                                  onChange={(e) => updateFfeItem(rt.id, it.id, "multiplier", e.target.value)}
                                  title="초도발주 배수(예산에 따라 개별 조정 가능)"
                                  className="w-14 text-right border border-slate-200 rounded-md px-2 py-1 text-sm focus:outline-none focus:ring-2 focus:ring-amber-500"
                                />
                              </td>
                              <td className="py-1.5 text-right text-slate-500">
                                {ffeItemQty(it, rt).toLocaleString("ko-KR")}
                                {it.cartonSize ? (
                                  <span className="text-[10px] text-amber-600 block">카톤×{it.cartonSize}</span>
                                ) : null}
                              </td>
                              <td className="py-1.5 text-right font-medium">
                                {won((it.unitPrice + (it.installUnitPrice || 0)) * ffeItemQty(it, rt))}
                              </td>
                              <td className="py-1.5 text-right font-medium text-teal-700">
                                {won(((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * ffeItemQty(it, rt))}
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
                    </>
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
              <span className="text-sm font-medium tracking-wide" title="이 카드 이름은 계산 방식(룸타입 무관 공통)을 뜻할 뿐, 실제 FF&E/OS&E 구분은 각 품목의 '구분' 칸을 따름">공통 품목 (전체 {grandTotal}실 기준)</span>
            </div>
            <div className="flex items-center gap-1.5">
              <button
                onClick={openOsePresetPicker}
                disabled={loadingOsePreset}
                title="Supabase 전사 표준 템플릿 중 룸타입에 무관한 공통(room 기준) 품목을 불러옴 — 침구/베개 등 룸타입별 품목은 각 룸타입 카드에서 불러오세요"
                className="text-xs border border-amber-300 text-amber-700 rounded-lg px-2.5 py-1 hover:bg-amber-50 disabled:opacity-50"
              >
                {loadingOsePreset ? "불러오는 중..." : "카탈로그 기본세트 불러오기"}
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
          {osePresetError && (
            <p className="text-xs text-red-600 mb-2">{osePresetError}</p>
          )}
          {presetPickerFor === "OSE" && (
            <div className="mb-3 bg-amber-50 border border-amber-200 rounded-lg p-3">
              <p className="text-[11px] text-slate-600 mb-2">
                불러올 카테고리를 선택하세요 (기본은 전체 선택됨)
              </p>
              <div className="flex flex-wrap gap-2 mb-3">
                {presetPickerCatCounts.map(([cat, count]) => (
                  <label
                    key={cat}
                    className="flex items-center gap-1.5 text-xs bg-white border border-slate-200 rounded-md px-2 py-1 cursor-pointer"
                  >
                    <input
                      type="checkbox"
                      checked={presetPickerSelectedCats.has(cat)}
                      onChange={() => togglePresetPickerCat(cat)}
                    />
                    {cat} ({count})
                  </label>
                ))}
              </div>
              <div className="flex justify-end gap-2">
                <button
                  onClick={cancelPresetPicker}
                  className="text-xs border border-slate-300 rounded-md px-3 py-1 hover:bg-white"
                >
                  취소
                </button>
                <button
                  onClick={confirmPresetPicker}
                  disabled={presetPickerSelectedCats.size === 0}
                  className="text-xs bg-amber-700 text-white px-3 py-1 rounded-md hover:bg-amber-800 disabled:opacity-50"
                >
                  선택한 카테고리 불러오기
                </button>
              </div>
            </div>
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
                  <th className="py-1.5 font-normal" title="회계 대분류 — '공통 품목' 카드에 있어도 실제로는 FF&E일 수 있음(예: 드라이기)">구분</th>
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
                      <select
                        value={it.categoryGroup || ""}
                        onChange={(e) => updateOseItem(it.id, "categoryGroup", e.target.value)}
                        title="대시보드 FF&E/OS&E 집계 기준. 카탈로그에서 불러온 품목은 자동 지정됨"
                        className={`text-xs border rounded-md px-1.5 py-1 ${
                          !it.categoryGroup
                            ? "border-rose-300 text-rose-500 bg-rose-50"
                            : "border-slate-200 text-slate-600"
                        }`}
                      >
                        <option value="">미지정</option>
                        <option value="FF&E">FF&E</option>
                        <option value="OS&E">OS&E</option>
                      </select>
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
                      {oseItemQty(it).toLocaleString("ko-KR")}
                      {it.cartonSize ? (
                        <span className="text-[10px] text-amber-600 block">카톤×{it.cartonSize}</span>
                      ) : null}
                    </td>
                    <td className="py-1.5 text-right font-medium">
                      {won((it.unitPrice + (it.installUnitPrice || 0)) * oseItemQty(it))}
                    </td>
                    <td className="py-1.5 text-right font-medium text-teal-700">
                      {won(((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * oseItemQty(it))}
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
        <LaborExpenseSection
          items={laborExpenses}
          onAdd={addLaborExpense}
          onAddDefaults={addDefaultLaborRoles}
          onUpdate={updateLaborExpense}
          onRemove={removeLaborExpense}
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
                  // FF&E/OS&E 구분은 카드 위치(룸타입별 vs 공통)가 아니라 품목별 categoryGroup 기준
                  let ffeTotal = 0;
                  let oseTotal = 0;
                  let unclassifiedTotal = 0;
                  roomTypes.forEach((rt) => {
                    const items = ffeItems[rt.id] || [];
                    items.forEach((it) => {
                      const amt = (it.unitPrice + (it.installUnitPrice || 0)) * ffeItemQty(it, rt);
                      if (it.categoryGroup === "FF&E") ffeTotal += amt;
                      else if (it.categoryGroup === "OS&E") oseTotal += amt;
                      else unclassifiedTotal += amt;
                    });
                  });
                  oseItems.forEach((it) => {
                    const amt = (it.unitPrice + (it.installUnitPrice || 0)) * oseItemQty(it);
                    if (it.categoryGroup === "FF&E") ffeTotal += amt;
                    else if (it.categoryGroup === "OS&E") oseTotal += amt;
                    else unclassifiedTotal += amt;
                  });
                  return (
                    <div className="space-y-1">
                      <div className="text-xs text-slate-500">
                        FF&E {won(ffeTotal)} + OS&E {won(oseTotal)}
                        {unclassifiedTotal > 0 && (
                          <span className="text-rose-500"> + 미지정 {won(unclassifiedTotal)}</span>
                        )}
                      </div>
                      <div className="text-lg font-semibold text-slate-800">
                        {won(ffeTotal + oseTotal + unclassifiedTotal)}
                      </div>
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
            // qtyFn(it): 품목별 실제 발주수량을 반환 — FF&E는 ffeItemQty(room/capacity/bed 3축), OS&E는 프로젝트 전체 객실수 비례(단순 곱)
            const sumBudget = (items, qtyFn) =>
              items.reduce((s, it) => s + (it.unitPrice + (it.installUnitPrice || 0)) * qtyFn(it), 0);
            const sumActual = (items, qtyFn) =>
              items.reduce((s, it) => s + ((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * qtyFn(it), 0);

            // FF&E/OS&E 대분류 집계는 "룸타입 카드냐 공통 리스트냐"가 아니라 품목별 categoryGroup 기준으로 나눔
            // (침구/타올처럼 계산은 룸타입별(bed/capacity축)이어도 실제 회계분류는 OS&E인 품목이 섞여있기 때문)
            const sumBudgetByGroup = (items, qtyFn, group) =>
              items
                .filter((it) => it.categoryGroup === group)
                .reduce((s, it) => s + (it.unitPrice + (it.installUnitPrice || 0)) * qtyFn(it), 0);
            const sumActualByGroup = (items, qtyFn, group) =>
              items
                .filter((it) => it.categoryGroup === group)
                .reduce((s, it) => s + ((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * qtyFn(it), 0);

            let ffeBudget = 0;
            let ffeActual = 0;
            let oseBudget = 0;
            let oseActual = 0;
            let unclassifiedBudget = 0;
            let unclassifiedActual = 0;
            roomTypes.forEach((rt) => {
              const items = ffeItems[rt.id] || [];
              const qtyFn = (it) => ffeItemQty(it, rt);
              ffeBudget += sumBudgetByGroup(items, qtyFn, "FF&E");
              ffeActual += sumActualByGroup(items, qtyFn, "FF&E");
              oseBudget += sumBudgetByGroup(items, qtyFn, "OS&E");
              oseActual += sumActualByGroup(items, qtyFn, "OS&E");
              const unclassified = items.filter((it) => it.categoryGroup !== "FF&E" && it.categoryGroup !== "OS&E");
              unclassifiedBudget += sumBudget(unclassified, qtyFn);
              unclassifiedActual += sumActual(unclassified, qtyFn);
            });
            {
              const qtyFn = (it) => oseItemQty(it);
              ffeBudget += sumBudgetByGroup(oseItems, qtyFn, "FF&E");
              ffeActual += sumActualByGroup(oseItems, qtyFn, "FF&E");
              oseBudget += sumBudgetByGroup(oseItems, qtyFn, "OS&E");
              oseActual += sumActualByGroup(oseItems, qtyFn, "OS&E");
              const unclassified = oseItems.filter((it) => it.categoryGroup !== "FF&E" && it.categoryGroup !== "OS&E");
              unclassifiedBudget += sumBudget(unclassified, qtyFn);
              unclassifiedActual += sumActual(unclassified, qtyFn);
            }
            const siteBudget = siteExpenses.reduce((s, it) => s + (it.budgetAmount || 0), 0);
            const siteActual = siteExpenses.reduce((s, it) => s + (it.actualAmount || 0), 0);
            const laborBudget = laborExpenses.reduce((s, it) => s + (it.budgetAmount || 0), 0);
            const laborActual = laborExpenses.reduce((s, it) => s + (it.actualAmount || 0), 0);
            const extraBudget = extraExpenses.reduce((s, it) => s + (it.budgetAmount || 0), 0);
            const extraActual = extraExpenses.reduce((s, it) => s + (it.actualAmount || 0), 0);

            const plannedTotal =
              ffeBudget + oseBudget + unclassifiedBudget + siteBudget + laborBudget + extraBudget;
            const actualTotal =
              ffeActual + oseActual + unclassifiedActual + siteActual + laborActual + extraActual;
            const perRoomBudget = grandTotal > 0 ? totalBudget / grandTotal : 0;
            const perRoomPlanned = grandTotal > 0 ? plannedTotal / grandTotal : 0;
            const remaining = totalBudget - actualTotal; // 잔여비 (총예산 - 실사용비)
            const planVsActual = plannedTotal - actualTotal; // 계획 대비 집행 차이

            const categoryRows = [
              { name: "FF&E", budget: ffeBudget, actual: ffeActual },
              { name: "OS&E", budget: oseBudget, actual: oseActual },
              ...(unclassifiedBudget > 0 || unclassifiedActual > 0
                ? [{ name: "미지정(FF&E/OS&E 구분 필요)", budget: unclassifiedBudget, actual: unclassifiedActual }]
                : []),
              { name: "현장지출", budget: siteBudget, actual: siteActual },
              { name: "인건비", budget: laborBudget, actual: laborActual },
              { name: "예산외 지출", budget: extraBudget, actual: extraActual },
            ];

            const ffeByRoomType = roomTypes
              .map((rt) => {
                const items = ffeItems[rt.id] || [];
                return {
                  name: codeFor(rt),
                  예산: sumBudget(items, (it) => ffeItemQty(it, rt)),
                  집행: sumActual(items, (it) => ffeItemQty(it, rt)),
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
                cur.budget += (it.unitPrice + (it.installUnitPrice || 0)) * ffeItemQty(it, rt);
                cur.actual += ((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * ffeItemQty(it, rt);
                itemAgg.set(it.name, cur);
              });
            });
            oseItems.forEach((it) => {
              if (!it.name) return;
              const cur = itemAgg.get(it.name) || { name: it.name, budget: 0, actual: 0 };
              cur.budget += (it.unitPrice + (it.installUnitPrice || 0)) * oseItemQty(it);
              cur.actual += ((it.actualUnitPrice || 0) + (it.installActualUnitPrice || 0)) * oseItemQty(it);
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
                    <p className="text-xs text-slate-500 mb-2" title="FF&E/OS&E 구분 없이 룸타입 카드 전체 합계">룸타입별 품목 예산 대비 집행</p>
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
