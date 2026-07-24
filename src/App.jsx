import { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import { Plus, X, Building2, LayoutGrid, Table2, Trash2, Upload, Save, Users, Loader2, LayoutDashboard } from "lucide-react";
import {
  Tooltip, Legend, ResponsiveContainer,
  BarChart, Bar, XAxis, YAxis, CartesianGrid,
} from "recharts";

import { getProjectIndex, getProjectData, saveProjectData, saveProjectIndex, getCatalog, saveCatalog } from "./api";

const TIERS = ["Flagship", "Premium", "Upper Select", "Select", "Essential"];
const BED_TYPES = ["싱글", "퀸"];
const BATHTUB = ["유", "무"];

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

export default function App() {
  const [projectName, setProjectName] = useState("");
  const [tier, setTier] = useState(TIERS[2]);
  const [totalBudget, setTotalBudget] = useState(0); // 전체 공사비 중 오픈바이징팀에 배정된 예산

  const [categories, setCategories] = useState(["트윈", "더블", "패밀리", "스위트"]);
  const [newCategory, setNewCategory] = useState("");

  const [irregularOptions, setIrregularOptions] = useState(["마사지체어", "발코니", "복층"]);
  const [newIrregular, setNewIrregular] = useState("");

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
        { id: nextId(), name: "", unitPrice: 0, actualUnitPrice: 0, qtyPerRoom: 1 },
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
        ...basicPreset.map((p) => ({ id: nextId(), name: p.name, unitPrice: 0, actualUnitPrice: 0, qtyPerRoom: 1 })),
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
    setOseItems((prev) => [...prev, { id: nextId(), name: "", unitPrice: 0, actualUnitPrice: 0, qtyPerRoom: 1 }]);
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
  });

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
        JSON.stringify([...rt.irregular].sort()) === JSON.stringify([...draft.irregular].sort())
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
        roomNumbers: [],
        byFloor: Object.fromEntries(floors.map((f) => [f, 0])),
      },
    ]);
    setDraft((d) => ({ ...d, irregular: [], mattressQty: 1 }));
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
    if (data.categories) setCategories(data.categories);
    if (data.irregularOptions) setIrregularOptions(data.irregularOptions);
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

  async function saveProject() {
    if (!currentProjectId) return;
    setIsSaving(true);
    setSaveError("");
    try {
      const payload = {
        projectName,
        tier,
        totalBudget,
        categories,
        irregularOptions,
        floors,
        roomTypes,
        ffeItems,
        oseItems,
        siteExpenses,
        laborExpenses,
        extraExpenses,
        basicPreset,
        savedAt: new Date().toISOString(),
      };
      const result = await saveProjectData(currentProjectId, payload);
      if (!result || result.error) throw new Error(result && result.error);
      setLastSaved(payload.savedAt);
      const updatedList = projectList.map((p) => (p.id === currentProjectId ? { ...p, name: projectName || p.name } : p));
      setProjectList(updatedList);
      await saveProjectIndex(updatedList);
    } catch (err) {
      setSaveError("저장에 실패했어요. 잠시 후 다시 시도해주세요.");
    } finally {
      setIsSaving(false);
    }
  }

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
        ...chosen.map((it) => ({ id: nextId(), name: it.name, unitPrice: it.unitPrice, actualUnitPrice: 0, qtyPerRoom: 1 })),
      ]);
    } else {
      setFfeItems((prev) => ({
        ...prev,
        [pickerOpenFor]: [
          ...(prev[pickerOpenFor] || []),
          ...chosen.map((it) => ({ id: nextId(), name: it.name, unitPrice: it.unitPrice, actualUnitPrice: 0, qtyPerRoom: 1 })),
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
            <button
              onClick={addRoomType}
              className="bg-amber-700 text-white text-sm px-4 py-2 rounded-lg hover:bg-amber-800 flex items-center gap-1.5"
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
                                  title={`${codeFor(rt)} · ${rt.category}`}
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
                  (sum, it) => sum + it.unitPrice * it.qtyPerRoom * roomCount,
                  0
                );
                const typeActualTotal = items.reduce(
                  (sum, it) => sum + (it.actualUnitPrice || 0) * it.qtyPerRoom * roomCount,
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
                          {rt.category} · 객실 {roomCount}실
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
                            <th className="py-1.5 font-normal text-right">예산단가</th>
                            <th className="py-1.5 font-normal text-right">집행단가</th>
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
                                {won(it.unitPrice * it.qtyPerRoom * roomCount)}
                              </td>
                              <td className="py-1.5 text-right font-medium text-teal-700">
                                {won((it.actualUnitPrice || 0) * it.qtyPerRoom * roomCount)}
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
                  <th className="py-1.5 font-normal text-right">예산단가</th>
                  <th className="py-1.5 font-normal text-right">집행단가</th>
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
                      {won(it.unitPrice * it.qtyPerRoom * grandTotal)}
                    </td>
                    <td className="py-1.5 text-right font-medium text-teal-700">
                      {won((it.actualUnitPrice || 0) * it.qtyPerRoom * grandTotal)}
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
                    return sum + items.reduce((s, it) => s + it.unitPrice * it.qtyPerRoom * roomCount, 0);
                  }, 0);
                  const oseTotal = oseItems.reduce(
                    (sum, it) => sum + it.unitPrice * it.qtyPerRoom * grandTotal,
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
              items.reduce((s, it) => s + it.unitPrice * it.qtyPerRoom * roomCountFn(it), 0);
            const sumActual = (items, roomCountFn) =>
              items.reduce((s, it) => s + (it.actualUnitPrice || 0) * it.qtyPerRoom * roomCountFn(it), 0);

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
                cur.budget += it.unitPrice * it.qtyPerRoom * roomCount;
                cur.actual += (it.actualUnitPrice || 0) * it.qtyPerRoom * roomCount;
                itemAgg.set(it.name, cur);
              });
            });
            oseItems.forEach((it) => {
              if (!it.name) return;
              const cur = itemAgg.get(it.name) || { name: it.name, budget: 0, actual: 0 };
              cur.budget += it.unitPrice * it.qtyPerRoom * grandTotal;
              cur.actual += (it.actualUnitPrice || 0) * it.qtyPerRoom * grandTotal;
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
      </div>
    </div>
  );
}
