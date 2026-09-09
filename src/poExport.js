// 발주서(PO) 엑셀 생성 — 브라우저에서 ExcelJS로 직접 생성해 다운로드.
// purchase-order-generator 스킬의 generate_po.py(디자인 확정판)를 그대로 포팅한 것.
// PDF는 만들지 않는다 — 사용자가 엑셀에서 열어 직접 PDF로 저장(A4·너비맞춤·반복헤더는 파일에 심어둠).
//
// 사용:
//   import { exportPurchaseOrders } from "./poExport";
//   await exportPurchaseOrders({ projectName, orderDate, deliveryDate, company, groups });
//   groups: [{ vendorName, vendorContactName, vendorContactPhone, items:[{name, brand, spec, unit, qty, price, note}], remark }]
//   업체가 1곳이면 xlsx 1개, 여러 곳이면 zip으로 묶어 내려받는다.

import ExcelJS from "exceljs";
import JSZip from "jszip";

const FONT = "Arial";
const ACCENT = "FF4472C4";      // 헤더/포인트
const ACCENT_DARK = "FF1F3864"; // 부제목/총합계 강조
const ACCENT_TINT = "FFD9E2F3"; // 총합계 배경
const BAND_FILL = "FFF2F2F2";   // 짝수행 밴딩
const WHITE = "FFFFFFFF";
const GREY = "FF808080";

// 페이지 분리 휴리스틱 (generate_po.py와 동일 상수 — 실제 인쇄 보고 튜닝)
const ROWS_PER_PAGE = 60;         // A4 세로·너비맞춤 축소 포함 실측치(2026-09-09, 기본 행높이 기준)
const SUMMARY_BLOCK_ROWS = 11;      // 합계3줄 + 발주조건(제목+4줄) + 사이 여백
const SUMMARY_MIN_GAP = 2;          // 품목 표와 합계 블록 사이 최소 빈 줄
const SUMMARY_BOTTOM_MARGIN = 3;    // 페이지 맨 아래에서 띄울 줄 수 (합계가 다음 장으로 밀리면 키울 것)

export const COMPANY_NAME = "(주)스페이스플래닝";
const LOGO_URL = `${import.meta.env.BASE_URL}po_logo.png`; // public/po_logo.png (SPACE PLANNING 로고, 약 7.5:1)

const ORDER_TERMS = [
  "상기 공급가액은 부가세 별도 금액이며, 부가세를 포함한 총 합계 기준으로 청구됩니다.",
  "납품 장소 및 시간은 현장 담당자와 사전 협의 후 확정합니다.",
  "검수 후 하자 발생 시 즉시 교체 또는 반품 처리합니다.",
  "수량 및 단가 변경 시 반드시 사전 협의 후 진행합니다.",
];

const thin = { style: "thin", color: { argb: "FF000000" } };
const BORDER = { top: thin, left: thin, bottom: thin, right: thin };

function font(opts = {}) {
  return { name: FONT, size: 10, ...opts };
}
function fill(argb) {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}
function remainingRowsOnPage(rowNum) {
  const pos = ((rowNum - 1) % ROWS_PER_PAGE) + 1;
  return ROWS_PER_PAGE - pos;
}

let logoCache = null;
async function fetchLogo() {
  if (logoCache !== null) return logoCache;
  try {
    const res = await fetch(LOGO_URL);
    if (!res.ok) throw new Error(String(res.status));
    logoCache = await res.arrayBuffer();
  } catch {
    logoCache = false; // 로고 없으면 회사명 텍스트로 대체
  }
  return logoCache;
}

// 업체 1곳 분량의 발주서 워크북 생성 → ArrayBuffer(xlsx)
export async function buildPoXlsx({
  projectName,
  orderDate,
  deliveryDate,
  companyName = COMPANY_NAME,
  companyContactName,
  companyContactPhone,
  vendorName,
  vendorContactName,
  vendorContactPhone,
  items,
  remark = "",
}) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("발주서", {
    pageSetup: {
      paperSize: 9, // A4
      orientation: "portrait",
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 },
    },
    headerFooter: {
      oddFooter: `&L&"Arial,Regular"&8발주서 · ${projectName} · 발주일자 ${orderDate}&R&"Arial,Regular"&8${companyName}`,
    },
  });

  const widths = [5, 20, 14, 13, 6, 6, 11, 13, 16];
  widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));

  // ── 로고 + 타이틀 ──
  ws.getRow(1).height = 34;
  ws.getRow(2).height = 14;
  ws.getRow(3).height = 5;

  const logo = await fetchLogo();
  if (logo) {
    const imgId = wb.addImage({ buffer: logo, extension: "png" });
    ws.addImage(imgId, { tl: { col: 0, row: 0 }, ext: { width: 260, height: 35 } });
  } else {
    ws.mergeCells("A1:D2");
    ws.getCell("A1").value = companyName;
    ws.getCell("A1").font = font({ size: 14, bold: true, color: { argb: ACCENT_DARK } });
    ws.getCell("A1").alignment = { horizontal: "left", vertical: "middle" };
  }

  ws.mergeCells("F1:I1");
  ws.getCell("F1").value = "발    주    서";
  ws.getCell("F1").font = font({ size: 18, bold: true, color: { argb: ACCENT_DARK } });
  ws.getCell("F1").alignment = { horizontal: "right", vertical: "middle" };

  ws.mergeCells("F2:I2");
  ws.getCell("F2").value = "P U R C H A S E   O R D E R";
  ws.getCell("F2").font = font({ size: 9, color: { argb: GREY } });
  ws.getCell("F2").alignment = { horizontal: "right", vertical: "middle" };

  ws.mergeCells("A3:I3");
  for (let c = 1; c <= 9; c++) ws.getCell(3, c).fill = fill(ACCENT);

  // ── 발주일자 / 납품요청일 ──
  let r = 5;
  ws.mergeCells(`A${r}:D${r}`);
  ws.getCell(`A${r}`).value = `발주일자 :  ${orderDate}`;
  ws.getCell(`A${r}`).font = font({ bold: true });
  ws.mergeCells(`F${r}:I${r}`);
  ws.getCell(`F${r}`).value = `납품요청일 :  ${deliveryDate}`;
  ws.getCell(`F${r}`).font = font({ bold: true });

  // ── 발주처 / 수신 업체 ──
  r += 2;
  ws.mergeCells(`A${r}:D${r}`);
  ws.getCell(`A${r}`).value = "[발주처]";
  ws.getCell(`A${r}`).font = font({ bold: true, color: { argb: ACCENT_DARK } });
  ws.mergeCells(`F${r}:I${r}`);
  ws.getCell(`F${r}`).value = "[수신 업체]";
  ws.getCell(`F${r}`).font = font({ bold: true, color: { argb: ACCENT_DARK } });
  r += 1;

  const infoRows = [
    ["회사명", companyName, "업체명", vendorName],
    ["담당자", companyContactName, "담당자", vendorContactName],
    ["연락처", companyContactPhone, "연락처", vendorContactPhone],
  ];
  infoRows.forEach(([aL, aV, bL, bV]) => {
    ws.getCell(`A${r}`).value = `${aL} :`;
    ws.getCell(`A${r}`).font = font({ bold: true });
    ws.mergeCells(`B${r}:D${r}`);
    ws.getCell(`B${r}`).value = ` ${aV || ""}`;
    ws.getCell(`B${r}`).font = font();
    ws.getCell(`F${r}`).value = `${bL} :`;
    ws.getCell(`F${r}`).font = font({ bold: true });
    ws.mergeCells(`G${r}:I${r}`);
    ws.getCell(`G${r}`).value = ` ${bV || ""}`;
    ws.getCell(`G${r}`).font = font();
    r += 1;
  });

  r += 1;
  ws.mergeCells(`A${r}:D${r}`);
  ws.getCell(`A${r}`).value = `프로젝트/현장 :  ${projectName}`;
  ws.getCell(`A${r}`).font = font({ bold: true });
  ws.mergeCells(`F${r}:I${r}`);
  ws.getCell(`F${r}`).value = `품목 ${items.length}건`;
  ws.getCell(`F${r}`).font = font({ bold: true });

  // ── 품목 표 ──
  r += 2;
  const headerRow = r;
  ["No", "품목명", "브랜드/제조사", "규격", "단위", "수량", "단가", "공급가액", "비고"].forEach((h, i) => {
    const c = ws.getCell(headerRow, i + 1);
    c.value = h;
    c.font = font({ bold: true, color: { argb: WHITE } });
    c.fill = fill(ACCENT);
    c.alignment = { horizontal: "center", vertical: "middle" };
    c.border = BORDER;
  });

  const firstItemRow = headerRow + 1;
  let row = firstItemRow;
  items.forEach((it, idx) => {
    const band = (idx + 1) % 2 === 0;
    const vals = [idx + 1, it.name, it.brand || "", it.spec || "", it.unit || "EA", it.qty, it.price, { formula: `F${row}*G${row}` }, it.note || ""];
    vals.forEach((v, i) => {
      const c = ws.getCell(row, i + 1);
      c.value = v;
      c.font = font();
      c.border = BORDER;
      if (band) c.fill = fill(BAND_FILL);
      if ([1, 3, 4, 5, 6].includes(i + 1)) c.alignment = { horizontal: "center" };
      if (i + 1 === 7 || i + 1 === 8) c.numFmt = "#,##0";
    });
    row += 1;
  });
  const lastItemRow = row - 1;

  // 합계 블록 위치: 현재 페이지에 들어가면 페이지 "하단"에 붙이고(품목 표와 합계 사이가 비어도 됨),
  // 안 들어가면 다음 페이지로 넘겨 위쪽에 둔다. 페이지 높이 추정은 ROWS_PER_PAGE 휴리스틱이라
  // 실제 인쇄에서 합계가 다음 장으로 밀리면 SUMMARY_BOTTOM_MARGIN을 키우면 된다.
  const remaining = remainingRowsOnPage(lastItemRow);
  const blockRows = SUMMARY_BLOCK_ROWS + (remark ? 2 : 0);
  if (remaining < blockRows + SUMMARY_MIN_GAP) {
    ws.getRow(lastItemRow).addPageBreak();
    r = lastItemRow + 1 + SUMMARY_MIN_GAP;
  } else {
    r = lastItemRow + 1 + Math.max(SUMMARY_MIN_GAP, remaining - blockRows - SUMMARY_BOTTOM_MARGIN);
  }

  if (remark) {
    ws.mergeCells(`A${r}:I${r}`);
    ws.getCell(`A${r}`).value = `비고: ${remark}`;
    ws.getCell(`A${r}`).font = font({ italic: true, size: 9 });
    r += 2;
  } else {
    r += 1;
  }

  const sumRow = r;
  ws.mergeCells(`A${sumRow}:G${sumRow}`);
  ws.getCell(`A${sumRow}`).value = "공급가액 합계";
  ws.getCell(`A${sumRow}`).font = font({ bold: true });
  ws.getCell(`A${sumRow}`).alignment = { horizontal: "right" };
  let c = ws.getCell(sumRow, 8);
  c.value = { formula: `SUM(H${firstItemRow}:H${lastItemRow})` };
  c.numFmt = "#,##0";
  c.font = font({ bold: true });

  const vatRow = sumRow + 1;
  ws.mergeCells(`A${vatRow}:G${vatRow}`);
  ws.getCell(`A${vatRow}`).value = "부가세 (10%)";
  ws.getCell(`A${vatRow}`).font = font({ bold: true });
  ws.getCell(`A${vatRow}`).alignment = { horizontal: "right" };
  c = ws.getCell(vatRow, 8);
  c.value = { formula: `H${sumRow}*0.1` };
  c.numFmt = "#,##0";
  c.font = font();

  const totalRow = vatRow + 1;
  ws.mergeCells(`A${totalRow}:G${totalRow}`);
  ws.getCell(`A${totalRow}`).value = "총 합계 (부가세 포함)";
  ws.getCell(`A${totalRow}`).font = font({ bold: true, size: 12, color: { argb: ACCENT_DARK } });
  ws.getCell(`A${totalRow}`).alignment = { horizontal: "right" };
  c = ws.getCell(totalRow, 8);
  c.value = { formula: `H${sumRow}+H${vatRow}` };
  c.numFmt = "#,##0";
  c.font = font({ bold: true, size: 12, color: { argb: ACCENT_DARK } });
  c.fill = fill(ACCENT_TINT);
  const thick = { style: "medium", color: { argb: ACCENT } };
  for (let col = 1; col <= 9; col++) ws.getCell(totalRow, col).border = { top: thick, bottom: thick };

  // ── 발주 조건 ──
  r = totalRow + 2;
  ws.mergeCells(`A${r}:I${r}`);
  ws.getCell(`A${r}`).value = "발주 조건";
  ws.getCell(`A${r}`).font = font({ bold: true, color: { argb: ACCENT_DARK } });
  r += 1;
  ORDER_TERMS.forEach((t) => {
    ws.mergeCells(`A${r}:I${r}`);
    ws.getCell(`A${r}`).value = `· ${t}`;
    ws.getCell(`A${r}`).font = font({ size: 9 });
    r += 1;
  });
  const lastRow = r - 1;

  ws.pageSetup.printArea = `A1:I${lastRow}`;
  ws.pageSetup.printTitlesRow = `${headerRow}:${headerRow}`;

  return wb.xlsx.writeBuffer();
}

function sanitize(name) {
  return String(name || "").replace(/[\\/:*?"<>|]/g, "").replace(/\s+/g, "");
}
function todayCompact(dateStr) {
  return String(dateStr || "").replace(/-/g, "");
}
function triggerDownload(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// 업체별 발주서 일괄 생성 → 1곳이면 xlsx, 2곳 이상이면 zip 다운로드
export async function exportPurchaseOrders({ projectName, orderDate, deliveryDate, company, groups }) {
  const stamp = todayCompact(orderDate);
  const proj = sanitize(projectName);
  const files = [];
  for (const g of groups) {
    const buf = await buildPoXlsx({
      projectName,
      orderDate,
      deliveryDate,
      companyName: COMPANY_NAME,
      companyContactName: company.name,
      companyContactPhone: company.phone,
      vendorName: g.vendorName,
      vendorContactName: g.vendorContactName || "(연락처 미등록)",
      vendorContactPhone: g.vendorContactPhone || "(연락처 미등록)",
      items: g.items,
      remark: g.remark || "",
    });
    files.push({ name: `${proj}_${sanitize(g.vendorName)}_발주서_${stamp}.xlsx`, buf });
  }

  const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  if (files.length === 1) {
    triggerDownload(new Blob([files[0].buf], { type: XLSX_MIME }), files[0].name);
    return { count: 1, filenames: [files[0].name] };
  }
  const zip = new JSZip();
  files.forEach((f) => zip.file(f.name, f.buf));
  const zipBlob = await zip.generateAsync({ type: "blob" });
  const zipName = `${proj}_발주서_${stamp}.zip`;
  triggerDownload(zipBlob, zipName);
  return { count: files.length, filenames: files.map((f) => f.name), zipName };
}
