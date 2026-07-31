"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type PriceColumn = {
  column: number;
  headerRow: number;
  header: string;
};

type PreviewRow = {
  row: number;
  before: number;
  after: number;
};

const PRICE_HEADER_PRIORITY = [
  "판매가격",
  "판매가",
  "상품판매가",
  "최종판매가",
  "할인가",
  "판매금액",
  "상품가격",
  "가격",
];

function normalizeHeader(value: unknown) {
  return String(value ?? "")
    .replace(/[\s\n\r\t_-]/g, "")
    .replace(/[()\[\]{}]/g, "")
    .trim()
    .toLowerCase();
}

function toNumber(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;

  const cleaned = value.replace(/,/g, "").replace(/[^0-9.-]/g, "");
  if (!cleaned) return null;

  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

function countNumbersBelow(
  sheet: XLSX.WorkSheet,
  column: number,
  startRow: number,
  endRow: number,
) {
  let count = 0;
  const limit = Math.min(endRow, startRow + 2000);

  for (let row = startRow; row <= limit; row++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
    if (!cell || cell.f) continue;
    if (toNumber(cell.v) !== null) count++;
  }

  return count;
}

function findPriceColumn(sheet: XLSX.WorkSheet): PriceColumn | null {
  if (!sheet["!ref"]) return null;

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const lastHeaderRow = Math.min(range.e.r, range.s.r + 39);
  const candidates: Array<PriceColumn & { priority: number; numericCount: number }> = [];

  for (let row = range.s.r; row <= lastHeaderRow; row++) {
    for (let column = range.s.c; column <= range.e.c; column++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
      if (!cell) continue;

      const original = String(cell.v ?? "").trim();
      const normalized = normalizeHeader(original);
      if (!normalized) continue;

      let priority = PRICE_HEADER_PRIORITY.findIndex(
        (header) => normalized === normalizeHeader(header),
      );

      if (priority < 0 && normalized.includes("판매") && normalized.includes("가격")) {
        priority = 1;
      } else if (priority < 0 && normalized.includes("판매가")) {
        priority = 2;
      } else if (priority < 0 && normalized.endsWith("가격")) {
        priority = 7;
      }

      if (priority < 0) continue;

      const numericCount = countNumbersBelow(sheet, column, row + 1, range.e.r);
      if (numericCount === 0) continue;

      candidates.push({
        column,
        headerRow: row,
        header: original,
        priority,
        numericCount,
      });
    }
  }

  candidates.sort((a, b) => {
    if (a.priority !== b.priority) return a.priority - b.priority;
    if (a.numericCount !== b.numericCount) return b.numericCount - a.numericCount;
    return a.headerRow - b.headerRow;
  });

  return candidates[0] ?? null;
}

export default function Home() {
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState("");
  const [sheetName, setSheetName] = useState("");
  const [percent, setPercent] = useState(10);
  const [status, setStatus] = useState("엑셀 파일을 올려 주세요.");

  const sheet = workbook && sheetName ? workbook.Sheets[sheetName] : null;
  const priceColumn = useMemo(() => (sheet ? findPriceColumn(sheet) : null), [sheet]);

  function calculate(value: number) {
    return value * (1 + percent / 100);
  }

  const preview = useMemo<PreviewRow[]>(() => {
    if (!sheet?.["!ref"] || !priceColumn) return [];

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const rows: PreviewRow[] = [];

    for (let row = priceColumn.headerRow + 1; row <= range.e.r && rows.length < 20; row++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: priceColumn.column })];
      if (!cell || cell.f) continue;

      const before = toNumber(cell.v);
      if (before === null) continue;

      rows.push({ row: row + 1, before, after: calculate(before) });
    }

    return rows;
  }, [sheet, priceColumn, percent]);

  async function handleUpload(file?: File) {
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      const book = XLSX.read(data, {
        type: "array",
        cellStyles: true,
        cellFormula: true,
        cellNF: true,
      });

      const firstSheetName = book.SheetNames[0] ?? "";
      const firstSheet = book.Sheets[firstSheetName];
      const detected = firstSheet ? findPriceColumn(firstSheet) : null;

      setWorkbook(book);
      setFileName(file.name);
      setSheetName(firstSheetName);
      setStatus(
        detected
          ? `‘${detected.header}’ 열을 자동으로 찾았습니다. 변경 퍼센트만 입력하세요.`
          : "가격 열을 자동으로 찾지 못했습니다. 판매가격·판매가·가격 헤더가 있는 시트를 선택해 주세요.",
      );
    } catch {
      setWorkbook(null);
      setFileName("");
      setSheetName("");
      setStatus("파일을 읽지 못했습니다. xlsx 또는 xls 파일인지 확인해 주세요.");
    }
  }

  function changeSheet(name: string) {
    setSheetName(name);
    const nextSheet = workbook?.Sheets[name];
    const detected = nextSheet ? findPriceColumn(nextSheet) : null;
    setStatus(
      detected
        ? `‘${detected.header}’ 열을 자동으로 찾았습니다. 변경 퍼센트만 입력하세요.`
        : "이 시트에서 가격 열을 자동으로 찾지 못했습니다.",
    );
  }

  function handleDownload() {
    if (!workbook || !sheet?.["!ref"] || !priceColumn) {
      setStatus("판매가격·판매가·가격 열을 자동으로 찾지 못해 변경하지 않았습니다.");
      return;
    }

    if (!Number.isFinite(percent)) {
      setStatus("변경 퍼센트를 숫자로 입력해 주세요.");
      return;
    }

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    let changed = 0;

    for (let row = priceColumn.headerRow + 1; row <= range.e.r; row++) {
      const address = XLSX.utils.encode_cell({ r: row, c: priceColumn.column });
      const cell = sheet[address];
      if (!cell || cell.f) continue;

      const before = toNumber(cell.v);
      if (before === null) continue;

      cell.v = calculate(before);
      cell.t = "n";
      changed++;
    }

    const isXls = fileName.toLowerCase().endsWith(".xls");
    const extension = isXls ? "xls" : "xlsx";
    const baseName = fileName.replace(/\.(xlsx|xls)$/i, "");
    const suffix = percent >= 0 ? `_${percent}퍼센트인상` : `_${Math.abs(percent)}퍼센트인하`;

    XLSX.writeFile(workbook, `${baseName}${suffix}.${extension}`, {
      bookType: extension as XLSX.BookType,
      cellStyles: true,
    });

    setStatus(
      `‘${priceColumn.header}’ 열의 ${changed.toLocaleString()}개 가격을 ${percent}% 변경했습니다.`,
    );
  }

  return (
    <main>
      <section className="hero">
        <p className="eyebrow">PRICE SHEET EDITOR</p>
        <h1>엑셀 판매가격 자동 일괄 수정</h1>
        <p>파일을 올리면 판매가격 열을 자동으로 찾아 퍼센트를 적용합니다.</p>
      </section>

      <section className="card">
        <label className="upload">
          <strong>엑셀 파일 업로드</strong>
          <span>{fileName || "xlsx 또는 xls 파일 선택"}</span>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={(event) => handleUpload(event.target.files?.[0])}
          />
        </label>

        {workbook && (
          <div className="grid">
            {workbook.SheetNames.length > 1 && (
              <label>
                시트
                <select value={sheetName} onChange={(event) => changeSheet(event.target.value)}>
                  {workbook.SheetNames.map((name) => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
              </label>
            )}

            <label>
              자동 인식된 가격 열
              <input
                value={
                  priceColumn
                    ? `${XLSX.utils.encode_col(priceColumn.column)}열 · ${priceColumn.header}`
                    : "찾지 못함"
                }
                readOnly
              />
            </label>

            <label>
              변경 퍼센트 (%)
              <input
                type="number"
                step="0.1"
                value={percent}
                onChange={(event) => setPercent(Number(event.target.value))}
              />
              <small>예: 10 입력 시 10% 인상, -10 입력 시 10% 인하</small>
            </label>
          </div>
        )}
      </section>

      {preview.length > 0 && (
        <section className="card">
          <h2>변경 미리보기</h2>
          <div className="table-wrap">
            <table>
              <thead>
                <tr><th>행</th><th>변경 전</th><th>변경 후</th></tr>
              </thead>
              <tbody>
                {preview.map((item) => (
                  <tr key={item.row}>
                    <td>{item.row}</td>
                    <td>{item.before.toLocaleString()}</td>
                    <td>{item.after.toLocaleString()}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      <section className="actions">
        <button onClick={handleDownload} disabled={!workbook || !priceColumn}>
          퍼센트 적용 후 다운로드
        </button>
        <p>{status}</p>
      </section>
    </main>
  );
}
