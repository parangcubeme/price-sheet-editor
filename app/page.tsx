"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type PreviewRow = {
  row: number;
  before: number;
  after: number;
};

function toNumber(value: unknown): number | null {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;

  const cleaned = value.replace(/,/g, "").replace(/[^0-9.-]/g, "");
  if (!cleaned) return null;

  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

export default function Home() {
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState("");
  const [sheetName, setSheetName] = useState("");
  const [selectedColumn, setSelectedColumn] = useState<number | null>(null);
  const [percent, setPercent] = useState(10);
  const [roundUnit, setRoundUnit] = useState(1);
  const [status, setStatus] = useState("엑셀 파일을 올려 주세요.");

  const sheet = workbook && sheetName ? workbook.Sheets[sheetName] : null;

  const columns = useMemo(() => {
    if (!sheet?.["!ref"]) return [];

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const result: Array<{ column: number; letter: string; count: number }> = [];

    for (let column = range.s.c; column <= range.e.c; column++) {
      let count = 0;

      for (let row = range.s.r; row <= range.e.r; row++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
        if (!cell || cell.f) continue;
        if (toNumber(cell.v) !== null) count++;
      }

      if (count > 0) {
        result.push({
          column,
          letter: XLSX.utils.encode_col(column),
          count,
        });
      }
    }

    return result;
  }, [sheet]);

  function calculate(value: number) {
    const adjusted = value * (1 + percent / 100);
    const unit = Math.max(1, roundUnit);

    if (percent >= 0) {
      return Math.ceil(adjusted / unit) * unit;
    }

    return Math.floor(adjusted / unit) * unit;
  }

  const preview = useMemo<PreviewRow[]>(() => {
    if (!sheet?.["!ref"] || selectedColumn === null) return [];

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const rows: PreviewRow[] = [];

    for (let row = range.s.r; row <= range.e.r && rows.length < 20; row++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: selectedColumn })];
      if (!cell || cell.f) continue;

      const before = toNumber(cell.v);
      if (before === null) continue;

      rows.push({
        row: row + 1,
        before,
        after: calculate(before),
      });
    }

    return rows;
  }, [sheet, selectedColumn, percent, roundUnit]);

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

      setWorkbook(book);
      setFileName(file.name);
      setSheetName(book.SheetNames[0] ?? "");
      setSelectedColumn(null);
      setStatus("파일을 읽었습니다. 가격이 있는 열과 변경 퍼센트를 선택하세요.");
    } catch {
      setWorkbook(null);
      setFileName("");
      setSheetName("");
      setSelectedColumn(null);
      setStatus("파일을 읽지 못했습니다. xlsx 또는 xls 파일인지 확인해 주세요.");
    }
  }

  function handleDownload() {
    if (!workbook || !sheet?.["!ref"] || selectedColumn === null) {
      setStatus("가격이 있는 열을 선택해 주세요.");
      return;
    }

    if (!Number.isFinite(percent)) {
      setStatus("변경 퍼센트를 숫자로 입력해 주세요.");
      return;
    }

    const range = XLSX.utils.decode_range(sheet["!ref"]);
    let changed = 0;

    for (let row = range.s.r; row <= range.e.r; row++) {
      const address = XLSX.utils.encode_cell({ r: row, c: selectedColumn });
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

    setStatus(`${changed.toLocaleString()}개 가격을 ${percent}% 변경했습니다.`);
  }

  return (
    <main>
      <section className="hero">
        <p className="eyebrow">PRICE SHEET EDITOR</p>
        <h1>엑셀 가격 퍼센트 일괄 수정</h1>
        <p>가격 열을 직접 선택하고 올리거나 내릴 퍼센트만 입력하면 됩니다.</p>
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
            <label>
              시트
              <select
                value={sheetName}
                onChange={(event) => {
                  setSheetName(event.target.value);
                  setSelectedColumn(null);
                }}
              >
                {workbook.SheetNames.map((name) => (
                  <option key={name} value={name}>
                    {name}
                  </option>
                ))}
              </select>
            </label>

            <label>
              가격 열
              <select
                value={selectedColumn ?? ""}
                onChange={(event) =>
                  setSelectedColumn(event.target.value === "" ? null : Number(event.target.value))
                }
              >
                <option value="">열 선택</option>
                {columns.map((item) => (
                  <option key={item.column} value={item.column}>
                    {item.letter}열 ({item.count.toLocaleString()}개 숫자)
                  </option>
                ))}
              </select>
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

            <label>
              금액 단위
              <select value={roundUnit} onChange={(event) => setRoundUnit(Number(event.target.value))}>
                <option value="1">1원 단위</option>
                <option value="10">10원 단위</option>
                <option value="100">100원 단위</option>
                <option value="1000">1,000원 단위</option>
              </select>
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
                <tr>
                  <th>행</th>
                  <th>변경 전</th>
                  <th>변경 후</th>
                </tr>
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
        <button onClick={handleDownload} disabled={!workbook || selectedColumn === null}>
          퍼센트 적용 후 다운로드
        </button>
        <p>{status}</p>
      </section>
    </main>
  );
}
