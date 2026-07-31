"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type RowPreview = { row: number; before: number; after: number };

const toNumber = (value: unknown): number | null => {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;
  const cleaned = value.replace(/,/g, "").replace(/[^0-9.-]/g, "");
  if (!cleaned) return null;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
};

export default function Home() {
  const [book, setBook] = useState<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState("");
  const [sheetName, setSheetName] = useState("");
  const [headerRow, setHeaderRow] = useState(1);
  const [column, setColumn] = useState<number | null>(null);
  const [fee, setFee] = useState(0);
  const [margin, setMargin] = useState(0);
  const [consumerRatio, setConsumerRatio] = useState(100);
  const [roundUnit, setRoundUnit] = useState(100);
  const [status, setStatus] = useState("엑셀 파일을 업로드해 주세요.");

  const sheet = book && sheetName ? book.Sheets[sheetName] : null;

  const columns = useMemo(() => {
    if (!sheet?.["!ref"]) return [];
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const row = Math.max(0, headerRow - 1);
    return Array.from({ length: range.e.c - range.s.c + 1 }, (_, i) => {
      const c = range.s.c + i;
      const address = XLSX.utils.encode_cell({ r: row, c });
      const label = String(sheet[address]?.v ?? `${XLSX.utils.encode_col(c)}열`);
      let count = 0;
      for (let r = row + 1; r <= range.e.r; r++) {
        if (toNumber(sheet[XLSX.utils.encode_cell({ r, c })]?.v) !== null) count++;
      }
      return { c, label, count, letter: XLSX.utils.encode_col(c) };
    }).filter((item) => item.count > 0);
  }, [sheet, headerRow]);

  const calculate = (value: number) => {
    const denominator = 1 - fee / 100 - margin / 100;
    if (denominator <= 0) return NaN;
    const raw = (value / denominator) * (consumerRatio / 100);
    return Math.ceil(raw / Math.max(1, roundUnit)) * Math.max(1, roundUnit);
  };

  const preview = useMemo<RowPreview[]>(() => {
    if (!sheet?.["!ref"] || column === null) return [];
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const result: RowPreview[] = [];
    for (let r = Math.max(0, headerRow); r <= range.e.r && result.length < 20; r++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c: column })];
      const before = toNumber(cell?.v);
      if (before === null || cell?.f) continue;
      result.push({ row: r + 1, before, after: calculate(before) });
    }
    return result;
  }, [sheet, column, headerRow, fee, margin, consumerRatio, roundUnit]);

  async function upload(file?: File) {
    if (!file) return;
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array", cellStyles: true, cellFormula: true, cellNF: true });
      setBook(workbook);
      setFileName(file.name);
      setSheetName(workbook.SheetNames[0] ?? "");
      setColumn(null);
      setStatus("파일을 읽었습니다. 시트와 가격 열을 선택해 주세요.");
    } catch {
      setStatus("파일을 읽지 못했습니다. xlsx 또는 xls 파일인지 확인해 주세요.");
    }
  }

  function download() {
    if (!book || !sheet?.["!ref"] || column === null) {
      setStatus("먼저 파일과 가격 열을 선택해 주세요.");
      return;
    }
    if (fee + margin >= 100) {
      setStatus("수수료율과 마진율의 합은 100%보다 작아야 합니다.");
      return;
    }
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    let changed = 0;
    for (let r = Math.max(0, headerRow); r <= range.e.r; r++) {
      const address = XLSX.utils.encode_cell({ r, c: column });
      const cell = sheet[address];
      const before = toNumber(cell?.v);
      if (!cell || before === null || cell.f) continue;
      cell.v = calculate(before);
      cell.t = "n";
      changed++;
    }
    const ext = fileName.toLowerCase().endsWith(".xls") ? "xls" : "xlsx";
    const base = fileName.replace(/\.(xlsx|xls)$/i, "");
    XLSX.writeFile(book, `${base}_가격수정.${ext}`, { bookType: ext as XLSX.BookType, cellStyles: true });
    setStatus(`${changed.toLocaleString()}개 가격을 변경해 다운로드했습니다.`);
  }

  return (
    <main>
      <section className="hero">
        <p className="eyebrow">PRICE SHEET EDITOR</p>
        <h1>엑셀 가격 일괄 변경기</h1>
        <p>원본 시트 구조는 유지하고 선택한 가격 열의 값만 바꿉니다.</p>
      </section>

      <section className="card">
        <label className="upload">
          <strong>1. 엑셀 업로드</strong>
          <span>{fileName || "xlsx 또는 xls 파일 선택"}</span>
          <input type="file" accept=".xlsx,.xls" onChange={(e) => upload(e.target.files?.[0])} />
        </label>

        {book && (
          <div className="grid">
            <label>시트<select value={sheetName} onChange={(e) => { setSheetName(e.target.value); setColumn(null); }}>{book.SheetNames.map((name) => <option key={name}>{name}</option>)}</select></label>
            <label>제목 행<input type="number" min="1" value={headerRow} onChange={(e) => setHeaderRow(Math.max(1, Number(e.target.value)))} /></label>
            <label className="wide">가격 열<select value={column ?? ""} onChange={(e) => setColumn(e.target.value === "" ? null : Number(e.target.value))}><option value="">선택하세요</option>{columns.map((item) => <option key={item.c} value={item.c}>{item.letter}열 · {item.label} ({item.count}개 숫자)</option>)}</select></label>
            <label>수수료율 (%)<input type="number" value={fee} onChange={(e) => setFee(Number(e.target.value))} /></label>
            <label>마진율 (%)<input type="number" value={margin} onChange={(e) => setMargin(Number(e.target.value))} /></label>
            <label>소비자가 비율 (%)<input type="number" value={consumerRatio} onChange={(e) => setConsumerRatio(Number(e.target.value))} /></label>
            <label>올림 단위<select value={roundUnit} onChange={(e) => setRoundUnit(Number(e.target.value))}><option value="1">1원</option><option value="10">10원</option><option value="100">100원</option><option value="1000">1,000원</option></select></label>
          </div>
        )}
      </section>

      {preview.length > 0 && (
        <section className="card">
          <h2>변경 미리보기</h2>
          <div className="table-wrap"><table><thead><tr><th>행</th><th>변경 전</th><th>변경 후</th></tr></thead><tbody>{preview.map((row) => <tr key={row.row}><td>{row.row}</td><td>{row.before.toLocaleString()}</td><td>{Number.isFinite(row.after) ? row.after.toLocaleString() : "계산 오류"}</td></tr>)}</tbody></table></div>
        </section>
      )}

      <section className="actions">
        <button onClick={download} disabled={!book || column === null}>가격 변경 후 다운로드</button>
        <p>{status}</p>
      </section>
    </main>
  );
}
