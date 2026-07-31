"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

type PriceTarget = {
  sheetName: string;
  headerRow: number;
  column: number;
};

type PreviewRow = {
  sheetName: string;
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

function normalizeHeader(value: unknown) {
  return String(value ?? "").replace(/\s+/g, "").trim();
}

function findPriceTargets(book: XLSX.WorkBook): PriceTarget[] {
  const targets: PriceTarget[] = [];

  for (const sheetName of book.SheetNames) {
    const sheet = book.Sheets[sheetName];
    if (!sheet?.["!ref"]) continue;

    const range = XLSX.utils.decode_range(sheet["!ref"]);

    for (let row = range.s.r; row <= range.e.r; row++) {
      for (let column = range.s.c; column <= range.e.c; column++) {
        const address = XLSX.utils.encode_cell({ r: row, c: column });
        const cell = sheet[address];

        if (cell && normalizeHeader(cell.v) === "판매가격") {
          targets.push({ sheetName, headerRow: row, column });
        }
      }
    }
  }

  return targets;
}

export default function Home() {
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [fileName, setFileName] = useState("");
  const [margin, setMargin] = useState(0);
  const [fee, setFee] = useState(0);
  const [targets, setTargets] = useState<PriceTarget[]>([]);
  const [status, setStatus] = useState("엑셀 파일을 올려 주세요.");

  const multiplier = 1 + (margin + fee) / 100;

  function calculate(value: number) {
    return Math.round(value * multiplier);
  }

  const preview = useMemo<PreviewRow[]>(() => {
    if (!workbook || targets.length === 0) return [];

    const rows: PreviewRow[] = [];

    for (const target of targets) {
      const sheet = workbook.Sheets[target.sheetName];
      if (!sheet?.["!ref"]) continue;

      const range = XLSX.utils.decode_range(sheet["!ref"]);

      for (let row = target.headerRow + 1; row <= range.e.r && rows.length < 20; row++) {
        const address = XLSX.utils.encode_cell({ r: row, c: target.column });
        const cell = sheet[address];
        if (!cell || cell.f) continue;

        const before = toNumber(cell.v);
        if (before === null) continue;

        rows.push({
          sheetName: target.sheetName,
          row: row + 1,
          before,
          after: calculate(before),
        });
      }

      if (rows.length >= 20) break;
    }

    return rows;
  }, [workbook, targets, margin, fee]);

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

      const foundTargets = findPriceTargets(book);

      setWorkbook(book);
      setFileName(file.name);
      setTargets(foundTargets);

      if (foundTargets.length === 0) {
        setStatus("'판매가격' 열을 찾지 못했습니다.");
      } else {
        setStatus(`'판매가격' 열 ${foundTargets.length.toLocaleString()}개를 찾았습니다. 마진율과 수수료율을 입력하세요.`);
      }
    } catch {
      setWorkbook(null);
      setTargets([]);
      setStatus("파일을 읽지 못했습니다. xlsx 또는 xls 파일인지 확인해 주세요.");
    }
  }

  function handleDownload() {
    if (!workbook) {
      setStatus("엑셀 파일을 먼저 올려 주세요.");
      return;
    }

    if (targets.length === 0) {
      setStatus("'판매가격' 열을 찾지 못해 수정할 수 없습니다.");
      return;
    }

    if (!Number.isFinite(margin) || !Number.isFinite(fee)) {
      setStatus("마진율과 수수료율을 숫자로 입력해 주세요.");
      return;
    }

    let changed = 0;

    for (const target of targets) {
      const sheet = workbook.Sheets[target.sheetName];
      if (!sheet?.["!ref"]) continue;

      const range = XLSX.utils.decode_range(sheet["!ref"]);

      for (let row = target.headerRow + 1; row <= range.e.r; row++) {
        const address = XLSX.utils.encode_cell({ r: row, c: target.column });
        const cell = sheet[address];
        if (!cell || cell.f) continue;

        const before = toNumber(cell.v);
        if (before === null) continue;

        cell.v = calculate(before);
        cell.t = "n";
        changed++;
      }
    }

    const isXls = fileName.toLowerCase().endsWith(".xls");
    const extension = isXls ? "xls" : "xlsx";
    const baseName = fileName.replace(/\.(xlsx|xls)$/i, "");

    XLSX.writeFile(workbook, `${baseName}_판매가격수정.${extension}`, {
      bookType: extension as XLSX.BookType,
      cellStyles: true,
    });

    setStatus(`완료: 판매가격 ${changed.toLocaleString()}개를 수정했습니다.`);
  }

  return (
    <main>
      <section className="hero">
        <p className="eyebrow">PRICE SHEET EDITOR</p>
        <h1>엑셀 판매가격 일괄 수정</h1>
        <p>엑셀을 올리고 마진율과 수수료율만 입력하면 판매가격이 자동으로 수정됩니다.</p>
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

        <div className="grid">
          <label>
            마진율 (%)
            <input
              type="number"
              step="0.1"
              value={margin}
              onChange={(event) => setMargin(Number(event.target.value))}
            />
          </label>

          <label>
            수수료율 (%)
            <input
              type="number"
              step="0.1"
              value={fee}
              onChange={(event) => setFee(Number(event.target.value))}
            />
          </label>
        </div>

        <p className="formula">
          판매가격 × [1 + (마진율 + 수수료율) ÷ 100]
        </p>
      </section>

      {preview.length > 0 && (
        <section className="card">
          <h2>변경 미리보기</h2>
          <div className="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>시트</th>
                  <th>행</th>
                  <th>변경 전</th>
                  <th>변경 후</th>
                </tr>
              </thead>
              <tbody>
                {preview.map((item) => (
                  <tr key={`${item.sheetName}-${item.row}`}>
                    <td>{item.sheetName}</td>
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
        <button onClick={handleDownload} disabled={!workbook || targets.length === 0}>
          판매가격 수정 후 다운로드
        </button>
        <p>{status}</p>
      </section>
    </main>
  );
}
