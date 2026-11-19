"use client";

import { useState, type ReactNode } from "react";

export type Source = { id: string; url: string; title: string };
export type TableData = { columns: string[]; rows: (string | number)[][] };

function TableView({ table }: { table: TableData }) {
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {table.columns.map((c, j) => (
              <th key={c} className={typeof table.rows[0]?.[j] === "number" ? "n" : undefined}>{c}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {table.rows.map((r, i) => (
            <tr key={i}>
              {r.map((v, j) => (
                <td key={j} className={typeof v === "number" ? "n" : undefined}>{typeof v === "number" ? v.toLocaleString("en-US") : v}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function download(table: TableData, file: string) {
  const esc = (v: string | number) => {
    const s = String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const csv = [table.columns, ...table.rows].map((r) => r.map(esc).join(",")).join("\n");
  const url = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  const a = document.createElement("a");
  a.href = url;
  a.download = file;
  a.click();
  URL.revokeObjectURL(url);
}

export default function Exhibit({ n, title, sub, badge, note, sources, table, file, toggle = true, children }: {
  n: number;
  title: string;
  sub?: string;
  badge?: string;
  note: string;
  sources: Source[];
  table: TableData;
  file: string;
  toggle?: boolean;
  children: ReactNode;
}) {
  const [asTable, setAsTable] = useState(false);
  return (
    <figure className="exhibit" aria-labelledby={`ex${n}`}>
      <p className="exhibit-n">Exhibit {n}</p>
      <h3 id={`ex${n}`}>{title}</h3>
      {sub && <p className="sub">{sub}</p>}
      {badge && <p className="badge">{badge}</p>}
      {asTable ? <TableView table={table} /> : children}
      <figcaption className="exhibit-foot">
        <span>
          Source: {note}
          {sources.length > 0 && "; "}
          {sources.map((s, i) => (
            <span key={s.id}>
              {i > 0 && ", "}
              <a href={s.url} title={s.title} target="_blank" rel="noopener noreferrer">{s.id}</a>
            </span>
          ))}
        </span>
        <span className="actions">
          {toggle && (
            <button type="button" className="textbtn" onClick={() => setAsTable(!asTable)}>
              {asTable ? "Show chart" : "Show table"}
            </button>
          )}
          <button type="button" className="textbtn" onClick={() => download(table, file)}>Download data (CSV)</button>
        </span>
      </figcaption>
    </figure>
  );
}
