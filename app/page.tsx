"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Activity, ArrowDownToLine, CalendarDays, Check, CheckCheck, CircleHelp,
  Eye, FileSpreadsheet, LoaderCircle, Pencil, Plus, Search, Trash2,
  UploadCloud, UsersRound, X,
} from "lucide-react";
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { type ProgressInput, validateProgress } from "@/lib/progress";

type Row = ProgressInput & { id: number; updatedAt: string };
type Parsed = { rows: ProgressInput[]; errors: string[]; filename: string };
const STORAGE_KEY = "timi-arajyntac-v1";

function storedRows(): Row[] {
  const value = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
  if (!Array.isArray(value)) throw new Error("Պահպանված տվյալները վնասված են։");
  return value;
}

function putRows(rows: Row[]) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
}

function armeniaMonth() {
  const parts = new Intl.DateTimeFormat("en-US", {
    timeZone: "Asia/Yerevan", year: "numeric", month: "2-digit",
  }).formatToParts(new Date());
  return `${parts.find(p => p.type === "year")?.value}-${parts.find(p => p.type === "month")?.value}`;
}

function monthLabel(value: string) {
  const [year, month] = value.split("-").map(Number);
  if (!year || !month) return value;
  return new Intl.DateTimeFormat("hy-AM", { year: "numeric", month: "long", timeZone: "UTC" })
    .format(new Date(Date.UTC(year, month - 1, 1)));
}

function monthFromCell(cell: unknown, XLSX: typeof import("xlsx")) {
  if (cell instanceof Date && !Number.isNaN(cell.getTime())) {
    return `${cell.getFullYear()}-${String(cell.getMonth() + 1).padStart(2, "0")}`;
  }
  if (typeof cell === "number") {
    const parsed = XLSX.SSF.parse_date_code(cell);
    if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, "0")}`;
  }
  const str = String(cell ?? "").trim();
  const iso = str.match(/^(\d{4})[-/.](\d{1,2})(?:[-/.]\d{1,2})?$/);
  if (iso) return `${iso[1]}-${iso[2].padStart(2, "0")}`;
  const dotted = str.match(/^\d{1,2}[./-](\d{1,2})[./-](\d{4})$/);
  if (dotted) return `${dotted[2]}-${dotted[1].padStart(2, "0")}`;
  return str;
}

const aliases: Record<string, string[]> = {
  memberName: ["աշխատակից", "անուն", "աշխատակցի անուն", "name", "employee", "member"],
  month: ["ամիս", "ժամանակաշրջան", "month", "period"],
  tasks: ["առաջադրանքներ", "թասկեր", "թասքեր", "tasks", "task"],
  observations: ["դիտարկումներ", "դիտարկում", "observations", "observation", "reviews"],
};

async function readSpreadsheet(file: File): Promise<Parsed> {
  const XLSX = await import("xlsx");
  const book = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const sheet = book.Sheets[book.SheetNames[0]];
  if (!sheet) throw new Error("Ֆայլում թերթ չի գտնվել։");
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: "", raw: true });
  if (!matrix.length) throw new Error("Ֆայլը դատարկ է։");
  const header = matrix[0].map(cell => String(cell).normalize("NFKC").trim().toLocaleLowerCase());
  const indices = Object.fromEntries(Object.entries(aliases).map(([key, options]) => [
    key, header.findIndex(cell => options.includes(cell)),
  ])) as Record<keyof typeof aliases, number>;
  if (Object.values(indices).some(index => index < 0)) {
    throw new Error("Առաջին տողում պետք է լինեն՝ Աշխատակից, Ամիս, Առաջադրանքներ, Դիտարկումներ սյունակները։");
  }
  const rows: ProgressInput[] = [];
  const errors: string[] = [];
  for (let i = 1; i < matrix.length; i++) {
    const line = matrix[i];
    if (!line || line.every(cell => String(cell ?? "").trim() === "")) continue;
    const candidate = {
      memberName: String(line[indices.memberName] ?? ""),
      month: monthFromCell(line[indices.month], XLSX),
      tasks: Number(String(line[indices.tasks] ?? "").replace(/\s/g, "")),
      observations: Number(String(line[indices.observations] ?? "").replace(/\s/g, "")),
    };
    const valid = validateProgress(candidate);
    if (valid) rows.push(valid);
    else errors.push(`${i + 1}-րդ տող`);
  }
  if (!rows.length && !errors.length) throw new Error("Տվյալներով տողեր չեն գտնվել։");
  return { rows, errors, filename: file.name };
}

async function downloadWorkbook(rows: ProgressInput[], filename: string) {
  const XLSX = await import("xlsx");
  const data = rows.map(row => ({
    "Աշխատակից": row.memberName,
    "Ամիս": row.month,
    "Առաջադրանքներ": row.tasks,
    "Դիտարկումներ": row.observations,
  }));
  const sheet = XLSX.utils.json_to_sheet(data, {
    header: ["Աշխատակից", "Ամիս", "Առաջադրանքներ", "Դիտարկումներ"],
  });
  sheet["!cols"] = [{ wch: 24 }, { wch: 14 }, { wch: 19 }, { wch: 18 }];
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, "Առաջընթաց");
  XLSX.writeFile(book, filename);
}

export default function Home() {
  const [rows, setRows] = useState<Row[]>([]);
  const [selectedMonth, setSelectedMonth] = useState(armeniaMonth);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState("");
  const [query, setQuery] = useState("");
  const [form, setForm] = useState<ProgressInput | null>(null);
  const [saving, setSaving] = useState(false);
  const [importOpen, setImportOpen] = useState(false);
  const [parsed, setParsed] = useState<Parsed | null>(null);
  const [importError, setImportError] = useState("");
  const [reading, setReading] = useState(false);
  const [helpOpen, setHelpOpen] = useState(false);
  const [toast, setToast] = useState("");
  const fileInput = useRef<HTMLInputElement>(null);

  const notify = (message: string) => {
    setToast(message);
    window.setTimeout(() => setToast(""), 4800);
  };

  const refresh = useCallback(async () => {
    try {
      setRows(storedRows());
      setLoadError("");
    } catch (error) {
      setLoadError(error instanceof Error ? error.message : "Տվյալները չհաջողվեց բեռնել։");
    } finally {
      setLoading(false);
    }
  }, []);
  useEffect(() => { void refresh(); }, [refresh]);

  const months = useMemo(
    () => [...new Set([armeniaMonth(), selectedMonth, ...rows.map(row => row.month)])].sort().reverse(),
    [rows, selectedMonth],
  );
  const monthRows = useMemo(
    () => rows.filter(row => row.month === selectedMonth),
    [rows, selectedMonth],
  );
  const filteredRows = useMemo(
    () => monthRows.filter(row => row.memberName.toLocaleLowerCase().includes(query.toLocaleLowerCase())),
    [monthRows, query],
  );
  const totals = useMemo(() => ({
    tasks: monthRows.reduce((sum, row) => sum + row.tasks, 0),
    observations: monthRows.reduce((sum, row) => sum + row.observations, 0),
    people: monthRows.length,
  }), [monthRows]);
  const maxValue = Math.max(1, ...monthRows.flatMap(row => [row.tasks, row.observations]));
  const trend = useMemo(() => months.slice(0, 6).reverse().map(month => {
    const entries = rows.filter(row => row.month === month);
    return {
      month,
      tasks: entries.reduce((sum, row) => sum + row.tasks, 0),
      observations: entries.reduce((sum, row) => sum + row.observations, 0),
    };
  }), [rows, months]);
  const trendMax = Math.max(1, ...trend.flatMap(item => [item.tasks, item.observations]));

  async function saveRecord(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!form || saving) return;
    setSaving(true);
    try {
      const valid = validateProgress(form);
      if (!valid) throw new Error("Ստուգեք անունը, ամիսը և թվերը։");
      const current = storedRows();
      const key = valid.memberName.toLocaleLowerCase();
      const duplicate = current.find(row => row.month === valid.month && row.memberName.toLocaleLowerCase() === key && row.id !== valid.id);
      const updated: Row = {
        ...valid,
        id: valid.id ?? duplicate?.id ?? Math.max(0, ...current.map(row => row.id)) + 1,
        updatedAt: new Date().toISOString(),
      };
      const next = current.filter(row => row.id !== updated.id && row.id !== duplicate?.id).concat(updated);
      putRows(next);
      setSelectedMonth(form.month);
      setForm(null);
      await refresh();
      notify("Գրառումը պահպանվեց");
    } catch (error) {
      notify(error instanceof Error ? error.message : "Պահպանման սխալ։");
    } finally {
      setSaving(false);
    }
  }

  async function removeRecord(row: Row) {
    if (!window.confirm(`${row.memberName} · ${monthLabel(row.month)} գրառումը ջնջե՞լ։`)) return;
    try {
      putRows(storedRows().filter(item => item.id !== row.id));
      await refresh();
      notify("Գրառումը ջնջվեց");
    } catch (error) {
      notify(error instanceof Error ? error.message : "Ջնջման սխալ։");
    }
  }

  async function handleFile(file?: File) {
    if (!file) return;
    setReading(true);
    setImportError("");
    setParsed(null);
    try {
      if (file.size > 8 * 1024 * 1024) throw new Error("Ֆայլը պետք է լինի 8 ՄԲ-ից փոքր։");
      const result = await readSpreadsheet(file);
      setParsed(result);
      if (result.errors.length) setImportError(`Ստուգեք ${result.errors.slice(0, 5).join(", ")}${result.errors.length > 5 ? " և մյուս սխալները" : ""}։`);
    } catch (error) {
      setImportError(error instanceof Error ? error.message : "Ֆայլը չհաջողվեց բացել։");
    } finally {
      setReading(false);
      if (fileInput.current) fileInput.current.value = "";
    }
  }

  async function importRows() {
    if (!parsed || parsed.errors.length || !parsed.rows.length || saving) return;
    setSaving(true);
    try {
      if (parsed.rows.length > 1000) throw new Error("Մեկ անգամում ներմուծեք առավելագույնը 1000 տող։");
      const next = storedRows();
      let nextId = Math.max(0, ...next.map(row => row.id));
      for (const row of parsed.rows) {
        const index = next.findIndex(item => item.month === row.month && item.memberName.toLocaleLowerCase() === row.memberName.toLocaleLowerCase());
        const value = { ...row, id: index >= 0 ? next[index].id : ++nextId, updatedAt: new Date().toISOString() };
        if (index >= 0) next[index] = value;
        else next.push(value);
      }
      putRows(next);
      setImportOpen(false);
      setParsed(null);
      await refresh();
      notify(`${parsed.rows.length} տող ներմուծվեց`);
    } catch (error) {
      setImportError(error instanceof Error ? error.message : "Ներմուծման սխալ։");
    } finally {
      setSaving(false);
    }
  }

  const zero = loading || !!loadError;
  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand"><div className="brand-mark"><Activity size={23} strokeWidth={2.5} /></div><span>ԹԻՄ<span className="brand-dot">.</span><small>առաջընթաց</small></span></div>
        <div className="side-label">ԱՇԽԱՏԱՆՔԱՅԻՆ ՏԱՐԱԾՔ</div>
        <div className="side-active"><Activity size={18} /><span>Ընդհանուր պատկերը</span><span className="active-line" /></div>
        <div className="side-bottom">
          <div className="side-note"><span className="side-note-icon"><FileSpreadsheet size={19} /></span><strong>Ձեր տվյալները՝ մեկ տեղում</strong><p>Թարմացրեք ձեռքով կամ ներմուծեք Excel ֆայլից։</p></div>
          <button className="side-help" onClick={() => setHelpOpen(true)}><CircleHelp size={18} /> Ինչպես օգտվել</button>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div className="breadcrumb">Վահանակ <span>/</span> <strong>Ընդհանուր պատկերը</strong></div>
          <button className="top-help" onClick={() => setHelpOpen(true)} aria-label="Օգնություն"><CircleHelp size={20} /></button>
        </header>
        <div className="content">
          <div className="page-heading">
            <div>
              <div className="eyebrow"><span className="live-pulse" /> ԹԻՄԻ ՎԻՃԱԿԱԳՐՈՒԹՅՈՒՆ</div>
              <h1>Թիմի առաջընթացը</h1>
              <p>Առաջադրանքներն ու դիտարկումները՝ պարզ և տեսանելի։</p>
            </div>
            <div className="heading-actions">
              {rows.length > 0 && <button className="action-secondary" onClick={() => void downloadWorkbook(rows, "timi-arajyntac-amboghj.xlsx")}><ArrowDownToLine size={18} /> Պահուստային Excel</button>}
              <button className="action-secondary" onClick={() => { setParsed(null); setImportError(""); setImportOpen(true); }}><UploadCloud size={18} /> Excel ներմուծել</button>
              <button className="action-primary" onClick={() => setForm({ memberName: "", month: selectedMonth, tasks: 0, observations: 0 })}><Plus size={18} /> Ավելացնել գրառում</button>
            </div>
          </div>

          <div className="period-line">
            <div className="period-title"><CalendarDays size={18} /><span>Ժամանակաշրջան</span></div>
            <div className="month-control">
              <select aria-label="Ընտրել ամիսը" value={selectedMonth} onChange={e => setSelectedMonth(e.target.value)}>
                {months.map(month => <option value={month} key={month}>{monthLabel(month)}</option>)}
              </select>
            </div>
          </div>
          {loadError && <div className="error-banner" role="alert">{loadError} <button onClick={() => { setLoading(true); void refresh(); }}>Կրկին փորձել</button></div>}

          <section className="stats-grid" aria-label="Ամփոփ ցուցանիշներ">
            <div className="stat-card stat-tasks"><div className="stat-top"><span>Կատարված առաջադրանքներ</span><span className="stat-icon"><CheckCheck size={21} /></span></div><strong>{zero ? "—" : totals.tasks.toLocaleString("hy-AM")}</strong><div className="stat-foot"><span className="small-line" /> Ընտրված ամսում</div></div>
            <div className="stat-card stat-observations"><div className="stat-top"><span>Իրականացված դիտարկումներ</span><span className="stat-icon"><Eye size={21} /></span></div><strong>{zero ? "—" : totals.observations.toLocaleString("hy-AM")}</strong><div className="stat-foot"><span className="small-line" /> Ընտրված ամսում</div></div>
            <div className="stat-card stat-people"><div className="stat-top"><span>Թիմի անդամներ</span><span className="stat-icon"><UsersRound size={21} /></span></div><strong>{zero ? "—" : totals.people.toLocaleString("hy-AM")}</strong><div className="stat-foot"><span className="small-line" /> Այս ամսվա տվյալներով</div></div>
          </section>

          <div className="data-grid">
            <section className="panel team-panel">
              <div className="panel-head">
                <div><div className="section-kicker">ԱՇԽԱՏԱԿԻՑՆԵՐ</div><h2>Անհատական արդյունքներ</h2></div>
                {monthRows.length > 0 && <span className="count-pill">{monthRows.length} մարդ</span>}
              </div>
              {monthRows.length > 0 && <div className="table-tools">
                <label className="search-box"><Search size={18} /><input value={query} onChange={e => setQuery(e.target.value)} placeholder="Որոնել աշխատակցին" aria-label="Որոնել աշխատակցին" /></label>
                <button className="export-btn" onClick={() => void downloadWorkbook(monthRows, `timi-arajyntac-${selectedMonth}.xlsx`)} title="Ներբեռնել ընտրված ամսվա Excel-ը"><ArrowDownToLine size={17} /><span>Արտահանել</span></button>
              </div>}
              {loading ? <div className="empty-state"><LoaderCircle className="spin" size={28} /><p>Տվյալները բեռնվում են…</p></div> :
                monthRows.length === 0 ? <div className="empty-state"><div className="empty-graphic"><FileSpreadsheet size={31} /></div><h3>Դեռ տվյալներ չկան</h3><p>Ավելացրեք առաջին գրառումը կամ ներմուծեք ձեր Excel ֆայլը։</p><button onClick={() => setForm({ memberName: "", month: selectedMonth, tasks: 0, observations: 0 })}>Ավելացնել գրառում <Plus size={16} /></button></div> :
                filteredRows.length === 0 ? <div className="no-search">Այս անունով աշխատակից չկա։</div> :
                <div className="rows">
                  <div className="table-labels"><span>ԱՇԽԱՏԱԿԻՑ</span><span>ԱՌԱՋԱԴՐԱՆՔՆԵՐ</span><span>ԴԻՏԱՐԿՈՒՄՆԵՐ</span><span className="sr-only">Գործողություններ</span></div>
                  {filteredRows.map((row, index) => <div className="member-row" key={row.id} style={{ animationDelay: `${Math.min(index * 45, 360)}ms` }}>
                    <div className="member-name"><span className="avatar" style={{ background: ["#dce8ff", "#d7f5ee", "#efe4ff", "#ffeacc"][index % 4], color: ["#254e9b", "#0a7665", "#6e44a9", "#8c5b12"][index % 4] }}>{row.memberName.trim().charAt(0).toLocaleUpperCase()}</span><span title={row.memberName}>{row.memberName}</span></div>
                    <div className="metric-cell"><strong>{row.tasks.toLocaleString("hy-AM")}</strong><span className="bar-track"><span className="bar-fill tasks-fill" style={{ width: `${Math.max(row.tasks ? 6 : 0, row.tasks / maxValue * 100)}%` }} /></span></div>
                    <div className="metric-cell"><strong>{row.observations.toLocaleString("hy-AM")}</strong><span className="bar-track"><span className="bar-fill obs-fill" style={{ width: `${Math.max(row.observations ? 6 : 0, row.observations / maxValue * 100)}%` }} /></span></div>
                    <div className="row-actions"><button aria-label={`${row.memberName}՝ խմբագրել`} title="Խմբագրել" onClick={() => setForm({ id: row.id, memberName: row.memberName, month: row.month, tasks: row.tasks, observations: row.observations })}><Pencil size={16} /></button><button aria-label={`${row.memberName}՝ ջնջել`} title="Ջնջել" onClick={() => void removeRecord(row)}><Trash2 size={16} /></button></div>
                  </div>)}
                </div>}
            </section>

            <section className="panel trend-panel">
              <div className="panel-head"><div><div className="section-kicker">ԴԻՆԱՄԻԿԱ</div><h2>Ամսական պատկերը</h2></div></div>
              <div className="legend"><span><i className="legend-task" /> Առաջադրանքներ</span><span><i className="legend-obs" /> Դիտարկումներ</span></div>
              {rows.length === 0 ? <div className="trend-empty"><div className="trend-empty-icon"><Activity size={24} /></div><p>Գրաֆիկը կհայտնվի, երբ տվյալներ ավելացնեք։</p></div> :
              <div className="chart-wrap">
                <div className="chart-guides"><span>{trendMax}</span><span>{Math.round(trendMax / 2)}</span><span>0</span></div>
                <div className="chart-plot">
                  <div className="gridline gridline-top" /><div className="gridline gridline-middle" /><div className="gridline gridline-bottom" />
                  <div className="chart-columns">
                    {trend.map(item => <div className="chart-column" key={item.month} title={`${monthLabel(item.month)} · ${item.tasks} առաջադրանք, ${item.observations} դիտարկում`}>
                      <div className="bars"><span className="chart-bar task-bar" style={{ height: `${Math.max(item.tasks ? 4 : 0, item.tasks / trendMax * 100)}%` }} /><span className="chart-bar obs-bar" style={{ height: `${Math.max(item.observations ? 4 : 0, item.observations / trendMax * 100)}%` }} /></div>
                      <span className="chart-month">{new Intl.DateTimeFormat("hy-AM", { month: "short", timeZone: "UTC" }).format(new Date(`${item.month}-01T00:00:00Z`))}</span>
                    </div>)}
                  </div>
                </div>
              </div>}
              <div className="trend-footer"><span className="trend-dash" /> Վերջին {Math.min(months.length, 6)} ամիսների տվյալները</div>
            </section>
          </div>
          <footer className="site-footer"><span>Թիմի առաջընթաց</span><span>Տվյալները պահպանվում են այս բրաուզերում · պահուստավորեք Excel-ով</span></footer>
        </div>
      </main>

      <Dialog open={!!form} onOpenChange={open => { if (!open) setForm(null); }}>
        <DialogContent className="app-dialog"><DialogHeader><div className="dialog-icon"><Pencil size={20} /></div><DialogTitle>{form?.id ? "Խմբագրել գրառումը" : "Նոր գրառում"}</DialogTitle><DialogDescription>Լրացրեք աշխատակցի ամսական արդյունքները։</DialogDescription></DialogHeader>
          {form && <form onSubmit={e => void saveRecord(e)} className="record-form">
            <label>Աշխատակցի անունը<Input autoFocus required maxLength={100} value={form.memberName} onChange={e => setForm({ ...form, memberName: e.target.value })} placeholder="Օրինակ՝ Անի Սարգսյան" /></label>
            <label>Ամիսը<Input required type="month" value={form.month} onChange={e => setForm({ ...form, month: e.target.value })} /></label>
            <div className="form-pair"><label>Առաջադրանքներ<Input required type="number" min="0" max="999999" step="1" value={form.tasks} onChange={e => setForm({ ...form, tasks: Number(e.target.value) })} /></label><label>Դիտարկումներ<Input required type="number" min="0" max="999999" step="1" value={form.observations} onChange={e => setForm({ ...form, observations: Number(e.target.value) })} /></label></div>
            <DialogFooter><Button type="button" variant="outline" onClick={() => setForm(null)}>Չեղարկել</Button><Button disabled={saving} type="submit" className="modal-primary">{saving ? <LoaderCircle className="spin" /> : <Check />} Պահպանել</Button></DialogFooter>
          </form>}
        </DialogContent>
      </Dialog>

      <Dialog open={importOpen} onOpenChange={open => { setImportOpen(open); if (!open) { setParsed(null); setImportError(""); } }}>
        <DialogContent className="app-dialog import-dialog"><DialogHeader><div className="dialog-icon"><FileSpreadsheet size={21} /></div><DialogTitle>Excel-ից ներմուծել</DialogTitle><DialogDescription>Ընտրեք .xlsx, .xls կամ .csv ֆայլ։ Առաջին թերթի տվյալները կներմուծվեն։</DialogDescription></DialogHeader>
          <div className="import-content">
            <div className="format-box"><strong>Պահանջվող սյունակները</strong><div className="column-chips"><span>Աշխատակից</span><span>Ամիս</span><span>Առաջադրանքներ</span><span>Դիտարկումներ</span></div><p>Ամիսը գրեք <code>2026-09</code> ձևաչափով։ Յուրաքանչյուր տողը մեկ աշխատակցի մեկ ամսվա արդյունքն է։</p></div>
            <button className="template-link" onClick={() => void downloadWorkbook([{ memberName: "Անի Սարգսյան", month: armeniaMonth(), tasks: 12, observations: 5 }], "timi-arajyntac-orinak.xlsx")}><ArrowDownToLine size={17} /> Ներբեռնել օրինակ ֆայլը</button>
            <input ref={fileInput} className="sr-only" type="file" accept=".xlsx,.xls,.csv" onChange={e => void handleFile(e.target.files?.[0])} aria-label="Ընտրել Excel ֆայլ" />
            <button className="upload-zone" onClick={() => fileInput.current?.click()}><UploadCloud size={29} /><strong>{reading ? "Ֆայլը կարդացվում է…" : parsed ? parsed.filename : "Ընտրել Excel ֆայլը"}</strong><span>Սեղմեք ֆայլ ընտրելու համար · առավելագույնը 8 ՄԲ</span></button>
            {importError && <p className="import-error" role="alert">{importError}</p>}
            {parsed && !parsed.errors.length && <div className="import-preview"><div><Check size={17} /> <strong>{parsed.rows.length} տող պատրաստ է ներմուծման</strong></div><p>Նույն աշխատակցի նույն ամսվա տվյալները կթարմացվեն։</p><div className="preview-lines">{parsed.rows.slice(0, 3).map((row, index) => <span key={index}>{row.memberName} · {row.month} · {row.tasks} / {row.observations}</span>)}</div></div>}
          </div>
          <DialogFooter><Button variant="outline" onClick={() => setImportOpen(false)}>Փակել</Button><Button className="modal-primary" disabled={!parsed || !!parsed.errors.length || saving || reading} onClick={() => void importRows()}>{saving ? <LoaderCircle className="spin" /> : <UploadCloud />} Ներմուծել</Button></DialogFooter>
        </DialogContent>
      </Dialog>

      <Dialog open={helpOpen} onOpenChange={setHelpOpen}><DialogContent className="app-dialog help-dialog"><DialogHeader><div className="dialog-icon"><CircleHelp size={21} /></div><DialogTitle>Ինչպես օգտվել</DialogTitle><DialogDescription>Կոդ գրել պետք չէ. ամեն ինչ կատարվում է այստեղ։</DialogDescription></DialogHeader>
        <ol className="help-steps"><li><span>01</span><div><strong>Ավելացրեք տվյալներ</strong><p>Սեղմեք «Ավելացնել գրառում», գրեք աշխատակցի անունը, ամիսը և երկու ցուցանիշները։</p></div></li><li><span>02</span><div><strong>Կամ ներմուծեք Excel</strong><p>Ներբեռնեք օրինակ ֆայլը, լրացրեք սյունակները և ներմուծեք այն։ Նույն անունն ու ամիսը կրկին ներմուծելիս թվերը կթարմացվեն։</p></div></li><li><span>03</span><div><strong>Պահպանեք Excel պատճենը</strong><p>Տվյալները մնում են այս բրաուզերում։ Սեղմեք «Պահուստային Excel»՝ պատճենը պահելու կամ այլ սարք տեղափոխելու համար։</p></div></li></ol>
        <Button className="modal-primary help-close" onClick={() => setHelpOpen(false)}>Հասկացա</Button>
      </DialogContent></Dialog>
      {toast && <div className="toast" role="status"><Check size={17} /> {toast}<button onClick={() => setToast("")} aria-label="Փակել ծանուցումը"><X size={16} /></button></div>}
    </div>
  );
}
