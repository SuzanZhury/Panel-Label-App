import { useMemo, useRef, useState, type ChangeEvent, type DragEvent, type ReactNode } from 'react';
import { QueryClient, QueryClientProvider } from '@tanstack/react-query';
import { ErrorBoundary } from '@/components/error-boundary';
import { Toaster } from '@/components/ui/toaster';
import { TooltipProvider } from '@/components/ui/tooltip';
import NotFound from '@/pages/not-found';
import * as XLSX from 'xlsx';
import { ArrowDown, ArrowUp, Check, ChevronLeft, ChevronRight, FileSpreadsheet, Info, LockKeyhole, Printer, RotateCcw, Upload, X, AlertCircle, Layers3 } from 'lucide-react';
import { createOriginalExcelLabelWorkbook } from '@/lib/excelExport';
import {
  Route,
  Switch,
  useLocation,
  Router as WouterRouter,
} from 'wouter';

const queryClient = new QueryClient();

type PanelRow = { panelLabel: string; quantity: number; panelType?: string; dim1?: string; dim2?: string };
type LabelSettings = { project: string; release: string; finish: string; arrowDirection: 'up' | 'down' };
type ExpandedLabel = Omit<PanelRow, 'quantity'>;

const emptySettings: LabelSettings = { project: '', release: '', finish: '', arrowDirection: 'up' };
const cleanCell = (value: unknown) => {
  if (value === null || value === undefined) return '';
  const text = String(value).trim();
  return /^#(VALUE|REF|N\/A|NAME\?|DIV\/0|NUM|NULL)!?$/i.test(text) ? '' : text;
};
const normalized = (value: unknown) => cleanCell(value).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

function getFirstSheetSettings(
  sheet: XLSX.WorkSheet,
  headerIndex: number,
): Partial<LabelSettings> {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '', raw: false }) as unknown[][];
  const metadataRows = matrix.slice(0, headerIndex);
  const valueAfterLabel = (matches: (key: string) => boolean) => {
    for (const row of metadataRows) {
      for (let columnIndex = 0; columnIndex < row.length; columnIndex++) {
        if (!matches(normalized(row[columnIndex]))) continue;
        const value = row.slice(columnIndex + 1).map(cleanCell).find(Boolean);
        if (value) return value;
      }
    }
    return undefined;
  };

  const projectName = valueAfterLabel(key => key === 'project name' || key === 'job name');
  const projectNumber = valueAfterLabel(key => key === 'project number' || key === 'job number');
  const standaloneProject = valueAfterLabel(key => key === 'project' || key === 'job' || key === 'project job');
  const releaseValue = valueAfterLabel(key => key === 'release' || key === 'release number');
  const project = projectNumber && projectName
    ? `${projectNumber} - ${projectName}`
    : projectNumber || projectName || standaloneProject;

  const finishLabels = new Set(['finish', 'color', 'colour']);
  const nonFinishValues = new Set(['size', 'lf pcs', 'accessory', 'finish', 'color', 'colour']);
  let finish: string | undefined;
  for (let rowIndex = 0; rowIndex < metadataRows.length && !finish; rowIndex++) {
    const row = metadataRows[rowIndex];
    for (let columnIndex = 0; columnIndex < row.length && !finish; columnIndex++) {
      if (!finishLabels.has(normalized(row[columnIndex]))) continue;
      const candidates = [
        row[columnIndex + 1],
        metadataRows[rowIndex + 1]?.[columnIndex],
        metadataRows[rowIndex - 1]?.[columnIndex + 1],
        metadataRows[rowIndex - 1]?.[columnIndex],
        metadataRows[rowIndex + 1]?.[columnIndex + 1],
      ];
      finish = candidates.map(cleanCell).find(value => value && !nonFinishValues.has(normalized(value)));
    }
  }

  const settings: Partial<LabelSettings> = {};
  if (project) settings.project = project;
  if (releaseValue) {
    settings.release = /^release\b/i.test(releaseValue) ? releaseValue : `Release ${releaseValue}`;
  }
  if (finish) settings.finish = finish;
  return settings;
}

function parseWorkbook(data: ArrayBuffer): { rows: PanelRow[]; settings: Partial<LabelSettings> } {
  const workbook = XLSX.read(data, { type: 'array', cellDates: false });
  const inputSheetName = workbook.SheetNames[0];
  const sheet = inputSheetName ? workbook.Sheets[inputSheetName] : undefined;
  if (!sheet || !inputSheetName) throw new Error('This workbook does not contain a worksheet to read.');
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: '', raw: false }) as unknown[][];
  const requiredHeaders = ['Release #', 'Panel Label', 'Qty', 'Panel Type', 'Dim 1', 'Dim 2'].map(normalized);
  const headerIndex = matrix.findIndex(row => {
    const rowHeaders = row.map(normalized);
    return requiredHeaders.every(required => rowHeaders.includes(required));
  });
  if (headerIndex < 0) throw new Error(`Could not find the required panel columns on the first worksheet (“${inputSheetName}”). Expected: Release #, Panel Label, Qty, Panel Type, Dim 1, Dim 2.`);
  const headers = matrix[headerIndex].map(normalized);
  const column = (name: string) => headers.indexOf(normalized(name));
  const labelCol = column('Panel Label');
  const qtyCol = column('Qty');
  const typeCol = column('Panel Type');
  const dim1Col = column('Dim 1');
  const dim2Col = column('Dim 2');
  const rows: PanelRow[] = [];
  const issues: string[] = [];
  matrix.slice(headerIndex + 1).forEach((line, i) => {
    const panelLabel = cleanCell(line[labelCol]);
    const qtyRaw = cleanCell(line[qtyCol]);
    if (!panelLabel) {
      const hasOtherValues = line.some((value, columnIndex) => columnIndex !== qtyCol && Boolean(cleanCell(value)));
      if (!hasOtherValues) return;
      issues.push(`Row ${headerIndex + i + 2}: panel label is missing.`);
      return;
    }
    if (!qtyRaw || !/^\d+$/.test(qtyRaw) || Number(qtyRaw) < 1) { issues.push(`Row ${headerIndex + i + 2}: quantity for “${panelLabel}” must be a positive whole number.`); return; }
    rows.push({
      panelLabel, quantity: Number(qtyRaw),
      panelType: cleanCell(line[typeCol]) || undefined,
      dim1: cleanCell(line[dim1Col]) || undefined,
      dim2: cleanCell(line[dim2Col]) || undefined,
    });
  });
  if (issues.length) {
    const detail = issues.slice(0, 3).join(' ');
    const remainder = issues.length > 3 ? ` And ${issues.length - 3} more row issue(s).` : '';
    throw new Error(`The workbook has panel rows that need correction. ${detail}${remainder}`);
  }
  if (!rows.length) {
    throw new Error(`No usable panel rows were found on the first worksheet (“${inputSheetName}”).`);
  }
  return { rows, settings: getFirstSheetSettings(sheet, headerIndex) };
}

function DirectionArrow({ direction, size = 20 }: { direction: 'up' | 'down'; size?: number }) {
  return direction === 'up' ? <ArrowUp size={size} strokeWidth={2.7} /> : <ArrowDown size={size} strokeWidth={2.7} />;
}

function PanelLabel({ label, settings, index, print = false }: { label: ExpandedLabel; settings: LabelSettings; index: number; print?: boolean }) {
  return <div className={print ? 'print-label' : 'label-card'} data-testid={print ? `print-label-${index}` : `preview-label-${index}`}>
    <div className="label-content">
      <img className={print ? '' : 'label-logo'} src="/panel-logo.png" alt="Parker Johnston" />
      <div className={print ? 'print-project' : 'label-project'}>{settings.project}</div>
      <div className={print ? 'print-release' : 'label-release'}>{settings.release}</div>
      <div className={print ? 'print-panel' : 'label-panel'}>PANEL:{label.panelLabel}</div>
      <div className={print ? 'print-finish' : 'label-finish'}>{settings.finish}</div>
    </div>
    <div className={print ? 'print-arrow' : 'label-arrow'}><DirectionArrow direction={settings.arrowDirection} size={print ? 25 : 20} /></div>
    {!print && <span className="label-index">{index + 1}</span>}
  </div>;
}

function Home() {
  const fileInput = useRef<HTMLInputElement>(null);
  const [rows, setRows] = useState<PanelRow[]>([]);
  const [settings, setSettings] = useState<LabelSettings>(emptySettings);
  const [fileName, setFileName] = useState('');
  const [error, setError] = useState('');
  const [excelExportError, setExcelExportError] = useState('');
  const [excelExportSuccess, setExcelExportSuccess] = useState('');
  const [excelExporting, setExcelExporting] = useState(false);
  const [rowNotes, setRowNotes] = useState<string[]>([]);
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [previewPage, setPreviewPage] = useState(1);
  const labels = useMemo(() => rows.flatMap(row => Array.from({ length: row.quantity }, () => ({ panelLabel: row.panelLabel, panelType: row.panelType, dim1: row.dim1, dim2: row.dim2 }))), [rows]);
  const pages = Math.ceil(labels.length / 10);
  const pageStartIndex = (previewPage - 1) * 10;
  const visibleLabels = labels.slice(pageStartIndex, pageStartIndex + 10);
  const setField = (key: keyof LabelSettings, value: string) => setSettings(current => ({ ...current, [key]: value }));

  const readFile = async (file?: File) => {
    if (!file) return;
    setError('');
    setExcelExportError('');
    setExcelExportSuccess('');
    setRowNotes([]);
    if (!/\.xlsx$/i.test(file.name)) { setError('Please choose an .xlsx workbook. Other file types are not supported.'); return; }
    setLoading(true);
    try {
      const parsed = parseWorkbook(await file.arrayBuffer());
      setRows(parsed.rows);
      setPreviewPage(1);
      setSettings({ ...emptySettings, ...parsed.settings });
      setFileName(file.name);
      const knownKeys = ['project', 'release', 'finish'] as const;
      const missing = knownKeys.filter(key => !parsed.settings[key]);
      setRowNotes(missing.length ? [`No ${missing.map(key => key === 'project' ? 'project/job' : key).join(', ')} value found in the workbook; fill it in below.`] : []);
    } catch (caught) {
      setRows([]);
      setPreviewPage(1);
      setFileName('');
      setError(caught instanceof Error ? caught.message : 'The workbook could not be read. Check that it is a valid .xlsx file.');
    } finally {
      setLoading(false);
      if (fileInput.current) fileInput.current.value = '';
    }
  };
  const onChange = (event: ChangeEvent<HTMLInputElement>) => { void readFile(event.target.files?.[0]); };
  const onDrop = (event: DragEvent<HTMLDivElement>) => { event.preventDefault(); setDragging(false); void readFile(event.dataTransfer.files?.[0]); };
  const clearWorkbook = () => { setRows([]); setPreviewPage(1); setFileName(''); setSettings(emptySettings); setError(''); setExcelExportError(''); setExcelExportSuccess(''); setRowNotes([]); };
  const previewOrder = visibleLabels.map((_, i) => ({ row: i % 5 + 1, col: Math.floor(i / 5) + 1 }));
  const printPages = Array.from({ length: pages }, (_, pageIndex) => labels.slice(pageIndex * 10, pageIndex * 10 + 10));
  const print = () => window.print();
  const downloadOriginalExcel = async () => {
    if (!labels.length) return;
    setExcelExporting(true);
    setExcelExportError('');
    setExcelExportSuccess('');
    try {
      const baseUrl = new URL(import.meta.env.BASE_URL, window.location.origin);
      const templateUrl = new URL('panel-label-template.xlsx', baseUrl);
      const response = await fetch(templateUrl);
      if (!response.ok) throw new Error('The original Excel label template could not be loaded.');
      const workbookBytes = createOriginalExcelLabelWorkbook(
        await response.arrayBuffer(),
        labels,
        settings,
      );
      const workbookBuffer = workbookBytes.slice().buffer as ArrayBuffer;
      const blob = new Blob([workbookBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const downloadUrl = URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      const fileBase = [settings.project, settings.release, 'Panel Labels']
        .filter(Boolean)
        .join(' - ')
        .replace(/[\\/:*?"<>|]/g, '-')
        .replace(/\s+/g, ' ')
        .trim();
      anchor.href = downloadUrl;
      anchor.download = `${fileBase || 'Panel Labels'}.xlsx`;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
      setExcelExportSuccess(`Excel labels downloaded: ${labels.length} labels; unused template positions removed.`);
    } catch (caught) {
      setExcelExportError(caught instanceof Error ? caught.message : 'The Excel label workbook could not be created.');
    } finally {
      setExcelExporting(false);
    }
  };

  return <div className="app-shell">
    <header className="topbar">
      <div className="brand"><img src="/panel-logo.png" alt="Parker Johnston" /><div className="brand-divider" /><span className="brand-caption">Panel label maker</span></div>
      <div className="secure-note"><span className="secure-dot" /> Your workbook stays on this device <LockKeyhole size={14} /></div>
    </header>
    <main className="main">
      <div className="intro">
        <div><div className="eyebrow">Production tools / labels</div><h1>Make every panel<br className="desktop-break" /> easy to identify.</h1><p>Turn your panel schedule into a print-ready sheet of shop labels.</p></div>
        <div className="step-chip"><strong>01</strong> Import workbook <span>—</span> <strong>02</strong> Review <span>—</span> <strong>03</strong> Print</div>
      </div>
      <div className="workspace">
        <section>
          <div className="panel" data-testid="section-workbook">
            <div className="panel-head"><div className="section-title"><span className="section-number">01</span><h2>Workbook & details</h2></div><span className="section-sub">LOCAL IMPORT</span></div>
            <div className="panel-body">
              <div className={`dropzone ${dragging ? 'dragging' : ''}`} data-testid="dropzone-workbook" role="button" tabIndex={0} onClick={() => fileInput.current?.click()} onKeyDown={event => { if (event.key === 'Enter' || event.key === ' ') fileInput.current?.click(); }} onDragOver={event => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={onDrop}>
                <input ref={fileInput} className="file-input" type="file" accept=".xlsx" data-testid="input-workbook" onChange={onChange} />
                <div className="upload-icon">{loading ? <RotateCcw size={19} className="spin" /> : <Upload size={19} />}</div>
                <strong>{loading ? 'Reading workbook…' : 'Drop your Excel workbook here'}</strong>
                <span>or <span className="browse">browse files</span> · .xlsx only</span>
              </div>
              {fileName && <div className="workbook-row" data-testid="status-workbook">
                <div className="file-mark"><FileSpreadsheet size={18} /></div><div className="file-info"><b>{fileName}</b><span>{rows.length} source {rows.length === 1 ? 'row' : 'rows'} · {labels.length} labels expanded</span></div>
                <button className="icon-button" type="button" aria-label="Remove workbook" data-testid="button-remove-workbook" onClick={clearWorkbook}><X size={17} /></button>
              </div>}
              {error && <div className="alert error" role="alert" data-testid="status-error"><AlertCircle size={16} /> <span>{error}</span></div>}
              {rowNotes.map((note, i) => <div className="alert info" key={i} data-testid="status-metadata"><Info size={15} /><span>{note}</span></div>)}
              <div style={{ margin: '23px 0 18px', borderTop: '1px solid #e9e6de' }} />
              <div className="fields">
                <div className="field full"><label htmlFor="project-input">Project / job</label><input id="project-input" data-testid="input-project" value={settings.project} placeholder="e.g. 21.202589 - Block 22" onChange={e => setField('project', e.target.value)} /></div>
                <div className="field"><label htmlFor="release-input">Release</label><input id="release-input" data-testid="input-release" value={settings.release} placeholder="e.g. Release 49" onChange={e => setField('release', e.target.value)} /></div>
                <div className="field"><label htmlFor="finish-input">Finish / color</label><input id="finish-input" data-testid="input-finish" value={settings.finish} placeholder="e.g. Bone White" onChange={e => setField('finish', e.target.value)} /></div>
                  <div className="field full"><label>Browser print arrow direction <span>DEFAULT: UP</span></label><div className="direction-options">
                  <button className={`direction-button ${settings.arrowDirection === 'up' ? 'selected' : ''}`} type="button" data-testid="button-arrow-up" aria-pressed={settings.arrowDirection === 'up'} onClick={() => setField('arrowDirection', 'up')}><ArrowUp size={16} /> Up</button>
                  <button className={`direction-button ${settings.arrowDirection === 'down' ? 'selected' : ''}`} type="button" data-testid="button-arrow-down" aria-pressed={settings.arrowDirection === 'down'} onClick={() => setField('arrowDirection', 'down')}><ArrowDown size={16} /> Down</button>
                </div></div>
              </div>
              <div className="meta-note"><Check size={13} /> Details apply to both exports; the Excel file keeps the template arrows.</div>
            </div>
          </div>
        </section>
        <section>
          <div className="panel" data-testid="section-preview">
            <div className="panel-head"><div className="section-title"><span className="section-number">02</span><h2>Review labels</h2></div><div className="preview-top"><Layers3 size={15} color="#668579" /><span className="count-pill" data-testid="text-label-count">{labels.length} LABEL{labels.length === 1 ? '' : 'S'}</span></div></div>
            <div className="preview-wrap">
              <div className="preview-page">
                <div className="preview-grid" data-testid="content-label-preview">
                  {visibleLabels.length ? visibleLabels.map((label, i) => <div key={`${label.panelLabel}-${pageStartIndex + i}`} style={{ gridRow: previewOrder[i].row, gridColumn: previewOrder[i].col }}><PanelLabel label={label} settings={settings} index={pageStartIndex + i} /></div>) : <div className="preview-placeholder" data-testid="empty-label-preview"><div className="placeholder-mark"><FileSpreadsheet size={24} /></div><b>Your labels will appear here</b><span>Import a workbook to review the layout before printing.</span></div>}
                </div>
              </div>
              <div className="preview-footer">
                <span>Letter · Portrait · 2 × 5 labels per page</span>
                {pages ? <div className="preview-pagination" aria-label="Label preview pages">
                  <button type="button" className="preview-page-button" data-testid="button-preview-previous" aria-label="Previous preview page" disabled={previewPage <= 1} onClick={() => setPreviewPage(page => Math.max(1, page - 1))}><ChevronLeft size={15} /></button>
                  <b className="preview-page-count" data-testid="text-preview-page" aria-live="polite">Page {previewPage} of {pages}</b>
                  <button type="button" className="preview-page-button" data-testid="button-preview-next" aria-label="Next preview page" disabled={previewPage >= pages} onClick={() => setPreviewPage(page => Math.min(pages, page + 1))}><ChevronRight size={15} /></button>
                </div> : <b>Awaiting workbook</b>}
              </div>
            </div>
          </div>
          <div className="print-action">
            <div className="print-copy"><strong>Use the original Excel artwork</strong><span>{labels.length ? `${labels.length} labels · Original logo, arrows, layout, and print settings` : 'Add a workbook to prepare your label file.'}</span></div>
            <div className="label-actions">
              <button className="excel-button" type="button" data-testid="button-download-excel" disabled={!labels.length || excelExporting} onClick={() => void downloadOriginalExcel()}>
                <FileSpreadsheet size={16} /> {excelExporting ? 'Preparing Excel…' : 'Download Excel labels'}
              </button>
              <button className="print-button" type="button" data-testid="button-print" disabled={!labels.length} onClick={print}><Printer size={16} /> Browser print</button>
            </div>
          </div>
          {excelExportError && <div className="alert error" role="alert" data-testid="status-excel-error"><AlertCircle size={16} /> <span>{excelExportError}</span></div>}
          {excelExportSuccess && <div className="alert info" role="status" data-testid="status-excel-success"><Check size={15} /><span>{excelExportSuccess}</span></div>}
          <div className="process-note"><Info size={14} /><span>Excel download uses the original label artwork, keeps only the labels needed for this quantity, and preserves the saved print settings. Browser print opens your browser dialog. Your uploaded workbook is read on this device and is never uploaded.</span></div>
        </section>
      </div>
    </main>
    <footer className="footer"><span>PARKER JOHNSTON · PANEL IDENTIFICATION</span><span>LOCAL WORKFLOW / NO UPLOAD</span></footer>
    <div id="print-sheet" data-testid="content-print-sheet" aria-hidden="true">
      {printPages.map((page, pageIndex) => <div className="print-page" key={pageIndex} data-testid={`print-page-${pageIndex + 1}`}>
        {page.map((label, i) => <div key={`${label.panelLabel}-${pageIndex * 10 + i}`} style={{ gridRow: i % 5 + 1, gridColumn: Math.floor(i / 5) + 1 }}><PanelLabel label={label} settings={settings} index={pageIndex * 10 + i} print /></div>)}
      </div>)}
    </div>
  </div>;
}

function Router() {
  return (
    // Keep a shared shell (sidebar, navbar) outside the boundary so it
    // survives a page crash.
    <RoutedErrorBoundary>
      <Switch>
        <Route path="/" component={Home} />
        <Route component={NotFound} />
      </Switch>
    </RoutedErrorBoundary>
  );
}

function RoutedErrorBoundary({ children }: { children: ReactNode }) {
  const [location] = useLocation();
  return <ErrorBoundary resetKey={location}>{children}</ErrorBoundary>;
}

function App() {
  return (
    <QueryClientProvider client={queryClient}>
      <TooltipProvider>
        <WouterRouter base={import.meta.env.BASE_URL.replace(/\/$/, '')}>
          <Router />
        </WouterRouter>
        <Toaster />
      </TooltipProvider>
    </QueryClientProvider>
  );
}

export default App;
