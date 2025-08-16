import { XlsxParser } from "../../../dist/index.js";

const input = document.getElementById("xlsxInput") as HTMLInputElement;
const sheetView = document.getElementById("sheetView")!;
const tabs = document.getElementById("tabs")!;

let parser: XlsxParser | null = null;
let workbook: any = null;
let current = 0;

const render = (index: number) => {
  if (!parser || !workbook) return;
  current = index;
  sheetView.innerHTML = parser.toHTMLSheet(workbook, index);
  // Make sheetView focusable to catch key events
  (sheetView as HTMLElement).setAttribute('tabindex', '0');
  (sheetView as HTMLElement).focus();
  [...tabs.children].forEach((el, i) => {
    el.classList.toggle('active', i === index);
  });
  attachSelectionHandlers();
};

input.addEventListener("change", async () => {
  const file = input.files?.[0];
  if (!file) return;

  const buffer = await file.arrayBuffer();
  parser = new XlsxParser();
  workbook = await parser.readFile(buffer, { dense: true, styles: true, drawings: true, skipHiddenRows: true });
  console.log(workbook)

  // Build tabs at bottom
  tabs.innerHTML = '';
  workbook.workSheets.forEach((ws: any, i: number) => {
    const btn = document.createElement('button');
    btn.className = 'tab';
    btn.textContent = ws.name || `Sheet ${i + 1}`;
    btn.addEventListener('click', () => render(i));
    tabs.appendChild(btn);
  });

  // Show first sheet by default
  render(0);
});

function attachSelectionHandlers() {
  const root = sheetView.querySelector('.xl') as HTMLTableElement | null;
  if (!root) return;
  let anchor: { row: number; col: number } | null = null;
  let draggingCells = false;
  let draggingRows = false;
  let draggingCols = false;
  let selectionRect: { r1: number; c1: number; r2: number; c2: number } | null = null;

  const getRC = (ref: string) => {
    const m = ref.match(/^([A-Z]+)(\d+)$/);
    if (!m) return null;
    const [, colLetters, rowStr] = m;
    let col = 0;
    for (let i = 0; i < colLetters.length; i++) {
      col = col * 26 + (colLetters.charCodeAt(i) - 64);
    }
    return { row: parseInt(rowStr, 10), col };
  };

  const clearSel = () => {
    root.querySelectorAll('td.sel').forEach(el => el.classList.remove('sel'));
    root.querySelectorAll('th.sel-h').forEach(el => el.classList.remove('sel-h'));
  };

  const applySel = (r1: number, c1: number, r2: number, c2: number) => {
    clearSel();
    const rStart = Math.min(r1, r2);
    const rEnd = Math.max(r1, r2);
    const cStart = Math.min(c1, c2);
    const cEnd = Math.max(c1, c2);
    selectionRect = { r1: rStart, c1: cStart, r2: rEnd, c2: cEnd };
    // cells
    for (let r = rStart; r <= rEnd; r++) {
      for (let c = cStart; c <= cEnd; c++) {
        const ref = rcToRef(r, c);
        const td = root.querySelector(`td[data-ref="${ref}"]`);
        if (td) td.classList.add('sel');
      }
    }
    // headers
    if (rStart === 1 && rEnd >= 1 && cStart <= cEnd) {
      // column selection highlight
      for (let c = cStart; c <= cEnd; c++) {
        const th = root.querySelector(`thead th.xl-col[data-col="${c}"]`);
        if (th) th.classList.add('sel-h');
      }
    }
    if (cStart === 1 && cEnd >= 1 && rStart <= rEnd) {
      // row selection highlight
      for (let r = rStart; r <= rEnd; r++) {
        const th = root.querySelector(`tbody tr:nth-child(${r}) > th.xl-row`);
        if (th) th.classList.add('sel-h');
      }
    }
  };

  const rcToRef = (row: number, col: number) => {
    let s = '';
    let n = col;
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return `${s}${row}`;
  };

  // Mouse interactions on cells
  // Start cell drag selection
  root.addEventListener('mousedown', (e) => {
    const td = (e.target as HTMLElement).closest('td[data-ref]') as HTMLTableCellElement | null;
    if (!td) return;
    const rc = getRC(td.dataset.ref || '');
    if (!rc) return;
    anchor = rc;
    draggingCells = true;
    applySel(anchor.row, anchor.col, anchor.row, anchor.col);
    e.preventDefault();
    const onUp = () => {
      draggingCells = false;
      draggingRows = false;
      draggingCols = false;
      anchor = null;
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mouseup', onUp);
  });

  root.addEventListener('mousemove', (e) => {
    if (!anchor) return;
    if (draggingCells) {
      const td = (e.target as HTMLElement).closest('td[data-ref]') as HTMLTableCellElement | null;
      if (!td) return;
      const rc = getRC(td.dataset.ref || '');
      if (!rc) return;
      applySel(anchor.row, anchor.col, rc.row, rc.col);
    } else if (draggingRows) {
      const th = (e.target as HTMLElement).closest('th.xl-row[data-row]') as HTMLTableCellElement | null;
      if (!th) return;
      const r = parseInt(th.dataset.row || '1', 10);
      const cols = root.querySelectorAll('thead th.xl-col').length;
      applySel(anchor.row, 1, r, cols);
    } else if (draggingCols) {
      const th = (e.target as HTMLElement).closest('th.xl-col[data-col]') as HTMLTableCellElement | null;
      if (!th) return;
      const c = parseInt(th.dataset.col || '1', 10);
      const rows = root.querySelectorAll('tbody tr').length;
      applySel(1, anchor.col, rows, c);
    }
  });

  // mouseup listener is attached per drag in mousedown handler

  // Row header drag selection
  root.addEventListener('mousedown', (e) => {
    const th = (e.target as HTMLElement).closest('th.xl-row[data-row]') as HTMLTableCellElement | null;
    if (!th) return;
    const r = parseInt(th.dataset.row || '1', 10);
    anchor = { row: r, col: 1 };
    draggingRows = true;
    const cols = root.querySelectorAll('thead th.xl-col').length;
    applySel(r, 1, r, cols);
    e.preventDefault();
  });

  // Column header drag selection
  root.addEventListener('mousedown', (e) => {
    const th = (e.target as HTMLElement).closest('th.xl-col[data-col]') as HTMLTableCellElement | null;
    if (!th) return;
    const c = parseInt(th.dataset.col || '1', 10);
    anchor = { row: 1, col: c };
    draggingCols = true;
    const rows = root.querySelectorAll('tbody tr').length;
    applySel(1, c, rows, c);
    e.preventDefault();
  });

  // Copy selection to clipboard as TSV on Ctrl/Cmd+C
  sheetView.addEventListener('keydown', async (e: KeyboardEvent) => {
    if (!selectionRect) return;
    const isCopy = (e.key === 'c' || e.key === 'C') && (e.ctrlKey || e.metaKey);
    if (!isCopy) return;
    const rows = [] as string[];
    for (let r = selectionRect.r1; r <= selectionRect.r2; r++) {
      const cols: string[] = [];
      for (let c = selectionRect.c1; c <= selectionRect.c2; c++) {
        const ref = rcToRef(r, c);
        const td = root.querySelector(`td[data-ref="${ref}"]`) as HTMLElement | null;
        const val = td ? (td.textContent || '') : '';
        // escape tabs/newlines within cell
        cols.push(val.replace(/\t/g, ' ').replace(/\r?\n/g, ' '));
      }
      rows.push(cols.join('\t'));
    }
    const tsv = rows.join('\n');
    try {
      await navigator.clipboard.writeText(tsv);
    } catch {
      const ta = document.createElement('textarea');
      ta.value = tsv;
      ta.style.position = 'fixed';
      ta.style.left = '-9999px';
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      ta.remove();
    }
    e.preventDefault();
  });
}
