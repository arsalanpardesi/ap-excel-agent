import { Cell, Sheet, Workbook, RangeRef, SheetEvent, CellProvenance } from './types.js';

function uuid() { return Math.random().toString(36).slice(2) + Date.now().toString(36); }

export class SheetModel {
  wb: Workbook;

  constructor() {
    this.wb = { id: uuid(), sheets: new Map(), checkpoints: [], events: [] };
    // Start with a single clean sheet
    this.createSheet('Sheet1');
    this.checkpoint('init');
  }

  // ===== Helpers =====
  private getSheet(name: string): Sheet {
    const s = this.wb.sheets.get(name);
    if (!s) throw new Error(`Sheet not found: ${name}`);
    return s;
  }

  private ensureSize(sheetName: string, rows: number, cols: number) {
    const s = this.getSheet(sheetName);
    while (s.rows.length < rows) s.rows.push([]);
    for (const r of s.rows) while (r.length < cols) r.push({ value: null });
  }

  static a1ToRc(ref: string): { r: number; c: number } {
    const m = ref.match(/^([A-Za-z]+)([0-9]+)$/);
    if (!m) throw new Error(`Bad A1 ref: ${ref}`);
    const colStr = m[1].toUpperCase();
    const row = parseInt(m[2], 10) - 1;
    let col = 0;
    for (let i = 0; i < colStr.length; i++) col = col * 26 + (colStr.charCodeAt(i) - 64);
    return { r: row, c: col - 1 };
  }

  static rcToA1(r: number, c: number): string {
    let n = c + 1, s = '';
    while (n > 0) { const rem = (n - 1) % 26; s = String.fromCharCode(65 + rem) + s; n = Math.floor((n - 1) / 26);} 
    return s + (r + 1);
  }

  private snapshot(range: RangeRef): Cell[][] {
    const s = this.getSheet(range.sheet);
    this.ensureSize(range.sheet, range.r2 + 1, range.c2 + 1);
    const out: Cell[][] = [];
    for (let r = range.r1; r <= range.r2; r++) {
      const row: Cell[] = [];
      for (let c = range.c1; c <= range.c2; c++) row.push({ ...s.rows[r][c] });
      out.push(row);
    }
    return out;
  }

  // ===== Events / Checkpoints =====
  private appendEvent(op: string, args: any, inverse: { op: string; args: any }, summary: string) {
    const ev: SheetEvent = { id: uuid(), ts: Date.now(), op, args, inverse, summary };
    this.wb.events.push(ev);
    return ev;
  }

  checkpoint(id?: string) {
    const cp = { id: id ?? uuid(), atEvent: this.wb.events.length, ts: Date.now() };
    this.wb.checkpoints.push(cp);
    return cp;
  }

  undo() {
    const ev = this.wb.events.pop();
    if (!ev) return null;
    this.dispatch(ev.inverse.op, ev.inverse.args, true);
    return ev;
  }

  // ===== SheetOps =====
  createSheet(name: string) {
    if (this.wb.sheets.has(name)) throw new Error('Sheet exists');
    this.wb.sheets.set(name, { name, rows: [] });
    this.appendEvent('createSheet', { name }, { op: 'deleteSheet', args: { name } }, `create ${name}`);
  }

  deleteSheet(name: string) {
    const s = this.getSheet(name);
    this.wb.sheets.delete(name);
    this.appendEvent('deleteSheet', { name }, { op: 'restoreSheet', args: { sheet: s } }, `delete ${name}`);
  }

  restoreSheet(sheet: Sheet) {
    this.wb.sheets.set(sheet.name, sheet);
    this.appendEvent('restoreSheet', { sheet }, { op: 'deleteSheet', args: { name: sheet.name } }, `restore ${sheet.name}`);
  }

  setValues(range: RangeRef, values: (string|number|null)[][], provenance?: CellProvenance[]) {
    const before = this.snapshot(range);
    const s = this.getSheet(range.sheet);
    this.ensureSize(range.sheet, range.r2 + 1, range.c2 + 1);
    for (let r=0; r<values.length; r++) {
      for (let c=0; c<values[r].length; c++) {
        const rr = range.r1 + r, cc = range.c1 + c;
        const prev = s.rows[rr][cc] ?? { value: null };
        s.rows[rr][cc] = {
          ...prev,
          value: values[r][c],
          formula: undefined,
          provenance: provenance ? [...provenance] : prev.provenance
        };
      }
    }
    this.appendEvent('setValues', { range, values, provenance }, { op: 'setCells', args: { range, cells: before } }, `setValues ${range.sheet}!${SheetModel.rcToA1(range.r1, range.c1)}`);
  }

  setCells(range: RangeRef, cells: Cell[][]) {
    const s = this.getSheet(range.sheet);
    this.ensureSize(range.sheet, range.r2 + 1, range.c2 + 1);
    for (let r=0; r<cells.length; r++) for (let c=0; c<cells[r].length; c++) s.rows[range.r1+r][range.c1+c] = cells[r][c];
    this.appendEvent('setCells', { range, cells }, { op: 'noop', args: {} }, 'setCells');
  }

  setFormulas(range: RangeRef, formulas: (string|null)[][]) {
    const before = this.snapshot(range);
    const s = this.getSheet(range.sheet);
    this.ensureSize(range.sheet, range.r2 + 1, range.c2 + 1);
    for (let r=0; r<formulas.length; r++) {
      for (let c=0; c<formulas[r].length; c++) {
        const rr = range.r1 + r, cc = range.c1 + c;
        const prev = s.rows[rr][cc] ?? { value: null };
        const f = formulas[r][c] ?? undefined;
        s.rows[rr][cc] = { ...prev, value: null, formula: f };
      }
    }
    this.appendEvent('setFormulas', { range, formulas }, { op: 'setCells', args: { range, cells: before } }, `setFormulas ${range.sheet}!${SheetModel.rcToA1(range.r1, range.c1)}`);
  }

  formatRange(range: RangeRef, format: string | null) {
    const before = this.snapshot(range);
    const s = this.getSheet(range.sheet);
    this.ensureSize(range.sheet, range.r2 + 1, range.c2 + 1);
    for (let r = range.r1; r <= range.r2; r++) {
      for (let c = range.c1; c <= range.c2; c++) {
        const prev = s.rows[r][c] ?? { value: null };
        s.rows[r][c] = { ...prev, format: format ?? undefined };
      }
    }
    this.appendEvent('formatRange', { range, format }, { op: 'setCells', args: { range, cells: before } }, `formatRange ${range.sheet} ${format ?? 'clear'}`);
  }

  linkProvenance(range: RangeRef, prov: CellProvenance[]) {
    const before = this.snapshot(range);
    const s = this.getSheet(range.sheet);
    for (let r = range.r1; r <= range.r2; r++) for (let c = range.c1; c <= range.c2; c++) {
      const cell = s.rows[r][c];
      if (!cell.provenance) cell.provenance = [];
      cell.provenance.push(...prov);
    }
    this.appendEvent('linkProvenance', { range, prov }, { op: 'setCells', args: { range, cells: before } }, `linkProvenance ${range.sheet}`);
  }

  // ===== Evaluation =====
  evaluateAll() {
    for (const s of this.wb.sheets.values()) {
      for (let r=0; r<s.rows.length; r++) for (let c=0; c<s.rows[r].length; c++) this.evaluateCell(s.name, r, c, new Set());
    }
  }

  private getCell(sheet: string, r: number, c: number): Cell { this.ensureSize(sheet, r+1, c+1); return this.getSheet(sheet).rows[r][c]; }

  private evalRef(sheet: string, a1: string, seen: Set<string>): number {
    const { r, c } = SheetModel.a1ToRc(a1);
    const v = this.evaluateCell(sheet, r, c, seen);
    if (typeof v === 'number') return v; if (v===null) return 0; const n = Number(v); return isNaN(n) ? 0 : n;
  }

  private sumTerm(sheet: string, term: string, seen: Set<string>): number {
    term = term.trim();
    if (/^[A-Za-z]+[0-9]+:[A-Za-z]+[0-9]+$/.test(term)) return this.sumRange(sheet, term, seen);
    if (/^[A-Za-z]+[0-9]+$/.test(term)) return this.evalRef(sheet, term, seen);
    const n = Number(term.replace(/,/g, ''));
    return Number.isFinite(n) ? n : 0;
  }

  private sumArgs(sheet: string, args: string, seen: Set<string>): number {
    return args.split(',').map(t => this.sumTerm(sheet, t, seen)).reduce((a,b)=>a+b, 0);
  }

  private sumRange(sheet: string, a1range: string, seen: Set<string>): number {
    const m = a1range.match(/^([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$/i);
    if (!m) return 0;
    const a = SheetModel.a1ToRc(m[1]+m[2]);
    const b = SheetModel.a1ToRc(m[3]+m[4]);
    let total = 0;
    for (let r = Math.min(a.r, b.r); r <= Math.max(a.r, b.r); r++)
      for (let c = Math.min(a.c, b.c); c <= Math.max(a.c, b.c); c++) {
        const v = this.evaluateCell(sheet, r, c, seen);
        const n = typeof v === 'number' ? v : Number(v);
        if (!isNaN(n)) total += n;
      }
    return total;
  }

  evaluateCell(sheet: string, r: number, c: number, seen = new Set<string>()): string | number | null {
    const key = `${sheet}:${r}:${c}`;
    if (seen.has(key)) return '#REF!'; // Return error string for circular reference
    seen.add(key);
    const cell = this.getCell(sheet, r, c);
    if (!cell.formula) return cell.value ?? null;
    const f = cell.formula.trim();
    if (!f.startsWith('=')) return cell.value ?? null;
    let expr = f.slice(1);

    expr = expr.replace(/SUM\(\s*([^)]+?)\s*\)/gi, (_m, args) => String(this.sumArgs(sheet, args, seen)));
    expr = expr.replace(/\b([A-Za-z]+[0-9]+)\b/g, (_m, a1) => String(this.evalRef(sheet, a1, seen)));

    try {
      const val = Function(`"use strict"; return (${expr})`)();
      return typeof val === 'number' ? val : Number(val) || 0;
    } catch {
      return '#ERROR!';
    }
  }

  // ===== Dispatch (for generic tool runner) =====
  dispatch(op: string, args: any, internal = false) {
    switch (op) {
      case 'createSheet': return this.createSheet(args.name);
      case 'deleteSheet': return this.deleteSheet(args.name);
      case 'restoreSheet': return this.restoreSheet(args.sheet);
      case 'setValues': return this.setValues(args.range, args.values, args.provenance);
      case 'setFormulas': return this.setFormulas(args.range, args.formulas);
      case 'formatRange': return this.formatRange(args.range, args.format ?? null);
      case 'setCells': return this.setCells(args.range, args.cells);
      case 'linkProvenance': return this.linkProvenance(args.range, args.prov ?? args.provenance ?? []);
      case 'noop': return;
      default: if (!internal) throw new Error(`Unknown op ${op}`);
    }
  }

  // ===== API helpers =====
  toJSON() {
    return {
      id: this.wb.id,
      sheets: Array.from(this.wb.sheets.values()).map(s => ({
        name: s.name,
        rows: s.rows.map(row => row.map(c => ({ value: c.value, formula: c.formula ?? null, format: c.format ?? null })))
      })),
      checkpoints: this.wb.checkpoints,
      events: this.wb.events.map(e => ({ id: e.id, ts: e.ts, op: e.op, summary: e.summary }))
    };
  }

  getProvenance(sheet: string, a1: string) {
    const { r, c } = SheetModel.a1ToRc(a1);
    const cell = this.getCell(sheet, r, c);
    return cell.provenance ?? [];
  }

  loadFromJSON(workbookData: { sheets: { name: string; rows: { value: any }[][] }[] }) {
    this.wb.sheets.clear();
    
    for (const s of workbookData.sheets) {
      const newSheet: Sheet = {
        name: s.name,
        // Ensure every cell is a valid Cell object, even if input is sparse
        rows: s.rows.map(row => 
          (row || []).map(cell => ({
            value: cell?.value ?? null,
            formula: undefined,
            format: undefined,
            provenance: undefined
          }))
        )
      };
      this.wb.sheets.set(s.name, newSheet);
    }

    // If no sheets were loaded, create a default one
    if (this.wb.sheets.size === 0) {
      this.createSheet('Sheet1');
    }

    // Reset history
    this.wb.events = [];
    this.wb.checkpoints = [];
    this.checkpoint('import');
    this.evaluateAll();
  }
}