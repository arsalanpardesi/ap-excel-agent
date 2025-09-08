export type RangeRef = { sheet: string; r1: number; c1: number; r2: number; c2: number };

export type CellProvenance = {
  docId: string; // arbitrary string id (e.g., URL or file id)
  snippet?: string;
  rationale?: string;
};

export type Cell = {
  value: string | number | null;
  formula?: string; // literal formula string starting with '='
  provenance?: CellProvenance[];
  /** Simple formatting: 'percent' | 'currency' | 'number' | 'text' */
  format?: string;
};

export type Sheet = {
  name: string;
  rows: Cell[][]; // rows[r][c]
};

export type Workbook = {
  id: string;
  sheets: Map<string, Sheet>;
  checkpoints: { id: string; atEvent: number; ts: number }[];
  events: SheetEvent[];
};

export type SheetEvent = {
  id: string;
  ts: number;
  op: string;
  args: any;
  inverse: { op: string; args: any };
  summary: string;
};

export type PlanStep = { op: string; args: any; explain?: string };
export type Plan = { steps: PlanStep[]; summary?: string };

/** 10-K parsed structure returned by the LLM */
export type Statement = {
  title: string;
  periods: string[];
  lines: { name: string; values: number[] }[];
  scale?: string;
  currency?: string;
};

export type TenKParsed = {
  income: Statement;
  balance: Statement;
  cashflow: Statement;
  meta?: { source?: string; company?: string; fiscalYearEnd?: string };
};
