import dotenv from 'dotenv';
import { Plan } from './types.js';
import { SheetModel } from './sheet.js';
import { GoogleGenAI } from '@google/genai';

dotenv.config();

// --- Ollama Configuration ---
const OLLAMA_BASE_URL = process.env.OLLAMA_BASE_URL || 'http://localhost:11434';
const OLLAMA_MODEL = process.env.OLLAMA_MODEL || 'qwen3:32b';

// --- Gemini Configuration & Initialization ---
const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
const geminiAI = GEMINI_API_KEY ? new GoogleGenAI({}) : null;

export type AgentHints = { sheetHint?: string; insertRow?: number }; // insertRow is 0-based

/* ---------- Context summarizer sent to the model ---------- */
function getCellText(cell: any) {
  if (!cell) return '';
  return cell.formula ? cell.formula : (cell.value ?? '');
}

function summarizeWorkbook(wb: SheetModel) {
  const out: any = { sheets: [] };

  for (const s of wb.wb.sheets.values()) {
    const rows = s.rows ?? [];
    const rowCount = rows.length;
    const colCount = rows.reduce((m: number, r: any[]) => Math.max(m, r?.length ?? 0), 0);

    const prevRows = Math.min(rowCount, 60);
    const prevCols = Math.min(colCount, 20);
    const previewA1: any[] = [];
    for (let r = 0; r < prevRows; r++) {
      const row: any[] = [];
      for (let c = 0; c < prevCols; c++) row.push(getCellText(rows[r]?.[c]));
      previewA1.push(row);
    }

    const headerA1: any[] = [];
    if (rowCount > 0) {
      for (let c = 0; c < Math.min(colCount, 50); c++) headerA1.push(getCellText(rows[0]?.[c]));
    }

    const labelsA: { row: number; text: string }[] = [];
    const capRows = Math.min(rowCount, 500);
    for (let r = 0; r < capRows; r++) labelsA.push({ row: r, text: String(getCellText(rows[r]?.[0] ?? '')).trim() });

    out.sheets.push({
      name: s.name,
      rows: rowCount,
      cols: colCount,
      headerA1,
      labelsA,
      previewA1
    });
  }
  return out;
}

/* ---------- System prompt / tool contract ---------- */
const SYSTEM_PROMPT = `
You are an AI spreadsheet operator. You receive:
- a user's goal,
- optional hints: { "sheetHint"?: string, "insertRow"?: number },
- and a JSON summary of the workbook, including for each sheet:
  - headerA1: first row cells (helps identify period columns),
  - labelsA: [{row, text}] for up to the first 500 rows of column A (captions),
  - previewA1: a small grid preview.

Return STRICT JSON:
{ "steps": PlanStep[], "summary": string }

PlanStep is exactly one of:
- { "op": "createSheet",   "args": { "name": string } }
- { "op": "setValues",     "args": { "range": { "sheet": string, "r1": number, "c1": number, "r2": number, "c2": number }, "values": (string|number|null)[][] } }
- { "op": "setFormulas",   "args": { "range": { "sheet": string, "r1": number, "c1": number, "r2": number, "c2": number }, "formulas": (string|null)[][] } }
- { "op": "formatRange",   "args": { "range": { "sheet": string, "r1": number, "c1": number, "r2": number, "c2": number }, "format": "percent" | "currency" | "number" | "text" } }
- { "op": "linkProvenance","args": { "range": { "sheet": string, "r1": number, "c1": number, "r2": number, "c2": number }, "prov": [{ "docId": string, "snippet"?: string, "rationale"?: string }] } }

Rules:
- **Indexing is 0-based**: All row/column numbers you receive (in labelsA) and provide (in ranges) are 0-based. HOWEVER, when writing Excel-style formulas (e.g., "=A1+B1"), you MUST convert the 0-based row index to a 1-based row number by adding 1. For example, to reference a cell at row index 4, use "5" in the formula string (e.g., "A5").
- Ensure 2D array sizes for values/formulas exactly match the range dimensions.
- If "sheetHint" is provided, prefer that exact sheet when writing unless it clearly doesn't exist.
- If "insertRow" is provided, place new calculation rows starting at that row index (0-based). Expand the sheet if necessary.
- **Captions**: For any new calculated row, ALWAYS write a descriptive caption in column A (c1==0), e.g., "Gross margin %".
- **Use captions to find inputs**: Use the sheet's labelsA (case-insensitive) to find row indices for:
    - Revenue: ["revenue", "net sales", "sales"]
    - Cost of revenue: ["cost of revenue", "cost of sales", "cogs"]
    - Gross profit: ["gross profit"]
  Prefer exact/starts-with matches, ignoring punctuation/whitespace.
- **Use headerA1** to align period columns (usually headers start at column 1). When writing across years, place formulas in the same data columns as Revenue etc.
- Write formulas, not hard-coded numbers, e.g.:
    - Gross margin % = Gross profit / Revenue
    - Net margin %   = Net income / Revenue
- When calculating percentage metrics, set "formatRange" to "percent" for the data columns you write (do not format the label cell).
- Keep plans <= 30 steps. No prose—return only JSON.
`.trim();

/* ---------- Utilities ---------- */
function messagesToPrompt(messages: { role: string; content: string }[]) {
  return messages.map(m => `${m.role.toUpperCase()}: ${m.content}`).join('\n\n');
}

function cleanToJsonString(raw: string) {
  return String(raw)
    .replace(/^\uFEFF/, '')
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/```$/i, '')
    .trim();
}

/* ---------- Streaming (SSE) ---------- */
export type AgentStreamEvent =
  | { type: 'status'; data: string }
  | { type: 'context'; data: any }
  | { type: 'token'; data: string }
  | { type: 'plan'; data: any }
  | { type: 'error'; data: string }
  | { type: 'done'; data: string };

// --- Gemini Streaming Function ---
async function streamFromGemini(messages: any[], on: (e: AgentStreamEvent) => void): Promise<string> {
    if (!geminiAI) {
      throw new Error('Gemini API key not configured. Please check your .env file.');
    }
  
    const userPrompt = messages.find(m => m.role === 'user')?.content;
    if (!userPrompt) {
        throw new Error('No user content found in messages for Gemini.');
    }
  
    const result = await geminiAI.models.generateContentStream({
      model: "gemini-2.5-pro", // Specify the model here
      contents: [{ role: "user", parts: [{ text: userPrompt }] }],
      config: {
        systemInstruction: [{ text: SYSTEM_PROMPT }],
        responseMimeType: "application/json",
        temperature: 0,
      }
    });
  
    let fullResponse = '';
    for await (const chunk of result) {
      const chunkText = chunk.text;
      if (chunkText) {
        on({ type: 'token', data: chunkText });
        fullResponse += chunkText;
      }
    }
    return fullResponse;
}

// --- Ollama Streaming Function ---
async function streamFromOllama(messages: any[], on: (e: AgentStreamEvent) => void): Promise<string> {
  const bodyChat: any = {
    model: OLLAMA_MODEL,
    messages,
    stream: true,
    format: 'json',
    options: { temperature: 0 }
  };

  let res: Response;
  try {
    res = await fetch(`${OLLAMA_BASE_URL}/api/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(bodyChat)
    });
  } catch (e) {
    throw new Error(`Could not reach Ollama at ${OLLAMA_BASE_URL} – ${String(e)}`);
  }

  if (res.status === 404) {
    const bodyGen: any = {
      model: OLLAMA_MODEL,
      prompt: messagesToPrompt(messages),
      stream: true,
      format: 'json',
      options: { temperature: 0 }
    };
    const r2 = await fetch(`${OLLAMA_BASE_URL}/api/generate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(bodyGen)
    });
    if (!r2.ok) throw new Error(`Ollama /api/generate ${r2.status}`);
    return await streamRead(r2, on, 'generate');
  }

  if (!res.ok) throw new Error(`Ollama /api/chat ${res.status}`);
  return await streamRead(res, on, 'chat');
}

async function streamRead(res: Response, on: (e: AgentStreamEvent)=>void, mode: 'chat'|'generate'): Promise<string> {
  const reader = res.body!.getReader();
  const decoder = new TextDecoder();
  let buf = '';
  let full = '';

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buf += decoder.decode(value, { stream: true });
    const lines = buf.split('\n');
    buf = lines.pop() || '';
    for (const line of lines) {
      const s = line.trim();
      if (!s) continue;
      try {
        const j = JSON.parse(s);
        const token = mode === 'chat' ? (j?.message?.content ?? '') : (j?.response ?? '');
        if (token) { on({ type: 'token', data: token }); full += token; }
        if (j?.done) break;
      } catch {
        // ignore partials
      }
    }
  }
  return full;
}

/* ---------- Public APIs ---------- */
export async function streamPlanAndExecute(
  goal: string,
  wb: SheetModel,
  hints: AgentHints | undefined,
  on: (e: AgentStreamEvent) => void,
  modelChoice: 'ollama' | 'gemini' // Parameter to select the model
) {
  const context = summarizeWorkbook(wb);
  on({ type: 'context', data: context });

  const messages = [
    { role: 'system', content: SYSTEM_PROMPT },
    { role: 'user', content: JSON.stringify({ goal, hints: hints ?? {}, context }) }
  ];

  on({ type: 'status', data: `contacting ${modelChoice} model...` });
  
  let raw: string;
  // Logic to switch between models
  if (modelChoice === 'gemini') {
    raw = await streamFromGemini(messages, on);
  } else {
    raw = await streamFromOllama(messages, on);
  }

  on({ type: 'status', data: 'parsing plan...' });

  let plan: Plan;
  try {
    plan = JSON.parse(cleanToJsonString(raw));
  } catch (e) {
    on({ type: 'error', data: `Model returned invalid JSON. Raw output: ${raw}` });
    throw e;
  }
  on({ type: 'plan', data: plan });

  const executed = executePlan(plan, wb);
  on({ type: 'status', data: `executed ${executed.length} step(s)` });
  return { plan: { ...plan, steps: executed }, workbook: wb.toJSON() };
}

function executePlan(plan: Plan, wb: SheetModel) {
  const executed: any[] = [];
  for (const step of plan.steps || []) {
    try {
      const sanitizedArgs = JSON.parse(JSON.stringify(step.args));
      if (sanitizedArgs.range) {
        if (sanitizedArgs.range.r3 !== undefined && sanitizedArgs.range.r2 === undefined) {
          sanitizedArgs.range.r2 = sanitizedArgs.range.r3;
          delete sanitizedArgs.range.r3;
        }
        if (sanitizedArgs.range.c3 !== undefined && sanitizedArgs.range.c2 === undefined) {
          sanitizedArgs.range.c2 = sanitizedArgs.range.c3;
          delete sanitizedArgs.range.c3;
        }
      }

      wb.dispatch(step.op, sanitizedArgs);
      executed.push(step);
    } catch (err) {
      executed.push({ ...step, explain: `FAILED: ${(err as Error).message}` });
      break;
    }
  }
  wb.evaluateAll();
  wb.checkpoint('agent');
  return executed;
}