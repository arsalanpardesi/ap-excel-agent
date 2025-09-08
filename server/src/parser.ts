import dotenv from 'dotenv';
import { z } from 'zod';
import type { TenKParsed } from './types.js';

dotenv.config();

const OLLAMA_BASE_URL = process.env.OLLAMA_BASE_URL || 'http://localhost:11434';
const OLLAMA_MODEL = process.env.OLLAMA_MODEL || 'qwen3:32b';

// ---- Zod schema for strict validation ----
const LineSchema = z.object({
  name: z.string(),
  values: z.array(z.number())
});

const StatementSchema = z.object({
  title: z.string(),
  periods: z.array(z.string()).min(1),
  lines: z.array(LineSchema),
  scale: z.string().optional(),
  currency: z.string().optional()
});

const ParsedSchema = z.object({
  income: StatementSchema,
  balance: StatementSchema,
  cashflow: StatementSchema,
  meta: z.object({
    source: z.string().optional(),
    company: z.string().optional(),
    fiscalYearEnd: z.string().optional()
  }).optional()
});

// ---- Prompt: explicit, schema-first, JSON-only, with example ----
const SYS = `
You convert a US public company's 10-K text into three structured statements.

Return ONLY strict JSON (no prose, no markdown, no comments) with keys:
{
  "income":   { "title": string, "periods": string[], "lines": [{"name": string, "values": number[]}], "scale"?: string, "currency"?: string },
  "balance":  { "title": string, "periods": string[], "lines": [{"name": string, "values": number[]}], "scale"?: string, "currency"?: string },
  "cashflow": { "title": string, "periods": string[], "lines": [{"name": string, "values": number[]}], "scale"?: string, "currency"?: string },
  "meta"?:    { "source"?: string, "company"?: string, "fiscalYearEnd"?: string }
}

Rules:
- Use the company's own period labels exactly as reported (e.g., "FY 2024", "Year Ended Dec 31, 2024", "Q4 2023").
- Preserve scale units as reported in headings (e.g., “in millions” -> put "scale":"millions", and keep numeric values in millions).
- For each "lines[].values", the length MUST equal the number of "periods". If a value is missing, put 0.
- Use plain numbers only (no strings with commas, no "—" or "N/A"; replace missing with 0).
- Do not add any keys beyond the schema.

Example (structure only):
{
  "income":   { "title":"Consolidated Statements of Operations", "periods":["2024","2023"], "lines":[{"name":"Revenue","values":[1234,1100]}], "scale":"millions","currency":"USD" },
  "balance":  { "title":"Consolidated Balance Sheets",          "periods":["12/31/24","12/31/23"], "lines":[{"name":"Cash","values":[100,90]}], "scale":"millions","currency":"USD" },
  "cashflow": { "title":"Consolidated Statements of Cash Flows", "periods":["2024","2023"], "lines":[{"name":"Net cash from ops","values":[300,280]}], "scale":"millions","currency":"USD" },
  "meta":     { "source":"10-K", "company":"Example, Inc.", "fiscalYearEnd":"Dec 31" }
}
`.trim();

function sanitizeJsonish(s: string): string {
  // Strip code fences & BOM
  s = s.replace(/^\uFEFF/, '').replace(/^```(?:json)?\s*/i, '').replace(/```$/i, '').trim();
  // Extract the largest {...} block if extra prose leaked
  const i = s.indexOf('{'); const j = s.lastIndexOf('}');
  if (i !== -1 && j !== -1 && j > i) s = s.slice(i, j + 1);

  // Replace “smart” quotes with straight quotes
  s = s.replace(/[“”]/g, '"').replace(/[‘’]/g, "'");
  // Remove trailing commas in JSON objects/arrays
  s = s.replace(/,(\s*[}\]])/g, '$1');
  // Replace NaN/Infinity with null
  s = s.replace(/\bNaN\b/g, 'null').replace(/\bInfinity\b/g, 'null').replace(/\b-Infinity\b/g, 'null');
  return s;
}

function tryParseStrict(s: string) {
  const clean = sanitizeJsonish(s);
  return JSON.parse(clean);
}

function messagesToPrompt(messages: { role: string; content: string }[]) {
  // Simple stitch for /api/generate
  return messages.map(m => `${m.role.toUpperCase()}: ${m.content}`).join('\n\n');
}

async function ollamaChat(messages: any[], forceJson = true) {
  // /api/chat with format:'json' (if supported by the daemon)
  const body: any = { model: OLLAMA_MODEL, messages, stream: false, options: { temperature: 0, num_ctx: 8192 } };
  if (forceJson) body.format = 'json';

  const res = await fetch(`${OLLAMA_BASE_URL}/api/chat`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body)
  });
  return res;
}

async function ollamaGenerate(messages: any[], forceJson = true) {
  // /api/generate as fallback or compatibility mode
  const prompt = messagesToPrompt(messages);
  const body: any = { model: OLLAMA_MODEL, prompt, stream: false, options: { temperature: 0, num_ctx: 8192 } };
  if (forceJson) body.format = 'json';

  const res = await fetch(`${OLLAMA_BASE_URL}/api/generate`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body)
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error(`Ollama /api/generate ${res.status}${txt ? ` – ${txt}` : ''}`);
  }
  const data = await res.json();
  const content = (data?.response ?? '').toString();
  return content;
}

async function callOllama(messages: any[]) {
  // 1) Try /api/chat with format: 'json'
  let res: Response;
  try {
    res = await ollamaChat(messages, true);
  } catch (e) {
    throw new Error(`Could not reach Ollama at ${OLLAMA_BASE_URL} – ${String(e)}`);
  }

  if (res.status === 404) {
    // 2) Fallback to /api/generate with format:'json'
    const content = await ollamaGenerate(messages, true);
    try { return tryParseStrict(content); } catch (_) {
      // 3) Repair pass (still on /api/generate, force JSON)
      const repaired = await repairWithModel(content);
      return tryParseStrict(repaired);
    }
  }

  if (!res.ok) {
    const body = await res.text().catch(() => '');
    if (res.status === 404 && /model not found/i.test(body)) {
      throw new Error(`Ollama says the model '${OLLAMA_MODEL}' is not installed. Try:\n  ollama pull ${OLLAMA_MODEL}`);
    }
    throw new Error(`Ollama /api/chat ${res.status}${body ? ` – ${body}` : ''}`);
  }

  const data = await res.json();
  const content = (data?.message?.content ?? '').toString();
  try { return tryParseStrict(content); } catch (_) {
    // 3) Repair pass if chat returned non-JSON
    const repaired = await repairWithModel(content);
    return tryParseStrict(repaired);
  }
}

async function repairWithModel(bad: string) {
  const repairUser = `
You emitted invalid JSON. Convert the following content into STRICT JSON that matches the schema from the system prompt. 
- Do not include any prose, code fences, or comments.
- Do not add extra keys.
- Ensure the "lines[].values" arrays exactly match the length of "periods"; fill missing values with 0.
Here is your previous output:
\`\`\`
${bad}
\`\`\`
`.trim();

  // Ask model to fix it, force JSON again
  const prompt = messagesToPrompt([
    { role: 'system', content: SYS },
    { role: 'user', content: repairUser }
  ]);

  const body: any = {
    model: OLLAMA_MODEL,
    prompt,
    stream: false,
    format: 'json',
    options: { temperature: 0, num_ctx: 8192 }
  };

  const res = await fetch(`${OLLAMA_BASE_URL}/api/generate`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body)
  });
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error(`Ollama repair failed ${res.status}${txt ? ` – ${txt}` : ''}`);
  }
  const data = await res.json();
  return (data?.response ?? '').toString();
}

// ---- Public API ----
export async function parseTenKToStructured(docText: string, sourceLabel: string): Promise<TenKParsed> {
  // Keep prompt bounded for local models
  const MAX_CHARS = 220_000;
  const text = docText.slice(0, MAX_CHARS);

  const messages = [
    { role: 'system', content: SYS },
    { role: 'user', content: JSON.stringify({ source: sourceLabel, text }) }
  ];

  const raw = await callOllama(messages);

  // Validate & normalize with Zod; also pad/truncate values to periods length
  const parsed = ParsedSchema.parse(raw);

  const normalize = (st: z.infer<typeof StatementSchema>) => {
    const n = st.periods.length;
    return {
      ...st,
      lines: st.lines.map(l => ({
        name: l.name,
        values: (l.values ?? []).slice(0, n).concat(Array(Math.max(0, n - (l.values ?? []).length)).fill(0))
      }))
    };
  };

  const clean: TenKParsed = {
    income: normalize(parsed.income),
    balance: normalize(parsed.balance),
    cashflow: normalize(parsed.cashflow),
    meta: parsed.meta
  };

  return clean;
}
