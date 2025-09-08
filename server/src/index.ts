import express from 'express';
import cors from 'cors';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';
import multer from 'multer';
import * as XLSX from 'xlsx';

import { SheetModel } from './sheet.js';
import { streamPlanAndExecute, AgentStreamEvent, AgentHints } from './agent.js';
import { extractTextFromPdf } from './pdf.js';
import { parseTenKToStructured } from './parser.js';

dotenv.config();

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT) || 3000;

const app = express();
app.use(cors());
app.use(express.json({ limit: '16mb' }));

const upload = multer({ storage: multer.memoryStorage() });
const model = new SheetModel();

/* ================= Workbook APIs ================= */

app.get('/api/workbook', (_req, res) => res.json(model.toJSON()));

app.post('/api/sheetOps', (req, res) => {
  const { op, args } = req.body || {};
  try {
    model.dispatch(op, args);
    model.evaluateAll();
    res.json({ ok: true, workbook: model.toJSON() });
  } catch (e) {
    res.status(400).json({ ok: false, error: (e as Error).message });
  }
});

app.post('/api/undo', (_req, res) => {
  const ev = model.undo();
  model.evaluateAll();
  res.json({ undone: ev?.summary ?? null, workbook: model.toJSON() });
});

app.get('/api/provenance', (req, res) => {
  const sheet = String(req.query.sheet || 'Sheet1');
  const cell = String(req.query.cell || 'A1');
  try {
    const prov = model.getProvenance(sheet, cell);
    res.json({ sheet, cell, provenance: prov });
  } catch (e) {
    res.status(400).json({ error: (e as Error).message });
  }
});

/* ========== Import/Ingest Endpoints ========== */
app.post('/api/ingest-10k', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ ok: false, error: 'No PDF uploaded' });

    const sourceLabel = req.file.originalname || '10-K.pdf';
    const text = await extractTextFromPdf(req.file.buffer);
    if (!text.trim()) return res.status(400).json({ ok: false, error: 'Could not extract text from PDF' });

    const parsed = await parseTenKToStructured(text, sourceLabel);

    const headerIncome = ['Line item', ...parsed.income.periods];
    try { (model as any).deleteSheet('P&L'); } catch {}
    (model as any).createSheet('P&L');
    (model as any).setValues({ sheet: 'P&L', r1: 0, c1: 0, r2: 0, c2: headerIncome.length - 1 }, [headerIncome]);
    parsed.income.lines.forEach((l, i) =>
      (model as any).setValues({ sheet: 'P&L', r1: i+1, c1: 0, r2: i+1, c2: headerIncome.length - 1 }, [[ l.name, ...l.values ]])
    );

    const headerBS = ['Line item', ...parsed.balance.periods];
    try { (model as any).deleteSheet('Balance Sheet'); } catch {}
    (model as any).createSheet('Balance Sheet');
    (model as any).setValues({ sheet: 'Balance Sheet', r1: 0, c1: 0, r2: 0, c2: headerBS.length - 1 }, [headerBS]);
    parsed.balance.lines.forEach((l, i) =>
      (model as any).setValues({ sheet: 'Balance Sheet', r1: i+1, c1: 0, r2: i+1, c2: headerBS.length - 1 }, [[ l.name, ...l.values ]])
    );

    const headerCF = ['Line item', ...parsed.cashflow.periods];
    try { (model as any).deleteSheet('Cash Flow'); } catch {}
    (model as any).createSheet('Cash Flow');
    (model as any).setValues({ sheet: 'Cash Flow', r1: 0, c1: 0, r2: 0, c2: headerCF.length - 1 }, [headerCF]);
    parsed.cashflow.lines.forEach((l, i) =>
      (model as any).setValues({ sheet: 'Cash Flow', r1: i+1, c1: 0, r2: i+1, c2: headerCF.length - 1 }, [[ l.name, ...l.values ]])
    );

    res.json({ ok: true, parsed, workbook: model.toJSON() });
  } catch (e) {
    res.status(500).json({ ok: false, error: (e as Error).message });
  }
});

app.post('/api/import-xlsx', (req, res) => {
  try {
    const workbookData = req.body;
    if (!workbookData || !Array.isArray(workbookData.sheets)) {
      throw new Error('Invalid workbook data format.');
    }
    model.loadFromJSON(workbookData);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ ok: false, error: (e as Error).message });
  }
});


/* ================= Export workbook as .xlsx ================= */
app.get('/api/export-xlsx', (_req, res) => {
  try {
    const book = XLSX.utils.book_new();
    const wb = (model as any).wb as { sheets: Map<string, any> };
    for (const [name, s] of wb.sheets.entries()) {
      const aoa: any[][] = s.rows.map((row: any[]) =>
        row.map((c: any) => (c?.formula ? c.formula : (c?.value ?? null)))
      );
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      XLSX.utils.book_append_sheet(book, ws, name.slice(0, 31));
    }
    const buffer = XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }) as Buffer;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="export.xlsx"');
    res.send(buffer);
  } catch (e) {
    res.status(500).json({ ok: false, error: (e as Error).message });
  }
});

/* ================= Agent endpoints ================= */

app.post('/api/agent', async (req, res) => { /* ... existing code ... */ });

app.get('/api/agent/stream', async (req, res) => {
  const modelChoice = String(req.query.model || 'ollama') as 'ollama' | 'gemini';
  const goal = String(req.query.goal || '');
  if (!goal) {
    res.writeHead(400).end('Missing goal');
    return;
  }

  const sheetHint = (req.query.sheetHint ? String(req.query.sheetHint) : undefined) || undefined;
  const insertRow = req.query.insertRow ? Math.max(0, Number(req.query.insertRow) - 1) : undefined;
  const hints: AgentHints = { sheetHint, insertRow };

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache, no-transform');
  res.setHeader('Connection', 'keep-alive');

  const send = (e: AgentStreamEvent) => {
    res.write(`event: ${e.type}\n`);
    res.write(`data: ${JSON.stringify(e.data)}\n\n`);
  };

  try {
    await streamPlanAndExecute(goal, model, hints, send, modelChoice);
    send({ type: 'done', data: 'ok' });
    res.end();
  } catch (e) {
    send({ type: 'error', data: (e as Error).message });
    res.end();
  }
});


/* ================= Static client ================= */
app.use('/', express.static(path.join(__dirname, '..', 'public')));

app.listen(PORT, () => {
  console.log(`Challenge POC (Ollama & Gemini) listening on http://localhost:${PORT}`);
  if (!process.env.GEMINI_API_KEY) {
    console.warn('⚠️  GEMINI_API_KEY not found in .env file. The Gemini model will not be available.');
  }
});