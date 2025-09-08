// server/src/pdf.ts
// Text extraction helper using pdfjs-dist (Node-friendly, no worker)
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

/** Ensure we pass a *plain* Uint8Array (not a Node Buffer) to pdfjs. */
function toUint8Array(input: Buffer | Uint8Array | ArrayBuffer): Uint8Array {
  // If it's already a plain Uint8Array (and not a Buffer), keep it
  // Note: Buffer is a subclass of Uint8Array, so explicitly exclude Buffer.
  // @ts-ignore Buffer global exists at runtime in Node
  if (input instanceof Uint8Array && !(globalThis.Buffer && Buffer.isBuffer(input))) {
    return input;
  }
  // If it's a Buffer, make a copy into a plain Uint8Array
  // @ts-ignore Buffer global exists at runtime in Node
  if (globalThis.Buffer && Buffer.isBuffer(input)) {
    return new Uint8Array(input); // copies into a non-Buffer Uint8Array
  }
  // If it's an ArrayBuffer, wrap it
  if (input instanceof ArrayBuffer) {
    return new Uint8Array(input);
  }
  // Fallback: try to construct a Uint8Array view
  return new Uint8Array(input as any);
}

/**
 * Extracts plain text from a PDF buffer using pdfjs-dist.
 * Works in Node without a worker.
 */
export async function extractTextFromPdf(buffer: Buffer | Uint8Array | ArrayBuffer): Promise<string> {
  const bytes = toUint8Array(buffer);

  // disableWorker keeps it simple in Node environments
  const pdf = await getDocument({ data: bytes, disableWorker: true, isEvalSupported: false }).promise;

  let full = '';
  try {
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const text = content.items
        .map((it: any) => (typeof it?.str === 'string' ? it.str : ''))
        .join(' ');
      full += text + '\n';
    }
  } finally {
    await pdf.cleanup?.();
    await pdf.destroy?.();
  }
  return full.trim();
}
