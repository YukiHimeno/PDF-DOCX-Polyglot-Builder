import { test, expect, vi } from 'vitest';
import { docxToPdf } from '../src/convert.js';
import { zipSync, strToU8 } from 'fflate';

test('DOCX conversion requires a valid package before starting the WASM engine', async () => {
  await expect(docxToPdf(zipSync({'bad':strToU8('x')}))).rejects.toThrow('DOCX');
});

test('DOCX conversion reports missing cross-origin isolation', async () => {
  const zip=zipSync({'[Content_Types].xml':strToU8('<Types><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),'_rels/.rels':strToU8('<Relationships><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),'word/document.xml':strToU8('<document><body><p>hello</p></body></document>')});
  vi.stubGlobal('crossOriginIsolated',false);
  try { await expect(docxToPdf(zip)).rejects.toThrow('跨源隔离'); } finally {vi.unstubAllGlobals();}
});
