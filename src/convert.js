import { WorkerBrowserConverter, createWasmPaths } from '@matbee/libreoffice-converter/browser';
import { inspectZip } from './polyglot';

export async function docxToPdf(bytes, report = () => {}) {
  inspectZip(bytes);
  if (!globalThis.crossOriginIsolated) throw Error('DOCX 自动转换需要跨源隔离（COOP/COEP 响应头），请检查部署配置。');
  const base = `${import.meta.env.BASE_URL}office-wasm/`;
  const binaryBase = import.meta.env.PROD ? 'https://erseco.github.io/libreoffice-document-converter/wasm/' : base;
  const fontResponse = await fetch(`${import.meta.env.BASE_URL}fonts/NotoSansSC-VF.ttf`);
  if (!fontResponse.ok) throw Error('无法加载中文字体');
  const converter = new WorkerBrowserConverter({
    ...createWasmPaths(base), browserWorkerJs: `${base}browser.worker.global.js`,
    sofficeWasm: `${binaryBase}soffice.wasm`, sofficeData: `${binaryBase}soffice.data`,
    fonts: [{ filename: 'NotoSansSC-VF.ttf', data: await fontResponse.arrayBuffer() }],
    onProgress: info => report(info.percent || 0, 100),
  });
  try {
    await converter.initialize();
    const result = await converter.convert(bytes, { inputFormat: 'docx', outputFormat: 'pdf' }, 'input.docx');
    if (!result.data?.length) throw Error('LibreOffice 未生成 PDF');
    return new Uint8Array(result.data);
  } finally { await converter.destroy(); }
}
