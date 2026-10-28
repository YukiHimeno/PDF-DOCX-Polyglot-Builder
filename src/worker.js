import { buildCandidate, inspectZip } from './polyglot';
self.onmessage = async ({data}) => {
 try {
  const pdf=new Uint8Array(data.pdf), docx=new Uint8Array(data.docx);
  self.postMessage({type:'log',message:'正在检查 DOCX 容器、CRC 与 OOXML 关系…',progress:18});
  const zip=inspectZip(docx);
  self.postMessage({type:'log',message:`DOCX 检查通过：${zip.entries.length} 个 ZIP 条目`,progress:32});
  const candidates=[]; const errors=[];
  for(const separator of [false,true]) { try { const bytes=buildCandidate(pdf,docx,separator); candidates.push({bytes:bytes.buffer,layout:separator?'PDF + 换行 + 修正偏移的 DOCX':'PDF + 修正偏移的 DOCX'}); } catch(e) { errors.push(e.message); } }
  if(!candidates.length) throw Error(errors.join('；'));
  self.postMessage({type:'candidates',candidates},candidates.map(c=>c.bytes));
 } catch(e) { self.postMessage({type:'error',message:e.message}); }
};
