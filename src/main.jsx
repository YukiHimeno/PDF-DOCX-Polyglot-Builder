import React, { useState, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import { ArrowDown, ArrowRight, Check, CheckCircle2, ChevronRight, FileText, FileType, LockKeyhole, Plus, ShieldCheck, Sparkles, X, AlertCircle, LoaderCircle, Fingerprint } from 'lucide-react';
import * as pdfjs from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import './style.css';
pdfjs.GlobalWorkerOptions.workerSrc=pdfWorker;
const formatSize=n=>n<1024?`${n} B`:n<1048576?`${(n/1024).toFixed(1)} KB`:`${(n/1048576).toFixed(2)} MB`;
function DropZone({kind,file,onChange,disabled}) {
 const ref=useRef(); const [drag,setDrag]=useState(false); const Icon=kind==='PDF'?FileText:FileType;
 const choose=f=>{if(f) onChange(f)};
 return <div className={`drop ${drag?'drag':''} ${file?'has-file':''}`} onDragOver={e=>{e.preventDefault();if(!disabled)setDrag(true)}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);if(!disabled)choose(e.dataTransfer.files[0])}}>
  <input ref={ref} type="file" accept={kind==='PDF'?'.pdf,application/pdf':'.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document'} onChange={e=>{choose(e.target.files[0]);e.target.value=''}} disabled={disabled} hidden/>
  {file?<><div className="file-icon"><Icon size={25}/></div><div className="file-info"><strong title={file.name}>{file.name}</strong><span>{kind} 文档 <i/> {formatSize(file.size)}</span></div><button className="icon-button" aria-label={`移除 ${kind} 文件`} disabled={disabled} onClick={()=>onChange(null)}><X size={18}/></button></>:<button className="drop-button" onClick={()=>ref.current.click()} disabled={disabled}><span className="upload-icon"><Icon size={24}/><span className="plus"><Plus size={12}/></span></span><strong>拖放 {kind} 文件到这里</strong><span>或 <u>点击选择文件</u></span><small>{kind==='PDF'?'支持 .pdf 格式':'支持 .docx 格式'}</small></button>}
 </div>;
}
async function validatePdf(bytes, onPage) {
 const task=pdfjs.getDocument({data:bytes.slice(), stopAtErrors:true}); const doc=await task.promise;
 if(!doc.numPages) throw Error('PDF 没有页面');
 try { for(let i=1;i<=doc.numPages;i++){const page=await doc.getPage(i); const viewport=page.getViewport({scale:Math.min(1,1000/Math.max(page.view[2]-page.view[0],page.view[3]-page.view[1]))}); const canvas=document.createElement('canvas'); canvas.width=Math.max(1,Math.ceil(viewport.width));canvas.height=Math.max(1,Math.ceil(viewport.height)); await page.render({canvasContext:canvas.getContext('2d'),viewport}).promise; page.cleanup(); onPage(i,doc.numPages); } return doc.numPages; } finally {await doc.destroy();}
}
function App(){
 const [pdf,setPdf]=useState(null),[docx,setDocx]=useState(null),[busy,setBusy]=useState(false),[progress,setProgress]=useState(0),[logs,setLogs]=useState([]),[error,setError]=useState(''),[result,setResult]=useState(null);
 const log=s=>setLogs(l=>[...l,s]);
 const update=(setter,f,ext)=>{setResult(null);setError('');setLogs([]);if(f&&!f.name.toLowerCase().endsWith(ext)){setter(null);setError(`请选择 ${ext} 格式的文件。`);return}setter(f)};
 async function generate(){
  setBusy(true);setProgress(5);setError('');setResult(null);setLogs(['开始本地处理，文件不会上传到服务器。']);
  try {
   if(!crypto.subtle) throw Error('SHA-256 需要安全上下文，请使用 HTTPS 或 localhost。');
   let p = pdf ? await pdf.arrayBuffer() : null, d = docx ? await docx.arrayBuffer() : null;
   if (!p || !d) {
    if (!d) throw Error('自动将 PDF 转为保留排版的可编辑 DOCX 暂不可用。请先使用 Microsoft Word 打开 PDF 并另存为 DOCX，然后上传两个文件。');
    let docxToPdf;
    try { ({ docxToPdf } = await import('./convert')); }
    catch (cause) { throw Error(`转换模块加载失败，请重启开发服务器并刷新页面（${cause.message}）`); }
    log('正在启动 LibreOffice WASM（首次加载约 240 MB），转换 DOCX 为 PDF…');
    p = (await docxToPdf(new Uint8Array(d), (i,n)=>{setProgress(5+Math.round(8*i/n));})).buffer;
    log('已通过 LibreOffice 生成 PDF；请核对字体和排版。');
   }
   log('正在验证原始 PDF 的每一页…'); await validatePdf(new Uint8Array(p),()=>{}); setProgress(15);
   const candidates=await new Promise((resolve,reject)=>{const w=new Worker(new URL('./worker.js',import.meta.url),{type:'module'});w.onmessage=({data})=>{if(data.type==='log'){log(data.message);setProgress(data.progress)}else if(data.type==='error'){w.terminate();reject(Error(data.message))}else if(data.type==='candidates'){w.terminate();resolve(data.candidates)}};w.onerror=()=>{w.terminate();reject(Error('处理线程异常，可能内存不足。'))};w.postMessage({pdf:p,docx:d},[p,d]);});
   let chosen, pages; const failures=[];
   for(const candidate of candidates){try{log(`验证候选布局：${candidate.layout}`); const bytes=new Uint8Array(candidate.bytes); pages=await validatePdf(bytes,(i,n)=>setProgress(40+Math.round(45*i/n)));chosen={bytes,layout:candidate.layout};log(`PDF.js 已渲染全部 ${pages} 页。`);break}catch(e){failures.push(`${candidate.layout}：${e.message}`);log(`候选布局失败：${e.message}`)}}
   if(!chosen) throw Error(`所有候选布局均未通过 PDF 渲染验证：${failures.join('；')}`);
   setProgress(90); const hash=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',chosen.bytes)),x=>x.toString(16).padStart(2,'0')).join('');
   const blob=new Blob([chosen.bytes],{type:'application/octet-stream'});setResult({blob,hash,pages,layout:chosen.layout});setProgress(100);log('构建完成。两个下载使用同一个二进制文件。');
  }catch(e){setError(e.message||'处理失败，请检查文件。');log(`处理失败：${e.message}`)}finally{setBusy(false)}
 }
 function download(ext){const url=URL.createObjectURL(result.blob);const a=document.createElement('a');a.href=url;a.download=`polyglot.${ext}`;a.click();setTimeout(()=>URL.revokeObjectURL(url),60000)}
 return <><header className="topbar"><a className="brand" href="#top"><span className="brand-mark"><span>p</span><span>d</span></span><span>polyglot<span className="brand-dot">.</span></span></a><div className="top-right"><span className="desktop-label">PDF + DOCX BUILDER</span><span className="top-divider"/><span className="privacy"><LockKeyhole size={14}/> 100% 本地处理</span></div></header>
 <main id="top"><div className="intro"><div className="eyebrow"><span className="eyebrow-dot"/> DOCUMENT TOOLKIT <span className="eyebrow-line"/> 01 / 01</div><h1>两个文档，<br/><span>一种全新可能。</span></h1><p>将 PDF 和 DOCX 融合为同一个文件。<br className="desktop"/>只需更改扩展名，即可以不同格式打开。</p><div className="intro-note"><LockKeyhole size={15}/> 文件始终留在你的设备上，安全又私密。</div></div>
  <section className="workspace" aria-label="创建融合文件"><div className="workspace-head"><div><div className="section-kicker">THE BUILDER <span>/</span> 文件融合</div><h2>准备好你的文件</h2><p>上传 DOCX 可自动生成 PDF，或分别上传 PDF 和 DOCX。</p></div><div className="step-badge"><span>01</span> / 03</div></div>
 <div className="upload-grid"><div><div className="field-label"><span><span className="number">01</span> PDF 文档</span><span>原始文件</span></div><DropZone kind="PDF" file={pdf} onChange={f=>update(setPdf,f,'.pdf')} disabled={busy}/></div><div className="connector"><Plus size={17}/></div><div><div className="field-label"><span><span className="number">02</span> Word 文档</span><span>原始文件</span></div><DropZone kind="DOCX" file={docx} onChange={f=>update(setDocx,f,'.docx')} disabled={busy}/></div></div>
  <div className="conversion-note" aria-live="polite">{pdf && docx ? '已选择两个文件，将分别保留各自内容。' : pdf ? 'PDF 自动转换为保留排版的可编辑 DOCX 暂不可用。请在 Word 中转换后上传 DOCX。' : docx ? '将使用 LibreOffice 在本地转换为 PDF（首次加载约 240 MB）；字体和排版可能与 Word 不同，请检查。' : '选择 Word 文档即可自动生成 PDF，也可同时选择两种文档。'}</div>
  <div className="action-row"><div className="action-hint"><ShieldCheck size={18}/><span>你的文件，仅由你的浏览器处理。</span></div><button className="primary" disabled={(!pdf&&!docx)||busy} onClick={generate}>{busy?<><LoaderCircle className="spin" size={17}/> 正在处理</>:<>生成融合文件 <ArrowRight size={18}/></>}</button></div>
 {(busy||logs.length>0)&&<div className="processing"><div className="process-title"><strong>{busy?'正在处理文件':error?'处理未完成':'处理完成'}</strong><span>{progress}%</span></div><div className="progress"><div style={{width:`${progress}%`}}/></div><div className="logs" aria-live="polite">{logs.map((x,i)=><div key={i}><span>›</span> {x}</div>)}</div></div>}
 {error&&<div className="error" role="alert"><AlertCircle size={18}/><span>{error}</span></div>}
 {result&&<div className="result"><div className="result-heading"><span className="success-icon"><Check size={19}/></span><div><h3>融合文件已生成</h3><p>{formatSize(result.blob.size)} · 两种扩展名，同一份二进制内容</p></div></div><div className="validation"><div><CheckCircle2 size={16}/> PDF.js 已渲染 {result.pages} 页</div><div><CheckCircle2 size={16}/> ZIP 数据与 OOXML 基础结构已检查</div><div className="pending"><AlertCircle size={16}/> Microsoft Word 原生打开：需要本地验证</div></div><div className="hash"><Fingerprint size={17}/><div><strong>SHA-256</strong><code>{result.hash}</code></div></div><div className="downloads"><button onClick={()=>download('pdf')}><ArrowDown size={17}/> 下载 .pdf</button><button onClick={()=>download('docx')}><ArrowDown size={17}/> 下载 .docx</button></div><p className="result-note">布局：{result.layout}。结构检查不能保证所有阅读器或 Word 版本的兼容性，请在目标软件中检查内容和排版。</p></div>}
 </section>
 <div className="below"><div className="below-title"><span>DESIGNED FOR SIMPLICITY</span><h2>一份文件，<br/>两种打开方式。</h2><p>保留各自的文档内容，让不同格式在同一个文件中共存。</p></div><div className="features"><div><span className="feature-icon"><FileText size={19}/></span><div><strong>保留原始内容</strong><p>PDF 与 DOCX 各自独立，保持原有内容。</p></div><ChevronRight size={16}/></div><div><span className="feature-icon"><ShieldCheck size={19}/></span><div><strong>隐私优先</strong><p>所有处理都在浏览器中完成，无需上传。</p></div><ChevronRight size={16}/></div><div><span className="feature-icon"><Sparkles size={19}/></span><div><strong>透明的验证</strong><p>检查文件结构，并清楚标注需要本地验证的项目。</p></div><ChevronRight size={16}/></div></div></div></main><footer><span>© {new Date().getFullYear()} POLYGLOT BUILDER</span><span>MADE FOR DOCUMENTS, MADE FOR YOU <span className="footer-dot">✳</span></span></footer></>;
}

createRoot(document.getElementById('root')).render(<App/>);
