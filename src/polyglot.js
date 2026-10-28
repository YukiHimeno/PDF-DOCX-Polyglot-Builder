import { unzipSync } from 'fflate';
import { XMLParser, XMLValidator } from 'fast-xml-parser';

const u16 = (v, p) => v.getUint16(p, true);
const u32 = (v, p) => v.getUint32(p, true);
const decoder = new TextDecoder();
const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', removeNSPrefix: true });
const list = x => x == null ? [] : Array.isArray(x) ? x : [x];
function crc32(bytes) { let c = 0xffffffff; for (const b of bytes) { c ^= b; for (let k=0;k<8;k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1)); } return (c ^ 0xffffffff) >>> 0; }
function xml(bytes, name) { if (!bytes) throw Error(`DOCX 缺少 ${name}`); const text = decoder.decode(bytes); if (XMLValidator.validate(text) !== true) throw Error(`${name} 不是有效 XML`); return parser.parse(text); }

export function inspectZip(bytes, checkContent = true) {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) if (u32(v,i) === 0x06054b50 && i+22+u16(v,i+20) === bytes.length) { end=i; break; }
  if (end < 0) throw Error('找不到有效的 ZIP EOCD；不支持尾部额外数据');
  const count=u16(v,end+10), size=u32(v,end+12), start=u32(v,end+16);
  if (u16(v,end+4) || u16(v,end+6) || u16(v,end+8)!==count || count===0xffff || start===0xffffffff || size===0xffffffff) throw Error('不支持分卷 ZIP 或 ZIP64');
  if (start+size !== end) throw Error('ZIP 中央目录偏移或长度无效');
  let p=start; const entries=[]; const names=new Set();
  for(let i=0;i<count;i++) {
    if(p+46>end || u32(v,p)!==0x02014b50) throw Error('ZIP 中央目录损坏');
    const flags=u16(v,p+8), method=u16(v,p+10), crc=u32(v,p+16), compressed=u32(v,p+20), expanded=u32(v,p+24), nl=u16(v,p+28), el=u16(v,p+30), cl=u16(v,p+32), off=u32(v,p+42), next=p+46+nl+el+cl;
    if(next>end || off+30>start || u32(v,off)!==0x04034b50 || compressed===0xffffffff || expanded===0xffffffff || off===0xffffffff) throw Error('ZIP 条目或本地文件头无效；不支持 ZIP64');
    if(flags & 1 || ![0,8].includes(method)) throw Error('不支持加密或此压缩方法的 DOCX');
    const name=decoder.decode(bytes.subarray(p+46,p+46+nl));
    if(names.has(name)) throw Error(`ZIP 存在重复条目：${name}`); names.add(name);
    const lnl=u16(v,off+26), lel=u16(v,off+28), data=off+30+lnl+lel;
    if(data+compressed>start || u16(v,off+6)!==flags || u16(v,off+8)!==method || decoder.decode(bytes.subarray(off+30,off+30+lnl))!==name) throw Error(`ZIP 本地文件头不匹配：${name}`);
    entries.push({name, crc, expanded, offsetField:p+42}); p=next;
  }
  if(p!==end) throw Error('ZIP 中央目录长度不匹配');
  if (checkContent) {
    let files; try { files=unzipSync(bytes); } catch(e) { throw Error(`ZIP 解压失败：${e.message}`); }
    for(const entry of entries) { if(entry.name.endsWith('/')) continue; const data=files[entry.name]; if(!data || data.length!==entry.expanded || crc32(data)!==entry.crc) throw Error(`ZIP 数据或 CRC 校验失败：${entry.name}`); }
    const types=xml(files['[Content_Types].xml'],'[Content_Types].xml');
    const rels=xml(files['_rels/.rels'],'_rels/.rels');
    const overrides=list(types.Types?.Override); const defaults=list(types.Types?.Default);
    const rel=list(rels.Relationships?.Relationship).find(r => /\/officeDocument$/.test(r['@_Type'] || '') && r['@_TargetMode'] !== 'External');
    if(!rel) throw Error('DOCX 缺少 officeDocument 根关系');
    const target=(rel['@_Target'] || '').replace(/^\//,'');
    if(!target || target.includes('..') || !files[target]) throw Error('DOCX 主文档关系目标不存在');
    const type=overrides.find(x=>x['@_PartName']===`/${target}`)?.['@_ContentType'] || defaults.find(x=>x['@_Extension']===target.split('.').pop())?.['@_ContentType'];
    if(!type || !/wordprocessingml\.(document|template)\.main\+xml/.test(type)) throw Error('DOCX 主文档内容类型无效或不受支持');
    const doc=xml(files[target],target); if(!doc.document?.body) throw Error('DOCX 主文档缺少 document/body');
  }
  return {end,start,entries};
}
export function buildCandidate(pdf, docx, separator = false) {
  const zip=inspectZip(docx); const shift=pdf.length+(separator?1:0);
  if(shift+docx.length>0xffffffff) throw Error('输出超过 ZIP 32 位偏移限制');
  const out=new Uint8Array(shift+docx.length); out.set(pdf); if(separator) out[pdf.length]=10; out.set(docx,shift);
  const v=new DataView(out.buffer); v.setUint32(shift+zip.end+16,zip.start+shift,true);
  for(const entry of zip.entries) { const pos=shift+entry.offsetField; v.setUint32(pos,u32(v,pos)+shift,true); }
  inspectZip(out); return out;
}
