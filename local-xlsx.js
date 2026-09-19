/* Local spreadsheet reader for Motabe — no uploads, no network */
async function readSpreadsheetLocally(file){
  const name=(file.name||'').toLowerCase();
  if(name.endsWith('.csv')) return parseLocalCSV(await file.text());
  if(!name.endsWith('.xlsx')) throw new Error('Only XLSX/CSV supported offline');
  if(typeof JSZip==='undefined') throw new Error('Local ZIP reader unavailable');
  const zip=await JSZip.loadAsync(await file.arrayBuffer());
  const xml=async p=>{const f=zip.file(p);return f?new DOMParser().parseFromString(await f.async('text'),'application/xml'):null};
  const els=(doc,name)=>doc?[...doc.getElementsByTagName('*')].filter(e=>e.localName===name):[];
  const children=(node,name)=>node?[...node.getElementsByTagName('*')].filter(e=>e.localName===name):[];
  const sharedDoc=await xml('xl/sharedStrings.xml');
  const shared=sharedDoc?els(sharedDoc,'si').map(si=>children(si,'t').map(t=>t.textContent||'').join('')):[];
  const wb=await xml('xl/workbook.xml');
  const rel=await xml('xl/_rels/workbook.xml.rels');
  if(!wb||!rel) throw new Error('Invalid workbook');
  const relMap={}; for(const r of els(rel,'Relationship')) relMap[r.getAttribute('Id')]=r.getAttribute('Target');
  const allSheets=els(wb,'sheet'); if(!allSheets.length)return [];
  let bestRows=[],bestScore=-1;
  for(const sh of allSheets){
    const rid=sh.getAttribute('r:id')||sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    let target=(relMap[rid]||'worksheets/sheet1.xml').replace(/^\//,'').replace(/^xl\//,'').replace(/^\.\.\//,'');
    const sheet=await xml('xl/'+target); if(!sheet)continue;
    const rows=[];
    for(const row of els(sheet,'row')){
      const vals={}; let max=-1;
      for(const c of children(row,'c')){
        const ref=c.getAttribute('r')||''; const letters=(ref.match(/[A-Z]+/i)||['A'])[0].toUpperCase();
        let col=0; for(const ch of letters) col=col*26+(ch.charCodeAt(0)-64); col--; max=Math.max(max,col);
        const type=c.getAttribute('t'); let value='';
        if(type==='inlineStr'){value=children(c,'t').map(t=>t.textContent||'').join('')}
        else {const v=children(c,'v')[0]?.textContent??''; value=type==='s'?(shared[Number(v)]??''):v}
        vals[col]=value;
      }
      if(max>=0)rows.push(Array.from({length:max+1},(_,i)=>vals[i]??''));
    }
    while(rows.length&&rows[0].every(x=>String(x).trim()===''))rows.shift(); if(!rows.length)continue;
    let hi=0,scoreBest=-1;
    for(let i=0;i<Math.min(rows.length,25);i++){
      const line=rows[i].map(x=>String(x??'').replace(/\s+/g,' ').trim()).join(' | ');
      const score=[/اسم\s*المدرب/i,/اسم\s*المقرر/i,/الأسبوع\s*التدريبي|الاسبوع\s*التدريبي/i,/عدد\s*المتدربين/i].filter(r=>r.test(line)).length;
      if(score>scoreBest){scoreBest=score;hi=i}
    }
    if(scoreBest>bestScore){bestScore=scoreBest;bestRows=[rows,hi]}
  }
  if(!bestRows.length)return [];
  const [rows,hi]=bestRows;
  const headers=rows[hi].map((h,i)=>String(h||`عمود ${i+1}`).replace(/\u00a0/g,' ').replace(/[\u200e\u200f]/g,'').replace(/\s+/g,' ').trim());
  return rows.slice(hi+1).filter(r=>r.some(x=>String(x).trim()!=='')).map(r=>Object.fromEntries(headers.map((h,i)=>[h,r[i]??''])));
}
function parseLocalCSV(text){
  const lines=text.replace(/^\uFEFF/,'').split(/\r?\n/).filter(x=>x.trim()); if(!lines.length)return [];
  const parse=line=>{let out=[],cur='',q=false;for(let i=0;i<line.length;i++){let ch=line[i];if(ch==='"'){if(q&&line[i+1]==='"'){cur+='"';i++}else q=!q}else if(ch===','&&!q){out.push(cur);cur=''}else cur+=ch}out.push(cur);return out};
  const h=parse(lines[0]).map(x=>x.trim());return lines.slice(1).map(line=>{let a=parse(line),o={};h.forEach((k,i)=>o[k]=a[i]??'');return o});
}
