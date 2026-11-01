/* ---- reading Word (.docx) versions of the reports ---- */
async function docxContent(JSZip,bytes){
  const zip=await JSZip.loadAsync(bytes); const xml=await zip.file('word/document.xml').async('string');
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  const kids=(el,name)=>{const o=[];for(let n=el.firstChild;n;n=n.nextSibling) if(n.nodeType===1&&n.localName===name)o.push(n);return o};
  const ptext=p=>{let s='';const walk=n=>{for(let c=n.firstChild;c;c=c.nextSibling){if(c.nodeType!==1)continue;if(c.localName==='t')s+=c.textContent;else if(c.localName==='br'||c.localName==='tab'||c.localName==='cr')s+='\n';else walk(c)}};walk(p);return s};
  const tables=[]; const paras=[];
  for(const tbl of Array.from(doc.getElementsByTagNameNS(WNS,'tbl'))){
    const rows=kids(tbl,'tr').map(tr=>kids(tr,'tc').map(tc=>Array.from(tc.getElementsByTagNameNS(WNS,'p')).map(ptext).join('\n').trim()));
    tables.push(rows);
  }
  for(const p of Array.from(doc.getElementsByTagNameNS(WNS,'p'))) { const t=ptext(p).trim(); if(t) paras.push(t) }
  return {tables,paras,zip,doc};
}
function docxKind(c){
  const flat=c.tables.flat(2).join(' ')+' '+c.paras.join(' ');
  if(/نسبة الحضور|متابعة سير العملية|اسم المعلم/.test(flat)) return 'form';
  if(/تحضير دروس|عنوان الدرس/.test(flat)) return 'teachers';
  if(/الحصص الغائبة|غياب الحصص/.test(flat)) return 'students';
  return null;
}
function parseStudentsDocx(c){
  const students=[]; const dateText=c.paras.find(p=>/تقرير/.test(p)&&/\d/.test(p))||'';
  for(const t of c.tables){
    let map=null;
    for(const row of t){
      const txt=row.map(norm);
      if(txt.includes('الحالة')&&txt.some(x=>/الطالب/.test(x))){ map={}; txt.forEach((h,i)=>{ if(h==='م')map.no=i; else if(/الطالب/.test(h))map.name=i; else if(/الفصل/.test(h))map.cls=i; else if(/الحاضرة/.test(h))map.pr=i; else if(/الغائبة/.test(h))map.ab=i; else if(/الحالة/.test(h))map.st=i }); continue }
      if(!map) continue;
      const st=txt[map.st]||''; if(!/حاضر|غائب/.test(st)) continue;
      const cls=txt[map.cls]||''; const m=cls.match(/(أول|ثاني|ثالث)\s*ش\s*(\d+)/);
      const list=s=>(s||'').split(/[،,\n]/).map(x=>x.trim()).filter(x=>x&&x!=='0'&&x!=='٠');
      students.push({no:+String(txt[map.no]||'').replace(/\D/g,''),name:txt[map.name]||'',classFull:cls,grade:m?m[1]:'',section:m?+m[2]:null,present:list(row[map.pr]),absent:list(row[map.ab]),status:/غائب/.test(st)?'غائب':'حاضر'});
    }
  }
  return {students,dateText};
}
function parseTeachersDocx(c){
  const teachers=[]; let cur=null, totals=null;
  const dateText=c.paras.find(p=>/تقرير/.test(p)&&/بتاريخ/.test(p))||'';
  const PW=['الأولى','الثانية','الثالثة','الرابعة','الخامسة','السادسة','السابعة','الثامنة','التاسعة','العاشرة'];
  for(const t of c.tables){
    let map=null;
    for(const row of t){
      const txt=row.map(norm);
      if(txt.includes('الاسم')&&txt.some(x=>PW.includes(x))){ map={per:{}}; txt.forEach((h,i)=>{ if(h==='م')map.no=i; else if(h==='الاسم')map.name=i; else if(PW.includes(h))map.per[i]=PW.indexOf(h)+1 }); continue }
      if(txt.some(x=>/الإجمالي/.test(x))){ const nums=txt.filter(x=>/^\d+$/.test(x)).map(Number); const hdrRow=t[t.indexOf(row)-1]?.map(norm)||[];
        if(nums.length>=4){ const pick=re=>{const i=hdrRow.findIndex(h=>re.test(h)); return i>=0&&/^\d+$/.test(txt[i])?+txt[i]:null};
          totals={all:pick(/كل الدروس/)??nums[0],async:pick(/غير متزامنة/)??nums[1],prepared:pick(/^عدد الدروس المعدة|المُعدة$/)??nums[2],notPrepared:pick(/الغير المعدة|غير المعدة/)??nums[3]} } continue }
      if(!map) continue;
      const name=norm(txt[map.name]||'');
      if(name){ cur={no:+String(txt[map.no]||'').replace(/\D/g,'')||teachers.length+1,name,lessons:[]}; teachers.push(cur) }
      if(!cur) continue;
      for(const [ci,period] of Object.entries(map.per)){
        const cell=row[ci]||''; if(!cell.trim()) continue;
        const re=/(أول|ثاني|ثالث)\s*ش\s*(\d+)/g; const idx=[]; let m;
        while((m=re.exec(cell))) idx.push({i:m.index,g:m[1],s:+m[2],end:re.lastIndex});
        idx.forEach((h,k)=>{ const seg=cell.slice(h.end,k+1<idx.length?idx[k+1].i:cell.length); const before=cell.slice(k?idx[k-1].end:0,h.i);
          cur.lessons.push({grade:h.g,section:h.s,period,ok:!/[X×✗]/.test(before.slice(-4)),details:[seg]}) });
      }
    }
  }
  return {teachers,totals,dateText};
}
// shared post-processing for both PDF and Word sources
function finishTeachers(T){
  let raw=0,rawAsync=0,rawOk=0;
  for(const t of T.teachers){
    const seen=new Map();
    for(const b of t.lessons){
      raw++; const d=norm(b.details.join(' '));
      b.async=/غير متزامن/.test(d); b.notPrepared=!b.ok||/لم يتم إعداد/.test(d);
      if(b.async)rawAsync++; if(!b.notPrepared)rawOk++;
      const m=d.match(/عنوان الدرس\s*:\s*(.*?)\s*(التوقيت|غير متزامن|لم يتم|رابط|$)/); b.title=m?fixTitle(m[1]):'';
      const key=b.grade+'|'+b.section+'|'+b.period;
      if(seen.has(key)){ const o=seen.get(key); if(!o.title&&b.title)o.title=b.title; b.drop=true } else seen.set(key,b);
    }
    t.lessons=t.lessons.filter(b=>!b.drop).sort((a,b)=>a.period-b.period);
    const titles=[...new Set(t.lessons.filter(b=>!b.notPrepared&&b.title).map(b=>b.title))];
    for(const b of t.lessons) if(!b.title) b.title=titles[0]||'';
    t.prepared=t.lessons.filter(b=>!b.notPrepared).length; t.notPrepared=t.lessons.filter(b=>b.notPrepared).length; t.async=t.lessons.filter(b=>b.async).length;
  }
  T.raw={all:raw,async:rawAsync,prepared:rawOk}; return T;
}

/* ---- form info (for the PDF version of the form) ---- */
async function formInfo(JSZip,bytes){
  const c=await docxContent(JSZip,bytes);
  const after=(label)=>{ for(const t of c.tables) for(const r of t) for(let i=0;i<r.length-1;i++) if(norm(r[i]).startsWith(label)) return norm(r[i+1]); return '' };
  let img=null; try{ const f=Object.keys(c.zip.files).find(n=>/^word\/media\/.*\.(png|jpe?g)$/i.test(n)); if(f){ const b64=await c.zip.file(f).async('base64'); img='data:image/'+(/png$/i.test(f)?'png':'jpeg')+';base64,'+b64 } }catch(e){}
  let headLines=[]; try{ const hs=Object.keys(c.zip.files).filter(n=>/^word\/header\d*\.xml$/.test(n)); for(const h of hs){ const x=await c.zip.file(h).async('string'); const lines=[...new Set((x.match(/<w:p[ >][\s\S]*?<\/w:p>/g)||[]).map(p=>(p.match(/<w:t[^>]*>([^<]*)/g)||[]).map(t=>t.replace(/<w:t[^>]*>/,'')).join('').trim()).filter(Boolean))]; if(lines.length>headLines.length) headLines=lines } }catch(e){}
  let preparer=[]; const xml=new XMLSerializer().serializeToString(c.doc); const tb=xml.match(/<w:txbxContent>([\s\S]*?)<\/w:txbxContent>/);
  if(tb){ preparer=(tb[1].match(/<w:p[ >][\s\S]*?<\/w:p>/g)||[]).map(p=>(p.match(/<w:t[^>]*>([^<]*)/g)||[]).map(t=>t.replace(/<w:t[^>]*>/,'')).join('').replace(/\s+/g,' ').trim()).filter(Boolean) }
  const stageRow=c.tables[0]&&c.tables[0].find(r=>r.some(x=>/ثانوي/.test(x)));
  return {school:after('اسم المدرسة'),principal:after('اسم مدير المدرسة'),sector:after('القطاع'),category:after('الفئة'),
    stage:/ثانوي/.test(JSON.stringify(c.tables[0]||[]))?'ثانوي':'',img,headLines,preparer,title:c.paras.find(p=>/تقرير متابعة/.test(p))||'تقرير متابعة سير العملية التعليمية أثناء تعليق الدراسة (عن بُعد)'};
}

/* ---- simple list document built on the form's own page header ---- */
async function makeListDocx(JSZip,formBytes,{title,headers,rows,widths,groups,dateText}){
  groups=groups&&groups.length?groups:[{title,rows}];
  const zip=await JSZip.loadAsync(formBytes);
  const xml=await zip.file('word/document.xml').async('string');
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  const kids=(el,name)=>{const o=[];for(let n=el.firstChild;n;n=n.nextSibling) if(n.nodeType===1&&n.localName===name&&n.namespaceURI===WNS)o.push(n);return o};
  const all=(el,name)=>Array.from(el.getElementsByTagNameNS(WNS,name));
  const text=el=>all(el,'t').map(t=>t.textContent).join('');
  const mk=n=>doc.createElementNS(WNS,'w:'+n);
  const body=all(doc,'body')[0];
  const les=all(doc,'tbl').find(t=>/اسم المعلم/.test(text(t)));
  if(!les) throw new Error('لم أجد جدول «اسم المعلم» في الاستمارة لاستخدامه قالبًا.');
  for(let n=body.firstChild;n;){const nx=n.nextSibling; if(!(n.nodeType===1&&n.localName==='sectPr')) body.removeChild(n); n=nx}
  const sect=kids(body,'sectPr')[0];
  if(dateText){ const p=mk('p'); const pPr=mk('pPr'); pPr.appendChild(mk('bidi')); const sp=mk('spacing'); sp.setAttributeNS(WNS,'w:after','120'); pPr.appendChild(sp); const jc=mk('jc'); jc.setAttributeNS(WNS,'w:val','center'); pPr.appendChild(jc); p.appendChild(pPr);
    const r=mk('r'); const rPr=mk('rPr'); const f=mk('rFonts'); ['ascii','hAnsi','cs'].forEach(k=>f.setAttributeNS(WNS,'w:'+k,'Sakkal Majalla')); rPr.appendChild(f); rPr.appendChild(mk('b')); rPr.appendChild(mk('bCs'));
    const col=mk('color'); col.setAttributeNS(WNS,'w:val','1F6F8B'); rPr.appendChild(col); for(const k of ['sz','szCs']){const e=mk(k); e.setAttributeNS(WNS,'w:val','32'); rPr.appendChild(e)} rPr.appendChild(mk('rtl')); r.appendChild(rPr);
    const tt=mk('t'); tt.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve'); tt.textContent=dateText; r.appendChild(tt); p.appendChild(r); body.insertBefore(p,sect||null) }
  groups.forEach((grp,gi)=>{
  if(gi>0){ const br=mk('p'); const r=mk('r'); const b=mk('br'); b.setAttributeNS(WNS,'w:type','page'); r.appendChild(b); br.appendChild(r); body.insertBefore(br,sect||null) }
  const tbl=les.cloneNode(true); body.insertBefore(tbl,sect||null);
  const rows0=kids(tbl,'tr'); const hdrIdx=rows0.findIndex(r=>/اسم المعلم/.test(text(r)));
  const tmpl=[rows0[hdrIdx+1].cloneNode(true),(rows0[hdrIdx+2]||rows0[hdrIdx+1]).cloneNode(true)];
  rows0.slice(hdrIdx+1).forEach(r=>tbl.removeChild(r));
  const n=headers.length;
  // resize grid to n columns
  const grid=kids(tbl,'tblGrid')[0]; const gcs=grid?kids(grid,'gridCol'):[]; const total=gcs.reduce((a,g)=>a+(+g.getAttributeNS(WNS,'w')||+g.getAttribute('w:w')||0),0)||13000;
  const ws=(widths||headers.map(()=>1)); const sum=ws.reduce((a,b)=>a+b,0); const tw=ws.map(w=>Math.round(total*w/sum));
  if(grid){ gcs.forEach(g=>grid.removeChild(g)); tw.forEach(w=>{const g=mk('gridCol'); g.setAttributeNS(WNS,'w:w',String(w)); grid.appendChild(g)}) }
  const shape=(tr,isTitle)=>{ let tcs=kids(tr,'tc');
    if(isTitle){ const span=kids(kids(tcs[0],'tcPr')[0]||mk('x'),'gridSpan')[0]; if(span) span.setAttributeNS(WNS,'w:val',String(n)); const w=all(tcs[0],'tcW')[0]; if(w) w.setAttributeNS(WNS,'w:w',String(total)); return }
    while(tcs.length>n){ tr.removeChild(tcs[tcs.length-1]); tcs=kids(tr,'tc') }
    while(tcs.length<n){ tr.appendChild(tcs[tcs.length-1].cloneNode(true)); tcs=kids(tr,'tc') }
    tcs.forEach((tc,i)=>{ const w=all(tc,'tcW')[0]; if(w){w.setAttributeNS(WNS,'w:w',String(tw[i]));w.setAttributeNS(WNS,'w:type','dxa')} }); };
  function setText(tc,s,sz){
    let ps=kids(tc,'p'); let p=ps[0]; ps.slice(1).forEach(x=>tc.removeChild(x));
    const firstR=all(p,'r')[0]; const pPr=kids(p,'pPr')[0]; let rPr=null;
    if(firstR&&kids(firstR,'rPr')[0]) rPr=kids(firstR,'rPr')[0].cloneNode(true); else if(pPr&&kids(pPr,'rPr')[0]) rPr=kids(pPr,'rPr')[0].cloneNode(true);
    for(let x=p.firstChild;x;){const nx=x.nextSibling; if(!(x.nodeType===1&&x.localName==='pPr')) p.removeChild(x); x=nx}
    const r=mk('r'); if(rPr){ if(sz){for(const tag of ['sz','szCs']){let e=kids(rPr,tag)[0]; if(!e){e=mk(tag);rPr.appendChild(e)} e.setAttributeNS(WNS,'w:val',String(sz))}} r.appendChild(rPr) }
    const t=mk('t'); t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve'); t.textContent=s; r.appendChild(t); p.appendChild(r);
  }
  const rws=kids(tbl,'tr');
  if(hdrIdx>0){ shape(rws[0],true); setText(kids(rws[0],'tc')[0],grp.title||title) }
  shape(rws[hdrIdx]); kids(rws[hdrIdx],'tc').forEach((tc,i)=>setText(tc,headers[i]));
  grp.rows.forEach((row,i)=>{ const tr=tmpl[i%2].cloneNode(true); shape(tr); const trh=all(tr,'trHeight')[0]; if(trh) trh.setAttributeNS(WNS,'w:val','360'); tbl.appendChild(tr); kids(tr,'tc').forEach((tc,k)=>setText(tc,String(row[k]??''),22)) });
  });
  zip.file('word/document.xml',new XMLSerializer().serializeToString(doc));
  return await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
