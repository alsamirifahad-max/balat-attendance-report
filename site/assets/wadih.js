'use strict';
/* Reader for the «واضح» (Wadeh) report «تقرير غياب الطلاب عن الحصص».
   The PDF holds one scanned-style JPEG per page and no text, so the table grid is found from the
   pixels and the cells are read with an on-device OCR engine (vendor/tesseract). Nothing leaves the browser.
   Section 1 «غائب عن حصص وحاضر في غيرها» has a «حضر في» column; section 2 «غائب في كل الحصص المرصودة»
   lists students absent all day. A student whose «حضر في» cell is empty is counted absent all day. */
(function(){
const BASE='vendor/tesseract/';
let tessP=null, workers=null;
function loadScript(src){ return new Promise((res,rej)=>{ const s=document.createElement('script'); s.src=src; s.onload=res; s.onerror=()=>rej(new Error('تعذّر تحميل قارئ الصور.')); document.head.appendChild(s) }) }
function tess(){ if(!tessP) tessP=window.Tesseract?Promise.resolve(window.Tesseract):loadScript(BASE+'tesseract.min.js?v=6.0.1').then(()=>window.Tesseract); return tessP }
async function getWorkers(onP){
  if(workers) return workers;
  const T=await tess();
  const n=Math.max(1,Math.min(3,(navigator.hardwareConcurrency||2)-1));
  const mk=()=>T.createWorker('ara',1,{workerPath:BASE+'worker.min.js',corePath:BASE,langPath:BASE+'lang',workerBlobURL:false,gzip:true,cacheMethod:'none',
    logger:m=>{ if(onP&&m&&m.status==='loading language traineddata') onP('تحميل قارئ النصوص العربية …') }});
  const list=[]; for(let i=0;i<n;i++) list.push(await mk());
  for(const w of list) await w.setParameters({tessedit_pageseg_mode:'6',preserve_interword_spaces:'1'});
  workers=list; return workers;
}

/* ---------- PDF → page JPEGs ---------- */
async function pageImages(bytes){
  const P=window.PDFLib, N=n=>P.PDFName.of(n);
  const doc=await P.PDFDocument.load(bytes,{ignoreEncryption:true,updateMetadata:false});
  const meta=((doc.getTitle()||'')+' '+(doc.getCreator()||'')+' '+(doc.getProducer()||''));
  const imgs=[];
  for(const p of doc.getPages()){
    let best=null; const res=p.node.Resources(); const xo=res&&res.lookupMaybe(N('XObject'),P.PDFDict);
    if(xo) for(const [,ref] of xo.entries()){
      const st=doc.context.lookup(ref); if(!st||!st.dict||!(st instanceof P.PDFRawStream)) continue;
      const d=st.dict; if(d.get(N('Subtype'))!==N('Image')) continue;
      let f=d.get(N('Filter')); if(f instanceof P.PDFArray){ const a=f.asArray(); f=a.length===1?a[0]:null }
      if(f!==N('DCTDecode')) continue;
      const w=d.lookup(N('Width')); const wn=w&&w.asNumber?w.asNumber():0;
      if(!best||wn>best.w) best={w:wn,bytes:st.contents};
    }
    imgs.push(best?best.bytes:null);
  }
  return {meta,imgs};
}
async function isImagePdf(bytes){ try{ const {imgs}=await pageImages(bytes); return imgs.length>0&&imgs.every(Boolean) }catch(e){ return false } }
async function looksLike(bytes){
  try{ const {meta,imgs}=await pageImages(bytes); const ok=imgs.length>0&&imgs.every(Boolean);
    return ok&&(/wadeh|واضح/i.test(meta)||/غياب الطلاب عن الحصص/.test(meta)) }catch(e){ return false }
}

/* ---------- pixel layout ---------- */
async function decode(jpeg){
  const bmp=await createImageBitmap(new Blob([jpeg],{type:'image/jpeg'}));
  const c=document.createElement('canvas'); c.width=bmp.width; c.height=bmp.height;
  const x=c.getContext('2d',{willReadFrequently:true}); x.drawImage(bmp,0,0); if(bmp.close) bmp.close();
  return {canvas:c,ctx:x,W:c.width,H:c.height,data:x.getImageData(0,0,c.width,c.height).data};
}
function runs(arr,min){ const out=[]; let s=-1; for(let i=0;i<=arr.length;i++){ const on=i<arr.length&&arr[i]; if(on&&s<0) s=i; if(!on&&s>=0){ if(i-s>=(min||1)) out.push([s,i-1]); s=-1 } } return out }
function layout(img){
  const {W,H,data}=img; const gray=new Uint8Array(W*H), dark=new Uint8Array(W*H), ink=new Uint8Array(W*H), white=new Uint8Array(W*H);
  for(let i=0,p=0;i<W*H;i++,p+=4){ const r=data[p],g=data[p+1],b=data[p+2], s=r+g+b, mx=Math.max(r,g,b), mn=Math.min(r,g,b);
    if(s>450&&s<725&&mx-mn<34) gray[i]=1; if(s<330) dark[i]=1; if(s<420) ink[i]=1; if(s>690) white[i]=1 }
  const x0=Math.round(W*.06), x1=Math.round(W*.94), span=x1-x0;
  const hl=new Uint8Array(H), hd=new Uint8Array(H);
  for(let y=0;y<H;y++){ let g=0,d=0; const o=y*W; for(let x=x0;x<x1;x++){ g+=gray[o+x]; d+=dark[o+x] } hl[y]=g>span*.6?1:0; hd[y]=d>span*.55?1:0 }
  const heads=runs(hd,18), lines=runs(hl,1).map(r=>r[0]);
  // group row separators into tables (a long gap ends a table)
  const groups=[]; let cur=null;
  const headBetween=(a,b)=>heads.some(h=>h[0]>a&&h[1]<b);
  for(const y of lines){ if(!cur||y-cur.last>260||headBetween(cur.last,y)){ cur={ys:[y],last:y}; groups.push(cur) } else if(y-cur.last>6){ cur.ys.push(y); cur.last=y } }
  const tables=[];
  const vline=(x,y)=>gray[y*W+x]||gray[y*W+x+1]||gray[y*W+x-1];
  for(const g of groups){
    const head=heads.filter(h=>h[1]<=g.ys[0]+2&&g.ys[0]-h[1]<=200).pop()||null;
    if(!head&&g.ys.length<2) continue;
    const ys=g.ys.slice(); if(head&&ys[0]-head[1]>6) ys.unshift(head[1]+1);
    const y0=ys[0], y1=Math.max(ys[ys.length-1],y0+40), hgt=y1-y0, col=new Uint8Array(W);
    for(let x=0;x<W;x++){ let c=0; for(let y=y0;y<y1;y++) c+=gray[y*W+x]; col[x]=c>hgt*.5?1:0 }
    const xs=runs(col,1).map(r=>Math.round((r[0]+r[1])/2));
    const t={head,ys,xs};
    if(xs.length>=2){ const xa=xs[0], xb=xs[xs.length-1];
      // a row cut by the page end continues below the last separator
      let y=ys[ys.length-1]+3; while(y<H&&vline(xa,y)&&vline(xb,y)) y++;
      if(y-ys[ys.length-1]>14&&y-ys[ys.length-1]<260){ ys.push(y); t.cut=true }
      // a row continued from the previous page sits above the first separator
      if(!head){ let u=ys[0]-3; while(u>0&&vline(xa,u)&&vline(xb,u)&&!hd[u]) u--; if(ys[0]-u>14&&ys[0]-u<260){ ys.unshift(u); t.pre=true } } }
    tables.push(t);
  }
  return {tables,gray,dark,ink,white};
}
function inkWidth(L,W,x0,x1,y0,y1,mask){ let a=1e9,b=-1; for(let x=x0;x<x1;x++){ for(let y=y0;y<y1;y++){ if(mask[y*W+x]){ if(x<a)a=x; if(x>b)b=x; break } } } return b<0?0:b-a+1 }
function inkCount(mask,W,x0,x1,y0,y1){ let n=0; for(let y=y0;y<y1;y++) for(let x=x0;x<x1;x++) n+=mask[y*W+x]; return n }

/* ---------- OCR helpers ---------- */
function strip(img,cells,pad,sc){
  // stack cell crops vertically (white gaps keep each cell on its own line); returns canvas + y slots
  pad=pad||18; sc=sc||1; const wmax=Math.max(...cells.map(c=>c.w))*sc; let h=pad; const slots=[];
  cells.forEach(c=>{ slots.push([h,h+c.h*sc]); h+=c.h*sc+pad });
  const cv=document.createElement('canvas'); cv.width=wmax+2*pad; cv.height=h; const x=cv.getContext('2d');
  x.fillStyle='#fff'; x.fillRect(0,0,cv.width,cv.height); x.imageSmoothingQuality='high';
  cells.forEach((c,i)=>{ x.drawImage(img.canvas,c.x,c.y,c.w,c.h,cv.width-pad-c.w*sc,slots[i][0],c.w*sc,c.h*sc) });
  // clean black-on-white text (removes the zebra row tint)
  const id=x.getImageData(0,0,cv.width,cv.height), d=id.data;
  for(let p=0;p<d.length;p+=4){ const l=(d[p]*299+d[p+1]*587+d[p+2]*114)/1000, v=l<175?0:255; d[p]=d[p+1]=d[p+2]=v }
  x.putImageData(id,0,0);
  return {cv,slots};
}
async function ocrStrip(w,st){
  const r=await w.recognize(st.cv,{},{text:true,blocks:true});
  const out=st.slots.map(()=>[]);
  const lines=[]; (r.data.blocks||[]).forEach(b=>(b.paragraphs||[]).forEach(p=>(p.lines||[]).forEach(l=>lines.push(l))));
  lines.forEach(l=>{ const cy=(l.bbox.y0+l.bbox.y1)/2; let k=st.slots.findIndex(s=>cy>=s[0]-9&&cy<=s[1]+9);
    if(k<0){ let best=1e9; st.slots.forEach((s,i)=>{ const d=Math.min(Math.abs(cy-s[0]),Math.abs(cy-s[1])); if(d<best){best=d;k=i} }) }
    out[k].push({y:l.bbox.y0,t:(l.text||'').trim()}) });
  return out.map(a=>a.sort((p,q)=>p.y-q.y).map(o=>o.t).join(' ').replace(/\s+/g,' ').trim());
}
const AR_DIG={'٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9','۰':'0','۱':'1','۲':'2','۳':'3','۴':'4','۵':'5','۶':'6','۷':'7','۸':'8','۹':'9'};
const digits=s=>String(s||'').replace(/[٠-٩۰-۹]/g,c=>AR_DIG[c]);
function splitFamily(n){ const w=n.split(' '); const last=w[w.length-1]||''; const i=last.lastIndexOf('ال');
  if(i>=2&&last.length-i>=4&&!/عبد$|ابو$|أبو$|بو$/.test(last.slice(0,i))) w.splice(w.length-1,1,last.slice(0,i),last.slice(i)); return w.join(' ') }
function cleanName(s){ return splitFamily(cleanName0(s)) }
function cleanName0(s){ return digits(s).replace(/[|\[\]{}()<>«»"'`~^_=+*#@!?؟:;.,،؛\\\/]/g,' ').replace(/\d+/g,' ').replace(/[A-Za-z]+/g,' ').replace(/\s+/g,' ').trim() }
function gradeOf(s){ const t=digits(s).replace(/[إأآ]/g,'ا');
  if(/ثا?ل?ث|ثلث|الث/.test(t)) return 'ثالث'; if(/ثا?ني|ثان|انى|اني/.test(t)) return 'ثاني'; if(/اول|اوك|ول/.test(t)) return 'أول'; return '' }
function clsOk(s){ const t=digits(s); const m=t.match(/\d{1,2}/); return !!gradeOf(t)&&!!m&&+m[0]>0&&+m[0]<=40 }
function periodsOf(s){ const t=digits(s); const set=new Set(); const re=/(?:ج|ح|خ|<|ع)\s*([1-8])|(?<!\d)([1-8])\s*(?:ج|ح|خ)/g; let m; while((m=re.exec(t))) set.add(+(m[1]||m[2])); return [...set].sort((a,b)=>a-b).map(String) }

/* ---------- main ---------- */
async function read(bytes,onP){
  onP=onP||(()=>{});
  onP('تجهيز صفحات تقرير واضح …');
  const {imgs}=await pageImages(bytes);
  const ws=await getWorkers(onP);
  const rows=[]; let section=1, done=0, wi=0, tail=null, dateTxt=''; const jobs=[];
  const ocrRows=(pr,src)=>{
    const list=[['name',pr,1],['cls',pr,2],['per',pr.filter(r=>r.section===1),2]].filter(j=>j[1].length);
    return Promise.all(list.map(([k,l,sc])=>{ const st=strip(src,l.map(r=>r.cells[k]),18,sc); const w=ws[(wi++)%ws.length];
      return ocrStrip(w,st).then(txt=>{ txt.forEach((t,i)=>l[i][k]=t) }) }))
      .then(()=>{ // second look, larger, at class cells that did not read cleanly
        const bad=pr.filter(r=>!clsOk(r.cls)); if(!bad.length) return;
        const st=strip(src,bad.map(r=>{ const c=r.cells.cls; return {x:c.x+2,y:c.y+2,w:c.w-4,h:c.h-4} }),40,3);
        return ocrStrip(ws[(wi++)%ws.length],st).then(txt=>{ txt.forEach((t,i)=>{ if(clsOk(t)||!bad[i].cls) bad[i].cls=t }) }) })
      .then(()=>pr.forEach(r=>delete r.cells));
  };
  for(let pi=0;pi<imgs.length;pi++){
    onP(`قراءة تقرير واضح: الصفحة ${pi+1} من ${imgs.length} …`);
    const img=await decode(imgs[pi]); const L=layout(img); const W=img.W, D=img.data;
    const red=(x0,x1,y0,y1)=>{ let n=0; for(let y=y0;y<y1;y+=2) for(let x=x0;x<x1;x+=2){ const p=(y*W+x)*4; if(D[p]>150&&D[p+1]<110&&D[p+2]<110) n++ } return n };
    const pr=[]; let lead=null, newTail=null;
    L.tables.forEach((t,ti)=>{
      const xs=t.xs.slice(-6); if(xs.length<6) return;
      const cx=i=>[xs[i]+3,xs[i+1]-3]; // cell i: 0 «حضر في/بعذر», 1 periods, 2 class, 3 name, 4 #
      // section from the header text width over the periods column (long «الحصص التي غاب عنها» vs short «الحصص»)
      if(t.head){ const [p0,p1]=cx(1); const hw=inkWidth(null,W,p0,p1,t.head[0]+4,t.head[1]-4,L.white); section=hw/(p1-p0)>0.32?1:2 }
      const last=t.ys.length-2;
      for(let r=0;r<=last;r++){
        const y0=t.ys[r]+2, y1=t.ys[r+1]-1, h=y1-y0;
        const isLead=ti===0&&r===0&&!!tail, isTail=!!t.cut&&r===last;
        if(h<8||(h<28&&!isLead&&!isTail)) continue;
        if(red(xs[0],xs[5],y0,y1)>40) continue;                              // a red section title, not a row
        const pres=inkCount(L.ink,W,cx(0)[0]+2,cx(0)[1]-2,y0+1,y1-1);
        const frag={section,y0,y1,xs,presentEmpty:pres<25};
        if(isLead){ lead=frag; continue }
        if(isTail){ newTail=frag; continue }
        if(inkCount(L.ink,W,cx(3)[0],cx(3)[1],y0,y1)<40) continue;           // empty row
        const cell=i=>{ const [a,b]=cx(i); return {x:a,y:y0+1,w:b-a,h:h-2} };
        const row={page:pi,section,h,presentEmpty:frag.presentEmpty,cells:{name:cell(3),cls:cell(2),per:cell(1)}};
        rows.push(row); pr.push(row);
      }
    });
    // a row split by the page break: glue the bottom of the previous page to the top of this one
    if(tail){
      const parts=[tail].concat(lead?[{canvas:img.canvas,f:lead}]:[]);
      const H2=parts.reduce((a,p)=>a+(p.f.y1-p.f.y0),0), cv=document.createElement('canvas'); cv.width=W; cv.height=H2+4;
      const x=cv.getContext('2d'); x.fillStyle='#fff'; x.fillRect(0,0,cv.width,cv.height); let yy=2;
      parts.forEach(p=>{ const hh=p.f.y1-p.f.y0; x.drawImage(p.canvas,0,p.f.y0,W,hh,0,yy,W,hh); yy+=hh });
      const xs=tail.f.xs, cx=i=>[xs[i]+3,xs[i+1]-3], cell=i=>{ const [a,b]=cx(i); return {x:a,y:2,w:b-a,h:H2} };
      const row={page:pi-1,section:tail.f.section,h:H2,presentEmpty:tail.f.presentEmpty&&(!lead||lead.presentEmpty),cells:{name:cell(3),cls:cell(2),per:cell(1)}};
      if(H2>=24){ rows.splice(rows.length-pr.length,0,row); jobs.push(ocrRows([row],{canvas:cv})) }
      tail=null;
    }
    if(newTail){ const hh=newTail.y1-newTail.y0, cv=document.createElement('canvas'); cv.width=W; cv.height=newTail.y1;
      cv.getContext('2d').drawImage(img.canvas,0,0,W,newTail.y1,0,0,W,newTail.y1); tail={canvas:cv,f:newTail} }
    if(pi===0){ const W0=img.W, c={x:Math.round(W0*.03),y:Math.round(img.H*.055),w:Math.round(W0*.42),h:Math.round(img.H*.06)};
      jobs.push(ocrStrip(ws[(wi++)%ws.length],strip(img,[c],18,1)).then(t=>{ dateTxt=digits(t[0]||'') }).catch(()=>{})) }
    if(pr.length) jobs.push(ocrRows(pr,img).then(()=>{ done++; onP(`قراءة تقرير واضح: اكتملت ${done} من ${imgs.length} صفحات …`) }));
    if(jobs.length>=ws.length+1) await Promise.all(jobs.splice(0));
  }
  if(tail&&tail.f.y1-tail.f.y0>=28){ const f=tail.f, xs=f.xs, cx=i=>[xs[i]+3,xs[i+1]-3], cell=i=>{ const [a,b]=cx(i); return {x:a,y:f.y0+1,w:b-a,h:f.y1-f.y0-2} };
    const row={page:imgs.length-1,section:f.section,h:f.y1-f.y0,presentEmpty:f.presentEmpty,cells:{name:cell(3),cls:cell(2),per:cell(1)}}; rows.push(row); jobs.push(ocrRows([row],{canvas:tail.canvas})) }
  await Promise.all(jobs);
  const out=[], counts={1:0,2:0};
  rows.forEach(r=>{ counts[r.section]++; const st=toStudent(r,counts[r.section]); if(st) out.push(st) });
  // periods for students absent all day: every period the school recorded that day
  const freq={}; out.filter(x=>x.sec===1).forEach(x=>x.absent.forEach(n=>{ freq[n]=(freq[n]||0)+1 }));
  const maxP=Math.max(6,...Object.keys(freq).filter(k=>freq[k]>=Math.max(3,out.length*.01)).map(Number));
  out.forEach(x=>{ if(x.sec===1) x.absent=x.absent.filter(n=>+n<=maxP) });
  out.forEach(x=>{ if(x.status==='غائب'&&x.sec===2) x.absent=Array.from({length:maxP},(_,i)=>String(i+1)) });
  const dm=[...dateTxt.matchAll(/(\d{4})\s*[\/\-]\s*(\d{1,2})\s*[\/\-]\s*(\d{1,2})/g)].map(m=>({y:+m[1],m:+m[2],d:+m[3]}));
  const hij=dm.find(x=>x.y>1400&&x.y<1600), gr=dm.find(x=>x.y>2000&&x.y<2100);
  const date={hijri:hij?`${hij.d}/${hij.m}/${hij.y}`:'',greg:gr?`${gr.y}-${String(gr.m).padStart(2,'0')}-${String(gr.d).padStart(2,'0')}`:'',raw:dateTxt};
  return {students:out,counts,pages:imgs.length,date};
}
function toStudent(r,n){
  const name=cleanName(r.name); if(!name) return null;
  const cls=digits(r.cls); const mm=cls.match(/\d+/g)||[]; const secN=(mm.find(n=>+n>0&&+n<=40)||'');
  let grade=gradeOf(cls);
  const absentDay=r.section===2||r.presentEmpty;
  const per=periodsOf(r.per);
  return {no:n,name,grade,section:secN,absent:per.length?per:(absentDay?['1','2','3','4','5','6']:[]),status:absentDay?'غائب':'حاضر',src:'wadih',sec:r.section,raw:{cls:r.cls,per:r.per}};
}
// fill missing grades from the section number (this school's sections map to one grade each)
function fixGrades(list){
  const map={}; list.forEach(s=>{ if(s.grade&&s.section){ const k=s.section; map[k]=map[k]||{}; map[k][s.grade]=(map[k][s.grade]||0)+1 } });
  list.forEach(s=>{ const m=map[s.section]; if(m){ const best=Object.keys(m).sort((a,b)=>m[b]-m[a])[0]; if(!s.grade||(m[s.grade]||0)*3<m[best]) s.grade=best } });
  return list;
}
// the report lists each class together: a row whose class could not be read takes it from matching neighbours
function fillFromNeighbours(list){
  const ok=x=>x&&x.grade&&x.section;
  for(let i=0;i<list.length;i++){ const x=list[i]; if(ok(x)) continue;
    let a=i-1; while(a>=0&&!ok(list[a])) a--; let b=i+1; while(b<list.length&&!ok(list[b])) b++;
    const A=list[a], B=list[b];
    if(A&&B&&A.sec===x.sec&&B.sec===x.sec&&A.grade===B.grade&&A.section===B.section){ x.grade=A.grade; x.section=A.section; x.guessed=true }
    else if(x.section&&!x.grade){ const c=[A,B].find(y=>y&&y.section===x.section); if(c){ x.grade=c.grade; x.guessed=true } } }
  return list;
}
async function parse(bytes,onP){ const r=await read(bytes,onP); r.students=fillFromNeighbours(fixGrades(r.students)); return r }
async function terminate(){ if(workers){ const w=workers; workers=null; for(const x of w) try{ await x.terminate() }catch(e){} } }
window.Wadih={looksLike,isImagePdf,parse,terminate,_layout:layout,_decode:decode,_pageImages:pageImages};
})();
