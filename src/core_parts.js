/* ---- pdf text ---- */
// Minimal content-stream text extractor honouring /ActualText (needed for Chrome-printed Arabic PDFs)
function makeExtractor(PDFLib){
const {PDFDocument,PDFName,PDFArray,PDFDict,PDFRawStream,PDFRef,decodePDFRawStream,PDFHexString,PDFString,PDFNumber}=PDFLib;
const isWS=c=>c===0x20||c===0x0a||c===0x0d||c===0x09||c===0x0c||c===0;
const isDelim=c=>"()<>[]{}/%".includes(String.fromCharCode(c));
function tokenize(b){
  const out=[];let i=0,n=b.length;
  while(i<n){
    const c=b[i];
    if(isWS(c)){i++;continue}
    if(c===0x25){while(i<n&&b[i]!==0x0a&&b[i]!==0x0d)i++;continue}
    if(c===0x2f){let j=i+1;while(j<n&&!isWS(b[j])&&!isDelim(b[j]))j++;out.push({t:'name',v:String.fromCharCode(...b.subarray(i+1,j))});i=j;continue}
    if(c===0x28){let depth=1,j=i+1,bytes=[];
      while(j<n&&depth>0){let d=b[j];
        if(d===0x5c){j++;const e=b[j];
          const map={110:10,114:13,116:9,98:8,102:12,40:40,41:41,92:92};
          if(e in map){bytes.push(map[e]);j++}
          else if(e>=48&&e<=55){let o='';while(o.length<3&&b[j]>=48&&b[j]<=55){o+=String.fromCharCode(b[j]);j++}bytes.push(parseInt(o,8))}
          else if(e===10||e===13){j++}
          else{bytes.push(e);j++}
          continue}
        if(d===0x28)depth++; if(d===0x29){depth--; if(depth===0){j++;break}}
        bytes.push(d);j++}
      out.push({t:'str',v:bytes});i=j;continue}
    if(c===0x3c&&b[i+1]===0x3c){out.push({t:'<<'});i+=2;continue}
    if(c===0x3e&&b[i+1]===0x3e){out.push({t:'>>'});i+=2;continue}
    if(c===0x3c){let j=i+1,h='';while(j<n&&b[j]!==0x3e){if(!isWS(b[j]))h+=String.fromCharCode(b[j]);j++}
      if(h.length%2)h+='0';const bytes=[];for(let k=0;k<h.length;k+=2)bytes.push(parseInt(h.substr(k,2),16));
      out.push({t:'str',v:bytes});i=j+1;continue}
    if(c===0x5b){out.push({t:'['});i++;continue}
    if(c===0x5d){out.push({t:']'});i++;continue}
    let j=i;while(j<n&&!isWS(b[j])&&!isDelim(b[j]))j++;
    if(j===i){i++;continue}
    const s=String.fromCharCode(...b.subarray(i,j));
    if(/^[+-]?(\d+\.?\d*|\.\d+)$/.test(s))out.push({t:'num',v:parseFloat(s)});
    else out.push({t:'op',v:s});
    i=j;
  }
  return out;
}
// parse operands into JS values
function parseOps(tokens){
  const ops=[];let stack=[];
  const st=[]; // container stack
  const push=v=>{ if(st.length){const top=st[st.length-1]; top.items.push(v)} else stack.push(v)};
  for(const tk of tokens){
    if(tk.t==='['){st.push({k:'a',items:[]});continue}
    if(tk.t==='<<'){st.push({k:'d',items:[]});continue}
    if(tk.t===']'||tk.t==='>>'){const top=st.pop(); let v;
      if(top.k==='a')v=top.items; else {v={};for(let i=0;i+1<top.items.length;i+=2){const k=top.items[i];v[k&&k.name?k.name:k]=top.items[i+1]}}
      push(v);continue}
    if(tk.t==='num'){push(tk.v);continue}
    if(tk.t==='name'){push({name:tk.v});continue}
    if(tk.t==='str'){push({bytes:tk.v});continue}
    if(tk.t==='op'){
      if(st.length){push({op:tk.v});continue}
      if(tk.v==='true'||tk.v==='false'||tk.v==='null'){push(tk.v==='true');continue}
      ops.push({op:tk.v,args:stack});stack=[];
    }
  }
  return ops;
}
function utf16(bytes){ // FEFF big endian
  let s='',i=0; if(bytes[0]===0xfe&&bytes[1]===0xff)i=2; else return String.fromCharCode(...bytes);
  for(;i+1<bytes.length;i+=2)s+=String.fromCharCode((bytes[i]<<8)|bytes[i+1]);
  return s;
}
function parseCMap(text){
  const map=new Map(); let m;
  const hex=h=>{let s='';for(let i=0;i<h.length;i+=4)s+=String.fromCharCode(parseInt(h.substr(i,4),16));return s};
  const bfchar=/beginbfchar([\s\S]*?)endbfchar/g;
  while((m=bfchar.exec(text))){const re=/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]*)>/g;let r;while((r=re.exec(m[1])))map.set(parseInt(r[1],16),hex(r[2].padEnd(4,'0')))}
  const bfrange=/beginbfrange([\s\S]*?)endbfrange/g;
  while((m=bfrange.exec(text))){const re=/<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<([0-9A-Fa-f]+)>|\[([^\]]*)\])/g;let r;
    while((r=re.exec(m[1]))){const a=parseInt(r[1],16),b=parseInt(r[2],16);
      if(r[4]!==undefined){const base=parseInt(r[4],16);for(let k=a;k<=b;k++)map.set(k,String.fromCharCode(base+k-a))}
      else{const arr=[...r[5].matchAll(/<([0-9A-Fa-f]+)>/g)].map(x=>hex(x[1]));for(let k=a;k<=b;k++)map.set(k,arr[k-a]||'')}}}
  return map;
}
const mul=(a,b)=>[a[0]*b[0]+a[1]*b[2],a[0]*b[1]+a[1]*b[3],a[2]*b[0]+a[3]*b[2],a[2]*b[1]+a[3]*b[3],a[4]*b[0]+a[5]*b[2]+b[4],a[4]*b[1]+a[5]*b[3]+b[5]];
async function extract(bytes){
  const doc=await PDFDocument.load(bytes,{ignoreEncryption:true,updateMetadata:false});
  const ctx=doc.context; const L=x=>x instanceof PDFRef?ctx.lookup(x):x;
  const streamBytes=s=>{s=L(s); if(s instanceof PDFRawStream) return decodePDFRawStream(s).decode(); return s.getContents? s.getContents(): new Uint8Array()};
  const fontCache=new Map();
  function fontInfo(ref){
    const key=ref.toString(); if(fontCache.has(key))return fontCache.get(key);
    const f=L(ref); const sub=f.lookup(PDFName.of('Subtype'))?.toString();
    const tu=f.get(PDFName.of('ToUnicode')); let map=new Map(); const info0={};
    if(tu){try{map=parseCMap(new TextDecoder('latin1').decode(streamBytes(tu)))}catch(e){}}
    const W=new Map(); let DW=1000; const num=x=>{x=L(x);return x&&x.asNumber?x.asNumber():(x&&x.numberValue!==undefined?x.numberValue:+x)};
    try{
      if(sub==='/Type0'){const df=L(L(f.get(PDFName.of('DescendantFonts'))).get(0));
        const dw=df.get(PDFName.of('DW')); if(dw)DW=num(dw);
        const w=L(df.get(PDFName.of('W'))); if(w){const a=[];for(let i=0;i<w.size();i++)a.push(L(w.get(i)));
          for(let i=0;i<a.length;){const first=num(a[i]); const nx=a[i+1];
            if(nx instanceof PDFArray){for(let k=0;k<nx.size();k++)W.set(first+k,num(nx.get(k)));i+=2}
            else{const last=num(nx),wv=num(a[i+2]);for(let k=first;k<=last;k++)W.set(k,wv);i+=3}}}}
      else{const fc=f.get(PDFName.of('FirstChar')),ws=L(f.get(PDFName.of('Widths')));
        if(fc&&ws){const f0=num(fc);for(let k=0;k<ws.size();k++)W.set(f0+k,num(ws.get(k)))} DW=0;
        const fm=L(f.get(PDFName.of('FontMatrix'))); if(fm){info0.scale=num(fm.get(0))*1000}}
    }catch(e){}
    const info={two:sub==='/Type0',map,W,DW,scale:info0.scale||1}; fontCache.set(key,info); return info;
  }
  const pages=[];
  for(const page of doc.getPages()){
    const {width,height}=page.getSize();
    const res=page.node.Resources(); const fonts=res?L(res.get(PDFName.of('Font'))):null;
    let c=page.node.Contents(); c=L(c); const parts=[];
    if(c instanceof PDFArray){for(let i=0;i<c.size();i++)parts.push(streamBytes(c.get(i)))} else if(c) parts.push(streamBytes(c));
    const total=parts.reduce((a,p)=>a+p.length+1,0); const buf=new Uint8Array(total); let o=0; for(const p of parts){buf.set(p,o);o+=p.length;buf[o++]=0x0a}
    const ops=parseOps(tokenize(buf));
    let ctm=[1,0,0,1,0,0]; const gs=[]; let tm=[1,0,0,1,0,0], tlm=[1,0,0,1,0,0], font=null, size=1, TL=0;
    const mc=[]; // marked content stack: {actual, rev, runs, pos}
    const runs=[];
    let group=null; // ReversedChars collector
    const pos=()=>{const m=mul(tm,ctm);return {x:m[4],y:height-m[5],s:Math.abs(size*Math.hypot(m[2],m[3]))}};
    const emit=(str,p,x2)=>{ if(!str)return; const r={str,x:p.x,y:p.y,s:p.s,x2:x2!==undefined?x2:p.x+p.s*0.5}; const g=mc.slice().reverse().find(e=>e.rev); if(g){g.parts.push(r)} else runs.push(r)};
    let Tc=0,Tw=0,Th=1;
    const adv=(code)=>{ if(!font)return; const w=(font.W.has(code)?font.W.get(code):font.DW)*font.scale/1000; let tx=(w*size+Tc+(code===32&&!font.two?Tw:0))*Th; tm=mul([1,0,0,1,tx,0],tm)};
    const codes=(bytes)=>{const out=[]; if(font&&font.two){for(let i=0;i+1<bytes.length;i+=2)out.push((bytes[i]<<8)|bytes[i+1])} else for(const b of bytes)out.push(b); return out};
    const show=(bytes)=>{ const cs=codes(bytes);
      const act=mc.slice().reverse().find(e=>e.actual!==undefined);
      if(act){ for(const c of cs){const q=pos(); act.p=q; if(act.mn===undefined||q.x<act.mn)act.mn=q.x; adv(c); const e=pos().x; if(act.mx===undefined||e>act.mx)act.mx=e} return }
      if(!font)return;
      for(const c of cs){ const u=font.map.get(c); let ch; if(font.two) ch=u&&u!=='\u0000'?u:''; else ch=u!==undefined?(u==='\u0000'?'':u):String.fromCharCode(c);
        const q=pos(); adv(c); if(ch) emit(ch,q,pos().x) }
    };
    const show_old=(bytes)=>{
      const act=mc.slice().reverse().find(e=>e.actual!==undefined);
      if(act){ act.p=pos(); act.n=(act.n||0)+1; return }
      if(!font)return; let s='';
      if(font.two){for(let i=0;i+1<bytes.length;i+=2){const code=(bytes[i]<<8)|bytes[i+1];const u=font.map.get(code);s+=u&&u!=='\u0000'?u:''}}
      else{for(const code of bytes){const u=font.map.get(code);s+=u!==undefined?(u==='\u0000'?'':u):String.fromCharCode(code)}}
      emit(s,pos());
    };
    const segs=[]; let cur=null;
    const dev=(x,y)=>{const m=ctm;return [x*m[0]+y*m[2]+m[4], height-(x*m[1]+y*m[3]+m[5])]};
    const addSeg=(a,b)=>{segs.push({x1:Math.min(a[0],b[0]),x2:Math.max(a[0],b[0]),y1:Math.min(a[1],b[1]),y2:Math.max(a[1],b[1])})};
    for(const {op,args} of ops){
      switch(op){
        case 're':{const [x,y,w,h]=args; const a=dev(x,y),b=dev(x+w,y+h); const r={x1:Math.min(a[0],b[0]),x2:Math.max(a[0],b[0]),y1:Math.min(a[1],b[1]),y2:Math.max(a[1],b[1])};
          if(r.x2-r.x1<2.5||r.y2-r.y1<2.5) segs.push(r); else {addSeg([r.x1,r.y1],[r.x1,r.y2]);addSeg([r.x2,r.y1],[r.x2,r.y2]);addSeg([r.x1,r.y1],[r.x2,r.y1]);addSeg([r.x1,r.y2],[r.x2,r.y2]);segs.push(Object.assign({box:1},r))}
          break}
        case 'm':cur=dev(args[0],args[1]);break;
        case 'l':{const p=dev(args[0],args[1]); if(cur)addSeg(cur,p); cur=p;break}
        case 'q':gs.push(ctm.slice());break;
        case 'Q':ctm=gs.pop()||[1,0,0,1,0,0];break;
        case 'cm':ctm=mul(args,ctm);break;
        case 'BT':tm=[1,0,0,1,0,0];tlm=[1,0,0,1,0,0];break;
        case 'Tf':{const nm=args[0]?.name; size=args[1]; const ref=fonts&&fonts.get(PDFName.of(nm)); font=ref?fontInfo(ref):null;break}
        case 'Tm':tm=args.slice(0,6);tlm=tm.slice();break;
        case 'Td':tlm=mul([1,0,0,1,args[0],args[1]],tlm);tm=tlm.slice();break;
        case 'TD':TL=-args[1];tlm=mul([1,0,0,1,args[0],args[1]],tlm);tm=tlm.slice();break;
        case 'TL':TL=args[0];break;
        case 'T*':tlm=mul([1,0,0,1,0,-TL],tlm);tm=tlm.slice();break;
        case 'Tj':show(args[0].bytes);break;
        case "'":tlm=mul([1,0,0,1,0,-TL],tlm);tm=tlm.slice();show(args[0].bytes);break;
        case '"':tlm=mul([1,0,0,1,0,-TL],tlm);tm=tlm.slice();show(args[2].bytes);break;
        case 'TJ':{for(const a of args[0]){ if(a&&a.bytes)show(a.bytes); else if(typeof a==='number'){tm=mul([1,0,0,1,-a/1000*size*Th,0],tm)}} break}
        case 'Tc':Tc=args[0];break; case 'Tw':Tw=args[0];break; case 'Tz':Th=args[0]/100;break;
        case 'BMC':mc.push({rev:args[0]?.name==='ReversedChars',parts:[]});break;
        case 'BDC':{const d=args[1]; const e={parts:[]}; if(d&&typeof d==='object'&&d.ActualText&&d.ActualText.bytes)e.actual=utf16(d.ActualText.bytes); if(args[0]?.name==='ReversedChars')e.rev=true; mc.push(e);break}
        case 'EMC':{const e=mc.pop(); if(e&&e.actual!==undefined&&e.p){emit(e.actual,{x:e.mn,y:e.p.y,s:e.p.s},e.mx)}
          if(e&&e.rev&&e.parts.length){ // visual order -> logical (right-to-left by x)
            const ps=e.parts.slice().sort((a,b)=>(b.x+b.x2)-(a.x+a.x2)); const str=ps.map(p=>p.str).join(''); const xs=ps.map(p=>p.x);
            const r={str,x:Math.min(...xs),y:ps[0].y,s:ps[0].s,x2:Math.max(...ps.map(p=>p.x2))}; const g=mc.slice().reverse().find(k=>k.rev); if(g)g.parts.push(r); else runs.push(r)}
          break}
      }
    }
    pages.push({width,height,runs,segs});
  }
  return pages;
}
return {extract};
}

/* ---- analysis ---- */
// ---------- generic helpers ----------
const AR=/[؀-ۿﭐ-﷿ﹰ-﻿]/;
const LTRTOK=/^[A-Za-z0-9٠-٩:.\-+\/()&'’,_%#=]+$/;
function norm(s){return s.replace(/ـ/g,'').replace(/\s+/g,' ').trim()}
// assemble runs (same visual line) into logical text
function lineText(runs){
  const t=runs.slice().sort((a,b)=>a.x-b.x);
  if(!t.some(r=>AR.test(r.str))) return norm(t.map(r=>r.str).join(''));
  // group LTR runs (with internal spaces) together, keep their order; reverse the rest
  const groups=[]; let i=0;
  while(i<t.length){
    if(LTRTOK.test(t[i].str)&&!AR.test(t[i].str)){
      let j=i, last=i;
      while(j<t.length && !AR.test(t[j].str) && (LTRTOK.test(t[j].str)||t[j].str.trim()==='')){ if(LTRTOK.test(t[j].str)) last=j; j++ }
      groups.push(t.slice(i,last+1).map(r=>r.str).join('')); i=last+1;
    } else { groups.push(t[i].str); i++ }
  }
  return norm(groups.reverse().join(''));
}
function toLines(runs,tol=3){
  const s=runs.filter(r=>r.str!=='').slice().sort((a,b)=>a.y-b.y||a.x-b.x); const lines=[];
  for(const r of s){ const L=lines.find(l=>Math.abs(l.y-r.y)<=tol); if(L){L.runs.push(r)} else lines.push({y:r.y,runs:[r]}) }
  lines.sort((a,b)=>a.y-b.y);
  return lines.map(l=>({y:l.y,text:lineText(l.runs),runs:l.runs})).filter(l=>l.text);
}
const cx=r=>(r.x+(r.x2??r.x))/2;
function verticals(pg,yIn){
  const xs=pg.segs.filter(s=>!s.box&&s.x2-s.x1<2.5&&s.y2-s.y1>6&&(yIn===undefined||(s.y1<=yIn+1&&s.y2>=yIn-1))).map(s=>(s.x1+s.x2)/2).sort((a,b)=>a-b);
  const out=[]; for(const x of xs){ if(!out.length||x-out[out.length-1]>3) out.push(x) } return out;
}
const PERIOD_WORDS=['الأولى','الثانية','الثالثة','الرابعة','الخامسة','السادسة','السابعة','الثامنة','التاسعة','العاشرة'];
const GRADE_RE=/(أول|ثاني|ثالث)/;

// ---------- teachers / lessons ----------
function parseTeachers(pages){
  const teachers=[]; let cur=null; let totals=null; let dateText='';
  const carry={}; // last box per column for page continuation
  pages.forEach((pg,pi)=>{
    const lines=toLines(pg.runs);
    if(pi===0){ const h=lines.find(l=>/تقرير/.test(l.text)&&/بتاريخ/.test(l.text)); if(h) dateText=h.text }
    // header row: runs that are period words
    const nameHdr=pg.runs.find(r=>r.str.trim()==='الاسم'); if(!nameHdr) return;
    const hy=nameHdr.y;
    const hdr=pg.runs.filter(r=>Math.abs(r.y-hy)<3&&(PERIOD_WORDS.includes(r.str.trim())||r.str.trim()==='الاسم'||r.str.trim()==='م')); const vx=verticals(pg,hy);
    const cell=r=>{const c=cx(r); let a=null,b=null; for(const x of vx){ if(x<=c)a=x; else {b=x;break} } return [a,b]};
    const cols={}; // name -> [x1,x2]
    for(const r of hdr){ cols[r.str.trim()]=cell(r) }
    const inCol=(r,k)=>{const c=cx(r),[a,b]=cols[k]||[];return a!==null&&a!==undefined&&b!==undefined&&b!==null&&c>a&&c<b};
    // horizontal separators in name column
    const [nx1,nx2]=cols['الاسم']||[0,0];
    const seps=[...new Set(pg.segs.filter(s=>!s.box&&s.y2-s.y1<2.5&&s.x1<nx2-2&&s.x2>nx1+2&&s.y1>hy+4&&s.y1<pg.height-30).map(s=>Math.round(s.y1)))].sort((a,b)=>a-b);
    const bounds=[]; let prev=hy+4; for(const y of seps){ if(y-prev>6){bounds.push([prev,y])} prev=y } bounds.push([prev,pg.height-28]);
    // totals row on last page
    const totLine=lines.find(l=>/الإجمالي/.test(l.text));
    if(totLine){ const nums=toLines(pg.runs.filter(r=>Math.abs(r.y-totLine.y)<4&&/\d/.test(r.str))).map(l=>l.runs).flat().sort((a,b)=>b.x-a.x).map(r=>r.str.trim());
      // join digit runs by x proximity
      const rr=pg.runs.filter(r=>Math.abs(r.y-totLine.y)<4&&/^\d$/.test(r.str.trim())).sort((a,b)=>a.x-b.x); const groups=[]; for(const r of rr){const g=groups[groups.length-1]; if(g&&r.x-g.x2<3){g.s+=r.str.trim();g.x2=r.x2} else groups.push({s:r.str.trim(),x:r.x,x2:r.x2})}
      const hl=lines.find(l=>/كل الدروس/.test(l.text)); 
      if(groups.length>=4){ const v=groups.sort((a,b)=>b.x-a.x).map(g=>+g.s); totals={all:v[0],async:v[1],prepared:v[2],notPrepared:v[3]} } }
    for(const [y1,y2] of bounds){
      const inRow=pg.runs.filter(r=>r.y>y1&&r.y<y2);
      if(!inRow.length) continue;
      const numRun=inRow.filter(r=>inCol(r,'م')&&/\d/.test(r.str));
      const nameRuns=inRow.filter(r=>inCol(r,'الاسم'));
      if(numRun.length&&nameRuns.length){
        const no=+numRun.sort((a,b)=>a.x-b.x).map(r=>r.str.trim()).join('');
        const name=toLines(nameRuns).map(l=>l.text).join(' ');
        cur={no,name:norm(name),lessons:[]}; teachers.push(cur);
        for(const k in carry) delete carry[k];
      }
      if(!cur) continue;
      // lesson boxes per period column
      PERIOD_WORDS.forEach((pw,pidx)=>{
        if(!cols[pw]) return;
        const cr=inRow.filter(r=>inCol(r,pw)); if(!cr.length) return;
        const [ca,cb]=cols[pw];
        // header rects in this column/row
        const rects=pg.segs.filter(s=>s.box&&s.x1>=ca-2&&s.x2<=cb+2&&s.y1>=y1-2&&s.y2<=y2+2&&(s.y2-s.y1)<60);
        const heads=[];
        for(const rc of rects){ const tr=cr.filter(r=>r.y>rc.y1-1&&r.y<rc.y2+1&&cx(r)>rc.x1&&cx(r)<rc.x2); const txt=toLines(tr).map(l=>l.text).join(' ');
          if(GRADE_RE.test(txt)&&/\d/.test(txt)) heads.push({rc,txt,runs:new Set(tr)}) }
        heads.sort((a,b)=>a.rc.y1-b.rc.y1);
        // dedupe overlapping rect detections
        const H=[]; for(const h of heads){ if(!H.length||Math.abs(H[H.length-1].rc.y1-h.rc.y1)>4) H.push(h) }
        const headRuns=new Set(H.flatMap(h=>[...h.runs]));
        const body=toLines(cr.filter(r=>!headRuns.has(r)));
        // continuation lines before first header
        const firstY=H.length?H[0].rc.y1:Infinity;
        const cont=body.filter(l=>l.y<firstY).map(l=>l.text);
        if(cont.length&&carry[pw]) carry[pw].details.push(...cont);
        H.forEach((h,i)=>{
          const nextY=i+1<H.length?H[i+1].rc.y1:Infinity;
          const det=body.filter(l=>l.y>h.rc.y1&&l.y<nextY).map(l=>l.text);
          const g=(h.txt.match(GRADE_RE)||[])[1]; const sec=+(h.txt.replace(/[^\d]/g,''));
          const ok=!/X|×|✗/i.test(h.txt);
          const box={grade:g,section:sec,period:pidx+1,ok,details:det};
          cur.lessons.push(box); carry[pw]=box;
        });
      });
    }
  });
  // finalize lessons
  let raw=0,rawAsync=0,rawOk=0;
  for(const t of teachers){
    const seen=new Map();
    for(const b of t.lessons){
      raw++; const d=norm(b.details.join(' '));
      b.async=/غير متزامن/.test(d); b.notPrepared=!b.ok||/لم يتم إعداد/.test(d);
      if(b.async)rawAsync++; if(!b.notPrepared)rawOk++;
      let m=d.match(/عنوان الدرس\s*:\s*(.*?)\s*(التوقيت|غير متزامن|لم يتم|رابط|$)/); b.title=m?fixTitle(m[1]):'';
      const key=b.grade+'|'+b.section+'|'+b.period;
      if(seen.has(key)){ const o=seen.get(key); if(!o.title&&b.title)o.title=b.title; o.dup=(o.dup||1)+1; b.drop=true } else seen.set(key,b);
    }
    t.lessons=t.lessons.filter(b=>!b.drop).sort((a,b)=>a.period-b.period);
    const titles=[...new Set(t.lessons.filter(b=>!b.notPrepared&&b.title).map(b=>b.title))];
    for(const b of t.lessons) if(!b.title) b.title=titles[0]||'';
    t.prepared=t.lessons.filter(b=>!b.notPrepared).length;
    t.notPrepared=t.lessons.filter(b=>b.notPrepared).length;
    t.async=t.lessons.filter(b=>b.async).length;
    t.titles=[...new Set(t.lessons.map(b=>b.title).filter(Boolean))];
  }
  return {teachers,totals,dateText,raw:{all:raw,async:rawAsync,prepared:rawOk}};
}

// ---------- students ----------
function parseStudents(pages){
  const students=[]; let dateText='';
  pages.forEach((pg,pi)=>{
    const lines=toLines(pg.runs);
    if(pi===0){const h=lines.find(l=>/تقرير/.test(l.text)); if(h) dateText=h.text}
    const hdrStatus=pg.runs.find(r=>r.str.trim()==='الحالة'); if(!hdrStatus) return;
    const hy=hdrStatus.y; const vx=verticals(pg);
    const H={}; for(const r of pg.runs.filter(r=>Math.abs(r.y-hy)<3)) H[r.str.trim()]=cx(r);
    const colOf=x=>{let a=-1; for(let i=0;i<vx.length;i++){ if(vx[i]<=x)a=i } return a};
    const key={}; for(const [k,x] of Object.entries(H)) key[colOf(x)]=k;
    const hs=[...new Set(pg.segs.filter(s=>!s.box&&s.y2-s.y1<2.5&&s.x2-s.x1>100&&s.y1>hy).map(s=>Math.round(s.y1)))].sort((a,b)=>a-b);
    for(let i=0;i+1<hs.length;i++){
      const y1=hs[i],y2=hs[i+1]; if(y2-y1<5) continue;
      const rr=pg.runs.filter(r=>r.y>y1&&r.y<y2); if(!rr.length) continue;
      const byCol={}; for(const r of rr){ const k=key[colOf(cx(r))]; if(!k) continue; (byCol[k]=byCol[k]||[]).push(r) }
      const txt=k=>byCol[k]?toLines(byCol[k],4).map(l=>l.text).join(' '):'';
      const status=txt('الحالة'); if(!/حاضر|غائب/.test(status)) continue;
      const cls=txt('الفصل'); const m=cls.match(/(أول|ثاني|ثالث)\s*ش\s*(\d+)/);
      const list=s=>s.split(/[،,]/).map(x=>x.trim()).filter(x=>x&&x!=='0'&&x!=='٠');
      students.push({no:+txt('م').replace(/\D/g,''),name:txt('الطالب'),classFull:cls,grade:m?m[1]:'',section:m?+m[2]:null,
        present:list(txt('الحصص الحاضرة')),absent:list(txt('الحصص الغائبة')),status:/غائب/.test(status)?'غائب':'حاضر'});
    }
  });
  return {students,dateText};
}

function fixTitle(s){ s=norm(s);
  s=s.replace(/\(\)\s*(\d+)/g,'($1)').replace(/(\d+)\s*\(\)/g,'($1)').replace(/\)\s*(\d+)\s*\(/g,'($1)');
  s=s.replace(/(\d{4})-(\d{4})/g,(m,a,b)=>+a>+b?b+'-'+a:m);
  return s.replace(/\s+([،,.:])/g,'$1').replace(/\s*\.$/,'').trim();
}
