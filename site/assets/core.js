'use strict';
const Core=(function(){
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

/* ---- report model ---- */
const GRADE_ORDER={'أول':1,'ثاني':2,'ثالث':3};
const SUBJECT_RULES=[
 // أسماء المواد الصريحة أولًا (مثل «الفيزياء 3 – لم يتم إعداد الدروس»)
 [/^\s*(الفيزياء|الكيمياء|الأحياء|الرياضيات|الجغرافيا|التاريخ|الاقتصاد|القرآن|الحاسب|التقنية الرقمية)(?=[\s\d\-–:]|$)/,'__NAME__'],
 [/تفسير|سورة|الآية|حديث|فقه|توحيد/,'الدراسات الإسلامية'],
 [/التفاعلات|الهيدروجين|الهيدروجيني|الأيونات|أيونات|الأملاح|التأكسد|الاتزان الكيميائي|الصيغة الأولية|الصيغة الجزيئية|الذرة|المول|الكيمياء|الأحماض|المحاليل|الروابط/,'كيمياء'],
 [/الأمن السيبراني|HTML|التقنيات الناشئة|البرمجة|الحاسب(?!ة)|قواعد البيانات|موقع إلكتروني/,'التقنية الرقمية'],
 [/المثلثات|المتباينات|كثيرات الحدود|الدوال|اللوغاريتم|القانون العام|المصفوفات|الاحتمال|الإحصاء/,'رياضيات'],
 [/الدفع|الزخم|المرايا|انكسار الضوء|حفظ الطاقة|الحركة|القوى|الكهرباء|المغناطيس|العدسات|الموجات/,'فيزياء'],
 [/الفطريات|الإسفنجيات|اللاسعات|النبات|الأزهار|الخلايا|هرمونات|الحيوانات|الوراثة/,'أحياء'],
 [/الصخور|القارات|الصفائح|قاع المحيط|الزلازل|البراكين|المعادن/,'علم الأرض والفضاء'],
 [/الدولة السعودية|التاريخ|معارك/,'تاريخ'],
 [/القراءة|الكتابة|النحو|البلاغة|الأدب/,'الكفايات اللغوية'],
 [/التفكير الناقد|المناظرة|وجاهة الشواهد|الفكر الضال/,'التفكير الناقد'],
 [/البحث العلمي|مصادر المعلومات|قواعد المعلومات/,'البحث ومصادر المعلومات'],
 [/كرة|اللياقة|البدنية/,'التربية البدنية'],
 [/إسعافات|الصحية/,'اللياقة والثقافة الصحية'],
 [/الإنفاق|المالي|الادخار/,'المعرفة المالية'],
 [/البيئية|البيئة|الطاقة المستدامة/,'علم البيئة'],
 [/السكرتارية|المهارات الإدارية|مهارات إدارية|الإدارة المكتبية|حل المشكلات|اتخاذ القرار|إدارة الوقت/,'مهارات إدارية'],
 [/النظم البحرية|الساحلية|التنمية المستدامة|تنمية مستدامة/,'تنمية مستدامة'],
 [/[A-Za-z]{3,}/,'لغة إنجليزية'],
 // عناوين عامة تُستخدم في الكفايات اللغوية بهذه المدرسة
 [/النشاطات التمهيدية|الاختبار القبلي|الغلق|الفلق|التلخيص/,'الكفايات اللغوية'],
];
// normalise Arabic text before matching: tatweel, invisible direction marks, hamza/alef and yaa variants
const normAr=x=>String(x||'').replace(/[ـ​-‏‪-‮⁦-⁩﻿]/g,'').replace(/[أإآ]/g,'ا').replace(/ى/g,'ي').replace(/\s+/g,' ').trim();
const SUBJECT_RULES_N=SUBJECT_RULES.map(([re,name])=>[new RegExp(normAr(re.source),re.flags),name]);
const SUBJECT_CANON={'الفيزياء':'فيزياء','الكيمياء':'كيمياء','الاحياء':'أحياء','الرياضيات':'رياضيات','الحاسب':'التقنية الرقمية','التاريخ':'تاريخ','القران':'القرآن الكريم','الجغرافيا':'جغرافيا','الاقتصاد':'اقتصاد','التقنية الرقمية':'التقنية الرقمية'};
function guessSubject(title){ const s=[]; for(const raw of String(title||'').split(/\s*\/\s*/)){ const part=normAr(raw); if(!part) continue; for(const [re,name] of SUBJECT_RULES_N){ const m=re.exec(part); if(m){ const label=name==='__NAME__'?(SUBJECT_CANON[m[1]]||m[1]):name; if(!s.includes(label)) s.push(label); break } } } return s.join(' / ') }
const HMONTHS=[['محرم',1],['صفر',2],['ربيع الأول',3],['ربيع الاول',3],['ربيع الآخر',4],['ربيع الثاني',4],['جمادى الأولى',5],['جمادى الاولى',5],['جمادى الآخرة',6],['جمادى الثانية',6],['رجب',7],['شعبان',8],['رمضان',9],['شوال',10],['ذو القعدة',11],['ذو الحجة',12]];
const GMONTHS={'يناير':1,'فبراير':2,'مارس':3,'أبريل':4,'ابريل':4,'مايو':5,'يونيو':6,'يوليو':7,'أغسطس':8,'اغسطس':8,'سبتمبر':9,'أكتوبر':10,'اكتوبر':10,'نوفمبر':11,'ديسمبر':12};
const DAYS=['الأحد','الاثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'];
function parseDates(tDate,sDate){
  let hijri=null,greg=null,m;
  if(tDate){ for(const [n,k] of HMONTHS){ const re=new RegExp('(\\d{1,2})\\s*'+n+'\\s*(\\d{4})'); if((m=tDate.match(re))){hijri={d:+m[1],m:k,y:+m[2]};break} }
    for(const [n,k] of Object.entries(GMONTHS)){ const re=new RegExp('(\\d{1,2})\\s*'+n+'\\s*(\\d{4})'); if((m=tDate.match(re))){greg=new Date(Date.UTC(+m[2],k-1,+m[1]));break} } }
  if(!greg&&sDate&&(m=sDate.match(/(\d{1,2})-(\d{1,2})-(\d{4})/))) greg=new Date(Date.UTC(+m[3],+m[2]-1,+m[1]));
  if(!hijri&&greg){ try{ const parts=new Intl.DateTimeFormat('en-u-ca-islamic-umalqura',{day:'numeric',month:'numeric',year:'numeric',timeZone:'UTC'}).formatToParts(greg); const g=k=>+parts.find(p=>p.type===k).value; hijri={d:g('day'),m:g('month'),y:g('year')} }catch(e){} }
  return {hijriText:hijri?`${hijri.d}/${hijri.m}/${hijri.y}هـ`:'',day:greg?DAYS[greg.getUTCDay()]:'',gregText:greg?greg.toISOString().slice(0,10):''};
}
const pct=(a,b)=>b?(Math.round(a/b*1000)/10).toString().replace(/\.0$/,'')+'%':'—';
function buildModel(T,S){
  const teachers=T.teachers.map(t=>{
    const groups={}; t.lessons.forEach(b=>{ (groups[b.grade]=groups[b.grade]||[]).push('ش'+b.section) });
    const cls=Object.entries(groups).map(([g,v])=>`${g} ث: ${v.join('، ')}`).join(' / ');
    const periods=t.lessons.map(b=>b.period).join('، ');
    const uniq=a=>{a=[...new Set(a.filter(Boolean))];return a.filter(x=>!a.some(y=>y!==x&&y.includes(x)))};
    const okTitles=uniq(t.lessons.filter(b=>!b.notPrepared).map(b=>b.title));
    const npTitles=[...new Set(t.lessons.filter(b=>b.notPrepared).map(b=>b.title).filter(Boolean))];
    let title=okTitles.join(' / ');
    const notes=[];
    if(t.prepared===0) title=(npTitles.join(' / ')||'')+(npTitles.length?' — ':'')+'لم يتم إعداد الدروس';
    else if(t.notPrepared){ notes.push(t.lessons.filter(b=>b.notPrepared).map(b=>`${b.grade} ش${b.section}`).join('، ')+' لم تُعدّ') }
    if(t.async&&t.async===t.lessons.length) notes.push('غير متزامن');
    else if(t.async) notes.push(`${t.async} غير متزامن`);
    const subject=guessSubject(okTitles.concat(npTitles).join(' / '));
    return {no:t.no,name:t.name,cls,periods,title:title+(notes.length?` (${notes.join('، ')})`:''),subject,lessons:t.lessons,prepared:t.prepared,notPrepared:t.notPrepared,async:t.async,present:t.prepared>0};
  });
  const st=S.students;
  const sp=st.filter(s=>s.status==='حاضر').length, sa=st.length-sp;
  const tp=teachers.filter(t=>t.present).length;
  const computed={all:T.raw.all,prepared:T.raw.prepared,notPrepared:T.raw.all-T.raw.prepared,async:T.raw.async};
  const totals=T.totals||computed;
  const sorted=st.slice().sort((a,b)=>(GRADE_ORDER[a.grade]||9)-(GRADE_ORDER[b.grade]||9)||(a.section||0)-(b.section||0)||a.name.localeCompare(b.name,'ar'));
  return {teachers,students:sorted,
    tStats:{total:teachers.length,present:tp,absent:teachers.length-tp},
    sStats:{total:st.length,present:sp,absent:sa},
    lessons:{...totals,source:T.totals?'system':'computed',computed},
    dates:parseDates(T.dateText,S.dateText),
    notPreparedTeachers:teachers.filter(t=>!t.present).map(t=>t.name)};
}
const periodCount=s=>s.absent.length;

/* ---- docx filling ---- */
const WNS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fillDocx(JSZip,docxBuf,model,opts){
  const zip=await JSZip.loadAsync(docxBuf);
  const xml=await zip.file('word/document.xml').async('string');
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  const kids=(el,name)=>{const o=[];for(let n=el.firstChild;n;n=n.nextSibling) if(n.nodeType===1&&n.localName===name&&n.namespaceURI===WNS)o.push(n);return o};
  const all=(el,name)=>Array.from(el.getElementsByTagNameNS(WNS,name));
  const text=el=>all(el,'t').map(t=>t.textContent).join('');
  const mk=name=>doc.createElementNS(WNS,'w:'+name);
  const warn=[];
  function setText(tc,s,sz){
    let ps=kids(tc,'p'); let p=ps[0]; if(!p){p=mk('p');tc.appendChild(p)}
    ps.slice(1).forEach(x=>tc.removeChild(x));
    const firstR=all(p,'r')[0];
    let rPr=null; const pPr=kids(p,'pPr')[0];
    if(firstR&&kids(firstR,'rPr')[0]) rPr=kids(firstR,'rPr')[0].cloneNode(true);
    else if(pPr&&kids(pPr,'rPr')[0]) rPr=kids(pPr,'rPr')[0].cloneNode(true);
    for(let n=p.firstChild;n;){const nx=n.nextSibling; if(!(n.nodeType===1&&n.localName==='pPr')) p.removeChild(n); n=nx}
    const r=mk('r');
    if(rPr){ if(sz){ for(const tag of ['sz','szCs']){ let e=kids(rPr,tag)[0]; if(!e){e=mk(tag);rPr.appendChild(e)} e.setAttributeNS(WNS,'w:val',String(sz)) } } r.appendChild(rPr) }
    const t=mk('t'); t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve'); t.textContent=s; r.appendChild(t); p.appendChild(r);
  }
  const tables=all(doc,'tbl');
  const find=pred=>tables.find(t=>pred(text(t)));
  const cellAfter=(tbl,label)=>{ for(const tr of kids(tbl,'tr')){ const tcs=kids(tr,'tc'); for(let i=0;i<tcs.length-1;i++) if(text(tcs[i]).trim().startsWith(label)) return tcs[i+1] } return null };
  const d=Object.assign({},model.dates,opts.dates||{});
  // general info
  const gen=find(s=>s.includes('اليوم')&&s.includes('التاريخ'));
  let school='';
  if(gen){ const a=cellAfter(gen,'اليوم'); if(a&&d.day) setText(a,d.day); const b=cellAfter(gen,'التاريخ'); if(b&&d.hijriText) setText(b,d.hijriText);
    const sc=cellAfter(gen,'اسم المدرسة'); if(sc) school=text(sc).trim() } else warn.push('جدول البيانات العامة');
  // discipline
  const disc=find(s=>s.includes('نسبة الحضور')&&s.includes('الاجمالي'));
  if(disc){ const rows=kids(disc,'tr'); const put=(tr,v)=>{const tcs=kids(tr,'tc'); v.forEach((x,i)=>tcs[i+1]&&setText(tcs[i+1],x))};
    const T=model.tStats,S=model.sStats;
    const tr=rows.find(r=>/المعلم/.test(text(r))); const sr=rows.find(r=>/الطلب|الطلاب/.test(text(r)));
    if(tr) put(tr,[T.total,T.present,T.absent,pct(T.present,T.total),pct(T.absent,T.total)].map(String));
    if(sr) put(sr,[S.total,S.present,S.absent,pct(S.present,S.total),pct(S.absent,S.total)].map(String));
  } else warn.push('جدول الانضباط');
  // platform
  const plat=find(s=>s.includes('نسبة دخول'));
  if(plat){ const r=kids(plat,'tr')[1]; const tcs=r?kids(r,'tc'):[];
    const vals=[null,opts.supervisor||'',opts.principal||'',pct(model.tStats.present,model.tStats.total),pct(model.sStats.present,model.sStats.total),`${model.lessons.prepared} من ${model.lessons.all}`];
    vals.forEach((v,i)=>{ if(v!==null&&tcs[i]&&(v!==''||text(tcs[i]).trim()==='')) setText(tcs[i],v) });
  } else warn.push('جدول منصة مدرستي');
  // lessons table
  const les=find(s=>s.includes('اسم المعلم')&&s.includes('عنوان الدرس'));
  const markHeader=tr=>{ let trPr=kids(tr,'trPr')[0]; if(!trPr){trPr=mk('trPr'); const ex=kids(tr,'tblPrEx')[0]; tr.insertBefore(trPr, ex?ex.nextSibling:tr.firstChild)} if(!kids(trPr,'tblHeader').length) trPr.appendChild(mk('tblHeader')) };
  // Word often splits one visible line into many runs ("ليوم" | " " | "الأحد" | "2" | "1" | "/"...).
  // Work on each paragraph's full text, then write it back into its first run.
  const DAY_RE=/(ليوم|يوم)(\s*)(الأحد|الاحد|الاثنين|الإثنين|الثلاثاء|الأربعاء|الاربعاء|الخميس|الجمعة|السبت)/g;
  const DATE_RE=/\d{1,2}\s*\/\s*\d{1,2}\s*\/\s*1\d{3}\s*هـ*/g;
  const rewritePara=(p,fn)=>{ const ts=all(p,'t'); if(!ts.length) return false; const before=ts.map(t=>t.textContent).join(''); const after=fn(before);
    if(after===before) return false; ts[0].textContent=after; ts[0].setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve'); ts.slice(1).forEach(t=>t.textContent=''); return true };
  const dateFix=str=>{ let o=str; if(d.day) o=o.replace(DAY_RE,(m,a,sp)=>a+(sp||' ')+d.day); if(d.hijriText) o=o.replace(DATE_RE,d.hijriText); return o };
  const fixTitleCell=(tc,newFirst)=>{ for(const p of all(tc,'p')) rewritePara(p,str=>{ let o=dateFix(str);
      if(newFirst&&school&&!o.includes(school)) o=o.replace(/لمدرسة\s*$/,'لمدرسة '+school+' ');
      return o }) };
  if(les){
    const rows=kids(les,'tr'); const hdrIdx=rows.findIndex(r=>/اسم المعلم/.test(text(r)));
    const t1=rows[hdrIdx+1], t2=rows[hdrIdx+2]||rows[hdrIdx+1];
    const tmpl=[t1.cloneNode(true),t2.cloneNode(true)];
    for(let i=0;i<=hdrIdx;i++) markHeader(rows[i]);
    if(hdrIdx>0){ const tc=kids(rows[0],'tc')[0]; if(school&&!text(tc).includes(school)) fixTitleCell(tc,true); else fixTitleCell(tc,false) }
    rows.slice(hdrIdx+1).forEach(r=>les.removeChild(r));
    model.teachers.forEach((t,i)=>{ const tr=tmpl[i%2].cloneNode(true); les.appendChild(tr); const tcs=kids(tr,'tc');
      [t.name,t.cls,t.periods,t.subject,t.title].forEach((v,k)=>tcs[k]&&setText(tcs[k],v,22)) });
    // absent students table (clone of lessons table)
    if(opts.absentList&&opts.absentList.length){
      const ab=les.cloneNode(true); const arows=kids(ab,'tr');
      if(hdrIdx>0){ const tc=kids(arows[0],'tc')[0]; setText(tc,`كشف الطلاب الغائبين ${school?'– مدرسة '+school:''} ليوم ${d.day||''} الموافق ${d.hijriText||''}`) }
      const htcs=kids(arows[hdrIdx],'tc'); ['اسم الطالب','الصف والشعبة','عدد حصص الغياب','الحالة','الحصص الغائبة'].forEach((v,k)=>htcs[k]&&setText(htcs[k],v));
      arows.slice(hdrIdx+1).forEach(r=>ab.removeChild(r));
      opts.absentList.forEach((s,i)=>{ const tr=tmpl[i%2].cloneNode(true); ab.appendChild(tr); const tcs=kids(tr,'tc');
        const trh=all(tr,'trHeight')[0]; if(trh) trh.setAttributeNS(WNS,'w:val','340');
        const vals=s.groupTitle?[s.groupTitle,'','','','']:[s.name,`${s.grade} ث – ش${s.section}`,String(s.count??s.absent.length),s.status,s.absent.join('، ')];
        vals.forEach((v,k)=>tcs[k]&&setText(tcs[k],v,22)) });
      const br=mk('p'); const r=mk('r'); const b=mk('br'); b.setAttributeNS(WNS,'w:type','page'); r.appendChild(b); br.appendChild(r);
      les.parentNode.insertBefore(br,les.nextSibling); br.parentNode.insertBefore(ab,br.nextSibling);
    }
  } else warn.push('جدول الحصص الافتراضية');
  // final sweep: any remaining "ليوم <day>" or Hijri date anywhere in the body
  for(const p of all(doc,'p')){ const tx=text(p); if(/\d\s*\/\s*\d/.test(tx)||/يوم/.test(tx)) rewritePara(p,dateFix) }
  if(opts.support&&opts.support.length){
    const ps=all(doc,'p'); const i=ps.findIndex(p=>/الدعم الذي تحتاجه/.test(text(p)));
    if(i>=0){ const items=[]; for(let k=i+1;k<ps.length;k++){ const pPr=kids(ps[k],'pPr')[0]; if(pPr&&kids(pPr,'numPr').length) items.push(ps[k]); else if(items.length) break }
      if(items.length){ while(items.length<opts.support.length){ const c=items[items.length-1].cloneNode(true); items[items.length-1].parentNode.insertBefore(c,items[items.length-1].nextSibling); items.push(c) }
        items.forEach((p,k)=>{ const fake=mk('tc'); p.parentNode.insertBefore(fake,p); fake.appendChild(p); setText(fake,opts.support[k]||''); fake.parentNode.insertBefore(p,fake); fake.parentNode.removeChild(fake) }) } }
  }
  for(const hf of Object.keys(zip.files).filter(n=>/^word\/(header|footer)\d*\.xml$/.test(n))){
    const hx=new DOMParser().parseFromString(await zip.file(hf).async('string'),'application/xml'); let changed=false;
    for(const p of Array.from(hx.getElementsByTagNameNS(WNS,'p'))){ const ts=Array.from(p.getElementsByTagNameNS(WNS,'t')); if(!ts.length) continue; const before=ts.map(t=>t.textContent).join(''); const after=dateFix(before);
      if(after!==before){ ts[0].textContent=after; ts.slice(1).forEach(t=>t.textContent=''); changed=true } }
    if(changed) zip.file(hf,new XMLSerializer().serializeToString(hx));
  }
  zip.file('word/document.xml',new XMLSerializer().serializeToString(doc));
  const out=await zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  return {bytes:out,warn};
}

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

return {makeExtractor,parseTeachers,parseStudents,buildModel,fillDocx,guessSubject,pct,docxContent,docxKind,parseStudentsDocx,parseTeachersDocx,finishTeachers,formInfo,makeListDocx,parseDates};
})();
