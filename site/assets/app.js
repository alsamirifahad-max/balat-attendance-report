'use strict';
(function(){

document.documentElement.setAttribute('dir','rtl');document.documentElement.setAttribute('lang','ar');
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const GO={'أول':1,'ثاني':2,'ثالث':3};
const GN={'أول':'أول ثانوي','ثاني':'ثاني ثانوي','ثالث':'ثالث ثانوي'};
const files={form:null,teachers:null,students:null,wadih:null};
let model=null, selSection='', teacherAbsent=[], formMeta=null;


/* ---- security guards ---- */
const MAX_FILE_BYTES=25*1024*1024;            // 25 MB per uploaded file
const MAX_PDF_PAGES=400;
const MAX_UNZIPPED_BYTES=80*1024*1024;        // zip-bomb guard for .docx
function sniff(bytes){ if(bytes.length>4&&bytes[0]===0x25&&bytes[1]===0x50&&bytes[2]===0x44&&bytes[3]===0x46) return 'pdf'; if(bytes.length>4&&bytes[0]===0x50&&bytes[1]===0x4b&&bytes[2]===0x03&&bytes[3]===0x04) return 'zip'; return null }
async function guardZip(bytes){ const z=await JSZip.loadAsync(bytes.slice(0)); let total=0, n=0;
  z.forEach((p,f)=>{ n++; const u=f._data&&f._data.uncompressedSize; total+=(typeof u==='number'?u:0) });
  if(n>2000||total>MAX_UNZIPPED_BYTES) throw new Error('ملف Word غير صالح أو كبير جدًا.');
  if(!z.file('word/document.xml')) throw new Error('الملف ليس مستند Word صالحًا.'); }
// neutralise spreadsheet formula injection (=, +, -, @, tab, CR at cell start)
function csvCell(v){ let s=String(v??''); if(/^[=+\-@\t\r]/.test(s)) s="'"+s; return '"'+s.replace(/"/g,'""')+'"' }

function b64ToBytes(b){const s=atob(b);const u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return u}
fetch('assets/form.docx',{credentials:'omit',cache:'no-cache'}).then(r=>{ if(!r.ok) throw new Error(r.status); return r.arrayBuffer() })
  .then(b=>{ if(!files.form){ files.form={name:'الاستمارة المحفوظة (بلاط الشهداء)',bytes:new Uint8Array(b),builtin:true}; refreshSlots() } })
  .catch(()=>setStatus('تعذّر تحميل الاستمارة المحفوظة. ارفع الاستمارة بصيغة Word.',true));
function setStatus(msg,err){const s=$('status');s.textContent=msg;s.classList.toggle('err',!!err)}
function busy(msg){const b=$('busy'); if(msg){b.textContent=msg;b.hidden=false} else b.hidden=true}

/* ---------- date ---------- */
const DAYS=['الأحد','الاثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'];
const HM=['محرم','صفر','ربيع الأول','ربيع الآخر','جمادى الأولى','جمادى الآخرة','رجب','شعبان','رمضان','شوال','ذو القعدة','ذو الحجة'];
const GM=['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];
function hijriOf(iso){ try{ const [y,m,d]=iso.split('-').map(Number); const p=new Intl.DateTimeFormat('en-u-ca-islamic-umalqura',{day:'numeric',month:'numeric',year:'numeric',timeZone:'UTC'}).formatToParts(new Date(Date.UTC(y,m-1,d))); const g=k=>parseInt(p.find(x=>x.type===k).value); return {d:g('day'),m:g('month'),y:g('year')} }catch(e){ return null } }
function todayISO(){ try{ return new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Riyadh',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date()) }catch(e){ return new Date().toISOString().slice(0,10) } }
let hijriManual=false;
function parseHijriNum(s){ const m=(s||'').match(/(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})/); return m?{d:+m[1],m:+m[2],y:+m[3]}:null }
function curDate(){
  const iso=$('fGreg').value; let day='',gregLong='';
  if(iso){ const [y,m,d]=iso.split('-').map(Number); day=DAYS[new Date(Date.UTC(y,m-1,d)).getUTCDay()]; gregLong=`${d} ${GM[m-1]} ${y}م` }
  const h=parseHijriNum($('fHijri').value);
  return {day,hijriText:h?`${h.d}/${h.m}/${h.y}هـ`:'',hijriLong:h?`${h.d} ${HM[h.m-1]||''} ${h.y}هـ`:'',gregLong,iso};
}
function applyGreg(iso,hijriFromFile){
  if(iso) $('fGreg').value=iso;
  let h=hijriFromFile?parseHijriNum(hijriFromFile):null;
  if(!h&&!hijriManual) h=hijriOf($('fGreg').value);
  if(h) $('fHijri').value=`${h.d}/${h.m}/${h.y}هـ`;
  updateStamp();
}
function updateStamp(){
  const d=curDate(); const full=[d.day,d.hijriLong].filter(Boolean).join(' ');
  $('stampDate').textContent=full; $('hijriLong').textContent=full||'—';
  if($('dbDay')){ $('dbDay').textContent=d.day||'—'; $('dbHijri').textContent=d.hijriLong||'—'; $('dbGreg').textContent=d.gregLong||''; }
  if($('sDate')) $('sDate').textContent='تاريخ الكشف: '+[d.day,d.hijriLong].filter(Boolean).join(' · ');
}
applyGreg(todayISO());
(()=>{ const iso=todayISO(); const h=hijriOf(iso); $('copyYear').textContent=`— ${h?h.y+'هـ / ':''}${iso.slice(0,4)}م` })();
let dateTouched=false;
$('fGreg').addEventListener('input',()=>{ dateTouched=true; hijriManual=false; applyGreg() });
$('fHijri').addEventListener('input',()=>{ dateTouched=true; hijriManual=true; updateStamp() });

/* ---------- files ---------- */
function refreshSlots(){
  for(const k of ['form','teachers','students','wadih']){
    if(!$('slot-'+k)) continue;
    const f=files[k]; $('slot-'+k).classList.toggle('ok',!!f);
    $('name-'+k).textContent=f?(f.builtin?f.name+' — ارفع ملفًا آخر لاستبدالها':f.name):'لم يُرفع ملف';
    if(f&&!f.builtin){ const x=document.createElement('button'); x.type='button'; x.textContent='إزالة'; x.title='إزالة هذا الملف';
      x.style.cssText='margin-inline-start:8px;font:inherit;font-size:.8rem;padding:1px 8px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--red);cursor:pointer';
      x.addEventListener('click',e=>{ e.preventDefault(); e.stopPropagation(); if(k==='form'){ files.form=null; formMeta=null; fetch('assets/form.docx',{credentials:'omit',cache:'no-cache'}).then(r=>r.arrayBuffer()).then(b=>{ if(!files.form){ files.form={name:'الاستمارة المحفوظة (بلاط الشهداء)',bytes:new Uint8Array(b),builtin:true}; refreshSlots() } }).catch(()=>{}) } else delete files[k]; refreshSlots(); setStatus('أُزيل الملف. اضغط «إنشاء التقرير» لتحديث النتائج.') });
      $('name-'+k).appendChild(x) }
    $('tag-'+k).textContent=f?'جاهز':(k==='form'?'Word':k==='wadih'?'PDF واضح':'PDF / Word / Excel');
  }
  $('run').disabled=!files.form||wadihBusy;
}
let wadihBusy=false;
function wadihProg(msg){ const el=$('prog-wadih'); if(!el) return; const m=String(msg||'').match(/(\d+) من (\d+)/); el.hidden=!msg; if(m) el.firstChild.style.width=Math.round(+m[1]/+m[2]*100)+'%'; if(msg) $('name-wadih').textContent=msg }
async function readWadih(f,bytes){
  wadihBusy=true; refreshSlots(); $('slot-wadih').classList.add('ok'); wadihProg('بدء قراءة تقرير واضح …');
  try{ const r=await window.Wadih.parse(bytes,m=>{ wadihProg(m); setStatus(m+' (قد تستغرق القراءة دقيقة أو دقيقتين، ابقَ على الصفحة)') });
    if(!r.students.length) throw new Error('لم أجد جدول الطلاب في «'+f.name+'». تأكد أنه «تقرير غياب الطلاب عن الحصص» من واضح.');
    files.wadih={name:f.name,bytes:new Uint8Array(0),data:r};
    if(r.date&&r.date.greg&&!dateTouched) applyGreg(r.date.greg);
    const miss=r.students.filter(x=>!x.grade||!x.section).length;
    return `تقرير واضح: قُرئ ${r.students.length} طالبًا — ${r.counts[1]} غائب عن حصص وحاضر في غيرها، و${r.counts[2]} غائب اليوم كاملًا.`+(miss?` (${miss} لم يتضح فصلهم في الصورة، وتظهر في الكشف بدون فصل.)`:'');
  } finally { wadihBusy=false; wadihProg('') }
}
async function classify(file){
  if(file.size>MAX_FILE_BYTES) throw new Error('الملف «'+file.name+'» أكبر من 25 ميجابايت.');
  const bytes=new Uint8Array(await file.arrayBuffer());
  const kind0=sniff(bytes);
  if(!kind0) throw new Error('الملف «'+file.name+'» ليس PDF أو Word.');
  const isZip=kind0==='zip';
  if(isZip){ await guardZip(bytes);
    const c=await Core.docxContent(JSZip,bytes.slice(0)); const k=Core.docxKind(c);
    return {kind:k,bytes,docx:c};
  }
  const {extract}=Core.makeExtractor(PDFLib); const pages=await extract(bytes.slice(0));
  if(pages.length>MAX_PDF_PAGES) throw new Error('ملف PDF طويل جدًا.');
  const txt=pages.slice(0,1).flatMap(p=>p.runs.map(r=>r.str)).join('');
  if(/متابعة سير العملية|نسبة الحضور/.test(txt)) return {kind:'formpdf',bytes,pages};
  if(/تحضير|عنوان|الدرس/.test(txt)) return {kind:'teachers',bytes,pages};
  if(/الحالة|غياب|الطالب/.test(txt)) return {kind:'students',bytes,pages};
  return {kind:null,bytes,pages};
}
const isSheet=f=>/\.(xlsx|xls|xlsm|csv)$/i.test(f.name||'');
// Excel sheets -> the same {tables,paras} shape that Core builds from Word files
async function sheetContent(f){
  if(f.size>MAX_FILE_BYTES) throw new Error('الملف «'+f.name+'» أكبر من 25 ميجابايت.');
  const buf=await f.arrayBuffer(); let wb;
  if(/\.csv$/i.test(f.name)){ let t=new TextDecoder('utf-8').decode(buf); if(t.indexOf('\uFFFD')>=0){ try{ t=new TextDecoder('windows-1256').decode(buf) }catch(e){} } wb=XLSX.read(t.replace(/^\uFEFF/,''),{type:'string'}) }
  else wb=XLSX.read(buf,{type:'array'});
  const tables=[], paras=[];
  wb.SheetNames.forEach(sn=>{ const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:'',raw:false,blankrows:false}).map(r=>r.map(v=>String(v==null?'':v).trim()));
    rows.forEach(r=>{ const f1=r.filter(Boolean); if(f1.length===1) paras.push(f1[0]) }); tables.push(rows) });
  return {tables,paras};
}
let nlMsg='', wMsg='';
async function addFiles(list,hint){
  let err=''; nlMsg=''; wMsg='';
  for(const f of list){
    setStatus('جارٍ قراءة '+f.name+' …');
    // «واضح» report: a PDF made of page images; read on the device with OCR
    if(window.Wadih&&/\.pdf$/i.test(f.name||'')&&f.size<=MAX_FILE_BYTES){
      try{ const bytes=new Uint8Array(await f.arrayBuffer());
        if(sniff(bytes)==='pdf'&&(await window.Wadih.looksLike(bytes.slice(0))||(hint==='wadih'&&await window.Wadih.isImagePdf(bytes.slice(0))))){
          wMsg=await readWadih(f,bytes); continue } }
      catch(e){ console.error(e); err=(e&&/[\u0600-\u06FF]/.test(e.message||''))?e.message:'تعذّرت قراءة تقرير واضح «'+f.name+'».'; continue }
    }
    // Excel / CSV: a Naqar report (teachers or absence) is read like its Word version; anything else is the platform list
    if(isSheet(f)&&hint!=='nologin'&&window.XLSX){
      try{ const c=await sheetContent(f); const k=Core.docxKind(c);
        if(k==='teachers'||k==='students'){ files[k]={name:f.name,bytes:new Uint8Array(0),docx:c}; continue }
        if(k==='form'){ err='الاستمارة تُرفع بصيغة Word حتى يمكن تعبئتها، وحتى ذلك الحين تُستخدم الاستمارة المحفوظة.'; continue }
      }catch(e){ console.error(e) }
    }
    if(isSheet(f)||hint==='nologin'){
      if(!window.NoLogin){ err='تعذّر تحميل قارئ ملف الطلاب الغائبين عن المنصة. أعد تحميل الصفحة.'; continue }
      try{ const inf=await window.NoLogin.loadFile(f);
        nlMsg=(inf.full&&inf.total!==inf.missing)?`ملف المنصة: في الملف ${inf.total} طالبًا، لم يدخل منهم ${inf.missing}${inf.refTxt?' (آخر دخول '+inf.refTxt+')':''}. تجدهم في قسم «الطلاب الذين لم يدخلوا المنصة» أسفل الصفحة.`:`ملف المنصة: ${inf.total} طالبًا لم يدخلوا المنصة، وتجدهم في قسمهم أسفل الصفحة.`;
      }catch(e){ console.error(e); err=(e&&e.message)||('تعذّرت قراءة «'+f.name+'».') }
      continue;
    }
    try{
      const r=await classify(f);
      // the platform list (students who did not log in) can come as PDF or Word too, in any slot
      if(window.NoLogin){ const inf=r.pages?await window.NoLogin.tryPages(r.pages,f.name):(r.docx?await window.NoLogin.tryDocx(r.bytes.slice(0),f.name):null);
        if(inf){ nlMsg=(inf.full&&inf.total!==inf.missing)?`ملف المنصة: في الملف ${inf.total} طالبًا، لم يدخل منهم ${inf.missing}${inf.refTxt?' (آخر دخول '+inf.refTxt+')':''}. تجدهم في قسم «الطلاب الذين لم يدخلوا المنصة» أسفل الصفحة.`:`ملف المنصة: ${inf.total} طالبًا لم يدخلوا المنصة، وتجدهم في قسمهم أسفل الصفحة.`; continue } }
      if(r.kind==='formpdf'){ err='الاستمارة بصيغة PDF لا يمكن تعبئتها. ارفعها بصيغة Word، وحتى ذلك الحين تُستخدم الاستمارة المحفوظة.'; continue }
      const k=r.kind||(hint!=='form'?hint:null);
      if(!k){ err='لم أتعرف على الملف «'+f.name+'». تأكد أنه تقرير نقار أو استمارة المتابعة.'; continue }
      files[k]={name:f.name,bytes:r.bytes,pages:r.pages,docx:r.docx};
      if(k==='form') formMeta=null;
    }catch(e){ console.error(e); err=(e&&/[\u0600-\u06FF]/.test(e.message||''))?e.message:'تعذّرت قراءة «'+f.name+'». تأكد أن الملف PDF أو Word سليم.' }
  }
  refreshSlots();
  const info=[wMsg,nlMsg].filter(Boolean).join(' ');
  if(err) setStatus(err+(info?' '+info:''),true); else if(info&&$('run').disabled) setStatus(info); else setStatus((info?info+' ':'')+'اضغط «إنشاء التقرير» متى شئت، والملفات كلها اختيارية.');
  if(model) updateDlNote();
}
for(const k of ['form','teachers','students','wadih']){
  const slot=$('slot-'+k), inp=$('in-'+k); if(!slot) continue;
  inp.addEventListener('change',()=>{addFiles([...inp.files],k);inp.value=''});
  slot.addEventListener('dragover',e=>{e.preventDefault();slot.classList.add('drag')});
  slot.addEventListener('dragleave',()=>slot.classList.remove('drag'));
  slot.addEventListener('drop',e=>{e.preventDefault();slot.classList.remove('drag');addFiles([...e.dataTransfer.files],k)});
}
if($('slot-nologin')){
  const slot=$('slot-nologin'), inp=$('in-nologin');
  inp.addEventListener('change',()=>{addFiles([...inp.files],'nologin');inp.value=''});
  slot.addEventListener('dragover',e=>{e.preventDefault();slot.classList.add('drag')});
  slot.addEventListener('dragleave',()=>slot.classList.remove('drag'));
  slot.addEventListener('drop',e=>{e.preventDefault();slot.classList.remove('drag');addFiles([...e.dataTransfer.files],'nologin')});
  document.addEventListener('nologin:cleared',()=>{ slot.classList.remove('ok'); $('name-nologin').textContent='لم يُرفع ملف'; $('tag-nologin').textContent='Excel'; setStatus('أُزيل ملف الطلاب الغائبين عن المنصة.'); if(model) updateDlNote() });
  document.addEventListener('nologin:data',e=>{ const d=e.detail||{}; slot.classList.add('ok');
    $('name-nologin').textContent=d.name+((d.full&&d.total!==d.missing)?` — لم يدخل ${d.missing} من ${d.total}`:` — ${d.missing} طالبًا لم يدخلوا`); $('tag-nologin').textContent='جاهز';
    const x=document.createElement('button'); x.type='button'; x.textContent='إزالة'; x.style.cssText='margin-inline-start:8px;font:inherit;font-size:.8rem;padding:1px 8px;border:1px solid var(--line);border-radius:6px;background:var(--card);color:var(--red);cursor:pointer';
    x.addEventListener('click',ev=>{ ev.preventDefault(); ev.stopPropagation(); if(window.NoLogin&&window.NoLogin.clear) window.NoLogin.clear() }); $('name-nologin').appendChild(x);
    if(model) updateDlNote() });
}
refreshSlots();

$('run').addEventListener('click',async()=>{
  $('run').disabled=true;
  try{
    const {extract}=Core.makeExtractor(PDFLib);
    setStatus('قراءة تحضير المعلمين …');
    // every file is optional: a missing report simply contributes nothing
    let T={teachers:[],raw:{all:0,prepared:0,async:0},dateText:''}; const tf=files.teachers;
    if(tf){
      if(tf.docx) T=Core.finishTeachers(Core.parseTeachersDocx(tf.docx));
      else T=Core.parseTeachers(tf.pages||await extract(tf.bytes.slice(0)));
      if(!T.teachers.length) throw new Error('لم أجد جدول المعلمين في ملف التحضير. تأكد أنه «تقرير تحضير دروس المعلمين» من نقار، أو احذفه وأنشئ التقرير بدونه.');
    }
    setStatus('قراءة غياب الطلاب …');
    let S={students:[],dateText:''}; const sf=files.students;
    if(sf){
      if(sf.docx) S=Core.parseStudentsDocx(sf.docx); else S=Core.parseStudents(sf.pages||await extract(sf.bytes.slice(0)));
      if(!S.students.length) throw new Error('لم أجد قائمة الطلاب في ملف الغياب. تأكد أنه «تقرير غياب الحصص للطلاب» من نقار، أو احذفه وأنشئ التقرير بدونه.');
    }
    const wf=files.wadih;
    if(wf&&wf.data){ const norm=n=>String(n||'').replace(/[إأآ]/g,'ا').replace(/ة/g,'ه').replace(/ى/g,'ي').replace(/\s+/g,'');
      const seen=new Set(S.students.map(x=>norm(x.name)));
      const ws=wf.data.students.filter(x=>!seen.has(norm(x.name))).map(x=>({no:x.no,name:x.name,classFull:x.grade?`${x.grade} ش${x.section}`:'',grade:x.grade,section:x.section?+x.section:null,present:[],absent:x.absent.slice(),status:x.status,src:'wadih'}));
      S={students:S.students.concat(ws),dateText:S.dateText} }
    model=Core.buildModel(T,S);
    if(wf&&wf.data&&wf.data.date&&wf.data.date.greg&&!model.dates.gregText){ const g=wf.data.date.greg; model.dates.gregText=g; const h=hijriOf(g); if(h) model.dates.hijriText=`${h.d}/${h.m}/${h.y}هـ` }
    { const mem=subjMem(); model.teachers.forEach(t=>{ if(!t.subject&&mem[t.name]) t.subject=mem[t.name] }) }
    teacherAbsent=model.teachers.map(t=>!t.present);
    let dateNote='';
    if(!dateTouched){ hijriManual=false; if(model.dates.gregText||model.dates.hijriText) applyGreg(model.dates.gregText||'',model.dates.hijriText) }
    else if(model.dates.hijriText&&model.dates.hijriText!==curDate().hijriText){ dateNote=` تنبيه: أبقيت التاريخ الذي اخترته (${curDate().day} ${curDate().hijriText})، بينما تاريخ ملفات نقار ${model.dates.day||''} ${model.dates.hijriText}.` }
    render();
    const parts=[]; if(tf) parts.push(`${model.tStats.total} معلمًا`); if(sf||wf) parts.push(`${model.sStats.total} طالبًا${wf?' (من '+(sf?'نقار وواضح':'واضح')+')':''}`); if(window.NoLogin&&window.NoLogin.hasData()) parts.push('ملف الطلاب الغائبين عن المنصة');
    setStatus((parts.length?'تم إنشاء التقرير من: '+parts.join('، '):'تم إنشاء الاستمارة بدون ملفات نقار، ويمكنك تعبئة بياناتها وتنزيلها')+'.'+dateNote);
  }catch(e){console.error(e);setStatus(e.message||'تعذّرت قراءة الملفات.',true)}
  $('run').disabled=false;
});

/* ---------- derived data ---------- */
function tStats(){ const total=model.teachers.length, absent=teacherAbsent.filter(Boolean).length; return {total,present:total-absent,absent} }
/* subjects typed manually are remembered in this browser only, per teacher name */
const SUBJ_KEY='balat.subjects.v1';
function subjMem(){ try{ const m=JSON.parse(localStorage.getItem(SUBJ_KEY)||'{}'); return (m&&typeof m==='object')?m:{} }catch(e){ return {} } }
function subjSave(name,v){ try{ const m=subjMem(); if(v) m[name]=v.slice(0,80); else delete m[name]; localStorage.setItem(SUBJ_KEY,JSON.stringify(m)) }catch(e){} }
function cnt(s){ return s.absent.length }
function bucketOf(n){ return n>=7?7:n }
const withStatus=()=>$('fPartial').checked;
const withPer=()=>!$('fPer')||$('fPer').checked;
const stFilter=()=>withStatus()?'':'غائب';
function chosenBuckets(){ return [1,2,3,4,5,6,7].filter(i=>$('b'+i).checked) }
function studentList(opts={}){
  const bs=chosenBuckets(), g=$('fGrade').value, st=stFilter(), q=opts.ignoreSearch?'':$('fSearch').value.trim();
  return model.students.filter(s=>{ const n=cnt(s); if(n<1||!bs.includes(bucketOf(n))) return false;
    if(g&&s.grade!==g) return false; if(st&&s.status!==st) return false;
    if(!opts.ignoreSection&&selSection&&s.grade+'|'+s.section!==selSection) return false;
    if(q&&!s.name.includes(q)) return false; return true });
}
function studentGroups(list){
  if(!$('fMerge').checked) return [{title:null,rows:list}];
  const out=[]; for(const g of ['أول','ثاني','ثالث']){ const rows=list.filter(s=>s.grade===g); if(rows.length) out.push({title:GN[g],grade:g,rows}) }
  const other=list.filter(s=>!GO[s.grade]); if(other.length) out.push({title:'أخرى',rows:other});
  return out;
}

/* ---------- render ---------- */
function render(){
  $('emptyState').hidden=true; $('results').hidden=false;
  renderFigures(); renderTeachers(); renderStudents(); renderAbsentTeachers(); renderWarnings(); updateDlNote(); updateSrcNote(); renderCut();
}
/* ---- student statistics source: Naqar absence report or the Madrasati platform list (user's choice) ---- */
function studentSrc(){
  const r=document.querySelector('input[name="sSrc"]:checked'); const v=r?r.value:'auto';
  const hasN=!!(files.students||files.wadih)&&model&&model.sStats.total>0, hasP=!!(window.NoLogin&&window.NoLogin.hasData());
  if(v==='naqar') return hasN?'naqar':(hasP?'platform':'naqar');
  if(v==='platform') return hasP?'platform':'naqar';
  return hasN?'naqar':(hasP?'platform':'naqar');
}
function platformStats(){
  const p=window.NoLogin&&window.NoLogin.stats&&window.NoLogin.stats(); if(!p) return null;
  let total=p.total;
  if(!p.full){ const t=parseInt(($('sTotal')&&$('sTotal').value)||'',10); total=t>0?t:(model&&model.sStats.total>0?model.sStats.total:0) }
  if(!total) return null;
  const absent=Math.min(p.absent,total); return {total,present:total-absent,absent};
}
function schoolTotal(){ const p=window.NoLogin&&window.NoLogin.stats&&window.NoLogin.stats(); if(p&&p.full&&p.total) return p.total; const t=parseInt(($('sTotal')&&$('sTotal').value)||'',10); return t>0?t:0 }
// «واضح» lists only students who missed periods: with the school total known, everyone else attended
function wadihStats(){ if(!files.wadih||files.students||!model) return null; const t=schoolTotal(); if(!t||t<model.sStats.total) return null; const a=model.sStats.absent; return {total:t,present:t-a,absent:a} }
function sStats(){ if(studentSrc()==='platform'){ const p=platformStats(); if(p) return p } return wadihStats()||model.sStats }
function updateSrcNote(){
  const el=$('sSrcNote'); if(!el) return;
  const src=studentSrc(), p=window.NoLogin&&window.NoLogin.stats&&window.NoLogin.stats();
  if($('sTotalWrap')) $('sTotalWrap').hidden=!((p&&!p.full)||(files.wadih&&!(p&&p.full)));
  if(!model){ el.textContent=''; return }
  const S=sStats();
  if(src==='platform'&&!platformStats()) el.textContent='اكتب إجمالي طلاب المدرسة ليُحسب الحضور من ملف المنصة، وحتى ذلك الحين تُستخدم أرقام نقار.';
  else if(src!=='platform'&&files.wadih&&!files.students&&!wadihStats()) el.textContent=`المصدر الحالي: تقرير واضح — الإجمالي ${S.total} (المذكورون في التقرير فقط)، الحاضرون ${S.present}، الغائبون ${S.absent}. اكتب إجمالي طلاب المدرسة أو ارفع ملف المنصة الكامل ليُحسب الحضور من كل طلاب المدرسة.`;
  else el.textContent=(src==='platform'?'المصدر الحالي: ملف منصة مدرستي':'المصدر الحالي: تقرير غياب الحصص من '+(files.students&&files.wadih?'نقار وواضح':files.wadih?'واضح':'نقار'))+` — الإجمالي ${S.total}، الحاضرون ${S.present}، الغائبون ${S.absent}.`;
}
function statData(){
  const T=tStats(),S=sStats(),L=model.lessons; const sync=Math.max(0,L.prepared-L.async);
  return [
    {key:'t',title:'المعلمون',main:'الحاضرون',v:T.present,of:T.total,y:['حضور',T.present],n:['غياب',T.absent],c:'var(--teal)',pc:'#1d6f68'},
    {key:'s',title:'الطلاب',main:'الحاضرون',v:S.present,of:S.total,y:['حضور',S.present],n:['غياب',S.absent],c:'var(--blue)',pc:'#2f6fb0'},
    {key:'l',title:'الحصص',main:'المعدّة',v:L.prepared,of:L.all,y:['معدّة',L.prepared],n:['غير معدّة',L.notPrepared],c:'var(--green)',pc:'#26794a'},
    {key:'a',title:'الحصص المعدّة',main:'المتزامنة',v:sync,of:L.prepared,y:['متزامنة',sync],n:['غير متزامنة',L.async],c:'var(--amber)',pc:'#a86412',amb:true},
  ];
}
function renderFigures(){
  const L=model.lessons; const match=L.source==='system'&&L.computed.all===L.all&&L.computed.prepared===L.prepared;
  $('figures').innerHTML=statData().map(d=>{ const py=d.of?d.y[1]/d.of*100:0;
    return `<div class="st" style="--c:${d.c}"><div class="k"><span>${d.title} ${d.main}</span>${d.key==='l'?(match?'<span class="chip ok">مطابق لنقار</span>':(L.source==='system'?'<span class="chip no">يختلف عن نقار</span>':'')):''}</div>
    <div class="v">${d.v} <small>من ${d.of}</small></div>
    <div class="duo" aria-hidden="true"><i style="width:${py}%;background:var(--green)"></i><i style="width:${100-py}%;background:var(${d.amb?'--amber':'--red'})"></i></div>
    <div class="pr"><div class="y">نسبة ${d.y[0]}<b><bdi>${Core.pct(d.y[1],d.of)}</bdi></b></div><div class="n ${d.amb?'amb':''}">نسبة ${d.n[0]}<b><bdi>${Core.pct(d.n[1],d.of)}</bdi></b></div></div>
    <div class="cnt"><span>${d.y[0]}: ${d.y[1]}</span><span>${d.n[0]}: ${d.n[1]}</span></div></div>` }).join('');
  updateStamp();
}
function renderTeachers(){
  $('n-t').textContent=model.teachers.length;
  $('tBody').innerHTML=model.teachers.map((t,i)=>{
    const st=!t.present?'<span class="pill a">لم يُعدّ</span>':(t.notPrepared?`<span class="pill p">${t.notPrepared} غير معدّة</span>`:'');
    return `<tr class="${!t.present?'flag':(t.notPrepared?'half':'')}"><td class="num">${i+1}</td><td>${esc(t.name)}</td><td>${esc(t.cls)}</td><td class="num">${esc(t.periods)}</td>
    <td><input type="text" id="subj-${i}" data-i="${i}" value="${esc(t.subject)}" class="${t.subject?'':'empty'}" aria-label="مادة ${esc(t.name)}"></td><td>${esc(t.title)}</td><td>${st}</td></tr>`}).join('');
  $('tBody').querySelectorAll('input').forEach(inp=>inp.addEventListener('input',()=>{model.teachers[+inp.dataset.i].subject=inp.value.trim();subjSave(model.teachers[+inp.dataset.i].name,inp.value.trim());inp.classList.toggle('empty',!inp.value.trim());renderWarnings();renderAbsentTeachers()}));
}
function renderStudents(){
  // bucket counts (respect grade/status filters)
  const g=$('fGrade').value, st=stFilter(); const bc={1:0,2:0,3:0,4:0,5:0,6:0,7:0};
  model.students.forEach(s=>{const n=cnt(s); if(n<1) return; if(g&&s.grade!==g) return; if(st&&s.status!==st) return; bc[bucketOf(n)]++});
  const minMode=isMin(); let acc=0; const cum={}; for(let i=7;i>=1;i--){ acc+=bc[i]; cum[i]=acc }
  for(let i=1;i<=7;i++){ const n=minMode?cum[i]:bc[i]; $('c'+i).textContent=n; $('b'+i).closest('.bk').style.opacity=n?'':'.55';
    $('lb'+i).textContent=minMode?BK_MIN[i]:BK_LABEL[i] }
  const lo=chosenBuckets()[0];
  $('modeHint').textContent=minMode?(lo?`يُعرض من غاب ${BK_TXT[lo]} أو أكثر`:'اختر حدًّا أدنى'):'يُعرض من غاب العدد المختار فقط';
  $('cAll').textContent=Object.values(bc).reduce((a,b)=>a+b,0);
  const base=studentList({ignoreSection:true,ignoreSearch:true});
  const counts={}; base.forEach(s=>{const k=s.grade+'|'+s.section; counts[k]=(counts[k]||0)+1});
  const keys=Object.keys(counts).sort((a,b)=>{const[ga,sa]=a.split('|'),[gb,sb]=b.split('|');return (GO[ga]||9)-(GO[gb]||9)||sa-sb});
  if(selSection&&!counts[selSection]) selSection='';
  $('secChips').innerHTML=`<button aria-pressed="${!selSection}" data-k="">الكل · <b>${base.length}</b></button>`+keys.map(k=>{const[gr,se]=k.split('|');return `<button aria-pressed="${selSection===k}" data-k="${k}"><bdi>${gr} ش${se}</bdi> · <b>${counts[k]}</b></button>`}).join('');
  $('secChips').querySelectorAll('button').forEach(b=>b.addEventListener('click',()=>{selSection=b.dataset.k;renderStudents();updateDlNote()}));
  const list=studentList(); const groups=studentGroups(list);
  $('n-s').textContent=list.length; $('sCount').textContent=`${list.length} طالبًا في الكشف الحالي`+($('fMerge').checked?` · ${groups.map(g=>g.title+': '+g.rows.length).join(' · ')}`:'');
  let html='';
  for(const gr of groups){ if(gr.title) html+=`<tr class="grp"><td colspan="${5+(withPer()?1:0)+(withStatus()?1:0)}">${esc(gr.title)} — ${gr.rows.length} طالبًا</td></tr>`;
    html+=gr.rows.map((s,i)=>`<tr><td class="num">${i+1}</td><td>${esc(s.name)}</td><td>${esc(GN[s.grade]||s.grade)}</td><td class="num"><bdi>ش ${s.section??''}</bdi></td><td class="num"><b>${cnt(s)}</b></td>${withPer()?`<td>${esc(s.absent.join('، '))}</td>`:''}${withStatus()?`<td>${s.status==='غائب'?'<span class="pill a">غائب</span>':'<span class="pill p">حاضر</span>'}</td>`:''}</tr>`).join('') }
  $('sBody').innerHTML=html||`<tr><td colspan="7" class="note">لا يوجد طلاب بهذه التصفية.</td></tr>`;
  $('thStatus').hidden=!withStatus(); $('thPer').hidden=!withPer();
  const hidden=model.students.filter(x=>x.status==='حاضر'&&cnt(x)>0&&(!g||x.grade===g)).length;
  $('partialHint').textContent=withStatus()?'':(hidden?`(${hidden} طالبًا حاضرًا فاتتهم حصص مخفيون)`:'');
}
function renderAbsentTeachers(){
  $('n-a').textContent=teacherAbsent.filter(Boolean).length;
  $('aBody').innerHTML=model.teachers.map((t,i)=>`<tr class="${teacherAbsent[i]?'flag':''}"><td><label class="check" for="ta-${i}"><input type="checkbox" id="ta-${i}" data-i="${i}" ${teacherAbsent[i]?'checked':''}> <span class="${teacherAbsent[i]?'bad':'note'}">${teacherAbsent[i]?'غائب':'حاضر'}</span></label></td><td>${esc(t.name)}</td><td>${esc(t.cls)}</td><td class="num">${esc(t.periods)}</td><td class="num">${t.prepared}</td><td class="num ${t.notPrepared?'bad':''}">${t.notPrepared}</td><td>${esc(t.subject)}</td></tr>`).join('');
  $('aBody').querySelectorAll('input').forEach(inp=>inp.addEventListener('change',()=>{teacherAbsent[+inp.dataset.i]=inp.checked;renderAbsentTeachers();renderFigures();renderWarnings()}));
}
function alertRows(){
  const rows=[];
  model.teachers.forEach((t,i)=>{ if(teacherAbsent[i]) rows.push(['معلم غائب',t.name,t.present?'حُدّد غائبًا يدويًا':'لم يُعدّ أي درس — '+t.cls]) });
  model.teachers.forEach((t,i)=>{ if(!teacherAbsent[i]&&t.notPrepared) rows.push(['حصة غير معدّة',t.name,t.lessons.filter(b=>b.notPrepared).map(b=>`${b.grade} ش${b.section} (الحصة ${b.period})`).join('، ')]) });
  model.teachers.forEach(t=>{ if(!t.subject) rows.push(['مادة غير محددة',t.name,'«'+t.title+'»']) });
  const L=model.lessons; if(L.source==='system'&&(L.computed.all!==L.all||L.computed.prepared!==L.prepared)) rows.push(['اختلاف في الإجمالي','—',`المقروء ${L.computed.all} (المعدّة ${L.computed.prepared}) مقابل نقار ${L.all} (المعدّة ${L.prepared})`]);
  return rows;
}
function renderWarnings(){
  const rows=alertRows(); $('n-w').textContent=rows.length;
  const by={}; rows.forEach(r=>{(by[r[0]]=by[r[0]]||[]).push(r)});
  $('warnBox').innerHTML=(Object.entries(by).map(([k,v])=>`<div style="display:grid;gap:6px"><h3>${esc(k)} (${v.length})</h3><ul class="list">${v.map(r=>`<li>${esc(r[1])} — ${esc(r[2])}</li>`).join('')}</ul></div>`).join('')||'<p class="note">لا توجد تنبيهات.</p>');
}
function updateDlNote(){
  const n=studentList({ignoreSearch:true}).length;
  let t=$('addAbsent').checked?`يُضاف في آخرها كشف بـ ${n} طالبًا حسب تصفية تبويب الطلاب الغائبين.`:'بدون كشف الغائبين.';
  if($('addNoLogin')){ const NL=window.NoLogin, m=NL&&NL.hasData()?NL.count():0;
    t+=$('addNoLogin').checked?(NL&&NL.hasData()?` ويُضاف كشف بـ ${m} طالبًا لم يدخلوا المنصة (${NL.summary()}).`:' (لم يُرفع ملف الطلاب الغائبين عن المنصة بعد.)'):' بدون كشف من لم يدخلوا المنصة.'; }
  $('dlNote').textContent=t;
}
const BK_LABEL=['','غائب حصة','غائب حصتين','غائب ٣ حصص','غائب ٤ حصص','غائب ٥ حصص','غائب ٦ حصص','غائب ٧ حصص'];
const BK_MIN=['','غائب حصة فأكثر','غائب حصتين فأكثر','غائب ٣ حصص فأكثر','غائب ٤ حصص فأكثر','غائب ٥ حصص فأكثر','غائب ٦ حصص فأكثر','غائب ٧ حصص فأكثر'];
const BK_TXT=['','حصة','حصتين','٣ حصص','٤ حصص','٥ حصص','٦ حصص','٧ حصص'];
const isMin=()=>$('mMin').checked;
for(let i=1;i<=7;i++){ const l=document.createElement('label'); l.className='bk'; l.htmlFor='b'+i; if(i===7) l.title='يشمل من غاب 7 حصص أو أكثر';
  l.innerHTML=`<input type="checkbox" id="b${i}" value="${i}" checked> <span id="lb${i}">${BK_MIN[i]}</span> <b id="c${i}"></b>`; $('buckets').appendChild(l); }
// "فأكثر" mode: picking a number selects it and every higher number (a minimum threshold)
let minLo=1; const applyMin=()=>{ for(let j=1;j<=7;j++) $('b'+j).checked=j>=minLo };
// click a number = it becomes the minimum; click the current minimum again = raise it by one
for(let i=1;i<=7;i++) $('b'+i).addEventListener('input',()=>{ if(!isMin()) return; minLo=(minLo===i)?i+1:i; applyMin() },true);
$('bAll').addEventListener('input',()=>{ if(isMin()){ minLo=$('bAll').checked?1:8; applyMin() } },true);
for(const r of ['mMin','mExact']) $(r).addEventListener('change',()=>{ if(isMin()){ minLo=chosenBuckets()[0]||1; applyMin() } $('bAll').checked=[1,2,3,4,5,6,7].every(i=>$('b'+i).checked); renderStudents(); updateDlNote() });
$('bAll').addEventListener('input',()=>{ for(let i=1;i<=7;i++) $('b'+i).checked=$('bAll').checked; renderStudents(); updateDlNote() });
['fGrade','fPartial','fSearch','fMerge','fPer','b1','b2','b3','b4','b5','b6','b7'].forEach(id=>$(id).addEventListener('input',()=>{ $('bAll').checked=[1,2,3,4,5,6,7].every(i=>$('b'+i).checked); renderStudents();updateDlNote()}));
$('addAbsent').addEventListener('change',updateDlNote);
if($('addNoLogin')) $('addNoLogin').addEventListener('change',updateDlNote);
/* remember the user's choices on this device */
const PREF=(()=>{ try{ return JSON.parse(localStorage.getItem('balat-prefs')||'{}') }catch(e){ return {} } })();
function savePref(k,v){ PREF[k]=v; try{ localStorage.setItem('balat-prefs',JSON.stringify(PREF)) }catch(e){} }
['addAbsent','addNoLogin','fPer'].forEach(id=>{ if(!$(id)) return; if(typeof PREF[id]==='boolean') $(id).checked=PREF[id]; $(id).addEventListener('change',()=>savePref(id,$(id).checked)) });
if(PREF.sSrc&&$('sSrc'+{auto:'Auto',naqar:'Naqar',platform:'Plat'}[PREF.sSrc])) $('sSrc'+{auto:'Auto',naqar:'Naqar',platform:'Plat'}[PREF.sSrc]).checked=true;
document.querySelectorAll('input[name="sSrc"]').forEach(r=>r.addEventListener('change',()=>savePref('sSrc',r.value)));
document.querySelectorAll('input[name="sSrc"]').forEach(r=>r.addEventListener('change',()=>{ if(model){ renderFigures(); updateSrcNote() } }));
if($('sTotal')){ try{ const v=localStorage.getItem('balat-sTotal'); if(v) $('sTotal').value=v }catch(e){}
  $('sTotal').addEventListener('input',()=>{ try{ localStorage.setItem('balat-sTotal',$('sTotal').value) }catch(e){} if(model){ renderFigures(); updateSrcNote() } }) }
document.addEventListener('nologin:data',e=>{ const d=e.detail||{}; if(d.full&&d.total!==d.missing&&$('sTotal')){ $('sTotal').value=d.total; try{ localStorage.setItem('balat-sTotal',String(d.total)) }catch(_){} }
  updateSrcNote(); if(model) renderFigures() });
document.addEventListener('nologin:cleared',()=>{ updateSrcNote(); if(model) renderFigures() });
document.addEventListener('nologin:change',()=>{ if(model) updateDlNote() });
const TABS=[['tab-t','pane-t'],['tab-s','pane-s'],['tab-a','pane-a'],['tab-c','pane-c'],['tab-w','pane-w']];
for(const [id] of TABS) $(id).addEventListener('click',()=>{for(const [i2,p2] of TABS){$(i2).setAttribute('aria-selected',i2===id);$(p2).hidden=i2!==id}});

/* ---------- export bars ---------- */
const ICON={
  docx:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M8.5 12.5l1.2 4.5 1.3-3.5 1.3 3.5 1.2-4.5"/></svg>',
  pdf:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M9 17v-4h1.3a1.2 1.2 0 0 1 0 2.4H9"/><path d="M14 13v4"/><path d="M14 13h1.6"/><path d="M14 15h1.2"/></svg>',
  print:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 9V3h10v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/></svg>',
  csv:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 4v16M15 4v16"/></svg>'
};
const NAMES={cut:'الطلاب المنقطعون تمامًا',stats:'إحصائيات_اليوم',form:'الاستمارة',lessons:'الحصص الافتراضية',students:'الطلاب الغائبين',teachers:'غياب المعلمين',alerts:'التنبيهات'};
document.querySelectorAll('.exp').forEach(el=>{
  const k=el.dataset.kind;
  el.innerHTML=`<span class="lab">${k==='form'?'تنزيل الاستمارة:':'تنزيل:'}</span>`+
    `<button class="word" data-fmt="docx">${ICON.docx}Word</button><button data-fmt="pdf">${ICON.pdf}PDF</button><button data-fmt="print">${ICON.print}طباعة</button>`+(k==='students'?`<button data-fmt="csv">${ICON.csv}Excel</button>`:'')+(k==='cut'?`<button data-fmt="xlsx">${ICON.csv}Excel</button>`:'');
  el.querySelectorAll('button').forEach(b=>b.addEventListener('click',()=>doExport(k,b.dataset.fmt,b)));
});

let downloadsP=null;
function getDownloads(){ return Promise.resolve(null) }
async function offer(filename,data,mime){
  const dl=await getDownloads();
  if(dl){
    try{ await dl.save({filename,data}); setStatus('تم حفظ '+filename+'.') }
    catch(e){ if(e&&e.code==='declined') setStatus('أُلغي التنزيل.'); else setStatus('تعذّر التنزيل: '+(e&&(e.message||e.code)||''),true) }
    return;
  }
  const url=URL.createObjectURL(data instanceof Blob?data:new Blob([data],{type:mime})); const a=document.createElement('a'); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),5000);
}
function fname(base,ext){ return safeName(fname0(base,ext)) }
function safeName(s){ return s.replace(/[\\/:*?"<>|\u0000-\u001f]/g,'_').slice(0,150) }
function fname0(base,ext){ const d=curDate(); return `${base}_${(d.hijriText||'').replace(/\//g,'-').replace(/هـ/g,'')}.${ext}`.replace(/_\./,'.').replace(/\s+/g,'_') }

function exportModel(){ return Object.assign({},model,{tStats:tStats(),sStats:sStats(),dates:Object.assign({},model.dates,curDate())}) }
function formAbsentRows(){
  const list=studentList({ignoreSearch:true}); const groups=studentGroups(list); const out=[];
  groups.forEach(g=>{ if(g.title) out.push({groupTitle:`${g.title} — ${g.rows.length} طالبًا`}); g.rows.forEach(s=>out.push(Object.assign({},s,{count:cnt(s)}))) });
  return out;
}
const SUPPORT=()=>$('fSupport').value.split('\n').map(s=>s.trim()).filter(Boolean);

function listSpec(kind){
  const d=curDate(); const when=`ليوم ${d.day} الموافق ${d.hijriText}`;
  const school=(formMeta&&formMeta.school)?' – مدرسة '+formMeta.school:'';
  if(kind==='lessons') return {orientation:'land',title:`الحصص الافتراضية${school} ${when}`,headers:['م','اسم المعلم/ة','الصف والشعب','الحصص','المادة','عنوان الدرس'],widths:[.45,2.1,2.1,1,1.6,3.6],
    groups:[{title:`استمارة متابعة سير العملية التعليمية (حضور حصص افتراضية)${school} ${when}`,rows:model.teachers.map((t,i)=>[i+1,t.name,t.cls,t.periods,t.subject,t.title])}]};
  if(kind==='students'){ const groups=studentGroups(studentList());
    const ws=withStatus(), wp=withPer();
    return {orientation:'port',headers:['م','اسم الطالب','الصف والشعبة','عدد الحصص'].concat(wp?['الحصص الغائبة']:[],ws?['الحالة']:[]),widths:[.5,3,1.6,.9].concat(wp?[2.6]:[],ws?[.9]:[]),
      groups:groups.map(g=>({title:`كشف الطلاب الغائبين${g.title?' – '+g.title:''}${school} ${when} (${g.rows.length} طالبًا)`,rows:g.rows.map((s,i)=>[i+1,s.name,`${s.grade} ث – ش${s.section}`,cnt(s),...(wp?[s.absent.join('، ')]:[]),...(ws?[s.status]:[])])}))} }
  if(kind==='teachers'){ const ab=model.teachers.filter((t,i)=>teacherAbsent[i]); const np=model.teachers.filter((t,i)=>!teacherAbsent[i]&&t.notPrepared);
    const groups=[{title:`المعلمون الغائبون${school} ${when} (${ab.length})`,rows:ab.map((t,i)=>[i+1,t.name,t.cls,t.periods,t.subject,t.present?'حُدّد غائبًا':'لم يُعدّ أي درس'])}];
    if(np.length) groups.push({title:`حصص غير معدّة لمعلمين حاضرين ${when}`,rows:np.map((t,i)=>[i+1,t.name,t.lessons.filter(b=>b.notPrepared).map(b=>`${b.grade} ش${b.section}`).join('، '),t.lessons.filter(b=>b.notPrepared).map(b=>b.period).join('، '),t.subject,`${t.notPrepared} غير معدّة`])});
    return {orientation:'port',headers:['م','اسم المعلم/ة','الصف والشعب','الحصص','المادة','ملاحظة'],widths:[.5,2.6,2,1,1.6,1.6],groups} }
  if(kind==='stats') return {orientation:'land',dateLine:true,headers:['البند','العدد','من إجمالي','النسبة المئوية','المقابل','العدد','النسبة المئوية'],widths:[2,1,1,1.3,1.6,1,1.3],
    groups:[{title:`إحصائيات اليوم${school} ${when}`,rows:statData().map(d=>[`${d.title} ${d.main}`,d.v,d.of,Core.pct(d.y[1],d.of),d.n[0],d.n[1],Core.pct(d.n[1],d.of)])}]};
  if(kind==='cut'){ const L=cutList(), cc=cutCols(), gsel=$('cGrade')&&$('cGrade').value; const base=`الطلاب المنقطعون تمامًا (غابوا عن الحصص ولم يدخلوا المنصة)${gsel?' – '+GN[gsel]:''}${school} ${when}`;
    const G=cutGroups(L);
    return {orientation:'port',headers:cutHeaders(cc),widths:cutWidths(cc),
      groups:G.map(g=>({title:g.title?`${base} – ${g.title} (${g.rows.length})`:`${base} (${L.length} طالبًا)`,rows:g.rows.map((c,i)=>cutCells(c,i,cc))}))} }
  if(kind==='alerts') return {orientation:'port',headers:['م','النوع','الاسم','التفاصيل'],widths:[.5,1.6,2.6,4],groups:[{title:`تنبيهات التقرير ${when}`,rows:alertRows().map((r,i)=>[i+1,...r])}]};
}

async function doExport(kind,fmt,btn){
  if(!model) return; btn.disabled=true;
  try{
    if(!formMeta) formMeta=await Core.formInfo(JSZip,files.form.bytes.slice(0));
    if(fmt==='xlsx'){
      const sp=listSpec(kind), d=curDate(), wb=XLSX.utils.book_new(); const aoa=[[(formMeta&&formMeta.school)||''],[`${d.day} ${d.hijriLong} – الموافق ${d.gregLong}`],[]];
      const merges=[]; sp.groups.forEach(g=>{ merges.push(aoa.length); aoa.push([g.title]); aoa.push(sp.headers); g.rows.forEach(r=>aoa.push(r)); aoa.push([]) });
      const ws=XLSX.utils.aoa_to_sheet(aoa); ws['!cols']=sp.headers.map(h=>({wch:/اسم/.test(h)?34:/الحصص الغائبة/.test(h)?22:14}));
      ws['!merges']=[0,1].concat(merges).map(r=>({s:{r,c:0},e:{r,c:sp.headers.length-1}})); ws['!views']=[{RTL:true}]; if(!wb.Workbook) wb.Workbook={Views:[{RTL:true}]};
      XLSX.utils.book_append_sheet(wb,ws,'المنقطعون'); const out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
      await offer(fname(NAMES[kind].replace(/\s/g,'_'),'xlsx'),new Blob([out]),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } else if(fmt==='csv'){
      const ws=withStatus(), wp=withPer(); const rows=[['م','اسم الطالب','الصف','الشعبة','عدد حصص الغياب'].concat(wp?['الحصص الغائبة']:[],ws?['الحالة في نقار']:[])];
      studentGroups(studentList()).forEach(g=>g.rows.forEach((s,i)=>rows.push([i+1,s.name,GN[s.grade]||s.grade,s.section,cnt(s)].concat(wp?[s.absent.join('، ')]:[],ws?[s.status]:[]))));
      await offer(fname('الطلاب_الغائبون','csv'),'﻿'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n'),'text/csv');
    } else if(fmt==='docx'){
      busy('تجهيز ملف Word …');
      let bytes;
      if(kind==='form'){ const r=await Core.fillDocx(JSZip,files.form.bytes.slice(0),exportModel(),{dates:curDate(),supervisor:$('sup').value,principal:$('pri').value,support:SUPPORT(),absentList:$('addAbsent').checked?formAbsentRows():[]}); bytes=r.bytes; if(r.warn.length) setStatus('لم أجد في الاستمارة: '+r.warn.join('، ')+'. عُبّئ الباقي.',true);
        if($('addNoLogin')&&$('addNoLogin').checked&&window.NoLogin&&window.NoLogin.hasData()){ const d=curDate(); bytes=await window.NoLogin.appendDocx(JSZip,bytes,{school:formMeta&&formMeta.school,day:d.day,hijri:d.hijriText}) } }
      else { const sp=listSpec(kind); const d=curDate(); sp.dateText=`اليوم: ${d.day}   |   التاريخ: ${d.hijriLong}   |   الموافق: ${d.gregLong}`; bytes=await Core.makeListDocx(JSZip,files.form.bytes.slice(0),sp) }
      await offer(fname(kind==='form'?'استمارة_المتابعة':NAMES[kind].replace(/\s/g,'_'),'docx'),new Blob([bytes]),'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    } else {
      const pages=kind==='form'?buildFormPages():(kind==='stats'?buildStatsPages():buildListPages(listSpec(kind)));
      const blob=await pagesToPdf(pages);
      const base=(fmt==='print'?'للطباعة_':'')+(kind==='form'?'استمارة_المتابعة':NAMES[kind].replace(/\s/g,'_'));
      await offer(fname(base,'pdf'),blob,'application/pdf');
    }
  }catch(e){ console.error(e); setStatus('تعذّر التنزيل: '+(e.message||e),true) }
  busy(); btn.disabled=false;
}

/* ---------- PDF pages (A4) ---------- */
function newPage(orient){
  const pg=document.createElement('div'); pg.className='pg '+orient;
  const head=document.createElement('div'); head.className='pg-head'; if(formMeta.img) head.style.backgroundImage=`url(${formMeta.img})`;
  head.innerHTML=`<div class="hl">${(formMeta.headLines||[]).map(esc).join('<br>')}</div>`;
  const d=curDate(); const band=document.createElement('div'); band.className='pg-date';
  band.innerHTML=`<span>اليوم: <b>${esc(d.day)}</b></span><span>التاريخ: <b>${esc(d.hijriLong)}</b></span><span>الموافق: <b>${esc(d.gregLong)}</b></span>${formMeta.school?`<span>المدرسة: <b>${esc(formMeta.school)}</b></span>`:''}`;
  const body=document.createElement('div'); body.className='pg-body';
  const foot=document.createElement('div'); foot.className='pg-foot';
  pg.append(head,band,body,foot); $('render').appendChild(pg); return {pg,body};
}
const fits=body=>body.scrollHeight<=body.clientHeight+1;
function tableHead(spec,title){ return `<colgroup>${spec.widths.map(w=>`<col style="width:${w}%">`).join('')}</colgroup><thead><tr><th class="ttl" colspan="${spec.headers.length}">${esc(title)}</th></tr><tr>${spec.headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead>` }
function addTable(pages,spec,group,orient){
  const sum=spec.widths.reduce((a,b)=>a+b,0); const sp=Object.assign({},spec,{widths:spec.widths.map(w=>w/sum*100)});
  let cur=newPage(orient); pages.push(cur.pg);
  let tbl=document.createElement('table'); tbl.className='pt'; tbl.innerHTML=tableHead(sp,group.title)+'<tbody></tbody>'; cur.body.appendChild(tbl);
  let tb=tbl.tBodies[0];
  const textCols=new Set(spec.headers.map((h,i)=>/اسم|عنوان|الحصص الغائبة|التفاصيل/.test(h)?i:-1));
  if(!group.rows.length){ tb.innerHTML=`<tr><td colspan="${spec.headers.length}">لا يوجد</td></tr>`; return }
  for(const row of group.rows){
    const tr=document.createElement('tr'); tr.innerHTML=row.map((v,i)=>`<td class="${textCols.has(i)?'lt':''}">${esc(v)}</td>`).join(''); tb.appendChild(tr);
    if(!fits(cur.body)&&tb.rows.length>1){ tb.removeChild(tr); cur=newPage(orient); pages.push(cur.pg);
      tbl=document.createElement('table'); tbl.className='pt'; tbl.innerHTML=tableHead(sp,group.title)+'<tbody></tbody>'; cur.body.appendChild(tbl); tb=tbl.tBodies[0]; tb.appendChild(tr) }
  }
}
function buildStatsPages(){
  $('render').innerHTML=''; const pages=[]; const cur=newPage('land'); pages.push(cur.pg); const d=curDate();
  cur.body.innerHTML=`<div class="p-title">إحصائيات اليوم — ${esc(d.day)} ${esc(d.hijriLong)}</div><div class="ps-grid">`+statData().map(x=>{const py=x.of?x.y[1]/x.of*100:0;
    return `<div class="ps" style="--pc:${x.pc}"><div class="k">${x.title} ${x.main}</div><div class="v">${x.v} <small>من ${x.of}</small></div>
    <div class="bar2"><i style="display:block;width:${py}%;background:#26794a"></i><i style="display:block;width:${100-py}%;background:${x.amb?'#a86412':'#b3372f'}"></i></div>
    <div class="row2"><div class="yy">نسبة ${x.y[0]} (${x.y[1]})<b>${Core.pct(x.y[1],x.of)}</b></div><div class="nn" ${x.amb?'style="background:#fbefdc"':''}>نسبة ${x.n[0]} (${x.n[1]})<b ${x.amb?'style="color:#a86412"':''}>${Core.pct(x.n[1],x.of)}</b></div></div></div>`}).join('')+`</div>`;
  const spec=listSpec('stats'); const tb=document.createElement('div'); tb.style.marginTop='10px'; cur.body.appendChild(tb);
  const sum=spec.widths.reduce((a,b)=>a+b,0); const sp=Object.assign({},spec,{widths:spec.widths.map(w=>w/sum*100)});
  tb.innerHTML=`<table class="pt">${tableHead(sp,spec.groups[0].title)}<tbody>${spec.groups[0].rows.map(r=>`<tr>${r.map(v=>`<td>${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  return pages;
}
/* ---- fully disconnected: absent in Naqar AND did not log in to the platform ---- */
function fmtPhone(p){ p=String(p||'').replace(/\D/g,''); const m=p.match(/(?:00)?9660?(5\d{8})$/)||p.match(/^0?(5\d{8})$/); return m?'0'+m[1]:(p||'') }
const CUT_COLS=[{k:'cnt',t:'عدد حصص الغياب'},{k:'per',t:'الحصص الغائبة'},{k:'st',t:'الحالة في نقار'},{k:'ph',t:'الجوال'},{k:'wa',t:'خانة التواصل (واتساب)',screen:true}];
function cutCols(){ const p=(typeof PREF!=='undefined'&&PREF.cutCols)||{}; const o={}; CUT_COLS.forEach(c=>{ o[c.k]=typeof p[c.k]==='boolean'?p[c.k]:c.k!=='st' }); return o }
function cutBase(){
  const NL=window.NoLogin; if(!model||!NL||!NL.hasData()) return {absent:[],cut:[]};
  const dayOnly=$('cDay')&&$('cDay').checked;
  const absent=model.students.filter(s=>cnt(s)>0&&(!dayOnly||s.status==='غائب'));
  const gi={'أول':1,'ثاني':2,'ثالث':3};
  const cut=absent.map(s=>{ const m=NL.matchStudent(s.name,s.grade,s.section); return m?{s,m,phone:fmtPhone(m.phone)}:null }).filter(Boolean)
    .sort((a,b)=>(gi[a.s.grade]||9)-(gi[b.s.grade]||9)||(+a.s.section||0)-(+b.s.section||0)||cnt(b.s)-cnt(a.s)||a.s.name.localeCompare(b.s.name,'ar'));
  return {absent,cut};
}
function cutList(){ const g=$('cGrade')?$('cGrade').value:''; return cutBase().cut.filter(c=>!g||c.s.grade===g) }
function cutGroups(L){
  if(!($('cGroup')&&$('cGroup').checked)) return [{title:'',rows:L}];
  const m=new Map(); L.forEach(c=>{ const k=`${GN[c.s.grade]||c.s.grade} – الشعبة ${c.s.section}`; if(!m.has(k)) m.set(k,[]); m.get(k).push(c) });
  return [...m].map(([title,rows])=>({title,rows}));
}
function cutCells(c,i,cc){ return [i+1,c.s.name,`${c.s.grade} ث – ش${c.s.section}`].concat(cc.cnt?[cnt(c.s)]:[],cc.per?[c.s.absent.join('، ')]:[],cc.st?[c.s.status||'—']:[],cc.ph?[c.phone||'—']:[]) }
function cutHeaders(cc){ return ['م','اسم الطالب','الصف والشعبة'].concat(cc.cnt?['عدد الحصص']:[],cc.per?['الحصص الغائبة']:[],cc.st?['الحالة في نقار']:[],cc.ph?['الجوال']:[]) }
function cutWidths(cc){ return [.5,3,1.5].concat(cc.cnt?[.9]:[],cc.per?[2.2]:[],cc.st?[1]:[],cc.ph?[1.5]:[]) }
function renderCutCols(){
  const box=$('cCols'); if(!box) return; const cc=cutCols();
  box.innerHTML='<span class="note"><b>الأعمدة في العرض والطباعة والتنزيل:</b></span>'+CUT_COLS.map(c=>`<label class="bk${c.screen?' scr':''}" for="cc-${c.k}"><input type="checkbox" id="cc-${c.k}" data-cc="${c.k}"${cc[c.k]?' checked':''}> ${c.t}${c.screen?' <span class="note">للشاشة فقط</span>':''}</label>`).join('')+
    '<button type="button" class="sm" id="ccMin">الاسم والفصل فقط</button><button type="button" class="sm" id="ccAll">إظهار الكل</button>';
}
function renderCut(){
  if(!$('cBody')) return; const NL=window.NoLogin;
  renderCutCols();
  const base=cutBase(), L=cutList(), cc=cutCols(); $('n-c').textContent=base.cut.length||'';
  const head=cutHeaders(cc).concat(cc.wa?['تواصل']:[]); $('cHead').innerHTML='<tr>'+head.map(h=>`<th>${esc(h)}</th>`).join('')+'</tr>';
  const span=head.length, msg=t=>{ $('cStats').innerHTML=''; $('cBody').innerHTML=`<tr><td colspan="${span}" class="empty-cell">${t}</td></tr>` };
  if(!model) return;
  if(!NL||!NL.hasData()) return msg('ارفع ملف «طلاب غائبون (منصة مدرستي)» في الأعلى لتظهر هنا قائمة المنقطعين تمامًا.');
  if(!model.students.length) return msg('ارفع تقرير «غياب الحصص للطلاب» من نقار ثم أنشئ التقرير لتظهر هنا قائمة المنقطعين تمامًا.');
  const A=base.absent.length, C=base.cut.length, byG=g=>base.cut.filter(c=>c.s.grade===g).length, withPh=base.cut.filter(c=>c.phone).length;
  const pc=(a,b)=>b?Core.pct(a,b):'—';
  $('cStats').innerHTML=[
    {c:'var(--red)',k:'المنقطعون تمامًا',v:C,sub:`من ${A} غائبًا في تقرير الغياب · <bdi>${pc(C,A)}</bdi>`,main:1},
    {c:'var(--teal)',k:'الصف الأول ثانوي',v:byG('أول'),sub:`<bdi>${pc(byG('أول'),C)}</bdi> من المنقطعين`},
    {c:'var(--blue)',k:'الصف الثاني ثانوي',v:byG('ثاني'),sub:`<bdi>${pc(byG('ثاني'),C)}</bdi> من المنقطعين`},
    {c:'#7a4f9e',k:'الصف الثالث ثانوي',v:byG('ثالث'),sub:`<bdi>${pc(byG('ثالث'),C)}</bdi> من المنقطعين`},
    {c:'var(--green)',k:'غابوا لكن دخلوا المنصة',v:A-C,sub:'يتابعون عن بُعد رغم الغياب'},
    {c:'#1f8f4e',k:'لهم رقم جوال',v:withPh,sub:'جاهزون للتواصل بالواتساب'}
  ].map(x=>`<div class="cst${x.main?' main':''}" style="--c:${x.c}"><div class="k">${x.k}</div><div class="v">${x.v}</div><div class="s">${x.sub}</div></div>`).join('');
  if(!L.length) return msg(C?'لا يوجد منقطعون في الصف المحدد.':'لا يوجد طالب غاب عن الحصص ولم يدخل المنصة معًا. ممتاز!');
  let h='';
  cutGroups(L).forEach(g=>{
    if(g.title) h+=`<tr class="grp"><td colspan="${span}">${esc(g.title)} <span class="pill a">${g.rows.length}</span></td></tr>`;
    g.rows.forEach((c,i)=>{ const cells=cutCells(c,i,cc); const href=cc.wa?NL.parentLink({name:c.s.name,phone:c.m.phone,g:c.m.g,s:c.m.s}):'';
      h+='<tr'+(c.s.status==='غائب'?' class="flag"':'')+'>'+cells.map((v,j)=>{ const hd=cutHeaders(cc)[j];
        return hd==='الجوال'?`<td class="ltr-cell">${esc(v)}</td>`:j===0||hd==='عدد الحصص'?`<td class="num">${esc(v)}</td>`:j===1?`<td class="nm">${esc(v)}</td>`:`<td>${esc(v)}</td>` }).join('')+
        (cc.wa?`<td>${href?`<a class="wa-btn" target="_blank" rel="noopener" href="${esc(href)}">واتساب</a>`:'—'}</td>`:'')+'</tr>' }) });
  $('cBody').innerHTML=h;
}
if($('cCols')){
  $('cCols').addEventListener('change',e=>{ const k=e.target.dataset&&e.target.dataset.cc; if(!k) return; const p=Object.assign({},cutCols()); p[k]=e.target.checked; savePref('cutCols',p); renderCut() });
  $('cCols').addEventListener('click',e=>{ const b=e.target.closest('button'); if(!b) return; const all=b.id==='ccAll'; const p={}; CUT_COLS.forEach(c=>p[c.k]=all); savePref('cutCols',p); renderCut() });
}
['cGrade','cGroup'].forEach(id=>{ if(!$(id)) return; const v=PREF['cut_'+id]; if(id==='cGroup'&&typeof v==='boolean') $(id).checked=v; if(id==='cGrade'&&typeof v==='string') $(id).value=v;
  $(id).addEventListener('change',()=>{ savePref('cut_'+id,id==='cGroup'?$(id).checked:$(id).value); renderCut() }) });
document.querySelectorAll('input[name="cMode"]').forEach(r=>r.addEventListener('change',renderCut));
document.addEventListener('nologin:data',()=>{ if(model) renderCut() });
document.addEventListener('nologin:cleared',()=>{ if(model) renderCut() });
document.addEventListener('nologin:change',()=>{ if(model) renderCut() });

/* ---- shared card + generic report pages (used by the platform section and the follow-up log) ---- */
function psCard(x){
  if(x.of==null) return `<div class="ps" style="--pc:${x.pc}"><div class="k">${esc(x.title)} — لم يدخلوا</div><div class="v">${x.v} <small>طالبًا</small></div>
    <div class="bar2"><i style="display:block;width:${((x.ratio||0)*100).toFixed(1)}%;background:#b3372f"></i></div>
    ${x.share?`<div class="row2" style="grid-template-columns:1fr"><div class="nn">حصته من قائمة الغائبين${x.listTotal?` (${x.listTotal})`:''}<b>${esc(x.share)}</b></div></div>`:''}</div>`;
  const py=x.of?x.inn/x.of*100:0;
  return `<div class="ps" style="--pc:${x.pc}"><div class="k">${esc(x.title)} — ${x.days?'متوسط يومي':'لم يدخلوا'}</div><div class="v">${x.v} <small>من ${x.of}</small></div>
  <div class="bar2"><i style="display:block;width:${py}%;background:#26794a"></i><i style="display:block;width:${100-py}%;background:#b3372f"></i></div>
  <div class="row2"><div class="yy">نسبة الدخول (${x.inn})<b>${esc(x.pIn)}</b></div><div class="nn">نسبة عدم الدخول (${x.v})<b>${esc(x.pOut)}</b></div></div></div>`;
}
function buildReportPages(spec){
  $('render').innerHTML=''; const pages=[]; const cur=newPage('land'); pages.push(cur.pg);
  let html=`<div class="p-title">${esc(spec.title)}</div>`;
  if(spec.cards&&spec.cards.length) html+=`<div class="ps-grid" style="grid-template-columns:repeat(${Math.min(4,spec.cards.length)},1fr)">${spec.cards.map(psCard).join('')}</div>`;
  if(spec.img) html+=`<img alt="" src="${spec.img}" style="display:block;width:100%;max-height:300px;object-fit:contain;margin-top:10px;border:1px solid #c9d6d8;border-radius:10px">`;
  cur.body.innerHTML=html;
  (spec.tables||[]).forEach(t=>{ if(t.rows&&t.rows.length) addTable(pages,{headers:t.headers,widths:t.widths||t.headers.map(()=>1)},{title:t.title,rows:t.rows},'land') });
  return pages;
}
async function ensureForm(){ if(!files.form) throw new Error('انتظر لحظة حتى تُحمَّل الاستمارة المحفوظة ثم أعد المحاولة.'); if(!formMeta) formMeta=await Core.formInfo(JSZip,files.form.bytes.slice(0)) }
window.BalatReport=async function(spec,fmt){
  try{ busy('تجهيز PDF …'); await ensureForm();
    const blob=await pagesToPdf(buildReportPages(spec));
    await offer(fname((fmt==='print'?'للطباعة_':'')+(spec.name||'تقرير'),'pdf'),blob,'application/pdf');
  }catch(e){ console.error(e); setStatus('تعذّر التنزيل: '+(e.message||e),true) }
  busy();
};

/* ---- full day report in one PDF ---- */
async function fullDayReport(){
  if(!model) return;
  const btn=$('fullReport'); btn.disabled=true;
  try{
    busy('تجهيز التقرير الكامل …'); await ensureForm();
    const NL=window.NoLogin;
    const builders=[()=>buildStatsPages(),()=>buildFormPages()];
    if(NL&&NL.hasData()&&NL.gradeCards()) builders.push(()=>buildGradePages());
    if(cutList().length) builders.push(()=>buildListPages(listSpec('cut')));
    let total=0; for(const b of builders) total+=b().length;
    const {jsPDF}=window.jspdf; let pdf=null, idx=0;
    try{ await document.fonts.ready }catch(e){}
    for(const b of builders){ const pages=b();
      for(const pg of pages){ idx++; busy(`تجهيز التقرير الكامل: صفحة ${idx} من ${total} …`);
        const land=pg.classList.contains('land'); const num=document.createElement('div'); num.className='pg-num'; num.textContent=`صفحة ${idx} من ${total}`; pg.appendChild(num);
        const canvas=await html2canvas(pg,{scale:1.6,backgroundColor:'#ffffff',logging:false,useCORS:true});
        const o=land?'landscape':'portrait'; if(!pdf) pdf=new jsPDF({orientation:o,unit:'mm',format:'a4',compress:true}); else pdf.addPage('a4',o);
        pdf.addImage(canvas.toDataURL('image/jpeg',0.82),'JPEG',0,0,land?297:210,land?210:297,undefined,'FAST'); } }
    $('render').innerHTML='';
    await offer(fname('تقرير_اليوم_الكامل','pdf'),pdf.output('blob'),'application/pdf');
  }catch(e){ console.error(e); setStatus('تعذّر إنشاء التقرير الكامل: '+(e.message||e),true) }
  busy(); btn.disabled=false;
}
if($('fullReport')) $('fullReport').addEventListener('click',fullDayReport);

/* per-grade platform statistics, drawn exactly like «إحصائيات اليوم» */
function buildGradePages(){
  const G=window.NoLogin&&window.NoLogin.gradeCards&&window.NoLogin.gradeCards(); const T=window.NoLogin.gradeSummary();
  $('render').innerHTML=''; const pages=[]; const cur=newPage('land'); pages.push(cur.pg); const d=curDate();
  const n=G.cards.length;
  cur.body.innerHTML=`<div class="p-title">نسبة الدخول لمنصة مدرستي حسب الصف — ${esc(d.day)} ${esc(d.hijriLong)}</div><div class="ps-grid" style="grid-template-columns:repeat(${n},1fr)">`+G.cards.map(x=>{ if(x) return psCard(Object.assign({listTotal:G.listTotal},x));
    if(x.of==null) return `<div class="ps" style="--pc:${x.pc}"><div class="k">${esc(x.title)} — لم يدخلوا</div><div class="v">${x.v} <small>طالبًا</small></div>
      <div class="bar2"><i style="display:block;width:${(x.ratio*100).toFixed(1)}%;background:#b3372f"></i></div>
      <div class="row2" style="grid-template-columns:1fr"><div class="nn">حصته من قائمة الغائبين (${G.listTotal})<b>${esc(x.share)}</b></div></div></div>`;
    const py=x.of?x.inn/x.of*100:0;
    return `<div class="ps" style="--pc:${x.pc}"><div class="k">${esc(x.title)} — لم يدخلوا</div><div class="v">${x.v} <small>من ${x.of}</small></div>
    <div class="bar2"><i style="display:block;width:${py}%;background:#26794a"></i><i style="display:block;width:${100-py}%;background:#b3372f"></i></div>
    <div class="row2"><div class="yy">نسبة الدخول (${x.inn})<b>${esc(x.pIn)}</b></div><div class="nn">نسبة عدم الدخول (${x.v})<b>${esc(x.pOut)}</b></div></div></div>`}).join('')+`</div>`;
  const spec={headers:T.headers,widths:T.headers.map((h,i)=>i?1:1.6)}; const sum=spec.widths.reduce((a,b)=>a+b,0); const sp=Object.assign({},spec,{widths:spec.widths.map(w=>w/sum*100)});
  const tb=document.createElement('div'); tb.style.marginTop='12px'; cur.body.appendChild(tb);
  tb.innerHTML=`<table class="pt">${tableHead(sp,'نسبة الدخول لمنصة مدرستي حسب الصف'+(G.full?'':' (الملف لا يحتوي على إجمالي الطلاب)'))}<tbody>${T.rows.map((r,i)=>`<tr${G.full&&i===T.rows.length-1?' style="font-weight:700"':''}>${r.map(v=>`<td>${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  const ST=window.NoLogin.sectionTable&&window.NoLogin.sectionTable();
  if(ST&&ST.rows.length) addTable(pages,{headers:ST.headers,widths:ST.headers.map((h,i)=>i===1?2:1)},{title:'ترتيب الشعب حسب نسبة الدخول لمنصة مدرستي'+(ST.full?'':' (حسب عدد غير الداخلين)'),rows:ST.rows},'land');
  return pages;
}
window.BalatGradeExport=async function(fmt){
  const btn=document.activeElement; try{
    if(!(window.NoLogin&&window.NoLogin.gradeCards&&window.NoLogin.gradeCards())) return;
    if(!files.form){ setStatus('انتظر لحظة حتى تُحمَّل الاستمارة المحفوظة ثم أعد المحاولة.',true); return }
    busy('تجهيز PDF …');
    if(!formMeta) formMeta=await Core.formInfo(JSZip,files.form.bytes.slice(0));
    const blob=await pagesToPdf(buildGradePages());
    await offer(fname((fmt==='print'?'للطباعة_':'')+'نسبة_الدخول_للمنصة_حسب_الصف','pdf'),blob,'application/pdf');
  }catch(e){ console.error(e); setStatus('تعذّر التنزيل: '+(e.message||e),true) }
  busy();
};
function buildListPages(spec){ $('render').innerHTML=''; const pages=[]; spec.groups.forEach(g=>addTable(pages,spec,g,spec.orientation)); return pages }
function buildFormPages(){
  $('render').innerHTML=''; const pages=[]; const m=formMeta, d=curDate(), T=tStats(), S=sStats(), L=model.lessons;
  const cur=newPage('land'); pages.push(cur.pg);
  const box=v=>v?'☑':'☐';
  cur.body.innerHTML=`<div class="p-title">${esc(m.title)}</div>
  <div class="p-sec">أولاً : بيانات عامة :</div>
  <table class="pt"><tbody>
   <tr><th class="hd">اليوم</th><td colspan="2">${esc(d.day)}</td><th class="hd" colspan="2">التاريخ</th><td colspan="2">${esc(d.hijriText)}</td><th class="hd">القطاع</th><td>${esc(m.sector)}</td><th class="hd">الفئة</th><td>${esc(m.category)}</td></tr>
   <tr><th class="hd" rowspan="2">المرحلة</th><td>طفولة مبكرة</td><td colspan="2">ابتدائي</td><td colspan="2">متوسط</td><td>ثانوي</td><th class="hd" rowspan="2">اسم المدرسة</th><td rowspan="2">${esc(m.school)}</td><th class="hd" rowspan="2">اسم مدير المدرسة</th><td rowspan="2">${esc(m.principal)}</td></tr>
   <tr><td>${box(false)}</td><td colspan="2">${box(false)}</td><td colspan="2">${box(false)}</td><td>${box(m.stage==='ثانوي')}</td></tr>
  </tbody></table>
  <div class="p-sec">ثانياً : متابعة الانضباط لكلاً من المعلمين/ات والطلبة :</div>
  <table class="pt" style="width:64%;margin-inline:auto"><thead><tr><th class="hd">الفئة</th><th class="hd">الاجمالي</th><th class="hd">الحضور</th><th class="hd">الغياب</th><th class="hd">نسبة الحضور</th><th class="hd">نسبة الغياب</th></tr></thead><tbody>
   <tr><td>المعلمين/ات</td><td>${T.total}</td><td style="color:#26794a">${T.present}</td><td style="color:#b3372f">${T.absent}</td><td style="color:#26794a">${Core.pct(T.present,T.total)}</td><td style="color:#b3372f">${Core.pct(T.absent,T.total)}</td></tr>
   <tr><td>الطلبة</td><td>${S.total}</td><td style="color:#26794a">${S.present}</td><td style="color:#b3372f">${S.absent}</td><td style="color:#26794a">${Core.pct(S.present,S.total)}</td><td style="color:#b3372f">${Core.pct(S.absent,S.total)}</td></tr>
  </tbody></table>
  <div class="p-sec">ثالثاً : متابعة الدخول لمنصة مدرستي :</div>
  <table class="pt"><thead><tr><th class="hd">الجدول المدرسي</th><th class="hd">عدد حضور الحصص لمشرف/ة المدرسة</th><th class="hd">عدد حضور الحصص لمدير/ة المدرسة</th><th class="hd">نسبة دخول المعلمين/ات</th><th class="hd">نسبة دخول الطلبة</th><th class="hd">عدد الحصص المنفذة لهذا اليوم</th></tr></thead><tbody>
   <tr><td>مكتمل</td><td>${esc($('sup').value)}</td><td>${esc($('pri').value)}</td><td>${Core.pct(T.present,T.total)}</td><td>${Core.pct(S.present,S.total)}</td><td>${L.prepared} من ${L.all}</td></tr>
  </tbody></table>
  <div class="p-sec">رابعاً : الدعم الذي تحتاجه المدرسة:</div>
  <div style="font-size:13px;font-weight:600;display:grid;gap:2px;padding-inline-start:14px">${(SUPPORT().length?SUPPORT():['','']).map((s,i)=>`<div>${i+1}. ${esc(s)}</div>`).join('')}</div>
  <div class="p-prep">${(m.preparer||[]).map((l,i)=>i===0?`<b>${esc(l)}</b>`:`<span>${esc(l)}</span>`).join('')}</div>`;
  const school=m.school?'لمدرسة '+m.school:'';
  addTable(pages,{headers:['اسم المعلم/ة','الصف','الحصة','المادة','عنوان الدرس'],widths:[2.2,2.2,1,1.6,3.6]},{title:`استمارة متابعة سير العملية التعليمية (حضور حصص افتراضية) ${school} بسبب تعليق الدراسة ليوم ${d.day} الموافق ${d.hijriText}`,rows:model.teachers.map(t=>[t.name,t.cls,t.periods,t.subject,t.title])},'land');
  if($('addAbsent').checked){ const groups=studentGroups(studentList({ignoreSearch:true}));
    groups.forEach(g=>addTable(pages,{headers:['م','اسم الطالب','الصف والشعبة','عدد حصص الغياب','الحصص الغائبة'].concat(withStatus()?['الحالة']:[]),widths:[.5,3,1.6,1,2.6].concat(withStatus()?[1]:[])},{title:`كشف الطلاب الغائبين${g.title?' – '+g.title:''} ${school} ليوم ${d.day} الموافق ${d.hijriText} (${g.rows.length} طالبًا)`,rows:g.rows.map((s,i)=>[i+1,s.name,`${s.grade} ث – ش${s.section}`,cnt(s),s.absent.join('، '),...(withStatus()?[s.status]:[])])},'land')) }
  if($('addNoLogin')&&$('addNoLogin').checked&&window.NoLogin&&window.NoLogin.hasData()&&window.NoLogin.count()){
    window.NoLogin.formGroups({school:m.school,day:d.day,hijri:d.hijriText}).forEach(g=>addTable(pages,{headers:g.headers,widths:g.widths},{title:g.title,rows:g.rows},'land')) }
  return pages;
}
async function pagesToPdf(pages){
  const {jsPDF}=window.jspdf; let pdf=null;
  try{ await document.fonts.ready }catch(e){}
  for(let i=0;i<pages.length;i++){
    busy(`تجهيز PDF: صفحة ${i+1} من ${pages.length} …`);
    const pg=pages[i]; const land=pg.classList.contains('land');
    const num=document.createElement('div'); num.className='pg-num'; num.textContent=`صفحة ${i+1} من ${pages.length}`; pg.appendChild(num);
    const canvas=await html2canvas(pg,{scale:1.6,backgroundColor:'#ffffff',logging:false,useCORS:true});
    const img=canvas.toDataURL('image/jpeg',0.82);
    const o=land?'landscape':'portrait';
    if(!pdf) pdf=new jsPDF({orientation:o,unit:'mm',format:'a4',compress:true}); else pdf.addPage('a4',o);
    const W=land?297:210,H=land?210:297; pdf.addImage(img,'JPEG',0,0,W,H,undefined,'FAST');
  }
  $('render').innerHTML='';
  return pdf.output('blob');
}

})();
