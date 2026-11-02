'use strict';
(function(){

document.documentElement.setAttribute('dir','rtl');document.documentElement.setAttribute('lang','ar');
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const GO={'أول':1,'ثاني':2,'ثالث':3};
const GN={'أول':'أول ثانوي','ثاني':'ثاني ثانوي','ثالث':'ثالث ثانوي'};
const files={form:null,teachers:null,students:null};
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
  for(const k of ['form','teachers','students']){
    const f=files[k]; $('slot-'+k).classList.toggle('ok',!!f);
    $('name-'+k).textContent=f?(f.builtin?f.name+' — ارفع ملفًا آخر لاستبدالها':f.name):'لم يُرفع ملف';
    $('tag-'+k).textContent=f?'جاهز':(k==='form'?'Word':'PDF / Word');
  }
  $('run').disabled=!(files.form&&files.teachers&&files.students);
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
  if(/متابعة سير العملية|نسبة الحضور/.test(txt)) return {kind:'formpdf',bytes};
  if(/تحضير|عنوان|الدرس/.test(txt)) return {kind:'teachers',bytes,pages};
  if(/الحالة|غياب|الطالب/.test(txt)) return {kind:'students',bytes,pages};
  return {kind:null,bytes};
}
async function addFiles(list,hint){
  let err='';
  for(const f of list){
    setStatus('جارٍ قراءة '+f.name+' …');
    try{
      const r=await classify(f);
      if(r.kind==='formpdf'){ err='الاستمارة بصيغة PDF لا يمكن تعبئتها. ارفعها بصيغة Word، وحتى ذلك الحين تُستخدم الاستمارة المحفوظة.'; continue }
      const k=r.kind||(hint!=='form'?hint:null);
      if(!k){ err='لم أتعرف على الملف «'+f.name+'». تأكد أنه تقرير نقار أو استمارة المتابعة.'; continue }
      files[k]={name:f.name,bytes:r.bytes,pages:r.pages,docx:r.docx};
      if(k==='form') formMeta=null;
    }catch(e){ console.error(e); err=(e&&/[\u0600-\u06FF]/.test(e.message||''))?e.message:'تعذّرت قراءة «'+f.name+'». تأكد أن الملف PDF أو Word سليم.' }
  }
  refreshSlots();
  if(err) setStatus(err,true); else setStatus($('run').disabled?'ارفع الملفات الناقصة.':'الملفات جاهزة. اضغط «إنشاء التقرير».');
}
for(const k of ['form','teachers','students']){
  const slot=$('slot-'+k), inp=$('in-'+k);
  inp.addEventListener('change',()=>{addFiles([...inp.files],k);inp.value=''});
  slot.addEventListener('dragover',e=>{e.preventDefault();slot.classList.add('drag')});
  slot.addEventListener('dragleave',()=>slot.classList.remove('drag'));
  slot.addEventListener('drop',e=>{e.preventDefault();slot.classList.remove('drag');addFiles([...e.dataTransfer.files],k)});
}
refreshSlots();

$('run').addEventListener('click',async()=>{
  $('run').disabled=true;
  try{
    const {extract}=Core.makeExtractor(PDFLib);
    setStatus('قراءة تحضير المعلمين …');
    let T; const tf=files.teachers;
    if(tf.docx) T=Core.finishTeachers(Core.parseTeachersDocx(tf.docx));
    else T=Core.parseTeachers(tf.pages||await extract(tf.bytes.slice(0)));
    if(!T.teachers.length) throw new Error('لم أجد جدول المعلمين في ملف التحضير. تأكد أنه «تقرير تحضير دروس المعلمين» من نقار.');
    setStatus('قراءة غياب الطلاب …');
    let S; const sf=files.students;
    if(sf.docx) S=Core.parseStudentsDocx(sf.docx); else S=Core.parseStudents(sf.pages||await extract(sf.bytes.slice(0)));
    if(!S.students.length) throw new Error('لم أجد قائمة الطلاب في ملف الغياب. تأكد أنه «تقرير غياب الحصص للطلاب» من نقار.');
    model=Core.buildModel(T,S);
    teacherAbsent=model.teachers.map(t=>!t.present);
    let dateNote='';
    if(!dateTouched){ hijriManual=false; if(model.dates.gregText||model.dates.hijriText) applyGreg(model.dates.gregText||'',model.dates.hijriText) }
    else if(model.dates.hijriText&&model.dates.hijriText!==curDate().hijriText){ dateNote=` تنبيه: أبقيت التاريخ الذي اخترته (${curDate().day} ${curDate().hijriText})، بينما تاريخ ملفات نقار ${model.dates.day||''} ${model.dates.hijriText}.` }
    render();
    setStatus(`تمت القراءة: ${model.tStats.total} معلمًا و${model.sStats.total} طالبًا.`+dateNote);
  }catch(e){console.error(e);setStatus(e.message||'تعذّرت قراءة الملفات.',true)}
  $('run').disabled=false;
});

/* ---------- derived data ---------- */
function tStats(){ const total=model.teachers.length, absent=teacherAbsent.filter(Boolean).length; return {total,present:total-absent,absent} }
function cnt(s){ return s.absent.length }
function bucketOf(n){ return n>=7?7:n }
const withStatus=()=>$('fPartial').checked;
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
  renderFigures(); renderTeachers(); renderStudents(); renderAbsentTeachers(); renderWarnings(); updateDlNote();
}
function statData(){
  const T=tStats(),S=model.sStats,L=model.lessons; const sync=Math.max(0,L.prepared-L.async);
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
  $('tBody').querySelectorAll('input').forEach(inp=>inp.addEventListener('input',()=>{model.teachers[+inp.dataset.i].subject=inp.value.trim();inp.classList.toggle('empty',!inp.value.trim());renderWarnings();renderAbsentTeachers()}));
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
  for(const gr of groups){ if(gr.title) html+=`<tr class="grp"><td colspan="${withStatus()?7:6}">${esc(gr.title)} — ${gr.rows.length} طالبًا</td></tr>`;
    html+=gr.rows.map((s,i)=>`<tr><td class="num">${i+1}</td><td>${esc(s.name)}</td><td>${esc(GN[s.grade]||s.grade)}</td><td class="num"><bdi>ش ${s.section??''}</bdi></td><td class="num"><b>${cnt(s)}</b></td><td>${esc(s.absent.join('، '))}</td>${withStatus()?`<td>${s.status==='غائب'?'<span class="pill a">غائب</span>':'<span class="pill p">حاضر</span>'}</td>`:''}</tr>`).join('') }
  $('sBody').innerHTML=html||`<tr><td colspan="7" class="note">لا يوجد طلاب بهذه التصفية.</td></tr>`;
  $('thStatus').hidden=!withStatus();
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
  $('dlNote').textContent=$('addAbsent').checked?`يُضاف في آخرها كشف بـ ${n} طالبًا حسب تصفية تبويب الطلاب الغائبين.`:'بدون كشف الغائبين.';
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
['fGrade','fPartial','fSearch','fMerge','b1','b2','b3','b4','b5','b6','b7'].forEach(id=>$(id).addEventListener('input',()=>{ $('bAll').checked=[1,2,3,4,5,6,7].every(i=>$('b'+i).checked); renderStudents();updateDlNote()}));
$('addAbsent').addEventListener('change',updateDlNote);
const TABS=[['tab-t','pane-t'],['tab-s','pane-s'],['tab-a','pane-a'],['tab-w','pane-w']];
for(const [id] of TABS) $(id).addEventListener('click',()=>{for(const [i2,p2] of TABS){$(i2).setAttribute('aria-selected',i2===id);$(p2).hidden=i2!==id}});

/* ---------- export bars ---------- */
const ICON={
  docx:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M8.5 12.5l1.2 4.5 1.3-3.5 1.3 3.5 1.2-4.5"/></svg>',
  pdf:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M9 17v-4h1.3a1.2 1.2 0 0 1 0 2.4H9"/><path d="M14 13v4"/><path d="M14 13h1.6"/><path d="M14 15h1.2"/></svg>',
  print:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 9V3h10v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/></svg>',
  csv:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 4v16M15 4v16"/></svg>'
};
const NAMES={stats:'إحصائيات_اليوم',form:'الاستمارة',lessons:'الحصص الافتراضية',students:'الطلاب الغائبين',teachers:'غياب المعلمين',alerts:'التنبيهات'};
document.querySelectorAll('.exp').forEach(el=>{
  const k=el.dataset.kind;
  el.innerHTML=`<span class="lab">${k==='form'?'تنزيل الاستمارة:':'تنزيل:'}</span>`+
    `<button class="word" data-fmt="docx">${ICON.docx}Word</button><button data-fmt="pdf">${ICON.pdf}PDF</button><button data-fmt="print">${ICON.print}طباعة</button>`+(k==='students'?`<button data-fmt="csv">${ICON.csv}Excel</button>`:'');
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

function exportModel(){ return Object.assign({},model,{tStats:tStats(),dates:Object.assign({},model.dates,curDate())}) }
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
    const ws=withStatus();
    return {orientation:'port',headers:['م','اسم الطالب','الصف والشعبة','عدد الحصص','الحصص الغائبة'].concat(ws?['الحالة']:[]),widths:[.5,3,1.6,.9,2.6].concat(ws?[.9]:[]),
      groups:groups.map(g=>({title:`كشف الطلاب الغائبين${g.title?' – '+g.title:''}${school} ${when} (${g.rows.length} طالبًا)`,rows:g.rows.map((s,i)=>[i+1,s.name,`${s.grade} ث – ش${s.section}`,cnt(s),s.absent.join('، '),...(ws?[s.status]:[])])}))} }
  if(kind==='teachers'){ const ab=model.teachers.filter((t,i)=>teacherAbsent[i]); const np=model.teachers.filter((t,i)=>!teacherAbsent[i]&&t.notPrepared);
    const groups=[{title:`المعلمون الغائبون${school} ${when} (${ab.length})`,rows:ab.map((t,i)=>[i+1,t.name,t.cls,t.periods,t.subject,t.present?'حُدّد غائبًا':'لم يُعدّ أي درس'])}];
    if(np.length) groups.push({title:`حصص غير معدّة لمعلمين حاضرين ${when}`,rows:np.map((t,i)=>[i+1,t.name,t.lessons.filter(b=>b.notPrepared).map(b=>`${b.grade} ش${b.section}`).join('، '),t.lessons.filter(b=>b.notPrepared).map(b=>b.period).join('، '),t.subject,`${t.notPrepared} غير معدّة`])});
    return {orientation:'port',headers:['م','اسم المعلم/ة','الصف والشعب','الحصص','المادة','ملاحظة'],widths:[.5,2.6,2,1,1.6,1.6],groups} }
  if(kind==='stats') return {orientation:'land',dateLine:true,headers:['البند','العدد','من إجمالي','النسبة المئوية','المقابل','العدد','النسبة المئوية'],widths:[2,1,1,1.3,1.6,1,1.3],
    groups:[{title:`إحصائيات اليوم${school} ${when}`,rows:statData().map(d=>[`${d.title} ${d.main}`,d.v,d.of,Core.pct(d.y[1],d.of),d.n[0],d.n[1],Core.pct(d.n[1],d.of)])}]};
  if(kind==='alerts') return {orientation:'port',headers:['م','النوع','الاسم','التفاصيل'],widths:[.5,1.6,2.6,4],groups:[{title:`تنبيهات التقرير ${when}`,rows:alertRows().map((r,i)=>[i+1,...r])}]};
}

async function doExport(kind,fmt,btn){
  if(!model) return; btn.disabled=true;
  try{
    if(!formMeta) formMeta=await Core.formInfo(JSZip,files.form.bytes.slice(0));
    if(fmt==='csv'){
      const ws=withStatus(); const rows=[['م','اسم الطالب','الصف','الشعبة','عدد حصص الغياب','الحصص الغائبة'].concat(ws?['الحالة في نقار']:[])];
      studentGroups(studentList()).forEach(g=>g.rows.forEach((s,i)=>rows.push([i+1,s.name,GN[s.grade]||s.grade,s.section,cnt(s),s.absent.join('، ')].concat(ws?[s.status]:[]))));
      await offer(fname('الطلاب_الغائبون','csv'),'﻿'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n'),'text/csv');
    } else if(fmt==='docx'){
      busy('تجهيز ملف Word …');
      let bytes;
      if(kind==='form'){ const r=await Core.fillDocx(JSZip,files.form.bytes.slice(0),exportModel(),{dates:curDate(),supervisor:$('sup').value,principal:$('pri').value,support:SUPPORT(),absentList:$('addAbsent').checked?formAbsentRows():[]}); bytes=r.bytes; if(r.warn.length) setStatus('لم أجد في الاستمارة: '+r.warn.join('، ')+'. عُبّئ الباقي.',true) }
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
function buildListPages(spec){ $('render').innerHTML=''; const pages=[]; spec.groups.forEach(g=>addTable(pages,spec,g,spec.orientation)); return pages }
function buildFormPages(){
  $('render').innerHTML=''; const pages=[]; const m=formMeta, d=curDate(), T=tStats(), S=model.sStats, L=model.lessons;
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
