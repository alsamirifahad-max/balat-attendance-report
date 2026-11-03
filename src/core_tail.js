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
