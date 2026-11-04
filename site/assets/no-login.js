/* خانة «الطلاب الذين لم يدخلوا المنصة» — وحدة مستقلة تمامًا.
   تعمل داخل Shadow DOM فلا تتأثر بتنسيقات الموقع ولا تؤثر عليها، ولا تعتمد على app.js.
   تحتاج: vendor/xlsx.full.min.js (قراءة وحفظ Excel) + مكتبات الموقع الموجودة jszip و html2canvas و jspdf. */
(function(){
"use strict";
var HOST=document.getElementById("nl-root"); if(!HOST||HOST.shadowRoot) return;
var R=HOST.attachShadow({mode:"open"});
R.innerHTML="<style>"+"\n:host{--bg:var(--paper,#f4f7f6);--muted:var(--ink-2,#4a5f5d);--accent:var(--teal,#1d6f68);--accent-ink:var(--on-teal,#fff);--accent-soft:var(--teal-soft,#e2f0ed);--miss:var(--red,#b3372f);--miss-soft:var(--red-soft,#f9e3e0);--font:var(--body,Tahoma,sans-serif)}\nh1,h2,h3{font-family:var(--display,inherit);font-weight:600}\n*,*::before,*::after{box-sizing:border-box}\n:host{display:block;color:var(--ink);font-family:var(--font);font-size:15px;line-height:1.6;margin-top:22px}\n.wrap{display:grid;grid-template-columns:minmax(0,1fr);gap:16px}\nh1{margin:0;font-size:24px;text-wrap:balance}\nh2{margin:0;font-size:16px}\n.sub,.note{color:var(--muted)}\n.note{font-size:13px;margin:0}\n.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px;display:grid;gap:12px;min-width:0}\n.row{display:flex;flex-wrap:wrap;gap:10px;align-items:center}\n.field{display:grid;gap:4px;font-size:13px;color:var(--muted)}\ninput,select{font:inherit;color:var(--ink);background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:7px 10px;min-width:0;max-width:100%}\nbutton{font:inherit;cursor:pointer}\n:focus-visible{outline:2px solid var(--accent);outline-offset:2px}\n\n/* grade segmented */\n.seg{display:flex;flex-wrap:wrap;gap:6px}\n.seg button{border:1px solid var(--line);background:var(--bg);color:var(--ink);border-radius:999px;padding:7px 16px}\n.seg button[aria-pressed=\"true\"]{background:var(--accent);border-color:var(--accent);color:var(--accent-ink);font-weight:600}\n.seg button .c{font-size:12px;opacity:.8;margin-inline-start:6px;font-variant-numeric:tabular-nums}\n\n/* section chips */\n.chips{display:flex;flex-wrap:wrap;gap:6px}\n.chip{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--line);background:var(--bg);border-radius:8px;padding:5px 10px;font-size:14px;cursor:pointer;user-select:none}\n.chip input{margin:0;accent-color:var(--accent)}\n.chip.on{border-color:var(--accent);background:var(--accent-soft)}\n.chip .c{color:var(--muted);font-size:12px;font-variant-numeric:tabular-nums}\n\n/* upload slot */\n.slot{border:2px dashed var(--line);border-radius:14px;padding:26px 16px;text-align:center;display:grid;gap:6px;cursor:pointer;background:var(--card)}\n.slot:hover,.slot.over{border-color:var(--accent);background:var(--accent-soft)}\n.slot .t{font-size:18px;font-weight:700}\n.slot .f{font-size:13px;color:var(--muted);overflow-wrap:anywhere}\n.slot .f b{color:var(--ink)}\n.slot input{display:none}\n\n/* actions */\n.actions{display:flex;flex-wrap:wrap;gap:8px}\n.btn{border:1px solid var(--line);background:var(--card);color:var(--ink);border-radius:10px;padding:9px 16px;font-weight:600;display:inline-flex;gap:8px;align-items:center}\n.btn .x{font-size:11px;font-weight:700;border-radius:4px;padding:1px 5px;color:#fff}\n.btn:disabled{opacity:.45;cursor:not-allowed}\n.gwrap{display:grid;gap:8px;padding-bottom:12px;border-bottom:1px solid var(--line)}\n.btn.sm{padding:6px 12px;font-size:13px}\n.gact{gap:6px}\n.gwrap .ghead{display:flex;flex-wrap:wrap;gap:6px 12px;align-items:baseline;justify-content:space-between}\n.gwrap .ghead b{font-size:14px}\n.gstats{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}\n.gs{border:1px solid var(--line);border-top:4px solid var(--pc);border-radius:12px;padding:12px 14px;display:grid;gap:8px;background:var(--card);min-width:0}\n.gs .gt{display:flex;justify-content:space-between;gap:8px;font-size:13.5px;color:var(--muted);font-weight:600}\n.gs .gtag{font-size:11.5px;font-weight:500;padding:1px 8px;border-radius:99px;background:var(--miss-soft);color:var(--miss);white-space:nowrap}\n.gs .gv{font-family:var(--display,inherit);font-size:30px;font-weight:600;line-height:1.1;color:var(--pc);font-variant-numeric:tabular-nums}\n.gs .gv small{font-size:15px;color:var(--muted);font-weight:500}\n.gs .gduo{display:flex;height:8px;border-radius:4px;overflow:hidden;background:var(--line)}\n.gs .gduo i{display:block;height:100%}\n.gs .gpr{display:grid;grid-template-columns:1fr 1fr;gap:6px}\n.gs .gpr.one{grid-template-columns:1fr}\n.gs .gpr div{border-radius:8px;padding:6px 8px;display:grid;font-size:12px;color:var(--muted)}\n.gs .gpr b{font-family:var(--display,inherit);font-size:18px;font-variant-numeric:tabular-nums}\n.gs .gpr .y{background:var(--green-soft,#e0f1e6)} .gs .gpr .y b{color:var(--green,#26794a)}\n.gs .gpr .n{background:var(--red-soft,#f9e3e0)} .gs .gpr .n b{color:var(--red,#b3372f)}\n.gs .gcnt{display:flex;justify-content:space-between;font-size:12px;color:var(--muted);font-variant-numeric:tabular-nums}\n@media (prefers-reduced-motion:no-preference){.gs .gduo i{transition:width .5s ease}}\n.x.pdf{background:#B42318}.x.doc{background:#2B579A}.x.xls{background:#217346}.x.prn{background:#5B6770}\n.check{display:inline-flex;gap:6px;align-items:center;font-size:14px}\n\n/* report preview */\n.report{display:grid;gap:14px}\n#repBody{display:grid;grid-template-columns:minmax(0,1fr);gap:18px}\n.report{grid-template-columns:minmax(0,1fr)}\n.grp{min-width:0}\n.rep-head{display:flex;flex-wrap:wrap;justify-content:space-between;gap:8px;align-items:end;border-bottom:2px solid var(--accent);padding-bottom:10px}\n.rep-head .school{font-weight:700}\n.rep-head .meta{font-size:13px;color:var(--muted);text-align:left}\n.total{font-size:14px}\n.total b{color:var(--miss);font-size:18px;font-variant-numeric:tabular-nums}\n.grp h3{margin:0 0 6px;font-size:15px;display:flex;gap:8px;align-items:center}\n.grp h3 .n{font-size:12px;font-weight:600;background:var(--miss-soft);color:var(--miss);border-radius:999px;padding:1px 10px}\n.scroll{overflow-x:auto}\ntable{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums}\nth,td{border:1px solid var(--line);padding:6px 8px;text-align:right;white-space:nowrap}\ntd.name{white-space:normal;min-width:180px}\nth{background:var(--accent-soft);font-weight:600;font-size:13px}\ntd.num{text-align:center;width:44px;color:var(--muted)}\ntd.ltr{direction:ltr;text-align:right}\n.empty{color:var(--muted);text-align:center;padding:28px 12px}\n.err{color:var(--miss);font-size:14px;margin:0}\n.toast{position:fixed;bottom:calc(18px + env(safe-area-inset-bottom,0px));left:50%;transform:translateX(-50%);background:var(--ink);color:var(--bg);padding:9px 16px;border-radius:8px;font-size:14px;z-index:30}\n.busy{position:fixed;inset:0;background:rgba(0,0,0,.35);display:grid;place-items:center;z-index:40;color:#fff;font-size:16px}\n\n@media (prefers-reduced-motion:no-preference){.slot,.chip,.seg button{transition:background .15s,border-color .15s}}\n[hidden]{display:none!important}\n"+"</style>"+"\n<div class=\"wrap\">\n  <header class=\"row\" style=\"justify-content:space-between;align-items:end\">\n    <div>\n      <h2 style=\"font-size:20px\">الطلاب الذين لم يدخلوا المنصة</h2>\n      <div class=\"sub\" id=\"todayTxt\"></div>\n    </div>\n  </header>\n\n  <!-- 1) Classification before upload -->\n  <section class=\"card\" aria-labelledby=\"fTitle\">\n    <h2 id=\"fTitle\">التصنيف</h2>\n    <div class=\"seg\" id=\"gradeSeg\" role=\"group\" aria-label=\"الصف\">\n      <button type=\"button\" data-g=\"\" aria-pressed=\"true\">الكل<span class=\"c\" id=\"gc-\"></span></button>\n      <button type=\"button\" data-g=\"1\" aria-pressed=\"false\">الأول ثانوي<span class=\"c\" id=\"gc-1\"></span></button>\n      <button type=\"button\" data-g=\"2\" aria-pressed=\"false\">الثاني ثانوي<span class=\"c\" id=\"gc-2\"></span></button>\n      <button type=\"button\" data-g=\"3\" aria-pressed=\"false\">الثالث ثانوي<span class=\"c\" id=\"gc-3\"></span></button>\n    </div>\n    <div class=\"row\" style=\"justify-content:space-between\">\n      <span class=\"note\">الشعب الدراسية</span>\n      <span class=\"row\" style=\"gap:6px\">\n        <button type=\"button\" class=\"btn\" id=\"secAll\" style=\"padding:4px 10px;font-weight:500;font-size:13px\">تحديد الكل</button>\n        <button type=\"button\" class=\"btn\" id=\"secNone\" style=\"padding:4px 10px;font-weight:500;font-size:13px\">إلغاء التحديد</button>\n      </span>\n    </div>\n    <div class=\"chips\" id=\"secChips\"><span class=\"note\">تظهر الشعب هنا بعد قراءة الملف.</span></div>\n  </section>\n\n  <!-- 2) The single upload slot -->\n  <label class=\"slot\" id=\"slot\" for=\"inFile\">\n    <span class=\"t\">ملف الطلاب الذين لم يدخلوا المنصة</span>\n    <span class=\"note\">ضع الملف هنا أو اضغط للاختيار · Excel أو Word أو PDF</span>\n    <span class=\"f\" id=\"fileTxt\">لم يُرفع ملف</span>\n    <input type=\"file\" id=\"inFile\" accept=\".xlsx,.xls,.xlsm,.csv,.docx,.pdf\">\n  </label>\n  <p class=\"err\" id=\"err\" hidden></p>\n\n  <!-- 3) Save / print -->\n  <section class=\"card\" aria-labelledby=\"aTitle\">\n    <div class=\"row\" style=\"justify-content:space-between\">\n      <h2 id=\"aTitle\">الحفظ والطباعة</h2>\n      <label class=\"check\" for=\"optSplit\"><input type=\"checkbox\" id=\"optSplit\" checked> جدول مستقل لكل شعبة</label>\n    </div>\n    <div class=\"gwrap\" id=\"gWrap\" hidden>\n      <div class=\"ghead\"><b>نسبة عدم الدخول حسب الصف</b><span class=\"note\" id=\"gNote\"></span></div>\n      <div class=\"gstats\" id=\"gStats\"></div>\n      <div class=\"actions gact\">\n        <button class=\"btn sm\" id=\"gPdf\" type=\"button\"><span class=\"x pdf\">PDF</span>حفظ الإحصائية PDF</button>\n        <button class=\"btn sm\" id=\"gDoc\" type=\"button\"><span class=\"x doc\">W</span>Word</button>\n        <button class=\"btn sm\" id=\"gXls\" type=\"button\"><span class=\"x xls\">X</span>Excel</button>\n        <button class=\"btn sm\" id=\"gPrn\" type=\"button\"><span class=\"x prn\">⎙</span>طباعة</button>\n      </div>\n    </div>\n    <div class=\"row\" style=\"gap:6px 14px\">\n      <span class=\"note\">البيانات الظاهرة:</span>\n      <span class=\"chips\" id=\"colPick\"><span class=\"note\">تظهر بعد قراءة الملف</span></span>\n    </div>\n    <div class=\"row\">\n      <label class=\"field\" for=\"school\">اسم المدرسة <input id=\"school\" value=\"ثانوية بلاط الشهداء بجدة\" style=\"width:240px\"></label>\n      <label class=\"field\" for=\"day\">التاريخ <input type=\"date\" id=\"day\"></label>\n    </div>\n    <div class=\"actions\">\n      <button class=\"btn\" id=\"bPdf\" type=\"button\" disabled><span class=\"x pdf\">PDF</span>حفظ PDF</button>\n      <button class=\"btn\" id=\"bDoc\" type=\"button\" disabled><span class=\"x doc\">W</span>حفظ Word</button>\n      <button class=\"btn\" id=\"bXls\" type=\"button\" disabled><span class=\"x xls\">X</span>حفظ Excel</button>\n      <button class=\"btn\" id=\"bPrn\" type=\"button\" disabled><span class=\"x prn\">⎙</span>طباعة</button>\n    </div>\n  </section>\n\n  <!-- 4) Result -->\n  <section class=\"card report\" id=\"report\" aria-live=\"polite\">\n    <div class=\"empty\" id=\"emptyBox\">\n      ضع ملف الطلاب الذين لم يدخلوا المنصة في الخانة أعلاه، وستظهر أسماؤهم هنا حسب الصف والشعبة.\n      <div style=\"margin-top:10px\"><button type=\"button\" class=\"btn\" id=\"bSample\" style=\"font-weight:500\">عرض مثال</button></div>\n    </div>\n    <div id=\"repBody\" hidden></div>\n  </section>\n\n  <p class=\"note\">يُقرأ الملف داخل جهازك فقط ولا يُرفع إلى أي مكان.</p>\n</div>\n<div class=\"toast\" id=\"toast\" hidden></div>\n<div class=\"busy\" id=\"busy\" hidden>جارٍ تجهيز الملف…</div>\n\n";
var gst=document.createElement("style"); gst.textContent="/* A4 pages used for PDF and printing */\n#nl-pages{position:absolute;top:0;left:-20000px}\n#nl-pages .page{width:794px;height:1123px;padding:44px 46px 56px;background:#fff;color:#111;position:relative;overflow:hidden;font-family:\"IBM Plex Sans Arabic\",Tahoma,\"Segoe UI\",Arial,sans-serif;font-size:14px;line-height:1.5;direction:rtl}\n#nl-pages .page .ph{display:flex;justify-content:space-between;align-items:end;border-bottom:2px solid #0F6B63;padding-bottom:8px;margin-bottom:12px}\n#nl-pages .page .ph .s{font-weight:700;font-size:15px}\n#nl-pages .page .ph .m{font-size:12px;color:#444;text-align:left}\n#nl-pages .page h1{font-size:20px;text-align:center;margin:4px 0 2px}\n#nl-pages .page .flt{text-align:center;font-size:13px;color:#333;margin-bottom:12px}\n#nl-pages .page h3{font-size:14px;margin:12px 0 4px;color:#0F6B63}\n#nl-pages .page table{width:100%;border-collapse:collapse;table-layout:fixed}\n#nl-pages .page td{overflow-wrap:anywhere}\n#nl-pages .page td img{max-width:100%;height:auto}\n#nl-pages .page th,#nl-pages .page td{border:1px solid #9AA5A3;padding:4px 6px;font-size:13px;text-align:right}\n#nl-pages .page th{background:#E2EFEC}\n#nl-pages .page td.num{text-align:center;color:#333}\n#nl-pages.dense .page th,#nl-pages.dense .page td{font-size:11px;padding:3px 4px}\n#nl-pages .page .pf{position:absolute;bottom:22px;left:46px;right:46px;display:flex;justify-content:space-between;font-size:11px;color:#555;border-top:1px solid #ccc;padding-top:6px}\n\n@media print{\n  @page{size:A4;margin:0}\n  body.nl-print{background:#fff!important;padding:0!important;margin:0!important}\n  body.nl-print>*:not(#nl-pages){display:none!important}\n  body.nl-print #nl-pages{position:static;left:auto}\n  #nl-pages .page{page-break-after:always;break-after:page}\n  #nl-pages .page:last-child{page-break-after:auto;break-after:auto}\n}\n"; document.head.appendChild(gst);
var PAGES=document.createElement("div"); PAGES.id="nl-pages"; PAGES.setAttribute("aria-hidden","true"); document.body.appendChild(PAGES);

"use strict";
var $=function(id){return R.getElementById(id);};
var GL={"1":"الأول ثانوي","2":"الثاني ثانوي","3":"الثالث ثانوي"};
var OPT=[{k:"id",t:"رقم الهوية"},{k:"phone",t:"الجوال"},{k:"user",t:"اسم المستخدم"},{k:"last",t:"آخر دخول"}];
var LTR={id:1,phone:1,user:1,last:1};
var lastInfo=null, students=[], fileName="", grade="", selSecs=new Set(), has={}, show={};

/* ---------- helpers ---------- */
function latin(s){return String(s).replace(/[٠-٩]/g,function(c){return String(c.charCodeAt(0)-0x0660);}).replace(/[۰-۹]/g,function(c){return String(c.charCodeAt(0)-0x06F0);});}
function str(v){return v==null?"":String(v).replace(/[\u200e\u200f\u202a-\u202e\u2066-\u2069]/g,"").replace(/\s+/g," ").trim();}
function norm(s){return latin(str(s)).replace(/[ً-ْـ]/g,"").replace(/[أإآٱ]/g,"ا").replace(/ة/g,"ه").replace(/ى/g,"ي");}
function esc(s){return String(s).replace(/[&<>"']/g,function(c){return{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c];});}
function xesc(s){return String(s).replace(/[&<>"]/g,function(c){return{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c];});}
function ymd(d){return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");}
function toast(m){var t=$("toast");t.textContent=m;t.hidden=false;clearTimeout(toast.t);toast.t=setTimeout(function(){t.hidden=true;},2200);}
function showErr(m){var e=$("err");e.textContent=m||"";e.hidden=!m;}
function dayDate(){return new Date(($("day").value||ymd(new Date()))+"T12:00:00");}
function hijri(d){try{return d.toLocaleDateString("ar-SA-u-ca-islamic-umalqura",{weekday:"long",day:"numeric",month:"long",year:"numeric"});}catch(e){return "";}}
function greg(d){return d.toLocaleDateString("ar-SA-u-ca-gregory-nu-latn",{day:"numeric",month:"long",year:"numeric"});}

/* ---------- class / section parsing ---------- */
function gradeOf(text){
  var t=norm(text); if(!t) return "";
  var m=t.match(/^\s*([1-3])\s*[\/\-\\]\s*\d{1,2}\s*$/); if(m) return m[1];
  var w=t.replace(/(?:فصل|شعب[هة]?|(?:^|\s)ش\.?)\s*[:\-]?\s*\d{1,2}/g," ").replace(/الثانوي[هة]?|ثانوي[هة]?|المرحله|مرحله|الصف|صف|الفصل|فصل/g," ");
  if(/الاول|اول|first/i.test(w)) return "1";
  if(/الثاني|ثاني|second/i.test(w)) return "2";
  if(/الثالث|ثالث|third/i.test(w)) return "3";
  m=w.match(/(?:^|\D)([1-3])(?:\D|$)/); if(m && w.replace(/\D/g,"").length===1) return m[1];
  return "";
}
function sectionOf(text,isSectionCol){
  var t=norm(text); if(!t) return "";
  var m=t.match(/^\s*[1-3]\s*[\/\-\\]\s*(\d{1,2})\s*$/); if(m) return m[1];
  m=t.match(/شعب[هة]?\s*[:\-]?\s*(\d{1,2}|[ا-ي])/); if(m) return m[1];
  m=t.match(/فصل\s*[:\-]?\s*(\d{1,2})/); if(m) return m[1];
  m=t.match(/(?:^|\s)ش\.?\s*[:\-]?\s*(\d{1,2})/); if(m) return m[1];
  if(isSectionCol){ m=t.match(/(\d{1,2})/); return m?m[1]:str(text); }
  m=t.match(/(?:ثانوي|ثانويه|الثانوي)\s*[\/\-]?\s*(\d{1,2})\s*$/); if(m) return m[1];
  m=t.match(/[\/\-]\s*(\d{1,2})\s*$/); if(m) return m[1];
  return "";
}

/* ---------- header detection ---------- */
function roleOf(cell){
  var c=norm(cell); if(!c||c.length>30||/\d|@/.test(c)) return "";
  if(/شعب/.test(c)) return "section";
  if(/مستخدم|username|user ?name|login ?id/i.test(c)) return "user";
  if(/دخول|login|last ?(seen|access)/i.test(c)) return "last";
  if(/(^|\s)(ال)?اسم(\s|$)/.test(c)&&!/ولي|الاب|المعلم|المدرس/.test(c)) return "name";
  if(/هويه|هوية|سجل|مدني|اقامه|رقم الطالب|رقم الهويه/.test(c)) return "id";
  if(/جوال|هاتف|موبايل|تواصل|ولي الامر/.test(c)) return "phone";
  if(/فصل/.test(c)&&!/دراسي/.test(c)) return "fasl";
  if(/صف|مرحله|grade|class/i.test(c)) return "grade";
  if(/^name$/i.test(c)) return "name";
  return "";
}
function headerOf(row){
  var map={},n=0;
  row.forEach(function(c,i){var r=roleOf(c); if(r&&map[r]==null){map[r]=i;n++;}});
  if(map.fasl!=null){ if(map.grade!=null&&map.section==null) map.section=map.fasl; else if(map.grade==null) map.grade=map.fasl; delete map.fasl; }
  if(map.name!=null&&n<2&&row.some(function(c){return /^(م|#|ت|ر|الرقم|التسلسل|مسلسل|تسلسل)$/.test(str(c));})) n=2;
  return (n>=2&&map.name!=null)?map:null;
}

/* ---------- extraction from generic rows ---------- */
function extract(rows){
  var out=[],head=null,ctxG="",ctxS="",seen=new Set();
  rows.forEach(function(raw){
    if(raw&&raw.__ctx!=null){ var g0=gradeOf(raw.__ctx),s0=sectionOf(raw.__ctx,false); if(g0){ctxG=g0;} if(s0) ctxS=s0; return; }
    var row=raw.map(str), filled=row.filter(Boolean);
    if(!filled.length) return;
    var h=headerOf(row); if(h){head=h;return;}
    if(filled.length<=2&&/صف|فصل|شعب|ثانوي/.test(norm(filled.join(" ")))&&!(head&&row[head.name]&&filled.length===2)){
      var j=filled.join(" "), g1=gradeOf(j), s1=sectionOf(j,false);
      if(g1||s1){ if(g1) ctxG=g1; ctxS=s1||ctxS; return; }
    }
    if(!head) return;
    var name=row[head.name]; if(!name||/^\d+$/.test(latin(name))||roleOf(name)==="name") return;
    var gTxt=head.grade!=null?row[head.grade]:"", sTxt=head.section!=null?row[head.section]:"";
    var g=gradeOf(gTxt)||(gTxt?"":ctxG)||ctxG, s=sTxt?sectionOf(sTxt,true):(sectionOf(gTxt,false)||ctxS);
    var id=head.id!=null?latin(row[head.id]):"", ph=head.phone!=null?latin(row[head.phone]):"", us=head.user!=null?str(row[head.user]):"", la=head.last!=null?latin(row[head.last]):"";
    var k=(id&&id.replace(/\D/g,"").length>5?"i"+id.replace(/\D/g,""):"n")+"|"+norm(name)+"|"+g+"|"+s;
    if(seen.has(k)) return; seen.add(k);
    out.push({name:name,id:id,phone:ph,user:us,last:la,g:g,gRaw:gTxt,s:s||"—"});
  });
  return out;
}

/* ---------- readers ---------- */
function readExcel(buf,isCsv){
  var wb;
  if(isCsv){ var t=new TextDecoder("utf-8").decode(buf); if(t.indexOf("�")>=0){try{t=new TextDecoder("windows-1256").decode(buf);}catch(e){}}
    wb=XLSX.read(t.replace(/^﻿/,""),{type:"string"}); }
  else wb=XLSX.read(buf,{type:"array"});
  var rows=[];
  wb.SheetNames.forEach(function(sn){
    rows.push({__ctx:sn});
    XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:"",raw:false,blankrows:false}).forEach(function(r){rows.push(r);});
  });
  return rows;
}
function readDocx(buf){
  return JSZip.loadAsync(buf).then(function(z){
    var f=z.file("word/document.xml"); if(!f) throw new Error("docx");
    return f.async("string");
  }).then(function(xml){
    var doc=new DOMParser().parseFromString(xml,"application/xml");
    var W="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var body=doc.getElementsByTagNameNS(W,"body")[0], rows=[];
    function text(el){var ts=el.getElementsByTagNameNS(W,"t"),s="";for(var i=0;i<ts.length;i++)s+=ts[i].textContent;return s;}
    function walk(node){
      for(var c=node.firstChild;c;c=c.nextSibling){
        if(c.nodeType!==1) continue;
        if(c.localName==="p"){var t=str(text(c)); if(t) rows.push({__ctx:t});}
        else if(c.localName==="tbl"){
          for(var r=c.firstChild;r;r=r.nextSibling){
            if(r.nodeType!==1||r.localName!=="tr") continue;
            var cells=[];
            for(var tc=r.firstChild;tc;tc=tc.nextSibling){
              if(tc.nodeType!==1||tc.localName!=="tc") continue;
              var ps=tc.getElementsByTagNameNS(W,"p"),parts=[];
              for(var k=0;k<ps.length;k++){var pt=str(text(ps[k]));if(pt)parts.push(pt);}
              cells.push(parts.join(" "));
              var span=tc.getElementsByTagNameNS(W,"gridSpan")[0];
              if(span){for(var q=1;q<+span.getAttributeNS(W,"val");q++)cells.push("");}
            }
            rows.push(cells);
          }
        } else if(c.localName==="sdt"||c.localName==="sdtContent"){walk(c);}
      }
    }
    walk(body);
    return rows;
  });
}
/* ---------- last-login interpretation ---------- */
var HM=[["جمادي الاولي",5],["جمادي الاول",5],["جمادي الاخره",6],["جمادي الاخر",6],["جمادي الثانيه",6],["جمادي الثاني",6],["ربيع الاول",3],["ربيع الاخر",4],["ربيع الثاني",4],
  ["ذو القعده",11],["ذي القعده",11],["ذو الحجه",12],["ذي الحجه",12],["محرم",1],["صفر",2],["رجب",7],["شعبان",8],["رمضان",9],["شوال",10]];
function loginKey(v){
  var t=norm(v); if(!t) return null;
  var m, y, mo, d, cal="G";
  for(var i=0;i<HM.length;i++){ if(t.indexOf(HM[i][0])>=0){ mo=HM[i][1]; var nums=(t.replace(HM[i][0]," ").match(/\d+/g)||[]).map(Number);
      y=nums.filter(function(n){return n>1300;})[0]; d=nums.filter(function(n){return n>=1&&n<=30;})[0]; cal="H"; break; } }
  if(!mo){
    if((m=t.match(/(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/))){y=+m[1];mo=+m[2];d=+m[3];}
    else if((m=t.match(/(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/))){d=+m[1];mo=+m[2];y=+m[3];}
    if(y&&y<1600) cal="H";
  }
  if(!y||!mo||!d) return null;
  return cal+y+"-"+String(mo).padStart(2,"0")+"-"+String(d).padStart(2,"0");
}
function isNeg(v){var t=norm(v);return !t||/لم|غير|لا يوجد|لايوجد|بدون|never|^no$|false|^[-–—0.\s]+$/i.test(t);}
var HMN=["محرم","صفر","ربيع الأول","ربيع الآخر","جمادى الأولى","جمادى الآخرة","رجب","شعبان","رمضان","شوال","ذو القعدة","ذو الحجة"];
function hijriToGreg(y,m,d){
  try{var f=new Intl.DateTimeFormat("en-u-ca-islamic-umalqura-nu-latn",{year:"numeric",month:"numeric",day:"numeric"});
    var approx=new Date(Math.round((y-1)*354.367+(m-1)*29.53+d)*86400000-42521587200000+227015*86400000*0);
    var base=new Date(622,6,16); approx=new Date(base.getTime()+((y-1)*354.367+(m-1)*29.53+(d-1))*86400000);
    for(var k=-40;k<=40;k++){var c=new Date(approx.getTime()+k*86400000),o={};f.formatToParts(c).forEach(function(p){o[p.type]=+p.value;});
      if(o.year===y&&o.month===m&&o.day===d) return ymd(c);}
  }catch(e){}
  return "";
}
function applyLogin(all){
  if(!all.some(function(x){return x.last;})) return {list:all,full:false,total:all.length};
  var keys=all.map(function(x){return isNeg(x.last)?null:loginKey(x.last);});
  var dated=keys.filter(Boolean);
  if(!dated.length){ // no parsable dates: anything non-negative counts as logged in
    return {list:all.filter(function(x){return isNeg(x.last);}),full:true,total:all.length};
  }
  var cal=dated.filter(function(k){return k[0]==="H";}).length>=dated.length/2?"H":"G";
  var ref=dated.filter(function(k){return k[0]===cal;}).sort().pop();
  var p=ref.slice(1).split("-").map(Number), refTxt, refGreg;
  if(cal==="H"){refTxt=p[2]+" "+HMN[p[1]-1]+" "+p[0]+"هـ"; refGreg=hijriToGreg(p[0],p[1],p[2]);}
  else {refGreg=ref.slice(1); refTxt=refGreg;}
  var list=all.filter(function(x,i){var k=keys[i]; return !k||k<ref;});
  return {list:list,full:true,total:all.length,refTxt:refTxt,refGreg:refGreg};
}
/* platform-list rows from a PDF read by Core.makeExtractor (glyph runs with x, x2, y; Arabic in visual order) */
function nlRowsFromPdfPages(pages){
  var AR=/[؀-ۿﭐ-﷿ﹰ-﻿]/;
  function logical(gl){ // visual glyph run (left->right) -> logical text, RTL paragraph
    var segs=[]; gl.forEach(function(g){ var t=AR.test(g.str)?"a":(/^\s+$/.test(g.str)?"s":"l"); var L=segs[segs.length-1]; if(L&&L.t===t) L.g.push(g); else segs.push({t:t,g:[g]}); });
    for(var i=1;i<segs.length-1;i++){ if(segs[i].t==="s"&&segs[i-1].t==="l"&&segs[i+1].t==="l"){ segs[i-1].g=segs[i-1].g.concat(segs[i].g,segs[i+1].g); segs.splice(i,2); i--; } }
    if(!segs.some(function(s){return s.t==="a";})) return [gl.map(function(g){return g.str;}).join("")];
    return segs.reverse().map(function(s){ return (s.t==="a"?s.g.slice().reverse():s.g).map(function(g){return g.str;}).join(""); });
  }
  var MON=/محرم|صفر|ربيع|جماد|رجب|شعبان|رمضان|شوال|القعد|الحج/;
  var out=[["م","رقم الهوية","الاسم","اسم المستخدم","الجوال","الفصل","آخر دخول"]], seen=0;
  pages.forEach(function(pg){
    var runs=(pg.runs||[]).filter(function(r){return r.str!=null&&r.str!=="";});
    var rows=[]; runs.forEach(function(r){ var R=rows.find(function(x){return Math.abs(x.y-r.y)<1.2;}); if(R) R.r.push(r); else rows.push({y:r.y,r:[r]}); });
    rows.sort(function(a,b){return a.y-b.y;}).forEach(function(R){
      // drawing order -> chunks: a new chunk when x jumps back or a visible gap opens
      var chunks=[], cur=[];
      R.r.forEach(function(g){ var p=cur[cur.length-1];
        if(p&&(g.x<p.x-0.5||(g.x-(p.x2!=null?p.x2:p.x))>2.6)){ chunks.push(cur); cur=[]; }
        cur.push(g); });
      if(cur.length) chunks.push(cur);
      var pieces=[]; chunks.forEach(function(c){ var t=logical(c).join(" ").replace(/\s+/g," ").trim(); if(t) pieces.push({t:t,x:c[0].x}); });
      var all=pieces.map(function(p){return p.t;}).join(" | ");
      var idm=all.match(/(?:^|\D)(\d{10})(?:\D|$)/)||(pieces.length>3&&/^\d{1,4}$/.test(pieces[0].t)&&/^\d{5,10}$/.test(pieces[1].t)?[0,pieces[1].t]:null); if(!idm) return;
      var rec={serial:"",id:idm[1],name:[],user:"",phone:"",grade:"",sec:"",last:""}, nums=[];
      pieces.forEach(function(p){
        var t=p.t;
        if(t.indexOf("@")>=0){ var em=t.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+/); if(em){ rec.user=em[0]; t=t.replace(em[0]," "); } }
        if(/لم\s*يتم\s*الدخول|لم\s*يدخل/.test(t)){ rec.last="لم يتم الدخول"; t=t.replace(/لم\s*يتم\s*الدخول|لم\s*يدخل/," "); }
        if(MON.test(t)&&/\d{3,4}/.test(t)&&t.replace(/[\d\/\s.\-]/g,"").length<=16){ rec.last=t.replace(/\s+/g," ").trim(); return; }
        var gm=t.match(/(أول|اول|ثاني|ثالث)\s*ش?/); if(gm&&/ش/.test(t)){ rec.grade=gm[1]; t=t.replace(gm[0]," ").replace(/ش/," "); }
        (t.match(/\d+/g)||[]).forEach(function(d){ nums.push({d:d,x:p.x}); });
        t=t.replace(/[\d|]+/g," ").replace(/\s+/g," ").trim();
        if(t&&AR.test(t)&&!/^ش$/.test(t)) rec.name.push(t);
      });
      var longs=[];
      nums.forEach(function(n){
        var d=n.d;
        if(d===rec.id) return;
        if(d.length>=9){ longs.push(d); return; }
        if(d.length<=4&&!rec.serial&&n.x<pieces[0].x+1){ rec.serial=d; return; }
        if(d.length<=2&&!rec.sec){ rec.sec=d; return; }
        if(d.length<=4&&!rec.serial) rec.serial=d;
      });
      longs.forEach(function(d){
        var ph=d.match(/^(9665\d{8}|05\d{8})(\d{1,2})$/);
        if(ph&&!rec.sec&&rec.grade){ rec.phone=rec.phone||ph[1]; rec.sec=ph[2]; return; }
        if(!rec.phone) rec.phone=d;
      });
      seen++;
      out.push([rec.serial||String(seen),rec.id,rec.name.join(" "),rec.user,rec.phone,rec.grade?(rec.grade+" ش "+rec.sec).trim():"",rec.last]);
    });
  });
  return seen?out:[];
}

function hasCore(){ return typeof Core!=="undefined"&&typeof PDFLib!=="undefined"; }
function readPdfBuf(buf){
  if(!hasCore()) return Promise.reject(new Error("قراءة PDF متاحة داخل موقع التقرير فقط."));
  return Core.makeExtractor(PDFLib).extract(new Uint8Array(buf)).then(nlRowsFromPdfPages);
}
function looksPlatform(rows){
  // a platform list: rows with a username/email or a «last login» value
  var recs=rows.filter(function(r){return Array.isArray(r);}), hits=0, n=0;
  recs.forEach(function(r){ var t=r.join(" "); if(/\d{5,10}/.test(t)){ n++; if(/@|لم\s*يتم\s*الدخول|آخر\s*دخول/.test(t)) hits++; } });
  var hdr=recs.some(function(r){ var h=headerOf(r.map(str)); return h&&(h.user!=null||h.last!=null); });
  return (hdr&&n>0)||(n>=3&&hits>=n*0.5);
}
function loadRows(rows,name){
  var all=extract(rows);
  if(!all.length) throw new Error("لم أجد أسماء طلاب في «"+name+"». تأكد أن فيه عمودًا عنوانه «اسم الطالب» أو «الاسم».");
  var info=applyLogin(all);
  info.byGrade={}; all.forEach(function(x){ var k=x.g||""; (info.byGrade[k]=info.byGrade[k]||{total:0,missing:0}).total++; });
  info.list.forEach(function(x){ var k=x.g||""; (info.byGrade[k]=info.byGrade[k]||{total:0,missing:0}).missing++; });
  setData(info.list,name,info);
  var out={name:name,total:all.length,missing:info.list.length,full:info.full,refTxt:info.refTxt||""};
  try{document.dispatchEvent(new CustomEvent("nologin:data",{detail:out}));}catch(e){}
  return out;
}
function loadFileP(file){
  var n=file.name.toLowerCase();
  if(/\.doc$/.test(n)) return Promise.reject(new Error("صيغة Word القديمة (.doc) غير مدعومة. افتح الملف في Word واحفظه بصيغة .docx ثم ارفعه."));
  if(file.size>25*1024*1024) return Promise.reject(new Error("الملف «"+file.name+"» أكبر من 25 ميجابايت."));
  return file.arrayBuffer().then(function(buf){
    var b=new Uint8Array(buf.slice(0,4));
    var isPdf=/\.pdf$/.test(n)||(b[0]===0x25&&b[1]===0x50&&b[2]===0x44&&b[3]===0x46);
    if(isPdf) return readPdfBuf(buf);
    return /\.docx$/.test(n)?readDocx(buf):readExcel(buf,/\.csv$/.test(n));
  }).catch(function(e){throw new Error((e&&/[\u0600-\u06FF]/.test(e.message||""))?e.message:"تعذّرت قراءة «"+file.name+"». تأكد أنه Excel أو Word أو PDF سليم وغير محمي بكلمة مرور.");})
  .then(function(rows){ return loadRows(rows,file.name); });
}
function tryPages(pages,name){
  try{ var rows=nlRowsFromPdfPages(pages); if(rows.length>1&&looksPlatform(rows)) return Promise.resolve(loadRows(rows,name)); }catch(e){ console.error(e); }
  return Promise.resolve(null);
}
function tryDocx(bytes,name){
  return readDocx(bytes).then(function(rows){ return looksPlatform(rows)?loadRows(rows,name):null; }).catch(function(){return null;});
}
function handleFile(file){
  showErr("");
  loadFileP(file).then(function(inf){
    toast(inf.full?("في الملف "+inf.total+" طالب، لم يدخل منهم "+inf.missing):("تمت قراءة "+inf.total+" طالب"));
  }).catch(function(e){showErr((e&&e.message)||"تعذّرت قراءة الملف.");});
}

/* ---------- state & filters ---------- */
function secKey(x){return x.g+"|"+x.s;}
function secLabel(g,s){return (GL[g]?"الصف "+GL[g]:"بدون صف")+(s&&s!=="—"?" - الشعبة "+s:"");}
var coll=new Intl.Collator("ar",{numeric:true});
function sortStudents(a,b){return (a.g||"9").localeCompare(b.g||"9")||coll.compare(a.s,b.s)||coll.compare(a.name,b.name);}
function setData(list,name){
  students=list.sort(sortStudents); fileName=name;
  has={}; show={}; OPT.forEach(function(o){has[o.k]=students.some(function(x){return x[o.k];}); show[o.k]=has[o.k];});
  selSecs=new Set(students.map(secKey));
  var inf=arguments[2];
  lastInfo=inf||{full:false,total:list.length};
  $("fileTxt").innerHTML="<b>"+esc(name)+"</b> · "+(inf&&inf.full&&inf.total!==students.length?
    ("في الملف "+inf.total+" طالب · دخل "+(inf.total-students.length)+" · <b style=\"color:var(--miss)\">لم يدخل "+students.length+"</b>"+(inf.refTxt?" (حسب آخر دخول بتاريخ "+esc(inf.refTxt)+")":"")):
    ("<b style=\"color:var(--miss)\">"+students.length+" طالبًا لم يدخلوا المنصة</b>"));
  if(inf&&inf.refGreg) $("day").value=inf.refGreg;
  render();
}
function visible(){
  return students.filter(function(x){return (!grade||x.g===grade)&&selSecs.has(secKey(x));});
}
function renderFilters(){
  ["","1","2","3"].forEach(function(g){var n=students.filter(function(x){return !g||x.g===g;}).length; $("gc-"+g).textContent=students.length?n:"";});
  R.querySelectorAll("#gradeSeg button").forEach(function(b){b.setAttribute("aria-pressed",b.dataset.g===grade?"true":"false");});
  var box=$("secChips");
  if(!students.length){box.innerHTML='<span class="note">تظهر الشعب هنا بعد قراءة الملف.</span>';return;}
  var m=new Map();
  students.forEach(function(x){if(grade&&x.g!==grade)return;var k=secKey(x);m.set(k,(m.get(k)||0)+1);});
  var keys=Array.from(m.keys()).sort(function(a,b){var A=a.split("|"),B=b.split("|");return (A[0]||"9").localeCompare(B[0]||"9")||coll.compare(A[1],B[1]);});
  box.innerHTML=keys.map(function(k,i){var p=k.split("|"),on=selSecs.has(k);
    var lab=grade?(p[1]==="—"?"بدون شعبة":"الشعبة "+p[1]):secLabel(p[0],p[1]);
    return '<label class="chip'+(on?" on":"")+'" for="sc'+i+'"><input type="checkbox" id="sc'+i+'" data-k="'+esc(k)+'"'+(on?" checked":"")+'>'+esc(lab)+'<span class="c">'+m.get(k)+'</span></label>';}).join("");
}
$("secChips").addEventListener("change",function(e){var k=e.target.dataset.k;if(!k)return;if(e.target.checked)selSecs.add(k);else selSecs.delete(k);render();});
$("secAll").addEventListener("click",function(){students.forEach(function(x){if(!grade||x.g===grade)selSecs.add(secKey(x));});render();});
$("secNone").addEventListener("click",function(){students.forEach(function(x){if(!grade||x.g===grade)selSecs.delete(secKey(x));});render();});
R.querySelectorAll("#gradeSeg button").forEach(function(b){b.addEventListener("click",function(){grade=b.dataset.g;render();});});

function groups(list){
  if(!$("optSplit").checked) return [{title:"",items:list}];
  var m=new Map();
  list.forEach(function(x){var k=secKey(x);if(!m.has(k))m.set(k,{title:secLabel(x.g,x.s),items:[]});m.get(k).items.push(x);});
  return Array.from(m.values());
}
function filterText(){
  var g=grade?"الصف "+GL[grade]:"جميع الصفوف";
  var all=students.filter(function(x){return !grade||x.g===grade;}).map(secKey), uniqAll=Array.from(new Set(all));
  var chosen=uniqAll.filter(function(k){return selSecs.has(k);});
  var s=chosen.length===uniqAll.length?"جميع الشعب":chosen.map(function(k){var p=k.split("|");return grade?"الشعبة "+p[1]:secLabel(p[0],p[1]);}).join("، ");
  return g+" · "+s;
}
function cols(){
  var c=[{k:"#",t:"م"},{k:"name",t:"اسم الطالب"}];
  if(show.id)c.push({k:"id",t:"رقم الهوية"});
  c.push({k:"g",t:"الصف"},{k:"s",t:"الشعبة"});
  OPT.forEach(function(o){if(o.k!=="id"&&show[o.k])c.push(o);});
  return c;
}
function val(x,k,i){if(k==="#")return String(i+1);if(k==="g")return GL[x.g]||x.gRaw||"—";if(k==="phone"){var p=String(x.phone||"").replace(/\D/g,"");if(/^(00)?9660?5\d{8}$/.test(p))return "0"+p.slice(-9);if(/^5\d{8}$/.test(p))return "0"+p;return x.phone||"";}return x[k]||"";}

function renderColPick(){
  var box=$("colPick"); if(!students.length){box.innerHTML='<span class="note">تظهر بعد قراءة الملف</span>';return;}
  box.innerHTML=OPT.map(function(o){var dis=!has[o.k];return '<label class="chip'+(show[o.k]?" on":"")+'" for="cp-'+o.k+'" title="'+(dis?"غير موجود في الملف":"")+'" style="'+(dis?"opacity:.45;cursor:not-allowed":"")+'"><input type="checkbox" id="cp-'+o.k+'" data-c="'+o.k+'"'+(show[o.k]?" checked":"")+(dis?" disabled":"")+'>'+o.t+(dis?' <span class="c">غير موجود</span>':"")+'</label>';}).join("");
}
/* ---------- absence rate per grade ---------- */
function pctTxt(a,b){ if(!b) return "—"; var v=Math.round(a/b*1000)/10; return String(v).replace(/\.0$/,"")+"%"; }
function gradeStats(){
  if(!students.length) return null;
  var full=!!(lastInfo&&lastInfo.full&&lastInfo.total!==students.length&&lastInfo.byGrade);
  var miss={}; students.forEach(function(x){ var k=x.g||""; miss[k]=(miss[k]||0)+1; });
  var keys=["1","2","3"];
  var allMiss=students.length, allTot=full?lastInfo.total:null;
  var rows=keys.map(function(k){ var m=miss[k]||0, t=full?((lastInfo.byGrade[k]||{}).total||0):null;
    return {k:k,label:GL[k]?"الصف "+GL[k]:"بدون صف",missing:m,total:t,pct:full?pctTxt(m,t):pctTxt(m,allMiss),ratio:full?(t?m/t:0):(allMiss?m/allMiss:0)}; });
  return {full:full,rows:rows,total:{label:"إجمالي المدرسة",missing:allMiss,total:allTot,pct:full?pctTxt(allMiss,allTot):"100%",ratio:full?(allTot?allMiss/allTot:0):1}};
}
function renderGrades(){
  var w=$("gWrap"), gc=gradeCards(); if(!w) return;
  if(!gc){ w.hidden=true; return; } w.hidden=false;
  $("gNote").textContent=gc.full?"من إجمالي طلاب كل صف في الملف":"حصة كل صف من قائمة غير الداخلين (الملف لا يحتوي على إجمالي الطلاب)";
  $("gStats").innerHTML=gc.cards.map(function(c){
    var pin=c.of?Math.max(0,Math.min(100,c.inn/c.of*100)):0;
    return '<div class="gs'+(c.isTotal?" total":"")+'" style="--pc:'+c.pc+'">'+
      '<div class="gt"><span>'+esc(c.title)+'</span><span class="gtag">لم يدخلوا</span></div>'+
      '<div class="gv">'+c.v+(c.of!=null?' <small>من '+c.of+'</small>':' <small>طالبًا</small>')+'</div>'+
      (c.of!=null?'<div class="gduo" aria-hidden="true"><i style="width:'+pin.toFixed(1)+'%;background:var(--green,#26794a)"></i><i style="width:'+(100-pin).toFixed(1)+'%;background:var(--red,#b3372f)"></i></div>'+
        '<div class="gpr"><div class="y">نسبة الدخول<b><bdi>'+esc(c.pIn)+'</bdi></b></div><div class="n">نسبة عدم الدخول<b><bdi>'+esc(c.pOut)+'</bdi></b></div></div>'+
        '<div class="gcnt"><span>دخلوا: '+c.inn+'</span><span>لم يدخلوا: '+c.v+'</span></div>'
      :'<div class="gduo" aria-hidden="true"><i style="width:'+(c.ratio*100).toFixed(1)+'%;background:var(--red,#b3372f)"></i></div>'+
        '<div class="gpr one"><div class="n">حصته من قائمة الغائبين<b><bdi>'+esc(c.share)+'</bdi></b></div></div>')+
      '</div>';
  }).join("");
}
var GPC={"1":"#1d6f68","2":"#2f6fb0","3":"#7a4f9e"};
function gradeCards(){ var st=gradeStats(); if(!st) return null;
  var rows=st.rows.map(function(r){return Object.assign({pc:GPC[r.k]||"#1d6f68"},r);}).concat(st.full?[Object.assign({pc:"#b3372f",isTotal:true},st.total)]:[]);
  return {full:st.full,listTotal:students.length,cards:rows.map(function(r){
    var inn=r.total!=null?r.total-r.missing:null;
    return {title:r.label,pc:r.pc,isTotal:!!r.isTotal,v:r.missing,of:r.total,inn:inn,
      pIn:r.total!=null?pctTxt(inn,r.total):null,pOut:r.total!=null?pctTxt(r.missing,r.total):null,share:pctTxt(r.missing,students.length),ratio:r.ratio};})};
}
function gradeSummaryRows(){ var st=gradeStats(); if(!st) return null;
  var rows=st.rows.concat(st.full?[st.total]:[]);
  return {full:st.full,headers:st.full?["الصف","إجمالي الطلاب","دخلوا","لم يدخلوا","نسبة الدخول","نسبة عدم الدخول"]:["الصف","لم يدخلوا","النسبة من قائمة الغائبين"],
    rows:rows.map(function(r){ return st.full?[r.label,String(r.total),String(r.total-r.missing),String(r.missing),pctTxt(r.total-r.missing,r.total),r.pct]:[r.label,String(r.missing),r.pct]; })};
}
function render(){
  renderFilters(); renderColPick(); renderGrades();
  try{document.dispatchEvent(new CustomEvent("nologin:change"));}catch(e){}
  var list=visible(), has=students.length>0;
  $("emptyBox").hidden=has; $("repBody").hidden=!has;
  ["bPdf","bDoc","bXls","bPrn"].forEach(function(id){$(id).disabled=!list.length;});
  var d=dayDate();
  $("todayTxt").textContent=hijri(d)+" · "+greg(d)+"م";
  if(!has) return;
  var c=cols();
  var html='<div class="rep-head"><div><div class="school">'+esc($("school").value)+'</div><div class="note">'+esc(filterText())+'</div></div><div class="meta">'+esc(hijri(d))+'<br>'+esc(greg(d))+'م</div></div>'+
    '<div class="total">عدد الطلاب الذين لم يدخلوا المنصة: <b>'+list.length+'</b></div>';
  if(!list.length) html+='<div class="empty">لا يوجد طلاب في التصنيف المحدد. غيّر الصف أو حدّد شعبًا.</div>';
  groups(list).forEach(function(g){
    html+='<div class="grp">'+(g.title?'<h3>'+esc(g.title)+'<span class="n">'+g.items.length+'</span></h3>':'')+
      '<div class="scroll"><table><thead><tr>'+c.map(function(x){return"<th>"+x.t+"</th>";}).join("")+'</tr></thead><tbody>'+
      g.items.map(function(x,i){return "<tr>"+c.map(function(col){var v=esc(val(x,col.k,i));
        return col.k==="#"?'<td class="num">'+v+"</td>":col.k==="name"?'<td class="name">'+v+"</td>":LTR[col.k]?'<td class="ltr">'+v+"</td>":"<td>"+v+"</td>";}).join("")+"</tr>";}).join("")+
      '</tbody></table></div></div>';
  });
  $("repBody").innerHTML=html;
}
$("optSplit").addEventListener("change",render);
$("colPick").addEventListener("change",function(e){var k=e.target.dataset.c;if(!k)return;show[k]=e.target.checked;render();});
$("school").addEventListener("input",render);
$("day").addEventListener("change",render);

/* ---------- upload wiring ---------- */
var inp=$("inFile"),slot=$("slot");
inp.addEventListener("change",function(){var f=inp.files[0];if(f)handleFile(f);inp.value="";});
["dragenter","dragover"].forEach(function(ev){slot.addEventListener(ev,function(e){e.preventDefault();slot.classList.add("over");});});
["dragleave","drop"].forEach(function(ev){slot.addEventListener(ev,function(e){e.preventDefault();slot.classList.remove("over");});});
slot.addEventListener("drop",function(e){var f=e.dataTransfer.files[0];if(f)handleFile(f);});

/* ---------- export helpers ---------- */
function baseName(){var g=grade?GL[grade]:"الكل";return "لم-يدخلوا-المنصة-"+g.replace(/\s+/g,"-")+"-"+($("day").value||ymd(new Date()));}
function save(blob,name){var a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(a.href);a.remove();},1500);}
function busy(on){$("busy").hidden=!on;}

/* A4 pages (shared by PDF and print) */
function ltrImg(text){
  var fam=getComputedStyle(document.body).fontFamily, sc=3, cv=document.createElement("canvas"), cx=cv.getContext("2d");
  var fs=PAGES.classList.contains("dense")?11:13; cx.font=fs+"px "+fam; var w=Math.ceil(cx.measureText(text).width)+2;
  cv.width=w*sc; cv.height=18*sc; cx=cv.getContext("2d"); cx.scale(sc,sc);
  cx.font=fs+"px "+fam; cx.fillStyle="#111"; cx.textBaseline="middle"; cx.direction="ltr"; cx.fillText(text,1,9.5);
  return '<img alt="" src="'+cv.toDataURL("image/png")+'" style="height:18px;width:'+w+'px;vertical-align:middle">';
}
function gsTableHtml(){ var g=gradeSummaryRows(); if(!g) return "";
  return '<h3>نسبة عدم الدخول حسب الصف</h3><table class="gsum"><thead><tr>'+g.headers.map(function(h){return "<th>"+esc(h)+"</th>";}).join("")+'</tr></thead><tbody>'+
    g.rows.map(function(r,i){return '<tr'+(g.full&&i===g.rows.length-1?' style="font-weight:700"':'')+'>'+r.map(function(v){return '<td style="text-align:center">'+esc(v)+'</td>';}).join("")+'</tr>';}).join("")+'</tbody></table><div style="height:6px"></div>'; }
function buildPages(forPdf){
  var host=PAGES; host.innerHTML="";
  var list=visible(), c=cols(), d=dayDate(), pages=[];
  var LIMIT=1123-64;
  function newPage(first){
    var p=document.createElement("div"); p.className="page";
    p.innerHTML='<div class="ph"><div class="s">'+esc($("school").value)+'</div><div class="m">'+esc(hijri(d))+'<br>'+esc(greg(d))+'م</div></div>'+
      (first?'<h1>الطلاب الذين لم يدخلوا المنصة</h1><div class="flt">'+esc(filterText())+' · العدد: '+list.length+'</div>'+gsTableHtml():'')+
      '<div class="pf"><span>'+esc($("school").value)+'</span><span class="pn"></span></div>';
    host.appendChild(p); pages.push(p); return p;
  }
  function fits(p){var el=p.querySelector(".pf").previousElementSibling;return !el||(el.offsetTop+el.offsetHeight)<=LIMIT;}
  var page=newPage(true);
  var PW={"#":4,name:24,id:12,g:10,s:6,phone:12,user:22,last:12}, pt=c.reduce(function(a,x){return a+PW[x.k];},0);
  var thead='<colgroup>'+c.map(function(x){return '<col style="width:'+(PW[x.k]/pt*100).toFixed(2)+'%">';}).join("")+'</colgroup><thead><tr>'+c.map(function(x){return"<th>"+x.t+"</th>";}).join("")+'</tr></thead>';
  if(c.length>6) host.classList.add("dense"); else host.classList.remove("dense");
  groups(list).forEach(function(g){
    var title=null,tbody=null;
    function start(cont){
      if(g.title){title=document.createElement("h3");title.textContent=g.title+" ("+g.items.length+")"+(cont?" — تابع":"");page.insertBefore(title,page.querySelector(".pf"));}
      var t=document.createElement("table");t.innerHTML=thead+"<tbody></tbody>";page.insertBefore(t,page.querySelector(".pf"));tbody=t.querySelector("tbody");
      return t;
    }
    var table=start(false), rowsInTable=0;
    g.items.forEach(function(x,i){
      var tr=document.createElement("tr");
      tr.innerHTML=c.map(function(col){var v=esc(val(x,col.k,i));return col.k==="#"?'<td class="num">'+v+"</td>":LTR[col.k]?'<td style="direction:ltr;text-align:right;white-space:nowrap">'+(forPdf&&v&&!/[\u0600-\u06FF]/.test(val(x,col.k,i))?ltrImg(val(x,col.k,i)):v)+"</td>":"<td>"+v+"</td>";}).join("");
      tbody.appendChild(tr);
      if(!fits(page)){
        tbody.removeChild(tr);
        if(!rowsInTable){table.remove();if(title)title.remove();}
        page=newPage(false); table=start(rowsInTable>0); rowsInTable=0;
        tbody.appendChild(tr);
      }
      rowsInTable++;
    });
  });
  pages.forEach(function(p,i){p.querySelector(".pn").textContent="صفحة "+(i+1)+" من "+pages.length;});
  return pages;
}

/* PDF */
$("bPdf").addEventListener("click",function(){
  if(!(window.jspdf||window.jsPDF)||!window.html2canvas) return toast("مكتبة PDF غير متاحة");
  busy(true);
  setTimeout(function(){
    var pages=buildPages(true), pdf=new ((window.jspdf&&window.jspdf.jsPDF)||window.jsPDF)({unit:"pt",format:"a4",orientation:"portrait"}), W=pdf.internal.pageSize.getWidth(), H=pdf.internal.pageSize.getHeight();
    var chain=Promise.resolve();
    pages.forEach(function(p,i){
      chain=chain.then(function(){return html2canvas(p,{scale:2,backgroundColor:"#ffffff",useCORS:true,logging:false});})
        .then(function(cv){if(i)pdf.addPage();pdf.addImage(cv.toDataURL("image/jpeg",0.92),"JPEG",0,0,W,H);});
    });
    chain.then(function(){save(pdf.output("blob"),baseName()+".pdf");toast("تم حفظ PDF");})
      .catch(function(){toast("تعذّر إنشاء PDF");})
      .then(function(){busy(false);PAGES.innerHTML="";});
  },30);
});

/* Print */
$("bPrn").addEventListener("click",function(){buildPages();document.body.classList.add("nl-print");setTimeout(function(){window.print();},50);});
window.addEventListener("afterprint",function(){if(document.body.classList.contains("nl-print")){document.body.classList.remove("nl-print");PAGES.innerHTML="";}});

/* Excel */
$("bXls").addEventListener("click",function(){
  var list=visible(), c=cols(), wb=XLSX.utils.book_new();
  function sheet(items,title){
    var aoa=[[$("school").value],["الطلاب الذين لم يدخلوا المنصة"+(title?" — "+title:"")],[hijri(dayDate())+" · "+greg(dayDate())+"م"],[],c.map(function(x){return x.t;})];
    items.forEach(function(x,i){aoa.push(c.map(function(col){var v=val(x,col.k,i);return col.k==="#"?i+1:v;}));});
    var ws=XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"]=c.map(function(col){return{wch:col.k==="#"?5:col.k==="name"?34:LTR[col.k]?18:14};});
    ws["!merges"]=[0,1,2].map(function(r){return{s:{r:r,c:0},e:{r:r,c:c.length-1}};});
    ws["!autofilter"]={ref:XLSX.utils.encode_range({s:{r:4,c:0},e:{r:4+items.length,c:c.length-1}})};
    return ws;
  }
  var gsum=gradeSummaryRows();
  if(gsum){ var gws=XLSX.utils.aoa_to_sheet([[$("school").value],["نسبة عدم الدخول حسب الصف"],[hijri(dayDate())+" · "+greg(dayDate())+"م"],[],gsum.headers].concat(gsum.rows));
    gws["!cols"]=gsum.headers.map(function(h,i){return {wch:i?16:22};}); XLSX.utils.book_append_sheet(wb,gws,"ملخص الصفوف"); }
  XLSX.utils.book_append_sheet(wb,sheet(list,filterText()),"الكل");
  if($("optSplit").checked){
    var used={};
    groups(list).forEach(function(g){
      var x0=g.items[0]||{}, SH={"1":"الأول ث","2":"الثاني ث","3":"الثالث ث"};
      var n=((SH[x0.g]||"بدون صف")+(x0.s&&x0.s!=="—"?" - ش "+x0.s:"")).replace(/[\[\]:*?\/\\]/g,"-").slice(0,26)||"ورقة";
      var base=n,k=2; while(used[n]){n=base+" ("+(k++)+")";} used[n]=1;
      XLSX.utils.book_append_sheet(wb,sheet(g.items,g.title),n);
    });
  }
  wb.Workbook={Views:[{RTL:true}]};
  XLSX.writeFile(wb,baseName()+".xlsx");
  toast("تم حفظ Excel");
});

/* Word (.docx) */
function wRun(t,o){o=o||{};var sz=o.sz||24;
  return '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'+(o.b?'<w:b/><w:bCs/>':'')+(o.color?'<w:color w:val="'+o.color+'"/>':'')+'<w:sz w:val="'+sz+'"/><w:szCs w:val="'+sz+'"/><w:rtl/></w:rPr><w:t xml:space="preserve">'+xesc(t)+'</w:t></w:r>';}
function wP(t,o){o=o||{};return '<w:p><w:pPr><w:bidi/>'+(o.center?'<w:jc w:val="center"/>':'')+'<w:spacing w:before="'+(o.before||0)+'" w:after="'+(o.after==null?80:o.after)+'"/></w:pPr>'+wRun(t,o)+'</w:p>';}
function wTableRows(rows,c){
  var fake=rows.map(function(r){var o={};c.forEach(function(col,j){o["__"+col.k]=r[j];});return o;});
  return wTable(fake,c,true);
}
function wTable(items,c,pre){
  var widths={"#":500,name:3900,id:1700,g:1500,s:800,phone:1700,user:3300,last:1900}, total=c.reduce(function(a,x){return a+widths[x.k];},0), scale=(c.length>6?14838:9906)/total;
  var ws=c.map(function(x){return Math.round(widths[x.k]*scale);});
  var b='<w:top w:val="single" w:sz="4" w:color="9AA5A3"/><w:left w:val="single" w:sz="4" w:color="9AA5A3"/><w:bottom w:val="single" w:sz="4" w:color="9AA5A3"/><w:right w:val="single" w:sz="4" w:color="9AA5A3"/><w:insideH w:val="single" w:sz="4" w:color="9AA5A3"/><w:insideV w:val="single" w:sz="4" w:color="9AA5A3"/>';
  function cell(t,w,head,center){return '<w:tc><w:tcPr><w:tcW w:w="'+w+'" w:type="dxa"/>'+(head?'<w:shd w:val="clear" w:color="auto" w:fill="E2EFEC"/>':'')+'</w:tcPr><w:p><w:pPr><w:bidi/>'+(center?'<w:jc w:val="center"/>':'')+'<w:spacing w:before="20" w:after="20"/></w:pPr>'+wRun(t,{b:head,sz:c.length>6?19:22})+'</w:p></w:tc>';}
  var x='<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>'+b+'</w:tblBorders><w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>'+ws.map(function(w){return'<w:gridCol w:w="'+w+'"/>';}).join("")+'</w:tblGrid>';
  x+='<w:tr><w:trPr><w:tblHeader/></w:trPr>'+c.map(function(col,i){return cell(col.t,ws[i],true,true);}).join("")+'</w:tr>';
  items.forEach(function(it,r){x+='<w:tr><w:trPr><w:cantSplit/></w:trPr>'+c.map(function(col,i){return cell(pre?String(it["__"+col.k]==null?"":it["__"+col.k]):val(it,col.k,r),ws[i],false,col.k==="#");}).join("")+'</w:tr>';});
  return x+'</w:tbl>';
}
$("bDoc").addEventListener("click",function(){
  if(!window.JSZip) return toast("مكتبة Word غير متاحة");
  var list=visible(), c=cols(), d=dayDate(), body="";
  body+=wP($("school").value,{b:true,sz:26,after:0});
  body+=wP(hijri(d)+" · "+greg(d)+"م",{sz:20,color:"555555",after:200});
  body+=wP("الطلاب الذين لم يدخلوا المنصة",{b:true,sz:34,center:true,after:60});
  body+=wP(filterText()+" · العدد: "+list.length,{sz:22,center:true,color:"333333",after:200});
  var gsum=gradeSummaryRows();
  if(gsum){ body+=wP("نسبة عدم الدخول حسب الصف",{b:true,sz:24,color:"0F6B63",after:60});
    var gc=gsum.headers.map(function(h,i){return {k:["g","id","s","phone","user","last"][i],t:h};});
    body+=wTableRows(gsum.rows,gc)+wP("",{after:120}); }
  groups(list).forEach(function(g){
    if(g.title) body+=wP(g.title+" ("+g.items.length+")",{b:true,sz:24,color:"0F6B63",before:200,after:60});
    body+=wTable(g.items,c);
  });
  var doc='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+body+
    '<w:sectPr>'+(c.length>6?'<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>':'<w:pgSz w:w="11906" w:h="16838"/>')+'<w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="500" w:footer="500" w:gutter="0"/><w:bidi/></w:sectPr></w:body></w:document>';
  var z=new JSZip();
  z.file("[Content_Types].xml",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file("_rels/.rels",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file("word/document.xml",doc);
  z.generateAsync({type:"blob",mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"}).then(function(b){save(b,baseName()+".docx");toast("تم حفظ Word");});
});

/* ---------- sample ---------- */
$("bSample").addEventListener("click",function(){
  var N=[["عبدالله محمد العتيبي","1",1],["فيصل سعد القحطاني","1",1],["تركي ناصر الشهري","1",2],["سلطان خالد الدوسري","1",3],
    ["مشاري فهد الحربي","2",1],["نواف عبدالرحمن المطيري","2",1],["ريان عبدالعزيز الزهراني","2",2],["يزيد ماجد الغامدي","2",4],
    ["عمر حمد السبيعي","3",1],["بدر صالح العنزي","3",2],["خالد إبراهيم الشمري","3",2],["محمد عادل البقمي","3",3]];
  setData(N.map(function(n,i){return{name:n[0]+" (مثال)",id:String(1100000000+i*7919),phone:"05"+String(10000000+i*731113).slice(0,8),user:"s"+(44100+i*37),last:i%3?"—":"1448/5/10",g:n[1],gRaw:"",s:String(n[2])};}),"بيانات مثال");
});


/* ---------- separate export of the per-grade statistics ---------- */
function gBase(){ return "نسبة-عدم-الدخول-حسب-الصف-"+($("day").value||ymd(new Date())); }
function buildGradePage(){
  var host=PAGES; host.innerHTML=""; var g=gradeSummaryRows(), st=gradeStats(), d=dayDate(); if(!g) return [];
  var p=document.createElement("div"); p.className="page";
  var bars=st.rows.concat(st.full?[st.total]:[]).map(function(r,i,a){ var last=st.full&&i===a.length-1;
    return '<div style="display:grid;grid-template-columns:150px 1fr 70px;gap:10px;align-items:center;margin:7px 0;'+(last?'font-weight:700':'')+'">'+
      '<div>'+esc(r.label)+'</div><div style="height:16px;background:#ECEFEC;border-radius:8px;overflow:hidden"><div style="height:100%;width:'+Math.max(0,Math.min(100,r.ratio*100)).toFixed(1)+'%;background:'+(last?'#8F2A22':'#B42318')+';border-radius:8px"></div></div>'+
      '<div style="text-align:left;direction:ltr">'+esc(r.pct)+'</div></div>'; }).join("");
  p.innerHTML='<div class="ph"><div class="s">'+esc($("school").value)+'</div><div class="m">'+esc(hijri(d))+'<br>'+esc(greg(d))+'م</div></div>'+
    '<h1>نسبة عدم الدخول للمنصة حسب الصف</h1><div class="flt">'+esc(st.full?"من إجمالي طلاب كل صف":"حصة كل صف من قائمة الطلاب الذين لم يدخلوا")+(fileName?' · '+esc(fileName):'')+'</div>'+
    gsTableHtml().replace('<h3>نسبة عدم الدخول حسب الصف</h3>','')+'<h3>التمثيل البياني</h3>'+bars+
    '<div class="pf"><span>'+esc($("school").value)+'</span><span class="pn">صفحة 1 من 1</span></div>';
  host.appendChild(p); return [p];
}
function makeDocxBlob(body){
  var doc='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+body+
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000" w:header="500" w:footer="500" w:gutter="0"/><w:bidi/></w:sectPr></w:body></w:document>';
  var z=new JSZip();
  z.file("[Content_Types].xml",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  z.file("_rels/.rels",'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  z.file("word/document.xml",doc);
  return z.generateAsync({type:"blob",mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
}
function gNeed(){ if(!gradeSummaryRows()){ toast("ارفع ملف الطلاب أولاً"); return false; } return true; }
$("gPdf").addEventListener("click",function(){
  if(!gNeed()) return;
  if(window.BalatGradeExport){ window.BalatGradeExport("pdf"); return; }
  if(!(window.jspdf||window.jsPDF)||!window.html2canvas) return toast("مكتبة PDF غير متاحة");
  busy(true);
  setTimeout(function(){
    var pages=buildGradePage(), pdf=new ((window.jspdf&&window.jspdf.jsPDF)||window.jsPDF)({unit:"pt",format:"a4",orientation:"portrait"}), W=pdf.internal.pageSize.getWidth(), H=pdf.internal.pageSize.getHeight();
    html2canvas(pages[0],{scale:2,backgroundColor:"#ffffff",useCORS:true,logging:false})
      .then(function(cv){ pdf.addImage(cv.toDataURL("image/jpeg",0.92),"JPEG",0,0,W,H); save(pdf.output("blob"),gBase()+".pdf"); toast("تم حفظ PDF"); })
      .catch(function(){toast("تعذّر إنشاء PDF");})
      .then(function(){busy(false);PAGES.innerHTML="";});
  },30);
});
$("gPrn").addEventListener("click",function(){ if(!gNeed()) return; if(window.BalatGradeExport){ window.BalatGradeExport("print"); return; } buildGradePage(); document.body.classList.add("nl-print"); setTimeout(function(){window.print();},50); });
$("gDoc").addEventListener("click",function(){
  if(!gNeed()) return; if(!window.JSZip) return toast("مكتبة Word غير متاحة");
  var g=gradeSummaryRows(), st=gradeStats(), d=dayDate(), body="";
  body+=wP($("school").value,{b:true,sz:26,after:0});
  body+=wP(hijri(d)+" · "+greg(d)+"م",{sz:20,color:"555555",after:200});
  body+=wP("نسبة عدم الدخول للمنصة حسب الصف",{b:true,sz:34,center:true,after:60});
  body+=wP(st.full?"من إجمالي طلاب كل صف":"حصة كل صف من قائمة الطلاب الذين لم يدخلوا",{sz:22,center:true,color:"333333",after:200});
  var gc=g.headers.map(function(h,i){return {k:["g","id","s","phone","user","last"][i],t:h};});
  body+=wTableRows(g.rows,gc);
  makeDocxBlob(body).then(function(b){save(b,gBase()+".docx");toast("تم حفظ Word");});
});
$("gXls").addEventListener("click",function(){
  if(!gNeed()) return;
  var g=gradeSummaryRows(), st=gradeStats(), wb=XLSX.utils.book_new();
  var ws=XLSX.utils.aoa_to_sheet([[$("school").value],["نسبة عدم الدخول للمنصة حسب الصف"],[hijri(dayDate())+" · "+greg(dayDate())+"م"],[st.full?"من إجمالي طلاب كل صف":"حصة كل صف من قائمة الطلاب الذين لم يدخلوا"],[],g.headers].concat(g.rows));
  ws["!cols"]=g.headers.map(function(h,i){return {wch:i?18:24};});
  ws["!merges"]=[0,1,2,3].map(function(r){return {s:{r:r,c:0},e:{r:r,c:g.headers.length-1}};});
  XLSX.utils.book_append_sheet(wb,ws,"ملخص الصفوف"); wb.Workbook={Views:[{RTL:true}]};
  XLSX.writeFile(wb,gBase()+".xlsx"); toast("تم حفظ Excel");
});

/* ---------- public API (used by the main page: top upload slot + the form) ---------- */
var FW={"#":.5,name:3,id:1.5,g:1.3,s:.7,phone:1.4,user:2.6,last:1.5};
function formGroups(meta){
  meta=meta||{}; var list=visible(), c=cols();
  var when=(meta.day||meta.hijri)?(" ليوم "+(meta.day||"")+(meta.hijri?" الموافق "+meta.hijri:"")):"";
  var school=meta.school?" – مدرسة "+meta.school:"";
  return groups(list).map(function(g){
    return {title:"الطلاب الذين لم يدخلوا المنصة"+(g.title?" – "+g.title:" – "+filterText())+school+when+" ("+g.items.length+" طالبًا)",
      headers:c.map(function(x){return x.t;}), widths:c.map(function(x){return FW[x.k]||1;}),
      rows:g.items.map(function(x,i){return c.map(function(col){return val(x,col.k,i);});})};
  });
}
function appendDocx(Z,bytes,meta){
  var gs=formGroups(meta); if(!gs.length||!visible().length) return Promise.resolve(bytes);
  var c=cols();
  return Z.loadAsync(bytes).then(function(z){
    var f=z.file("word/document.xml"); if(!f) return bytes;
    return f.async("string").then(function(xml){
      var body='<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
      gs.forEach(function(g,i){
        body+=wP(g.title,{b:true,sz:24,color:"0F6B63",before:i?240:0,after:80});
        body+=wTableRows(g.rows,c);
      });
      var at=xml.lastIndexOf("<w:sectPr"); if(at<0) at=xml.lastIndexOf("</w:body>");
      xml=xml.slice(0,at)+body+xml.slice(at);
      z.file("word/document.xml",xml);
      return z.generateAsync({type:"uint8array"});
    });
  });
}
window.NoLogin={
  loadFile:loadFileP,
  tryPages:tryPages,
  tryDocx:tryDocx,
  count:function(){return students.length?visible().length:0;},
  hasData:function(){return students.length>0;},
  stats:function(){ if(!students.length||!lastInfo) return null; return {full:!!(lastInfo.full&&lastInfo.total!==students.length),total:lastInfo.total,absent:students.length}; },
  summary:function(){return students.length?filterText():"";},
  formGroups:formGroups,
  gradeCards:gradeCards,
  gradeSummary:gradeSummaryRows,
  appendDocx:appendDocx,
  clear:function(){ students=[]; fileName=""; lastInfo=null; selSecs=new Set(); has={}; show={}; $("fileTxt").textContent="لم يُرفع ملف"; showErr(""); render();
    try{document.dispatchEvent(new CustomEvent("nologin:cleared"));}catch(e){} }
};

/* ---------- boot ---------- */
$("day").value=ymd(new Date());
render();

})();
