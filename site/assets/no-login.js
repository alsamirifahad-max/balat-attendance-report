/* خانة «الطلاب الذين لم يدخلوا المنصة» — وحدة مستقلة تمامًا.
   تعمل داخل Shadow DOM فلا تتأثر بتنسيقات الموقع ولا تؤثر عليها، ولا تعتمد على app.js.
   تحتاج: vendor/xlsx.full.min.js (قراءة وحفظ Excel) + مكتبات الموقع الموجودة jszip و html2canvas و jspdf. */
(function(){
"use strict";
var HOST=document.getElementById("nl-root"); if(!HOST||HOST.shadowRoot) return;
var R=HOST.attachShadow({mode:"open"});
R.innerHTML="<style>"+"\n:host{--bg:var(--paper,#f4f7f6);--muted:var(--ink-2,#4a5f5d);--accent:var(--teal,#1d6f68);--accent-ink:var(--on-teal,#fff);--accent-soft:var(--teal-soft,#e2f0ed);--miss:var(--red,#b3372f);--miss-soft:var(--red-soft,#f9e3e0);--font:var(--body,Tahoma,sans-serif)}\nh1,h2,h3{font-family:var(--display,inherit);font-weight:600}\n*,*::before,*::after{box-sizing:border-box}\n:host{display:block;color:var(--ink);font-family:var(--font);font-size:15px;line-height:1.6;margin-top:22px}\n.wrap{display:grid;grid-template-columns:minmax(0,1fr);gap:16px}\nh1{margin:0;font-size:24px;text-wrap:balance}\nh2{margin:0;font-size:16px}\n.sub,.note{color:var(--muted)}\n.note{font-size:13px;margin:0}\n.card{background:var(--card);border:1px solid var(--line);border-radius:16px;box-shadow:var(--shadow,0 1px 2px rgba(16,48,47,.05),0 6px 22px rgba(16,48,47,.07));padding:16px;display:grid;gap:12px;min-width:0;grid-template-columns:minmax(0,1fr)}\n.row{display:flex;flex-wrap:wrap;gap:10px;align-items:center}\n.field{display:grid;gap:4px;font-size:13px;color:var(--muted)}\ninput,select{font:inherit;color:var(--ink);background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:7px 10px;min-width:0;max-width:100%}\nbutton{font:inherit;cursor:pointer}\n:focus-visible{outline:2px solid var(--accent);outline-offset:2px}\n\n/* grade segmented */\n.seg{display:flex;flex-wrap:wrap;gap:6px}\n.seg button{border:1px solid var(--line);background:var(--bg);color:var(--ink);border-radius:999px;padding:7px 16px}\n.seg button[aria-pressed=\"true\"]{background:var(--accent);border-color:var(--accent);color:var(--accent-ink);font-weight:600}\n.seg button .c{font-size:12px;opacity:.8;margin-inline-start:6px;font-variant-numeric:tabular-nums}\n\n/* section chips */\n.chips{display:flex;flex-wrap:wrap;gap:6px}\n.chip{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--line);background:var(--bg);border-radius:8px;padding:5px 10px;font-size:14px;cursor:pointer;user-select:none}\n.chip input{margin:0;accent-color:var(--accent)}\n.chip.on{border-color:var(--accent);background:var(--accent-soft)}\n.chip.scr{border-style:dashed}\n#colPick .btn.sm{margin-inline-start:4px}\n.chip .c{color:var(--muted);font-size:12px;font-variant-numeric:tabular-nums}\n\n/* upload slot */\n.slot{border:2px dashed var(--line);border-radius:14px;padding:26px 16px;text-align:center;display:grid;gap:6px;cursor:pointer;background:var(--card)}\n.slot:hover,.slot.over{border-color:var(--accent);background:var(--accent-soft)}\n.slot .t{font-size:18px;font-weight:700}\n.slot .f{font-size:13px;color:var(--muted);overflow-wrap:anywhere}\n.slot .f b{color:var(--ink)}\n.slot input{display:none}\n\n/* actions */\n.actions{display:flex;flex-wrap:wrap;gap:8px}\n.btn{border:1px solid var(--line);background:var(--card);color:var(--ink);border-radius:10px;padding:9px 16px;font-weight:600;display:inline-flex;gap:8px;align-items:center}\n.btn .x{font-size:11px;font-weight:700;border-radius:4px;padding:1px 5px;color:#fff}\n.btn:disabled{opacity:.45;cursor:not-allowed}\n.gwrap{display:grid;gap:8px;padding-bottom:12px;border-bottom:1px solid var(--line)}\n.btn.sm{padding:6px 12px;font-size:13px}\n.gact{gap:6px}\n.gwrap .ghead{display:flex;flex-wrap:wrap;gap:6px 12px;align-items:baseline;justify-content:space-between}\n.gwrap .ghead b{font-size:14px}\n.gstats{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px}\n.gs{border:1px solid var(--line);border-top:4px solid var(--pc);border-radius:12px;padding:12px 14px;display:grid;gap:8px;background:var(--card);min-width:0}\n.gs .gt{display:flex;justify-content:space-between;gap:8px;font-size:13.5px;color:var(--muted);font-weight:600}\n.gs .gtag{font-size:11.5px;font-weight:500;padding:1px 8px;border-radius:99px;background:var(--miss-soft);color:var(--miss);white-space:nowrap}\n.gs .gv{font-family:var(--display,inherit);font-size:30px;font-weight:600;line-height:1.1;color:var(--pc);font-variant-numeric:tabular-nums}\n.gs .gv small{font-size:15px;color:var(--muted);font-weight:500}\n.gs .gduo{display:flex;height:8px;border-radius:4px;overflow:hidden;background:var(--line)}\n.gs .gduo i{display:block;height:100%}\n.gs .gpr{display:grid;grid-template-columns:1fr 1fr;gap:6px}\n.gs .gpr.one{grid-template-columns:1fr}\n.gs .gpr div{border-radius:8px;padding:6px 8px;display:grid;font-size:12px;color:var(--muted)}\n.gs .gpr b{font-family:var(--display,inherit);font-size:18px;font-variant-numeric:tabular-nums}\n.gs .gpr .y{background:var(--green-soft,#e0f1e6)} .gs .gpr .y b{color:var(--green,#26794a)}\n.gs .gpr .n{background:var(--red-soft,#f9e3e0)} .gs .gpr .n b{color:var(--red,#b3372f)}\n.secrank{display:grid;gap:8px;margin-top:4px}\n.seclist{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:6px 14px}\n.sr{display:grid;grid-template-columns:26px minmax(92px,auto) 1fr 52px 72px;gap:8px;align-items:center;font-size:13px;padding:4px 0;border-bottom:1px dashed var(--line)}\n.sr .rk{width:24px;height:24px;border-radius:50%;display:grid;place-items:center;font-size:12px;font-weight:700;background:var(--bg);color:var(--muted)}\n.sr .sl{font-weight:600;white-space:nowrap}\n.sr .sb{height:8px;border-radius:4px;background:var(--line);overflow:hidden}\n.sr .sb i{display:block;height:100%;border-radius:4px;background:var(--green,#26794a)}\n.sr.mid .sb i{background:var(--amber,#a86412)} .sr.lo .sb i{background:var(--red,#b3372f)}\n.sr .sp{font-weight:700;font-variant-numeric:tabular-nums;text-align:left}\n.sr .sc{color:var(--muted);font-size:12px;font-variant-numeric:tabular-nums;white-space:nowrap}\ndetails.tpl{border:1px solid var(--line);border-radius:10px;padding:8px 12px;background:var(--bg)}\ndetails.tpl summary{cursor:pointer;font-weight:600;font-size:14px}\ndetails.tpl .tplbody{display:grid;gap:10px;margin-top:10px}\ndetails.tpl textarea{width:100%;font:inherit;font-size:13.5px;line-height:1.6;resize:vertical}\na.wa{display:inline-flex;align-items:center;gap:4px;font-size:12.5px;font-weight:600;text-decoration:none;color:#fff;background:#1f8f4e;border-radius:6px;padding:3px 10px;white-space:nowrap}\na.wa:hover{background:#18763f}\n.grp h3 .gbtns{display:inline-flex;gap:6px;margin-inline-start:auto}\n.grp h3 .gbtns .btn,.grp h3 .gbtns a{font-size:12px;padding:3px 10px;font-weight:600}\n.hdays{display:flex;flex-wrap:wrap;gap:6px}\n.hd{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--line);border-radius:8px;padding:3px 4px 3px 10px;font-size:13px;background:var(--bg);font-variant-numeric:tabular-nums}\n.hd small{color:var(--muted)}\n.hx{border:0;background:transparent;color:var(--muted);cursor:pointer;font-size:12px;border-radius:5px;padding:2px 6px}\n.hx:hover,.hx.sure{background:var(--miss-soft);color:var(--miss)}\n.hx.sure::after{content:\" حذف؟\"}\n.chartbox{border:1px solid var(--line);border-radius:12px;padding:8px;background:var(--card);min-width:0;overflow:hidden}\n.chartbox canvas{display:block;max-width:100%;width:100%}\n.cntpill{display:inline-block;border-radius:99px;padding:1px 10px;font-size:12.5px;font-weight:700;background:var(--miss-soft);color:var(--miss);font-variant-numeric:tabular-nums}\n.cntpill.all{background:var(--miss);color:#fff}\n@media (max-width:560px){ .sr{grid-template-columns:24px minmax(80px,auto) 1fr 48px} .sr .sc{display:none} .seclist{grid-template-columns:1fr} }\n.gs .gcnt{display:flex;justify-content:space-between;font-size:12px;color:var(--muted);font-variant-numeric:tabular-nums}\n@media (prefers-reduced-motion:no-preference){.gs .gduo i{transition:width .5s ease}}\n.x.pdf{background:#B42318}.x.doc{background:#2B579A}.x.xls{background:#217346}.x.prn{background:#5B6770}\n.check{display:inline-flex;gap:6px;align-items:center;font-size:14px}\n\n/* report preview */\n.report{display:grid;gap:14px}\n#repBody{display:grid;grid-template-columns:minmax(0,1fr);gap:18px}\n.report{grid-template-columns:minmax(0,1fr)}\n.grp{min-width:0}\n.rep-head{display:flex;flex-wrap:wrap;justify-content:space-between;gap:8px;align-items:end;border-bottom:2px solid var(--accent);padding-bottom:10px}\n.rep-head .school{font-weight:700}\n.rep-head .meta{font-size:13px;color:var(--muted);text-align:left}\n.total{font-size:14px}\n.total b{color:var(--miss);font-size:18px;font-variant-numeric:tabular-nums}\n.grp h3{margin:0 0 6px;font-size:15px;display:flex;gap:8px;align-items:center}\n.grp h3 .n{font-size:12px;font-weight:600;background:var(--miss-soft);color:var(--miss);border-radius:999px;padding:1px 10px}\n.scroll{overflow-x:auto}\ntable{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums}\nth,td{border:1px solid var(--line);padding:6px 8px;text-align:right;white-space:nowrap}\ntd.name{white-space:normal;min-width:180px}\nth{background:var(--accent-soft);font-weight:600;font-size:13px}\ntd.num{text-align:center;width:44px;color:var(--muted)}\ntd.ltr{direction:ltr;text-align:right}\n.empty{color:var(--muted);text-align:center;padding:28px 12px}\n.err{color:var(--miss);font-size:14px;margin:0}\n.toast{position:fixed;bottom:calc(18px + env(safe-area-inset-bottom,0px));left:50%;transform:translateX(-50%);background:var(--ink);color:var(--bg);padding:9px 16px;border-radius:8px;font-size:14px;z-index:30}\n.busy{position:fixed;inset:0;background:rgba(0,0,0,.35);display:grid;place-items:center;z-index:40;color:#fff;font-size:16px}\n\n@media (prefers-reduced-motion:no-preference){.slot,.chip,.seg button{transition:background .15s,border-color .15s}}\n[hidden]{display:none!important}\n"+"</style>"+"\n<div class=\"wrap\">\n  <header class=\"row\" style=\"justify-content:space-between;align-items:end\">\n    <div>\n      <h2 style=\"font-size:20px\">الطلاب الذين لم يدخلوا المنصة</h2>\n      <div class=\"sub\" id=\"todayTxt\"></div>\n    </div>\n  </header>\n\n  <!-- 1) Classification before upload -->\n  <section class=\"card\" aria-labelledby=\"fTitle\">\n    <h2 id=\"fTitle\">التصنيف</h2>\n    <div class=\"seg\" id=\"gradeSeg\" role=\"group\" aria-label=\"الصف\">\n      <button type=\"button\" data-g=\"\" aria-pressed=\"true\">الكل<span class=\"c\" id=\"gc-\"></span></button>\n      <button type=\"button\" data-g=\"1\" aria-pressed=\"false\">الأول ثانوي<span class=\"c\" id=\"gc-1\"></span></button>\n      <button type=\"button\" data-g=\"2\" aria-pressed=\"false\">الثاني ثانوي<span class=\"c\" id=\"gc-2\"></span></button>\n      <button type=\"button\" data-g=\"3\" aria-pressed=\"false\">الثالث ثانوي<span class=\"c\" id=\"gc-3\"></span></button>\n    </div>\n    <div class=\"row\" style=\"justify-content:space-between\">\n      <span class=\"note\">الشعب الدراسية</span>\n      <span class=\"row\" style=\"gap:6px\">\n        <button type=\"button\" class=\"btn\" id=\"secAll\" style=\"padding:4px 10px;font-weight:500;font-size:13px\">تحديد الكل</button>\n        <button type=\"button\" class=\"btn\" id=\"secNone\" style=\"padding:4px 10px;font-weight:500;font-size:13px\">إلغاء التحديد</button>\n      </span>\n    </div>\n    <div class=\"chips\" id=\"secChips\"><span class=\"note\">تظهر الشعب هنا بعد قراءة الملف.</span></div>\n  </section>\n\n  <!-- 2) The single upload slot -->\n  <label class=\"slot\" id=\"slot\" for=\"inFile\">\n    <span class=\"t\">ملف الطلاب الذين لم يدخلوا المنصة</span>\n    <span class=\"note\">ضع الملف هنا أو اضغط للاختيار · Excel أو Word أو PDF</span>\n    <span class=\"f\" id=\"fileTxt\">لم يُرفع ملف</span>\n    <input type=\"file\" id=\"inFile\" accept=\".xlsx,.xls,.xlsm,.csv,.docx,.pdf\">\n  </label>\n  <p class=\"err\" id=\"err\" hidden></p>\n\n  <!-- 3) Save / print -->\n  <section class=\"card\" aria-labelledby=\"aTitle\">\n    <div class=\"row\" style=\"justify-content:space-between\">\n      <h2 id=\"aTitle\">الحفظ والطباعة</h2>\n      <label class=\"check\" for=\"optSplit\"><input type=\"checkbox\" id=\"optSplit\" checked> جدول مستقل لكل شعبة</label>\n    </div>\n    <div class=\"gwrap\" id=\"gWrap\" hidden>\n      <label class=\"field\" for=\"sgSel\" style=\"max-width:260px\">عرض إحصائيات وترتيب <select id=\"sgSel\"><option value=\"\">كل الصفوف</option><option value=\"1\">الأول ثانوي فقط</option><option value=\"2\">الثاني ثانوي فقط</option><option value=\"3\">الثالث ثانوي فقط</option></select></label>\n      <div class=\"ghead\"><b>نسبة عدم الدخول حسب الصف</b><span class=\"note\" id=\"gNote\"></span></div>\n      <div class=\"gstats\" id=\"gStats\"></div>\n      <div class=\"secrank\" id=\"secRank\" hidden>\n        <div class=\"ghead\"><b>ترتيب الشعب حسب نسبة الدخول</b><span class=\"note\" id=\"secRankNote\"></span></div>\n        <div class=\"seclist\" id=\"secList\"></div>\n      </div>\n      <div class=\"actions gact\">\n        <button class=\"btn sm\" id=\"gPdf\" type=\"button\"><span class=\"x pdf\">PDF</span>حفظ الإحصائية PDF</button>\n        <button class=\"btn sm\" id=\"gDoc\" type=\"button\"><span class=\"x doc\">W</span>Word</button>\n        <button class=\"btn sm\" id=\"gXls\" type=\"button\"><span class=\"x xls\">X</span>Excel</button>\n        <button class=\"btn sm\" id=\"gPrn\" type=\"button\"><span class=\"x prn\">⎙</span>طباعة</button>\n      </div>\n    </div>\n    <details class=\"tpl\">\n      <summary>نصوص رسائل الواتساب (ولي الأمر · رائد الفصل)</summary>\n      <div class=\"tplbody\">\n        <label class=\"field\" for=\"tplParent\">رسالة ولي الأمر (زر «واتساب» بجانب كل طالب)<textarea id=\"tplParent\" rows=\"5\"></textarea></label>\n        <label class=\"field\" for=\"tplTeacher\">رسالة رائد الفصل (أزرار رأس كل شعبة)<textarea id=\"tplTeacher\" rows=\"6\"></textarea></label>\n        <p class=\"note\">تُستبدل تلقائيًا: {الاسم} {الفصل} {اليوم} {التاريخ} {المدرسة} {العدد} {القائمة}. تُحفظ نصوصك في هذا الجهاز.</p>\n        <div><button class=\"btn sm\" id=\"tplReset\" type=\"button\">استعادة النصوص الافتراضية</button></div>\n      </div>\n    </details>\n    <div class=\"row\" style=\"gap:6px 14px\">\n      <span class=\"note\"><b>الأعمدة الظاهرة في العرض والطباعة والتنزيل:</b></span>\n      <span class=\"chips\" id=\"colPick\"><span class=\"note\">تظهر بعد قراءة الملف</span></span>\n    </div>\n    <div class=\"row\">\n      <label class=\"field\" for=\"school\">اسم المدرسة <input id=\"school\" value=\"ثانوية بلاط الشهداء بجدة\" style=\"width:240px\"></label>\n      <label class=\"field\" for=\"day\">التاريخ <input type=\"date\" id=\"day\"></label>\n    </div>\n    <div class=\"actions\">\n      <button class=\"btn\" id=\"bPdf\" type=\"button\" disabled><span class=\"x pdf\">PDF</span>حفظ PDF</button>\n      <button class=\"btn\" id=\"bDoc\" type=\"button\" disabled><span class=\"x doc\">W</span>حفظ Word</button>\n      <button class=\"btn\" id=\"bXls\" type=\"button\" disabled><span class=\"x xls\">X</span>حفظ Excel</button>\n      <button class=\"btn\" id=\"bPrn\" type=\"button\" disabled><span class=\"x prn\">⎙</span>طباعة</button>\n    </div>\n  </section>\n\n  <!-- 4) Result -->\n  <section class=\"card report\" id=\"report\" aria-live=\"polite\">\n    <div class=\"empty\" id=\"emptyBox\">\n      ضع ملف الطلاب الذين لم يدخلوا المنصة في الخانة أعلاه، وستظهر أسماؤهم هنا حسب الصف والشعبة.\n      <div style=\"margin-top:10px\"><button type=\"button\" class=\"btn\" id=\"bSample\" style=\"font-weight:500\">عرض مثال</button></div>\n    </div>\n    <div id=\"repBody\" hidden></div>\n  </section>\n\n  <!-- 5) Follow-up log -->\n  <section class=\"card\" aria-labelledby=\"hTitle\" id=\"histCard\">\n    <div class=\"row\" style=\"justify-content:space-between\">\n      <h2 id=\"hTitle\">سجل المتابعة الأسبوعي والشهري</h2>\n      <label class=\"check\" for=\"hAuto\"><input type=\"checkbox\" id=\"hAuto\" checked> حفظ كل يوم تلقائيًا عند رفع الملف</label>\n    </div>\n    <p class=\"note\" id=\"hCount\"></p>\n    <p class=\"note\" id=\"hSavedMsg\" style=\"color:var(--accent)\"></p>\n    <div class=\"hdays\" id=\"hDays\"></div>\n    <div class=\"row\">\n      <label class=\"field\" for=\"hPer\">الفترة\n        <select id=\"hPer\"><option value=\"week\">هذا الأسبوع</option><option value=\"7\">آخر 7 أيام</option><option value=\"30\">آخر 30 يومًا</option><option value=\"all\">كل الأيام المحفوظة</option><option value=\"custom\">من تاريخ إلى تاريخ</option></select></label>\n      <span class=\"row\" id=\"hCustom\" hidden>\n        <label class=\"field\" for=\"hFrom\">من <input type=\"date\" id=\"hFrom\"></label>\n        <label class=\"field\" for=\"hTo\">إلى <input type=\"date\" id=\"hTo\"></label>\n      </span>\n    </div>\n    <div class=\"empty\" id=\"hEmpty\">لا توجد أيام محفوظة في هذه الفترة. كل مرة ترفع ملف منصة مدرستي يُحفظ يومه هنا تلقائيًا، ثم تظهر المتابعة والرسم البياني وكشف المتكرر غيابهم.</div>\n    <div id=\"hBody\" hidden style=\"display:grid;gap:14px;grid-template-columns:minmax(0,1fr);min-width:0\">\n      <p class=\"note\" id=\"hRangeTxt\"></p>\n      <div class=\"gstats\" id=\"hCards\"></div>\n      <div class=\"chartbox\" id=\"hChartWrap\"><canvas id=\"hChart\" role=\"img\" aria-label=\"نسبة عدم الدخول يومًا بيوم\"></canvas></div>\n      <div class=\"row\" style=\"justify-content:space-between\">\n        <b>الطلاب المتكرر عدم دخولهم <span class=\"note\" id=\"hRepCount\"></span></b>\n        <span class=\"row\">\n          <label class=\"field\" for=\"hMin\">عدد الأيام <select id=\"hMin\"></select></label>\n          <label class=\"field\" for=\"hGrade\">الصف <select id=\"hGrade\"><option value=\"\">الكل</option><option value=\"1\">الأول ثانوي</option><option value=\"2\">الثاني ثانوي</option><option value=\"3\">الثالث ثانوي</option></select></label>\n        </span>\n      </div>\n      <div id=\"hRep\"></div>\n    </div>\n    <div class=\"actions\">\n      <button class=\"btn\" id=\"hPdf\" type=\"button\" disabled><span class=\"x pdf\">PDF</span>تقرير المتابعة PDF</button>\n      <button class=\"btn\" id=\"hDoc\" type=\"button\" disabled><span class=\"x doc\">W</span>Word</button>\n      <button class=\"btn\" id=\"hXls\" type=\"button\" disabled><span class=\"x xls\">X</span>Excel</button>\n      <button class=\"btn\" id=\"hPrn\" type=\"button\" disabled><span class=\"x prn\">⎙</span>طباعة</button>\n    </div>\n    <div class=\"row\" style=\"gap:8px\">\n      <button class=\"btn sm\" id=\"hBackup\" type=\"button\">حفظ نسخة احتياطية من السجل</button>\n      <label class=\"btn sm\" for=\"hImport\" style=\"cursor:pointer\">استرجاع نسخة احتياطية<input type=\"file\" id=\"hImport\" accept=\".json,application/json\" hidden></label>\n      <span class=\"note\">السجل محفوظ في هذا المتصفح فقط. احفظ نسخة احتياطية دوريًا، واسترجعها على أي جهاز آخر.</span>\n    </div>\n  </section>\n\n  <p class=\"note\">يُقرأ الملف داخل جهازك فقط ولا يُرفع إلى أي مكان.</p>\n</div>\n<div class=\"toast\" id=\"toast\" hidden></div>\n<div class=\"busy\" id=\"busy\" hidden>جارٍ تجهيز الملف…</div>\n\n";
var gst=document.createElement("style"); gst.textContent="/* A4 pages used for PDF and printing */\n#nl-pages{position:absolute;top:0;left:-20000px}\n#nl-pages .page{width:794px;height:1123px;padding:44px 46px 56px;background:#fff;color:#111;position:relative;overflow:hidden;font-family:\"IBM Plex Sans Arabic\",Tahoma,\"Segoe UI\",Arial,sans-serif;font-size:14px;line-height:1.5;direction:rtl}\n#nl-pages .page .ph{display:flex;justify-content:space-between;align-items:end;border-bottom:2px solid #0F6B63;padding-bottom:8px;margin-bottom:12px}\n#nl-pages .page .ph .s{font-weight:700;font-size:15px}\n#nl-pages .page .ph .m{font-size:12px;color:#444;text-align:left}\n#nl-pages .page h1{font-size:20px;text-align:center;margin:4px 0 2px}\n#nl-pages .page .flt{text-align:center;font-size:13px;color:#333;margin-bottom:12px}\n#nl-pages .page h3{font-size:14px;margin:12px 0 4px;color:#0F6B63}\n#nl-pages .page table{width:100%;border-collapse:collapse;table-layout:fixed}\n#nl-pages .page td{overflow-wrap:anywhere}\n#nl-pages .page td img{max-width:100%;height:auto}\n#nl-pages .page th,#nl-pages .page td{border:1px solid #9AA5A3;padding:4px 6px;font-size:13px;text-align:right}\n#nl-pages .page th{background:#E2EFEC;color:#16302f}\n#nl-pages .page td{background:#fff;color:#111}\n#nl-pages .page tbody tr:nth-child(even) td{background:#eaf4f1}\n#nl-pages .page td.num{text-align:center;color:#333}\n#nl-pages.dense .page th,#nl-pages.dense .page td{font-size:11px;padding:3px 4px}\n#nl-pages .page .pf{position:absolute;bottom:22px;left:46px;right:46px;display:flex;justify-content:space-between;font-size:11px;color:#555;border-top:1px solid #ccc;padding-top:6px}\n\n@media print{\n  @page{size:A4;margin:0}\n  body.nl-print{background:#fff!important;padding:0!important;margin:0!important}\n  body.nl-print>*:not(#nl-pages){display:none!important}\n  body.nl-print #nl-pages{position:static;left:auto}\n  #nl-pages .page{page-break-after:always;break-after:page}\n  #nl-pages .page:last-child{page-break-after:auto;break-after:auto}\n}\n"; document.head.appendChild(gst);
var PAGES=document.createElement("div"); PAGES.id="nl-pages"; PAGES.setAttribute("aria-hidden","true"); document.body.appendChild(PAGES);

"use strict";
var $=function(id){return R.getElementById(id);};
var GL={"1":"الأول ثانوي","2":"الثاني ثانوي","3":"الثالث ثانوي"};
var OPT=[{k:"id",t:"رقم الهوية"},{k:"g",t:"الصف"},{k:"s",t:"الشعبة"},{k:"phone",t:"الجوال"},{k:"user",t:"اسم المستخدم"},{k:"last",t:"آخر دخول"},{k:"wa",t:"خانة التواصل (واتساب)",screen:true}];
var OPT_FIXED={g:1,s:1};
function colPref(k){ var pf=lsGet("nl-cols",{}); return pf[k]!==false; }
var LTR={id:1,phone:1,user:1,last:1};
var lastGroups=[], lastInfo=null, students=[], fileName="", grade="", selSecs=new Set(), has={}, show={};

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
  info.bySec={}; all.forEach(function(x){ if(!x.g||!x.s||x.s==="—") return; var k=x.g+"|"+x.s; info.bySec[k]=(info.bySec[k]||0)+1; });
  setData(info.list,name,info);
  hRecord(info);
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
  var pref=lsGet("nl-cols",{}); has={}; show={}; OPT.forEach(function(o){has[o.k]=OPT_FIXED[o.k]?true:o.k==="wa"?students.some(function(x){return intlPhone(x.phone);}):students.some(function(x){return x[o.k];}); show[o.k]=has[o.k]&&pref[o.k]!==false;});
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
  OPT.forEach(function(o){if(!o.screen&&show[o.k])c.push(o);});
  return c;
}
function val(x,k,i){if(k==="#")return String(i+1);if(k==="g")return GL[x.g]||x.gRaw||"—";if(k==="phone"){var p=String(x.phone||"").replace(/\D/g,"");if(/^(00)?9660?5\d{8}$/.test(p))return "0"+p.slice(-9);if(/^5\d{8}$/.test(p))return "0"+p;return x.phone||"";}return x[k]||"";}

function renderColPick(){
  var box=$("colPick"); if(!students.length){box.innerHTML='<span class="note">تظهر بعد قراءة الملف</span>';return;}
  box.innerHTML=OPT.map(function(o){var dis=!has[o.k];return '<label class="chip'+(show[o.k]?" on":"")+(o.screen?" scr":"")+'" for="cp-'+o.k+'" title="'+(dis?"غير موجود في الملف":"")+'" style="'+(dis?"opacity:.45;cursor:not-allowed":"")+'"><input type="checkbox" id="cp-'+o.k+'" data-c="'+o.k+'"'+(show[o.k]?" checked":"")+(dis?" disabled":"")+'>'+o.t+(dis?' <span class="c">غير موجود</span>':(o.screen?' <span class="c">للشاشة فقط</span>':""))+'</label>';}).join("")+
    '<button type="button" class="btn sm" id="colMin">اسم الطالب والفصل فقط</button><button type="button" class="btn sm" id="colAll">إظهار الكل</button>';
}
/* ---------- absence rate per grade ---------- */
function pctTxt(a,b){ if(!b) return "—"; var v=Math.round(a/b*1000)/10; return String(v).replace(/\.0$/,"")+"%"; }
function gradeStats(){
  if(!students.length) return null;
  var full=!!(lastInfo&&lastInfo.full&&lastInfo.total!==students.length&&lastInfo.byGrade);
  var miss={}; students.forEach(function(x){ var k=x.g||""; miss[k]=(miss[k]||0)+1; });
  var sg=statsGrade(), keys=sg?[sg]:["1","2","3"];
  var allMiss=students.length, allTot=full?lastInfo.total:null;
  var rows=keys.map(function(k){ var m=miss[k]||0, t=full?((lastInfo.byGrade[k]||{}).total||0):null;
    return {k:k,label:GL[k]?"الصف "+GL[k]:"بدون صف",missing:m,total:t,pct:full?pctTxt(m,t):pctTxt(m,allMiss),ratio:full?(t?m/t:0):(allMiss?m/allMiss:0)}; });
  return {full:full,one:!!sg,showTotal:full&&!sg,rows:rows,total:{label:"إجمالي المدرسة",missing:allMiss,total:allTot,pct:full?pctTxt(allMiss,allTot):"100%",ratio:full?(allTot?allMiss/allTot:0):1}};
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
function statsGrade(){ var e=$("sgSel"), v=e?e.value:""; return GL[v]?v:""; }
function statsGradeLabel(){ var g=statsGrade(); return g?"الصف "+GL[g]:""; }
function gradeCards(){ var st=gradeStats(); if(!st) return null;
  var rows=st.rows.map(function(r){return Object.assign({pc:GPC[r.k]||"#1d6f68"},r);}).concat(st.showTotal?[Object.assign({pc:"#b3372f",isTotal:true},st.total)]:[]);
  return {full:st.full,one:st.one,gradeLabel:statsGradeLabel(),listTotal:students.length,cards:rows.map(function(r){
    var inn=r.total!=null?r.total-r.missing:null;
    return {title:r.label,pc:r.pc,isTotal:!!r.isTotal,v:r.missing,of:r.total,inn:inn,
      pIn:r.total!=null?pctTxt(inn,r.total):null,pOut:r.total!=null?pctTxt(r.missing,r.total):null,share:pctTxt(r.missing,students.length),ratio:r.ratio};})};
}
function gradeSummaryRows(){ var st=gradeStats(); if(!st) return null;
  var rows=st.rows.concat(st.showTotal?[st.total]:[]);
  return {full:st.full,hasTotal:st.showTotal,gradeLabel:statsGradeLabel(),headers:st.full?["الصف","إجمالي الطلاب","دخلوا","لم يدخلوا","نسبة الدخول","نسبة عدم الدخول"]:["الصف","لم يدخلوا","النسبة من قائمة الغائبين"],
    rows:rows.map(function(r){ return st.full?[r.label,String(r.total),String(r.total-r.missing),String(r.missing),pctTxt(r.total-r.missing,r.total),r.pct]:[r.label,String(r.missing),r.pct]; })};
}
function render(){
  renderFilters(); renderColPick(); renderGrades(); renderSections();
  try{document.dispatchEvent(new CustomEvent("nologin:change"));}catch(e){}
  var list=visible(), has=students.length>0;
  $("emptyBox").hidden=has; $("repBody").hidden=!has;
  ["bPdf","bDoc","bXls","bPrn"].forEach(function(id){$(id).disabled=!list.length;});
  var d=dayDate();
  $("todayTxt").textContent=hijri(d)+" · "+greg(d)+"م";
  if(!has) return;
  var c=cols(); var waCol=!!show.wa;
  var html='<div class="rep-head"><div><div class="school">'+esc($("school").value)+'</div><div class="note">'+esc(filterText())+'</div></div><div class="meta">'+esc(hijri(d))+'<br>'+esc(greg(d))+'م</div></div>'+
    '<div class="total">عدد الطلاب الذين لم يدخلوا المنصة: <b>'+list.length+'</b></div>';
  if(!list.length) html+='<div class="empty">لا يوجد طلاب في التصنيف المحدد. غيّر الصف أو حدّد شعبًا.</div>';
  var GR=groups(list); lastGroups=GR;
  GR.forEach(function(g,gi){
    var tbtn=(g.title&&$("optSplit").checked)?'<span class="gbtns"><button type="button" class="btn sm" data-tcopy="'+gi+'">نسخ رسالة الرائد</button><a class="wa" target="_blank" rel="noopener" href="'+esc(waLink("",teacherMsg(g)))+'">واتساب للرائد</a></span>':'';
    html+='<div class="grp">'+(g.title?'<h3>'+esc(g.title)+'<span class="n">'+g.items.length+'</span>'+tbtn+'</h3>':'')+
      '<div class="scroll"><table><thead><tr>'+c.map(function(x){return"<th>"+x.t+"</th>";}).join("")+(waCol?'<th>تواصل</th>':'')+'</tr></thead><tbody>'+
      g.items.map(function(x,i){return "<tr>"+c.map(function(col){var v=esc(val(x,col.k,i));
        return col.k==="#"?'<td class="num">'+v+"</td>":col.k==="name"?'<td class="name">'+v+"</td>":LTR[col.k]?'<td class="ltr">'+v+"</td>":"<td>"+v+"</td>";}).join("")+
        (waCol?'<td>'+(intlPhone(x.phone)?'<a class="wa" target="_blank" rel="noopener" href="'+esc(waLink(x.phone,parentMsg(x)))+'" title="رسالة لولي الأمر">واتساب</a>':'—')+'</td>':'')+"</tr>";}).join("")+
      '</tbody></table></div></div>';
  });
  $("repBody").innerHTML=html;
}
$("optSplit").addEventListener("change",function(){lsSet("nl-split",$("optSplit").checked);render();});
(function(){ var e=$("sgSel"); if(!e) return; var v=lsGet("nl-sgrade",""); if(GL[v]) e.value=v; e.addEventListener("change",function(){lsSet("nl-sgrade",e.value);render();}); })();
$("colPick").addEventListener("change",function(e){var k=e.target.dataset.c;if(!k)return;show[k]=e.target.checked;var pf=lsGet("nl-cols",{});pf[k]=e.target.checked;lsSet("nl-cols",pf);render();renderHist();});
$("colPick").addEventListener("click",function(e){var b=e.target.closest&&e.target.closest("button");if(!b||(b.id!=="colMin"&&b.id!=="colAll"))return;var on=b.id==="colAll",pf=lsGet("nl-cols",{});
  OPT.forEach(function(o){ if(!has[o.k]) return; var v=on||!!OPT_FIXED[o.k]; show[o.k]=v; pf[o.k]=v; }); lsSet("nl-cols",pf); render(); renderHist();});
$("school").addEventListener("input",function(){lsSet("nl-school",$("school").value);render();});
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
    g.rows.map(function(r,i){return '<tr'+(g.hasTotal&&i===g.rows.length-1?' style="font-weight:700"':'')+'>'+r.map(function(v){return '<td style="text-align:center">'+esc(v)+'</td>';}).join("")+'</tr>';}).join("")+'</tbody></table><div style="height:6px"></div>'; }
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
function gBase(){ return "نسبة-عدم-الدخول-حسب-الصف-"+(statsGrade()?GL[statsGrade()].replace(/\s+/g,"-")+"-":"")+($("day").value||ymd(new Date())); }
function buildGradePage(){
  var host=PAGES; host.innerHTML=""; var g=gradeSummaryRows(), st=gradeStats(), d=dayDate(); if(!g) return [];
  var p=document.createElement("div"); p.className="page";
  var bars=st.rows.concat(st.showTotal?[st.total]:[]).map(function(r,i,a){ var last=st.showTotal&&i===a.length-1;
    return '<div style="display:grid;grid-template-columns:150px 1fr 70px;gap:10px;align-items:center;margin:7px 0;'+(last?'font-weight:700':'')+'">'+
      '<div>'+esc(r.label)+'</div><div style="height:16px;background:#ECEFEC;border-radius:8px;overflow:hidden"><div style="height:100%;width:'+Math.max(0,Math.min(100,r.ratio*100)).toFixed(1)+'%;background:'+(last?'#8F2A22':'#B42318')+';border-radius:8px"></div></div>'+
      '<div style="text-align:left;direction:ltr">'+esc(r.pct)+'</div></div>'; }).join("");
  p.innerHTML='<div class="ph"><div class="s">'+esc($("school").value)+'</div><div class="m">'+esc(hijri(d))+'<br>'+esc(greg(d))+'م</div></div>'+
    '<h1>نسبة عدم الدخول للمنصة حسب الصف'+(statsGrade()?' — '+esc(statsGradeLabel()):'')+'</h1><div class="flt">'+esc(st.full?"من إجمالي طلاب كل صف":"حصة كل صف من قائمة الطلاب الذين لم يدخلوا")+(fileName?' · '+esc(fileName):'')+'</div>'+
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
  body+=wP("نسبة عدم الدخول للمنصة حسب الصف"+(statsGrade()?" — "+statsGradeLabel():""),{b:true,sz:34,center:true,after:60});
  body+=wP(st.full?"من إجمالي طلاب كل صف":"حصة كل صف من قائمة الطلاب الذين لم يدخلوا",{sz:22,center:true,color:"333333",after:200});
  var gc=g.headers.map(function(h,i){return {k:["g","id","s","phone","user","last"][i],t:h};});
  body+=wTableRows(g.rows,gc);
  var stb=sectionTable(); if(stb){ body+=wP("ترتيب الشعب حسب نسبة الدخول"+(statsGrade()?" — "+statsGradeLabel():""),{b:true,sz:24,color:"0F6B63",before:240,after:60});
    body+=wTableRows(stb.rows,stb.headers.map(function(h,i){return {k:["#","g","id","s","phone","user"][i],t:h};})); }
  makeDocxBlob(body).then(function(b){save(b,gBase()+".docx");toast("تم حفظ Word");});
});
$("gXls").addEventListener("click",function(){
  if(!gNeed()) return;
  var g=gradeSummaryRows(), st=gradeStats(), wb=XLSX.utils.book_new();
  var ws=XLSX.utils.aoa_to_sheet([[$("school").value],["نسبة عدم الدخول للمنصة حسب الصف"+(statsGrade()?" — "+statsGradeLabel():"")],[hijri(dayDate())+" · "+greg(dayDate())+"م"],[st.full?"من إجمالي طلاب كل صف":"حصة كل صف من قائمة الطلاب الذين لم يدخلوا"],[],g.headers].concat(g.rows));
  ws["!cols"]=g.headers.map(function(h,i){return {wch:i?18:24};});
  ws["!merges"]=[0,1,2,3].map(function(r){return {s:{r:r,c:0},e:{r:r,c:g.headers.length-1}};});
  XLSX.utils.book_append_sheet(wb,ws,"ملخص الصفوف");
  var stb=sectionTable(); if(stb){ var ws2=XLSX.utils.aoa_to_sheet([["ترتيب الشعب حسب نسبة الدخول"+(statsGrade()?" — "+statsGradeLabel():"")],[],stb.headers].concat(stb.rows)); ws2["!cols"]=stb.headers.map(function(h,i){return {wch:i===1?24:14};}); XLSX.utils.book_append_sheet(wb,ws2,"ترتيب الشعب"); }
  wb.Workbook={Views:[{RTL:true}]};
  XLSX.writeFile(wb,gBase()+".xlsx"); toast("تم حفظ Excel");
});

/* ========== saved settings ========== */
function lsGet(k,d){try{var v=localStorage.getItem(k);return v==null?d:JSON.parse(v);}catch(e){return d;}}
function lsSet(k,v){try{localStorage.setItem(k,JSON.stringify(v));return true;}catch(e){return false;}}

/* ========== WhatsApp messages ========== */
var DEF_PARENT="السلام عليكم ورحمة الله وبركاته\nولي أمر الطالب/ {الاسم}\nنفيدكم بعدم دخول ابنكم منصة مدرستي يوم {اليوم} {التاريخ}، نأمل متابعته وحثّه على الدخول وحضور الدروس.\n{المدرسة}";
var DEF_TEACHER="الأستاذ الفاضل رائد فصل {الفصل}\nالسلام عليكم ورحمة الله\nالطلاب الذين لم يدخلوا منصة مدرستي يوم {اليوم} {التاريخ} ({العدد}):\n{القائمة}\n{المدرسة}";
function fillTpl(t,o){return String(t).replace(/\{(الاسم|الفصل|اليوم|التاريخ|المدرسة|العدد|القائمة)\}/g,function(m,k){return o[k]!=null?o[k]:m;});}
function intlPhone(p){p=latin(String(p||"")).replace(/\D/g,""); var m=p.match(/(?:00)?9660?(5\d{8})$/)||p.match(/^0?(5\d{8})$/); return m?"966"+m[1]:"";}
function dayParts(dt){var d=dt||dayDate(); var day=d.toLocaleDateString("ar-SA",{weekday:"long"}), h;
  try{h=d.toLocaleDateString("ar-SA-u-ca-islamic-umalqura",{day:"numeric",month:"long",year:"numeric"});}catch(e){h=greg(d);} return {day:day,date:h};}
function tplParent(){return $("tplParent")&&$("tplParent").value.trim()?$("tplParent").value:DEF_PARENT;}
function tplTeacher(){return $("tplTeacher")&&$("tplTeacher").value.trim()?$("tplTeacher").value:DEF_TEACHER;}
function clsShort(g,s){return (GL[g]||"")+(s&&s!=="—"?" - "+s:"");}
function parentMsg(x,extra){var dp=dayParts(); var o={"الاسم":x.name,"الفصل":clsShort(x.g,x.s),"اليوم":dp.day,"التاريخ":dp.date,"المدرسة":$("school").value};
  if(extra) Object.keys(extra).forEach(function(k){o[k]=extra[k];}); return fillTpl(tplParent(),o);}
function waLink(phone,text){var p=intlPhone(phone); return "https://wa.me/"+p+"?text="+encodeURIComponent(text);}
function teacherMsg(g){var dp=dayParts(); var x=g.items[0]||{};
  return fillTpl(tplTeacher(),{"الفصل":clsShort(x.g,x.s),"اليوم":dp.day,"التاريخ":dp.date,"المدرسة":$("school").value,"العدد":String(g.items.length),
    "القائمة":g.items.map(function(s,i){return (i+1)+". "+s.name;}).join("\n")});}
function copyText(t){ var done=function(){toast("نُسخت الرسالة");};
  if(navigator.clipboard&&navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(done,function(){fallbackCopy(t);}); else fallbackCopy(t); }
function fallbackCopy(t){var ta=document.createElement("textarea");ta.value=t;ta.style.position="fixed";ta.style.opacity="0";((typeof R!=="undefined"&&R&&R.appendChild)?R:document.body).appendChild(ta);ta.select();try{document.execCommand("copy");toast("نُسخت الرسالة");}catch(e){toast("انسخ الرسالة يدويًا");}ta.remove();}

/* ========== matching a Naqar student to the platform list ========== */
function toks(s){return norm(s).replace(/عبد\s+/g,"عبد").split(" ").filter(function(t){return t&&t!=="بن"&&t!=="ابن"&&t!=="بنت";});}
function matchStudent(name,gWord,sec){
  var g={"أول":"1","ثاني":"2","ثالث":"3"}[gWord]||""; var a=toks(name); if(!a.length) return null;
  var cands=students.filter(function(x){return (!g||x.g===g)&&(sec==null||sec===""||String(x.s)===String(sec));});
  var best=null;
  cands.forEach(function(x){ var b=toks(x.name), sc=0;
    if(a.join(" ")===b.join(" ")) sc=3;
    else if(a[0]===b[0]&&a[a.length-1]===b[b.length-1]) sc=2;
    else if(a.length>=2&&b.length>=2&&a[0]===b[0]&&a[1]===b[1]&&(b.indexOf(a[a.length-1])>=0||a.indexOf(b[b.length-1])>=0)) sc=1;
    if(sc&&(!best||sc>best.sc)) best={sc:sc,x:x}; });
  return best?best.x:null;
}

/* ========== sections ranking ========== */
function sectionRows(){
  if(!students.length) return null;
  var full=!!(lastInfo&&lastInfo.full&&lastInfo.total!==students.length&&lastInfo.bySec);
  var miss={}; students.forEach(function(x){ if(!x.g||x.s==="—") return; var k=x.g+"|"+x.s; miss[k]=(miss[k]||0)+1; });
  var keys=full?Object.keys(lastInfo.bySec).filter(function(k){return k.split("|")[0]&&k.split("|")[1]!=="—";}):Object.keys(miss);
  var sg=statsGrade(); if(sg) keys=keys.filter(function(k){return k.split("|")[0]===sg;});
  var rows=keys.map(function(k){ var p=k.split("|"), m=miss[k]||0, t=full?lastInfo.bySec[k]:null;
    return {k:k,g:p[0],s:p[1],label:clsShort(p[0],p[1]),missing:m,total:t,inn:t!=null?t-m:null,rate:t?(t-m)/t:null}; });
  rows.sort(function(a,b){ return full?((b.rate-a.rate)||(a.missing-b.missing)):(a.missing-b.missing); });
  rows.forEach(function(r,i){r.rank=i+1;});
  return {full:full,gradeLabel:statsGradeLabel(),rows:rows};
}
function renderSections(){
  var box=$("secRank"); if(!box) return; var sr=sectionRows();
  if(!sr||!sr.rows.length){ box.hidden=true; return; } box.hidden=false;
  $("secRankNote").textContent=sr.full?"من الأعلى دخولًا إلى الأقل":"مرتبة من الأقل غيابًا عن المنصة إلى الأكثر (الملف لا يحتوي على إجمالي الشعب)";
  var max=sr.full?1:Math.max.apply(null,sr.rows.map(function(r){return r.missing;}).concat([1]));
  $("secList").innerHTML=sr.rows.map(function(r){
    var w=sr.full?r.rate*100:(r.missing/max*100), cls=sr.full?(r.rate>=0.8?"hi":(r.rate>=0.6?"mid":"lo")):"lo";
    return '<div class="sr '+cls+'"><span class="rk">'+r.rank+'</span><span class="sl">'+esc(r.label)+'</span>'+
      '<span class="sb" aria-hidden="true"><i style="width:'+Math.max(0,Math.min(100,w)).toFixed(1)+'%"></i></span>'+
      '<span class="sp"><bdi>'+(sr.full?pctTxt(r.inn,r.total):r.missing)+'</bdi></span>'+
      '<span class="sc">'+(sr.full?r.inn+" من "+r.total:"لم يدخلوا")+'</span></div>';
  }).join("");
}
function sectionTable(){ var sr=sectionRows(); if(!sr) return null;
  return {full:sr.full,gradeLabel:sr.gradeLabel,headers:sr.full?["الترتيب","الصف والشعبة","إجمالي الطلاب","دخلوا","لم يدخلوا","نسبة الدخول"]:["الترتيب","الصف والشعبة","لم يدخلوا"],
    rows:sr.rows.map(function(r){return sr.full?[String(r.rank),r.label,String(r.total),String(r.inn),String(r.missing),pctTxt(r.inn,r.total)]:[String(r.rank),r.label,String(r.missing)];})};
}

/* ========== follow-up log (stored in this browser only) ========== */
var HKEY="nl-history-v1";
function hLoad(){var h=lsGet(HKEY,null); return (h&&h.days)?h:{v:1,days:{}};}
function hStore(h){ if(!lsSet(HKEY,h)){ toast("تعذّر الحفظ في السجل: مساحة المتصفح ممتلئة. صدّر نسخة احتياطية واحذف أيامًا قديمة."); return false; } return true; }
function hRecord(info){
  if(!$("hAuto").checked||!students.length) return;
  var key=(info&&info.refGreg)||$("day").value||ymd(new Date());
  var full=!!(info&&info.full&&info.total!==students.length);
  var byG=null, byS=null;
  if(full&&info.byGrade){ byG={}; Object.keys(info.byGrade).forEach(function(k){byG[k]=info.byGrade[k].total;}); }
  if(full&&info.bySec){ byS=info.bySec; }
  var h=hLoad();
  h.days[key]={full:full,total:full?info.total:null,byGrade:byG,bySec:byS,file:fileName,saved:Date.now(),
    list:students.map(function(x){return [x.id||"",x.name,x.g||"",x.s||"",x.phone||""];})};
  if(hStore(h)){ var dp=dayParts(new Date(key+"T12:00:00")); $("hSavedMsg").textContent="حُفظ يوم "+dp.day+" "+dp.date+" في السجل."; renderHist(); }
}
function hRange(){
  var v=$("hPer").value, t=new Date(), to=ymd(t), from, d;
  $("hCustom").hidden=v!=="custom";
  if(v==="week"){ d=new Date(t); d.setDate(d.getDate()-d.getDay()); from=ymd(d); }
  else if(v==="7"){ d=new Date(t); d.setDate(d.getDate()-6); from=ymd(d); }
  else if(v==="30"){ d=new Date(t); d.setDate(d.getDate()-29); from=ymd(d); }
  else if(v==="all"){ from="0000-00-00"; to="9999-99-99"; }
  else { from=$("hFrom").value||"0000-00-00"; to=$("hTo").value||"9999-99-99"; }
  return {from:from,to:to};
}
function hAgg(){
  var h=hLoad(), r=hRange();
  var keys=Object.keys(h.days).filter(function(k){return k>=r.from&&k<=r.to;}).sort();
  var days=keys.map(function(k){return Object.assign({key:k},h.days[k]);});
  var map={};
  days.forEach(function(d){ d.list.forEach(function(s){ var id=s[0]?("i"+s[0]):("n"+norm(s[1])+"|"+s[2]+"|"+s[3]);
    var o=map[id]||(map[id]={id:s[0],name:s[1],g:s[2],s:s[3],phone:s[4],n:0,dates:[]}); o.n++; o.dates.push(d.key); if(s[4]) o.phone=s[4]; }); });
  var series=days.map(function(d){
    var g={}; ["1","2","3"].forEach(function(k){ var t=d.byGrade&&d.byGrade[k]; var m=d.list.filter(function(s){return s[2]===k;}).length; g[k]={m:m,t:t||null}; });
    return {key:d.key,full:d.full,total:d.total,missing:d.list.length,g:g};
  });
  return {all:Object.keys(h.days).sort(),days:days,series:series,students:Object.keys(map).map(function(k){return map[k];}),range:r};
}
function shortDay(key){ var d=new Date(key+"T12:00:00"); try{ return d.toLocaleDateString("ar-SA-u-ca-islamic-umalqura-nu-latn",{day:"numeric",month:"numeric"}); }catch(e){ return key.slice(5); } }
function longDay(key){ var dp=dayParts(new Date(key+"T12:00:00")); return dp.day+" "+dp.date; }
function hCards(A){
  var fullDays=A.series.filter(function(s){return s.full&&s.total;});
  var rows=[["1","الصف "+GL["1"],GPC["1"]],["2","الصف "+GL["2"],GPC["2"]],["3","الصف "+GL["3"],GPC["3"]]];
  var cards=rows.map(function(r){
    var T=0,M=0,n=0; fullDays.forEach(function(s){ var x=s.g[r[0]]; if(x&&x.t){T+=x.t;M+=x.m;n++;} });
    if(!n){ var mm=0; A.series.forEach(function(s){mm+=s.g[r[0]].m;}); var dn=A.series.length||1;
      return {title:r[1],pc:r[2],v:Math.round(mm/dn),of:null,share:"",avgOnly:true,days:A.series.length}; }
    var v=Math.round(M/n), of=Math.round(T/n);
    return {title:r[1],pc:r[2],v:v,of:of,inn:of-v,pIn:pctTxt(T-M,T),pOut:pctTxt(M,T),days:n};
  });
  var T=0,M=0; fullDays.forEach(function(s){T+=s.total;M+=s.missing;});
  if(fullDays.length){ var n=fullDays.length; cards.push({title:"إجمالي المدرسة",pc:"#b3372f",isTotal:true,v:Math.round(M/n),of:Math.round(T/n),inn:Math.round((T-M)/n),pIn:pctTxt(T-M,T),pOut:pctTxt(M,T),days:n}); }
  return cards;
}
function drawChart(cv,A,w,hgt,forPrint){
  var dpr=forPrint?2:(window.devicePixelRatio||1); cv.width=w*dpr; cv.height=hgt*dpr; cv.style.width=forPrint?"":w+"px"; cv.style.height=forPrint?"":hgt+"px";
  var c=cv.getContext("2d"); c.setTransform(dpr,0,0,dpr,0,0);
  var cs=getComputedStyle((typeof HOST!=="undefined"&&HOST)||document.body), ink=forPrint?"#22302f":(cs.getPropertyValue("--ink")||"#16302f").trim()||"#16302f";
  var muted=forPrint?"#6b7678":"#8a9997", grid=forPrint?"#e3e8e6":"rgba(128,140,138,.25)";
  if(forPrint){ c.fillStyle="#fff"; c.fillRect(0,0,w,hgt); } else c.clearRect(0,0,w,hgt);
  var fam=getComputedStyle(document.body).fontFamily;
  var S=A.series; var pctMode=S.some(function(s){return s.full&&s.total;});
  var lines=[{k:"1",c:GPC["1"],n:"الصف "+GL["1"]},{k:"2",c:GPC["2"],n:"الصف "+GL["2"]},{k:"3",c:GPC["3"],n:"الصف "+GL["3"]},{k:"T",c:"#b3372f",n:"إجمالي المدرسة"}];
  function val(s,k){ if(k==="T") return pctMode?(s.total?s.missing/s.total*100:null):s.missing; var x=s.g[k]; return pctMode?(x.t?x.m/x.t*100:null):x.m; }
  var vals=[]; S.forEach(function(s){lines.forEach(function(l){var v=val(s,l.k); if(v!=null) vals.push(v);});});
  var max=vals.length?Math.max.apply(null,vals):10; var step=pctMode?10:Math.max(1,Math.ceil(max/5/5)*5); max=Math.max(step,Math.ceil(max/step)*step);
  var L=46,Rr=16,T=34,B=40, pw=w-L-Rr, ph=hgt-T-B;
  c.font="12px "+fam; c.textBaseline="middle"; c.strokeStyle=grid; c.lineWidth=1; c.fillStyle=muted; c.textAlign="right";
  for(var y=0;y<=max+0.001;y+=step){ var yy=T+ph-(y/max)*ph; c.beginPath(); c.moveTo(L,yy); c.lineTo(L+pw,yy); c.stroke(); c.fillText(String(Math.round(y))+(pctMode?"%":""),L-6,yy); }
  var n=S.length; function xAt(i){ return n<=1?L+pw/2:L+pw-(i/(n-1))*pw; } // right-to-left time axis
  c.textAlign="center"; c.textBaseline="top";
  var every=Math.max(1,Math.ceil(n/12));
  S.forEach(function(s,i){ if(i%every===0||i===n-1) c.fillText(shortDay(s.key),xAt(i),T+ph+8); });
  lines.forEach(function(l){
    var pts=S.map(function(s,i){var v=val(s,l.k); return v==null?null:[xAt(i),T+ph-(v/max)*ph];});
    c.strokeStyle=l.c; c.lineWidth=l.k==="T"?3:2; c.beginPath(); var started=false;
    pts.forEach(function(p){ if(!p){started=false;return;} if(!started){c.moveTo(p[0],p[1]);started=true;} else c.lineTo(p[0],p[1]); }); c.stroke();
    c.fillStyle=l.c; pts.forEach(function(p){ if(p){ c.beginPath(); c.arc(p[0],p[1],l.k==="T"?4:3,0,Math.PI*2); c.fill(); } });
  });
  // legend (top, right-aligned)
  c.textBaseline="middle"; c.textAlign="right"; c.font="12px "+fam; var lx=w-Rr;
  lines.forEach(function(l){ c.fillStyle=l.c; c.fillRect(lx-12,10,12,4); c.fillStyle=forPrint?"#22302f":ink; c.fillText(l.n,lx-18,12); lx-=c.measureText(l.n).width+40; });
  c.fillStyle=muted; c.textAlign="left"; c.fillText(pctMode?"نسبة عدم الدخول":"عدد غير الداخلين",L,12);
}
function chartImg(A){ var cv=document.createElement("canvas"); drawChart(cv,A,1000,330,true); return cv.toDataURL("image/png"); }
function hRepeat(A){
  var min=+$("hMin").value||1, g=$("hGrade").value;
  return A.students.filter(function(s){return s.n>=min&&(!g||s.g===g);}).sort(function(a,b){return (b.n-a.n)||(a.g||"9").localeCompare(b.g||"9")||coll.compare(String(a.s),String(b.s))||coll.compare(a.name,b.name);});
}
var hChartA=null, hChartW=0, hChartT=0;
function hChartDraw(){ if(!hChartA) return; var cv=$("hChart"), wrap=$("hChartWrap"); cv.style.width="0px"; var cw=wrap.clientWidth-18; if(!(cw>0)) cw=700; hChartW=cw;
  drawChart(cv,hChartA,Math.max(240,cw),Math.max(200,Math.min(320,cw*0.5)),false); cv.style.width="100%"; cv.style.height="auto"; }
window.addEventListener("resize",function(){ clearTimeout(hChartT); hChartT=setTimeout(function(){ var w=$("hChartWrap"); if(hChartA&&w&&Math.abs((w.clientWidth-18)-hChartW)>8) hChartDraw(); },150); });
function renderHist(){
  var box=$("hBody"); if(!box) return;
  var A=hAgg(), all=A.all;
  $("hCount").textContent=all.length?("الأيام المحفوظة في هذا الجهاز: "+all.length+" (من "+longDay(all[0])+" إلى "+longDay(all[all.length-1])+")"):"لا توجد أيام محفوظة بعد";
  $("hDays").innerHTML=all.slice().reverse().map(function(k){return '<span class="hd" data-k="'+k+'">'+esc(shortDay(k))+' <small>'+esc(new Date(k+"T12:00:00").toLocaleDateString("ar-SA",{weekday:"short"}))+'</small><button type="button" class="hx" data-del="'+k+'" aria-label="حذف هذا اليوم">✕</button></span>';}).join("");
  var hasAny=A.days.length>0;
  $("hEmpty").hidden=hasAny; box.hidden=!hasAny;
  ["hPdf","hDoc","hXls","hPrn"].forEach(function(id){$(id).disabled=!hasAny;});
  // keep threshold options in range
  var mx=Math.max(1,A.days.length), sel=$("hMin"), cur=hMinTouched?(+sel.value||1):Math.min(3,mx);
  sel.innerHTML=""; for(var i=1;i<=mx;i++){ var o=document.createElement("option"); o.value=i; o.textContent=i===1?"يوم واحد فأكثر":(i===2?"يومان فأكثر":i+" أيام فأكثر"); sel.appendChild(o); }
  sel.value=String(Math.min(cur,mx));
  if(!hasAny) return;
  $("hRangeTxt").textContent="الفترة: "+longDay(A.days[0].key)+" — "+longDay(A.days[A.days.length-1].key)+" ("+A.days.length+" "+(A.days.length===1?"يوم":"أيام")+")";
  $("hCards").innerHTML=hCards(A).map(function(c){
    if(c.avgOnly) return '<div class="gs" style="--pc:'+c.pc+'"><div class="gt"><span>'+esc(c.title)+'</span><span class="gtag">متوسط يومي</span></div><div class="gv">'+c.v+' <small>طالبًا لم يدخلوا</small></div><div class="gcnt"><span>خلال '+c.days+' أيام</span></div></div>';
    var pin=c.of?Math.max(0,Math.min(100,c.inn/c.of*100)):0;
    return '<div class="gs" style="--pc:'+c.pc+'"><div class="gt"><span>'+esc(c.title)+'</span><span class="gtag">متوسط يومي</span></div>'+
      '<div class="gv">'+c.v+' <small>من '+c.of+'</small></div>'+
      '<div class="gduo" aria-hidden="true"><i style="width:'+pin.toFixed(1)+'%;background:var(--green,#26794a)"></i><i style="width:'+(100-pin).toFixed(1)+'%;background:var(--red,#b3372f)"></i></div>'+
      '<div class="gpr"><div class="y">نسبة الدخول<b><bdi>'+esc(c.pIn)+'</bdi></b></div><div class="n">نسبة عدم الدخول<b><bdi>'+esc(c.pOut)+'</bdi></b></div></div>'+
      '<div class="gcnt"><span>خلال '+c.days+(c.days===1?" يوم":" أيام")+'</span></div></div>';
  }).join("");
  hChartA=A; hChartDraw();
  var rep=hRepeat(A), D=A.days.length, anyPh=rep.some(function(s){return s.phone;}), hasPh=anyPh&&colPref("phone"), hasWa=anyPh&&colPref("wa");
  $("hRepCount").textContent=rep.length+" طالبًا";
  $("hRep").innerHTML=rep.length?('<div class="scroll"><table><thead><tr><th>م</th><th>اسم الطالب</th><th>الصف والشعبة</th><th>أيام عدم الدخول</th>'+(hasPh?'<th>الجوال</th>':'')+(hasWa?'<th>تواصل</th>':'')+'</tr></thead><tbody>'+
    rep.map(function(s,i){ var msg=parentMsg(s,{"اليوم":"","التاريخ":"("+s.n+" من "+D+" أيام خلال الفترة من "+dayParts(new Date(A.days[0].key+"T12:00:00")).date+" إلى "+dayParts(new Date(A.days[D-1].key+"T12:00:00")).date+")"});
      return '<tr><td class="num">'+(i+1)+'</td><td class="name">'+esc(s.name)+'</td><td>'+esc(clsShort(s.g,s.s)||"—")+'</td><td><span class="cntpill'+(s.n>=Math.max(2,D)?" all":"")+'">'+s.n+' من '+D+'</span></td>'+
        (hasPh?'<td class="ltr">'+esc(val(s,"phone",0))+'</td>':'')+(hasWa?'<td>'+(intlPhone(s.phone)?'<a class="wa" target="_blank" rel="noopener" href="'+esc(waLink(s.phone,msg))+'">واتساب</a>':'—')+'</td>':'')+'</tr>'; }).join("")+'</tbody></table></div>')
    :'<div class="empty">لا يوجد طلاب بهذا العدد من الأيام في الفترة المحددة.</div>';
}
function hExportSpec(){
  var A=hAgg(), D=A.days.length, rep=hRepeat(A), min=+$("hMin").value||1;
  var hg=$("hGrade").value;
  var title="تقرير متابعة الدخول لمنصة مدرستي"+(hg&&GL[hg]?" — الصف "+GL[hg]:"")+" — من "+longDay(A.days[0].key)+" إلى "+longDay(A.days[D-1].key);
  var daily={title:"نسبة عدم الدخول يومًا بيوم",headers:["اليوم","الإجمالي","لم يدخلوا","نسبة عدم الدخول","الصف "+GL["1"],"الصف "+GL["2"],"الصف "+GL["3"]],
    rows:A.series.map(function(s){ function gp(k){var x=s.g[k]; return x.t?(x.m+" ("+pctTxt(x.m,x.t)+")"):String(x.m);}
      return [longDay(s.key),s.total!=null?String(s.total):"—",String(s.missing),s.total?pctTxt(s.missing,s.total):"—",gp("1"),gp("2"),gp("3")]; }),widths:[2.2,1,1,1.2,1.3,1.3,1.3]};
  var repT={title:"الطلاب الذين لم يدخلوا المنصة "+(min===1?"يومًا واحدًا فأكثر":(min===2?"يومين فأكثر":min+" أيام فأكثر"))+" خلال الفترة ("+rep.length+" طالبًا)",
    headers:["م","اسم الطالب","الصف والشعبة","أيام عدم الدخول"].concat(colPref("phone")?["الجوال"]:[]),widths:[.5,3,1.4,1.2].concat(colPref("phone")?[1.4]:[]),
    rows:rep.map(function(s,i){return [String(i+1),s.name,clsShort(s.g,s.s)||"—",s.n+" من "+D].concat(colPref("phone")?[val(s,"phone",0)]:[]);})};
  return {A:A,title:title,cards:hCards(A),daily:daily,rep:repT};
}
function hBase(){var A=hAgg(); return "تقرير-متابعة-المنصة-"+(A.days.length?A.days[0].key+"-إلى-"+A.days[A.days.length-1].key:ymd(new Date()));}


/* ---------- wiring: messages, log ---------- */
$("tplParent").value=lsGet("nl-tpl-parent",DEF_PARENT); $("tplTeacher").value=lsGet("nl-tpl-teacher",DEF_TEACHER);
$("tplParent").addEventListener("change",function(){lsSet("nl-tpl-parent",$("tplParent").value);render();});
$("tplTeacher").addEventListener("change",function(){lsSet("nl-tpl-teacher",$("tplTeacher").value);render();});
$("tplReset").addEventListener("click",function(){$("tplParent").value=DEF_PARENT;$("tplTeacher").value=DEF_TEACHER;lsSet("nl-tpl-parent",DEF_PARENT);lsSet("nl-tpl-teacher",DEF_TEACHER);render();toast("أُعيدت النصوص الافتراضية");});
$("repBody").addEventListener("click",function(e){var b=e.target.closest("[data-tcopy]"); if(!b) return; var g=lastGroups[+b.dataset.tcopy]; if(g) copyText(teacherMsg(g));});
$("hAuto").checked=lsGet("nl-hauto",true); $("hAuto").addEventListener("change",function(){lsSet("nl-hauto",$("hAuto").checked);});
$("hPer").value=lsGet("nl-hper","week");
var hMinTouched=false;
["hPer","hFrom","hTo","hMin","hGrade"].forEach(function(id){$(id).addEventListener("change",function(){ if(id==="hPer") lsSet("nl-hper",$("hPer").value); if(id==="hMin") hMinTouched=true; renderHist(); });});
$("hDays").addEventListener("click",function(e){var b=e.target.closest("[data-del]"); if(!b) return;
  if(!b.classList.contains("sure")){ b.classList.add("sure"); setTimeout(function(){b.classList.remove("sure");},3000); return; }
  var h=hLoad(); delete h.days[b.dataset.del]; hStore(h); renderHist(); toast("حُذف اليوم من السجل");});
var hRT; window.addEventListener("resize",function(){clearTimeout(hRT);hRT=setTimeout(renderHist,150);});
$("hBackup").addEventListener("click",function(){ var h=hLoad(); if(!Object.keys(h.days).length) return toast("السجل فارغ");
  save(new Blob([JSON.stringify(h)],{type:"application/json"}),"نسخة-سجل-متابعة-المنصة-"+ymd(new Date())+".json"); toast("حُفظت النسخة الاحتياطية"); });
$("hImport").addEventListener("change",function(){ var f=$("hImport").files[0]; $("hImport").value=""; if(!f) return;
  f.text().then(function(t){ var x=JSON.parse(t); if(!x||!x.days) throw 0; var h=hLoad(), n=0;
    Object.keys(x.days).forEach(function(k){ if(/^\d{4}-\d{2}-\d{2}$/.test(k)&&x.days[k]&&Array.isArray(x.days[k].list)){ h.days[k]=x.days[k]; n++; } });
    if(hStore(h)){ renderHist(); toast("استُرجع "+n+" يومًا إلى السجل"); } }).catch(function(){ toast("الملف ليس نسخة احتياطية صالحة من السجل"); }); });
function hWord(){ var X=hExportSpec(), d=dayDate(), body="";
  body+=wP($("school").value,{b:true,sz:26,after:0});
  body+=wP(X.title,{b:true,sz:30,center:true,before:120,after:120});
  var cc=X.cards.filter(function(c){return !c.avgOnly;});
  if(cc.length){ body+=wP("متوسط الدخول خلال الفترة",{b:true,sz:24,color:"0F6B63",after:60});
    body+=wTableRows(cc.map(function(c){return [c.title,String(c.of),String(c.inn),String(c.v),c.pIn,c.pOut];}),[{k:"g",t:"الصف"},{k:"id",t:"متوسط الإجمالي"},{k:"s",t:"متوسط من دخلوا"},{k:"phone",t:"متوسط من لم يدخلوا"},{k:"user",t:"نسبة الدخول"},{k:"last",t:"نسبة عدم الدخول"}]); }
  [X.daily,X.rep].forEach(function(t){ body+=wP(t.title,{b:true,sz:24,color:"0F6B63",before:240,after:60});
    body+=t.rows.length?wTableRows(t.rows,t.headers.map(function(h,i){return {k:["g","id","s","phone","user","last","name"][i],t:h};})):wP("لا يوجد.",{sz:22}); });
  makeDocxBlob(body).then(function(b){save(b,hBase()+".docx");toast("تم حفظ Word");}); }
function hExcel(){ var X=hExportSpec(), wb=XLSX.utils.book_new();
  [[X.daily,"الأيام"],[X.rep,"المتكرر غيابهم"]].forEach(function(p){ var t=p[0]; var ws=XLSX.utils.aoa_to_sheet([[$("school").value],[X.title],[t.title],[],t.headers].concat(t.rows));
    ws["!cols"]=t.headers.map(function(h,i){return {wch:i===1&&p[1]!=="الأيام"?32:18};}); XLSX.utils.book_append_sheet(wb,ws,p[1]); });
  wb.Workbook={Views:[{RTL:true}]}; XLSX.writeFile(wb,hBase()+".xlsx"); toast("تم حفظ Excel"); }
function hPdf(fmt){ var X=hExportSpec();
  if(!window.BalatReport) return toast("تقرير PDF متاح داخل موقع التقرير");
  window.BalatReport({title:X.title,cards:X.cards,img:chartImg(X.A),tables:[X.daily,X.rep],name:"تقرير_متابعة_المنصة"},fmt); }
$("hPdf").addEventListener("click",function(){hPdf("pdf");});
$("hPrn").addEventListener("click",function(){hPdf("print");});
$("hDoc").addEventListener("click",hWord);
$("hXls").addEventListener("click",hExcel);

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
  matchStudent:function(n,g,s){var x=matchStudent(n,g,s); return x?{name:x.name,id:x.id||"",phone:x.phone||"",g:x.g,s:x.s}:null;},
  parentLink:function(x){ return intlPhone(x.phone)?waLink(x.phone,parentMsg(x)):""; },
  sectionTable:sectionTable,
  gradeCards:gradeCards,
  gradeSummary:gradeSummaryRows,
  appendDocx:appendDocx,
  clear:function(){ students=[]; fileName=""; lastInfo=null; selSecs=new Set(); has={}; show={}; $("fileTxt").textContent="لم يُرفع ملف"; showErr(""); render();
    try{document.dispatchEvent(new CustomEvent("nologin:cleared"));}catch(e){} }
};

/* ---------- boot ---------- */
$("day").value=ymd(new Date());
$("school").value=lsGet("nl-school",$("school").value); $("optSplit").checked=lsGet("nl-split",true);
renderHist();
render();

})();
