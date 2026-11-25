/* ============ TP_SEARCH + TP_NOTE_BOX — תצוגת 4 הנושאים (בקשות המעתד 25.11, אושרו) ============
   1. החיפוש: אותו מנוע של הכלי המלא — מק״ט שהודבק מ-SAP עם תווים נסתרים (תווי כיוון, רווח ברוחב
      אפס, BOM, טאב, רווח קשיח) או עם מקף/רווח באמצע נמצא.
   2. רשימה נפתחת, כמו בכלי המלא: כל פריטי הדוח, הנושא של כל תוצאה; Enter / לחיצה פותחים את הכרטיס
      בנושא הראשון של הפריט (או ב«כל הפריטים»); ↑↓ מעבר; Esc סגירה.
   3. ההערה מ«טופל» הקודם — תיבה בולטת בראש הכרטיס, מעל מצב המלאי. */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname,XL=path.join(SD,'zmrp-demo.xlsx');
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const out=[];let bad=0;const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
(async()=>{const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const ctx=await b.newContext({viewport:{width:1440,height:900}});
 await ctx.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
 const p=await ctx.newPage();const errs=[];p.on('pageerror',e=>errs.push(e.message));
 await p.goto('file://'+path.join(SD,'..','index.html')+'?topics');
 await p.evaluate(()=>{try{localStorage.clear()}catch(_){}});
 await p.setInputFiles('#f',XL);
 await p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await p.waitForSelector('#lite table');

 /* פריט מטרה: הראשון בנושא «מלאי 0 / צפוי להיגמר» */
 const PN=await p.evaluate(()=>{TP.tab='short';TP.cat={};tpRender();return TP.list[0].pn});
 const type=async v=>{await p.fill('#lite [data-tp="q"]',v);await p.waitForTimeout(450)};
 const st=()=>p.evaluate(()=>({list:(TP.list||[]).map(r=>r.pn),open:!!document.querySelector('#tpsug.open'),
  sug:[...document.querySelectorAll('#tpsug .sgi')].map(x=>x.querySelector('.p').textContent),tab:TP.tab,card:TP.open?TP.open.pn:null}));

 /* 1 · הדבקה מ-SAP */
 for(const [n,v] of [['תו כיוון (LRM) בהתחלה','‎'+PN],['רווח ברוחב אפס באמצע',PN.slice(0,3)+'​'+PN.slice(3)],['BOM + טאב בסוף','﻿'+PN+'\t'],
   ['רווח קשיח מסביב',' '+PN+' '],['מקף באמצע',PN.slice(0,3)+'-'+PN.slice(3)],['רווח באמצע',PN.slice(0,3)+' '+PN.slice(3)]]){
  await type(v);const S=await st();
  ok(`הדבקה עם ${n}: הפריט בטבלה ובראש הרשימה הנפתחת`,S.list.includes(PN)&&S.sug[0]===PN,`טבלה ${S.list.length} · ראשון ${S.sug[0]}`)}

 /* 2 · הרשימה הנפתחת */
 /* שאילתה עם כמה תוצאות: המילה הראשונה בתיאור של פריט שהיא מופיעה בעוד פריטים */
 const Q2=await p.evaluate(()=>{for(const r of ALL){const w=String(r.desc||'').split(/\s+/).find(x=>x.length>=3);if(!w)continue;
   const n=searchAll(parseQ(w),8).rows.length;if(n>=3)return w}return null});
 await type(Q2);let S=await st();
 const H=await p.evaluate(()=>({hdr:(document.querySelector('#tpsug .sgh')||{}).textContent||'',topic:[...document.querySelectorAll('#tpsug .sgi .c i')].map(x=>x.textContent)}));
 ok('הקלדה פותחת רשימה נפתחת — עד 8 תוצאות, עם מספר ההתאמות בדוח',S.open&&S.sug.length>0&&S.sug.length<=8&&/פריטים מתאימים בדוח/.test(H.hdr),`${S.sug.length} · ${H.hdr}`);
 ok('לכל תוצאה — הנושא שלה (או «לא בנושא»)',H.topic.length>=S.sug.length,H.topic.slice(0,3).join(' | '));
 const G=await p.evaluate(q=>{const Qy=parseQ(q);return searchAll(Qy,8).rows.map(r=>r.pn)},Q2);
 ok('אותן תוצאות ואותו סדר כמו בחיפוש של הכלי המלא',JSON.stringify(G)===JSON.stringify(S.sug));
 await p.keyboard.press('ArrowDown');await p.keyboard.press('ArrowDown');
 const sel=await p.evaluate(()=>{const o=document.querySelector('#tpsug .sgi.on');return o?o.querySelector('.p').textContent:null});
 ok('↓↓ מסמן את התוצאה השנייה',sel===S.sug[1],sel);
 await p.keyboard.press('Enter');await p.waitForTimeout(300);
 const E=await st();const want=await p.evaluate(pn=>{const r=ALL.find(x=>x.pn===pn);return tpSugTopic(r)},S.sug[1]);
 ok('Enter: הכרטיס נפתח על התוצאה המסומנת, בנושא הראשון שלה, והרשימה נסגרת',E.card===S.sug[1]&&E.tab===want&&!E.open,`${E.card} · ${E.tab}`);
 await p.evaluate(()=>{tpdClose();TP.tab='short';TP.cat={};tpRender()});
 await type(PN);await p.keyboard.press('Escape');S=await st();
 ok('Esc סוגר את הרשימה, והחיפוש בטבלה נשאר',!S.open&&S.list.includes(PN));
 await type(PN);await p.click('#tpsug .sgi');await p.waitForTimeout(300);S=await st();
 ok('לחיצה על תוצאה פותחת את הכרטיס שלה',S.card===PN&&!S.open,S.card);
 await p.evaluate(()=>tpdClose());
 await type('zzzz-אין-כזה');const N=await p.evaluate(()=>(document.querySelector('#tpsug .sgn')||{}).textContent||'');
 ok('אין התאמה: «אין פריט מתאים בדוח»',/אין פריט מתאים/.test(N),N);
 /* פריט שטופל — נפתח ב«כל הפריטים» */
 await p.evaluate(pn=>{setMark(ALL.find(x=>x.pn===pn),'handled','')},PN);
 await type(PN);const D=await p.evaluate(()=>[...document.querySelectorAll('#tpsug .sgi .c i')].map(x=>x.textContent));
 ok('פריט שטופל: «✓ טופל» ברשימה',D.includes('✓ טופל'),D.join(' | '));
 await p.keyboard.press('Enter');await p.waitForTimeout(300);S=await st();
 ok('…ונפתח בלשונית «כל הפריטים»',S.card===PN&&S.tab==='all',S.tab);
 await p.evaluate(()=>tpdClose());await type('');

 /* 3 · תיבת ההערה */
 const B=await p.evaluate(pn=>{const r=ALL.find(x=>x.pn===pn),k=markKey(r);setMark(r,'handled','הזמנה דחופה מהספק');MARKS[k]={...MARKS[k],ts:Date.now()-8*864e5};apply();
  TP.tab='short';TP.cat={};tpRender();tpdShow(r);const db=document.querySelector('#tpd .db'),f=db.firstElementChild,nt=f&&f.querySelector('.nt');
  const v={cls:f&&f.className,next:f&&f.nextElementSibling&&f.nextElementSibling.className,nt:nt?nt.textContent:'',fs:nt?parseFloat(getComputedStyle(nt).fontSize):0,
   fw:nt?+getComputedStyle(nt).fontWeight:0,bg:f?getComputedStyle(f).backgroundColor:'',meta:(f&&f.querySelector('.mt')||{}).textContent||''};tpdClose();return v},PN);
 ok('ההערה — תיבה ראשונה בכרטיס, מעל «מצב המלאי»',/tpprev/.test(B.cls)&&/tpstock/.test(B.next),`${B.cls} → ${B.next}`);
 ok('ההערה עצמה גדולה ומודגשת (16px), עם 📝',B.nt.includes('הזמנה דחופה מהספק')&&B.nt.includes('📝')&&B.fs>=16&&B.fw>=700,`${B.fs}px ${B.fw}`);
 ok('רקע צבעוני, ומתחת התאריך ו«וחזר לרשימה»',B.bg!=='rgba(0, 0, 0, 0)'&&/סומן «טופל» ב-/.test(B.meta)&&/וחזר לרשימה/.test(B.meta),B.bg+' · '+B.meta);

 /* 4 · TP_CARD_FOB — מחיר FOB בכותרת מצב המלאי */
 await p.setViewportSize({width:1366,height:768});
 const F=await p.evaluate(()=>{const withP=ALL.find(r=>!r.priceMissing&&r.price>0&&tpTopicsOf(r).length),noP=ALL.find(r=>(r.priceMissing||!(r.price>0))&&tpTopicsOf(r).length);
  const card=r=>{if(!r)return null;TP.tab=tpTopicsOf(r)[0];TP.cat={};tpRender();tpdShow(r);const h=document.querySelector('#tpd .tpstock h3'),f=h&&h.querySelector('.tpfob');
   const v={h:h?h.textContent:'',f:f?f.textContent:'',miss:!!f&&f.classList.contains('miss'),fits:!!h&&h.scrollWidth<=h.clientWidth+1,price:r.price,cur:r.currency};tpdClose();return v};
  return {a:card(withP),b:card(noP)}});
 ok('בכותרת «מצב המלאי» — מחיר FOB ליחידה, במטבע של הפריט',!!F.a&&/^מצב המלאי/.test(F.a.h)&&/מחיר FOB/.test(F.a.f)&&!F.a.miss&&F.a.f.replace(/[^\d.]/g,'').length>0,F.a&&F.a.f);
 ok('פריט בלי מחיר: «מחיר FOB חסר»',!F.b||(/מחיר FOB חסר/.test(F.b.f)&&F.b.miss),F.b?F.b.f:'אין פריט כזה בנושאים');
 ok('ב-1366: הכותרת לא גולשת',F.a&&F.a.fits);

 ok('אין שגיאות JS',errs.length===0,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(bad?`\n${bad} כשלים`:`\nכל ${out.length} הבדיקות עברו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
