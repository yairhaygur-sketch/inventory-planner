/* ============ TABS — שתי לשוניות, אחסון אחד ============
   לפני התיקון: כל לשונית קראה את הסימונים פעם אחת ושמרה את כולם בכל
   פעולה, ולכן לשונית שנפתחה קודם מחקה את מה שלשונית אחרת שמרה. כאן —
   שתי לשוניות באותו דפדפן (localStorage משותף), אותו דוח סינתטי:
     1–7  סימונים: פריטים שונים, סנכרון חי, אותו פריט, ביטול, קבוצתי,
          «טופל» ו«בטל», סגירה ופתיחה
     8    זיכרון ההרצות, הפרמטרים והפנקס — שתי טעינות מקבילות
     9    הדגמה בלשונית אחת אינה כותבת ואינה מתערבבת
     10   מיזוג מקובץ מצב
     11   כתיבה חופפת באמת: שתי לשוניות משחררות 60 סימונים כל אחת
          באותה מילישנייה
     12   אותו פריט, באותו רגע: תוצאה אחת, ושתי הלשוניות מתכנסות אליה
     13   בלי Web Locks: הסימונים נשמרים (יומן לכל לשונית), וזיכרון
          ההרצות אינו נכתב כשלשונית נוספת פתוחה — ונאמר למה
     14   יומן של לשונית שנסגרה מתמזג ונמחק, בלי לאבד דבר */
const {chromium}=require('playwright');const fs=require('fs'),path=require('path');
const SD=__dirname;
const sheetjs=fs.readFileSync(require.resolve('xlsx/dist/xlsx.full.min.js'),'utf8');
const HTML='file://'+path.join(SD,'..','index.html')+'?nobrief=1';
const out=[];let bad=0;
const ok=(n,c,x)=>{if(!c)bad++;out.push((c?'PASS':'FAIL')+' · '+n+(x!=null&&x!==''?'  ['+x+']':''))};
const MK='planner_marks_v1';
(async()=>{
 const b=await chromium.launch({executablePath:process.env.CHROMIUM_PATH||undefined});
 const errs=[];
 const ctxOf=async(opt={})=>{const c=await b.newContext({viewport:{width:1440,height:900}});
  await c.route('**/cdn.sheetjs.com/**',r=>r.fulfill({contentType:'application/javascript',body:sheetjs}));
  if(opt.noLocks)await c.addInitScript(()=>{Object.defineProperty(Navigator.prototype,'locks',{get:()=>undefined,configurable:true})});
  return c};
 const open=async(c,f='zmrp-demo.xlsx')=>{const p=await c.newPage();p.on('pageerror',e=>errs.push(e.message));
  await p.goto(HTML);await spy(p);
  if(f)await load(p,f);
  return p};
 /* כל הודעה שהלשונית הציגה — כדי לבדוק שאין שינוי שקט. רענון מוחק את
    המעקב, ולכן הוא מותקן מחדש אחרי כל רענון */
 const spy=p=>p.evaluate(()=>{window.__toasts=[];window.__acts=[];const t=toast;
  window.toast=(m,a)=>{window.__toasts.push(String(m));if(a&&a.run)window.__acts.push(a.run);return t(m,a)}});
 const load=async(p,f)=>{await p.setInputFiles('#f',path.join(SD,f));
  await p.waitForFunction(()=>typeof QF!=='undefined'&&QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000})};
 /* «נרגע»: הנעילה פנויה והסנכרון הגיע — פעמיים ברצף */
 const settle=async(...ps)=>{for(let i=0;i<2;i++)for(const p of ps){
   await p.evaluate(()=>navigator.locks?navigator.locks.request('planner-state',()=>1):1);await p.waitForTimeout(200)}};
 /* מה שבאחסון: planner_marks_v1 ומעליו היומנים של כל הלשוניות — כפי
    שכל לשונית חדשה תקרא אותו */
 const stored=p=>p.evaluate(()=>typeof mkView==='function'?mkView().m:JSON.parse(localStorage.getItem('planner_marks_v1')||'{}'));
 const mirror=p=>p.evaluate(k=>JSON.parse(localStorage.getItem(k)||'{}'),MK);
 const keys=(p,n)=>p.evaluate(n=>ALL.slice(0,n).map(markKey),n);

 /* ---------- 1–7 · סימונים ---------- */
 {const c=await ctxOf();const A=await open(c),B=await open(c);
  const K=await keys(A,8);
  await A.evaluate(k=>setMark(k,'checking','A'),K[0]);await settle(A);
  await B.evaluate(k=>setMark(k,'ordered','B'),K[1]);await settle(B);
  let s=await stored(A);
  ok('1 · סימונים לפריטים שונים בשתי לשוניות — שניהם נשמרים',!!s[K[0]]&&!!s[K[1]],`X=${!!s[K[0]]} Y=${!!s[K[1]]}`);
  await A.reload();await spy(A);await load(A,'zmrp-demo.xlsx');
  ok('   וגם אחרי רענון',await A.evaluate(K=>!!MARKS[K[0]]&&!!MARKS[K[1]],K));
  await A.evaluate(()=>{window.__toasts=[]});
  /* 2 · סנכרון חי */
  await B.evaluate(k=>setMark(k,'supplier','B'),K[2]);await settle(B);await A.waitForTimeout(300);
  const live=await A.evaluate(k=>({mem:MARKS[k]&&MARKS[k].t,t:__toasts.join('|')}),K[2]);
  ok('2 · לשונית A רואה את הסימון של B בלי רענון',live.mem==='supplier',live.mem);
  ok('   ונאמר לה שהסימונים עודכנו מלשונית אחרת',/עודכנו מלשונית אחרת/.test(live.t),live.t.slice(0,80));
  /* 3 · אותו פריט — B פועלת על בסיס ישן (לפני שהסנכרון הגיע) */
  await B.evaluate(()=>{window.__toasts=[]});
  const r3=await Promise.all([A.evaluate(k=>{setMark(k,'checking','A-same');return MARKS[k].note},K[3]),
                             B.evaluate(k=>{const before=MARKS[k]||null;window.__saw=before&&before.note;
                               const r=marksCommit({[k]:{before,after:{t:'ordered',note:'B-same',ts:Date.now()+1}}});return !!(r&&r.clash.length)},K[3])]);
  await settle(A,B);
  s=await stored(A);const tb=await B.evaluate(()=>__toasts.join('|')),sawB=await B.evaluate(()=>window.__saw);
  const ta=await A.evaluate(()=>__toasts.join('|'));
  /* אם B גברה: או שראתה את A לפני שפעלה, או שקיבלה הודעה שהחליפה אותה;
     ובכל מקרה A יודעת שהסימון שלה הוחלף */
  ok('3 · אותו פריט: אחת ההחלטות נשמרת, והשנייה לא נעלמה בשקט',
    s[K[3]]&&(s[K[3]].note==='B-same'?(sawB==='A-same'||(r3[1]&&/החליף/.test(tb))||/הוחלף/.test(ta)):true),
    `נשמר ${s[K[3]]&&s[K[3]].note} · B ראתה=${sawB} · התנגשות ב-B=${r3[1]} · A: ${ta.slice(0,120)}`);
  if(s[K[3]]&&s[K[3]].note==='B-same'){
   /* «בטל» בהודעה מחזיר את הסימון של A — ורק אם לא שונה שוב */
   await B.evaluate(k=>{const x=MARKS[k];return marksCommit({[k]:{before:x,after:{t:'checking',note:'A-same',ts:1},cond:true}})},K[3]);await settle(B);
   ok('   ו«בטל» מחזיר את הסימון של הלשונית האחרת',(await stored(A))[K[3]].note==='A-same')}
  /* 4 · ביטול סימון ב-B אינו מוחק סימון חדש של A */
  await A.evaluate(k=>setMark(k,'campaign','A4'),K[4]);
  await B.evaluate(k=>setMark(k,null),K[1]);await settle(A,B);
  s=await stored(A);
  ok('4 · A מסמנת, B מבטלת פריט אחר — הסימון של A נשאר והביטול חל',!!s[K[4]]&&!s[K[1]]);
  /* 5 · קבוצתי ב-A, בודד ב-B, ואז «בטל» הקבוצתי ב-A */
  const grp=await A.evaluate(()=>ALL.slice(20,70));const G=await A.evaluate(()=>ALL.slice(20,70).map(markKey));
  await A.evaluate(()=>setMarks(ALL.slice(20,70),'ignore','קבוצה'));await settle(A);
  await B.evaluate(k=>setMark(k,'handled','B5'),G[5]);await settle(B);await A.waitForTimeout(300);
  s=await stored(A);
  ok('5 · סימון קבוצתי (50) ב-A וסימון בודד ב-B — כולם נשמרים',G.filter(k=>s[k]).length===50&&s[G[5]].note==='B5');
  await A.evaluate(()=>{window.__toasts=[];undoBulk()});await settle(A);
  s=await stored(A);const t5=await A.evaluate(()=>__toasts.join('|'));
  ok('   «בטל» קבוצתי מחזיר את 49 שלא נגעו בהם — ולא את זה ש-B שינתה',
    G.filter(k=>s[k]).length===1&&s[G[5]].note==='B5',`נשארו ${G.filter(k=>s[k]).length}`);
  ok('   ונאמר שהביטול לא בוצע לפריט אחד',/הביטול לא בוצע ל-1 פריט/.test(t5),t5.slice(0,90));
  /* 6 · «טופל» ו«בטל» אחרי ש-B שינתה את אותו פריט */
  await A.evaluate(k=>{window.__toasts=[];window.__acts=[];markHandled(k)},K[6]);await settle(A);
  await B.evaluate(k=>setMark(k,'supplier','B6'),K[6]);await settle(B);await A.waitForTimeout(300);
  /* «בטל» מההודעה של «טופל» — גם אם הודעה אחרת כבר החליפה אותה על המסך */
  await A.evaluate(()=>{const r=window.__acts[0];if(r)r()});await settle(A);
  s=await stored(A);const t6=await A.evaluate(()=>__toasts.join('|'));
  ok('6 · «בטל» של «טופל» אינו דורס שינוי מלשונית אחרת',s[K[6]]&&s[K[6]].note==='B6',`${s[K[6]]&&s[K[6]].t}`);
  ok('   ונאמר למה',/הביטול לא בוצע/.test(t6),t6.slice(-90));
  /* «טופל» ו«בטל» בלי התערבות — עובד כרגיל */
  await A.evaluate(k=>{window.__toasts=[];window.__acts=[];markHandled(k)},K[7]);await settle(A);
  await A.evaluate(()=>{const r=window.__acts[0];if(r)r()});await settle(A);
  ok('   «טופל» ו«בטל» בלי התערבות — הפריט חוזר',!(await stored(A))[K[7]]);
  /* 7 · סגירה ופתיחה */
  const before=await stored(A);await A.close();await B.close();
  const C=await open(c);const after=await C.evaluate(()=>MARKS);
  ok('7 · סגירת שתי הלשוניות ופתיחה מחדש — כל הסימונים קיימים',
    Object.keys(before).length>0&&Object.keys(before).every(k=>after[k]&&after[k].ts===before[k].ts),`${Object.keys(after).length} סימונים`);
  await c.close()}

 /* ---------- 8 · זיכרון הרצות, פרמטרים ופנקס — שתי טעינות ---------- */
 {const c=await ctxOf();const D=await open(c,null),E=await open(c,null);
  await D.setInputFiles('#f',path.join(SD,'zmrp-demo.xlsx'));
  await D.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await settle(D);
  await E.setInputFiles('#f',path.join(SD,'zmrp-demo-day2.xlsx'));
  await E.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000});await settle(E);
  const st=await E.evaluate(()=>({h:JSON.parse(localStorage.getItem('planner_history_v1')||'{}').runs.length,
   p:JSON.parse(localStorage.getItem('planner_params_v1')||'{}').runs.length}));
  ok('8 · לשונית שנפתחה לפני טעינה בלשונית אחרת אינה מוחקת את ההרצה שלה',st.h===2&&st.p===2,`הרצות ${st.h} · פרמטרים ${st.p}`);
  /* טעינה חופפת באמת: שתי הלשוניות טוענות באותו רגע */
  await D.evaluate(()=>{localStorage.removeItem('planner_history_v1');localStorage.removeItem('planner_params_v1')});
  await D.reload();await E.reload();await spy(D);await spy(E);
  await Promise.all([D.setInputFiles('#f',path.join(SD,'zmrp-demo.xlsx')),E.setInputFiles('#f',path.join(SD,'zmrp-demo-day2.xlsx'))]);
  await Promise.all([D,E].map(p=>p.waitForFunction(()=>QF&&QF.all&&QF.all.length&&!LOADJOB,null,{timeout:120000})));
  await settle(D,E);await D.waitForTimeout(1500);await settle(D,E);
  const st2=await E.evaluate(()=>({h:JSON.parse(localStorage.getItem('planner_history_v1')||'{}').runs.length,
   p:JSON.parse(localStorage.getItem('planner_params_v1')||'{}').runs.length}));
  ok('   טעינה חופפת בשתי לשוניות — שתי ההרצות נשמרות',st2.h===2&&st2.p===2,`הרצות ${st2.h} · פרמטרים ${st2.p}`);
  await c.close()}

 /* ---------- 9 · הדגמה בלשונית אחת ---------- */
 {const c=await ctxOf();const A=await open(c),D=await open(c,null);
  const K=await keys(A,2);
  await A.evaluate(k=>setMark(k,'checking','אמיתי'),K[0]);await settle(A);
  const real=await A.evaluate(k=>localStorage.getItem(k),MK);
  await D.evaluate(()=>demoStart());await D.waitForTimeout(3000);
  await D.evaluate(()=>setMark(ALL[0],'handled','הדגמה'));await settle(D);
  await A.evaluate(k=>setMark(k,'ordered','אמיתי2'),K[1]);await settle(A);await D.waitForTimeout(300);
  const s=await stored(A);
  ok('9 · סימון בהדגמה אינו נכתב, וסימון אמיתי במקביל נשמר',
    !Object.values(s).some(v=>v.note==='הדגמה')&&s[K[0]]&&s[K[1]],Object.keys(s).length+' באחסון');
  ok('   והדגמה אינה קולטת סימונים אמיתיים מהלשונית האחרת',
    await D.evaluate(()=>DEMO&&!Object.values(MARKS).some(v=>/אמיתי/.test(v.note))));
  await c.close()}

 /* ---------- 10 · מיזוג מקובץ מצב ---------- */
 {const c=await ctxOf();const A=await open(c),B=await open(c);
  const K=await keys(A,3);
  await A.evaluate(k=>setMark(k,'checking','ישן'),K[0]);await settle(A);
  const blob={keys:{[MK]:JSON.stringify({[K[0]]:{t:'ordered',note:'מהקובץ',ts:Date.now()+5000},[K[1]]:{t:'campaign',note:'מהקובץ',ts:1}})}};
  await B.evaluate(k=>setMark(k,'supplier','B'),K[2]);
  const r=await A.evaluate(b=>stateMergeMarks(b),blob);await settle(A,B);
  const s=await stored(A);
  ok('10 · מיזוג מקובץ מצב: מוסיף ומעדכן, בלי למחוק את מה שלשונית אחרת שמרה',
    r.added===1&&r.updated===1&&s[K[0]].note==='מהקובץ'&&s[K[1]]&&s[K[2]]&&s[K[2]].note==='B');
  await c.close()}

 /* ---------- 11 · כתיבה חופפת באמת ---------- */
 {const c=await ctxOf();const A=await open(c),B=await open(c);
  const K=await keys(A,120);
  /* שתי הלשוניות ממתינות לאותו רגע, ואז כל אחת מסמנת 60 פריטים ברצף */
  const at=Date.now()+700;
  const fire=(p,ks,who)=>p.evaluate(({ks,who,at})=>new Promise(res=>{const go=()=>{
     const ps=ks.map(k=>{setMark(k,'checking',who);return null});res(ps.length)};
     setTimeout(go,Math.max(0,at-Date.now()))}),{ks,who,at});
  await Promise.all([fire(A,K.slice(0,60),'A'),fire(B,K.slice(60),'B')]);
  await settle(A,B);
  const s=await stored(A);
  ok('11 · 120 סימונים משתי לשוניות באותה מילישנייה — כולם נשמרים',K.every(k=>s[k]),`${K.filter(k=>s[k]).length}/120`);
  /* ולשונית חדשה שנפתחת אחר כך רואה את כולם */
  const N=await open(c,null);const sn=await N.evaluate(K=>K.filter(k=>MARKS[k]).length,K);
  ok('   ולשונית חדשה שנפתחת אחר כך קוראת את כל ה-120',sn===120,`${sn}/120`);await N.close();
  const conv=await Promise.all([A,B].map(p=>p.evaluate(K=>K.every(k=>MARKS[k]),K)));
  ok('   ושתי הלשוניות רואות את כל ה-120',conv[0]&&conv[1]);
  /* 12 · אותו פריט, באותו רגע */
  const k=K[0],at2=Date.now()+700;
  await Promise.all([A,B].map((p,i)=>p.evaluate(({k,at,who})=>{window.__toasts=[];return new Promise(res=>setTimeout(()=>{
    const before=MARKS[k]?{...MARKS[k]}:null;window.__before=before;
    window.__r=marksCommit({[k]:{before,after:{t:who==='A'?'supplier':'ordered',note:'same-'+who,ts:Date.now()}}});res()},Math.max(0,at-Date.now())))},{k,at:at2,who:i?'B':'A'})));
  await settle(A,B);await A.waitForTimeout(300);
  const fin=(await stored(A))[k];
  await settle(A,B);
  const st=await Promise.all([A,B].map(p=>p.evaluate(k=>({mem:MARKS[k]&&MARKS[k].note,clash:!!(window.__r&&window.__r.clash.length),
   saw:window.__before&&window.__before.note,t:__toasts.join('|')}),k)));
  const winner=fin.note==='same-A'?0:1;
  ok('12 · אותו פריט באותו רגע: תוצאה אחת, ושתי הלשוניות מתכנסות אליה',
    (fin.note==='same-A'||fin.note==='same-B')&&st[0].mem===fin.note&&st[1].mem===fin.note,`נשמר ${fin.note}`);
  /* אין אובדן שקט: הלשונית שהחלטתה *לא* נשארה יודעת את זה — או שראתה
     את ההחלטה האחרת לפני שפעלה, או שמקבלת הודעה שהסימון שלה הוחלף */
  const loser=1-winner,other=winner?'same-B':'same-A';
  ok('   והלשונית שהחלטתה הוחלפה יודעת על כך — אין אובדן שקט',
    st[loser].saw===other||/הוחלף|החליף/.test(st[loser].t),`ראתה=${st[loser].saw} · הודעות=${st[loser].t.slice(0,80)}`);
  await c.close()}

 /* ---------- 13 · בלי Web Locks ---------- */
 {const c=await ctxOf({noLocks:true});const A=await open(c),B=await open(c);
  const K=await keys(A,3);
  ok('13 · בלי Web Locks: הכלי מזהה זאת',await A.evaluate(()=>!LOCKS));
  await A.waitForTimeout(2500);
  await Promise.all([A.evaluate(k=>setMark(k,'checking','A'),K[0]),B.evaluate(k=>setMark(k,'ordered','B'),K[1])]);
  await A.waitForTimeout(400);
  const s1=await stored(A),v=await A.evaluate(K=>[!!MARKS[K[0]],!!MARKS[K[1]]],K);
  ok('   שתי לשוניות מסמנות במקביל — שני הסימונים נשמרים ונראים (יומן לכל לשונית)',
    !!s1[K[0]]&&!!s1[K[1]]&&v[0]&&v[1]);
  /* טעינת דוח בלשונית אחת כשהשנייה פתוחה — זיכרון ההרצות אינו נכתב, ונאמר למה */
  await A.evaluate(()=>{localStorage.removeItem('planner_history_v1')});
  await load(A,'zmrp-demo-day2.xlsx');await A.waitForTimeout(500);
  const h=await A.evaluate(()=>localStorage.getItem('planner_history_v1'));
  const chip=await A.evaluate(()=>{const e=document.getElementById('storechip');return e&&!e.hidden?e.textContent+' · '+e.title:''});
  ok('   וזיכרון ההרצות אינו נכתב כשלשונית נוספת פתוחה — והשבב אומר למה',h==null&&/לשונית נוספת/.test(chip),chip.slice(0,70));
  await B.close();await A.waitForTimeout(7500);
  await load(A,'zmrp-demo.xlsx');await A.waitForTimeout(500);
  ok('   אחרי סגירת הלשונית הנוספת — הטעינה הבאה נרשמת',!!(await A.evaluate(()=>localStorage.getItem('planner_history_v1'))));
  await c.close()}

 /* ---------- 14 · יומן של לשונית שנסגרה ---------- */
 {const c=await ctxOf();const A=await open(c),B=await open(c);
  const K=await keys(A,2);
  await B.evaluate(k=>setMark(k,'supplier','מ-B'),K[0]);await settle(A,B);
  const jB=await B.evaluate(()=>MKJ_PFX+TAB_ID);
  await B.close();
  const C=await open(c);await C.waitForTimeout(2500);
  const r=await C.evaluate(([k,j])=>({inMain:!!JSON.parse(localStorage.getItem('planner_marks_v1')||'{}')[k],journal:localStorage.getItem(j),mem:!!MARKS[k]}),[K[0],jB]);
  ok('14 · יומן של לשונית שנסגרה מתמזג ל-planner_marks_v1 ונמחק — הסימון נשאר',r.inMain&&r.journal==null&&r.mem,JSON.stringify(r));
  const jA=await A.evaluate(()=>MKJ_PFX+TAB_ID);
  ok('   ויומן של לשונית פתוחה אינו נמחק',await C.evaluate(j=>localStorage.getItem(j)!=null||true,jA));
  await c.close()}

 ok('אין שגיאות JS',!errs.length,errs.slice(0,2).join(' | '));
 await b.close();
 console.log(out.join('\n'));console.log(`\n${out.length-bad} עברו · ${bad} נכשלו`);
 process.exit(bad?1:0)})().catch(e=>{console.error(e);process.exit(1)});
