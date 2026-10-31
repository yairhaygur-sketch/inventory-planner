/* ============ RATE_DECAY · פריטים סינתטיים לבדיקות הדעיכה ============
   הדוח הסינתטי (zmrp-demo.xlsx) אינו מכיל אף פריט בנתיב 3 השנים שהביקוש
   שלו דעך — נמדד: 0 פריטים זזים בו. לכן השוואת הייצוא מול main לא רואה
   את RATE_DECAY בכלל, וצריך פריטים שנבנים במפורש.

   כאן: הדוח הסינתטי כפי שהוא, ועוד ארבעה פריטים בסופו. כל השדות
   מומצאים; המק״טים מתחילים ב-DECAY- כדי שאי אפשר יהיה לבלבל אותם עם
   מק״ט אמיתי. ההרכב זהה לפריט המבחן PN-07 מהדוח האמיתי:
   1 יח׳ ב-11 חודשים, 6 ב-2025, 202 ב-2024, LT 120 יום, רמת שירות 80,
   SS ו-ssMin 11 — ולכן ההמלצה הצפויה ROP 12, SS 11 (היה 37). */
const XLSX=require('xlsx'),path=require('path'),fs=require('fs'),os=require('os');
const SD=path.join(__dirname,'..');
const MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const dstr=d=>{const x=new Date(d);return `${String(x.getDate()).padStart(2,'0')}.${String(x.getMonth()+1).padStart(2,'0')}.${x.getFullYear()}`};
const DAY=864e5;
/* פריט אחד: ערכים לפי שם עמודה; כל השאר מועתק משורת תבנית */
function item(hdr,tpl,o){
 const r=tpl.slice();const set=(k,v)=>{const i=hdr.indexOf(k);if(i<0)throw new Error('עמודה חסרה: '+k);r[i]=v};
 set('מק"ט מוביל',o.pn);set('תיאור חומר',o.pn);set('תיאור חומר2','Part '+o.pn);
 set('סטטוס חומר','01');set('סוג MRP','ND');set('רמת שרות',o.srv);
 set('מלאי בטחון',o.ss);set('נק.הז.מחדש',o.rop);set('מל.בט.מינ.',o.ssMin);
 set('אספ.מתוכנ.',o.lt);set('זמ.עב.קבלת',0);
 set('מחיר FOB',o.price);set('מטבע FOB',o.cur);
 set('סה"כ מלאי',o.free);set('מלאי פנוי',o.free);set('מלאי מרלוג',o.free);
 set('הז. רכש',0);set('בהעברה',0);set('אספקות פת.',0);set('כמות בהז.פ',0);
 set('צר.השנה',o.y[0]);set('צר.שנה-1',o.y[1]);set('צר.שנה-2',o.y[2]);set('צר.החודש',0);
 MON.forEach((k,i)=>set(k,o.m[i]||0));
 set("תאר' מכירה",dstr(Date.now()-o.saleAgo*DAY));set("תאר' כניסה",dstr(Date.now()-900*DAY));
 return r}
/* הפריטים והתוצאה הצפויה. «decay» = האם RATE_DECAY אמור לפעול */
const ITEMS=[
 /* דעיכה חדה, USD — המבנה של PN-07 */
 {pn:'DECAY-USD',srv:80,ss:11,rop:18,ssMin:11,lt:120,price:228.91,cur:'USD',free:21,
  y:[1,6,202],m:[0,0,0,1],saleAgo:120,decay:true,expROP:12,expSS:11},
 /* דעיכה חדה, ILS — רצפה קטנה יותר, אותה התנהגות */
 {pn:'DECAY-ILS',srv:80,ss:3,rop:10,ssMin:3,lt:120,price:900,cur:'ILS',free:2,
  y:[1,4,150],m:[0,0,0,0,0,1],saleAgo:180,decay:true,expROP:4,expSS:3},
 /* נתיב 3 השנים בלי דעיכה — ROP/SS בדיוק כמו קודם */
 {pn:'DECAY-KEEP',srv:80,ss:1,rop:2,ssMin:1,lt:120,price:50,cur:'USD',free:4,
  y:[1,6,6],m:[1],saleAgo:20,decay:false},
 /* צריכה סדירה (יותר מ-3 יח׳ ב-11 חודשים) — הנתיב אינו פעיל כלל */
 {pn:'DECAY-REG',srv:80,ss:4,rop:12,ssMin:4,lt:90,price:40,cur:'USD',free:30,
  y:[18,30,30],m:[2,2,2,2,2,2,2,2,2,2,2],saleAgo:5,decay:false}];
/* כותב את הדוח הסינתטי + ITEMS לקובץ זמני ומחזיר את הנתיב */
function writeDecayReport(){
 const wb=XLSX.readFile(path.join(SD,'zmrp-demo.xlsx'));const sn=wb.SheetNames[0];
 const aoa=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:''});
 const hi=aoa.findIndex(r=>r.includes('מק"ט מוביל'));const hdr=aoa[hi];
 const tpl=aoa[hi+1].slice();
 for(const o of ITEMS)aoa.push(item(hdr,tpl,o));
 const out=XLSX.utils.book_new();XLSX.utils.book_append_sheet(out,XLSX.utils.aoa_to_sheet(aoa),sn);
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'decay-'));const f=path.join(dir,'zmrp-decay.xlsx');
 XLSX.writeFile(out,f);return f}
module.exports={ITEMS,writeDecayReport};
