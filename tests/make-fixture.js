const XLSX=require('xlsx');
const MON=Array.from({length:11},(_,i)=>'צר.חודש-'+(i+1));
const hdr=['מק"ט מוביל','תיאור חומר','תיאור חומר2','שם ספק','סטטוס חומר','תיאור',
 'סוג MRP','ABC','רמת שרות','מלאי בטחון','נק.הז.מחדש','אספ.מתוכנ.','זמ.עב.קבלת','מל.בט.מינ.',
 'מחיר FOB','מטבע FOB','סה"כ מלאי','מלאי פנוי','מלאי מרלוג','הז. רכש','אספקות פת.','כמות בהז.פ',
 'תיא.קבוצ.חומרים','קב.חו.חיצו','טקסט ארוך','תב.אח.הש.','ת.היר.1 מח','ת.היר.2 מח',
 'היררכייה1','היררכייה1','היררכייה2','היררכייה2','היררכייה3','היררכייה3','אביזרים',
 'צר.השנה','צר.שנה-1','צר.שנה-2','צר.החודש',...MON,"תאר' מכירה","תאר' כניסה"];
const BR=['ZT','LINK&CO'],SUP=['אלפא רכיבים','בטא ייבוא','גמא חלפים','דלתא מוטורס'],
 GRP=['מנוע','חשמל','מרכב','בלמים','מתלים','פנים'],AB=['A','B','C'],
 MODEL=['ZT T5','ZT X7','LINK 01','LINK 03','ZT E-Van'];
const R=(a,b)=>a+Math.random()*(b-a), I=(a,b)=>Math.round(R(a,b));
const dstr=d=>{const x=new Date(d);return `${String(x.getDate()).padStart(2,'0')}.${String(x.getMonth()+1).padStart(2,'0')}.${x.getFullYear()}`};
const now=Date.now(),DAY=864e5;
const rows=[];
for(let i=0;i<900;i++){
 const kind=Math.random();
 let months,y0,y1,y2,free,cust,po,stock,saleAgo,entAgo;
 if(kind<.25){ // בוער: לקוח ממתין ללא כיסוי
  months=MON.map(()=>I(0,14)); y0=I(40,120);y1=I(40,120);y2=I(30,110);
  free=I(0,3);cust=I(5,40);po=0;stock=free+I(0,2);saleAgo=I(3,40);entAgo=I(60,300);
 }else if(kind<.45){ // מניעת חוסר
  months=MON.map(()=>I(2,20)); y0=I(60,200);y1=I(60,200);y2=I(50,180);
  free=I(2,10);cust=I(0,4);po=I(0,6);stock=free+I(0,5);saleAgo=I(1,25);entAgo=I(40,200);
 }else if(kind<.62){ // עודף
  months=MON.map(()=>I(0,4)); y0=I(5,25);y1=I(5,30);y2=I(5,30);
  free=I(200,900);cust=0;po=0;stock=free+I(0,30);saleAgo=I(30,200);entAgo=I(100,500);
 }else if(kind<.78){ // מת
  months=MON.map(()=>0); y0=0;y1=0;y2=0;
  free=I(10,300);cust=0;po=0;stock=free;saleAgo=I(1300,2600);entAgo=I(1300,2600);
 }else if(kind<.9){ // איטי
  months=MON.map(()=>Math.random()<.25?I(1,3):0); y0=I(1,6);y1=I(1,8);y2=I(1,8);
  free=I(80,400);cust=0;po=I(0,2);stock=free;saleAgo=I(200,900);entAgo=I(300,1000);
 }else{ // פעיל רגיל
  months=MON.map(()=>I(3,25)); y0=I(80,300);y1=I(80,300);y2=I(70,280);
  free=I(30,200);cust=I(0,10);po=I(0,20);stock=free+I(0,40);saleAgo=I(1,20);entAgo=I(10,90);
 }
 const price=Math.random()<.19?'':+R(20,3200).toFixed(2);
 const sell=price===''?'':'USD';
 const br=BR[I(0,BR.length-1)],md=MODEL[I(0,MODEL.length-1)],gp=GRP[I(0,GRP.length-1)];
 const pn=`${I(10,99)}${String.fromCharCode(65+I(0,25))}${I(100000,999999)}`;
 rows.push([pn,`חלק ${gp} ${md} דגם ${i}`,`Part ${gp} ${i}`,SUP[I(0,SUP.length-1)],
  ['01','04','14','01','01'][I(0,4)],'פעיל',['VB','VB','PD','V1'][I(0,3)],AB[I(0,2)],+R(90,99).toFixed(1),I(0,20),I(0,40),
  I(14,180),I(1,14),I(0,10),price,sell,stock,free,I(0,50),po,I(0,5),cust,
  gp,br,'מתכנן א׳','12 חודשים',md,gp,
  '100',br,'200',md,'300',gp,Math.random()<.15?'X':'',
  y0,y1,y2,I(0,10),...months,
  dstr(now-saleAgo*DAY),dstr(now-entAgo*DAY)]);
}
const ws=XLSX.utils.aoa_to_sheet([['דוח ZMRP — תכנון מלאי'],[],hdr,...rows]);
const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'ZMRP');
XLSX.writeFile(wb,__dirname+'/zmrp-demo.xlsx');
console.log('wrote',rows.length,'rows,',hdr.length,'cols');
