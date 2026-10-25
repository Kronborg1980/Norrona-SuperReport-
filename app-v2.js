(function(){
"use strict";
const $=id=>document.getElementById(id);
const MN=["","Januar","Februar","Marts","April","Maj","Juni","Juli","August","September","Oktober","November","December"];
const EN=["January","February","March","April","May","June","July","August","September","October","November","December"];
const DEP={TFS:"Taxfree / Sjóbúðin",FNB:"F&B / Restaurant",REC:"Reception",ACC:"Housekeeping",GAL:"Kabys"};
const COUNTRY={DE:"Tyskland",FO:"Færøerne",DK:"Danmark",IS:"Island",NL:"Nederlandene",FR:"Frankrig",CH:"Schweiz",PL:"Polen",SE:"Sverige",NO:"Norge",GB:"Storbritannien",AT:"Østrig",IT:"Italien",ES:"Spanien",CZ:"Tjekkiet",BE:"Belgien",FI:"Finland",US:"USA"};
const pax=new Map(),jets=new Map();
let errors=[],tableCSV=[],reportText="";
const savedSegments=[];let segmentCSV=[];
const esc=v=>String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmt=(v,n=0)=>v==null||!Number.isFinite(v)?"—":Number(v).toLocaleString("da-DK",{minimumFractionDigits:n,maximumFractionDigits:n});
const money=v=>v==null?"—":fmt(v,0)+" DKK";
const percent=v=>v==null?"—":fmt(v,1)+" %";
const diffPercent=(a,b)=>a==null||b==null||b===0?null:100*(a/b-1);
const num=x=>typeof x==="number"&&Number.isFinite(x)?x:typeof x==="string"&&x.trim()?Number(x.replace(/\s/g,"").replace(",",".")):null;
const colLabel=c=>(c||"UK")+" · "+(COUNTRY[c]||c||"Ukendt");
const sum=a=>a.reduce((t,n)=>t+(n||0),0);
const sepYear=(y,m)=>String(y)+"-"+String(m).padStart(2,"0");
const toMonth=x=>{const key=String(x??"").trim().toLowerCase(),lookup={januar:1,februar:2,marts:3,april:4,maj:5,juni:6,juli:7,august:8,september:9,oktober:10,november:11,december:12,jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};return lookup[key]||Number(key)||0};
function parsePax(wb,file){
 const byYear=new Map();let rows=0,rejected=0,found=false;
 const required=["DEPARTURETIME - År","DEPARTURETIME - Måned","DEPARTURETIME - Dag","COUNTRY_ID"];
 for(const name of wb.SheetNames){
 const grid=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,raw:false,blankrows:false,defval:""});
 const hdr=grid.findIndex(row=>Array.isArray(row)&&required.every(k=>row.some(c=>String(c).trim()===k)));
 if(hdr<0)continue;found=true;
 const ix=Object.fromEntries(grid[hdr].map((v,i)=>[String(v).trim(),i]));
 const at=(r,k)=>ix[k]==null?"":String(r[ix[k]]??"").trim();
 for(let i=hdr+1;i<grid.length;i++){
  const row=grid[i];if(!row||!row.length)continue;
  const y=Number(at(row,"DEPARTURETIME - År")),m=toMonth(at(row,"DEPARTURETIME - Måned")),day=Number(at(row,"DEPARTURETIME - Dag"));
  const valid=Number.isInteger(y)&&y>=2010&&y<=2100&&Number.isInteger(m)&&m>=1&&m<=12&&Number.isInteger(day)&&day>=1&&day<=31;
  if(!valid){rejected++;continue}
  const check=new Date(Date.UTC(y,m-1,day));
  if(check.getUTCFullYear()!==y||check.getUTCMonth()!==m-1||check.getUTCDate()!==day){rejected++;continue}
  if(!byYear.has(y))byYear.set(y,{year:y,months:new Map(),count:0,file});
  const yr=byYear.get(y);
  if(!yr.months.has(m))yr.months.set(m,{all:new Map(),withoutNSH:new Map(),segments:new Map(),total:0,noShow:0,days:new Set()});
  const obj=yr.months.get(m),ct=at(row,"COUNTRY_ID").toUpperCase()||"UK",status=at(row,"bookingRowStatusCode").toUpperCase();
  const category=at(row,"Category").toUpperCase(),age=at(row,"AgeGroup"),guest=at(row,"GUEST_TYPE_ID").toUpperCase();
  const segmentKey=JSON.stringify([ct,category,age,guest,status]);
  obj.segments.set(segmentKey,(obj.segments.get(segmentKey)||0)+1);
  obj.all.set(ct,(obj.all.get(ct)||0)+1);obj.total++;yr.count++;obj.days.add(day);
  if(status==="NSH")obj.noShow++;else obj.withoutNSH.set(ct,(obj.withoutNSH.get(ct)||0)+1);
  rows++;
 }
 }
 if(!found||rows===0)throw Error("Passagerfilen mangler de nødvendige kolonner eller gyldige rækker");
 for(const record of byYear.values()){record.limitRisk=rows>=149500;record.rejected=rejected;record.file=file}
 return {byYear,rows,rejected};
}
function parseJet(wb,file){
 if(!wb.Sheets.input||!wb.Sheets.main)throw Error("JetReport mangler fanerne input og main");
 const input=XLSX.utils.sheet_to_json(wb.Sheets.input,{header:1,raw:true,defval:null,blankrows:false});
 const grid=XLSX.utils.sheet_to_json(wb.Sheets.main,{header:1,raw:true,defval:null,blankrows:false});
 const meta={};for(const row of input)if(row[2]!=null&&row[3]!=null&&meta[String(row[2]).trim()]==null)meta[String(row[2]).trim()]=row[3];
 const dep=String(meta.Department||"").split("|")[0].trim(),y=Number(meta.Year??meta["YEAR:"]),cutoff=Number(meta.Month??meta["Til MONTH:"]);
 if(!DEP[dep]||!Number.isInteger(y)||!Number.isInteger(cutoff)||cutoff<1||cutoff>12)throw Error("JetReport har ukendt afdeling, år eller måned");
 const titles=grid.findIndex(r=>String(r[8]??"").startsWith("January ")&&String(r[11]??"").startsWith("February "));
 const rowBy=label=>grid.filter(r=>String(r[7]??"").trim().toLowerCase()===label.toLowerCase()).at(-1);
 const income=rowBy("Income total"),paxRow=rowBy("Pax:"),stated=rowBy("Sales pr. Pax:"),gross=rowBy("Gross-Profit"),result=rowBy("Total"),gpPct=rowBy("Gross-Profit %");
 if(titles<0||!income||!paxRow||!result)throw Error("JetReport mangler Income total, Pax:, Total eller månedskolonner");
 const resultObj={year:y,dep,file,months:new Map(),cutoff,checks:[]};
 for(let m=1;m<=cutoff;m++){
  const col=8+(m-1)*3,title=String(grid[titles]?.[col]??"");
  if(!title.startsWith(EN[m-1]))throw Error("Ugyldigt månedshoved "+MN[m]+" i "+file);
  const value=(row,offset=1)=>{if(!row)return null;const n=num(row[col+offset]);return Number.isFinite(n)?n:null};
  const entry={month:m,income:value(income),budget:value(income,0),pax:value(paxRow),spendReported:value(stated),gp:value(gross),gpPctReported:value(gpPct),result:value(result)};
  if(entry.income===null||entry.pax===null)throw Error("Mangler omsætning eller pax i "+MN[m]+" "+file);
  entry.spend=entry.pax>0?entry.income/entry.pax:null;
  entry.spendGap=entry.spendReported!=null&&entry.spend!=null?entry.spend-entry.spendReported:null;
  if(entry.spendGap!=null&&Math.abs(entry.spendGap)>.10)resultObj.checks.push(MN[m]+": Sales pr. Pax afviger "+fmt(entry.spendGap,2)+" DKK");
  if(entry.gpPctReported!=null&&entry.gp!=null&&entry.income>0&&Math.abs(entry.gp*100/entry.income-entry.gpPctReported)>.2)resultObj.checks.push(MN[m]+": Dækningsgrad afviger");
  resultObj.months.set(m,entry);
 }
 return resultObj;
}
function allYears(){return [...new Set([...pax.keys(),...[...jets.values()].map(j=>j.year)])].sort((a,b)=>b-a)}
function deptCodes(){return [...new Set([...jets.values()].map(j=>j.dep))].sort()}
function officialMonth(y,m,dept){
 const src=dept==="ALL"?[...jets.values()].filter(j=>j.year===y):[jets.get(y+"|"+dept)].filter(Boolean);
 const rows=src.map(j=>j.months.get(m)).filter(Boolean);
 if(!rows.length||rows.length!==src.length)return null;
 const paxValues=rows.map(r=>r.pax).filter(v=>v>0);
 const consistent=paxValues.length>0&&paxValues.every(v=>Math.abs(v-paxValues[0])<.5);
 const revenue=sum(rows.map(r=>r.income));
 return {revenue,budget:sum(rows.map(r=>r.budget)),pax:consistent?paxValues[0]:null,
 spend:consistent?revenue/paxValues[0]:null,
 gp:dept==="ALL"?null:rows[0].gp,result:dept==="ALL"?null:rows[0].result,
 spendReported:dept==="ALL"?null:rows[0].spendReported,spendDifference:dept==="ALL"||!consistent||rows[0].spendReported==null?null:revenue/paxValues[0]-rows[0].spendReported,
 sourceCount:rows.length,sourceNames:src.map(j=>j.dep),paxMismatch:!consistent};
}
function countryMonth(y,m,withoutNoShow){
 const x=pax.get(y)?.months.get(m);if(!x)return null;
 const countries=withoutNoShow?x.withoutNSH:x.all;
 return {count:sum([...countries.values()]),countries,totalAll:x.total,noShow:x.noShow,days:x.days};
}
function periodMonths(y,dept,selection){
 if(selection!=="YTD")return [Number(selection)];
 const financial=[...jets.values()].filter(j=>j.year===y&&(dept==="ALL"||j.dep===dept));
 let end=0;
 if(financial.length)end=Math.max(...financial.map(j=>Math.max(...j.months.keys(),0)));
 else end=Math.max(...(pax.get(y)?.months.keys()||[]),0);
 return Array.from({length:end},(_,i)=>i+1);
}
function countryAggregate(year,months,withoutNoShow){
 let missing=[],countries=new Map(),nsh=0,all=0;const present=[];
 for(const m of months){
  const data=countryMonth(year,m,withoutNoShow);
  if(!data){missing.push(m);continue}
  present.push(m);nsh+=data.noShow;all+=data.totalAll;
  for(const [country,n] of data.countries)countries.set(country,(countries.get(country)||0)+n)
 }
 return {countries,total:sum([...countries.values()]),noShow:nsh,all,missing,present};
}
function financeAggregate(year,months,dept){
 let missing=[],monthly=[],income=0,budget=0,pax=0,gp=0,result=0,hasError=false,validGP=true;
 for(const m of months){
  const row=officialMonth(year,m,dept);if(!row){missing.push(m);continue}
  monthly.push({month:m,...row});income+=row.revenue;budget+=row.budget;
  if(row.pax===null){hasError=true}else pax+=row.pax;
  if(row.gp===null||row.result===null)validGP=false;else{gp+=row.gp;result+=row.result}
 }
 return {monthly,missing,income,budget,pax:hasError||!monthly.length?null:pax,
 spend:!hasError&&pax>0?income/pax:null,gp:validGP&&monthly.length?gp:null,
 result:validGP&&monthly.length?result:null,gpPct:validGP&&income>0?100*gp/income:null,
 complete:missing.length===0&&monthly.length===months.length,hasError};
}
function model(){
 const year=Number($("yearA").value)||0,yearB=Number($("yearB").value)||0,dept=$("department").value||"ALL",period=$("period").value||"YTD",withoutNoShow=$("status").value==="without";
 const months=periodMonths(year,dept,period),pa=countryAggregate(year,months,withoutNoShow),pb=yearB?countryAggregate(yearB,months,withoutNoShow):null;
 const fa=financeAggregate(year,months,dept),fb=yearB?financeAggregate(yearB,months,dept):null;
 const paxComparable=!!pb&&pa.missing.length===0&&pb.missing.length===0;
 const financeComparable=!!fb&&fa.complete&&fb.complete&&!fa.hasError&&!fb.hasError;
 return {year,yearB,dept,period,months,pa,pb,fa,fb,paxComparable,financeComparable,withoutNoShow};
}
function option(id,opts,preferred){const node=$(id);const prev=preferred??node.value;node.innerHTML=opts.map(o=>'<option value="'+esc(o[0])+'">'+esc(o[1])+'</option>').join("");if([...node.options].some(x=>x.value===String(prev)))node.value=String(prev)}
function table(columns,rows){return '<div class="table-wrap"><table><thead><tr>'+columns.map(x=>"<th>"+esc(x)+"</th>").join("")+"</tr></thead><tbody>"+rows.map(row=>"<tr>"+row.map(c=>"<td>"+c+"</td>").join("")+"</tr>").join("")+"</tbody></table></div>"}
function stat(label,value,foot){return '<div class="card kpi"><div class="sub tiny">'+esc(label)+'</div><div class="big">'+value+'</div><div class="sub tiny">'+foot+'</div></div>'}
function pill(text,cls){return '<span class="pill '+(cls||"")+'">'+esc(text)+'</span>'}
function notes(m){
 const lines=[];
 if(m.fa.monthly.length===0)lines.push("Ingen JetReport for den valgte afdeling og periode.");
 if(m.fa.missing.length&&m.fa.monthly.length)lines.push("Regnskab mangler for: "+m.fa.missing.map(x=>MN[x]).join(", ")+". Nøgletallene omfatter kun de indlæste regnskabsmåneder.");
 if(m.fa.hasError)lines.push("JetReports-pax er ikke ens på tværs af indlæste afdelinger. Omsætning pr. pax skjules.");
 if(m.pa.missing.length)lines.push("Passagerfil mangler: "+m.pa.missing.map(x=>MN[x]).join(", ")+".");
 if(m.period==="YTD"&&m.fa.monthly.length===0&&m.months.length)lines.push("YTD bygger på passagerfilens måneder, da regnskab mangler.");
 $("notices").innerHTML=lines.map(v=>'<p class="notice">'+esc(v)+'</p>').join("");
}
function drawBars(m){
 const entries=[...new Set([...m.pa.countries.keys(),...(m.pb?.countries.keys()||[])])];
 entries.sort((a,b)=>(m.pa.countries.get(b)||0)-(m.pa.countries.get(a)||0));const max=Math.max(1,...entries.slice(0,12).map(k=>Math.max(m.pa.countries.get(k)||0,m.pb?.countries.get(k)||0)));
 $("countryBars").innerHTML=entries.slice(0,12).map(k=>{
 const a=m.pa.countries.get(k)||0,b=m.pb?.countries.get(k)||0;
 return '<div class="country-row"><span class="truncate" title="'+esc(colLabel(k))+'">'+esc(colLabel(k))+'</span><div class="twobar">'+
 '<div class="track" title="'+m.year+': '+fmt(a)+'"><i style="width:'+(100*a/max)+'%;background:#32b7e5"></i></div>'+
 (m.pb?'<div class="track" title="'+m.yearB+': '+fmt(b)+'"><i style="width:'+(100*b/max)+'%;background:#e7ae50"></i></div>':"")+
 '</div><div class="right small">'+fmt(a)+(m.pb?'<div class="gold">'+fmt(b)+'</div>':"")+'</div></div>';
 }).join("")||'<p class="empty">Ingen passagertal</p>';
 $("countryLegend").innerHTML='<span><i class="swatch" style="background:#32b7e5"></i>'+m.year+'</span>'+(m.pb?'<span><i class="swatch" style="background:#e7ae50"></i>'+m.yearB+'</span>':"");
}
function drawSpend(m){
 const ys=[m.year,m.yearB].filter(Boolean),max=Math.max(1,...ys.flatMap(y=>Array.from({length:12},(_,i)=>officialMonth(y,i+1,m.dept)?.spend??null).filter(x=>x!=null)));
 $("spendGraph").innerHTML=Array.from({length:12},(_,i)=>{
 const a=officialMonth(m.year,i+1,m.dept)?.spend??null,b=m.yearB?officialMonth(m.yearB,i+1,m.dept)?.spend??null:null;
 return '<div class="month-chart"><div class="columns"><span class="bar-a" title="'+MN[i+1]+' '+m.year+': '+fmt(a,2)+'" style="height:'+(a==null?0:100*a/max)+'%"></span>'+
 (m.yearB?'<span class="bar-b" title="'+MN[i+1]+' '+m.yearB+': '+fmt(b,2)+'" style="height:'+(b==null?0:100*b/max)+'%"></span>':"")+'</div><span>'+MN[i+1].slice(0,3)+'</span></div>'
 }).join("");
 $("spendLegend").innerHTML='<span><i class="swatch" style="background:#32b7e5"></i>'+m.year+'</span>'+(m.yearB?'<span><i class="swatch" style="background:#e7ae50"></i>'+m.yearB+'</span>':"");
}
function render(){
 const m=model();notes(m);renderSources();
 const scope=Array.from({length:m.months.length},(_,i)=>MN[m.months[i]]).join(", ");
 $("scope").textContent=(DEP[m.dept]||"Alle indlæste afdelinger")+" · "+(m.period==="YTD"?"År til dato":MN[Number(m.period)])+" "+m.year+" · "+(m.fa.monthly.length?"Regnskab: "+m.fa.monthly.length+" måned(er)":"Intet regnskab");
 const f=m.fa,p=m.pa;
 $("kpis").innerHTML=stat("Omsætning",f.monthly.length?money(f.income):"—","Faktisk · JetReports")+
 stat("Omsætning pr. pax",f.spend!=null?fmt(f.spend,2)+" DKK":"—","Omsætning ÷ JetReports-pax")+
 stat("JetReports-pax",f.pax!=null?fmt(f.pax):"—","Samme skibspassagerer tælles kun én gang")+
 stat("Passagerregistreringer",p.present.length?fmt(p.total):"—",m.withoutNoShow?"Uden no-shows":"Alle statusser, inklusive no-shows")+
 stat("Dækningsgrad",f.gpPct!=null?percent(f.gpPct):"—",f.gpPct!=null?"Dækningsbidrag: "+money(f.gp):"Kun enkelt afdeling")+
 stat("Afdelingsresultat",f.result!=null?money(f.result):"—",f.result!=null?"Faktisk · JetReports":"Vises kun for enkelt afdeling");
 $("compareHeading").textContent=m.yearB?"Sammenligning "+m.year+" mod "+m.yearB:"Sammenligning mellem år";
 if(m.financeComparable){
 $("compareNotice").textContent="Økonomiske regnskabstal er indlæst for samme måneder i begge år.";
 $("yoy").innerHTML=stat("Omsætning · ændring",percent(diffPercent(f.income,m.fb.income)),"Samme måneder · "+m.year+" vs. "+m.yearB)+
 stat("DKK/pax · ændring",percent(diffPercent(f.spend,m.fb.spend)),"JetReports-pax · samme måneder")+
 stat("JetReports-pax · ændring",percent(diffPercent(f.pax,m.fb.pax)),"Samme skibspassagerer");
 }else{
 const hasA=f.monthly.length>0,hasB=!!m.fb&&m.fb.monthly.length>0;
 $("compareNotice").textContent=!m.yearB?"Vælg et sammenligningsår.":hasA&&!hasB?
 "Regnskab for "+m.yearB+" er ikke indlæst for "+(DEP[m.dept]||"de valgte afdelinger")+". Passagermix kan stadig sammenlignes nedenfor.":hasA&&hasB?
 "Økonomisk sammenligning afventer JetReports for samme måneder i begge år.":"Økonomiske sammenligninger kræver JetReports fra begge år.";
 $("yoy").innerHTML="";
 }
 if(m.yearB&&m.paxComparable){
 const change=m.pa.total-m.pb.total,rel=diffPercent(m.pa.total,m.pb.total);
 $("paxCompare").innerHTML='<h3 style="margin:0 0 12px">Passagerregistreringer · '+m.year+' mod '+m.yearB+'</h3>'+
 '<div class="grid g3">'+
 stat("Passagermix "+m.year,fmt(m.pa.total),"Indlæst passagerfil")+
 stat("Passagermix "+m.yearB,fmt(m.pb.total),"Indlæst passagerfil")+
 stat("Forskel i antal",(change>0?"+":"")+fmt(change),percent(rel)+" fra "+m.yearB+" til "+m.year)+
 '</div><p class="sub small" style="margin:11px 0 0">Landefordelingen og ændringer i procentpoint finder du under Passagermix. Passagerregistreringer er ikke det samme som officielle JetReports-pax.</p>';
 }else{
 $("paxCompare").innerHTML=m.yearB?
 '<p class="sub small">Passagerregistreringer for '+m.yearB+' kan endnu ikke sammenlignes for de valgte måneder. Kontrollér passagerfilen.</p>':"";
 }
 const bFinance=m.yearB&&[...jets.values()].some(j=>j.year===m.yearB&&(m.dept==="ALL"||j.dep===m.dept));
 $("spendGraphInfo").textContent="Grafen viser indlæste regnskabsmåneder for hele året, uanset valgt enkeltmåned. DKK/pax = omsætning ÷ JetReports-pax. "+
 (m.yearB&&!bFinance?"Guld mangler, fordi JetReports for "+m.yearB+" ikke er indlæst. Pax "+m.yearB+" kan ses i sammenligningen ovenfor.":"");
 drawSpend(m);
 $("spendMonthTable").innerHTML=table(["Måned","Omsætning A","Jet-pax A","DKK/pax A","Omsætning B","Jet-pax B","DKK/pax B"],m.months.map(mm=>{
 const a=officialMonth(m.year,mm,m.dept),b=m.yearB?officialMonth(m.yearB,mm,m.dept):null;
 return [esc(MN[mm]),a?money(a.revenue):"—",a?.pax!=null?fmt(a.pax):"—",a?.spend!=null?fmt(a.spend,2):"—",
 b?money(b.revenue):"—",b?.pax!=null?fmt(b.pax):"—",b?.spend!=null?fmt(b.spend,2):"—"]}));
 drawBars(m);
 const countries=[...new Set([...p.countries.keys(),...(m.pb?.countries.keys()||[])])].sort((a,b)=>(p.countries.get(b)||0)-(p.countries.get(a)||0));
 const rows=countries.map(k=>{
  const a=p.countries.get(k)||0,b=m.pb?.countries.get(k)||0,aa=p.total?100*a/p.total:null,bb=m.pb?.total?100*b/m.pb.total:null;
  return [esc(colLabel(k)),fmt(a),percent(aa),m.pb&&m.pb.missing.length===0?fmt(b):"—",
   m.pb&&m.pb.missing.length===0?percent(bb):"—",m.paxComparable?(a-b>=0?"+":"")+fmt(a-b):"—",
   m.paxComparable&&aa!=null&&bb!=null?(aa-bb>=0?"+":"")+fmt(aa-bb,1)+" pp":"—"];
 });
 tableCSV=[["Land","Registreringer "+m.year,"Andel "+m.year+" %","Registreringer "+(m.yearB||"B"),"Andel B %","Forskel","Mix pp"],
 ...countries.map(k=>{const a=p.countries.get(k)||0,b=m.pb?.countries.get(k)||0,aa=p.total?100*a/p.total:null,bb=m.pb?.total?100*b/m.pb.total:null;return [k,a,aa,m.pb&&m.pb.missing.length===0?b:"",m.pb?.total?bb:"",m.paxComparable?a-b:"",m.paxComparable&&aa!=null&&bb!=null?aa-bb:""]})];
 $("countriesTable").innerHTML=table(["Nationalitet",m.year+" pax",m.year+" andel",(m.yearB||"B")+" pax",(m.yearB||"B")+" andel","Forskel pax","Ændring i andel"],rows);
 $("paxTotal").textContent=p.present.length?fmt(p.total)+" registreringer / "+p.countries.size+" nationaliteter":"Ingen passagerdata";
 const dRows=[];
 for(const dep of deptCodes()){
 const o=financeAggregate(m.year,m.months,dep);
 dRows.push([esc(DEP[dep]),o.monthly.length?money(o.income):"—",o.pax!=null?fmt(o.pax):"—",o.spend!=null?fmt(o.spend,2):"—",
 o.gpPct!=null?percent(o.gpPct):"—",o.result!=null?money(o.result):"—",o.monthly.length+"/"+m.months.length]);
 }
 $("departmentTable").innerHTML=table(["Afdeling","Omsætning","JetReports-pax","DKK pr. pax","DB%","Resultat","Måneder"],dRows);
 const allJetRows=m.months.map(mm=>{const a=officialMonth(m.year,mm,m.dept),pa=countryMonth(m.year,mm,m.withoutNoShow);
 return [esc(MN[mm]),a?money(a.revenue):"—",a?.pax!=null?fmt(a.pax):"—",a?.spend!=null?fmt(a.spend,2):"—",
 a?.spendReported!=null?fmt(a.spendReported,2):"—",a?.spendDifference!=null?fmt(a.spendDifference,3):"—",
 pa?fmt(pa.count):"—",a?.pax!=null&&pa?((pa.count-a.pax>=0?"+":"")+fmt(pa.count-a.pax)):"—"]});
 $("dataTable").innerHTML=table(["Måned","Omsætning","Jet pax","Beregnet DKK/pax","Jet-felt Sales pr. Pax","Afvigelse DKK","Paxfil registreringer","Afvigelse pax"],allJetRows);
 const sep=pax.get(m.year)?.months.get(9);
 $("dataNotes").innerHTML=
 '<p><strong>Omsætning pr. pax</strong> = summen af omsætning / summen af JetReports-pax for præcis de samme måneder. Aldrig gennemsnit af månedlige DKK/pax. Den beregnede månedsværdi kontrolleres mod JetReports-feltet Sales pr. Pax.</p>'+
 '<p><strong>Nationaliteter</strong> er passagerregistreringer fra paxfilen, ikke køb. Ingen omsætning fordeles på nationaliteter.</p>'+
 '<p><strong>Afdelingsvalg “Alle”</strong> summerer kun indlæste afdelingers omsætning; JetReports-pax lægges aldrig sammen. DB og resultat for hele skibet vises ikke som konsoliderede tal.</p>'+
 (sep?'<p><strong>September '+m.year+':</strong> '+fmt(sep.total)+' registreringer, '+fmt(sep.days.size)+' forskellige afgangsdage. Sidste afgangsdato er ikke i sig selv tegn på en ufuldstændig måned.</p>':"")+
 [...jets.values()].filter(j=>j.year===m.year&&(m.dept==="ALL"||j.dep===m.dept)).flatMap(j=>j.checks.map(c=>'<p class="alert">'+esc(j.file+": "+c)+'</p>')).join("");
 makeReport(m,countries);
}
function makeReport(m,countries){
 const f=m.fa,p=m.pa,top=countries.slice(0,8).map(k=>{const n=p.countries.get(k)||0;return colLabel(k)+": "+fmt(n)+" ("+percent(p.total?100*n/p.total:null)+")"}).join("; ");
 const lines=[
 "NORRÖNA · KOMMERCIEL RAPPORT",
 (DEP[m.dept]||"Alle indlæste afdelinger")+" · "+(m.period==="YTD"?"År til dato":MN[Number(m.period)])+" "+m.year,
 "",
 "VERIFICEREDE REGNSKABSNØGLETAL",
 "Omsætning: "+money(f.monthly.length?f.income:null),
 "JetReports-passagerer: "+fmt(f.pax),
 "Omsætning pr. pax: "+(f.spend==null?"—":fmt(f.spend,2)+" DKK"),
 "Dækningsbidrag: "+money(f.gp),
 "Dækningsgrad: "+percent(f.gpPct),
 "Afdelingsresultat: "+money(f.result),
 "Regnskabsmåneder: "+f.monthly.map(v=>MN[v.month]).join(", "),
 "",
 "PASSAGERMIX",
 "Registreringer: "+fmt(p.present.length?p.total:null),
 "Bookingstatus: "+(m.withoutNoShow?"uden no-shows":"alle registreringer"),
 "Top nationaliteter: "+(top||"ingen"),
 "",
 "KONTROL OG FORBEHOLD",
 "Omsætning/pax er altid beregnet ud fra JetReports-omsætning og JetReports-pax.",
 "Passagerfilen indeholder registreringer og kan afvige fra JetReports-pax.",
 "Nationaliteters omsætning kan ikke beregnes af disse kilder. Ingen simulering anvendt."
 ];
 reportText=lines.join("\n");$("reportText").textContent=reportText;
}
function refresh(preserve){
 const yrs=allYears(),saved={yearA:$("yearA").value,yearB:$("yearB").value,department:$("department").value,period:$("period").value};
 option("yearA",yrs.map(y=>[y,y]),preserve?saved.yearA:yrs[0]);
 option("yearB",[["","Intet sammenligningsår"],...yrs.map(y=>[y,y])],preserve?saved.yearB:yrs[1]||"");
 const depts=deptCodes();
 option("department",[["ALL","Alle indlæste afdelinger"],...depts.map(d=>[d,DEP[d]])],preserve?saved.department:depts.length===1?depts[0]:"ALL");
 option("period",[["YTD","År til dato (regnskab)"],...MN.slice(1).map((v,i)=>[i+1,v])],preserve?saved.period:"YTD");
 render();
}
function renderSources(){
 const all=[...pax.values()].map(v=>({key:"P|"+v.year,text:"PAX "+v.year,file:v.file})).concat([...jets.values()].map(v=>({key:"J|"+v.year+"|"+v.dep,text:v.dep+" "+v.year,file:v.file})));
 $("sources").innerHTML=all.map(v=>'<span class="file-chip"><strong>'+esc(v.text)+'</strong><span class="truncate">'+esc(v.file)+'</span><button type="button" data-del="'+esc(v.key)+'" title="Fjern kilde">×</button></span>').join("");
 $("loadStatus").textContent=pax.size+" pax-år · "+jets.size+" JetReports";
}
async function importFiles(files){
 const valid=[...files].filter(f=>/\.(xlsx|xls|xlsm)$/i.test(f.name));if(!valid.length)return;
 $("loadStatus").textContent="Indlæser "+valid.length+" filer …";
 const faults=[],loaded=[];
 for(const f of valid){
  try{
   if(typeof XLSX==="undefined")throw Error("Excel-biblioteket kunne ikke indlæses");
   const wb=XLSX.read(await f.arrayBuffer(),{type:"array",cellDates:false});
   if(wb.Sheets.input&&wb.Sheets.main){const record=parseJet(wb,f.name);jets.set(record.year+"|"+record.dep,record);loaded.push(record.dep+" "+record.year)}
   else{const record=parsePax(wb,f.name);for(const [year,yr] of record.byYear){pax.set(year,yr);loaded.push("PAX "+year)}}
  }catch(e){faults.push(f.name+": "+e.message)}
 }
 errors=faults;
 refresh(false);
 $("feedback").textContent=loaded.length?"Indlæst: "+loaded.join(", "):"";
 if(faults.length)$("feedback").textContent+=" · "+faults.join(" · ");
}
function csv(rows,filename){
 const content="\ufeff"+rows.map(r=>r.map(v=>'"'+String(v??"").replace(/"/g,'""')+'"').join(";")).join("\r\n");
 const url=URL.createObjectURL(new Blob([content],{type:"text/csv;charset=utf-8"}));
 const a=document.createElement("a");a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
async function powerPoint(){
 if(!window.PptxGenJS){
  await new Promise((resolve,reject)=>{const s=document.createElement("script");s.src="https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js";s.onload=resolve;s.onerror=()=>reject(Error("PowerPoint-modulet kunne ikke hentes"));document.head.appendChild(s)});
 }
 const PPT=window.PptxGenJS,ppt=new PPT();ppt.layout="LAYOUT_WIDE";ppt.author="Smyril Line";ppt.title="Norröna · nøgletal";ppt.subject="Passagermix og regnskab";const m=model(),f=m.fa;
 const add=(title,lines)=>{const slide=ppt.addSlide();slide.background={color:"112838"};slide.addText(title,{x:.7,y:.45,w:11.6,h:.75,fontSize:28,bold:true,color:"FFFFFF",fontFace:"Aptos"});slide.addShape(ppt.ShapeType.line,{x:.7,y:1.35,w:11.9,h:0,line:{color:"DAA74C",width:2}});let y=1.68;for(const l of lines){slide.addText(l,{x:.85,y,w:11.5,h:.7,fontSize:20,color:"D6E9F4",fontFace:"Aptos"});y+=.78}};
 add("MS Norröna · nøgletal",[m.year+" · "+(DEP[m.dept]||"Alle indlæste afdelinger"),"Omsætning: "+money(f.monthly.length?f.income:null),"JetReports-pax: "+fmt(f.pax),"Omsætning/pax: "+(f.spend==null?"—":fmt(f.spend,2)+" DKK"),"Dækningsgrad: "+percent(f.gpPct)]);
 const top=[...m.pa.countries].sort((a,b)=>b[1]-a[1]).slice(0,7);
 add("Passagermix · faktiske registreringer",top.map(([k,n])=>colLabel(k)+": "+fmt(n)+" · "+percent(m.pa.total?100*n/m.pa.total:null)));
 add("Datagrundlag og kontrol",["Regnskab: JetReports · ikke passagerfil","Omsætning/pax: omsætning / officielt pax","Passagermix: nationalitetsregistreringer","Manglende regnskab vises som —","Ingen modellerede køb pr. nationalitet"]);
 await ppt.writeFile({fileName:"Norrona_noegletal_"+m.year+".pptx"});
}
function activate(tab){
 for(const b of document.querySelectorAll("[data-tab]"))b.classList.toggle("on",b.dataset.tab===tab);
 for(const section of document.querySelectorAll(".tab"))section.classList.toggle("active",section.id==="tab-"+tab);
}
function wire(){
 $("choose").onclick=()=>$("inputFiles").click();
 $("drop").onclick=()=>$("inputFiles").click();
 $("drop").onkeydown=e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();$("inputFiles").click()}};
 $("inputFiles").onchange=async e=>{await importFiles(e.target.files);e.target.value=""};
 document.addEventListener("dragover",e=>{if(e.dataTransfer?.types?.includes("Files")){e.preventDefault();$("drop").classList.add("hover")}});
 document.addEventListener("dragleave",e=>{if(e.relatedTarget==null)$("drop").classList.remove("hover")});
 document.addEventListener("drop",async e=>{if(e.dataTransfer?.files?.length){e.preventDefault();$("drop").classList.remove("hover");await importFiles(e.dataTransfer.files)}});
 $("sources").addEventListener("click",e=>{const key=e.target.closest("[data-del]")?.dataset.del;if(!key)return;const [type,y,d]=key.split("|");if(type==="P")pax.delete(Number(y));else jets.delete(y+"|"+d);refresh(true)});
 for(const id of ["yearA","yearB","department","period","status"])$(id).onchange=()=>render();
 document.querySelector("nav").onclick=e=>{const x=e.target.closest("[data-tab]");if(x)activate(x.dataset.tab)};
 $("exportCSV").onclick=()=>csv(tableCSV,"norrona_passagermix.csv");
 $("print").onclick=()=>{activate("report");window.print()};
 $("exportPPT").onclick=()=>{powerPoint().catch(e=>$("feedback").textContent="PowerPoint: "+e.message)};
 $("reset").onclick=()=>{pax.clear();jets.clear();refresh(false);$("feedback").textContent=""};
}
wire();refresh(false);
window.NorronaDiagnostics={parsePax,parseJet,officialMonth,financeAggregate,countryAggregate,model,getPax:()=>pax,getJets:()=>jets};
})();