import {exportExactFloorPlan} from "./excel-export.js?v=20261010d";
import {parseCsv,reportType,buildEvents} from "./csv-import.js?v=20261010e";
let pdfjsLib=null;

const $=id=>document.getElementById(id);
const DAYS=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"];
const COLOR={
  "Flex":{Definite:"#D9EAD3",Tentative:"#6AA84F"},
  "Pleaser":{Definite:"#FFF2CC",Tentative:"#F1C232"},
  "Buffet":{Definite:"#C9DAF8",Tentative:"#3C78D8"},
  "Field Trip":{Definite:"#D9D2E9",Tentative:"#674EA7"},
  "Youth":{Definite:"#F4CCCC",Tentative:"#E06666"}
};
const STORAGE="fpmatic-v1";
const state={events:[],week:"",day:0,selection:"",master:null,files:[],showBayNumbers:true,layout:null};
let noticeTimer;
function notify(msg){$("toast").textContent=msg;$("toast").classList.add("visible");clearTimeout(noticeTimer);noticeTimer=setTimeout(()=>$("toast").classList.remove("visible"),3400)}
function pad(n){return String(n).padStart(2,"0")}
function dateIso(y,m,d){return [y,pad(m),pad(d)].join("-")}
function usToIso(s){const m=String(s||"").match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);return m?dateIso(+m[3],+m[1],+m[2]):""}
function localDate(s){const a=(s||"").split("-").map(Number);return new Date(a[0],a[1]-1,a[2],12)}
function isoDate(d){return dateIso(d.getFullYear(),d.getMonth()+1,d.getDate())}
function addDays(s,n){const d=localDate(s);d.setDate(d.getDate()+n);return isoDate(d)}
function weekOf(iso){const d=localDate(iso);d.setDate(d.getDate()-(d.getDay()+6)%7);return isoDate(d)}
function niceDate(s,opts){return localDate(s).toLocaleDateString("en-US",opts||{month:"short",day:"numeric"})}
function clockMin(v){if(v===undefined||v===null||v==="")return null;if(/^\d{1,2}:\d\d$/.test(v)){const [h,m]=v.split(":").map(Number);return h*60+m}const a=String(v).match(/(\d{1,2})(?::(\d\d))?\s*(AM|PM)/i);if(!a)return null;let h=+a[1]%12;if(a[3].toUpperCase()==="PM")h+=12;return h*60+Number(a[2]||0)}
function time24(s){const m=clockMin(s);return m===null?"":pad(Math.floor(m/60))+":"+pad(m%60)}
function fmtMin(v){if(v===null||v===undefined)return"";const h=Math.floor(v/60)%24;return String((h%12)||12)+(v%60?pad(v%60):"")}
function fmtRange(a,b){return fmtMin(clockMin(a))+"-"+fmtMin(clockMin(b))}
function isPM(e){const a=clockMin(e.start),b=clockMin(e.end);return a!==null&&b!==null&&(a>=960||b>=1140)}
function period(e){return isPM(e)?"PM":"AM"}
function escapeHtml(s){return String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}
function shortName(s){return String(s||"").replace(/\s*[-–]\s*\d\d?[./-]\d\d?[./-]\d{2,4}\s*$/,"").trim()}
function displayLines(e){
  const out=[e.name||"Unnamed event",(e.space?e.space+" ":"")+fmtRange(e.start,e.end)];
  if(e.golfStart&&e.golfEnd&&(clockMin(e.golfStart)!==clockMin(e.start)||clockMin(e.golfEnd)!==clockMin(e.end)))out.push("Golf "+fmtRange(e.golfStart,e.golfEnd));
  out.push((e.guests===null||e.guests===undefined||e.guests==="")?"NA ppl":String(e.guests)+"ppl");
  if(e.package)out.push(e.package);
  return out;
}
function extractLines(items){
  const buckets=[];
  for(const a of items){const str=(a.str||"").trim();if(!str)continue;const x=a.transform[4],y=a.transform[5];let row=buckets.find(b=>Math.abs(b.y-y)<2.15);if(!row){row={y,items:[]};buckets.push(row)}row.items.push({x,str})}
  buckets.sort((a,b)=>b.y-a.y);
  return buckets.map(b=>b.items.sort((a,c)=>a.x-c.x).map(x=>x.str).join(" ").replace(/\s+/g," ").trim());
}
function findField(lines,label){
  const rx=new RegExp("(?:^|\\s)"+label+"\\s*:\\s*(.*)","i");
  const l=lines.find(x=>rx.test(x));
  if(!l)return"";
  return (l.match(rx)||[])[1]||"";
}
function scanTimes(s){
  const m=String(s||"").match(/((?:1[0-2]|0?[1-9]):[0-5]\d\s*[AP]M)\s*[-–—]\s*((?:1[0-2]|0?[1-9]):[0-5]\d\s*[AP]M)/i);
  return m?[time24(m[1]),time24(m[2])]:null;
}
function findSectionRange(pages,title){
  for(const page of pages){
    const lines=page.lines;
    for(let i=0;i<lines.length;i++){
      if(!new RegExp("^"+title+"(?:\\s|$)","i").test(lines[i])||/TOTAL/i.test(lines[i]))continue;
      for(let j=i+1;j<Math.min(i+65,lines.length);j++){
        const l=lines[j];
        if(/^(BILLING SUMMARY|PAYMENT SUMMARY|EVENT NOTES|Asset|Beverages|Food|Packages|Game Enhancement|Labor|Incentives|Limited Inventory|Event Space Rental)/i.test(l))break;
        const range=scanTimes(l);if(range)return range;
      }
    }
  }
  return null;
}
function baySegments(s){
  const cleaned=String(s||"").replace(/(?:BAY LEVEL ASSIGNMENT\(S\)|GOLF BAYS)\s*:/gi,"");
  const hits=[...cleaned.matchAll(/(\d{3})(?:\s*[-–—]\s*(\d{3}))?/g)];
  const ranges=[];
  for(const m of hits){const lo=+m[1],hi=+(m[2]||m[1]);if(lo>hi||Math.floor(lo/100)!==Math.floor(hi/100)||!([1,2,3].includes(Math.floor(lo/100)))||lo%100<1||hi%100>34)continue;ranges.push({start:lo,end:hi,floor:Math.floor(lo/100)})}
  return ranges.filter((r,i)=>ranges.findIndex(x=>x.start===r.start&&x.end===r.end)===i);
}
function spaceFrom(text,header){
  const search=(header||"")+"\n"+text.slice(0,4500);
  const north=/Signature (?:Room )?North|Sig North/i.test(search);
  const south=/Signature (?:Room )?South|Sig South/i.test(search);
  const lower=/Lower Lounge/i.test(search);
  const upper=/Upper Lounge/i.test(search);
  const terrace=/(?:Terrace)\s*(?:-|Hourly|Rental|$)/im.test(search)&&/Terrace/i.test(header+text.slice(0,2200));
  const p=[];if(north&&south)p.push("Sig N+S");else if(north)p.push("Sig N");else if(south)p.push("Sig S");
  if(lower)p.push("LL");if(upper)p.push("UL");if(terrace)p.push("Terrace");
  return p.join(", ");
}
function categoryAndPackage(text,booking,catering,name){
  const s=text.slice(0,21000);
  if(/(?:Play\s*\+\s*Eat\s*-\s*Field Trip|Field Trip\s*-\s*(?:Two|Turkey|Pizza)|field trip)/i.test(s)&&!/(?:Field Trip.*only in notes)/i.test(s))return {category:"Field Trip",package:"Field Trip"};
  if(/Youth Platter|Youth - Fountain|(?:Birthday|Bday)/i.test(s+" "+name)&&(/Youth Event/i.test(booking)||/Youth Platter/i.test(s)))return {category:"Youth",package:"Youth"};
  if(/Crowd Pleaser/i.test(s))return {category:"Pleaser",package:"Pleaser"};
  if(/Fan Fav(?:orite)?|Teeline Taster|Tee Line Taster/i.test(s))return {category:"Pleaser",package:/Taster/i.test(s)?"Taster":"Fan Fave"};
  const foods=[
    [/Fajita Fiesta/i,"Fajita"],[/Ultimate Bites/i,"Ultimate"],[/Smokehouse BBQ/i,"Smokehouse"],
    [/(?:The )?Premium/i,"Premium"],[/Most Valuable Buffet|The MVB\b/i,"MVB"],
    [/Rise\s*&\s*Shine/i,"Rise & Shine"],[/Continental/i,"Continental"],
    [/Fairway Feast/i,"Fairway Feast"],[/Breakfast Buffet/i,"Breakfast"],
    [/(?:Italian|Taco|BBQ|Buffet)\s*(?:Bar|Buffet|Menu|Package)/i,null]
  ];
  for(const [re,val] of foods)if(re.test(s))return {category:"Buffet",package:val||(s.match(re)||[])[0]};
  if(/Youth Event/i.test(booking))return {category:"Youth",package:"Youth"};
  if(/Corporate|Social|Internal/i.test(booking)&&/Flex|Food & Beverage Minimum|F&B Minimum|Individual Tabs|Order off Menu|in the bays/i.test(s+" "+catering))return {category:"Flex",package:"Flex"};
  return {category:"Flex",package:"Flex"};
}
function eventFromPages(pages,beo,file){
  const head=pages.find(p=>p.lines.some(l=>/Manager Catering BEO/i.test(l)))||pages[0];
  const first=head.lines;
  const pre=pages.map(p=>p.lines.join("\n")).join("\n");
  const date=usToIso(findField(first,"Event Date"));
  const times=scanTimes(findField(first,"Event Time"));
  const statusRaw=findField(first,"Event Status");
  const countMatch=findField(first,"Event Guest Count").match(/^\d+/);
  const status=/^Definite\b/i.test(statusRaw)?"Definite":"Tentative";
  let name="";
  const titleLine=first.findIndex(l=>/Manager Catering BEO/i.test(l));
  if(titleLine!==-1){
    const next=(first[titleLine+1]||"").split(/\s+Event Date\s*:/i)[0].trim();
    name=next.split(/\s+-\s+(?=[^-]+-\s*\d{1,2}[./]\d{1,2}[./]\d{2,4})/)[0].trim();
    if(name.includes(" - "))name=name.split(" - ")[0];
  }
  if(!name||/^(Event|Account|Contact)/i.test(name)){
    const ix=first.findIndex(l=>/BEO\s*#\s*\d+/i.test(l));
    const next=first[ix+1]||"";
    name=next.split(/Date Printed:/i)[0].trim();
  }
  if(!name)name="Event "+beo;
  const golfField=findField(first,"GOLF BAYS");
  const noteLine=pages.map(p=>findField(p.lines,"BAY LEVEL ASSIGNMENT\\(S\\)")).find(v=>baySegments(v).length)||"";
  const bayRanges=baySegments(golfField).length?baySegments(golfField):baySegments(noteLine);
  const bays=bayRanges.map(r=>r.start===r.end?String(r.start):r.start+"-"+r.end).join(", ");
  const booking=findField(first,"BOOKING TYPE");
  const catering=findField(first,"CATERING LOCATION");
  const x=categoryAndPackage(pre,booking,catering,name);
  const golf=findSectionRange(pages,"Gameplay");
  const space=spaceFrom(pre,catering);
  return {id:"beo-"+beo,beo,name:shortName(name),date,start:times?times[0]:"",end:times?times[1]:"",golfStart:golf&&times&&(golf[0]!==times[0]||golf[1]!==times[1])?golf[0]:"",golfEnd:golf&&times&&(golf[0]!==times[0]||golf[1]!==times[1])?golf[1]:"",guests:countMatch?Number(countMatch[0]):null,bays,category:x.category,package:x.package,status,space,source:file};
}
async function parsePdf(file){
  if(!pdfjsLib){
    pdfjsLib=await import("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs");
    pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs";
  }
  const bytes=new Uint8Array(await file.arrayBuffer());
  const pdf=await pdfjsLib.getDocument({data:bytes}).promise;
  const byBEO=new Map();let previous="";
  for(let n=1;n<=pdf.numPages;n++){
    const pg=await pdf.getPage(n);const lines=extractLines((await pg.getTextContent()).items);
    const first=lines.slice(0,13).join(" ");
    let id=(first.match(/\bBEO\s*#\s*(\d{6,9})\b/i)||[])[1]||"";
    if(!id)id=previous;
    if(id){if(!byBEO.has(id))byBEO.set(id,[]);byBEO.get(id).push({lines,page:n});previous=id}
    if(n%7===0){notify("Reading "+file.name+" · page "+n+" / "+pdf.numPages);await new Promise(r=>setTimeout(r,0))}
  }
  const out=[];
  for(const [id,pages] of byBEO){
    if(!pages.some(p=>p.lines.some(l=>/Manager Catering BEO/i.test(l))))continue;
    out.push(eventFromPages(pages,id,file.name));
  }
  return out;
}
function persist(){try{localStorage.setItem(STORAGE,JSON.stringify({events:state.events,week:state.week,day:state.day,files:state.files}))}catch(e){/* local storage may be disabled */}}
function restore(){try{const d=JSON.parse(localStorage.getItem(STORAGE)||"null");if(d&&Array.isArray(d.events)){state.events=d.events;state.week=d.week;state.day=d.day||0;state.files=d.files||[]}}catch(e){}}
function overlaps(a,b){return a.start<=b.end&&b.start<=a.end}
function timeCollision(a,b){
  const aa=clockMin(a.golfStart||a.start),ab=clockMin(a.golfEnd||a.end),ba=clockMin(b.golfStart||b.start),bb=clockMin(b.golfEnd||b.end);
  if([aa,ab,ba,bb].some(x=>x===null))return false;
  return aa<(bb<=ba?bb+1440:bb)&&ba<(ab<=aa?ab+1440:ab);
}
function validation(e){
  const issues=[];
  if(e.importFlags)for(const flag of e.importFlags)if(!/No bay assignment|Missing event date or time/.test(flag))issues.push(flag);
  if(!e.date)issues.push("Missing event date");
  if(!e.start||!e.end)issues.push("Missing event time");
  if(!baySegments(e.bays).length)issues.push("Missing or invalid bay assignment");
  return issues;
}
function weeksEvents(){return state.events.filter(e=>e.date&&weekOf(e.date)===state.week)}
function dayEvents(){return weeksEvents().filter(e=>e.date===addDays(state.week,state.day))}
function allIssues(){
  const out=[];
  const current=weeksEvents();
  for(const e of current)for(const msg of validation(e))out.push({event:e,msg});
  for(let i=0;i<current.length;i++)for(let j=i+1;j<current.length;j++){
    const a=current[i],b=current[j];if(a.date!==b.date||!timeCollision(a,b))continue;
    const overlap=baySegments(a.bays).some(x=>baySegments(b.bays).some(y=>x.floor===y.floor&&overlaps(x,y)));
    if(overlap)out.push({event:a,msg:"Bay/time conflict with "+b.name,other:b.id});
  }
  return out;
}
function color(e){return (COLOR[e.category]||COLOR.Flex)[e.status]||COLOR.Flex.Definite}
function actualStatus(e){return e.status==="Definite"?"Definite":"Tentative"}
function layoutDay(events){
  const result={};const overflow=[];
  for(const floor of [3,2,1])for(const part of ["AM","PM"]){
    const key=floor+"-"+part;
    const records=[];
    for(const event of events){if(period(event)!==part)continue;
      for(const range of baySegments(event.bays))if(range.floor===floor)records.push({event,range,row:-1,height:3});
    }
    records.sort((a,b)=>(clockMin(a.event.start)||0)-(clockMin(b.event.start)||0)||a.range.start-b.range.start);
    for(const a of records)if(records.some(b=>a!==b&&overlaps(a.range,b.range)))a.height=2;
    const occupancy=Array.from({length:5},()=>Array(34).fill(null));
    for(const r of records){
      // Stack back-to-back 2-row events at rows 1–2 and 3–4, leaving row 5 blank.
      const starts=r.height===2?[0,2,3,1]:[0,2,1];
      for(const start of starts){
        let free=true;
        for(let y=start;y<start+r.height&&free;y++)for(let x=r.range.start%100-1;x<r.range.end%100;x++)if(occupancy[y][x]){free=false;break}
        if(!free)continue;
        r.row=start;
        for(let y=start;y<start+r.height;y++)for(let x=r.range.start%100-1;x<r.range.end%100;x++)occupancy[y][x]=r.event.id;
        break;
      }
      if(r.row<0)overflow.push({event:r.event,msg:"Not enough vertical room for "+r.range.start+"-"+r.range.end+" ("+part+")"});
    }
    result[key]=records;
  }
  return {records:result,overflow};
}
function couch(floor,bay){
  const x=bay%100;
  return floor===3?(x<=14||x>=29):floor===2?(x<=14||x>=29):(x<=8||x>=29);
}
function neutral(floor,bay){return floor===3&&(bay%100===15||bay%100===16)||floor===1&&(bay%100>=9&&bay%100<=14)}
function renderLandmarks(floor){
  const entries=floor===3?[[9,16,"Terrace"],[17,18,"DOORS","doors"],[21,21,"DOOR","doors"],[25,31,"Upper Lounge"]]:
    floor===2?[[7,10,"Sig North"],[11,14,"Sig South"],[15,20,"3 Heater Bays"],[21,22,"DOORS","doors"]]:
    [[9,14,"3 Heater Bays"],[24,26,"Lower Lounge"],[29,30,"DOORS","doors"]];
  return entries.map(a=>'<div class="landmark '+(a[3]||"")+'" style="grid-column:'+a[0]+' / '+(a[1]+1)+'">'+escapeHtml(a[2])+'</div>').join("");
}
function renderPeriod(floor,part,recordList){
  const plots=recordList.filter(r=>r.row>=0).map(r=>{
    const e=r.event;
    const lines=displayLines(e);
    const tiny=r.range.end-r.range.start<3||r.height===2;
    const hasConflict=allIssues().some(i=>i.event.id===e.id&&i.msg.startsWith("Bay/time conflict"));
    return '<button type="button" class="event-plot'+(tiny?" plot-compact":"")+(e.status==="Tentative"?" tentative":"")+(state.selection===e.id?" selected":"")+(hasConflict?" conflict":"")+'" data-event="'+escapeHtml(e.id)+'" title="'+escapeHtml(e.name+" | "+r.range.start+"-"+r.range.end+" | "+fmtRange(e.start,e.end))+'" style="grid-column:'+(r.range.start%100)+' / '+(r.range.end%100+1)+';grid-row:'+(r.row+1)+' / span '+r.height+';background:'+color(e)+';color:'+(e.status==="Tentative"&&e.category!=="Pleaser"?"#fff":"#151515")+'">'+lines.map((t,i)=>'<span'+(i===0?' class="plot-name"':'')+'>'+escapeHtml(t)+'</span>').join("")+'</button>';
  }).join("");
  return '<div class="period-wrap"><div class="period-label '+part.toLowerCase()+'">'+part+'</div><div class="event-grid">'+plots+'</div></div>';
}
function floorBoard(floor,layout){
  const ids=Array.from({length:34},(_,i)=>floor*100+i+1);
  const hdr=ids.map(n=>'<div class="bay-no '+(couch(floor,n)?"couch":neutral(floor,n)?"neutral":"")+'">'+(state.showBayNumbers?n:"")+'</div>').join("");
  const title={1:"FIRST FLOOR",2:"SECOND FLOOR",3:"THIRD FLOOR"}[floor];
  return '<section class="floor-board"><div class="floor-title">'+title+'</div><div class="floor-main"><div class="landmarks">'+renderLandmarks(floor)+'</div><div class="bay-headers">'+hdr+'</div>'+renderPeriod(floor,"AM",layout.records[floor+"-AM"])+renderPeriod(floor,"PM",layout.records[floor+"-PM"])+'<div class="bay-headers">'+hdr+'</div></div></section>';
}
function render(){
  const has=state.events.length>0;
  document.body.classList.toggle("is-landing",!has);
  $("intro").hidden=has;$("workspace").hidden=!has;$("export").disabled=!has;
  if(!has)return;
  if(!state.week)state.week=weekOf(state.events.find(e=>e.date)?.date||isoDate(new Date()));
  const start=state.week,end=addDays(start,6),current=weeksEvents(),day=dayEvents();
  $("weekTitle").textContent=niceDate(start,{month:"long",day:"numeric"})+" – "+niceDate(end,{month:"long",day:"numeric",year:"numeric"});
  $("weekPickerLabel").textContent=niceDate(start)+" – "+niceDate(end);
  $("fileSummary").textContent=state.files.length+" report"+(state.files.length===1?"":"s")+" · "+state.events.length+" events";
  const issues=allIssues(),layout=layoutDay(day),missing=current.filter(e=>validation(e).length>0).length;
  $("statEvents").textContent=current.length;
  $("statPlotted").textContent=current.filter(e=>baySegments(e.bays).length>0).length;
  $("statNeeds").textContent=new Set(issues.map(x=>x.event.id).concat(layout.overflow.map(x=>x.event.id))).size;
  $("statStatus").textContent=current.filter(e=>e.status==="Definite").length+" / "+current.length;
  $("dayTabs").innerHTML=DAYS.map((x,i)=>{
    const n=current.filter(e=>e.date===addDays(start,i)).length;
    return '<button class="day-tab'+(state.day===i?" active":"")+'" data-day="'+i+'">'+x.slice(0,3)+'<small>'+niceDate(addDays(start,i),{month:"short",day:"numeric"})+' · '+n+'</small></button>';
  }).join("");
  $("dayHeading").textContent=DAYS[state.day]+" · "+niceDate(addDays(start,state.day),{month:"long",day:"numeric"});
  $("daySubheading").textContent="102 bays · AM / PM";
  $("dayTotal").textContent=day.length+" event"+(day.length===1?"":"s");
  $("floorBoards").innerHTML=[3,2,1].map(f=>floorBoard(f,layout)).join("");
  const search=$("eventSearch").value.trim().toLowerCase(),f=$("eventFilter").value;
  const selectedIssues=new Set(issues.map(x=>x.event.id).concat(layout.overflow.map(x=>x.event.id)));
  const shown=day.slice().sort((a,b)=>(clockMin(a.start)||0)-(clockMin(b.start)||0)||a.name.localeCompare(b.name)).filter(e=>(!search||[e.name,e.bays,e.package,e.beo].join(" ").toLowerCase().includes(search))&&(f==="all"||f==="issues"&&selectedIssues.has(e.id)||period(e)===f));
  $("eventListCount").textContent=shown.length;
  $("eventList").innerHTML=shown.length?shown.map(e=>{
    const warn=selectedIssues.has(e.id);const bays=baySegments(e.bays);
    return '<div class="event-item'+(state.selection===e.id?" selected":"")+'" data-event="'+escapeHtml(e.id)+'"><i class="swatch" style="background:'+color(e)+'"></i><div class="event-item-main"><strong>'+escapeHtml(e.name)+'</strong><span>'+escapeHtml(fmtRange(e.start,e.end))+' · '+escapeHtml(e.package)+' · '+e.guests+' guests</span></div><div class="event-item-side '+(warn?"item-warning":"")+'">'+(warn?"⚠ Review":bays.length?escapeHtml(e.bays):"No bays")+'<br>'+period(e)+'</div></div>';
  }).join(""):'<div class="empty">No events match this view.</div>';
  const dailyIssues=issues.filter(x=>x.event.date===addDays(start,state.day)).concat(layout.overflow);
  $("issueCount").textContent=dailyIssues.length;
  $("issues").innerHTML=dailyIssues.length?dailyIssues.map(x=>'<div class="issue" data-event="'+escapeHtml(x.event.id)+'"><strong>'+escapeHtml(x.event.name)+'</strong>'+escapeHtml(x.msg)+'</div>').join(""):'<div class="no-issues">✓ No issues detected for this day.</div>';
  persist();
}
let editId="";
function openEditor(id){
  let e=state.events.find(x=>x.id===id);
  editId=id;
  if(!e){const date=addDays(state.week,state.day);e={name:"",date,start:"",end:"",golfStart:"",golfEnd:"",guests:"",bays:"",space:"",package:"Flex",category:"Flex",status:"Definite"};}
  state.selection=e.id||"";
  $("editorTitle").textContent=id?"Edit event":"New event";
  for(const [id,key] of Object.entries({editName:"name",editDate:"date",editGuests:"guests",editStart:"start",editEnd:"end",editGolfStart:"golfStart",editGolfEnd:"golfEnd",editBays:"bays",editSpace:"space",editPackage:"package",editCategory:"category",editStatus:"status"})){$(id).value=e[key]??""}
  $("deleteEvent").hidden=!id;
  $("editorNote").textContent=e.beo?"BEO #"+e.beo+" · "+e.source:"Manual event";
  $("editorDialog").showModal();render();
}
function closeEditor(){$("editorDialog").close();state.selection="";render()}
function saveEditor(){
  const data={name:$("editName").value.trim(),date:$("editDate").value,start:$("editStart").value,end:$("editEnd").value,golfStart:$("editGolfStart").value,golfEnd:$("editGolfEnd").value,guests:$("editGuests").value===""?null:+$("editGuests").value,bays:$("editBays").value.trim(),space:$("editSpace").value,package:$("editPackage").value.trim(),category:$("editCategory").value,status:$("editStatus").value};
  if(!data.name||!data.date||!data.start||!data.end)return notify("Event name, date, and times are required");
  if(!!data.golfStart!==!!data.golfEnd)return notify("Enter both golf times, or leave both blank");
  if(editId){Object.assign(state.events.find(e=>e.id===editId),data)}
  else{data.id="manual-"+Date.now()+"-"+Math.random().toString(36).slice(2,5);state.events.push(data)}
  state.week=weekOf(data.date);state.day=(localDate(data.date).getDay()+6)%7;closeEditor();notify("Event saved");
}

const csvFiles={staffing:null,menus:null};
async function registerCsvFile(file,requestedType){
  if(!file)return;
  if(!/\.csv$/i.test(file.name))throw new Error("Please choose a CSV report");
  const rows=parseCsv(await file.text()),type=reportType(rows);
  if(requestedType&&requestedType!==type)throw new Error("That file is "+(type==="menus"?"Catering Menus":"Event Staffing")+", not "+(requestedType==="menus"?"Catering Menus":"Event Staffing"));
  csvFiles[type]={file,rows};
  const ready=Boolean(csvFiles.staffing&&csvFiles.menus);
  const set=(type,id,box)=>{
    const picked=csvFiles[type];$(id).textContent=picked?picked.file.name:"Choose CSV";
    $(box).classList.toggle("ready",Boolean(picked));
  };
  set("staffing","staffFileName","staffBox");set("menus","menuFileName","menuBox");
  $("buildFromReports").disabled=!ready;
  $("reportFeedback").textContent=ready?"":"Select both reports.";
  if($("workspace").hidden===false)notify(ready?"Reports ready — build to update floor plan":"Report selected — upload the other CSV");
}
async function acceptCsvFiles(files,autoBuild=false){
  try{
    for(const file of Array.from(files||[]))await registerCsvFile(file);
    if(csvFiles.staffing&&csvFiles.menus&&autoBuild)await runCsvReports();
  }catch(error){
    console.error("CSV import error",error);
    $("reportFeedback").textContent=error.message||String(error);
    notify("CSV: "+(error.message||String(error)));
  }
}
async function runCsvReports(){
  if(!csvFiles.staffing||!csvFiles.menus)return notify("Choose Event Staffing and Catering Menus");
  const btn=$("buildFromReports");btn.disabled=true;btn.textContent="Building…";
  try{
    const incoming=buildEvents(csvFiles.staffing.rows,csvFiles.menus.rows,csvFiles.staffing.file.name,csvFiles.menus.file.name);
    if(!incoming.events.length)throw new Error("No events found in Event Staffing");
    state.events=incoming.events;
    state.files=[csvFiles.staffing.file.name,csvFiles.menus.file.name];
    const dates=incoming.events.filter(e=>e.date).map(e=>e.date).sort();
    state.week=dates.length?weekOf(dates[0]):weekOf(isoDate(new Date()));
    state.day=0;state.selection="";
    $("eventSearch").value="";$("eventFilter").value="all";
    render();
    notify(incoming.events.length+" events · "+incoming.matched+" catering matches"+(incoming.unmatched.length?" · "+incoming.unmatched.length+" unmatched menus":""));
  }catch(error){
    console.error("FPmatic report build failed",error);
    $("reportFeedback").textContent=error.message||String(error);
    notify(error.message||String(error));
  }finally{btn.textContent="Build floor plan";btn.disabled=false}
}
async function importFiles(files){
  const all=[...files].filter(f=>/\.pdf$/i.test(f.name));if(!all.length)return notify("Choose PDF files");
  $("newUpload").disabled=true;
  let found=0;
  try{
    for(const file of all){
      notify("Reading "+file.name+"…");
      const incoming=await parsePdf(file);
      for(const e of incoming){
        const ix=state.events.findIndex(x=>x.id===e.id);
        if(ix<0)state.events.push(e);else state.events[ix]=e;
      }
      found+=incoming.length;
      if(!state.files.includes(file.name))state.files.push(file.name);
    }
    const dates=state.events.filter(e=>e.date).map(e=>e.date).sort();
    if(dates.length){const newDates=state.events.filter(e=>e.date&&all.some(f=>f.name===e.source)).map(e=>e.date).sort();state.week=weekOf(newDates[0]||dates[0]);}
    state.day=0;state.selection="";
    render();notify("Imported "+found+" BEO event"+(found===1?"":"s"));
    if(found===0)notify("No Manager Catering BEO events found. Please check the PDF.");
  }catch(err){console.error(err);notify("PDF parsing error: "+err.message)}
  finally{$("newUpload").disabled=false;$("pdfInput").value=""}
}
async function exportWorkbook(){
  if(!state.events.length)return;
  const button=$("export");
  button.disabled=true;
  const originalText=button.textContent;
  button.textContent="Preparing Excel…";
  try{
    const exported=await exportExactFloorPlan({
      state,weekEvents:weeksEvents(),layoutDay,period,displayLines,addDays
    });
    notify("Downloaded "+exported.filename+" — original master formatting preserved.");
  }catch(err){
    console.error("FPmatic exact-template export failed:",err);
    notify("Export failed: "+(err.message||String(err)));
  }finally{
    button.disabled=false;
    button.textContent=originalText;
  }
}
$("staffInput").addEventListener("change",async e=>{try{await registerCsvFile(e.target.files[0],"staffing")}catch(error){$("reportFeedback").textContent=error.message;notify(error.message)}e.target.value=""});
$("menuInput").addEventListener("change",async e=>{try{await registerCsvFile(e.target.files[0],"menus")}catch(error){$("reportFeedback").textContent=error.message;notify(error.message)}e.target.value=""});
$("csvInput").addEventListener("change",async e=>{await acceptCsvFiles(e.target.files,true);e.target.value=""});
$("buildFromReports").addEventListener("click",runCsvReports);
$("newUpload").addEventListener("click",()=>$("csvInput").click());
$("beoLanding").addEventListener("click",()=>$("pdfInput").click());
$("beoUpload").addEventListener("click",()=>$("pdfInput").click());
$("pdfInput").addEventListener("change",e=>importFiles(e.target.files));
const drop=$("reportsDropzone");
for(const ev of ["dragenter","dragover"])drop.addEventListener(ev,e=>{e.preventDefault();for(const box of drop.querySelectorAll(".report-file"))box.classList.add("dragging")});
for(const ev of ["dragleave","drop"])drop.addEventListener(ev,e=>{e.preventDefault();for(const box of drop.querySelectorAll(".report-file"))box.classList.remove("dragging")});
drop.addEventListener("drop",e=>acceptCsvFiles(e.dataTransfer.files,false));
$("resetAll").addEventListener("click",()=>{if(!confirm("Clear all imported events and edits to start a new week?"))return;state.events=[];state.week="";state.day=0;state.selection="";state.files=[];state.master=null;csvFiles.staffing=null;csvFiles.menus=null;$("staffFileName").textContent="Choose CSV";$("menuFileName").textContent="Choose CSV";$("staffBox").classList.remove("ready");$("menuBox").classList.remove("ready");$("buildFromReports").disabled=true;$("reportFeedback").textContent="";try{localStorage.removeItem(STORAGE)}catch(e){}render();notify("Ready for a new week");});
$("templateInput").addEventListener("change",async e=>{const file=e.target.files[0];if(!file)return;state.master=await file.arrayBuffer();$("templateFileName").textContent=file.name+" ready";notify("Original Excel master loaded for export");});
$("export").addEventListener("click",exportWorkbook);
$("prevWeek").addEventListener("click",()=>{state.week=addDays(state.week,-7);state.day=0;render()});
$("nextWeek").addEventListener("click",()=>{state.week=addDays(state.week,7);state.day=0;render()});
$("dayTabs").addEventListener("click",e=>{const b=e.target.closest("[data-day]");if(b){state.day=+b.dataset.day;state.selection="";render()}});
$("floorBoards").addEventListener("click",e=>{const b=e.target.closest("[data-event]");if(b)openEditor(b.dataset.event)});
$("eventList").addEventListener("click",e=>{const b=e.target.closest("[data-event]");if(b)openEditor(b.dataset.event)});
$("issues").addEventListener("click",e=>{const b=e.target.closest("[data-event]");if(b)openEditor(b.dataset.event)});
$("eventSearch").addEventListener("input",render);
$("eventFilter").addEventListener("change",render);
$("toggleNames").addEventListener("click",()=>{state.showBayNumbers=!state.showBayNumbers;$("toggleNames").textContent=state.showBayNumbers?"Hide bays":"Show bays";render()});
$("addEvent").addEventListener("click",()=>openEditor(""));
$("closeEditor").addEventListener("click",closeEditor);
$("cancelEditor").addEventListener("click",closeEditor);
$("editorForm").addEventListener("submit",e=>{e.preventDefault();saveEditor()});
$("deleteEvent").addEventListener("click",()=>{if(editId&&confirm("Delete this event from the floor plan?")){state.events=state.events.filter(x=>x.id!==editId);closeEditor();notify("Event deleted")}});
$("editorDialog").addEventListener("close",()=>{state.selection="";render()});
restore();render();
