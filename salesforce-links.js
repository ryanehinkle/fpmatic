// Salesforce report link extraction + conservative multifactor event matching.
// Keep parsing and scoring browser-only, never transmit clipboard data.
function decodeEntities(input){
  const entity={"amp":"&","lt":"<","gt":">","quot":'"',"apos":"'","nbsp":" ","#39":"'","ndash":"–","mdash":"—"};
  return String(input||"").replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi,(_,code)=>{
    const key=code.toLowerCase();
    if(key.startsWith("#")){
      const n=parseInt(key.slice(key[1]==="x"?2:1),key[1]==="x"?16:10);
      return Number.isFinite(n)&&n>=0&&n<=0x10ffff?String.fromCodePoint(n):"";
    }
    return entity[key]??("&"+code+";");
  });
}
function textOfHtml(html){return decodeEntities(String(html||"").replace(/<br\s*\/?>/gi," ").replace(/<\/(?:div|p|span|td|th)>/gi," ").replace(/<[^>]*>/g," ")).replace(/\s+/g," ").trim()}
export function canonicalOpportunity(name){
  return String(name||"").normalize("NFKD").replace(/[\u0300-\u036f]/g,"")
    .replace(/[\u2018\u2019\u2032]/g,"'").replace(/[^a-z0-9]+/gi," ").trim().toLowerCase().replace(/\s+/g," ");
}
export function opportunityDate(full){
  const m=String(full||"").match(/(?:^|[-\s])(\d{1,2})[./](\d{1,2})[./](\d{2}|\d{4})\s*$/);
  if(!m)return "";
  return (m[3].length===2?"20"+m[3]:m[3])+"-"+m[1].padStart(2,"0")+"-"+m[2].padStart(2,"0");
}
export function shortOpportunityName(full){
  let s=String(full||"").replace(/\s*-\s*\d{1,2}[./]\d{1,2}[./]\d{2,4}\s*$/,"").trim();
  const parts=s.split(/\s+-\s+/);
  if(parts.length>1)parts.pop();
  return parts.join(" - ")||s;
}
export function safeSalesforceUrl(href,base=""){
  try{
    const decoded=decodeEntities(href);
    const address=new URL(decoded,base||undefined);
    if(address.protocol!=="https:"||address.username||address.password)return "";
    const host=address.hostname.toLowerCase();
    if(!["salesforce.com","force.com","salesforce-sites.com"].some(s=>host===s||host.endsWith("."+s)))return "";
    return address.href;
  }catch{return ""}
}
function parseClock(input){
  const s=String(input||"").trim();
  const m=s.match(/^(\d{1,2})(?::(\d{2}))?\s*([AP]M)?$/i);
  if(!m)return null;
  let h=+m[1],min=+(m[2]||0);
  if(min>59||h>23||m[3]&&!(h>=1&&h<=12))return null;
  if(m[3])h=(h%12)+(m[3].toLowerCase()==="pm"?12:0);
  return h*60+min;
}
function rangesFromText(bays){
  const set=new Set();
  for(const m of String(bays||"").matchAll(/\b([123]\d{2})\s*[-–]\s*([123]\d{2})\b|\b([123]\d{2})\b/g)){
    const a=+(m[1]||m[3]),b=+(m[2]||m[3]);if(a>b||b-a>34||Math.floor(a/100)!==Math.floor(b/100))continue;
    for(let n=a;n<=b;n++)set.add(n);
  }
  return set;
}
function nameSimilarity(a,b){
  a=canonicalOpportunity(a);b=canonicalOpportunity(b);
  if(!a||!b)return 0;
  if(a===b)return 1;
  const x=a.split(" "),y=b.split(" ");
  const sx=new Set(x.filter(t=>t.length>=2)),sy=new Set(y.filter(t=>t.length>=2));
  const shared=[...sx].filter(t=>sy.has(t)).length;
  const dice=2*shared/(sx.size+sy.size||1);
  const bigrams=s=>{const arr=[];for(let i=0;i<s.length-1;i++)arr.push(s.slice(i,i+2));return arr};
  const bx=bigrams(a.replace(/\s/g,"")),by=bigrams(b.replace(/\s/g,""));
  const counts=new Map();for(const t of bx)counts.set(t,(counts.get(t)||0)+1);
  let common=0;for(const t of by)if(counts.get(t)>0){common++;counts.set(t,counts.get(t)-1)}
  const character=2*common/(bx.length+by.length||1);
  return Math.max(0.62*dice+0.38*character,(a.includes(b)||b.includes(a))?0.83*Math.min(a.length,b.length)/Math.max(a.length,b.length)+0.17:0);
}
function sameClock(a,b){
  const x=parseClock(a),y=parseClock(b);
  if(x===null||y===null)return false;
  return Math.abs(x-y)<=5||Math.abs((x%720)-(y%720))<=5;
}
function floorplanTime(text){
  const lines=String(text||"").split(/\r?\n/);
  const line=lines.find(s=>/\b\d{1,4}\s*[-–]\s*\d{1,4}\b/.test(s)&&!/\bGolf\b/i.test(s))||"";
  const m=line.match(/\b(\d{1,4})\s*[-–]\s*(\d{1,4})\b/);
  if(!m)return {};
  const p=n=>{let x=String(n);if(x.length>2)x=x.slice(0,-2)+":"+x.slice(-2);return x};
  return {start:p(m[1]),end:p(m[2])};
}
function guessTimeCells(cells,anchorIndex){
  const start=String(cells[anchorIndex-3]||"").trim();
  const end=String(cells[anchorIndex-2]||"").trim();
  if(parseClock(start)!==null&&parseClock(end)!==null)return {start,end};
  const prior=cells.slice(0,anchorIndex).filter(x=>parseClock(x)!==null);
  return {start:prior.at(-2)||"",end:prior.at(-1)||""};
}
function guessRowMeta(cells,title){
  const idx=cells.findIndex(s=>canonicalOpportunity(s)===canonicalOpportunity(title));
  if(idx<0)return {};
  const times=guessTimeCells(cells,idx);
  const end=cells.slice(idx+1);
  // Salesforce report column positions after Opportunity Name:
  // Stage, Event Card, Bay Numbers, # Bays, Final Balance, Guest Count.
  const bays=/\b[123]\d{2}\s*[-–]\s*[123]\d{2}\b/.test(end[2]||"")?end[2]:
    end.slice(0,6).find(s=>/^\s*[123]\d{2}\s*(?:[-–]\s*[123]\d{2})?\s*$/.test(s))||"";
  const guest=String(end[5]||"").replace(/,/g,"").trim();
  return {...times,bays,guests:/^\d+$/.test(guest)?+guest:null};
}
function tableCells(html){
  return [...html.matchAll(/<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]\s*>/gi)].map(m=>textOfHtml(m[1]));
}
function parseRow(html,index,anchorName){
  const upto=html.slice(0,index),start=upto.toLowerCase().lastIndexOf("<tr");
  const end=html.indexOf("</tr>",index);
  if(start<0||end<0)return {};
  const row=html.slice(start,end+5);
  return guessRowMeta(tableCells(row),anchorName);
}
export function extractSalesforceLinks(html){
  const raw=String(html||"");
  const links=[],byName=new Map();
  const source=raw.match(/SourceURL:\s*(https:\/\/[^\s<]+)/i)?.[1]||"";
  let base=safeSalesforceUrl(source);
  if(!base){
    const first=[...raw.matchAll(/\bhref\s*=\s*["'](https:\/\/[^"']+)["']/gi)].find(m=>safeSalesforceUrl(m[1]));
    if(first)base=safeSalesforceUrl(first[1]);
  }
  for(const m of raw.matchAll(/<a\b([^>]*)>([\s\S]*?)<\/a\s*>/gi)){
    const href=(m[1].match(/\bhref\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/i)||[]);
    const url=safeSalesforceUrl(href[1]||href[2]||href[3]||"",base);
    const name=textOfHtml(m[2]);
    const date=opportunityDate(name);
    if(!url||!date)continue;
    const key=canonicalOpportunity(name),row=parseRow(raw,m.index,name);
    if(byName.has(key)){
      const existing=links[byName.get(key)];
      for(const prop of ["start","end","bays","guests"])if(!existing[prop]&&row[prop])existing[prop]=row[prop];
      continue;
    }
    byName.set(key,links.length);
    links.push({name,url,date,shortName:shortOpportunityName(name),...row});
  }
  // DOM fallback handles Salesforce's role=row div-based tables when available.
  if(typeof DOMParser==="function"){
    try{
      const doc=new DOMParser().parseFromString(raw,"text/html");
      for(const anchor of doc.querySelectorAll("a[href]")){
        const key=canonicalOpportunity(anchor.textContent);
        const i=byName.get(key);if(i===undefined)continue;
        const tr=anchor.closest("tr,[role=row]");
        if(!tr)continue;
        const cells=[...tr.querySelectorAll("td,th,[role=cell],[role=gridcell]")].map(x=>x.textContent.replace(/\s+/g," ").trim());
        const row=guessRowMeta(cells,anchor.textContent);
        for(const prop of ["start","end","bays","guests"])if(links[i][prop]===undefined||links[i][prop]==="")links[i][prop]=row[prop];
      }
    }catch{/* HTML regex path remains available */}
  }
  return links;
}
function overlapRatio(a,b){
  const x=rangesFromText(a),y=rangesFromText(b);
  if(!x.size||!y.size)return null;
  const intersection=[...x].filter(n=>y.has(n)).length;
  return intersection?intersection/Math.max(x.size,y.size):0;
}
function score(candidate,record){
  const a=record.opportunityName||record.name||"";
  const name=Math.max(nameSimilarity(a,candidate.name),nameSimilarity(a,candidate.shortName),nameSimilarity(record.name,candidate.shortName));
  const start=Boolean(record.start&&candidate.start&&sameClock(record.start,candidate.start));
  const end=Boolean(record.end&&candidate.end&&sameClock(record.end,candidate.end));
  const bay=overlapRatio(record.bays,candidate.bays);
  const guest=record.guests!==null&&record.guests!==undefined&&candidate.guests!==null&&candidate.guests!==undefined
    ?Math.abs(Number(record.guests)-Number(candidate.guests)):null;
  let points=name*65+(name>=0.95?14:0)+(start?16:0)+(end?16:0)+
    (bay!==null?(bay===1?29:Math.round(bay*23)):0)+(guest===0?13:guest!==null&&guest<=2?7:0);
  // Strong contradictory evidence should beat a superficially similar title.
  if(candidate.start&&record.start&&!start)points-=9;
  if(candidate.end&&record.end&&!end)points-=9;
  if(bay===0)points-=12;
  const timePair=start&&end,baySupport=bay!==null&&bay>=0.5;
  const eligible=name>=0.70&&(points>=52)||name>=0.36&&
    (timePair&&guest!==null&&guest<=2||baySupport&&(start||end||guest!==null&&guest<=2))&&points>=58;
  return {points,eligible,name,timePair,baySupport};
}
export function matchingSalesforceLink(record,links){
  if(!record||!Array.isArray(links)||!links.length)return null;
  const date=record.date;
  if(!date)return null;
  const onDate=links.filter(x=>x.date===date);
  if(!onDate.length)return null;
  const full=canonicalOpportunity(record.opportunityName);
  if(full){
    const exact=onDate.filter(x=>canonicalOpportunity(x.name)===full);
    if(exact.length===1)return exact[0];
  }
  const ranked=onDate.map(x=>({item:x,...score(x,record)})).filter(x=>x.eligible)
    .sort((a,b)=>b.points-a.points);
  if(!ranked.length)return null;
  if(ranked.length>1&&ranked[0].points-ranked[1].points<9)return null;
  return ranked[0].item;
}
export function linkForWorksheetText(text,date,events,links,bays=""){
  const lines=String(text||"").split(/\r?\n/);
  const name=lines[0]?.trim()||"";
  if(!name)return null;
  const times=floorplanTime(text);
  const guests=Number((String(text).match(/(\d+)\s*ppl/i)||[])[1]);
  const record={name,date,start:times.start,end:times.end,
    guests:Number.isFinite(guests)?guests:null,bays};
  const fromEvents=(events||[]).filter(e=>e.date===date&&e.salesforceUrl)
    .map(e=>({name:e.opportunityName||e.name,shortName:e.name,url:e.salesforceUrl,date:e.date,
      bays:e.bays,start:e.start,end:e.end,guests:e.guests}));
  return matchingSalesforceLink(record,fromEvents)?.url||
    matchingSalesforceLink(record,links)?.url||null;
}
