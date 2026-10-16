/*
 * Reconcile a live weekly floorplan against authoritative Event Staffing
 * and Catering Menus reports. Operates only on individual worksheet XML parts.
 * Never changes the original workbook, staffing tabs, or unrelated rows.
 */
const DAY_NAMES=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"];
const EVENT_ROWS=[[4,8],[9,13],[20,24],[25,29],[36,40],[41,45]];
function loc(a){const m=/^([A-Z]+)(\d+)$/.exec(a);if(!m)return null;return {c:[...m[1]].reduce((v,ch)=>v*26+ch.charCodeAt(0)-64,0),r:+m[2]}}
function col(c){let s="";while(c){c--;s=String.fromCharCode(65+c%26)+s;c=Math.floor(c/26)}return s}
function ref(a){return col(a.c1)+a.r1+":"+col(a.c2)+a.r2}
function areaFromRef(s){const [a,b=a]=s.split(":"),x=loc(a),y=loc(b);return {c1:x.c,c2:y.c,r1:x.r,r2:y.r}}
function hit(a,b){return a.c1<=b.c2&&a.c2>=b.c1&&a.r1<=b.r2&&a.r2>=b.r1}
function contains(a,c,r){return c>=a.c1&&c<=a.c2&&r>=a.r1&&r<=a.r2}
function subtract(a,b){if(!hit(a,b))return [a];const k={c1:Math.max(a.c1,b.c1),c2:Math.min(a.c2,b.c2),r1:Math.max(a.r1,b.r1),r2:Math.min(a.r2,b.r2)};const x=[];if(a.r1<k.r1)x.push({...a,r2:k.r1-1});if(k.r2<a.r2)x.push({...a,r1:k.r2+1});if(a.c1<k.c1)x.push({c1:a.c1,c2:k.c1-1,r1:k.r1,r2:k.r2});if(k.c2<a.c2)x.push({c1:k.c2,c2:a.c2,r1:k.r1,r2:k.r2});return x}
function normalized(s){return String(s||"").normalize("NFKD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[’‘]/g,"'").replace(/[^a-z0-9]/g,"")}
function decode(s){return String(s||"").replace(/&(?:amp|lt|gt|quot|apos|#(\d+)|#x([a-f0-9]+));/gi,(m,dec,hex)=>dec?String.fromCharCode(+dec):hex?String.fromCharCode(parseInt(hex,16)):({"&amp;":"&","&lt;":"<","&gt;":">","&quot;":'"',"&apos;":"'"}[m]||m))}
function encode(s){return String(s??"").replace(/[&<>"']/g,ch=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&apos;"}[ch]))}
function readShared(xml){return [...String(xml||"").matchAll(/<si(?:\s[^>]*)?>[\s\S]*?<\/si>/g)].map(m=>decode([...m[0].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(t=>t[1]).join("")))}
function readCellNodes(xml,shared){
  const result=new Map();
  const re=/<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*\/>|<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*>[\s\S]*?<\/c>/g;
  for(const m of xml.matchAll(re)){
    const address=m[1]||m[2],tag=m[0],type=(tag.match(/\bt="([^"]+)"/)||[])[1]||"",s=+(tag.match(/\bs="(\d+)"/)||[])[1]||0;
    const v=(tag.match(/<v>([\s\S]*?)<\/v>/)||[])[1]||"";
    const rich=[...tag.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(x=>x[1]).join("");
    const text=type==="s"?(shared[+v]||""):type==="inlineStr"?decode(rich):type==="str"?decode(v):"";
    result.set(address,{tag,style:s,text});
  }
  return result;
}
function parseExistingEvent(text){
  const lines=String(text||"").split(/\r?\n/).map(x=>x.trim()).filter(Boolean);
  if(lines.length<2||!lines.some(x=>/\b\d+\s*ppl\b/i.test(x)))return null;
  const guests=(text.match(/(\d+)\s*ppl\b/i)||[])[1];
  const packageLine=lines.find(x=>/^(?:FX|Flex|Fajita|Smokehouse|Pleaser|Youth|Field Trip|Ultimate|Ult Bites|Premium|MVB|Buffet|Internal|Taster|Rise & Shine)$/i.test(x))||"";
  const extra=lines.filter((x,i)=>i>0&&( /^(?:EAO|EA |EA:|EA Bar|EA Host|.*(?:HH\d+|online learning))/i.test(x)||/\b(?:EAO|EA Bar)\s*[-:]/i.test(x)));
  return {name:lines[0],guests:guests?+guests:null,package:packageLine,description:lines.slice(1),extra};
}
function eventInfo(area,text){
  const p=parseExistingEvent(text);if(!p)return null;
  const floor=area.r1>=36?1:area.r1>=20?2:3;
  return {...p,area,floor,bays:floor*100+area.c1-4+"-"+(floor*100+area.c2-4),startRow:area.r1};
}
function timeLabel(s){const [h,m]=String(s||"").split(":").map(Number);if(!Number.isFinite(h))return "";return (h%12||12)+(m?String(m).padStart(2,"0"):"")}
function desiredText(e,displayLines,old){
  const lines=displayLines(e);
  // Preserve the staffing/EA notes attached to an existing event when it moves.
  for(const line of old?.extra||[])if(!lines.some(x=>normalized(x)===normalized(line)))lines.push(line);
  return lines.join("\n");
}
function matchEvent(e,old){
  const name=normalized(e.name),known=normalized(old.name);
  if(!name||!known)return false;
  if(name===known)return true;
  // Some workbooks omit an exclamation mark, possessive, "Event!", etc.
  return Math.min(name.length,known.length)>=9&&
    (name.startsWith(known)||known.startsWith(name))&&
    Math.min(name.length,known.length)/Math.max(name.length,known.length)>0.83;
}
function pickRow(e,range,staffAreas,protectedEvents,reserved,preferred){
  const floor=Math.floor(range.start/100),part=(e.start&&e.end&&(Number(e.start.slice(0,2))>=16||Number(e.end.slice(0,2))>=19))?"PM":"AM";
  const top=({3:4,2:20,1:36}[floor])+(part==="PM"?5:0);
  const count=preferred?.r2-preferred?.r1+1===2?2:3;
  const candidates=[];
  if(preferred&&preferred.r1>=top&&preferred.r2<=top+4)candidates.push(preferred.r1);
  for(const row of count===2?[top,top+2,top+3,top+1]:[top,top+2,top+1])if(!candidates.includes(row))candidates.push(row);
  for(const row of candidates){
    if(row+count-1>top+4)continue;
    const area={c1:4+range.start%100,c2:4+range.end%100,r1:row,r2:row+count-1};
    if(protectedEvents.some(x=>hit(x.area,area))||reserved.some(x=>hit(x,area)))continue;
    // A staffing block may be divided, but its assignment must have room to stay.
    if(staffAreas.some(x=>hit(x.area,area)&&!subtract(x.area,area).length&&x.cell.text.trim()))continue;
    return area;
  }
  return null;
}
function parseBays(s){
  const ranges=[];
  for(const m of String(s||"").matchAll(/(\d{3})(?:\s*[-–]\s*(\d{3}))?/g)){
    const a=+m[1],b=+(m[2]||m[1]);if(a>b||Math.floor(a/100)!==Math.floor(b/100)||a%100<1||b%100>34)continue;
    ranges.push({start:a,end:b});
  }
  return ranges;
}
function setCell(xml,updates){
  const seen=new Set();
  xml=xml.replace(/<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*\/>|<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*>[\s\S]*?<\/c>/g,(all,a,b)=>{
    const key=a||b,u=updates.get(key);
    if(!u)return all;seen.add(key);return u;
  });
  const missing=[...updates].filter(([key])=>!seen.has(key));
  if(!missing.length)return xml;
  const rows=new Map();
  for(const [address,value] of missing){
    const p=loc(address);if(!rows.has(p.r))rows.set(p.r,[]);
    rows.get(p.r).push([p.c,address,value]);
  }
  xml=xml.replace(/<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g,(block,n)=>{
    const entries=rows.get(+n);if(!entries)return block;
    rows.delete(+n);
    let b=block;
    for(const [colIndex,_,tag] of entries.sort((a,b)=>b[0]-a[0])){
      const next=[...b.matchAll(/<c\b[^>]*\br="([A-Z]{1,3}\d+)"/g)].find(x=>loc(x[1]).c>colIndex);
      if(next)b=b.slice(0,next.index)+tag+b.slice(next.index);
      else b=b.replace("</row>",tag+"</row>");
    }
    return b;
  });
  if(rows.size){const extra=[...rows].sort((a,b)=>a[0]-b[0]).map(([r,entries])=>'<row r="'+r+'">'+entries.sort((a,b)=>a[0]-b[0]).map(x=>x[2]).join("")+'</row>').join("");xml=xml.replace("</sheetData>",extra+"</sheetData>")}
  return xml;
}
function eventCell(ref,text,style){return '<c r="'+ref+'" s="'+style+'" t="inlineStr"><is><t xml:space="preserve">'+encode(text)+'</t></is></c>'}
function blankCell(ref,style){return '<c r="'+ref+'" s="'+style+'"/>'}
export function reconcileDay(xml,shared,events,styles,displayLines){
  const cells=readCellNodes(xml,shared);
  const mergeNodes=[...xml.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"[^>]*\/>/g)].map(m=>areaFromRef(m[1]));
  const original=[];
  for(const m of mergeNodes){
    const address=col(m.c1)+m.r1,cell=cells.get(address);
    if(!cell?.text)continue;
    const info=eventInfo(m,cell.text);
    if(info&&m.c1>=5&&m.c2<=38&&EVENT_ROWS.some(([a,b])=>m.r1>=a&&m.r2<=b))original.push({...info,cell,address});
  }
  const matched=new Map(),used=new Set();
  for(const e of events){
    const matches=original.filter(o=>!used.has(o)&&matchEvent(e,o));
    if(matches.length===1){matched.set(e,matches[0]);used.add(matches[0])}
  }
  const protectedEvents=original.filter(o=>!used.has(o));
  const staff=mergeNodes.flatMap(area=>{
    const address=col(area.c1)+area.r1,cell=cells.get(address);
    if(!cell?.text.trim()||original.some(o=>o.address===address))return [];
    if(area.c1<5||area.c2>38||!EVENT_ROWS.some(([a,b])=>area.r1>=a&&area.r2<=b))return [];
    return [{area,cell,address}];
  });
  const changes=[],warnings=[],planned=[],reserved=[];
  for(const e of events){
    const old=matched.get(e),bays=parseBays(e.bays);
    if(!bays.length){warnings.push(e.name+": no assigned bays");continue}
    const ranges=bays.map(range=>pickRow(e,range,staff,protectedEvents,reserved,old?.area));
    if(ranges.some(x=>!x)){warnings.push(e.name+": could not place without overwriting another event/staff assignment");continue}
    ranges.forEach(r=>reserved.push(r));
    if(old){
      if(old.guests!==null&&e.guests!==null&&old.guests!==e.guests)
        changes.push({name:e.name,type:"guests",before:old.guests+"ppl",after:e.guests+"ppl"});
      const oldBays=old.bays,updatedBays=bays.map(x=>x.start+"-"+x.end).join(", ");
      if(normalized(oldBays)!==normalized(updatedBays))
        changes.push({name:e.name,type:"bays",before:oldBays,after:updatedBays});
      if(old.package&&e.package&&normalized(old.package)!==normalized(e.package)&&!(old.package==="FX"&&e.package==="Flex"))
        changes.push({name:e.name,type:"package",before:old.package,after:e.package});
    }else changes.push({name:e.name,type:"added",before:"Not on floorplan",after:e.bays});
    for(const area of ranges)planned.push({area,event:e,old,text:desiredText(e,displayLines,old)});
  }
  if(!planned.length)return {xml,changes,warnings,placed:0};
  const updates=new Map(),deleted=new Set(),newMerges=mergeNodes.slice();
  const removing=new Set([...planned].map(p=>p.old).filter(Boolean));
  function removeMerge(a){const i=newMerges.findIndex(v=>ref(v)===ref(a));if(i>=0)newMerges.splice(i,1)}
  function addMerge(a){if(a.c1!==a.c2||a.r1!==a.r2)newMerges.push(a)}
  function clearArea(area,style){
    for(let r=area.r1;r<=area.r2;r++)for(let c=area.c1;c<=area.c2;c++)
      updates.set(col(c)+r,blankCell(col(c)+r,styles.clear(style)));
  }
  // First clear the old matched blocks. If a staff allocation directly below
  // has identical bay columns, extend that staff merge into the vacated space
  // (Stephen above Kelly is an actual example).
  for(const old of removing){
    removeMerge(old.area);
    clearArea(old.area,old.cell.style);
    const neighbor=staff.find(s=>s.area.c1===old.area.c1&&s.area.c2===old.area.c2&&
      (s.area.r1===old.area.r2+1||s.area.r2+1===old.area.r1)&&
      EVENT_ROWS.some(([a,b])=>Math.min(s.area.r1,old.area.r1)>=a&&Math.max(s.area.r2,old.area.r2)<=b));
    if(neighbor){
      removeMerge(neighbor.area);
      const enlarged={...neighbor.area,r1:Math.min(neighbor.area.r1,old.area.r1),r2:Math.max(neighbor.area.r2,old.area.r2)};
      addMerge(enlarged);
      const top=col(enlarged.c1)+enlarged.r1;
      updates.set(neighbor.address,blankCell(neighbor.address,styles.clear(neighbor.cell.style)));
      updates.set(top,eventCell(top,neighbor.cell.text,styles.center(neighbor.cell.style)));
      neighbor.area=enlarged;
    }else{
      addMerge(old.area);
    }
  }
  // Merge original staff/blank sections around each new booking. Reposition a
  // displaced staff label within the largest leftover fragment, never discard.
  for(const {area,event,old,text} of planned){
    for(const existing of newMerges.slice()){
      if(!hit(existing,area))continue;
      const ownedStaff=staff.find(s=>ref(s.area)===ref(existing));
      removeMerge(existing);
      const pieces=subtract(existing,area).filter(x=>x.c1<=x.c2&&x.r1<=x.r2);
      if(ownedStaff&& !pieces.length){warnings.push(event.name+": staff block fully covered");continue}
      for(const piece of pieces)addMerge(piece);
      if(ownedStaff&&contains(area,ownedStaff.area.c1,ownedStaff.area.r1)&&pieces.length){
        const best=pieces.slice().sort((a,b)=>(b.c2-b.c1+1)*(b.r2-b.r1+1)-(a.c2-a.c1+1)*(a.r2-a.r1+1))[0];
        updates.set(ownedStaff.address,blankCell(ownedStaff.address,styles.clear(ownedStaff.cell.style)));
        const anchor=col(best.c1)+best.r1;
        updates.set(anchor,eventCell(anchor,ownedStaff.cell.text,styles.center(ownedStaff.cell.style)));
      }
    }
    addMerge(area);
    for(let r=area.r1;r<=area.r2;r++)for(let c=area.c1;c<=area.c2;c++){
      const address=col(c)+r,sourceStyle=cells.get(address)?.style||old?.cell.style||0;
      const style=styles.event(sourceStyle,event,r===area.r1&&c===area.c1?area.r2-area.r1+1===2:false);
      updates.set(address,r===area.r1&&c===area.c1?eventCell(address,text,style):blankCell(address,style));
    }
  }
  xml=setCell(xml,updates);
  const unique=[...new Map(newMerges.map(m=>[ref(m),m])).values()];
  xml=xml.replace(/<mergeCells\b([^>]*)>[\s\S]*?<\/mergeCells>/,block=>
    block.replace(/\bcount="\d+"/,'count="'+unique.length+'"')
      .replace(/<mergeCell\b[^>]*\/>/g,"")
      .replace("</mergeCells>",unique.map(m=>'<mergeCell ref="'+ref(m)+'"/>').join("")+"</mergeCells>"));
  return {xml,changes,warnings,placed:planned.length};
}
export function workbookDayPaths(workbookXml){
  const out={};
  for(const tag of workbookXml.match(/<sheet\b[^>]*\/>/g)||[]){
    const name=decode((tag.match(/\bname="([^"]+)"/)||[])[1]||"");
    const d=DAY_NAMES.findIndex(day=>name===day||name.startsWith(day+" "));
    if(d<0||name.startsWith("AM ")||/\bAM$/.test(name))continue;
    if(!(d in out)){
      const index=+(tag.match(/\bsheetId="(\d+)"/)||[])[1];
      if(index)out[d]="xl/worksheets/sheet"+index+".xml";
    }
  }
  return out;
}
export function getLiveEvents(xml,shared){
  const cells=readCellNodes(xml,shared);
  const merged=(xml.match(/<mergeCell\b[^>]*ref="([^"]+)"[^>]*\/>/g)||[]).map(x=>areaFromRef((x.match(/\bref="([^"]+)"/)||[])[1]));
  const result=[];
  for(const area of merged){
    const anchor=col(area.c1)+area.r1,cell=cells.get(anchor);
    if(!cell)continue;
    const isFloor=area.c1>=5&&area.c2<=38&&EVENT_ROWS.some(([a,b])=>area.r1>=a&&area.r2<=b);
    if(!isFloor)continue;
    result.push({area,text:cell.text,style:cell.style,isEvent:Boolean(parseExistingEvent(cell.text))});
  }
  return result;
}
export const parseFloorplanCell=readCellNodes;
export const parseFloorplanShared=readShared;
