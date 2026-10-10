/*
 * Format-preserving Excel export for FPmatic.
 * We never read the original workbook through XLSX parsers or rebuild worksheets.
 * Every original ZIP part stays byte-identical after decompression except:
 *   1. xl/styles.xml: append event-only styles
 *   2. xl/worksheets/sheet1..7.xml: change only assigned event cells + merges
 * This preserves widths, heights, frozen views, landmarks, print settings,
 * borders, reference tabs, existing tables and all other master formatting.
 */
const DAYS=["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"];
const COLORS={
  Flex:["D9EAD3","6AA84F"],
  Pleaser:["FFF2CC","F1C232"],
  Buffet:["C9DAF8","3C78D8"],
  "Field Trip":["D9D2E9","674EA7"],
  Youth:["F4CCCC","E06666"]
};
function xmlEscape(value){
  return String(value ?? "").replace(/[<>&"']/g, ch=>({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&apos;"
  }[ch]));
}
function replaceSection(xml,sectionName,transform){
  const re=new RegExp("<"+sectionName+"\\b[^>]*>[\\s\\S]*?<\\/"+sectionName+">");
  const match=xml.match(re);
  if(!match)throw new Error("The Excel master has no "+sectionName+" section");
  const replacement=transform(match[0]);
  return xml.replace(match[0],replacement);
}
function recolorStyles(xml){
  const colorToFill=new Map();
  let fillCount=0;
  xml=replaceSection(xml,"fills",section=>{
    const tags=[...section.matchAll(/<fill(?:\s[^>]*)?>[\s\S]*?<\/fill>/g)];
    fillCount=tags.length;
    tags.forEach((hit,index)=>{
      const rgb=(hit[0].match(/<fgColor\b[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/)||[])[1];
      if(rgb)colorToFill.set(rgb.toUpperCase(),index);
    });
    let toAdd="";
    Object.values(COLORS).flat().forEach(color=>{
      if(colorToFill.has(color))return;
      colorToFill.set(color,fillCount++);
      toAdd+='<fill><patternFill patternType="solid"><fgColor rgb="FF'+color+'"/><bgColor indexed="64"/></patternFill></fill>';
    });
    return section.replace(/<fills\b([^>]*)>/,(_,attrs)=>'<fills'+attrs.replace(/\bcount="\d+"/,'count="'+fillCount+'"')+'>').replace("</fills>",toAdd+"</fills>");
  });
  let whiteFontId=-1;
  xml=replaceSection(xml,"fonts",section=>{
    const count=Number((section.match(/<fonts\b[^>]*count="(\d+)"/)||[])[1]);
    if(!Number.isFinite(count))throw new Error("Excel template font table is invalid");
    whiteFontId=count;
    const extra='<font><b/><sz val="20"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>';
    return section.replace(/<fonts\b([^>]*)>/,(_,attrs)=>'<fonts'+attrs.replace(/\bcount="\d+"/,'count="'+(count+1)+'"')+'>').replace("</fonts>",extra+"</fonts>");
  });
  const styleIds={};
  xml=replaceSection(xml,"cellXfs",section=>{
    const styles=[...section.matchAll(/<xf(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/xf>)/g)];
    const start=styles.length;
    let additions="";
    let n=0;
    for(const [category,colors] of Object.entries(COLORS)){
      styleIds[category]={};
      for(let tentative=0;tentative<2;tentative++){
        const status=tentative?"Tentative":"Definite",fill=colorToFill.get(colors[tentative]);
        styleIds[category][status]={};
        for(const compact of [false,true]){
          const fontId=tentative && ["Buffet","Flex","Field Trip","Youth"].includes(category)?whiteFontId:(compact?9:18);
          const name=compact?"compact":"normal";
          styleIds[category][status][name]=start+n++;
          additions+='<xf numFmtId="0" fontId="'+fontId+'" fillId="'+fill+'" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>';
        }
      }
    }
    return section.replace(/<cellXfs\b([^>]*)>/,(_,attrs)=>'<cellXfs'+attrs.replace(/\bcount="\d+"/,'count="'+(start+n)+'"')+'>').replace("</cellXfs>",additions+"</cellXfs>");
  });
  return {xml,styleIds};
}
function colName(n){
  let s="";
  while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}
  return s;
}
function rectToCell(range,baseRow){
  const colStart=4+range.start%100,colEnd=4+range.end%100;
  return {c1:colStart,c2:colEnd,r1:baseRow,r2:baseRow};
}
function plotSheet(xml,events,layout,styleIds,period,displayLines){
  const assignments=new Map(),newMerges=[];
  let assigned=0;
  const ranges=Object.values(layout.records).flat();
  for(const rec of ranges){
    if(rec.row<0)continue;
    const event=rec.event;
    const p=period(event);
    const base={3:4,2:20,1:36}[rec.range.floor];
    if(!base)continue;
    const r1=base+(p==="PM"?5:0)+rec.row,r2=r1+rec.height-1;
    const c1=4+rec.range.start%100,c2=4+rec.range.end%100;
    const fillStyle=(styleIds[event.category]||styleIds.Flex)[event.status==="Tentative"?"Tentative":"Definite"][rec.height===2?"compact":"normal"];
    const first=colName(c1)+r1,last=colName(c2)+r2;
    for(let row=r1;row<=r2;row++)for(let col=c1;col<=c2;col++){
      const address=colName(col)+row;
      if(assignments.has(address))throw new Error("Events overlap in the export at "+address);
      assignments.set(address,{
        style:fillStyle,
        text:(row===r1&&col===c1)?displayLines(event).join("\n"):null
      });
    }
    newMerges.push(first+":"+last);
    assigned++;
  }
  if(!assignments.size)return {xml,count:assigned};
  let replacements=0;
  // Every grid coordinate in the original master is already a <c> element.
  // Rewrite these elements only. Leave every non-event cell and sheet XML
  // attribute untouched, including row heights and existing merge definitions.
  xml=xml.replace(/<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*?\/>|<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*>[\s\S]*?<\/c>/g,(original,selfClosingRef,normalRef)=>{
    const ref=selfClosingRef||normalRef;
    const change=assignments.get(ref);
    if(!change)return original;
    replacements++;
    const s=String(change.style);
    if(change.text===null)return '<c r="'+ref+'" s="'+s+'"/>';
    return '<c r="'+ref+'" s="'+s+'" t="inlineStr"><is><t xml:space="preserve">'+xmlEscape(change.text)+'</t></is></c>';
  });
  if(replacements!==assignments.size){
    throw new Error("Master is missing "+(assignments.size-replacements)+" expected bay cells; no file was exported.");
  }
  xml=replaceSection(xml,"mergeCells",section=>{
    const existing=new Set([...section.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"/g)].map(x=>x[1]));
    for(const ref of newMerges)if(existing.has(ref))throw new Error("Master already has an event merge at "+ref);
    const count=Number((section.match(/<mergeCells\b[^>]*count="(\d+)"/)||[])[1]||0);
    return section.replace(/<mergeCells\b([^>]*)>/,(_,attrs)=>'<mergeCells'+attrs.replace(/\bcount="\d+"/,'count="'+(count+newMerges.length)+'"')+'>')
      .replace("</mergeCells>",newMerges.map(ref=>'<mergeCell ref="'+ref+'"/>').join("")+"</mergeCells>");
  });
  return {xml,count:assigned};
}
function formatDownloadDate(date){
  const [year,month,day]=date.split("-");
  return month+"."+day+"."+year.slice(-2);
}
async function downloadBinary(data,name){
  const url=URL.createObjectURL(new Blob([data],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
  const anchor=document.createElement("a");
  anchor.href=url;anchor.download=name;document.body.appendChild(anchor);anchor.click();anchor.remove();
  setTimeout(()=>URL.revokeObjectURL(url),30000);
}
export async function exportExactFloorPlan({state,weekEvents,layoutDay,period,displayLines,addDays}){
  if(typeof JSZip!=="function")throw new Error("Local ZIP export library failed to load");
  const masterBytes=state.master||(await (async()=>{
    const response=await fetch("./master.xlsx",{cache:"no-cache"});
    if(!response.ok)throw new Error("Original floor plan master could not be loaded");
    return await response.arrayBuffer();
  })());
  const zip=await JSZip.loadAsync(masterBytes,{checkCRC32:true});
  const styleFile=zip.file("xl/styles.xml");
  if(!styleFile)throw new Error("Original master has no Excel formatting table");
  const {xml:updatedStyles,styleIds}=recolorStyles(await styleFile.async("string"));
  // Do not touch any other workbook part besides styles.xml and event bay cells.
  const modified=new Map();
  let count=0;
  for(let day=0;day<7;day++){
    const path="xl/worksheets/sheet"+(day+1)+".xml";
    const original=zip.file(path);
    if(!original)throw new Error("The master is missing the "+DAYS[day]+" worksheet");
    const events=weekEvents.filter(e=>e.date===addDays(state.week,day));
    const {xml,count:dayCount}=plotSheet(await original.async("string"),events,layoutDay(events),styleIds,period,displayLines);
    modified.set(path,xml);
    count+=dayCount;
  }
  zip.file("xl/styles.xml",updatedStyles);
  for(const [path,xml] of modified)zip.file(path,xml);
  const result=await zip.generateAsync({type:"blob",compression:"DEFLATE",compressionOptions:{level:6}});
  const filename=formatDownloadDate(state.week)+" - "+formatDownloadDate(addDays(state.week,6))+" Floorplan.xlsx";
  await downloadBinary(result,filename);
  return {filename,events:count};
}
