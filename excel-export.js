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
  // Append only the colors/fonts we need; keep every original master style intact.
  const colorToFill=new Map();
  xml=replaceSection(xml,"fills",section=>{
    const fills=[...section.matchAll(/<fill(?:\s[^>]*)?>[\s\S]*?<\/fill>/g)];
    let count=fills.length;
    fills.forEach((hit,index)=>{
      const rgb=(hit[0].match(/<fgColor\b[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/)||[])[1];
      if(rgb)colorToFill.set(rgb.toUpperCase(),index);
    });
    let appended="";
    for(const color of Object.values(COLORS).flat()){
      if(colorToFill.has(color))continue;
      colorToFill.set(color,count++);
      appended+='<fill><patternFill patternType="solid"><fgColor rgb="FF'+color+'"/><bgColor indexed="64"/></patternFill></fill>';
    }
    return section.replace(/<fills\b([^>]*)>/,(_,attrs)=>
      '<fills'+attrs.replace(/\bcount="\d+"/,'count="'+count+'"')+'>'
    ).replace("</fills>",appended+"</fills>");
  });

  const fontIds={};
  xml=replaceSection(xml,"fonts",section=>{
    const count=Number((section.match(/<fonts\b[^>]*count="(\d+)"/)||[])[1]);
    if(!Number.isInteger(count))throw new Error("Invalid Excel master font table");
    let appended="",next=count;
    for(const compact of [false,true])for(const white of [false,true]){
      const key=(compact?"compact":"normal")+(white?"White":"Black");
      fontIds[key]=next++;
      appended+='<font><b/><sz val="'+(compact?18:20)+'"/><color rgb="'+(white?"FFFFFFFF":"FF000000")+'"/><name val="Calibri"/><family val="2"/></font>';
    }
    return section.replace(/<fonts\b([^>]*)>/,(_,attrs)=>
      '<fonts'+attrs.replace(/\bcount="\d+"/,'count="'+next+'"')+'>'
    ).replace("</fonts>",appended+"</fonts>");
  });

  const xfsMatch=xml.match(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/);
  if(!xfsMatch)throw new Error("Invalid Excel master cell style table");
  const originals=[...xfsMatch[0].matchAll(/<xf(?:\s[^>]*)?(?:\/>|>[\s\S]*?<\/xf>)/g)].map(x=>x[0]);
  const created=[],cached=new Map();

  function setAttribute(tag,key,value){
    const attr=new RegExp("\\b"+key+'="[^"]*"');
    return attr.test(tag)?tag.replace(attr,key+'="'+value+'"'):tag.replace(/\/?>$/,match=>' '+key+'="'+value+'"'+match);
  }
  function cloneBaseXf(baseIndex,fillId,fontId){
    // Start with the EXACT source cell's XF, so its borderId (including thick
    // left/right outlines and AM/PM top/bottom rules) is never discarded.
    let original=originals[baseIndex];
    if(!original)throw new Error("Unknown style "+baseIndex+" in master");
    const tag=original.match(/^<xf\b[^>]*\/?>/);
    if(!tag)throw new Error("Invalid master style "+baseIndex);
    let next=tag[0];
    next=setAttribute(next,"fillId",fillId);
    next=setAttribute(next,"fontId",fontId);
    next=setAttribute(next,"applyFill",1);
    next=setAttribute(next,"applyFont",1);
    next=setAttribute(next,"applyAlignment",1);
    original=original.replace(tag[0],next);
    const alignment='<alignment horizontal="center" vertical="center" wrapText="1"/>';
    if(next.endsWith("/>")){
      return original.replace(/\/>$/,'>'+alignment+'</xf>');
    }
    if(/<alignment\b[^>]*\/>/.test(original)){
      return original.replace(/<alignment\b[^>]*\/>/,alignment);
    }
    return original.replace("</xf>",alignment+"</xf>");
  }

  const styleIds={
    resolve(baseIndex,event,compact){
      const category=COLORS[event.category]?event.category:"Flex";
      const tentative=event.status==="Tentative";
      const color=COLORS[category][tentative?1:0];
      const white=tentative&&category!=="Pleaser";
      const fontId=fontIds[(compact?"compact":"normal")+(white?"White":"Black")];
      const fillId=colorToFill.get(color);
      const key=[baseIndex,fillId,fontId].join(":");
      if(!cached.has(key)){
        cached.set(key,originals.length+created.length);
        created.push(cloneBaseXf(baseIndex,fillId,fontId));
      }
      return cached.get(key);
    }
  };

  function finalize(stylesXml){
    if(!created.length)return stylesXml;
    return replaceSection(stylesXml,"cellXfs",section=>{
      const count=originals.length+created.length;
      return section.replace(/<cellXfs\b([^>]*)>/,(_,attrs)=>
        '<cellXfs'+attrs.replace(/\bcount="\d+"/,'count="'+count+'"')+'>'
      ).replace("</cellXfs>",created.join("")+"</cellXfs>");
    });
  }
  return {xml,styleIds,finalize};
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
    const compact=rec.height===2;
    const first=colName(c1)+r1,last=colName(c2)+r2;
    for(let row=r1;row<=r2;row++)for(let col=c1;col<=c2;col++){
      const address=colName(col)+row;
      if(assignments.has(address))throw new Error("Events overlap in the export at "+address);
      assignments.set(address,{
        event,compact,
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
    const originalStyle=Number((original.match(/\bs="(\d+)"/)||[])[1]||0);
    const s=String(styleIds.resolve(originalStyle,change.event,change.compact));
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
  const {xml:updatedStyles,styleIds,finalize}=recolorStyles(await styleFile.async("string"));
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
  zip.file("xl/styles.xml",finalize(updatedStyles));
  for(const [path,xml] of modified)zip.file(path,xml);
  const result=await zip.generateAsync({type:"blob",compression:"DEFLATE",compressionOptions:{level:6}});
  const filename=formatDownloadDate(state.week)+" - "+formatDownloadDate(addDays(state.week,6))+" Floorplan.xlsx";
  await downloadBinary(result,filename);
  return {filename,events:count};
}
