import {reconcileDay,workbookDayPaths,parseFloorplanShared} from "./floorplan-sync.js?v=20261019c";
import {applyEventHyperlinks} from "./sheet-hyperlinks.js?v=20261017a";
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
    for(const color of [...Object.values(COLORS).flat(),"FFFFFF"]){
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
  const originals=[...xfsMatch[0].matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)].map(x=>x[0]);
  const created=[],cached=new Map();
  const bordersMatch=xml.match(/<borders\b[^>]*>[\s\S]*?<\/borders>/);
  if(!bordersMatch)throw new Error("Invalid Excel master border table");
  const originalBorders=[...bordersMatch[0].matchAll(/<border\b[^>]*\/>|<border\b[^>]*>[\s\S]*?<\/border>/g)].map(m=>m[0]);
  const addedBorders=[],borderCache=new Map();

  function setAttribute(tag,key,value){
    const attr=new RegExp("\\b"+key+'="[^"]*"');
    return attr.test(tag)?tag.replace(attr,key+'="'+value+'"'):tag.replace(/\/?>$/,match=>' '+key+'="'+value+'"'+match);
  }
  function addBorder(sides){
    if(!sides)return null;
    const names=["left","right","top","bottom"];
    const key=names.map(name=>sides[name]||"").join("|");
    if(borderCache.has(key))return borderCache.get(key);
    const xmlBorder='<border>'+names.map(name=>{
      const style=sides[name];
      const lineStyle=style==="light"?"thin":style;
      const lineColor=style==="light"?"FFD9D9D9":"FF000000";
      return style?'<'+name+' style="'+lineStyle+'"><color rgb="'+lineColor+'"/></'+name+'>':'<'+name+'/>';
    }).join("")+'<diagonal/></border>';
    const id=originalBorders.length+addedBorders.length;
    addedBorders.push(xmlBorder);
    borderCache.set(key,id);
    return id;
  }
  function centeredAlignment(original,vertical="center"){
    const alignment=original.match(/<alignment\b[^>]*\/>/);
    if(alignment){
      // Preserve rotation, indentation, etc. of pre-existing MOD labels.
      let next=alignment[0];
      next=setAttribute(next,"horizontal","center");
      next=setAttribute(next,"vertical",vertical);
      next=setAttribute(next,"wrapText",1);
      return original.replace(alignment[0],next);
    }
    const alignmentTag='<alignment horizontal="center" vertical="'+vertical+'" wrapText="1"/>';
    const opening=original.match(/^<xf\b[^>]*\/>/);
    if(opening)return original.replace(/\/>$/,'>'+alignmentTag+'</xf>');
    return original.replace("</xf>",alignmentTag+"</xf>");
  }
  function cloneBaseXf(baseIndex,fillId,fontId,borderId,center,vertical="center"){
    let original=originals[baseIndex];
    if(!original)throw new Error("Unknown style "+baseIndex+" in master");
    const opening=original.match(/^<xf\b[^>]*\/?>/);
    if(!opening)throw new Error("Invalid master style "+baseIndex);
    let next=opening[0];
    if(fillId!==null){next=setAttribute(next,"fillId",fillId);next=setAttribute(next,"applyFill",1)}
    if(fontId!==null){next=setAttribute(next,"fontId",fontId);next=setAttribute(next,"applyFont",1)}
    if(borderId!==null){next=setAttribute(next,"borderId",borderId);next=setAttribute(next,"applyBorder",1)}
    if(center)next=setAttribute(next,"applyAlignment",1);
    original=original.replace(opening[0],next);
    return center?centeredAlignment(original,vertical):original;
  }
  function createStyle(baseIndex,fillId,fontId,sides,center,vertical="center"){
    const borderId=addBorder(sides);
    const key=[baseIndex,fillId,fontId,borderId,center?1:0,vertical].join(":");
    if(!cached.has(key)){
      cached.set(key,originals.length+created.length);
      created.push(cloneBaseXf(baseIndex,fillId,fontId,borderId,center,vertical));
    }
    return cached.get(key);
  }
  function sectionBorderEdge(styleId,side){
    const xf=originals[styleId]||originals[0]||"";
    const bid=Number((xf.match(/\bborderId="(\d+)"/)||[])[1]||0);
    const border=originalBorders[bid]||"";
    const tag=(border.match(new RegExp("<"+side+"\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/"+side+">)"))||[])[0]||"";
    const weight=(tag.match(/\bstyle="([^"]+)"/)||[])[1]||"";
    if(!weight)return false;
    if(/^(medium|thick|double|mediumDashed)$/.test(weight))return true;
    if(weight!=="thin")return false;
    const hex=(tag.match(/<color\b[^>]*\brgb="([0-9A-Fa-f]{6,8})"/)||[])[1];
    if(!hex)return true; // Excel's normal grid lines are not actual styled borders.
    const rgb=hex.slice(-6);
    const channels=[0,2,4].map(i=>parseInt(rgb.slice(i,i+2),16));
    return Math.max(...channels)<125;
  }
  const styleIds={
    hardDivider(leftId,rightId,direction){
      return direction==="vertical"?
        sectionBorderEdge(leftId,"right")||sectionBorderEdge(rightId,"left"):
        sectionBorderEdge(leftId,"bottom")||sectionBorderEdge(rightId,"top");
    },
    resolve(baseIndex,event,compact,sides){
      const category=COLORS[event.category]?event.category:"Flex";
      const tentative=event.status==="Tentative";
      const color=COLORS[category][tentative?1:0];
      const fontId=fontIds[(compact?"compact":"normal")+"Black"];
      return createStyle(baseIndex,colorToFill.get(color),fontId,sides,true);
    },
    blank(baseIndex,sides){
      return createStyle(baseIndex,null,null,sides,true);
    },
    center(baseIndex){
      return createStyle(baseIndex,null,null,null,true);
    },
    staff(baseIndex){
      // Center staff and events both horizontally and vertically while
      // retaining the master cell's font, color, borders and rich text.
      return createStyle(baseIndex,null,null,null,true,"center");
    },
    clear(baseIndex){
      return createStyle(baseIndex,colorToFill.get("FFFFFF"),null,null,true);
    },
    event(baseIndex,event,compact=false){
      const category=COLORS[event.category]?event.category:"Flex";
      const rgb=COLORS[category][event.status==="Tentative"?1:0];
      return createStyle(baseIndex,colorToFill.get(rgb),fontIds[(compact?"compact":"normal")+"Black"],null,true);
    }
  };

  function finalize(stylesXml){
    if(!created.length)return stylesXml;
    stylesXml=replaceSection(stylesXml,"borders",section=>{
      if(!addedBorders.length)return section;
      return section.replace(/<borders\b([^>]*)>/,(_,attrs)=>
        '<borders'+attrs.replace(/\bcount="\d+"/,'count="'+(originalBorders.length+addedBorders.length)+'"')+'>'
      ).replace("</borders>",addedBorders.join("")+"</borders>");
    });
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

function colNumber(letters){
  let number=0;for(const ch of letters)number=number*26+ch.charCodeAt(0)-64;return number;
}
function location(ref){
  const a=ref.match(/^([A-Z]+)(\d+)$/);
  if(!a)throw new Error("Invalid merged cell: "+ref);
  return {c:colNumber(a[1]),r:Number(a[2])};
}
function rect(ref){
  const [start,end=start]=ref.split(":");
  const a=location(start),b=location(end);
  return {c1:a.c,r1:a.r,c2:b.c,r2:b.r};
}
function reference(x){return colName(x.c1)+x.r1+":"+colName(x.c2)+x.r2}
function intersection(a,b){
  const c1=Math.max(a.c1,b.c1),c2=Math.min(a.c2,b.c2),
    r1=Math.max(a.r1,b.r1),r2=Math.min(a.r2,b.r2);
  return c1<=c2&&r1<=r2?{c1,r1,c2,r2}:null;
}
function subtractRectangle(original,cut){
  const i=intersection(original,cut);
  if(!i)return [original];
  const remaining=[];
  // Horizontal sections above/below the new event, and vertical sections
  // to its left/right. Mirrors the hand-built "Test Event" workbook exactly.
  if(original.r1<i.r1)remaining.push({c1:original.c1,c2:original.c2,r1:original.r1,r2:i.r1-1});
  if(i.r2<original.r2)remaining.push({c1:original.c1,c2:original.c2,r1:i.r2+1,r2:original.r2});
  if(original.c1<i.c1)remaining.push({c1:original.c1,c2:i.c1-1,r1:i.r1,r2:i.r2});
  if(i.c2<original.c2)remaining.push({c1:i.c2+1,c2:original.c2,r1:i.r1,r2:i.r2});
  return remaining;
}
function isMergedRectangle(x){return x.c1!==x.c2||x.r1!==x.r2}
function readCells(xml){
  const found=new Map();
  const re=/<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*?\/>|<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*>[\s\S]*?<\/c>/g;
  for(const m of xml.matchAll(re)){
    const ref=m[1]||m[2],content=m[0],style=Number((content.match(/\bs="(\d+)"/)||[])[1]||0);
    const hasData=/<(?:v|f|is)\b/.test(content);
    found.set(ref,{style,content,hasData});
  }
  return found;
}
function insertMissingCells(xml,missing){
  if(!missing.size)return xml;
  const rows=new Map();
  for(const [ref,cell] of missing){
    const spot=location(ref);
    if(!rows.has(spot.r))rows.set(spot.r,[]);
    rows.get(spot.r).push({column:spot.c,text:cell});
  }
  const processed=new Set();
  xml=xml.replace(/(<row\b[^>]*\br="(\d+)"[^>]*>)([\s\S]*?)(<\/row>)/g,(whole,opening,rowText,body,closing)=>{
    const row=Number(rowText),additions=rows.get(row);
    if(!additions)return whole;
    processed.add(row);
    additions.sort((a,b)=>a.column-b.column);
    let updated=body;
    for(const addition of additions){
      const existing=[...updated.matchAll(/<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*\/?>/g)];
      const next=existing.find(e=>location(e[1]).c>addition.column);
      const at=next?next.index:updated.length;
      updated=updated.slice(0,at)+addition.text+updated.slice(at);
    }
    return opening+updated+closing;
  });
  // A blank row may be omitted entirely by Excel. Create only that row,
  // keeping every pre-existing row, height and grid style byte-for-byte.
  for(const [row,additions] of rows){
    if(processed.has(row))continue;
    additions.sort((a,b)=>a.column-b.column);
    const markup='<row r="'+row+'">'+additions.map(item=>item.text).join("")+'</row>';
    const next=[...xml.matchAll(/<row\\b[^>]*\\br="(\\d+)"[^>]*>/g)].find(m=>Number(m[1])>row);
    if(next)xml=xml.slice(0,next.index)+markup+xml.slice(next.index);
    else xml=xml.replace('</sheetData>',markup+'</sheetData>');
  }
  return xml;
}
function sectionBorder(area,x,y,originalSection=null,isEvent=false){
  // The original merged section owns its outline, not each new fragment.
  // Cutting a blank merge around an event must NOT create black internal lines.
  // Event edges are pale gray, like the spreadsheet's ordinary gridlines.
  const floorTops=[4,9,20,25,36,41],floorBottoms=[8,13,24,29,40,45];
  const sides={left:null,right:null,top:null,bottom:null};
  const external={
    left:x===5,right:x===38,
    top:floorTops.includes(y),bottom:floorBottoms.includes(y)
  };
  const eventEdge={
    left:x===area.c1,right:x===area.c2,
    top:y===area.r1,bottom:y===area.r2
  };
  for(const side of ["left","right","top","bottom"]){
    if(external[side]){sides[side]="medium";continue}
    const originalEdge=originalSection&&({
      left:x===originalSection.c1,right:x===originalSection.c2,
      top:y===originalSection.r1,bottom:y===originalSection.r2
    })[side];
    // Preserve existing section dividers, but never run them through the
    // middle of a colored event that crosses an old section boundary.
    if(originalEdge&&(!isEvent||eventEdge[side])){sides[side]="thin";continue}
    if(isEvent&&eventEdge[side])sides[side]="light";
  }
  return sides;
}
function centerExistingText(xml,styleIds){
  // Keep pre-entered text (including rotated MOD names), but center its alignment.
  // Do not alter the original rotation, value, formula, merged range or border.
  return xml.replace(/<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*?\/>|<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*>[\s\S]*?<\/c>/g,(original,emptyRef,filledRef)=>{
    const ref=emptyRef||filledRef;
    if(!/\bt="(?:s|inlineStr|str)"/.test(original)||!/<(?:is|v)\b/.test(original))return original;
    const pos=location(ref);
    if(pos.r>47||pos.c>39)return original;
    const base=Number((original.match(/\bs="(\d+)"/)||[])[1]||0);
    const center=styleIds.center(base);
    return original.replace(/(<c\b[^>]*?)\bs="\d+"/,'$1s="'+center+'"');
  });
}
function plotSheet(xml,events,layout,styleIds,period,displayLines){
  const records=Object.values(layout.records).flat().filter(rec=>rec.row>=0);
  xml=centerExistingText(xml,styleIds);
  if(!records.length)return {xml,count:0,skipped:[]};
  const originalCells=readCells(xml);
  const gridStyles=[...originalCells].map(([ref,cell])=>({ref,cell,...location(ref)})).filter(x=>x.c>=5&&x.c<=38&&x.r>=4&&x.r<=45);
  const prefilled=gridStyles.filter(x=>x.cell.hasData);
  const mergeBlock=xml.match(/<mergeCells\b[^>]*>[\s\S]*?<\/mergeCells>/);
  if(!mergeBlock)throw new Error("The Excel master is missing its merged-cell layout");
  const originalMerges=[...mergeBlock[0].matchAll(/<mergeCell\b[^>]*ref="([^"]+)"[^>]*\/>/g)].map(x=>({ref:x[1],area:rect(x[1])}));
  const occupied=new Set(),placements=[],skipped=[];
  const eventCells=new Map();
  for(const rec of records){
    const event=rec.event;
    const floorBase={3:4,2:20,1:36}[rec.range.floor];
    if(!floorBase)continue;
    const r1=floorBase+(period(event)==="PM"?5:0)+rec.row;
    const box={c1:4+rec.range.start%100,c2:4+rec.range.end%100,r1,r2:r1+rec.height-1};
    if(box.c1<5||box.c2>38)throw new Error("Event bays are outside the floor plan");
    const conflicts=originalMerges.filter(m=>intersection(m.area,box))
      .filter(m=>prefilled.some(p=>p.c>=m.area.c1&&p.c<=m.area.c2&&p.r>=m.area.r1&&p.r<=m.area.r2));
    // Don't silently replace the text in an already populated section. Those
    // blocks were manually placed in the uploaded weekly master.
    let hasContent=conflicts.length>0;
    if(!hasContent){
      for(let y=box.r1;y<=box.r2&&!hasContent;y++)for(let x=box.c1;x<=box.c2;x++){
        if(originalCells.get(colName(x)+y)?.hasData){hasContent=true;break}
      }
    }
    if(hasContent){
      skipped.push(event.name+" ("+reference(box)+") — prefilled master section");
      continue;
    }
    for(let y=box.r1;y<=box.r2;y++)for(let x=box.c1;x<=box.c2;x++){
      const address=colName(x)+y;
      if(occupied.has(address))throw new Error("Two imported events overlap at "+address);
      occupied.add(address);
      eventCells.set(address,{event,compact:rec.height===2,
        area:box,originalSection:null,text:y===box.r1&&x===box.c1?displayLines(event).join("\n"):null});
    }
    placements.push(box);
  }
  if(!placements.length)return {xml,count:0,skipped};
  const newMerges=[];
  const leftoverAreas=[];
  const leftoverOrigins=new Map();
  for(const original of originalMerges){
    let parts=[original.area];
    for(const placed of placements)parts=parts.flatMap(part=>subtractRectangle(part,placed));
    const changed=parts.length!==1||!parts.length||reference(parts[0])!==reference(original.area);
    if(changed){
      leftoverAreas.push(...parts);
      for(const part of parts)leftoverOrigins.set(reference(part),original.area);
    }
    for(const p of parts)if(isMergedRectangle(p))newMerges.push(reference(p));
  }
  for(const placed of placements)if(isMergedRectangle(placed))newMerges.push(reference(placed));
  // The untouched merged regions are retained; modified blanks are re-merged
  // around each inserted event. Never leave overlapping mergeCell entries.
  const scan=new Set();
  for(const m of newMerges){
    if(scan.has(m))throw new Error("Duplicate Excel merge "+m);
    scan.add(m);
  }
  // Blank remnants of a previously merged bay section must be realigned and
  // outlined too. Otherwise the old individual cell-grid borders remain visible.
  const sectionCells=new Map();
  for(const area of leftoverAreas){
    const relevant=area.c1>=5&&area.c2<=38&&
      [4,9,20,25,36,41].some(top=>area.r1>=top&&area.r2<=top+4);
    if(!relevant)continue;
    for(let row=area.r1;row<=area.r2;row++)for(let col=area.c1;col<=area.c2;col++){
      const address=colName(col)+row;
      if(!eventCells.has(address))sectionCells.set(address,{area,originalSection:leftoverOrigins.get(reference(area))||null,event:null,compact:false,text:null});
    }
  }
  const changes=new Map([...sectionCells,...eventCells]);
  // Find the nearest existing template style, preferably on the same edge of
  // the old merged region. This handles absent <c> nodes in premerged blocks.
  function baseStyle(address){
    const original=originalCells.get(address);
    if(original)return original.style;
    const loc=location(address);
    const containing=originalMerges.find(m=>intersection(m.area,{c1:loc.c,c2:loc.c,r1:loc.r,r2:loc.r}));
    if(!containing)return 0;
    const edgeCandidates=[];
    for(const {c,r,cell} of gridStyles){
      const spot={c,r};
      if(spot.c<containing.area.c1||spot.c>containing.area.c2||
        spot.r<containing.area.r1||spot.r>containing.area.r2)continue;
      const edgeSame=(loc.c===containing.area.c1&&spot.c===loc.c)||
        (loc.c===containing.area.c2&&spot.c===loc.c)||
        (loc.r===containing.area.r1&&spot.r===loc.r)||
        (loc.r===containing.area.r2&&spot.r===loc.r);
      const distance=Math.abs(loc.c-spot.c)+Math.abs(loc.r-spot.r);
      edgeCandidates.push({distance:distance+(edgeSame?0:100),style:cell.style});
    }
    edgeCandidates.sort((a,b)=>a.distance-b.distance);
    return edgeCandidates[0]?.style||originalCells.get(colName(containing.area.c1)+containing.area.r1)?.style||0;
  }
  const missing=new Map(),rewritten=new Set();
  function cellXml(address,change){
    const loc=location(address);
    const source=change.originalSection||(change.event?originalMerges.find(m=>intersection(m.area,{c1:loc.c,c2:loc.c,r1:loc.r,r2:loc.r}))?.area:null);
    const sides=sectionBorder(change.area,loc.c,loc.r,source,Boolean(change.event));
    const base=baseStyle(address);
    const id=change.event?styleIds.resolve(base,change.event,change.compact,sides):styleIds.blank(base,sides);
    const style=String(id);
    if(change.text===null)return '<c r="'+address+'" s="'+style+'"/>';
    return '<c r="'+address+'" s="'+style+'" t="inlineStr"><is><t xml:space="preserve">'+xmlEscape(change.text)+'</t></is></c>';
  }
  const cellRe=/<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*?\/>|<c\b(?=[^>]*\br="([A-Z]{1,3}\d+)")[^>]*>[\s\S]*?<\/c>/g;
  xml=xml.replace(cellRe,(whole,short,long)=>{
    const ref=short||long,change=changes.get(ref);
    if(!change)return whole;
    rewritten.add(ref);
    return cellXml(ref,change);
  });
  for(const [ref,change] of changes)if(!rewritten.has(ref))missing.set(ref,cellXml(ref,change));
  // Rebuild a blank anchor when the original merged section's anchor was
  // covered by an event and the adjacent remainder needs its own merge anchor.
  for(const area of leftoverAreas){
    const anchor=colName(area.c1)+area.r1;
    if(!originalCells.has(anchor)&&!missing.has(anchor)&&!eventCells.has(anchor))
      missing.set(anchor,'<c r="'+anchor+'" s="'+baseStyle(anchor)+'"/>');
  }
  xml=insertMissingCells(xml,missing);
  xml=replaceSection(xml,"mergeCells",section=>{
    const count=newMerges.length;
    return section.replace(/<mergeCells\b([^>]*)>/,(_,attrs)=>
      '<mergeCells'+attrs.replace(/\bcount="\d+"/,'count="'+count+'"')+'>'
    ).replace(/<mergeCell\b[^>]*\/>/g,"")
      .replace("</mergeCells>",newMerges.map(ref=>'<mergeCell ref="'+ref+'"/>').join("")+"</mergeCells>");
  });
  return {xml,count:placements.length,skipped};
}

/**
 * Lock the master workbook's visual row geometry.
 *
 * Google Sheets' XLSX export often writes the user's visibly taller
 * floor-divider/header rows as implicit 15 pt defaults. Importing that XLSX
 * again collapses the 6-row bay-header bands (14–19, 30–35).
 * Match the photographed 21 pt original appearance for those header rows
 * while honoring any custom source row heights. Other rows retain the
 * original effective height. Cells, colors, merges, column widths and
 * explicitly sized rows are never resized.
 */
function lockMasterRowHeights(outputXml,masterXml){
  const defaults=(masterXml.match(/<sheetFormatPr\b[^>]*\/?>/)||[])[0]||"";
  const match=defaults.match(/\bdefaultRowHeight="([0-9]+(?:\.[0-9]+)?)"/);
  if(!match)return outputXml; // Don't guess a height absent from this master.
  const defaultHeight=match[1];
  const masterRows=new Map(
    [...masterXml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>/g)].map(m=>[m[1],m[0]])
  );
  // Excel templates exported from Google Sheets commonly serialize the
  // six-row floor-transition/header bands as bare 15 pt defaults, even when
  // the original Google Sheets floorplan visually uses ~21 pt rows. The
  // screenshots show this exact 6*4 px shortfall after reimport at 50% zoom.
  // Restore the source floorplan's visible section spacing only when a row
  // had NO explicit height. Respect all explicitly sized uploaded masters.
  const isSectionBand=(n)=>(n>=14&&n<=19)||(n>=30&&n<=35);
  const sectionHeight="21";
  const fixedRowTag=(tag,n)=>{
    let row=tag;
    const explicit=(row.match(/\bht="([0-9]+(?:\.[0-9]+)?)"/)||[])[1];
    const height=explicit||(isSectionBand(n)?sectionHeight:defaultHeight);
    if(!/\bht="[^"]*"/.test(row))row=row.replace(/>$/,' ht="'+height+'">');
    if(/\bcustomHeight="[^"]*"/.test(row))
      row=row.replace(/\bcustomHeight="[^"]*"/,'customHeight="1"');
    else row=row.replace(/>$/,' customHeight="1">');
    return row;
  };
  outputXml=outputXml.replace(/<row\b[^>]*\br="(\d+)"[^>]*>/g,(row,n)=>
    fixedRowTag(masterRows.get(n)||row,Number(n)));
  // The exporter must never change column geometry. Defensive restoration
  // makes custom-uploaded masters authoritative if another stage did so.
  const masterCols=(masterXml.match(/<cols\b[^>]*>[\s\S]*?<\/cols>/)||[])[0];
  if(masterCols)outputXml=outputXml.replace(/<cols\b[^>]*>[\s\S]*?<\/cols>/,masterCols);
  return outputXml;
}

function renameWeeklyTabs(xml,weekStart,addDays){
  // Change names only, never worksheet identities or ordering.
  return xml.replace(/<sheet\b[^>]*\/>/g,(tag)=>{
    const name=(tag.match(/\bname="([^"]+)"/)||[])[1];
    const day=DAYS.findIndex(d=>name===d||new RegExp("^"+d+" \\d{1,2}\\.\\d{1,2}$").test(name));
    if(day<0)return tag;
    const date=addDays(weekStart,day);
    const parts=date.split("-");
    const label=DAYS[day]+" "+Number(parts[1])+"."+parts[2];
    return tag.replace(/\bname="[^"]+"/,'name="'+xmlEscape(label)+'"');
  });
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
export async function buildExactFloorPlan({state,weekEvents,layoutDay,period,displayLines,addDays}){
  if(typeof JSZip!=="function")throw new Error("Local ZIP export library failed to load");
  const masterBytes=state.master||(await (async()=>{
    const response=await fetch("./master.xlsx",{cache:"no-cache"});
    if(!response.ok)throw new Error("Original floor plan master could not be loaded");
    return await response.arrayBuffer();
  })());
  const zip=await JSZip.loadAsync(masterBytes);
  const styleFile=zip.file("xl/styles.xml");
  if(!styleFile)throw new Error("Original master has no Excel formatting table");
  const {xml:updatedStyles,styleIds,finalize}=recolorStyles(await styleFile.async("string"));
  const workbookFile=zip.file("xl/workbook.xml");
  if(!workbookFile)throw new Error("Excel master is missing the worksheet names");
  const workbookXml=await workbookFile.async("string");
  const dayPaths=workbookDayPaths(workbookXml);
  const sharedFile=zip.file("xl/sharedStrings.xml");
  const shared=sharedFile?parseFloorplanShared(await sharedFile.async("string")):[];
  const modified=new Map(),changesByDay=[],skipped=[];
  let count=0;
  for(let day=0;day<7;day++){
    const path=dayPaths[day];
    if(!path)throw new Error("Missing "+DAYS[day]+" worksheet");
    const original=zip.file(path);
    if(!original)throw new Error("Could not find "+DAYS[day]+" worksheet data");
    const events=weekEvents.filter(e=>e.date===addDays(state.week,day));
    const masterSheetXml=await original.async("string");
    if(state.master){
      const result=reconcileDay(masterSheetXml,shared,events,styleIds,displayLines);
      modified.set(path,lockMasterRowHeights(result.xml,masterSheetXml));
      count+=result.placed;
      changesByDay[day]=result.changes;
      skipped.push(...result.warnings.map(w=>DAYS[day]+": "+w));
    }else{
      const result=plotSheet(masterSheetXml,events,layoutDay(events),styleIds,period,displayLines);
      modified.set(path,lockMasterRowHeights(result.xml,masterSheetXml));
      count+=result.count;
      changesByDay[day]=[];
      skipped.push(...result.skipped.map(w=>DAYS[day]+": "+w));
    }
  }
  // Add links to worksheet event anchors without changing a single cell
  // style, font, border, merge, color or visible text.
  const hyperlinked=await applyEventHyperlinks(zip,modified,dayPaths,shared,weekEvents,state.salesforceLinks||[],state.week,addDays);
  zip.file("xl/workbook.xml",renameWeeklyTabs(workbookXml,state.week,addDays));
  zip.file("xl/styles.xml",finalize(updatedStyles));
  for(const [path,xml] of modified)zip.file(path,xml);
  const result=await zip.generateAsync({type:"blob",compression:"DEFLATE",compressionOptions:{level:6}});
  const filename=formatDownloadDate(state.week)+" - "+formatDownloadDate(addDays(state.week,6))+" Floorplan.xlsx";
  return {blob:result,filename,events:count,skipped,changesByDay,hyperlinked};
}
export async function exportExactFloorPlan(options){
  const result=options.cachedResult||await buildExactFloorPlan(options);
  await downloadBinary(result.blob,result.filename);
  return result;
}
