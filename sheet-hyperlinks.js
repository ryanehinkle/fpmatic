import {parseFloorplanCell} from "./floorplan-sync.js?v=20261016g";
import {linkForWorksheetText} from "./salesforce-links.js?v=20261016g";
const XLSX_REL="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const XML_REL="http://schemas.openxmlformats.org/package/2006/relationships";
function escapeXml(s){return String(s).replace(/[&<>"']/g,ch=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&apos;"}[ch]))}
function unusedRid(existing,extra=[]){
  const ids=new Set([...existing.matchAll(/\bId="([^"]+)"/g)].map(x=>x[1]).concat(extra));
  let n=1;while(ids.has("rId"+n))n++;
  return "rId"+n;
}
export function addSheetHyperlinks(xml,rels,items){
  if(!items?.length)return {xml,rels,added:0};
  let relationshipXml=rels||'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="'+XML_REL+'"></Relationships>';
  if(!/<Relationships\b/.test(relationshipXml))throw new Error("Unsupported Excel worksheet relations");
  const hyperlinkTags=[],relationships=[];
  const oldHyperlinks=xml.match(/<hyperlinks\b[^>]*>[\s\S]*?<\/hyperlinks>/)?.[0]||"";
  const cleaned=new Set(items.map(x=>x.ref));
  const existing=oldHyperlinks.replace(/<hyperlink\b[^>]*\/>/g,tag=>{
    const r=(tag.match(/\bref="([^"]+)"/)||[])[1];
    return cleaned.has(r)?"":tag;
  });
  const extras=[];
  for(const {ref,url} of items){
    if(!/^[A-Z]{1,3}\d+$/.test(ref))continue;
    const parsed=new URL(url);
    if(parsed.protocol!=="https:")continue;
    const id=unusedRid(relationshipXml,extras);
    extras.push(id);
    hyperlinkTags.push('<hyperlink ref="'+ref+'" r:id="'+id+'"/>');
    relationships.push('<Relationship Id="'+id+'" Type="'+XLSX_REL+'" Target="'+escapeXml(url)+'" TargetMode="External"/>');
  }
  if(!hyperlinkTags.length)return {xml,rels,added:0};
  relationshipXml=relationshipXml.replace(/<Relationships\b([^>]*)\/>/,
    (_,attrs)=>'<Relationships'+attrs+'></Relationships>');
  relationshipXml=relationshipXml.replace("</Relationships>",relationships.join("")+"</Relationships>");
  if(!/\bxmlns:r=/.test(xml.match(/<worksheet\b[^>]*>/)?.[0]||"")){
    xml=xml.replace(/<worksheet\b/, '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
  }
  if(oldHyperlinks){
    xml=xml.replace(oldHyperlinks,existing.replace("</hyperlinks>",hyperlinkTags.join("")+"</hyperlinks>"));
  }else{
    const block='<hyperlinks>'+hyperlinkTags.join("")+'</hyperlinks>';
    // SpreadsheetML hyperlink ordering: after dataValidations/mergeCells,
    // before page layout and drawing metadata.
    const before=xml.search(/<(?:printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/);
    if(before>=0)xml=xml.slice(0,before)+block+xml.slice(before);
    else xml=xml.replace("</worksheet>",block+"</worksheet>");
  }
  return {xml,rels:relationshipXml,added:hyperlinkTags.length};
}
function excelColumnNumber(s){return [...s].reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0)}
function anchoredBayRanges(xml){
  const out=new Map();
  for(const m of xml.matchAll(/<mergeCell\b[^>]*ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"[^>]*\/>/g)){
    const first=excelColumnNumber(m[1]),last=excelColumnNumber(m[3]),row=+m[2];
    if(first<5||last>38||first>last)continue;
    const floor=row>=36?1:row>=20?2:row>=4&&row<=13?3:0;
    if(!floor)continue;
    out.set(m[1]+m[2],(floor*100+first-4)+"-"+(floor*100+last-4));
  }
  return out;
}
export async function applyEventHyperlinks(zip,modified,dayPaths,shared,events,links,week,addDays){
  let added=0;
  for(let day=0;day<7;day++){
    const path=dayPaths[day],xml=modified.get(path);
    if(!path||!xml)continue;
    const date=addDays(week,day),cells=parseFloorplanCell(xml,shared),items=[],baysByAnchor=anchoredBayRanges(xml);
    for(const [ref,cell] of cells){
      if(!cell.text||!/^[A-Z]{1,3}\d+$/.test(ref))continue;
      const numeric=(ref.match(/\d+/)||[])[0],row=Number(numeric);
      // Event cells only, never a staffing or reference worksheet.
      if(![4,9,20,25,36,41].some(top=>row>=top&&row<=top+4))continue;
      const url=linkForWorksheetText(cell.text,date,events,links,baysByAnchor.get(ref)||"");
      if(url)items.push({ref,url});
    }
    if(!items.length)continue;
    const relationshipPath=path.replace("/worksheets/","/worksheets/_rels/")+".rels";
    const oldRels=zip.file(relationshipPath);
    const result=addSheetHyperlinks(xml,oldRels?await oldRels.async("string"):null,items);
    modified.set(path,result.xml);
    zip.file(relationshipPath,result.rels);
    added+=result.added;
  }
  return added;
}
