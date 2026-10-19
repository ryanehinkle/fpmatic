// Run with: node tests/floorplan-regression.mjs
// Verifies the real bundled master keeps thick borders and back-to-back
// event stacks are contiguous in the emitted XLSX.
import assert from "node:assert/strict";
import fs from "node:fs";
import { createRequire } from "node:module";
const require=createRequire(import.meta.url);
globalThis.JSZip=require("../jszip.min.js");

const source=fs.readFileSync(new URL("../app.js",import.meta.url),"utf8");
const core=source.slice(source.indexOf("function layoutDay(events){"),source.indexOf("\nfunction couch("));
assert(core.startsWith("function layoutDay"),"Could not locate app placement logic");

const clockMin=s=>{const [h,m]=s.split(":").map(Number);return h*60+m};
const period=e=>clockMin(e.start)>=960||clockMin(e.end)>=1140?"PM":"AM";
const baySegments=str=>[...str.matchAll(/(\d{3})(?:-(\d{3}))?/g)].map(x=>({start:Number(x[1]),end:Number(x[2]||x[1]),floor:Math.floor(Number(x[1])/100)}));
const overlaps=(a,b)=>a.start<=b.end&&b.start<=a.end;
const layoutDay=new Function("period","baySegments","clockMin","overlaps",core+"\nreturn layoutDay")(period,baySegments,clockMin,overlaps);
const monday="2026-10-12";
function event(id,bays,start,end,category="Buffet"){
  return {id,name:id,bays,start,end,date:monday,category,package:"Fajita",status:"Definite",guests:20,space:"",golfStart:"",golfEnd:""};
}
const data=[
  event("First event","301-304","11:00","12:30"),
  event("Second event","301-304","13:00","15:00"),
  event("Right-edge event","329-334","11:00","13:00"),
  event("Left second-floor event","201-203","10:00","13:00")
];
const plot=layoutDay(data);
const stack=plot.records["3-AM"].filter(x=>x.event.id==="First event"||x.event.id==="Second event");
assert.deepEqual(stack.map(x=>[x.row,x.height]),[[0,2],[2,2]],"Stacked 2-high events should occupy rows 0-1 and 2-3");
assert.equal(plot.overflow.length,0);

const original=fs.readFileSync(new URL("../master.xlsx",import.meta.url));
globalThis.fetch=async()=>({ok:true,arrayBuffer:async()=>original.buffer.slice(original.byteOffset,original.byteOffset+original.byteLength)});
const originalZip=await JSZip.loadAsync(original);
const sourceModule=fs.readFileSync(new URL("../excel-export.js",import.meta.url),"utf8");
const syncSource=fs.readFileSync(new URL("../floorplan-sync.js",import.meta.url),"utf8").replace(/^export /gm,"");
const exporter=await import("data:text/javascript;base64,"+Buffer.from(syncSource+"\nasync function applyEventHyperlinks(){return 0;}\n"+sourceModule.replace(/^import[^\n]*\n/gm,"")).toString("base64"));

let downloadedBlob,downloadedName;
const makeURL=URL.createObjectURL,freeURL=URL.revokeObjectURL;
URL.createObjectURL=blob=>{downloadedBlob=blob;return "blob:fpmatic-regression"};
URL.revokeObjectURL=()=>{};
globalThis.document={
  body:{appendChild(){}},
  createElement(){return {set href(x){},set download(x){downloadedName=x},click(){},remove(){}}}
};
try{
 await exporter.exportExactFloorPlan({
   state:{master:null,week:monday},weekEvents:data,layoutDay,
   period,displayLines:e=>[e.name,e.start+"-"+e.end,e.guests+"ppl",e.package],
   addDays:(date,n)=>{const d=new Date(date+"T12:00:00Z");d.setUTCDate(d.getUTCDate()+n);return d.toISOString().slice(0,10)}
 });
}finally{URL.createObjectURL=makeURL;URL.revokeObjectURL=freeURL}
assert(downloadedBlob,"No file was downloaded");
assert.match(downloadedName,/Floorplan\.xlsx$/);
const outputZip=await JSZip.loadAsync(await downloadedBlob.arrayBuffer());
function read(zip,path){return zip.file(path).async("string")}
const [oldSheet,newSheet,oldStyles,newStyles]=await Promise.all([
  read(originalZip,"xl/worksheets/sheet1.xml"),read(outputZip,"xl/worksheets/sheet1.xml"),
  read(originalZip,"xl/styles.xml"),read(outputZip,"xl/styles.xml")
]);
function cell(xml,ref){
 const self=new RegExp('<c\\b(?=[^>]*\\br="'+ref+'")[^>]*\\/>');
 const expanded=new RegExp('<c\\b(?=[^>]*\\br="'+ref+'")[^>]*>[\\s\\S]*?<\\/c>');
 const match=xml.match(self)||xml.match(expanded);
 assert(match,"Missing master cell "+ref);
 return match[0];
}
function styleIndex(c){return Number((c.match(/\bs="(\d+)"/)||[])[1]||0)}
function xf(xml,id){
 const inner=(xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)||[])[1];
 assert(inner,"Missing master XF table");
 const xfs=[...inner.matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)].map(x=>x[0]);
 assert(id<xfs.length,"Invalid XF reference "+id);
 return xfs[id];
}
function borderId(xml,cellRef){
 return Number((xf(xml,styleIndex(cell(xml,cellRef))).match(/\bborderId="(\d+)"/)||[])[1]||0)
}
function border(xml,cellRef,sheetXml){
 const id=Number((xf(xml,styleIndex(cell(sheetXml,cellRef))).match(/\bborderId="(\d+)"/)||[])[1]||0);
 const content=(xml.match(/<borders\b[^>]*>([\s\S]*?)<\/borders>/)||[])[1];
 assert(content,"Missing borders");
 const borders=[...content.matchAll(/<border\b[^>]*\/>|<border\b[^>]*>[\s\S]*?<\/border>/g)].map(x=>x[0]);
 assert(borders[id],"Unknown border "+id);
 return borders[id];
}
function side(borderXml,name){
 return (borderXml.match(new RegExp('<'+name+'\\b[^>]*style="([^"]+)"'))||[])[1]||"";
}
// Only the perimeter of each event is drawn; outer-floor lines remain heavier.
assert.equal(side(border(newStyles,"E4",newSheet),"left"),"medium");
assert.equal(side(border(newStyles,"E4",newSheet),"top"),"medium");
assert.equal(side(border(newStyles,"E5",newSheet),"bottom"),"thin");
assert.equal(side(border(newStyles,"F4",newSheet),"left"),"");
assert.equal(side(border(newStyles,"F5",newSheet),"top"),"");
assert.equal(side(border(newStyles,"F5",newSheet),"right"),"");
assert.equal(side(border(newStyles,"H4",newSheet),"right"),"thin");
assert.equal(side(border(newStyles,"AL4",newSheet),"right"),"medium");
for(const ref of ["E8","AL8","E9","AM4","AM9"]){
 assert.equal(cell(newSheet,ref),cell(oldSheet,ref),"Untouched master grid cell changed at "+ref);
}
// Existing cell content is never changed when its alignment is centered.
for(const ref of ["D4","D9"]){
 const before=cell(oldSheet,ref).replace(/\bs="\d+"/,'');
 const after=cell(newSheet,ref).replace(/\bs="\d+"/,'');
 assert.equal(after,before,"Template label content changed at "+ref);
}

// Compare ORIGINAL master sizes against every generated day. The master
// usually omits ht= for bay-number/landmark rows; these must be explicitly
// locked to that master default so Sheets cannot auto-fit them on import.
function sourceRowSizes(xml){
 const defaultHeight=Number((xml.match(/<sheetFormatPr\b[^>]*defaultRowHeight="([^"]+)"/)||[])[1]);
 assert(Number.isFinite(defaultHeight)&&defaultHeight>0,"Master needs a valid default row height");
 const rows=new Map([...xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>/g)].map(m=>{
   const explicit=(m[0].match(/\bht="([^"]+)"/)||[])[1];
   return [+m[1],explicit?Number(explicit):defaultHeight];
 }));
 return {rows,defaultHeight};
}
for(let day=1;day<=7;day++){
 const path="xl/worksheets/sheet"+day+".xml";
 const source=await read(originalZip,path),exported=await read(outputZip,path);
 const src=sourceRowSizes(source),dst=sourceRowSizes(exported);
 const columnBlock=xml=>(xml.match(/<cols\b[^>]*>[\s\S]*?<\/cols>/)||[])[0];
 assert.equal(columnBlock(exported),columnBlock(source),"Bay column widths changed on day "+day);
 const sectionRow=r=>(r>=14&&r<=19)||(r>=30&&r<=35);
 for(const [r,height] of src.rows){
   const originalRow=(source.match(new RegExp('<row\\b[^>]*\\br="'+r+'"[^>]*>'))||[])[0]||"";
   const custom=/\\bht="[^"]+"/.test(originalRow);
   const expected=!custom&&sectionRow(r)?21:height;
   assert.equal(dst.rows.get(r),expected,"Row "+r+" lost its original geometry or calibrated header height on day "+day);
 }
 for(const r of [14,15,16,17,18,19,30,31,32,33,34,35]){
   const tag=(exported.match(new RegExp('<row\\b[^>]*\\br="'+r+'"[^>]*>'))||[])[0];
   assert(tag&&tag.includes('customHeight="1"'),
     "Bay header or transition row "+r+" not locked on day "+day);
   const originalTag=(source.match(new RegExp('<row\\b[^>]*\\br="'+r+'"[^>]*>'))||[])[0]||"";
   const explicit=(originalTag.match(/\\bht="([^"]+)"/)||[])[1];
   assert.equal(Number((tag.match(/\bht="([^"]+)"/)||[])[1]),explicit?Number(explicit):21,
     "Bay header row "+r+" did not match the original visual spacing on day "+day);
 }
 for(const r of [4,5,8,9,13,20,21,24,25,29,36,40,41,45]){
   const tag=(exported.match(new RegExp('<row\\b[^>]*\\br="'+r+'"[^>]*>'))||[])[0];
   assert(tag&&tag.includes('customHeight="1"'),
     "Event row "+r+" isn't fixed-height on day "+day);
   assert.equal(Number((tag.match(/\bht="([^"]+)"/)||[])[1]),src.rows.get(r));
 }
}
console.log("PASS: All seven sheets preserve custom row heights and bay widths; floor-divider rows match the 21pt reference screenshot.");

const workbook=await read(outputZip,"xl/workbook.xml");
for(const [i,name] of ["Monday 10.12","Tuesday 10.13","Wednesday 10.14","Thursday 10.15","Friday 10.16","Saturday 10.17","Sunday 10.18"].entries()){
 assert(workbook.includes('name="'+name+'"'),"Missing dated tab: "+name);
}
assert(workbook.includes('name="Specialty Roles"'),"Reference tab renamed");
const insertedMerges=[...newSheet.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"/g)].map(m=>m[1]);
assert(insertedMerges.includes("E4:H5"),"First event merge should cover 2 rows");
assert(insertedMerges.includes("E6:H7"),"Second event should follow immediately");
assert(!insertedMerges.some(m=>m==="E7:H8"||m==="E8:H9"),"Incorrect gap placement");
for(const pathname of Object.keys(originalZip.files)){
 if(originalZip.files[pathname].dir)continue;
 if(pathname==="xl/styles.xml"||pathname==="xl/workbook.xml"||/^xl\/worksheets\/sheet[1-7]\.xml$/.test(pathname))continue;
 assert.deepEqual(
   new Uint8Array(await outputZip.file(pathname).async("uint8array")),
   new Uint8Array(await originalZip.file(pathname).async("uint8array")),
   "Unrelated original master component changed: "+pathname
 );
}
console.log("PASS: two-row events stack directly, row five stays empty, and original heavy borders and all non-event workbook components are preserved.");
process.exit(0);
