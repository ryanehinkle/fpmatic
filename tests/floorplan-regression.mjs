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
const originalZip=await JSZip.loadAsync(original);
const sourceModule=fs.readFileSync(new URL("../excel-export.js",import.meta.url),"utf8");
const exporter=await import("data:text/javascript;base64,"+Buffer.from(sourceModule).toString("base64"));

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
   state:{master:new Uint8Array(original),week:monday},weekEvents:data,layoutDay,
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
 const regex=new RegExp('<c\\b(?=[^>]*\\br="'+ref+'")[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)');
 const match=xml.match(regex);
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
for(const ref of ["E4","E5","E6","E7","H4","H7","AG4","AL4","AL5","AL6","E20","E22"]){
 assert.equal(
   Number((xf(newStyles,styleIndex(cell(newSheet,ref))).match(/\bborderId="(\d+)"/)||[])[1]||0),
   Number((xf(oldStyles,styleIndex(cell(oldSheet,ref))).match(/\bborderId="(\d+)"/)||[])[1]||0),
   "Template border changed on "+ref
 );
}
for(const ref of ["E8","AL8","E9","D4","D8","AM4","AM9","D9"]){
 assert.equal(cell(newSheet,ref),cell(oldSheet,ref),"Master non-event cell unexpectedly rewritten at "+ref);
}
const insertedMerges=[...newSheet.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"/g)].map(m=>m[1]);
assert(insertedMerges.includes("E4:H5"),"First event merge should cover 2 rows");
assert(insertedMerges.includes("E6:H7"),"Second event should follow immediately");
assert(!insertedMerges.some(m=>m==="E7:H8"||m==="E8:H9"),"Incorrect gap placement");
for(const pathname of Object.keys(originalZip.files)){
 if(pathname==="xl/styles.xml"||/^xl\/worksheets\/sheet[1-7]\.xml$/.test(pathname))continue;
 assert.deepEqual(
   new Uint8Array(await outputZip.file(pathname).async("uint8array")),
   new Uint8Array(await originalZip.file(pathname).async("uint8array")),
   "Unrelated original master component changed: "+pathname
 );
}
console.log("PASS: two-row events stack directly, row five stays empty, and original heavy borders and all non-event workbook components are preserved.");
process.exit(0);
