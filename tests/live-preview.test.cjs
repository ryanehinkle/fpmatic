const assert=require("node:assert/strict"),fs=require("node:fs"),vm=require("node:vm"),path=require("node:path");
const JSZip=require("../jszip.min.js");
const sync=fs.readFileSync(path.join(__dirname,"..","floorplan-sync.js"),"utf8").replace(/^export /gm,"");
const preview=fs.readFileSync(path.join(__dirname,"..","floorplan-preview.js"),"utf8")
  .replace(/^import[^\n]*\n/,"").replace(/^export /gm,"");
const compiled=vm.runInNewContext(sync+"\n"+preview+"\n({readWorkbookPreview,workbookDayPaths})",{JSZip,console,Map,Set});
(async()=>{
 const bytes=fs.readFileSync(path.join(__dirname,"..","master.xlsx"));
 const result=await compiled.readWorkbookPreview(bytes);
 assert.equal(result.days.length,7);
 for(const day of result.days){
   assert(Array.isArray(day.regions),"Missing preview day data");
   assert(day.regions.every(g=>g.area.c1>=5&&g.area.c2<=38),"Preview extends outside the event bay grid");
 }
 // A blank master may contain zero bay merges. Insert a synthetic merged event
 // so the preview parser has an actual booking rectangle to verify.
 const copy=await JSZip.loadAsync(bytes),source=await copy.file("xl/worksheets/sheet1.xml").async("string");
 const withEvent=source.replace("</mergeCells>",'<mergeCell ref="E4:H6"/></mergeCells>');
 copy.file("xl/worksheets/sheet1.xml",withEvent);
 const withMerge=await compiled.readWorkbookPreview(await copy.generateAsync({type:"nodebuffer"}));
 const region=withMerge.days[0].regions.find(g=>g.area.c1===5&&g.area.c2===8&&g.area.r1===4&&g.area.r2===6);
 assert(region,"Monday first bay section missing");
 assert(/^(#(?:[A-Fa-f0-9]{6})|transparent)$/.test(region.fill),"Invalid preview fill");
 assert.equal(compiled.workbookDayPaths('<sheets><sheet name="Monday 10.12" sheetId="1"/><sheet name="AM Monday 10.12" sheetId="2"/><sheet name="Tuesday 10.13" sheetId="3"/><sheet name="Wednesday 10.14" sheetId="5"/><sheet name="Thursday 10.15" sheetId="7"/><sheet name="Friday 10.16" sheetId="9"/><sheet name="Saturday 10.17" sheetId="11"/><sheet name="Sunday 10.18" sheetId="12"/></sheets>')[4],"xl/worksheets/sheet9.xml");
 console.log("PASS: actual original workbook preview parses all seven days, merges, bay locations, fills, and worksheet IDs");
})().catch(e=>{console.error(e);process.exitCode=1});
