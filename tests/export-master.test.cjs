// Regression test: exporting events may ONLY alter intended bay cells,
// merges and append brand-new styles. All original master parts stay intact.
const fs=require("node:fs");
const vm=require("node:vm");
const assert=require("node:assert/strict");
const JSZip=require("../jszip.min.js");

const source=fs.readFileSync(require("node:path").join(__dirname,"..","excel-export.js"),"utf8")
  .replace("export async function exportExactFloorPlan","async function exportExactFloorPlan");
const {recolorStyles,plotSheet}=vm.runInNewContext(
  source+"\n({recolorStyles,plotSheet})",
  {}, {timeout:5000}
);
(async()=>{
  const template=fs.readFileSync(require("node:path").join(__dirname,"..","master.xlsx"));
  const original=await JSZip.loadAsync(template,{checkCRC32:true});
  const requiredSheets=Array.from({length:11},(_,i)=>"xl/worksheets/sheet"+(i+1)+".xml");
  for(const name of requiredSheets)assert(original.file(name), "Missing original worksheet "+name);
  const oldStyles=await original.file("xl/styles.xml").async("string");
  const {xml:updatedStyles,styleIds}=recolorStyles(oldStyles);
  const oldCount=+(oldStyles.match(/<cellXfs\b[^>]*count="(\d+)"/)||[])[1];
  const newCount=+(updatedStyles.match(/<cellXfs\b[^>]*count="(\d+)"/)||[])[1];
  assert.equal(newCount,oldCount+20,"Original style indices must remain stable");
  assert.equal(styleIds.Buffet.Definite.normal>=oldCount,true);
  const originalDay=await original.file(requiredSheets[0]).async("string");
  const event={name:"Test Company",category:"Buffet",status:"Definite",start:"18:00",end:"20:00"};
  const out=plotSheet(originalDay,[event],{
    records:{"3-PM":[{event,range:{floor:3,start:309,end:313},row:0,height:3}]}
  },styleIds,()=>"PM",e=>[e.name,"6-8","30ppl","Fajita"]);
  assert.equal(out.count,1);
  assert(out.xml.includes('mergeCell ref="M9:Q11"'),"New event must be merged at correct bays");
  assert(out.xml.includes('t="inlineStr"><is><t xml:space="preserve">Test Company'));
  assert(out.xml.includes('<c r="E3"'),"Bay headers must be preserved");
  assert(out.xml.includes('<c r="AO1"'),"Staffing grid must be preserved");
  original.file("xl/styles.xml",updatedStyles);
  original.file(requiredSheets[0],out.xml);
  const bytes=await original.generateAsync({type:"nodebuffer",compression:"DEFLATE"});
  const roundtrip=await JSZip.loadAsync(bytes,{checkCRC32:true});
  assert.equal((await roundtrip.file(requiredSheets[0]).async("string")),out.xml);
  assert.equal((await roundtrip.file("xl/styles.xml").async("string")),updatedStyles);
  for(const name of requiredSheets.slice(1)){
    const before=(await JSZip.loadAsync(template)).file(name);
    assert.equal(await before.async("string"),await roundtrip.file(name).async("string"),
      "Unrelated worksheet was changed: "+name);
  }
  console.log("PASS — exact original master preserved; event cell/merge/styling patch valid");
})().catch(e=>{console.error(e);process.exitCode=1;});
