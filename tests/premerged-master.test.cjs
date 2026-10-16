// Synthetic equivalent of the user's pre-merged weekly master.
// No customer data or copies of uploaded workbooks are committed.
const assert=require("node:assert/strict");
const fs=require("node:fs");
const vm=require("node:vm");
const JSZip=require("../jszip.min.js");
const path=require("node:path");
const source=fs.readFileSync(path.join(__dirname,"..","excel-export.js"),"utf8")
  .replace(/^import[^\n]*\n/gm,"").replace(/^export /gm,"");
const {plotSheet,recolorStyles}=vm.runInNewContext(source+"\n({plotSheet,recolorStyles})");
(async()=>{
  const zip=await JSZip.loadAsync(fs.readFileSync(path.join(__dirname,"..","master.xlsx")));
  const original=await zip.file("xl/worksheets/sheet1.xml").async("string");
  const styles=await zip.file("xl/styles.xml").async("string");
  const styled=recolorStyles(styles);
  assert(original.includes("<mergeCells"));
  assert(!original.includes('mergeCell ref="P25:AA29"'));
  // Pre-build one large empty region, as in the customer's week (2).
  let premerged=original.replace("</mergeCells>",'<mergeCell ref="P25:AA29"/></mergeCells>');
  // Google Sheets/Excel often omit cells inside a pre-existing large merge.
  for(const col of ["S","T","U","V","W","X"])for(const row of [26,27]){
    const re=new RegExp('<c\\b(?=[^>]*\\br="'+col+row+'")[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)');
    premerged=premerged.replace(re,"");
  }
  const e={id:"test",name:"Test Event",category:"Buffet",package:"Fajita",status:"Tentative",start:"15:00",end:"19:00",date:"2026-10-19",guests:20};
  const layout={records:{"2-PM":[{event:e,range:{start:215,end:220,floor:2},row:0,height:3}]}};
  const rendered=plotSheet(premerged,[e],layout,styled.styleIds,()=>"PM",x=>[x.name,"3-7","20ppl","Fajita"]);
  assert.equal(rendered.count,1);
  assert.deepEqual(Array.from(rendered.skipped),[]);
  const merges=[...rendered.xml.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"/g)].map(x=>x[1]);
  for(const name of ["P25:R27","S25:X27","Y25:AA27","P28:AA29"])
    assert(merges.includes(name),"Expected split region "+name);
  assert(!merges.includes("P25:AA29"),"Overlapped original merge must be replaced");
  assert(rendered.xml.includes("Test Event"),"Event text was not written");
  assert(/<c r="S26" s="\d+"\/>/.test(rendered.xml),"Missing interior event cell not reinserted");
  assert(/<c r="S27" s="\d+"\/>/.test(rendered.xml),"Missing bottom event cell not reinserted");
  assert(/<c r="P28" s="\d+"\/>/.test(rendered.xml),"Blank lower region lost its anchor");
  const count=+(rendered.xml.match(/<mergeCells count="(\d+)"/)||[])[1];
  assert.equal(count,merges.length,"Merge count mismatch");
  assert(rendered.xml.includes('<c r="D25"'),"AM/PM boundary changed");

  // Section boundaries are dark ONLY where the uploaded master had them.
  // Newly created leftover merge fragments must not get black cut lines.
  function sideDetails(ref,side){
    const cell=(rendered.xml.match(new RegExp('<c\\b(?=[^>]*\\br="'+ref+'")[^>]*\\/?>'))||[])[0];
    assert(cell,"Missing modified cell "+ref);
    const styleId=Number((cell.match(/\bs="(\d+)"/)||[])[1]);
    const stylesheet=styled.finalize(styled.xml);
    const xfs=[...(stylesheet.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)||[])[1]
      .matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)].map(m=>m[0]);
    const borderId=Number((xfs[styleId].match(/\bborderId="(\d+)"/)||[])[1]||0);
    const borders=[...(stylesheet.match(/<borders\b[^>]*>([\s\S]*?)<\/borders>/)||[])[1]
      .matchAll(/<border\b[^>]*\/>|<border\b[^>]*>[\s\S]*?<\/border>/g)].map(m=>m[0]);
    const xml=borders[borderId]||"";
    const tag=(xml.match(new RegExp("<"+side+"\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/"+side+">)"))||[])[0]||"";
    return {style:(tag.match(/\bstyle="([^"]+)"/)||[])[1]||"",
      color:(tag.match(/\brgb="([^"]+)"/)||[])[1]||""};
  }
  assert.equal(sideDetails("R25","right").style,"","No dark outline at the new left cut");
  assert.equal(sideDetails("Y25","left").style,"","No dark outline at the new right cut");
  assert.equal(sideDetails("P28","top").style,"","No dark divider above the bottom blank fragment");
  assert.equal(sideDetails("S26","left").color,"FFD9D9D9","Event outlines must be pale gray");
  assert.equal(sideDetails("X26","right").color,"FFD9D9D9","Event right edge must be pale gray");
  assert.equal(sideDetails("P26","left").color,"FF000000","Original left section border lost");
  assert.equal(sideDetails("AA26","right").color,"FF000000","Original right section border lost");
  // Tentative font should be black, not the former white override.
  const staffStyleId=styled.styleIds.staff(0);
  const linkedStyleId=styled.styleIds.link(0);
  const finalized=styled.finalize(styled.xml);
  const eventCell=(rendered.xml.match(/<c r="S25" s="(\d+)"/)||[])[1];
  assert(eventCell,"Missing event anchor");
  const xfNodes=[...(finalized.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)||[])[1].matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)];
  const chosen=xfNodes[Number(eventCell)][0];
  const fontId=Number((chosen.match(/fontId="(\d+)"/)||[])[1]);
  const fonts=[...(finalized.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/)||[])[1].matchAll(/<font(?:\s[^>]*)?>[\s\S]*?<\/font>|<font\/>/g)];
  assert(fonts[fontId][0].includes("FF000000"),"Tentative font must be black");
  const allXfs=[...(finalized.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/)||[])[1].matchAll(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g)].map(x=>x[0]);
  const centeredStaff=allXfs[staffStyleId];
  assert(centeredStaff,"Missing generated staff alignment style");
  assert(/<alignment\b[^>]*horizontal="center"/.test(centeredStaff),"Staff name is not horizontally centered");
  assert(/<alignment\b[^>]*vertical="center"/.test(centeredStaff),"Staff name is not vertically centered");
  const originalStyle=allXfs[0];
  for(const attribute of ["fontId","fillId","borderId"]){
    const value=(originalStyle.match(new RegExp("\\b"+attribute+'="(\\d+)"'))||[])[1]||"0";
    const updated=(centeredStaff.match(new RegExp("\\b"+attribute+'="(\\d+)"'))||[])[1]||"0";
    assert.equal(updated,value,"Centering altered original staff "+attribute);
  }
  const linkedStyle=allXfs[linkedStyleId];
  assert(linkedStyle,"No linked event style generated");
  for(const attribute of ["fillId","borderId"]){
    const before=(originalStyle.match(new RegExp("\\b"+attribute+'="(\\d+)"'))||[])[1]||"0";
    const after=(linkedStyle.match(new RegExp("\\b"+attribute+'="(\\d+)"'))||[])[1]||"0";
    assert.equal(after,before,"Linked event changed original "+attribute);
  }
  const linkedFontId=Number((linkedStyle.match(/\bfontId="(\d+)"/)||[])[1]||0);
  const fontXml=(finalized.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/)||[])[1]||"";
  const fonts=[...fontXml.matchAll(/<font\b[^>]*\/>|<font\b[^>]*>[\s\S]*?<\/font>/g)].map(m=>m[0]);
  assert(fonts[linkedFontId]?.includes('<u val="none"/>'),"Linked font must explicitly disable underlining");
  assert(/<color\b/.test(fonts[linkedFontId]||""),"Linked font needs an explicit original color to prevent blue text");
  // A manually pre-populated block is never overwritten by imported events.
  const occupied=premerged.replace(/<c r="P25"[^>]*\/>/, '<c r="P25" s="2" t="inlineStr"><is><t>Already Assigned</t></is></c>');
  const result2=plotSheet(occupied,[e],layout,styled.styleIds,()=>"PM",x=>[x.name]);
  assert.equal(result2.count,0);
  assert.equal(result2.skipped.length,1);
  assert(result2.xml.includes("Already Assigned"),"Prefilled event content was altered");
  assert(result2.xml.includes('mergeCell ref="P25:AA29"'),"Prefilled merged region was changed");
  console.log("PASS: premerged weekly master split into adjacent white sections around new event; tentative text black; occupied sections intact.");
})().catch(e=>{console.error(e);process.exitCode=1});
