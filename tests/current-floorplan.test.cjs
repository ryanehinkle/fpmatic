const fs=require("node:fs"),vm=require("node:vm"),assert=require("node:assert/strict"),path=require("node:path");
const source=fs.readFileSync(path.join(__dirname,"..","floorplan-sync.js"),"utf8").replace(/^export /gm,"");
const {reconcileDay,workbookDayPaths}=vm.runInNewContext(source+"\n({reconcileDay,workbookDayPaths})",{});
const xml='<worksheet><sheetData>'+
'<row r="9"><c r="Q9" s="2" t="inlineStr"><is><t>Bryce S SL 5\nHH04</t></is></c><c r="AB9" s="2" t="inlineStr"><is><t>Other employee</t></is></c></row>'+
'<row r="25"><c r="O25" s="3" t="inlineStr"><is><t>Stephen&#39;s Surprise Birthday Party\nSig S 630-9\n25ppl\nUlt Bites\nEA Bar- Nick T 5</t></is></c></row>'+
'<row r="28"><c r="O28" s="2" t="inlineStr"><is><t>Kelly G 5 SL\n*online learning*\nHH11</t></is></c></row>'+
'</sheetData><mergeCells count="3"><mergeCell ref="O25:S27"/><mergeCell ref="O28:S29"/><mergeCell ref="Q9:X13"/></mergeCells></worksheet>';
const styles={clear:()=>1,event:()=>3,center:()=>2};
const display=e=>[e.name,"6:30-9",e.guests+"ppl",e.package];
const updated={id:"stephen",name:"Stephen's Surprise Birthday Party",bays:"313-316",start:"18:30",end:"21:00",guests:40,package:"Ultimate",category:"Buffet",status:"Definite"};
const x=reconcileDay(xml,[],[updated],styles,display);
function cell(s,ref){const re=new RegExp('<c\\b(?=[^>]*\\br="'+ref+'")[^>]*>[\\s\\S]*?<\\/c>');return (s.match(re)||[])[0]||""}
function merges(s){return [...s.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"/g)].map(x=>x[1])}
assert.equal(x.placed,1,JSON.stringify(x.warnings));
assert(x.changes.some(x=>x.type==="guests"&&x.before==="25ppl"&&x.after==="40ppl"));
assert(x.changes.some(x=>x.type==="bays"&&x.before==="211-215"&&x.after==="313-316"));
assert(cell(x.xml,"Q9").includes("Stephen"),"Stephen should be moved to Bryce's section");
assert(cell(x.xml,"Q9").includes("Nick T 5"),"Attached event staffing note should move with Stephen");
assert(cell(x.xml,"O25").includes("Kelly"),"Kelly should expand into Stephen's old section");
assert(!cell(x.xml,"O28").includes("Kelly"),"Kelly's old cell must not retain duplicate text");
assert(cell(x.xml,"Q12").includes("Bryce"),"Bryce staffing must survive event placement");
assert(cell(x.xml,"AB9").includes("Other employee"),"Unrelated staffing was overwritten");
assert(merges(x.xml).includes("O25:S29"),"Stephen's vacated area did not merge with Kelly");
assert(merges(x.xml).includes("Q9:T11"),"New Stephen bay assignment not merged");
assert(merges(x.xml).includes("Q12:X13"),"Bryce's remaining staffing area not merged");
assert(!merges(x.xml).includes("O25:S27"),"Old event merge was not removed");
const same={...updated,bays:"211-215",guests:40};
const y=reconcileDay(xml,[],[same],styles,display);
assert.equal(y.placed,1);
assert(merges(y.xml).includes("O25:S27"),"In-place update removed original merge");
assert(merges(y.xml).includes("O28:S29"),"In-place update should retain Kelly's original merge");
assert(cell(y.xml,"O25").includes("40ppl"),"In-place guest update failed");
assert(cell(y.xml,"O28").includes("Kelly"),"In-place update modified Kelly");
assert.equal(workbookDayPaths('<sheets><sheet name="Monday 10.12" sheetId="1"/><sheet name="AM Monday 10.12" sheetId="2"/><sheet name="Tuesday 10.13" sheetId="3"/><sheet name="Tuesday 10.13 AM" sheetId="4"/><sheet name="Wednesday 10.14" sheetId="5"/><sheet name="Thursday 10.15" sheetId="7"/><sheet name="Friday 10.16" sheetId="9"/><sheet name="Saturday 10.17" sheetId="11"/><sheet name="Sunday 10.18" sheetId="12"/></sheets>')[4],"xl/worksheets/sheet9.xml");

// Mid-shift employee sections span all ten AM/PM rows. They must remain
// continuous until an event actually needs the space, then split without
// losing the original staff shared string or font assignment.
const tall='<worksheet><sheetData><row r="36">'+
  '<c r="E36" s="62" t="s"><v>0</v></c>'+
  '<c r="M36" s="91" t="s"><v>1</v></c></row>'+
  '</sheetData><mergeCells count="2"><mergeCell ref="E36:L45"/><mergeCell ref="M36:T45"/></mergeCells></worksheet>';
const richStyles={...styles,staff:()=>417};
const longEvent={name:"New Midday Booking",bays:"101-103",start:"12:00",end:"15:00",guests:20,package:"Fajita",category:"Buffet",status:"Definite"};
const shifted=reconcileDay(tall,["Aidan R DBL\nw/ Erick P","Violet B 12-8"],[longEvent],richStyles,display);
assert.equal(shifted.placed,1,"Midday event should fit inside tall first-floor staff allocation");
assert(merges(shifted.xml).includes("E39:L45"),"Remaining staffed merge must span the AM/PM divide");
assert(merges(shifted.xml).includes("H36:L38"),"The adjacent upper staff section must remain merged: "+JSON.stringify(merges(shifted.xml)));
assert(merges(shifted.xml).includes("M36:T45"),"Unaffected 12–8 section must stay ten rows tall");
assert(cell(shifted.xml,"E39").includes('t="s"><v>0</v>'),"Moved staff label lost its original rich/shared-string reference");
assert(cell(shifted.xml,"E39").includes('s="417"'),"Moved staff label did not retain its source style and bottom alignment");
assert(cell(shifted.xml,"M36").includes('t="s"><v>1</v>'),"Unrelated mid-shift staff label was changed");
const keepLong=reconcileDay(tall,["Aidan R DBL\nw/ Erick P","Violet B 12-8"],[],richStyles,display);
assert(merges(keepLong.xml).includes("E36:L45")&&merges(keepLong.xml).includes("M36:T45"),"Unchanged tall staffing merges were modified");
assert(cell(keepLong.xml,"E36").includes('t="s"><v>0</v>'),"Unchanged 12–8 staff text was rewritten");
// When an inserted event splits a staffing section without covering its
// top-left cell, staff still belongs in the broad lower section, not a
// one/two-bay strip on the left (Izack example).
const skinny='<worksheet><sheetData><row r="20"><c r="O20" s="22" t="inlineStr"><is><t>Izack P 930</t></is></c></row></sheetData>'+
  '<mergeCells count="1"><mergeCell ref="O20:W24"/></mergeCells></worksheet>';
const izackEvent={name:"Joey Birthday",bays:"213-217",start:"10:00",end:"12:00",guests:19,package:"Flex",category:"Flex",status:"Definite"};
const relocated=reconcileDay(skinny,[],[izackEvent],richStyles,display);
assert.equal(relocated.placed,1,JSON.stringify(relocated.warnings));
assert(merges(relocated.xml).includes("O23:W24"),"Broad lower staff merge disappeared");
assert(merges(relocated.xml).includes("O20:P22"),"Unoccupied side strip should remain");
assert(cell(relocated.xml,"O23").includes("Izack P 930"),"Staff name should move to the bottom section");
assert(!cell(relocated.xml,"O20").includes("Izack P 930"),"Staff name still stranded in the narrow strip");

// Blank unmerged bays should become a single rectangular section, but do not
// merge across an existing dark divider or swallow an employee label.
const emptyGrid='<worksheet><sheetData><row r="4"><c r="E4" s="1"/><c r="F4" s="1"/><c r="G4" s="1"/></row>'+
  '<row r="5"><c r="E5" s="1"/><c r="F5" s="1"/><c r="G5" s="1"/></row>'+
  '</sheetData><mergeCells count="1"><mergeCell ref="H4:AL13"/></mergeCells></worksheet>';
const blank=reconcileDay(emptyGrid,[],[],styles,display);
assert(merges(blank.xml).includes("E4:G8"),"Blank AM bay section was not merged");
assert(merges(blank.xml).includes("E9:G13"),"Blank PM bay section was not merged");
const dividerStyles={...styles,hardDivider:(a,b,direction)=>direction==="vertical"&&a===1&&b===2};
const divided=emptyGrid.replace('<c r="G4" s="1"/>','<c r="G4" s="2"/>').replace('<c r="G5" s="1"/>','<c r="G5" s="2"/>');
const split=reconcileDay(divided,[],[],dividerStyles,display);
assert(!merges(split.xml).some(m=>m==="E4:G8"),"Merged blank cells across a dark original divider");
const labeled=emptyGrid.replace('<c r="F4" s="1"/>','<c r="F4" s="1" t="inlineStr"><is><t>Staff One</t></is></c>');
const untouched=reconcileDay(labeled,[],[],styles,display);
assert(cell(untouched.xml,"F4").includes("Staff One"),"Blank-area merging removed an employee label");
assert(!merges(untouched.xml).some(m=>/^[A-Z]+\d+:[A-Z]+\d+$/.test(m)&&(()=>{const [a,b]=m.split(":");return a==="E4"&&b==="G8"})()),"Blank merge covered an employee");
console.log("PASS: moved birthday expands Kelly, preserves Bryce, in-place updates retain staffing, custom 16-sheet workbook selects real weekdays");
