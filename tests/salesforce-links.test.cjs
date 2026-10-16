const fs=require("node:fs"),vm=require("node:vm"),assert=require("node:assert/strict"),path=require("node:path");
const base=p=>fs.readFileSync(path.join(__dirname,"..",p),"utf8").replace(/^import[^\n]*\n/gm,"").replace(/^export /gm,"");
const all=base("floorplan-sync.js")+"\n"+base("salesforce-links.js")+"\n"+base("sheet-hyperlinks.js");
const api=vm.runInNewContext(all+"\n({extractSalesforceLinks,matchingSalesforceLink,linkForWorksheetText,addSheetHyperlinks,applyEventHyperlinks})",{URL,Map,Set,console});
const id="006xx00000123ABCDE";
const href="https://topgolf.lightning.force.com/lightning/r/Opportunity/"+id+"/view?tab=Details&from=report";
const label="Stephen's Surprise Birthday Party - Jane Smith - 10.16.26";
const html='<table><tr><td><a class="slds-truncate" href="'+href.replace("&","&amp;")+'"><span>'+
  'Stephen&#39;s Surprise Birthday Party - Jane Smith - 10.16.26</span></a></td></tr>'+
  '<tr><td><a href="https://evil.example/steal">Malicious Event - Person - 10.16.26</a></td></tr>'+
  '<tr><td><a href="'+href+'">Stephen&#39;s Surprise Birthday Party - Jane Smith - 10.16.26</a></td></tr>'+
  '<tr><td><a href="https://topgolf.lightning.force.com/lightning/r/Opportunity/OTHER/view">Tom&#39;s Lessons - Contact - 10.16.26</a></td></tr></table>';
const found=api.extractSalesforceLinks(html);
assert.equal(found.length,2,"Only Salesforce opportunities should survive, without duplicates");
const event={name:"Stephen's Surprise Birthday Party",date:"2026-10-16",opportunityName:label};
const match=api.matchingSalesforceLink(event,found);
assert(match,"Full opportunity name did not match");
assert.equal(match.url,href,"Ampersand or origin lost in Salesforce link");
assert.equal(api.linkForWorksheetText("Stephen's Surprise Birthday Party\n6-9\n25ppl\nUltimate","2026-10-16",[event],found),href);
assert.equal(api.linkForWorksheetText("Unlinked Event\n12ppl","2026-10-16",[],found),null);
assert.equal(api.matchingSalesforceLink({name:"Different event",date:"2026-10-16",opportunityName:"Different event - Contact - 10.16.26"},found),null);
assert(!api.extractSalesforceLinks('<a href="javascript:alert(1)">Bad - Person - 10.16.26</a>').length);
const oldXml='<?xml version="1.0"?><worksheet><sheetData><row r="4">'+
  '<c r="E4" s="22" t="inlineStr"><is><t>Stephen&#39;s Surprise Birthday Party\n6-9\n25ppl\nUltimate</t></is></c>'+
  '<c r="F4" s="34"/></row></sheetData><mergeCells count="1"><mergeCell ref="E4:H6"/></mergeCells>'+
  '<pageMargins left="0.5" right="0.5"/></worksheet>';
const originalAnchor=(oldXml.match(/<c r="E4"[\s\S]*?<\/c>/)||[])[0];
const result=api.addSheetHyperlinks(oldXml,null,[{ref:"E4",url:href}]);
assert.equal(result.added,1);
assert(result.xml.includes('<hyperlink ref="E4" r:id="rId1"/>'));
assert(result.xml.indexOf('<hyperlinks>')>result.xml.indexOf('</mergeCells>'));
assert(result.xml.indexOf('</hyperlinks>')<result.xml.indexOf('<pageMargins'));
assert(result.xml.includes('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'));
assert.equal((result.xml.match(/<c r="E4"[\s\S]*?<\/c>/)||[])[0],originalAnchor,"Visible Excel event text or style changed");
assert(result.rels.includes('TargetMode="External"'));
assert(result.rels.includes(href.replace("&","&amp;")));
const second=api.addSheetHyperlinks(result.xml,result.rels,[{ref:"E4",url:href}]);
assert.equal((second.xml.match(/<hyperlink ref="E4"/g)||[]).length,1,"Duplicated hyperlink for same event");
assert.equal((second.xml.match(/<hyperlink\b/g)||[]).length,1,"Unexpected repeated hyperlink");

// The user's actual Salesforce report has name/contact/date, time, bay and
// guest columns. Fuzzy names must use all fields, not just the first title.
const reportRow=(start,end,name,bays,guests,link)=>
  '<tr><td>'+start+'</td><td>'+end+'</td><td>Online Booking</td>'+
  '<td><a href="'+link+'">'+name+'</a></td><td>Executed</td>'+
  '<td>1005</td><td>'+bays+'</td><td>9</td><td>USD 0.00</td><td>'+guests+'</td></tr>';
const rowHtml='<table>'+
  reportRow("10:00 AM","12:00 PM","Sipley Elementary - Michael Davis - 10.13.26","226-234",52,
    "https://topgolf.lightning.force.com/lightning/r/Opportunity/SIPLEY/view")+
  reportRow("4:00 PM","7:00 PM","BMO Bank - Janna Lewis - 10.13.26","230-234",30,
    "https://topgolf.lightning.force.com/lightning/r/Opportunity/BMO/view")+
  reportRow("3:00 PM","5:00 PM","Sipley Student Party - Contact - 10.13.26","309-310",14,
    "https://topgolf.lightning.force.com/lightning/r/Opportunity/OTHER/view")+
  '</table>';
const parsed=api.extractSalesforceLinks(rowHtml);
assert.equal(parsed.length,3);
const sipley=parsed.find(x=>/Sipley Elementary/.test(x.name));
assert.equal(sipley.start,"10:00 AM");
assert.equal(sipley.end,"12:00 PM");
assert.equal(sipley.bays,"226-234");
assert.equal(sipley.guests,52);
const matchedSchool=api.matchingSalesforceLink({name:"Sipley Elem School",date:"2026-10-13",
  start:"10:00",end:"12:00",bays:"226-234",guests:52},parsed);
assert(matchedSchool&&/SIPLEY/.test(matchedSchool.url),"Partial school name with matching time/bays/guest should resolve");
const linkedSchool=api.linkForWorksheetText("Sipley Elementary\n10-12\n52ppl\nField Trip",
  "2026-10-13",[],parsed,"226-234");
assert(linkedSchool&&linkedSchool.includes("/SIPLEY/"),"Existing workbook event was not linked using supporting report details");
const repeated=api.linkForWorksheetText("BMO Bank\n4-7\n30ppl\nFajita",
  "2026-10-13",[],parsed,"230-234");
assert(repeated&&repeated.includes("/BMO/"),"Repeated event label failed to inherit the link");
const shortRepeated=api.linkForWorksheetText("BMO Bank\n4-7\nFajita",
  "2026-10-13",[],parsed,"230-234");
assert(shortRepeated&&shortRepeated.includes("/BMO/"),"Event repeated without guest line should still link");
assert.equal(api.matchingSalesforceLink({name:"Sipley Elem School",date:"2026-10-14",
  start:"10:00",end:"12:00",bays:"226-234",guests:52},parsed),null,"Never match across event dates");
const spreadsheetRepeated=api.addSheetHyperlinks(oldXml,null,[{ref:"E4",url:href},{ref:"F4",url:href}],n=>n+1000);
assert.equal(spreadsheetRepeated.added,2,"All appearances of an event must receive a hyperlink");
assert(spreadsheetRepeated.xml.includes('<hyperlink ref="F4"'),"Repeated event appearance had no link");
assert(spreadsheetRepeated.xml.includes('s="1022"'),"Hyperlinked event did not receive explicit visual-preservation font");
const originalSheetWithExisting=result.xml.replace("</hyperlinks>",'<hyperlink ref="G9" r:id="rId999"/></hyperlinks>');
const preserved=api.addSheetHyperlinks(originalSheetWithExisting,result.rels,[{ref:"E4",url:href}]);
assert(preserved.xml.includes('ref="G9"'),"Unrelated hyperlink was removed");
const JSZip=require("../jszip.min.js");
(async()=>{
 const original=await JSZip.loadAsync(fs.readFileSync(path.join(__dirname,"..","master.xlsx")));
 const source=await original.file("xl/worksheets/sheet1.xml").async("string");
 const add=source.replace(/<c r="E4"[^>]*(?:\/>|>[\s\S]*?<\/c>)/,originalAnchor);
 const resultMap=new Map([["xl/worksheets/sheet1.xml",add]]);
 const count=await api.applyEventHyperlinks(original,resultMap,{0:"xl/worksheets/sheet1.xml"},[],[{...event,salesforceUrl:href}],found,"2026-10-16",(_,day)=>day===0?"2026-10-16":"2026-10-17");
 assert.equal(count,1,"Linked event wasn't found in worksheet");
 const final=resultMap.get("xl/worksheets/sheet1.xml");
 assert(final.includes('ref="E4" r:id="rId'));
 assert((await original.file("xl/worksheets/_rels/sheet1.xml.rels").async("string")).includes(id));
 assert(final.includes(originalAnchor),"Event formatting changed after links were added");
 // A second AM worksheet can repeat opportunities and must also get links,
 // without moving cells or changing the existing text/section formatting.
 const workbook=await original.file("xl/workbook.xml").async("string");
 original.file("xl/workbook.xml",workbook.replace("</sheets>",
   '<sheet name="AM Tuesday 10.13" sheetId="88" r:id="rId88"/></sheets>'));
 const extra='<worksheet><sheetData><row r="5"><c r="K5" s="17" t="inlineStr"><is>'+
   '<t>BMO Bank&#10;4-7&#10;30ppl&#10;Fajita</t></is></c></row></sheetData></worksheet>';
 original.file("xl/worksheets/sheet88.xml",extra);
 const touched=new Map([["xl/worksheets/sheet1.xml",await original.file("xl/worksheets/sheet1.xml").async("string")]]);
 const auxCount=await api.applyEventHyperlinks(original,touched,{0:"xl/worksheets/sheet1.xml"},[],
   [],parsed,"2026-10-12",(week,day)=>"2026-10-"+String(12+day).padStart(2,"0"));
 assert(auxCount>=1,"Additional AM worksheet event did not receive a link");
 const updated=touched.get("xl/worksheets/sheet88.xml");
 assert(updated.includes('ref="K5"'),"Repeated BMO opportunity missing from AM worksheet");
 assert(updated.includes('s="17"'),"Extra worksheet event formatting changed unexpectedly");
 console.log("PASS: pasted Salesforce HTML, full Opportunity Name matching, safe URLs, invisible Excel hyperlinks, and unaffected formatting");
})().catch(e=>{console.error(e);process.exitCode=1});
