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
 console.log("PASS: pasted Salesforce HTML, full Opportunity Name matching, safe URLs, invisible Excel hyperlinks, and unaffected formatting");
})().catch(e=>{console.error(e);process.exitCode=1});
