// The uploaded master—not an assumed uniform sheet—is the size authority.
const fs=require("node:fs"),vm=require("node:vm"),assert=require("node:assert/strict");
const path=require("node:path");
const source=fs.readFileSync(path.join(__dirname,"..","excel-export.js"),"utf8")
  .replace(/^import[^\n]*\n/gm,"").replace(/^export /gm,"");
const {lockMasterRowHeights}=vm.runInNewContext(source+"\n({lockMasterRowHeights})",{});
const master='<?xml version="1.0"?><worksheet><sheetFormatPr defaultRowHeight="13.25" defaultColWidth="8.43"/>'+
  '<cols><col min="5" max="38" width="10.14" customWidth="1"/></cols>'+
  '<sheetData>'+
  '<row r="14"><c r="E14" s="12"/></row>'+
  '<row r="19" ht="16.5" customHeight="1"><c r="E19" s="22"/></row>'+
  '<row r="20" ht="52.5" customHeight="1"><c r="E20" s="35"/></row>'+
  '<row r="30" hidden="1" outlineLevel="1"><c r="E30" s="1"/></row>'+
  '<row r="31"><c r="E31" s="4"/></row>'+
  '<row r="32" ht="11.75"><c r="E32" s="5"/></row>'+
  '<row r="34" customHeight="0"><c r="E34" s="8"/></row>'+
  '</sheetData><mergeCells count="0"></mergeCells></worksheet>';
const output=master
  .replace('width="10.14"','width="28.0"')
  .replace(' ht="16.5" customHeight="1"',' ht="61.0"')
  .replace(' ht="52.5" customHeight="1"',' ht="10.0"')
  .replace('row r="31"','row r="31" ht="45.0"')
  .replace('<row r="34" customHeight="0"','<row r="34" ht="66" customHeight="0"');
const fixed=lockMasterRowHeights(output,master);
const tag=n=>(fixed.match(new RegExp('<row\\b[^>]*\\br="'+n+'"[^>]*>'))||[])[0]||"";
const height=n=>Number((tag(n).match(/\bht="([^"]+)"/)||[])[1]);
for(const [n,expected] of [[14,21],[19,16.5],[20,52.5],[30,21],[31,21],[32,11.75],[34,21]]){
 assert.equal(height(n),expected,"Original master row "+n+" changed size");
 assert(tag(n).includes('customHeight="1"'),"Row "+n+" still subject to auto-sizing");
}
assert(tag(30).includes('hidden="1"')&&tag(30).includes('outlineLevel="1"'),
  "Source hidden/grouped metadata wasn't preserved");
assert(fixed.includes('<col min="5" max="38" width="10.14" customWidth="1"/>'),
  "Original bay column widths were changed");
assert(fixed.includes('<c r="E20" s="35"/>'),"Cell style was altered");
assert(fixed.includes('<mergeCells count="0"></mergeCells>'),"Existing merges changed");
assert.equal(lockMasterRowHeights(fixed,master),fixed,"Dimension locking should be idempotent");
console.log("PASS: 21pt floor-transition headers and original custom row heights, hidden flags, columns, cells and merges preserved");
