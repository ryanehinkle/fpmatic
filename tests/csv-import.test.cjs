// Report-import regression checks; run via node tests/csv-import.test.cjs.
// Uses synthetic examples to avoid committing real client/event details.
const fs=require("node:fs");
const vm=require("node:vm");
const assert=require("node:assert/strict");
const path=require("node:path");
const source=fs.readFileSync(path.join(__dirname,"..","csv-import.js"),"utf8").replace(/export function /g,"function ");
const {parseCsv,reportType,buildEvents}=vm.runInNewContext(source+";({parseCsv,reportType,buildEvents})",{},{timeout:5000});
const staffingCsv=[
 '"Event End Time","Stage","Guest Count","Bay Numbers","Product Name","Start Time","End Time","Event Date","Event Start Time","Opportunity Name"',
 '"3:00 PM","Definite","20","211-214","Signature Room South","12:00 PM","1:00 PM","10/15/2026","12:00 PM","Test Corporation - Smith, Jane - 10.15.26"',
 '"3:00 PM","Definite","20","211-214","Event Golf Bays","1:00 PM","3:00 PM","10/15/2026","12:00 PM","Test Corporation - Smith, Jane - 10.15.26"',
 '"3:00 PM","Definite","20","211-214","Topgolf Gameplay Per Person - 2 Hours","1:00 PM","3:00 PM","10/15/2026","12:00 PM","Test Corporation - Smith, Jane - 10.15.26"',
 '"1:00 PM","Definite","80","101-120","Field Trip - Two (2) Hours Topgolf","11:00 AM","1:00 PM","10/14/2026","11:00 AM","School Trip - John Smith - 10.14.26"',
 '"1:00 PM","Definite","80","101-120","Event Golf Bays","11:00 AM","1:00 PM","10/14/2026","11:00 AM","School Trip - John Smith - 10.14.26"',
 '"8:00 PM","Agreement Sent","16","317-318","Event Golf Bays","6:00 PM","8:00 PM","10/17/2026","6:00 PM","Test Birthday - Casey - 10.17.26"',
 '"4:00 PM","Definite","12","","Event Golf Bays","2:00 PM","4:00 PM","10/17/2026","2:00 PM","Missing Bays - Casey - 10.17.26"',
 '"6:00 PM","Definite","12","331-332","Multi-Bay Game Play - Per Bay - 2 Hours","4:00 PM","6:00 PM","10/18/2026","4:00 PM","Flex Event - Casey - 10.18.26"'
].join("\n");
const menuCsv=[
 '"Event Start","Quantity","Product Name","Catering Location","Opportunity Name"',
 '"12:00 PM","20","Smokehouse BBQ","In room","Test Corporation - Smith, Jane - 10.15.26"',
 '"11:00 AM","80","Chocolate Chip Cookie - Field Trip","Field Trip Menu","School Trip - John Smith - 10.14.26"',
 '"6:00 PM","2","Crowd Pleaser - Queso, Guac, Salsa, Wings","In the Golf Bays","Test Birthday - Casey - 10.17.26"'
].join("\n");
const staff=parseCsv(staffingCsv),menus=parseCsv(menuCsv);
assert.equal(reportType(staff),"staffing");assert.equal(reportType(menus),"menus");
const result=buildEvents(staff,menus);
assert.equal(result.events.length,5);
assert.equal(result.matched,3);assert.equal(result.unmatched.length,0);
const byName=name=>result.events.find(e=>e.name===name);
const company=byName("Test Corporation");
assert(company);assert.equal(company.date,"2026-10-15");assert.equal(company.bays,"211-214");
assert.equal(company.space,"Sig S");assert.equal(company.package,"Smokehouse");
assert.equal(company.golfStart,"13:00");assert.equal(company.golfEnd,"15:00");
const school=byName("School Trip");assert.equal(school.category,"Field Trip");
assert.equal(school.guests,80);assert.equal(school.golfStart,"");
const birthday=byName("Test Birthday");assert.equal(birthday.category,"Pleaser");assert.equal(birthday.status,"Tentative");
const noBays=byName("Missing Bays");assert.equal(noBays.bays,"");assert(noBays.importFlags.includes("No bay assignment"));
assert.equal(byName("Flex Event").category,"Flex");
assert.equal(parseCsv('"A","B"\r\n"one, two","quoted ""value"""\r\n')[0].B,'quoted "value"');
assert.throws(()=>reportType(parseCsv("Unknown,Another\n1,2")),/Unrecognized CSV/);

const more=[
 '"3:00 PM","Definite","80","201-208","Signature Room North","12:00 PM","3:00 PM","10/18/2026","12:00 PM","Full Rooms Event - Contact - 10.18.26"',
 '"3:00 PM","Definite","80","201-208","Signature Room South","12:00 PM","3:00 PM","10/18/2026","12:00 PM","Full Rooms Event - Contact - 10.18.26"',
 '"1:00 PM","Definite","12","133-134","Event Golf Bays","9:00 AM","1:00 PM","10/18/2026","9:00 AM","Tom's Lessons - Contact - 10.18.26"'
].join("\n");
const merged=buildEvents(parseCsv(staffingCsv+"\n"+more),menus);
assert.equal(merged.events.length,6,"Only the Full Sig event should be added, not Tom's Lessons");
assert.equal(merged.events.find(x=>x.name==="Full Rooms Event").space,"Full Sig");
assert(!merged.events.some(x=>x.name==="Tom's Lessons"));
console.log("PASS — CSV joins, package colors, duplicate product rows, bay assignments, room/golf hours, tentative statuses, and missing-bay warnings.");
