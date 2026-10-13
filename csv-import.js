// Salesforce report import. Event Staffing is authoritative for each booking;
// Catering Menus supplies the food/package classification. Group product rows by Opportunity Name.
export function parseCsv(text){
  text=String(text||"").replace(/^\uFEFF/,"");
  const rows=[];let row=[],value="",quoted=false;
  for(let i=0;i<text.length;i++){
    const c=text[i];
    if(quoted){
      if(c==='"'&&text[i+1]==='"'){value+='"';i++}
      else if(c==='"')quoted=false;
      else value+=c;
    }else if(c==='"')quoted=true;
    else if(c===','){row.push(value);value=""}
    else if(c==='\n'||c==='\r'){
      if(c==='\r'&&text[i+1]==='\n')i++;
      row.push(value);value="";
      if(row.some(v=>v.trim()))rows.push(row);
      row=[];
    }else value+=c;
  }
  if(quoted)throw new Error("CSV has an unfinished quoted value.");
  row.push(value);if(row.some(v=>v.trim()))rows.push(row);
  if(!rows.length)throw new Error("CSV is empty.");
  const headings=rows.shift().map(h=>h.trim().replace(/^\uFEFF/,""));
  if(new Set(headings).size!==headings.length)throw new Error("CSV contains duplicate column names");
  return rows.map((cells,index)=>{
    const out={};
    headings.forEach((h,k)=>out[h]=(cells[k]||"").trim());
    out._line=index+2;return out;
  });
}
export function reportType(rows){
  if(!rows.length)throw new Error("Report has no rows");
  const keys=Object.keys(rows[0]);
  if(["Opportunity Name","Event Date","Event Start Time","Event End Time","Guest Count","Bay Numbers","Product Name"].every(k=>keys.includes(k)))return "staffing";
  if(["Opportunity Name","Event Start","Product Name","Catering Location"].every(k=>keys.includes(k)))return "menus";
  throw new Error("Unrecognized CSV. Choose an Event Staffing or List of Catering Menus report.");
}
const pad=n=>String(n).padStart(2,"0");
function isoDate(s){
  const m=String(s||"").match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if(!m)return "";
  const year=+m[3],month=+m[1],day=+m[2];
  const d=new Date(Date.UTC(year,month-1,day));
  return d.getUTCFullYear()===year&&d.getUTCMonth()+1===month&&d.getUTCDate()===day?year+"-"+pad(month)+"-"+pad(day):"";
}
function mins(time){
  const match=String(time||"").match(/^\s*(\d{1,2}):(\d\d)\s*(AM|PM)\s*$/i);
  if(!match)return null;
  const h=+match[1],m=+match[2];if(h<1||h>12||m>59)return null;
  return (h%12+(match[3].toUpperCase()==="PM"?12:0))*60+m;
}
function clock(time){const t=mins(time);return t===null?"":pad(Math.floor(t/60))+":"+pad(t%60)}
function mostCommon(values){
  const counts=new Map();
  for(const v of values){if(!v)continue;counts.set(v,(counts.get(v)||0)+1)}
  return [...counts].sort((a,b)=>b[1]-a[1])[0]?.[0]||"";
}
function validBayRanges(s){
  return [...String(s||"").matchAll(/(\d{3})(?:\s*[-–—]\s*(\d{3}))?/g)]
    .map(x=>[+x[1],+(x[2]||x[1])])
    .filter(([a,b])=>a<=b&&Math.floor(a/100)===Math.floor(b/100)&&a%100>=1&&b%100<=34&&a>=101&&b<=334)
    .map(([a,b])=>a===b?String(a):a+"-"+b);
}
function bayValue(rows){
  const ordered=[
    ...rows.filter(r=>/^\s*Event Golf Bays\s*$/i.test(r["Product Name"]||"")),
    ...rows.filter(r=>/game\s*play|gameplay|field trip|complimentary bay/i.test(r["Product Name"]||"")),
    ...rows
  ];
  return mostCommon(ordered.map(r=>validBayRanges(r["Bay Numbers"]||"").join(", ")).filter(Boolean));
}
function opportunityTitle(name){
  const stripped=String(name||"").replace(/\s*-\s*\d{1,2}[./]\d{1,2}[./]\d{2,4}\s*$/,"").trim();
  const parts=stripped.split(/\s+-\s+/);
  if(parts.length>1)parts.pop();
  return (parts.join(" - ")||stripped||String(name)).trim().replace(/([A-Za-z])\?s\b/g,"$1\u0027s");
}
function findSpace(staff,menus){
  const products=staff.map(r=>r["Product Name"]||"").join(" | ");
  const location=menus.map(r=>r["Catering Location"]||"").join(" | ");
  const south=/Signature Room South|Sig(?:nature)?\s*Room\s*S\b/i.test(products)||/Signature South|Sig Room S\b/i.test(location);
  const north=/Signature Room North|Sig(?:nature)?\s*Room\s*N\b/i.test(products)||/Signature North|Sig Room N\b/i.test(location);
  const out=[];
  if(north&&south)out.push("Full Sig");else if(north)out.push("Sig N");else if(south)out.push("Sig S");
  if(/\bLower Lounge\b/i.test(products))out.push("LL");
  if(/\bUpper Lounge\b/i.test(products))out.push("UL");
  if(/\bTerrace\b/i.test(products))out.push("Terrace");
  return out.join(", ");
}
function mainMenu(products,staffProducts){
  const food=products.map(p=>String(p||"")).filter(Boolean);
  const staff=staffProducts.join(" | ");
  if(food.some(p=>/field\s*trip/i.test(p))||/Field Trip.*Topgolf/i.test(staff))return {category:"Field Trip",package:"Field Trip"};
  if(food.some(p=>/Youth Platter|Youth Wings|Youth Burgers|Youth Package/i.test(p))||/Youth Gameplay/i.test(staff))return {category:"Youth",package:"Youth"};
  if(food.some(p=>/Crowd Pleaser/i.test(p)))return {category:"Pleaser",package:"Pleaser"};
  const named=[
    [/Fajita Fiesta/i,"Fajita"],[/Ultimate Bites/i,"Ultimate"],[/Smokehouse BBQ/i,"Smokehouse"],
    [/\bThe Premium\b/i,"Premium"],[/\bRise\s*&\s*Shine\b/i,"Rise & Shine"],
    [/Most Valuable Buffet|\bMVB\b/i,"MVB"],[/Fairway Feast/i,"Fairway Feast"],
    [/Breakfast Buffet/i,"Breakfast"],[/\bContinental\b/i,"Continental"]
  ];
  const packages=named.filter(([re])=>food.some(p=>re.test(p))).map(x=>x[1]);
  if(packages.length)return {category:"Buffet",package:packages.join(" + ")};
  return {category:"Flex",package:"Flex"};
}
function golfWindow(rows){
  const bay=rows.find(r=>/^\s*Event Golf Bays\s*$/i.test(r["Product Name"])&&clock(r["Start Time"])&&clock(r["End Time"]));
  if(bay)return [clock(bay["Start Time"]),clock(bay["End Time"])];
  const preferred=rows.find(r=>/^(?:Youth Gameplay|Field Trip.*Topgolf|Topgolf Gameplay Per Person|Multi-Bay Game Play|2-4 Bay Game Play|Complimentary Bay)/i.test(r["Product Name"])&&clock(r["Start Time"])&&clock(r["End Time"]));
  return preferred?[clock(preferred["Start Time"]),clock(preferred["End Time"])]:null;
}
export function buildEvents(staffingRows,menuRows,staffFile="Event Staffing.csv",menuFile="Catering Menus.csv"){
  if(reportType(staffingRows)!=="staffing"||reportType(menuRows)!=="menus")throw new Error("Select both report types.");
  const groups=new Map(),catering=new Map();
  for(const r of staffingRows){
    const key=r["Opportunity Name"];if(!key)continue;
    if(!groups.has(key))groups.set(key,[]);groups.get(key).push(r);
  }
  for(const r of menuRows){
    const key=r["Opportunity Name"];if(!key)continue;
    if(!catering.has(key))catering.set(key,[]);catering.get(key).push(r);
  }
  const events=[];
  for(const [op,rows] of groups){
    if(/^Tom['’]s Lessons(?:\s*-|\s*$)/i.test(op.trim()))continue;
    const menus=catering.get(op)||[];
    const date=mostCommon(rows.map(r=>isoDate(r["Event Date"])));
    const start=clock(mostCommon(rows.map(r=>r["Event Start Time"])));
    const end=clock(mostCommon(rows.map(r=>r["Event End Time"])));
    const status=/^Definite$/i.test(mostCommon(rows.map(r=>r.Stage)))?"Definite":"Tentative";
    const guestsText=mostCommon(rows.map(r=>r["Guest Count"]));
    const guests=/^\d+$/.test(guestsText)?Number(guestsText):null;
    const bay=bayValue(rows),products=rows.map(r=>r["Product Name"]||"");
    const food=menus.map(r=>r["Product Name"]||"");
    const menu=mainMenu(food,products),golf=golfWindow(rows);
    const different=golf&&golf[0]&&golf[1]&&(golf[0]!==start||golf[1]!==end);
    const flags=[];
    if(!bay)flags.push("No bay assignment");
    if(!date||!start||!end)flags.push("Missing event date or time");
    if(menu.category==="Flex"&&menus.length)flags.push("Review catering package");
    events.push({
      id:"csv-"+date+"-"+encodeURIComponent(op).replace(/%/g,"").slice(0,130),
      name:opportunityTitle(op),date,start,end,
      golfStart:different?golf[0]:"",golfEnd:different?golf[1]:"",
      guests,bays:bay,category:menu.category,package:menu.package,status,
      space:findSpace(rows,menus),source:staffFile+" + "+menuFile,
      opportunityName:op,importFlags:flags
    });
  }
  const unmatched=[...catering.keys()].filter(k=>!groups.has(k));
  return {events,unmatched,matched:[...catering.keys()].filter(k=>groups.has(k)).length,staffCount:groups.size,menuCount:catering.size};
}
