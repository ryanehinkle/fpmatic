import {workbookDayPaths,parseFloorplanCell,parseFloorplanShared} from "./floorplan-sync.js?v=20261016a";

const floors=[
  {number:3,am:4,pm:9},
  {number:2,am:20,pm:25},
  {number:1,am:36,pm:41}
];
function colIndex(s){return [...s].reduce((v,c)=>v*26+c.charCodeAt(0)-64,0)}
function toCell(s){
  const m=/^([A-Z]+)(\d+)$/.exec(s);return m?{c:colIndex(m[1]),r:+m[2]}:null;
}
function fromRef(s){
  const [left,right=left]=s.split(":"),a=toCell(left),b=toCell(right);
  return a&&b?{c1:a.c,r1:a.r,c2:b.c,r2:b.r}:null;
}
function rgb(tag){
  const v=(tag.match(/\brgb="([^"]+)"/)||[])[1];
  return v?"#"+v.slice(-6):null;
}
function sheetStyles(xml){
  function tags(section,name){
    const text=(xml.match(new RegExp("<"+section+"\\b[^>]*>([\\s\\S]*?)<\\/"+section+">"))||[])[1]||"";
    return [...text.matchAll(new RegExp("<"+name+"\\b[^>]*\\/>|<"+name+"\\b[^>]*>[\\s\\S]*?<\\/"+name+">","g"))].map(m=>m[0]);
  }
  const fills=tags("fills","fill").map(s=>{
    const solid=/patternType="solid"/.test(s);
    return solid?(rgb((s.match(/<fgColor\b[^>]*\/>/)||[])[0]||"")||"transparent"):"transparent";
  });
  const borders=tags("borders","border").map(s=>{
    const obj={};
    for(const side of ["top","bottom","left","right"]){
      const fragment=(s.match(new RegExp("<"+side+"\\b[^>]*(?:\\/>|>[\\s\\S]*?<\\/"+side+">)"))||[])[0]||"";
      const weight=(fragment.match(/\bstyle="([^"]+)"/)||[])[1]||"";
      obj[side]={weight,color:rgb(fragment)||"#111111"};
    }
    return obj;
  });
  const formats=tags("cellXfs","xf").map(s=>({
    fillId:+((s.match(/\bfillId="(\d+)"/)||[])[1]||0),
    borderId:+((s.match(/\bborderId="(\d+)"/)||[])[1]||0)
  }));
  return {fills,borders,formats};
}
function styleFor(styles,index){
  const fmt=styles.formats[index]||styles.formats[0]||{fillId:0,borderId:0};
  return {fill:styles.fills[fmt.fillId]||"transparent",border:styles.borders[fmt.borderId]||{}};
}
function effectiveBorder(cells,styles,area,side){
  const rank={none:0,hair:1,dotted:1,dashed:1,thin:2,medium:3,thick:4,double:4};
  const test=[];
  for(let c=area.c1;c<=area.c2;c++)for(let r=area.r1;r<=area.r2;r++){
    if((side==="top"&&r!==area.r1)||(side==="bottom"&&r!==area.r2)||(side==="left"&&c!==area.c1)||(side==="right"&&c!==area.c2))continue;
    let cc=c,chars="";while(cc>0){cc--;chars=String.fromCharCode(65+cc%26)+chars;cc=Math.floor(cc/26)}
    const cell=cells.get(chars+r);
    if(!cell)continue;
    const border=styleFor(styles,cell.style).border[side];
    if(border?.weight)test.push(border);
  }
  test.sort((a,b)=>(rank[b.weight]||0)-(rank[a.weight]||0));
  return test[0]||null;
}
function cssBorder(b){
  if(!b)return "0";
  const w=/^(?:medium|thick|double)$/.test(b.weight)?2:1;
  return w+"px solid "+b.color;
}
function gridRegion(area,text,cells,styles){
  let c=area.c1,str="";while(c>0){c--;str=String.fromCharCode(65+c%26)+str;c=Math.floor(c/26)}
  const ref=str+area.r1,cell=cells.get(ref);
  const base=styleFor(styles,cell?.style||0);
  const edges={};
  for(const side of ["top","bottom","left","right"])edges[side]=cssBorder(effectiveBorder(cells,styles,area,side));
  return {area,text:text||"",fill:base.fill,borders:edges};
}
export async function readWorkbookPreview(blob){
  if(typeof JSZip!=="function")throw new Error("Excel ZIP library was not loaded");
  const zip=await JSZip.loadAsync(blob);
  const workbook=await zip.file("xl/workbook.xml").async("string");
  const paths=workbookDayPaths(workbook);
  const stylesheet=sheetStyles(await zip.file("xl/styles.xml").async("string"));
  const sharedFile=zip.file("xl/sharedStrings.xml");
  const shared=sharedFile?parseFloorplanShared(await sharedFile.async("string")):[];
  const days=[];
  for(let day=0;day<7;day++){
    const path=paths[day];if(!path)throw new Error("Missing day "+day+" in workbook");
    const xml=await zip.file(path).async("string"),cells=parseFloorplanCell(xml,shared);
    const covered=new Set(),regions=[];
    const merges=[...xml.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"[^>]*\/>/g)].map(x=>fromRef(x[1])).filter(Boolean);
    for(const area of merges){
      if(area.c1<5||area.c2>38)continue;
      const floor=floors.find(f=>[f.am,f.pm].some(r=>area.r1>=r&&area.r2<=r+4));
      if(!floor)continue;
      for(let r=area.r1;r<=area.r2;r++)for(let c=area.c1;c<=area.c2;c++)covered.add(c+":"+r);
      let col=area.c1,letter="";while(col){col--;letter=String.fromCharCode(65+col%26)+letter;col=Math.floor(col/26)}
      regions.push(gridRegion(area,cells.get(letter+area.r1)?.text||"",cells,stylesheet));
    }
    for(const [addr,cell] of cells){
      if(!cell.text.trim())continue;
      const p=toCell(addr);
      if(p.c<5||p.c>38||covered.has(p.c+":"+p.r))continue;
      if(!floors.some(f=>[f.am,f.pm].some(r=>p.r>=r&&p.r<=r+4)))continue;
      const area={c1:p.c,c2:p.c,r1:p.r,r2:p.r};
      regions.push(gridRegion(area,cell.text,cells,stylesheet));
    }
    days.push({regions,path});
  }
  return {days};
}
