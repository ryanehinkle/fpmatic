// Pure Salesforce opportunity-link parser. Everything stays in the browser.
function entityDecode(input){
  const named={"amp":"&","lt":"<","gt":">","quot":'"',"apos":"'","nbsp":" ","#39":"'","colon":":","sol":"/"};
  return String(input||"").replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi,(_,code)=>{
    const key=code.toLowerCase();
    if(key.startsWith("#x"))return String.fromCodePoint(parseInt(key.slice(2),16));
    if(key.startsWith("#")&&key!=="#39")return String.fromCodePoint(parseInt(key.slice(1),10));
    return named[key]??("&"+code+";");
  });
}
export function canonicalOpportunity(name){
  return String(name||"").normalize("NFKC")
    .replace(/[\u2018\u2019\u2032?](?=s\b)/gi,"'")
    .replace(/\s+/g," ").trim().toLocaleLowerCase("en-US");
}
export function shortOpportunityName(full){
  const name=String(full||"").replace(/\s*-\s*\d{1,2}[./]\d{1,2}[./]\d{2,4}\s*$/,"").trim();
  const fields=name.split(/\s+-\s+/);
  if(fields.length>1)fields.pop(); // Salesforce trailing contact.
  return fields.join(" - ")||name;
}
export function opportunityDate(full){
  const m=String(full||"").match(/(?:^|\s*-\s*)(\d{1,2})[./](\d{1,2})[./](\d{2}|\d{4})\s*$/);
  if(!m)return "";
  const year=m[3].length===2?"20"+m[3]:m[3];
  return year+"-"+m[1].padStart(2,"0")+"-"+m[2].padStart(2,"0");
}
export function safeSalesforceUrl(href){
  try{
    const address=new URL(entityDecode(href));
    if(address.protocol!=="https:"||address.username||address.password)return "";
    const host=address.hostname.toLowerCase();
    const approved=["salesforce.com","force.com","salesforce-sites.com"];
    if(!approved.some(s=>host===s||host.endsWith("."+s)))return "";
    return address.href;
  }catch{return ""}
}
export function extractSalesforceLinks(html){
  const links=[],seen=new Set();
  // Salesforce copies report cells as HTML anchors, often with nested spans.
  const anchors=String(html||"").matchAll(/<a\b([^>]*)>([\s\S]*?)<\/a\s*>/gi);
  for(const match of anchors){
    const rawHref=(match[1].match(/\bhref\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/i)||[]);
    const href=safeSalesforceUrl(rawHref[1]||rawHref[2]||rawHref[3]||"");
    if(!href)continue;
    const title=entityDecode(match[2].replace(/<br\s*\/?>/gi," ").replace(/<[^>]*>/g," "))
      .replace(/\s+/g," ").trim();
    if(!title||!opportunityDate(title))continue;
    const key=canonicalOpportunity(title);
    if(seen.has(key))continue;
    seen.add(key);
    links.push({name:title,url:href,date:opportunityDate(title),shortName:shortOpportunityName(title)});
  }
  return links;
}
export function matchingSalesforceLink(event,links){
  if(!event||!Array.isArray(links))return null;
  const full=canonicalOpportunity(event.opportunityName);
  if(full){
    const exact=links.find(x=>canonicalOpportunity(x.name)===full);
    if(exact)return exact;
  }
  const name=canonicalOpportunity(event.name);
  if(!name||!event.date)return null;
  const matches=links.filter(x=>x.date===event.date&&canonicalOpportunity(x.shortName)===name);
  return matches.length===1?matches[0]:null;
}
export function linkForWorksheetText(text,date,events,links){
  const first=String(text||"").split(/\r?\n/)[0].trim();
  if(!first||!/\b\d+\s*ppl\b/i.test(text||""))return null;
  const name=canonicalOpportunity(first);
  const matches=(events||[]).filter(e=>e.date===date&&canonicalOpportunity(e.name)===name&&e.salesforceUrl);
  const choices=new Set(matches.map(e=>e.salesforceUrl));
  if(choices.size===1)return [...choices][0];
  if(choices.size>1)return null;
  const direct=(links||[]).filter(x=>x.date===date&&canonicalOpportunity(x.shortName)===name);
  return direct.length===1?direct[0].url:null;
}
