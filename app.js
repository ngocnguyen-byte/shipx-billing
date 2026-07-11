
/* ============================================================
   ShipX Billing — unified multi-service tool (hosted, multi-user)
   Auth + shared rate cards + billing history + live sync via Supabase.
   Calculation engine below is identical to the offline build.
   ============================================================ */

const money = (n,cur="$") => n==null||isNaN(n) ? "—" : cur+Number(n).toLocaleString("en-SG",{minimumFractionDigits:2,maximumFractionDigits:2});
const num = v => { if(v==null||v==="") return null; const n=parseFloat(String(v).replace(/[^0-9.\-]/g,"")); return isNaN(n)?null:n; };
const round2 = n => Math.round((n+Number.EPSILON)*100)/100;
/* Excel formula cell with cached value (SheetJS drops bare {f}) */
const fcell = (formula,val) => ({t:'n', f:formula, v:(val==null||isNaN(val)?0:val)});
const sumBy = (lines,k) => round2(lines.reduce((s,o)=>s+(num(o[k])||0),0));
const esc = s => String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
const norm = s => String(s??"").toLowerCase().replace(/[^a-z0-9]/g,"");
const todayISO = () => new Date().toISOString().slice(0,10);
function toISO(d){ if(d instanceof Date) return d.toISOString().slice(0,10);
  if(typeof d==="number" && d>30000 && d<60000){ const dt=new Date(Date.UTC(1899,11,30)); dt.setUTCDate(dt.getUTCDate()+d); return dt.toISOString().slice(0,10);}
  const s=String(d??"").trim(); const m=s.match(/^(\d{4}-\d{2}-\d{2})/); return m?m[1]:s; }
function periodOf(iso){ const m=String(iso).match(/^(\d{4})-(\d{2})/); if(!m) return {y:"?",q:"?",m:"?"};
  const q="Q"+Math.ceil(parseInt(m[2])/3); return {y:m[1], q:m[1]+" "+q, m:m[1]+"-"+m[2]}; }

/* ---------- storage (Supabase-backed, shared + live) ----------
   In-memory caches keep the existing (synchronous) render code working;
   they are filled on boot and kept fresh by realtime subscriptions. */
let RECORDS = [];      // cache of billing_records (mapped to record objects)
let RATES   = {};      // cache: "service_id:card_id" -> rows[]
let CURRENT_USER = null;
let SB = null;         // supabase client

const rkey = id => id.indexOf(":")>=0 ? id : id+":main";
function loadRecords(){ return RECORDS; }
function loadSavedRates(id){ return RATES[rkey(id)] || null; }
function saveRatesFor(id, rows){
  if(rows==null) return;
  const k=rkey(id); RATES[k]=rows;                       // optimistic
  const [service_id, card_id="main"]=k.split(":");
  if(SB) SB.from("rate_cards").upsert({service_id, card_id, rows,
    updated_at:new Date().toISOString(), updated_by:CURRENT_USER}).then(({error})=>{
      if(error) console.error("rate save failed", error); });
}

/* ============================================================
   SERVICE DEFINITIONS
   Each service: id,name,group,tags,outputs,description,
     rate:{cols,rows,keyCol}, input:{cols:[{key,label,aliases,required}]},
     calc(rows,rate) -> {lines, review, columns, currency}
   A "line" is one billed row; must include .amount, and ideally .customer & .date & .cost
   ============================================================ */

const R2S = 1.34; // default USD->SGD (IOSS)

/* ---- Singpost ePAC rate cards (embedded) : [zone,dest,code,itemS$,kgS$] ---- */
const _SPL=[["A","Malaysia","MY",2.35,3.2],["B","Brunei","BN",2.65,9.2],["B","Hong Kong","HK",2.4,5.1],["B","Indonesia","ID",2.65,5.4],["B","Philippines","PH",2.65,7.5],["B","Taiwan","TW",2.4,7.6],["B","Thailand","TH",2.65,5.3],["B","Bangladesh","BD",3.65,28.55],["C","Bhutan","BT",2.55,4.6],["C","Cambodia","KH",2.65,8.1],["C","China","CN",2.65,6.7],["C","Fiji","FJ",4.25,12.1],["C","India","IN",2.35,8.3],["C","Kiribati","KI",2.85,41.4],["C","Laos","LA",2.65,5.6],["C","Macau","MO",4.25,9.8],["C","Maldives","MV",4.25,10.6],["C","Mongolia","MN",3.15,17.75],["C","Myanmar","MM",2.55,10.8],["C","Nepal","NP",4.25,16.1],["C","Pakistan","PK",4.25,9.8],["C","Papua New Guinea","PG",4.25,11.1],["C","Solomon Island","SB",2.55,48],["C","South Korea","KR",2.75,6.2],["C","Sri Lanka","LK",4.25,9],["C","Tuvalu","TV",2.85,41.4],["C","Vanuatu","VU",4.25,46.2],["C","Vietnam","VN",1.85,4.1],["C","Tahiti (French Polynesia)","PF",3.75,30.5],["C","Countries in Asia & Oceania","Zone C",3.75,30.5],["R","Australia","AU",5.05,14.2],["R","Japan","JP",3.8,13.35],["R","New Zealand","NZ",5.05,13.1],["R","Albania","AL",4.95,36.6],["S","Armenia","AM",2.9,33.1],["S","Austria","AT",3.05,33.8],["S","Belarus","BY",2.9,29.9],["S","Belgium","BE",6.85,17.4],["S","Bosnia Herzegovina","BA",3.25,30.3],["S","Bulgaria","BG",2.2,10.4],["S","Canada","CA",5.2,21.1],["S","Croatia","HR",2.6,34.9],["S","Cyprus","CY",3.1,27.7],["S","Czech","CZ",3.25,28.5],["S","Denmark","DK",3.1,34],["S","Estonia","EE",2.8,12.6],["S","Finland","FI",7.35,14.8],["S","France","FR",6.75,13.5],["S","Georgia","GE",2.45,13.1],["S","Germany","DE",3.05,27.35],["S","Gibraltar","GI",2.3,39.55],["S","Greece","GR",3.1,30.9],["S","Hungary","HU",2.65,33.7],["S","Iceland","IS",3.1,37.4],["S","Ireland","IE",6.85,16.55],["S","Italy","IT",6.25,12.8],["S","Jersey","JE",2.55,42.15],["S","Kazakhstan","KZ",2.75,13.25],["S","Latvia","LV",3.1,35.55],["S","Lithuania","LT",2.65,10.75],["S","Luxembourg","LU",3.1,38.65],["S","Malta","MT",3.1,32.9],["S","Netherlands","NL",4.8,11],["S","Norway","NO",5.9,10.8],["S","Poland","PL",4.35,13.1],["S","Portugal","PT",6.25,13.3],["S","Republic of Serbia","RS",4.05,19.9],["S","Romania","RO",3.1,29.8],["S","Slovakia","SK",3.1,32.85],["S","Slovenia","SI",3.1,38.65],["S","Spain","ES",6.05,14.05],["S","Sweden","SE",7.15,14.2],["S","Switzerland","CH",7,16.2],["S","Turkey","TR",2.65,31.3],["S","Ukraine","UA",2.25,10.7],["S","United Kingdom","GB",5.3,15.5],["S","Moldova","MD",5.05,38.65],["S","Azerbaijan","AZ",5.05,38.65],["S","Republic of Montenegro","ME",5.05,38.65],["S","Kyrgyzstan","KG",5.05,38.65],["S","Monaco","MC",5.05,38.65],["S","Uzbekistan","UZ",5.05,38.65],["S","Guernsey Island and Alderney","GG",5.05,38.65],["S","Countries in Europe","Zone S",5.05,38.65],["T","Algeria","DZ",5.3,14.2],["T","Argentina","AR",3.25,44.6],["T","Aruba","AW",2.2,18.5],["T","Bahrain","BH",5.15,10.7],["T","Brazil","BR",6.25,27],["T","Chile","CL",2.8,36.9],["T","Colombia","CO",3.25,42.3],["T","Curacao","CW",2.2,36.1],["T","Democratic Rep of Congo","CD",3.1,20.1],["T","Dominican Republic","DO",2.35,21.25],["T","Ecuador","EC",3.95,44.7],["T","Egypt","EG",2.35,8],["T","Eswatini","SZ",2.3,27.15],["T","Guadeloupe","GP",6.3,23.5],["T","Iran","IR",4.25,13.1],["T","Israel","IL",5.1,18.9],["T","Jordan","JO",3.85,42.35],["T","Kuwait","KW",3.25,23],["T","Martinique","MQ",6.3,23.5],["T","Mauritius","MU",3.85,24],["T","Mexico","MX",2.2,17.8],["T","Morocco","MA",2.35,11.8],["T","Oman","DM",2.5,8.3],["T","Paraguay","PY",3.85,42.35],["T","Peru","PE",3.05,36.1],["T","Reunion","RE",6.3,22.8],["T","Saudi Arabia","SA",6.1,13.8],["T","Sint Maarten","SX",2.2,36.1],["T","South Africa","ZA",3.85,25.6],["T","United Arab Emirates","AE",2.2,9.1],["T","Uruguay","UT",2.55,25.3],["T","Zimbabwe","ZW",2.35,19.1],["T","Qatar","QA",3.95,44.7],["T","Cameroon","CM",3.95,44.7],["T","Lebanon","LB",3.95,44.7],["T","Costa Rica","CR",3.95,44.7],["T","Countries in Africa & ROW","Zone T",3.95,44.7],["U","USA","US",7.2,17.7],["U","USA & Territories","Zone U",7.2,17.7],["S","Andorra","AD",5.05,38.65],["T","Bermuda","BM",3.95,44.7],["T","Bolivia","BO",3.95,44.7],["T","El Salvador","SV",3.95,44.7],["S","Faroe Islands","FO",5.05,38.65],["T","French Guiana","GF",3.95,44.7],["T","Ghana","GH",3.95,44.7],["T","Guatemala","GT",3.95,44.7],["C","New Caledonia","NC",3.75,30.5],["C","Pacific Islands (Palau)","PW",3.75,30.5],["T","Panama","PA",3.95,44.7],["C","Samoa Western","WS",3.75,30.5],["T","St. Pierre & Miquelon","PM",3.95,44.7],["T","Zambia","ZM",3.95,44.7]];
const _SPS=[["A","Malaysia","MY",2.47,3.36],["B","Brunei","BN",2.79,9.66],["B","Hong Kong","HK",2.52,5.36],["B","Indonesia","ID",2.79,5.67],["B","Philippines","PH",2.79,7.88],["B","Taiwan","TW",2.52,7.98],["B","Thailand","TH",2.79,5.57],["B","Bangladesh","BD",3.84,29.98],["C","Bhutan","BT",2.68,4.83],["C","Cambodia","KH",2.79,8.51],["C","China","CN",2.79,7.04],["C","Fiji","FJ",4.47,12.71],["C","India","IN",2.47,8.72],["C","Kiribati","KI",3,43.47],["C","Laos","LA",2.79,5.88],["C","Macau","MO",4.47,10.29],["C","Maldives","MV",4.47,11.13],["C","Mongolia","MN",3.31,18.64],["C","Myanmar","MM",2.68,11.34],["C","Nepal","NP",4.47,16.91],["C","Pakistan","PK",4.47,10.29],["C","Papua New Guinea","PG",4.47,11.66],["C","Solomon Island","SB",2.68,50.4],["C","South Korea","KR",2.89,6.51],["C","Sri Lanka","LK",4.47,9.45],["C","Tuvalu","TV",3,43.47],["C","Vanuatu","VU",4.47,48.51],["C","Vietnam","VN",1.95,4.31],["C","Tahiti (French Polynesia)","PF",3.94,32.03],["C","Countries in Asia & Oceania","Zone C",3.94,32.03],["R","Australia","AU",5.31,14.91],["R","Japan","JP",3.99,14.02],["R","New Zealand","NZ",5.31,13.76],["R","Albania","AL",5.2,38.43],["S","Armenia","AM",3.05,34.76],["S","Austria","AT",3.21,35.49],["S","Belarus","BY",3.05,31.4],["S","Belgium","BE",7.2,18.27],["S","Bosnia Herzegovina","BA",3.42,31.82],["S","Bulgaria","BG",2.31,10.92],["S","Canada","CA",5.46,22.16],["S","Croatia","HR",2.73,36.65],["S","Cyprus","CY",3.26,29.09],["S","Czech","CZ",3.42,29.93],["S","Denmark","DK",3.26,35.7],["S","Estonia","EE",2.94,13.23],["S","Finland","FI",7.72,15.54],["S","France","FR",7.09,14.18],["S","Georgia","GE",2.58,13.76],["S","Germany","DE",3.21,28.72],["S","Gibraltar","GI",2.42,41.53],["S","Greece","GR",3.26,32.45],["S","Hungary","HU",2.79,35.39],["S","Iceland","IS",3.26,39.27],["S","Ireland","IE",7.2,17.38],["S","Italy","IT",6.57,13.44],["S","Jersey","JE",2.68,44.26],["S","Kazakhstan","KZ",2.89,13.92],["S","Latvia","LV",3.26,37.33],["S","Lithuania","LT",2.79,11.29],["S","Luxembourg","LU",3.26,40.59],["S","Malta","MT",3.26,34.55],["S","Netherlands","NL",5.04,11.55],["S","Norway","NO",6.2,11.34],["S","Poland","PL",4.57,13.76],["S","Portugal","PT",6.57,13.97],["S","Republic of Serbia","RS",4.26,20.9],["S","Romania","RO",3.26,31.29],["S","Slovakia","SK",3.26,34.5],["S","Slovenia","SI",3.26,40.59],["S","Spain","ES",6.36,14.76],["S","Sweden","SE",7.51,14.91],["S","Switzerland","CH",7.35,17.01],["S","Turkey","TR",2.79,32.87],["S","Ukraine","UA",2.37,11.24],["S","United Kingdom","GB",5.57,16.28],["S","Moldova","MD",5.31,40.59],["S","Azerbaijan","AZ",5.31,40.59],["S","Republic of Montenegro","ME",5.31,40.59],["S","Kyrgyzstan","KG",5.31,40.59],["S","Monaco","MC",5.31,40.59],["S","Uzbekistan","UZ",5.31,40.59],["S","Guernsey Island and Alderney","GG",5.31,40.59],["S","Countries in Europe","Zone S",5.31,40.59],["T","Algeria","DZ",5.57,14.91],["T","Argentina","AR",3.42,46.83],["T","Aruba","AW",2.31,19.43],["T","Bahrain","BH",5.41,11.24],["T","Brazil","BR",6.57,28.35],["T","Chile","CL",2.94,38.75],["T","Colombia","CO",3.42,44.42],["T","Curacao","CW",2.31,37.91],["T","Democratic Rep of Congo","CD",3.26,21.11],["T","Dominican Republic","DO",2.47,22.32],["T","Ecuador","EC",4.15,46.94],["T","Egypt","EG",2.47,8.4],["T","Eswatini","SZ",2.42,28.51],["T","Guadeloupe","GP",6.62,24.68],["T","Iran","IR",4.47,13.76],["T","Israel","IL",5.36,19.85],["T","Jordan","JO",4.05,44.47],["T","Kuwait","KW",3.42,24.15],["T","Martinique","MQ",6.62,24.68],["T","Mauritius","MU",4.05,25.2],["T","Mexico","MX",2.31,18.69],["T","Morocco","MA",2.47,12.39],["T","Oman","DM",2.63,8.72],["T","Paraguay","PY",4.05,44.47],["T","Peru","PE",3.21,37.91],["T","Reunion","RE",6.62,23.94],["T","Saudi Arabia","SA",6.41,14.49],["T","Sint Maarten","SX",2.31,37.91],["T","South Africa","ZA",4.05,26.88],["T","United Arab Emirates","AE",2.31,9.56],["T","Uruguay","UT",2.68,26.57],["T","Zimbabwe","ZW",2.47,20.06],["T","Qatar","QA",4.15,46.94],["T","Cameroon","CM",4.15,46.94],["T","Lebanon","LB",4.15,46.94],["T","Costa Rica","CR",4.15,46.94],["T","Countries in Africa & ROW","Zone T",4.15,46.94],["U","USA","US",7.56,18.59],["U","USA & Territories","Zone U",7.56,18.59],["S","Andorra","AD",5.31,40.59],["T","Bermuda","BM",4.15,46.94],["T","Bolivia","BO",4.15,46.94],["T","El Salvador","SV",4.15,46.94],["S","Faroe Islands","FO",5.31,40.59],["T","French Guiana","GF",4.15,46.94],["T","Ghana","GH",4.15,46.94],["T","Guatemala","GT",4.15,46.94],["C","New Caledonia","NC",3.94,32.03],["C","Pacific Islands (Palau)","PW",3.94,32.03],["T","Panama","PA",4.15,46.94],["C","Samoa Western","WS",3.94,32.03],["T","St. Pierre & Miquelon","PM",4.15,46.94],["T","Zambia","ZM",4.15,46.94]];
const SP_CARD_COLS=[{key:"zone",label:"Zone"},{key:"dest",label:"Destination"},{key:"code",label:"Code"},{key:"item",label:"Item S$",num:true},{key:"kg",label:"Kg S$",num:true}];
const _mkcard=a=>a.map(r=>({zone:r[0],dest:r[1],code:r[2],item:r[3],kg:r[4]}));
const SP_LINS_CARD=_mkcard(_SPL), SP_SGL_CARD=_mkcard(_SPS);

const SERVICES = [
/* ---------------- CCL ---------------- */
{
  id:"ccl", name:"CCL — Clearance", group:"Verified", tags:["Billing","Recon"], status:"ready",
  description:"Clearance = Weight × rate(consignee). Total = Clearance + Permit(consignee).",
  rate:{ keyCol:"consignee",
    cols:[{key:"consignee",label:"Consignee"},{key:"permit",label:"Permit (SGD)",num:true},
          {key:"clearance",label:"Clearance /kg",num:true},{key:"cost",label:"Cost (SGD) — optional",num:true}],
    rows:[{consignee:"Quantium",permit:58,clearance:0.3,cost:null},
          {consignee:"BPost",permit:53,clearance:0,cost:null},
          {consignee:"Uniglobe",permit:53,clearance:0.3,cost:null}] },
  input:{ cols:[
    {key:"date",label:"DATE",aliases:["date"],required:true},
    {key:"ship",label:"SHIPMENT",aliases:["shipment"],required:true},
    {key:"pcs",label:"PCS",aliases:["pcs"],required:false},
    {key:"weight",label:"Weight",aliases:["weight","actualweight"],required:true},
    {key:"cons",label:"Consignee",aliases:["consignee"],required:true} ]},
  calc(rows,rate){
    const map={}; rate.rows.forEach(r=>map[norm(r.consignee)]={...r});
    const lines=[],review=[];
    rows.forEach(r=>{
      const key=norm(r.cons); const rc=map[key]; const wt=num(r.weight)||0;
      if(!r.cons){ return; }
      if(!rc){ review.push({...r,reason:"Unknown consignee — no rate"}); return; }
      const clearance=round2(wt*(num(rc.clearance)||0));
      const permit=num(rc.permit)||0;
      const total=round2(clearance+permit);
      const cost=rc.cost!=null?num(rc.cost):null;
      lines.push({date:toISO(r.date),ship:r.ship,pcs:r.pcs,weight:wt,customer:r.cons,
        clearance,permit,amount:total,cost});
    });
    return {lines,review,currency:"$",
      columns:[{k:"date",l:"Date"},{k:"ship",l:"Shipment"},{k:"pcs",l:"PCS",num:true},
        {k:"weight",l:"Weight",num:true},{k:"customer",l:"Consignee"},
        {k:"clearance",l:"Clearance",num:true,money:true},{k:"permit",l:"Permit",num:true,money:true},
        {k:"amount",l:"Total (SGD)",num:true,money:true,tot:true}]};
  }
},
/* ---------------- Dom SG ---------------- */
{
  id:"domsg", name:"Dom SG", group:"Verified", tags:["Billing"], status:"ready",
  description:"Domestic Singapore delivery. Price by service code and weight tier (0-3kg, 3.01-5kg, then per-kg).",
  rate:{ keyCol:"service",
    cols:[{key:"service",label:"Service Code"},{key:"t1",label:"0–3 KG",num:true},
          {key:"t2",label:"3.01–5 KG",num:true},{key:"extra",label:"Each KG >5",num:true},{key:"cost",label:"Cost (SGD) — optional",num:true}],
    rows:[{service:"DOM123",t1:2.9,t2:3.6,extra:0.5,cost:null},
          {service:"DOMN",t1:3.4,t2:3.7,extra:0.5,cost:null}] },
  input:{ cols:[
    {key:"date",label:"Shipment Date",aliases:["shipmentdate","date"],required:true},
    {key:"ref",label:"Ref Num",aliases:["refnum","ref"],required:true},
    {key:"track",label:"Tracking Number",aliases:["trackingnumber"],required:false},
    {key:"consignee",label:"Consignee Name",aliases:["consigneename"],required:false},
    {key:"svc",label:"Service Code",aliases:["servicecode"],required:true},
    {key:"weight",label:"Actual Weight",aliases:["actualweight","acutalweight","weight"],required:true} ]},
  calc(rows,rate){
    const map={}; rate.rows.forEach(r=>map[norm(r.service)]={...r});
    const lines=[],review=[];
    rows.forEach(r=>{
      const rc=map[norm(r.svc)]; const wt=num(r.weight)||0;
      if(!rc){ review.push({...r,reason:"Unknown service code — no rate"}); return; }
      let price; const t1=num(rc.t1)||0,t2=num(rc.t2)||0,ex=num(rc.extra)||0;
      if(wt<=3) price=t1; else if(wt<=5) price=t2; else price=round2(t2+Math.ceil(wt-5)*ex);
      lines.push({date:toISO(r.date),ref:r.ref,track:r.track,customer:r.svc,weight:wt,
        amount:round2(price),cost:rc.cost!=null?num(rc.cost):null});
    });
    return {lines,review,currency:"$",
      columns:[{k:"date",l:"Date"},{k:"ref",l:"Ref"},{k:"track",l:"Tracking"},
        {k:"customer",l:"Service"},{k:"weight",l:"Weight",num:true},
        {k:"amount",l:"Price (SGD)",num:true,money:true,tot:true}]};
  }
},
/* ---------------- IOSS ---------------- */
{
  id:"ioss", name:"IOSS", group:"Verified", tags:["Billing"], status:"ready",
  description:"Per-shipment IOSS fee by customer account (USD → SGD). Fee × count per customer.",
  fx:R2S,
  rate:{ keyCol:"account",
    cols:[{key:"account",label:"Customer Account"},{key:"fee",label:"Fee (USD)",num:true}],
    rows:[{account:"SINLOP_001",fee:0.25},{account:"SINSGL_001",fee:0.4},{account:"SINSUP_001",fee:0.25},
          {account:"SINUTT_001",fee:0.25},{account:"SINXPB_002",fee:0.25},{account:"SINXPB_003",fee:0.25},
          {account:"SINUDG_001",fee:0.4},{account:"SINSUP_003",fee:0.25},{account:"SINMIX_001",fee:0.25},
          {account:"SINJSP_001",fee:0.25},{account:"SINUTT_003",fee:0.25}] },
  input:{ cols:[
    {key:"date",label:"Date Shipment Created",aliases:["dateshipmentcreated","date"],required:true},
    {key:"acct",label:"Customer Account Number",aliases:["customeraccountnumber","customeraccount"],required:true},
    {key:"track",label:"Tracking Number",aliases:["trackingnumber"],required:false},
    {key:"country",label:"Country",aliases:["country"],required:false} ]},
  calc(rows,rate){
    const fx=this.fx||R2S;
    const map={}; rate.rows.forEach(r=>map[norm(r.account)]=num(r.fee));
    const lines=[],review=[];
    rows.forEach(r=>{
      const fee=map[norm(r.acct)];
      if(fee==null){ review.push({...r,reason:"Unknown customer account — no IOSS fee"}); return; }
      const sgd=round2(fee*fx);
      lines.push({date:toISO(r.date),track:r.track,customer:r.acct,country:r.country,
        feeUsd:fee,amount:sgd,cost:null});
    });
    return {lines,review,currency:"$",note:`Converted at 1 USD = ${fx} SGD`,
      columns:[{k:"date",l:"Date"},{k:"track",l:"Tracking"},{k:"customer",l:"Account"},
        {k:"country",l:"Country"},{k:"feeUsd",l:"Fee (USD)",num:true},
        {k:"amount",l:"Billing (SGD)",num:true,money:true,tot:true}]};
  }
},
/* ---------------- Linehaul ---------------- */
{
  id:"linehaul", name:"Linehaul (bypass bag)", group:"Verified", tags:["Billing","GP"], status:"ready",
  description:"Linehaul = Chargeable Weight × rate per kg (by destination). Cost per kg gives GP.",
  rate:{ keyCol:"dest",
    cols:[{key:"dest",label:"Dest Code"},{key:"name",label:"Dest Name"},
          {key:"rate",label:"Rate /kg (SGD)",num:true},{key:"cost",label:"Cost /kg (SGD)",num:true}],
    rows:[{dest:"AT",name:"Austria",rate:5.2,cost:5.2},{dest:"BE",name:"Belgium",rate:5.2,cost:5.2},
          {dest:"DE",name:"Germany",rate:5.2,cost:5.2},{dest:"FR",name:"France",rate:5.2,cost:5.2},
          {dest:"NL",name:"Netherlands",rate:5.2,cost:5.2}] },
  input:{ cols:[
    {key:"date",label:"Date of Dispatch",aliases:["dateofdispatch","date"],required:true},
    {key:"cn35",label:"CN35",aliases:["cn35"],required:false},
    {key:"actual",label:"Actual Weight",aliases:["actualweightkg","actualweight"],required:false},
    {key:"charge",label:"Weight to Charge (KG)",aliases:["weighttochargekg","weighttocharge","chargeableweight"],required:true},
    {key:"dest",label:"Destination",aliases:["destination","destcode"],required:true},
    {key:"rateIn",label:"Rate p/kg",aliases:["ratepkg","rate"],required:false} ]},
  calc(rows,rate){
    const map={}; rate.rows.forEach(r=>map[norm(r.dest)]={...r});
    const lines=[],review=[];
    rows.forEach(r=>{
      const rc=map[norm(r.dest)]; const wt=num(r.charge)||0;
      // fall back to rate present in the input file if no card entry
      const rateUse = rc? num(rc.rate) : num(r.rateIn);
      if(rateUse==null){ review.push({...r,reason:"No rate for destination "+(r.dest||"?")}); return; }
      const amount=round2(wt*rateUse);
      const cost = rc && rc.cost!=null ? round2(wt*num(rc.cost)) : null;
      lines.push({date:toISO(r.date),cn35:r.cn35,customer:r.dest,weight:wt,rate:rateUse,amount,cost});
    });
    return {lines,review,currency:"$",
      columns:[{k:"date",l:"Date"},{k:"cn35",l:"CN35"},{k:"customer",l:"Dest"},
        {k:"weight",l:"Chg Wt",num:true},{k:"rate",l:"Rate/kg",num:true,money:true},
        {k:"amount",l:"Billing (SGD)",num:true,money:true,tot:true}]};
  }
},
/* ---------------- Pickup Surcharge (generator) ---------------- */
{
  id:"pickup", name:"Pickup Surcharge", group:"Verified", tags:["Billing","No input"], status:"ready",
  generator:true,
  description:"No input file. Pickups every Monday & Thursday. Two pickups per day at a fixed price.",
  rate:{ keyCol:"vendor",
    cols:[{key:"vendor",label:"Pickup"},{key:"price",label:"Price (SGD)",num:true}],
    rows:[{vendor:"Anglers (1st pickup)",price:15},{vendor:"JustShip (2nd pickup)",price:15}] },
  genFields:[{key:"month",label:"Billing month",type:"month"}],
  generate(rate,opts){
    const month=opts.month||todayISO().slice(0,7);
    const [y,m]=month.split("-").map(Number);
    const days=new Date(y,m,0).getDate();
    const lines=[];
    for(let d=1; d<=days; d++){
      const dt=new Date(y,m-1,d); const wd=dt.getDay(); // 1=Mon 4=Thu
      if(wd===1||wd===4){
        const iso=`${month}-${String(d).padStart(2,"0")}`;
        rate.rows.forEach(r=>lines.push({date:iso,customer:r.vendor,amount:round2(num(r.price)||0),cost:null}));
      }
    }
    return {lines,review:[],currency:"$",
      columns:[{k:"date",l:"Pickup Date"},{k:"customer",l:"Pickup"},{k:"amount",l:"Price (SGD)",num:true,money:true,tot:true}]};
  }
},
/* ---------------- AME (Postal) ---------------- */
{
  id:"ame", name:"AME (Postal)", group:"Rate-card", tags:["Billing","Recon","GP"], status:"beta",
  description:"Per-shipment postal billing by destination account. Weight-slab rate cards live per account (SINUDG, SINSGL, SINKLA, SINJSP, SINKRS, SINAGT). Simplified here to PC + KG-rate; upload the full slab card to refine.",
  rate:{ keyCol:"account",
    cols:[{key:"account",label:"Customer Account"},{key:"pc",label:"PC Rate (SGD)",num:true},
          {key:"kg",label:"KG Rate (SGD)",num:true},{key:"cost",label:"Cost (SGD) — optional",num:true}],
    rows:[{account:"SINUDG_001",pc:0,kg:0,cost:null},{account:"SINSGL_001",pc:0,kg:0,cost:null},
          {account:"SINKLA_001",pc:0,kg:0,cost:null},{account:"SINJSP_001",pc:0,kg:0,cost:null},
          {account:"SINKRS_001",pc:0,kg:0,cost:null},{account:"SINAGT_001",pc:0,kg:0,cost:null}] },
  input:{ cols:[
    {key:"track",label:"Shipment Tracking ID",aliases:["shipmenttrackingid","shipmenttrackingid"],required:true},
    {key:"weight",label:"Actual Weight",aliases:["actualweight","weight"],required:true},
    {key:"dest",label:"Destination Country Code",aliases:["destinationcountrycode","destination"],required:false},
    {key:"acct",label:"Customer Account",aliases:["customeraccount","customeraccountnumber"],required:true} ]},
  calc(rows,rate){
    const map={}; rate.rows.forEach(r=>map[norm(r.account)]={...r});
    const lines=[],review=[];
    rows.forEach(r=>{
      const rc=map[norm(r.acct)]; const wt=num(r.weight)||0;
      if(!rc){ review.push({...r,reason:"Unknown account — no rate card"}); return; }
      const pc=num(rc.pc)||0, kg=num(rc.kg)||0;
      const amount=round2(pc + wt*kg);
      if(amount===0){ review.push({...r,reason:"Rate is 0 — upload the weight-slab card for "+r.acct}); }
      lines.push({track:r.track,customer:r.acct,dest:r.dest,weight:wt,amount,
        cost:rc.cost!=null?num(rc.cost):null});
    });
    return {lines,review,currency:"$",
      columns:[{k:"track",l:"Tracking"},{k:"customer",l:"Account"},{k:"dest",l:"Dest"},
        {k:"weight",l:"Weight",num:true},{k:"amount",l:"Billing (SGD)",num:true,money:true,tot:true}]};
  }
},
/* ---------------- Singpost D&T ---------------- */
{
  id:"sp_dt", name:"Singpost D&T", group:"Rate-card", tags:["Billing","Count"], status:"ready",
  description:"Duties & Taxes from Linscomm. Total = Amount SGD + Admin (Admin = 3% of Amount SGD). Upload one or many Linscomm D&T files (sheet 'Amilo') — they combine into one billing.",
  rate:{ keyCol:"k", cols:[{key:"k",label:"Setting"},{key:"v",label:"Value",num:true}],
    rows:[{k:"Admin %",v:3}] },
  input:{ cols:[
    {key:"track",label:"Tracking Number",aliases:["trackingnumber","tracking"],required:true},
    {key:"usd",label:"Amount USD",aliases:["amountusd"],required:false},
    {key:"sgd",label:"Amount SGD",aliases:["amountsgd"],required:true} ]},
  calc(rows,rate){
    const pct=(num(rate.rows[0]&&rate.rows[0].v)||0)/100;
    const lines=[],review=[];
    rows.forEach(r=>{
      const sgd=num(r.sgd);
      if(sgd==null){ review.push({...r,reason:"Missing Amount SGD"}); return; }
      const admin=round2(sgd*pct), total=round2(sgd+admin);
      lines.push({track:r.track,customer:"Singpost D&T",usd:num(r.usd),sgd:round2(sgd),admin,amount:total,cost:round2(sgd)});
    });
    return {lines,review,currency:"$",
      columns:[{k:"track",l:"Tracking Number"},{k:"usd",l:"Amount USD",num:true},
        {k:"sgd",l:"Amount SGD",num:true,money:true},{k:"admin",l:"Admin SGD",num:true,money:true},
        {k:"amount",l:"Total SGD",num:true,money:true,tot:true}]};
  },
  buildWorkbook(res,st){
    const pct=(num(st.rate.rows[0]&&st.rate.rows[0].v)||3);
    const N=res.lines.length;
    const aoa=[[null,null,null,null,fcell("SUM(E3:E"+(N+2)+")",sumBy(res.lines,"amount"))],
      ["Tracking Number","Amount USD","Amount SGD","Admin SGD","Total SGD"]];
    res.lines.forEach((o,i)=>{ const R=i+3;
      aoa.push([o.track,o.usd,o.sgd,fcell("C"+R+"*"+pct+"%",o.admin),fcell("D"+R+"+C"+R,o.amount)]); });
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),"SGL");
    return {wb, name:"SGL - Singpost D&T"};
  }
},
/* ---------------- Singpost Postage ---------------- */
{
  id:"sp_post", name:"Singpost Postage", group:"Rate-card", tags:["Recon","Billing","GP"], status:"ready",
  description:"Reconcile Linscomm's billing (did they charge the correct ePAC rate?) and bill SG Link. Upload the Linscomm billing export: Date, Docket No, Mode, Schm, Weight, Quantity, Rate, Postage S$, Country. Two rate cards below: Linscomm (cost) and SG Link (sell).",
  rateCards:[
    {id:"lins", label:"Rate from Linscomm (cost)", keyCol:"code", cols:SP_CARD_COLS, rows:SP_LINS_CARD},
    {id:"sgl", label:"Rate to SG Link (sell)", keyCol:"code", cols:SP_CARD_COLS, rows:SP_SGL_CARD}
  ],
  input:{ cols:[
    {key:"date",label:"Date",aliases:["dateordinary","date"],required:false},
    {key:"docket",label:"Docket No",aliases:["docketno","docket"],required:false},
    {key:"mode",label:"Mode",aliases:["mode"],required:false},
    {key:"schm",label:"Schm",aliases:["schm","scheme"],required:false},
    {key:"weight",label:"Weight (Kg)",aliases:["weightkg","weight"],required:true},
    {key:"qty",label:"Quantity",aliases:["quantity","qty"],required:true},
    {key:"rate",label:"Rate",aliases:["rate"],required:false},
    {key:"postage",label:"Postage S$",aliases:["postages","postage"],required:false},
    {key:"country",label:"Country",aliases:["country"],required:true} ]},
  calc(rows,rate,st){
    const lins={},sgl={},byName={};
    st.rateCards.lins.rows.forEach(r=>{ lins[norm(r.code)]=r; if(r.dest) byName[norm(r.dest)]=r.code; });
    st.rateCards.sgl.rows.forEach(r=>{ sgl[norm(r.code)]=r; });
    const lines=[],review=[]; let billedTot=0,expTot=0,sglTot=0,qtyTot=0,wtTot=0;
    rows.forEach(r=>{
      const country=r.country, qty=num(r.qty)||0, wt=num(r.weight)||0, billed=num(r.postage);
      const code=byName[norm(country)];
      if(!code){ review.push({...r,reason:"Country not in rate card: "+(country||"(blank)")}); return; }
      const lc=lins[norm(code)], sc=sgl[norm(code)];
      if(!sc){ review.push({...r,reason:"No SG Link rate for code "+code}); return; }
      const expected=round2((lc.item||0)*qty+(lc.kg||0)*wt);
      const sgAmt=round2((sc.item||0)*qty+(sc.kg||0)*wt);
      billedTot+=billed||0; expTot+=expected; sglTot+=sgAmt; qtyTot+=qty; wtTot+=wt;
      lines.push({date:toISO(r.date),docket:r.docket,customer:country,country,code,weight:wt,qty,
        billed:billed!=null?round2(billed):null,expected,amount:sgAmt,cost:billed!=null?round2(billed):expected});
    });
    const variance=round2(billedTot-expTot), margin=round2(sglTot-billedTot);
    const recon={ title:"Reconciliation with vendor (Linscomm)",
      metrics:[
        {label:"Billed line items",value:lines.length},
        {label:"Total articles",value:qtyTot},
        {label:"Total weight (kg)",value:round2(wtTot)},
        {label:"Linscomm billed postage (S$)",value:round2(billedTot),money:true},
        {label:"Rate-card expected (S$)",value:round2(expTot),money:true},
        {label:"Rate variance (S$)",value:variance,money:true},
        {label:"Billing to SG Link (S$)",value:round2(sglTot),money:true},
        {label:"SG Link margin / GP (S$)",value:margin,money:true}
      ],
      verdict: Math.abs(variance)<1
        ? "✓ Linscomm applied the correct ePAC rate on all "+lines.length+" billed lines (variance S$"+variance.toFixed(2)+" = rounding only)."
        : "⚠ Rate variance S$"+variance.toFixed(2)+" across "+lines.length+" lines — check the Diff_Postage column; Linscomm may have mis-rated some lines."
    };
    return {lines,review,currency:"$",recon,
      columns:[{k:"date",l:"Date"},{k:"docket",l:"Docket"},{k:"country",l:"Country"},{k:"code",l:"Code"},
        {k:"weight",l:"Weight",num:true},{k:"qty",l:"Qty",num:true},
        {k:"billed",l:"Linscomm S$",num:true,money:true},{k:"expected",l:"Expected S$",num:true,money:true},
        {k:"amount",l:"SG Link S$",num:true,money:true,tot:true}]};
  },
  buildWorkbook(res,st){ // Output 2 — Billing to SG Link, live VLOOKUP on SG Link card
    const N=res.lines.length;
    const aoa=[[null,null,null,null,fcell("SUM(E3:E"+(N+2)+")",sumBy(res.lines,"weight")),fcell("SUM(F3:F"+(N+2)+")",sumBy(res.lines,"qty")),null,null,null,fcell("SUM(J3:J"+(N+2)+")",sumBy(res.lines,"amount"))],
      ["Date\nOrdinary","Docket No","Mode","Schm","Weight (Kg)","Quantity","Rate","Country","Country Code","Postage to SG Link (SGD)"]];
    res.lines.forEach((o,i)=>{ const R=i+3;
      aoa.push([o.date,o.docket,"AIR","EP",o.weight,o.qty,"CNT",o.country,o.code,
        fcell("VLOOKUP(I"+R+",'Rate to SG Link'!$C:$E,2,0)*F"+R+"+VLOOKUP(I"+R+",'Rate to SG Link'!$C:$E,3,0)*E"+R, o.amount)]); });
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),"Jun");
    const sg=[["SG LINK RATE CARD"],["Zone","Destination","Code","Item S$","Kg S$"]];
    st.rateCards.sgl.rows.forEach(r=>sg.push([r.zone,r.dest,r.code,r.item,r.kg]));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sg),"Rate to SG Link");
    return {wb, name:"Singpost Postage Billing_SGL"};
  },
  buildRecon(res,st){ // Output 1 — Reconciliation, live VLOOKUP on Linscomm card
    const N=res.lines.length; const wb=XLSX.utils.book_new();
    const sum=[["Linscomm SingPost Reconciliation & SG Link Billing"],[]];
    res.recon.metrics.forEach(x=>sum.push([x.label,x.value]));
    sum.push([]); sum.push(["VERDICT"]); sum.push([res.recon.verdict]);
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sum),"Reconciliation Summary");
    const aoa=[[null,null,null,null,fcell("SUM(E3:E"+(N+2)+")",sumBy(res.lines,"weight")),fcell("SUM(F3:F"+(N+2)+")",sumBy(res.lines,"qty")),null,fcell("SUM(H3:H"+(N+2)+")",sumBy(res.lines,"billed")),null,null,fcell("SUM(K3:K"+(N+2)+")",sumBy(res.lines,"expected")),fcell("SUM(L3:L"+(N+2)+")",round2(sumBy(res.lines,"billed")-sumBy(res.lines,"expected")))],
      ["Date\nOrdinary","Docket No","Mode","Schm","Weight (Kg)","Quantity","Rate","Postage S$","Country","Country Code","Postage - From Rate","Diff_Postage"]];
    res.lines.forEach((o,i)=>{ const R=i+3;
      aoa.push([o.date,o.docket,"AIR","EP",o.weight,o.qty,"CNT",o.billed,o.country,o.code,
        fcell("VLOOKUP(J"+R+",'Rate from Linscomm'!$C:$E,2,0)*F"+R+"+VLOOKUP(J"+R+",'Rate from Linscomm'!$C:$E,3,0)*E"+R, o.expected),
        fcell("K"+R+"-H"+R, round2((o.expected||0)-(o.billed||0)))]); });
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),"Billing Output");
    const li=[["RATES FROM LINS COMMUNICATION"],["Zone","Destination","Code","Item S$","Kg S$"]];
    st.rateCards.lins.rows.forEach(r=>li.push([r.zone,r.dest,r.code,r.item,r.kg]));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(li),"Rate from Linscomm");
    return {wb, name:"Reconciliation Linscomm Singpost"};
  }
},
/* ---------------- FedEx ---------------- */
{
  id:"fedex", name:"FedEx", group:"Complex", tags:["Billing","Recon","GP","Multi-cust"], status:"beta",
  description:"Multi-customer zone/weight billing. Full engine uses year rate cards + cost cards + customer mapping. This module validates the input, maps customers, and flags unknown-customer / missing-rate rows for review; wire the amount column or upload the rate cards to bill.",
  rate:{ keyCol:"account",
    cols:[{key:"account",label:"Bill-to Account / Ref"},{key:"customer",label:"Customer"},{key:"markup",label:"Markup %",num:true}],
    rows:[{account:"203937214",customer:"Amilo MY",markup:0},
          {account:"SINBWS_001",customer:"Brainworks",markup:0}] },
  input:{ cols:[
    {key:"awb",label:"Air Waybill Number",aliases:["airwaybillnumber","awb","trackingno"],required:true},
    {key:"acct",label:"Bill To / Reference",aliases:["billto","shipperreference1","billtoaccountnumber","account"],required:false},
    {key:"date",label:"Ship Date",aliases:["shipdateformatted","shipdate","date"],required:false},
    {key:"amount",label:"Amount / Original Amount Due",aliases:["originalamountdue","amount","totalstandardcharges"],required:false} ]},
  calc(rows,rate){
    const map={}; rate.rows.forEach(r=>map[norm(r.account)]={...r});
    const lines=[],review=[];
    rows.forEach(r=>{
      const rc=map[norm(r.acct)];
      const cust = rc? rc.customer : null;
      if(!cust){ review.push({...r,reason:"Unknown customer — map "+(r.acct||"(blank)")+" in the rate card"}); return; }
      const cost=num(r.amount);
      if(cost==null){ review.push({...r,customer:cust,reason:"No amount / rate — needs rate-card lookup"}); return; }
      const markup=(num(rc.markup)||0)/100;
      lines.push({awb:r.awb,date:toISO(r.date),customer:cust,amount:round2(cost*(1+markup)),cost:round2(cost)});
    });
    return {lines,review,currency:"$",
      columns:[{k:"awb",l:"AWB"},{k:"date",l:"Date"},{k:"customer",l:"Customer"},
        {k:"cost",l:"Cost",num:true,money:true},{k:"amount",l:"Billing (SGD)",num:true,money:true,tot:true}]};
  }
},
];

const SVC = Object.fromEntries(SERVICES.map(s=>[s.id,s]));

/* ============================================================
   APP STATE + ROUTER
   ============================================================ */
let route={view:"dashboard",id:null};
const state={}; // per service working state

function getRate(svc){
  if(!svc.rate) return null;
  const saved=loadSavedRates(svc.id);
  if(saved) return {rows:saved.map(r=>({...r}))};
  return {rows:svc.rate.rows.map(r=>({...r}))};
}
function getRateCards(svc){
  if(!svc.rateCards) return null;
  const o={};
  svc.rateCards.forEach(c=>{ const saved=loadSavedRates(svc.id+":"+c.id);
    o[c.id]={rows:(saved||c.rows).map(r=>({...r}))}; });
  return o;
}

/* ---------- Sidebar ---------- */
function renderNav(){
  const groups={};
  SERVICES.forEach(s=>{ (groups[s.group]=groups[s.group]||[]).push(s); });
  let h=`<div class="nav-item ${route.view==='dashboard'?'active':''}" onclick="go('dashboard')"><span class="ico">▚</span> Dashboard</div>`;
  h+=`<div class="nav-item ${route.view==='records'?'active':''}" onclick="go('records')"><span class="ico">≣</span> Records</div>`;
  for(const g of ["Verified","Rate-card","Complex"]){
    if(!groups[g]) continue;
    h+=`<div class="nav-group">${g}</div>`;
    groups[g].forEach(s=>{
      const act=route.view==='service'&&route.id===s.id?'active':'';
      const ready=s.status==='ready'?'ready':'';
      h+=`<div class="nav-item ${act} ${ready}" onclick="go('service','${s.id}')"><span class="ico">▸</span> ${esc(s.name)}<span class="dot"></span></div>`;
    });
  }
  document.getElementById("nav").innerHTML=h;
}
function go(view,id){ route={view,id:id||null}; render(); window.scrollTo(0,0); }

/* ============================================================
   RENDER
   ============================================================ */
function render(){
  renderNav();
  const v=document.getElementById("view"), ta=document.getElementById("topActions");
  ta.innerHTML="";
  if(route.view==="dashboard"){ setTitle("Dashboard","Overview"); renderDashboard(v); }
  else if(route.view==="records"){ setTitle("Records","History"); renderRecords(v); }
  else if(route.view==="service"){ renderService(v, SVC[route.id]); }
}
function setTitle(t,crumb){ document.getElementById("pageTitle").textContent=t; document.getElementById("crumb").textContent=crumb||t; }

/* ---------------- SERVICE VIEW ---------------- */
function renderService(v, svc){
  setTitle(svc.name, svc.group+" service");
  if(!state[svc.id]) state[svc.id]={rate:getRate(svc), rateCards:getRateCards(svc), result:null, files:[], opts:{}, adj:[]};
  const st=state[svc.id];

  const statusTag = svc.status==="ready"
    ? `<span class="tag green">Verified</span>` : `<span class="tag amber">Beta · confirm mapping</span>`;

  let h=`<div class="banner info" style="margin-top:0">${esc(svc.description)}</div>`;

  /* Rate card card(s) */
  if(svc.rateCards){
    svc.rateCards.forEach(c=>{
      h+=`<div class="card"><div class="flexhead">
          <div><div class="step">Rate card</div><h3>${esc(c.label)} ${statusTag}</h3>
          <p class="sub">${c.rows.length} rows. Edit inline, add rows, or load a file. Saved to this browser.</p></div>
          <div style="display:flex;gap:8px">
            <button class="ghost sm" onclick="addRateRowN('${svc.id}','${c.id}')">+ Add row</button>
            <button class="subtle sm" onclick="rateFileInputN('${svc.id}','${c.id}')">Load file…</button>
            <button class="subtle sm" onclick="resetRateN('${svc.id}','${c.id}')">Reset</button>
          </div></div>
          <div class="tbl-scroll"><table id="rateTbl_${svc.id}_${c.id}"></table></div>
        </div>`;
    });
  } else if(svc.rate){
    h+=`<div class="card"><div class="flexhead">
        <div><div class="step">Rate card</div><h3>${esc(svc.name)} rates ${statusTag}</h3>
        <p class="sub">Edit any value inline, add rows, or load a rate-card file. Saved to this browser.</p></div>
        <div style="display:flex;gap:8px">
          <button class="ghost sm" onclick="addRateRow('${svc.id}')">+ Add row</button>
          <button class="subtle sm" onclick="rateFileInput('${svc.id}')">Load file…</button>
          <button class="subtle sm" onclick="resetRate('${svc.id}')">Reset</button>
        </div></div>
        <div class="tbl-scroll" style="max-height:none"><table id="rateTbl_${svc.id}"></table></div>
      </div>`;
  }

  /* Step 1 — input (or generator) */
  if(svc.generator){
    h+=`<div class="card"><div class="step">Step 1 · Options</div><h3>Generate monthly billing</h3>
        <p class="sub">No input file needed. Choose the month and generate.</p>
        <div class="filters">`;
    svc.genFields.forEach(f=>{
      const val=st.opts[f.key]||todayISO().slice(0,7);
      h+=`<label>${esc(f.label)} <input type="${f.type}" value="${val}" onchange="setOpt('${svc.id}','${f.key}',this.value)"></label>`;
    });
    h+=`<button onclick="runGenerator('${svc.id}')">Generate</button></div></div>`;
  } else {
    h+=`<div class="card"><div class="step">Step 1 · Upload input</div><h3>Billing input file(s)</h3>
        <p class="sub">Expected columns: ${svc.input.cols.map(c=>`<code>${esc(c.label)}</code>`).join(" ")}. Extra columns are ignored. You can drop several files at once.</p>
        <div class="drop" id="drop_${svc.id}"><p style="margin:0"><b>Drop Excel/CSV files here</b> or click to choose</p>
          <small>.xlsx · .xls · .csv &nbsp;·&nbsp; files never leave this computer</small></div>
        <div id="filepills_${svc.id}" style="margin-top:10px"></div>
        <div id="validation_${svc.id}"></div></div>`;
  }

  /* Manual adjustments */
  h+=`<div class="card"><div class="flexhead">
      <div><div class="step">Optional · Manual adjustments</div><h3>Add non-input lines</h3>
      <p class="sub">For one-off charges that aren't in the input file (e.g. CCL "returns" pickup). These are added to the billing and reconcile the total.</p></div>
      <button class="ghost sm" onclick="addAdj('${svc.id}')">+ Add adjustment</button></div>
      <div class="tbl-scroll" style="max-height:none" id="adjWrap_${svc.id}"><table id="adjTbl_${svc.id}"></table></div></div>`;

  /* Results */
  h+=`<div id="results_${svc.id}"></div>`;
  v.innerHTML=h;

  if(svc.rateCards) svc.rateCards.forEach(c=>renderRateTableN(svc.id,c.id));
  else if(svc.rate) renderRateTable(svc.id);
  renderAdjTable(svc.id);
  if(!svc.generator) wireDrop(svc.id);
  renderFilePills(svc.id);
  if(st.result) renderResult(svc.id);
}

/* ---------- Rate table ---------- */
function renderRateTable(id){
  const svc=SVC[id], st=state[id];
  let h=`<thead><tr>`;
  svc.rate.cols.forEach(c=>h+=`<th class="${c.num?'num':''}">${esc(c.label)}</th>`);
  h+=`<th style="width:34px"></th></tr></thead><tbody>`;
  st.rate.rows.forEach((row,i)=>{
    h+=`<tr>`;
    svc.rate.cols.forEach(c=>{
      const val=row[c.key]??"";
      h+=`<td class="${c.num?'num':''}"><input class="cell-in" style="width:100%" value="${esc(val)}"
        onchange="editRate('${id}',${i},'${c.key}',this.value)"></td>`;
    });
    h+=`<td><span class="x" title="Delete row" style="cursor:pointer;color:var(--danger);font-weight:700" onclick="delRateRow('${id}',${i})">✕</span></td></tr>`;
  });
  document.getElementById("rateTbl_"+id).innerHTML=h+"</tbody>";
}
function editRate(id,i,key,val){ const c=SVC[id].rate.cols.find(c=>c.key===key);
  state[id].rate.rows[i][key]= c.num ? num(val) : val; persistRate(id); if(state[id].result) rerun(id); }
function addRateRow(id){ const blank={}; SVC[id].rate.cols.forEach(c=>blank[c.key]=c.num?0:""); state[id].rate.rows.push(blank);
  renderRateTable(id); persistRate(id); }
function delRateRow(id,i){ state[id].rate.rows.splice(i,1); renderRateTable(id); persistRate(id); if(state[id].result) rerun(id); }
function resetRate(id){ state[id].rate={rows:SVC[id].rate.rows.map(r=>({...r}))}; saveRatesFor(id,state[id].rate.rows); renderRateTable(id); if(state[id].result) rerun(id); }
function persistRate(id){ saveRatesFor(id, state[id].rate.rows); }

/* ---------- Multi rate-card (services with svc.rateCards) ---------- */
function cardDef(id,cardId){ return SVC[id].rateCards.find(c=>c.id===cardId); }
function renderRateTableN(id,cardId){
  const c=cardDef(id,cardId), rows=state[id].rateCards[cardId].rows;
  let h=`<thead><tr>`; c.cols.forEach(col=>h+=`<th class="${col.num?'num':''}">${esc(col.label)}</th>`);
  h+=`<th style="width:34px"></th></tr></thead><tbody>`;
  rows.forEach((row,i)=>{ h+=`<tr>`;
    c.cols.forEach(col=>h+=`<td class="${col.num?'num':''}"><input class="cell-in" style="width:100%" value="${esc(row[col.key]??"")}"
      onchange="editRateN('${id}','${cardId}',${i},'${col.key}',this.value)"></td>`);
    h+=`<td><span style="cursor:pointer;color:var(--danger);font-weight:700" onclick="delRateRowN('${id}','${cardId}',${i})">✕</span></td></tr>`; });
  document.getElementById("rateTbl_"+id+"_"+cardId).innerHTML=h+"</tbody>";
}
function editRateN(id,cardId,i,key,val){ const col=cardDef(id,cardId).cols.find(c=>c.key===key);
  state[id].rateCards[cardId].rows[i][key]= col.num?num(val):val; persistRateN(id,cardId); if(state[id].result) rerun(id); }
function addRateRowN(id,cardId){ const blank={}; cardDef(id,cardId).cols.forEach(c=>blank[c.key]=c.num?0:"");
  state[id].rateCards[cardId].rows.push(blank); renderRateTableN(id,cardId); persistRateN(id,cardId); }
function delRateRowN(id,cardId,i){ state[id].rateCards[cardId].rows.splice(i,1); renderRateTableN(id,cardId); persistRateN(id,cardId); if(state[id].result) rerun(id); }
function resetRateN(id,cardId){ state[id].rateCards[cardId]={rows:cardDef(id,cardId).rows.map(r=>({...r}))};
  saveRatesFor(id+":"+cardId, state[id].rateCards[cardId].rows); renderRateTableN(id,cardId); if(state[id].result) rerun(id); }
function persistRateN(id,cardId){ saveRatesFor(id+":"+cardId, state[id].rateCards[cardId].rows); }
function rateFileInputN(id,cardId){
  const inp=document.createElement("input"); inp.type="file"; inp.accept=".xlsx,.xls,.csv";
  inp.onchange=e=>{ const f=e.target.files[0]; if(!f) return;
    const rd=new FileReader(); rd.onload=ev=>{
      const wb=XLSX.read(ev.target.result,{type:"array"});
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1}).filter(r=>r.some(c=>c!==""&&c!=null));
      applyRateFileN(id,cardId,rows);
    }; rd.readAsArrayBuffer(f); };
  inp.click();
}
function applyRateFileN(id,cardId,rows){
  const c=cardDef(id,cardId);
  if(!rows.length){ alert("Empty rate file."); return; }
  // find header row (row containing the key column name)
  let hr=0; for(let i=0;i<Math.min(rows.length,6);i++){ const hd=rows[i].map(norm); if(hd.some(h=>h.includes("code")||h.includes("destination"))){ hr=i; break; } }
  const hdr=rows[hr].map(norm);
  const colIdx={}; c.cols.forEach(col=>{ colIdx[col.key]=hdr.findIndex(h=>h===norm(col.key)||h.includes(norm(col.label.split("(")[0]))); });
  if(colIdx[c.keyCol]<0){ alert("Could not find the '"+c.keyCol+"' column in that file."); return; }
  const nw=[];
  rows.slice(hr+1).forEach(r=>{ if(r[colIdx[c.keyCol]]==null||r[colIdx[c.keyCol]]==="") return;
    const o={}; c.cols.forEach(col=>{ const i=colIdx[col.key]; let v=i>=0?r[i]:(col.num?0:""); o[col.key]=col.num?num(v):v; }); nw.push(o); });
  if(nw.length){ state[id].rateCards[cardId].rows=nw; persistRateN(id,cardId); renderRateTableN(id,cardId); if(state[id].result) rerun(id);
    toast(id,`Loaded ${nw.length} rows into ${c.label}.`,"ok"); }
  else alert("No rate rows found.");
}

/* ---------- Manual adjustments ---------- */
const ADJ_COLS=[{key:"date",label:"Date",num:false},{key:"desc",label:"Description",num:false},
  {key:"customer",label:"Customer",num:false},{key:"amount",label:"Amount (SGD)",num:true},{key:"cost",label:"Cost (SGD)",num:true}];
function renderAdjTable(id){
  const st=state[id], wrap=document.getElementById("adjWrap_"+id); if(!wrap) return;
  if(!st.adj.length){ wrap.innerHTML=`<p class="muted" style="padding:6px 2px;margin:0">No manual adjustments.</p>`; return; }
  let h=`<thead><tr>`; ADJ_COLS.forEach(c=>h+=`<th class="${c.num?'num':''}">${esc(c.label)}</th>`); h+=`<th style="width:34px"></th></tr></thead><tbody>`;
  st.adj.forEach((row,i)=>{ h+=`<tr>`;
    ADJ_COLS.forEach(c=>h+=`<td class="${c.num?'num':''}"><input class="cell-in" style="width:100%" value="${esc(row[c.key]??"")}"
      onchange="editAdj('${id}',${i},'${c.key}',this.value)"></td>`);
    h+=`<td><span style="cursor:pointer;color:var(--danger);font-weight:700" onclick="delAdj('${id}',${i})">✕</span></td></tr>`; });
  document.getElementById("adjTbl_"+id).innerHTML=h+"</tbody>";
}
function addAdj(id){ state[id].adj.push({date:todayISO(),desc:"",customer:"Adjustment",amount:0,cost:null});
  renderAdjTable(id); if(state[id].result) rerun(id); }
function editAdj(id,i,key,val){ const c=ADJ_COLS.find(c=>c.key===key); state[id].adj[i][key]=c.num?num(val):val;
  if(state[id].result) rerun(id); }
function delAdj(id,i){ state[id].adj.splice(i,1); renderAdjTable(id); if(state[id].result) rerun(id); }
function adjLines(id){
  return (state[id].adj||[]).filter(a=>num(a.amount)!=null).map(a=>({
    date:toISO(a.date), customer:a.customer||a.desc||"Adjustment", desc:a.desc,
    amount:round2(num(a.amount)||0), cost: num(a.cost)!=null?round2(num(a.cost)):null, _adj:true }));
}

function rateFileInput(id){
  const inp=document.createElement("input"); inp.type="file"; inp.accept=".xlsx,.xls,.csv";
  inp.onchange=e=>{ const f=e.target.files[0]; if(!f) return;
    const rd=new FileReader(); rd.onload=ev=>{
      const wb=XLSX.read(ev.target.result,{type:"array"});
      const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1}).filter(r=>r.some(c=>c!==""&&c!=null));
      applyRateFile(id,rows);
    }; rd.readAsArrayBuffer(f); };
  inp.click();
}
function applyRateFile(id,rows){
  const svc=SVC[id];
  if(!rows.length){ alert("Empty rate file."); return; }
  const hdr=rows[0].map(x=>norm(x));
  const colIdx={}; svc.rate.cols.forEach(c=>{ colIdx[c.key]=hdr.findIndex(h=>h.includes(norm(c.label.split("(")[0]))||h.includes(norm(c.key))); });
  const keyCol=svc.rate.keyCol;
  if(colIdx[keyCol]<0){ alert("Could not find the key column ("+keyCol+") in that file. Header row must contain a matching name."); return; }
  const nw=[];
  rows.slice(1).forEach(r=>{
    if(r[colIdx[keyCol]]==null||r[colIdx[keyCol]]==="") return;
    const o={}; svc.rate.cols.forEach(c=>{ const i=colIdx[c.key]; let val=i>=0?r[i]:(c.num?0:""); o[c.key]=c.num?num(val):val; });
    nw.push(o);
  });
  if(nw.length){ state[id].rate.rows=nw; persistRate(id); renderRateTable(id); if(state[id].result) rerun(id);
    toast(id,`Loaded ${nw.length} rate rows from file.`,"ok"); }
  else alert("No rate rows found in that file.");
}

/* ---------- File upload ---------- */
function wireDrop(id){
  const drop=document.getElementById("drop_"+id); if(!drop) return;
  drop.onclick=()=>{ const inp=document.createElement("input"); inp.type="file"; inp.multiple=true; inp.accept=".xlsx,.xls,.csv";
    inp.onchange=e=>handleFiles(id,[...e.target.files]); inp.click(); };
  ["dragover","dragenter"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add("hot");}));
  ["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove("hot");}));
  drop.addEventListener("drop",e=>{ if(e.dataTransfer.files.length) handleFiles(id,[...e.dataTransfer.files]); });
}
function handleFiles(id,files){
  const st=state[id]; let pending=files.length; let allRows=[]; let names=[];
  files.forEach(f=>{
    const rd=new FileReader();
    rd.onload=ev=>{
      const wb=XLSX.read(ev.target.result,{type:"array",cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true}).filter(r=>r.some(c=>c!==""&&c!=null));
      allRows.push({name:f.name,rows}); names.push(f.name);
      if(--pending===0) ingest(id,allRows);
    };
    rd.readAsArrayBuffer(f);
  });
}
function renderFilePills(id){
  const el=document.getElementById("filepills_"+id); if(!el) return;
  const st=state[id];
  el.innerHTML=(st.files||[]).map((f,i)=>`<span class="pill">📄 ${esc(f)} <span class="x" onclick="removeFile('${id}',${i})">✕</span></span>`).join("")
    + (st.files&&st.files.length?`<button class="subtle sm" onclick="restart('${id}')" style="vertical-align:middle">↺ Restart</button>`:"");
}
function removeFile(id,i){ state[id].files.splice(i,1); state[id].parsed.splice(i,1); reprocess(id); }
function restart(id){ state[id].files=[]; state[id].parsed=[]; state[id].result=null;
  document.getElementById("validation_"+id).innerHTML=""; document.getElementById("results_"+id).innerHTML="";
  renderFilePills(id); }

/* ---------- Validation + ingest ---------- */
function ingest(id, fileRows){
  const svc=SVC[id], st=state[id];
  st.files=st.files||[]; st.parsed=st.parsed||[];
  const valEl=document.getElementById("validation_"+id);
  let reports=[];
  fileRows.forEach(fr=>{
    const {mapped,report}=validateFile(svc,fr.rows,fr.name);
    st.files.push(fr.name); st.parsed.push(mapped);
    reports.push(report);
  });
  valEl.innerHTML=reports.join("");
  renderFilePills(id);
  reprocess(id);
}
function validateFile(svc,rows,name){
  if(!rows.length) return {mapped:[],report:`<div class="banner err">⚠ <b>${esc(name)}</b>: empty file.</div>`};
  const hdr=rows[0].map(x=>norm(x));
  const idx={}; const missing=[];
  svc.input.cols.forEach(c=>{
    let i=-1;
    for(const a of [c.key,...(c.aliases||[])]){ i=hdr.findIndex(h=>h===norm(a)); if(i>=0) break; }
    if(i<0) for(const a of [c.key,...(c.aliases||[])]){ i=hdr.findIndex(h=>h.includes(norm(a))); if(i>=0) break; }
    idx[c.key]=i;
    if(i<0 && c.required) missing.push(c.label);
  });
  if(missing.length)
    return {mapped:[],report:`<div class="banner err">⚠ <b>${esc(name)}</b> doesn't match the template. Missing required column(s): <b>${missing.map(esc).join(", ")}</b>. Fix and re-upload.</div>`};
  const mapped=rows.slice(1).map(r=>{
    const o={}; svc.input.cols.forEach(c=>{ o[c.key]= idx[c.key]>=0 ? r[idx[c.key]] : null; }); return o;
  }).filter(o=>Object.values(o).some(v=>v!==null&&v!==""));
  const extra=hdr.length - svc.input.cols.filter(c=>idx[c.key]>=0).length;
  return {mapped, report:`<div class="banner ok">✓ <b>${esc(name)}</b>: columns match (${mapped.length} rows).</div>`};
}
function reprocess(id){
  const st=state[id]; const rows=[].concat(...(st.parsed||[]));
  if(!rows.length){ st.result=null; document.getElementById("results_"+id).innerHTML=""; return; }
  rerun(id, rows);
  renderFilePills(id);
}
function rerun(id, rows){
  const svc=SVC[id], st=state[id];
  const data = rows || [].concat(...(st.parsed||[]));
  st.result = svc.calc.call(svc, data, st.rate, st);
  st.result.lines = st.result.lines.concat(adjLines(id));
  renderResult(id);
}
function runGenerator(id){
  const svc=SVC[id], st=state[id];
  st.result = svc.generate.call(svc, st.rate, st.opts);
  st.result.lines = st.result.lines.concat(adjLines(id));
  renderResult(id);
}
function setOpt(id,key,val){ state[id].opts[key]=val; }

/* ---------- Result rendering ---------- */
function renderResult(id){
  const svc=SVC[id], st=state[id], res=st.result;
  const el=document.getElementById("results_"+id);
  const totals=summarize(res);
  const cur=res.currency||"$";

  let h="";

  /* Output 1 — reconciliation (services that produce it) */
  if(res.recon){
    h+=`<div class="card" style="border-color:var(--accent2)"><div class="flexhead">
      <div><div class="step">Output 1 · Reconciliation with vendor</div><h3>${esc(res.recon.title)}</h3></div>
      ${svc.buildRecon?`<button class="ghost" onclick="downloadRecon('${id}')">⭳ Download reconciliation</button>`:""}
      </div><div class="metrics">`;
    res.recon.metrics.forEach(m=>{
      h+=`<div class="metric"><div class="lbl">${esc(m.label)}</div><div class="val" style="font-size:19px">${m.money?money(m.value,cur):esc(m.value)}</div></div>`;
    });
    h+=`</div><div class="banner ${res.recon.verdict.indexOf("✓")>=0?'ok':'warn'}">${esc(res.recon.verdict)}</div></div>`;
  }

  h+=`<div class="card"><div class="flexhead">
    <div><div class="step">${res.recon?'Output 2 · ':'Step 2 · Review output'}</div><h3>Customer billing</h3></div>
    <div style="display:flex;gap:8px">
      <button class="ghost" onclick="downloadResult('${id}')">⭳ Download billing</button>
      <button onclick="saveToRecords('${id}')">✔ Save to records</button>
    </div></div>`;

  h+=`<div class="metrics">
      <div class="metric accent"><div class="lbl">Total billing</div><div class="val">${money(totals.amount,cur)}</div></div>
      <div class="metric"><div class="lbl">Shipments</div><div class="val">${totals.shipments}</div></div>
      <div class="metric"><div class="lbl">Line items</div><div class="val">${res.lines.length}</div></div>`;
  if(totals.hasCost)
    h+=`<div class="metric good"><div class="lbl">Gross profit</div><div class="val">${money(totals.gp,cur)} · ${totals.gpPct.toFixed(1)}%</div></div>`;
  else h+=`<div class="metric"><div class="lbl">Gross profit</div><div class="val" style="font-size:14px;color:var(--ink2)">add cost in rate card →</div></div>`;
  h+=`</div>`;

  if(res.note) h+=`<p class="muted">${esc(res.note)}</p>`;

  if(res.review.length){
    h+=`<div class="banner warn">⚠ <b>${res.review.length} row(s) need review</b> — excluded from the total:<br>`
      + res.review.slice(0,8).map(r=>`• ${esc(r.ship||r.track||r.awb||r.ref||"?")} — ${esc(r.reason)}`).join("<br>")
      + (res.review.length>8?`<br><i>…and ${res.review.length-8} more (in the download's "Needs review" sheet)</i>`:"")
      + `</div>`;
  }

  // table
  const cols=res.columns, hasCost=totals.hasCost;
  h+=`<div class="tbl-scroll" style="margin-top:14px"><table><thead><tr>`;
  cols.forEach(c=>h+=`<th class="${c.num?'num':''}">${esc(c.l)}</th>`);
  if(hasCost) h+=`<th class="num">Cost</th><th class="num">GP</th>`;
  h+=`</tr></thead><tbody>`;
  res.lines.slice(0,300).forEach(o=>{
    h+=`<tr>`;
    cols.forEach(c=>{
      let val=o[c.k]; if(c.money&&val!=null) val=Number(val).toFixed(2); else if(val==null) val="";
      h+=`<td class="${c.num?'num':''} ${c.tot?'tot':''}">${esc(val)}</td>`;
    });
    if(hasCost){ const gp=o.cost!=null?round2(o.amount-o.cost):null;
      h+=`<td class="num">${o.cost!=null?o.cost.toFixed(2):"—"}</td><td class="num">${gp!=null?gp.toFixed(2):"—"}</td>`; }
    h+=`</tr>`;
  });
  // grand total row
  const span=cols.length-1;
  h+=`<tr><td class="tot" colspan="${span}" style="text-align:right">Grand total</td>
      <td class="num tot">${totals.amount.toFixed(2)}</td>${hasCost?'<td></td><td class="num tot">'+totals.gp.toFixed(2)+'</td>':''}</tr>`;
  h+=`</tbody></table></div>`;
  if(res.lines.length>300) h+=`<p class="muted">Showing first 300 of ${res.lines.length} rows. The download has them all.</p>`;
  h+=`</div>`;

  // per-customer breakdown
  const by=breakdownByCustomer(res);
  if(by.length>1){
    h+=`<div class="card"><div class="step">Step 3 · Summary</div><h3>By customer</h3>
        <div class="tbl-scroll" style="max-height:none"><table><thead><tr><th>Customer</th><th class="num">Shipments</th>
        <th class="num">Billing</th>${hasCost?'<th class="num">GP</th><th class="num">GP %</th>':''}</tr></thead><tbody>`;
    by.forEach(b=>{
      h+=`<tr><td>${esc(b.customer)}</td><td class="num">${b.shipments}</td><td class="num">${money(b.amount,cur)}</td>`;
      if(hasCost){ const pct=b.amount?b.gp/b.amount*100:0; h+=`<td class="num">${b.hasCost?money(b.gp,cur):"—"}</td><td class="num">${b.hasCost?pct.toFixed(1)+"%":"—"}</td>`; }
      h+=`</tr>`;
    });
    h+=`</tbody></table></div></div>`;
  }

  el.innerHTML=h;
  el.scrollIntoView({behavior:"smooth",block:"nearest"});
}

/* ---------- aggregation helpers ---------- */
function summarize(res){
  let amount=0,cost=0,hasCost=false; const ships=new Set();
  res.lines.forEach(o=>{ amount+=o.amount||0; if(o.cost!=null){cost+=o.cost;hasCost=true;}
    ships.add(o.ship||o.track||o.awb||o.ref||o.cn35||o.docket||JSON.stringify([o.date,o.customer,o.amount])); });
  const gp=amount-cost;
  return {amount:round2(amount),cost:round2(cost),hasCost,gp:round2(gp),gpPct:amount?gp/amount*100:0,shipments:ships.size};
}
function breakdownByCustomer(res){
  const by={};
  res.lines.forEach(o=>{ const k=o.customer||"—"; by[k]=by[k]||{customer:k,shipments:0,amount:0,cost:0,hasCost:false};
    by[k].shipments++; by[k].amount+=o.amount||0; if(o.cost!=null){by[k].cost+=o.cost;by[k].hasCost=true;} });
  return Object.values(by).map(b=>({...b,amount:round2(b.amount),gp:round2(b.amount-b.cost)})).sort((a,b)=>b.amount-a.amount);
}

/* ---------- Download ---------- */
function downloadRecon(id){
  const svc=SVC[id], st=state[id], res=st.result;
  if(!res||!svc.buildRecon){ alert("No reconciliation to download."); return; }
  const {wb,name}=svc.buildRecon(res,st);
  XLSX.writeFile(wb, `${name}_${todayISO()}.xlsx`);
}
function downloadResult(id){
  const svc=SVC[id], st=state[id], res=st.result; if(!res||!res.lines.length){ alert("Nothing to download yet."); return; }
  if(svc.buildWorkbook){ const {wb,name}=svc.buildWorkbook(res,st); XLSX.writeFile(wb, `${name}_${todayISO()}.xlsx`); return; }
  const totals=summarize(res);
  const header=res.columns.map(c=>c.l).concat(totals.hasCost?["Cost","GP"]:[]);
  const aoa=[[svc.name+" — Billing"],["Total billing (SGD):",totals.amount],["Shipments:",totals.shipments],[],header];
  res.lines.forEach(o=>{
    const row=res.columns.map(c=>o[c.k]);
    if(totals.hasCost) row.push(o.cost, o.cost!=null?round2(o.amount-o.cost):null);
    aoa.push(row);
  });
  if(res.review.length){ aoa.push([]); aoa.push(["NEEDS REVIEW — not billed"]);
    aoa.push(Object.keys(res.review[0]).filter(k=>k!=="reason").concat("Reason"));
    res.review.forEach(r=>aoa.push(Object.keys(r).filter(k=>k!=="reason").map(k=>r[k]).concat(r.reason))); }
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Billing");
  // customer summary sheet
  const by=breakdownByCustomer(res);
  const s2=[["Customer","Shipments","Billing (SGD)"].concat(totals.hasCost?["GP","GP %"]:[])];
  by.forEach(b=>s2.push([b.customer,b.shipments,b.amount].concat(totals.hasCost?[b.gp, b.amount?round2(b.gp/b.amount*100):0]:[])));
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s2),"By customer");
  XLSX.writeFile(wb, `${svc.name.replace(/[^\w]+/g,"_")}_Billing_${todayISO()}.xlsx`);
}

/* ---------- Records ---------- */
function saveToRecords(id){
  const svc=SVC[id], res=state[id].result; if(!res||!res.lines.length){ alert("Nothing to save."); return; }
  const totals=summarize(res);
  const months=[...new Set(res.lines.map(o=>periodOf(o.date||"").m).filter(m=>m!=="?"))];
  const rec={
    id:(crypto.randomUUID?crypto.randomUUID():"rec_"+Date.now()), serviceId:id, service:svc.name, savedBy:CURRENT_USER, savedAt:new Date().toISOString(),
    files:(state[id].files||[]).slice(), month:months[0]||todayISO().slice(0,7),
    totals, byCustomer:breakdownByCustomer(res).map(b=>({customer:b.customer,shipments:b.shipments,amount:b.amount,gp:b.hasCost?b.gp:null})),
    lines:res.lines, columns:res.columns, currency:res.currency||"$", reviewCount:res.review.length
  };
  RECORDS.unshift(rec); insertRecordDB(rec);
  toast(id,`Saved to records — ${money(totals.amount)} across ${totals.shipments} shipments.`,"ok");
}
function toast(id,msg,kind){
  const el=document.getElementById("results_"+id)||document.getElementById("view");
  const d=document.createElement("div"); d.className="banner "+(kind||"ok"); d.innerHTML="✓ "+esc(msg);
  el.prepend(d); setTimeout(()=>d.remove(),4000);
}

/* ---------------- RECORDS VIEW ---------------- */
function renderRecords(v){
  const recs=loadRecords();
  if(!recs.length){ v.innerHTML=`<div class="card"><div class="empty"><div class="big">≣</div>
    No saved runs yet.<br>Run a service and click <b>Save to records</b> to build history.</div></div>`; return; }
  let h=`<div class="card"><div class="flexhead"><div><div class="step">History</div>
      <h3>Saved billing runs</h3><p class="sub">Every saved output is stored in this browser. Click to view or re-download.</p></div>
      <button class="danger sm" onclick="clearRecords()">Clear all</button></div>
    <div class="tbl-scroll" style="max-height:none"><table><thead><tr>
      <th>Saved</th><th>Service</th><th>Month</th><th class="num">Billing</th><th class="num">Shipments</th>
      <th class="num">GP</th><th>Files</th><th></th></tr></thead><tbody>`;
  recs.forEach(r=>{
    h+=`<tr><td>${esc(r.savedAt.slice(0,16).replace("T"," "))}</td><td>${esc(r.service)}</td><td>${esc(r.month)}</td>
      <td class="num">${money(r.totals.amount)}</td><td class="num">${r.totals.shipments}</td>
      <td class="num">${r.totals.hasCost?money(r.totals.gp):"—"}</td>
      <td class="muted">${esc((r.files||[]).join(", ")||(r.serviceId==="pickup"?"generated":"—"))}</td>
      <td style="white-space:nowrap"><button class="subtle sm" onclick="redownload('${r.id}')">⭳</button>
        <button class="danger sm" onclick="delRecord('${r.id}')">✕</button></td></tr>`;
  });
  v.innerHTML=h+`</tbody></table></div></div>`;
}
function clearRecords(){ if(confirm("Delete ALL saved records for the whole team?")){ RECORDS=[]; if(SB) SB.from("billing_records").delete().neq("id","00000000-0000-0000-0000-000000000000").then(()=>{}); render(); } }
function delRecord(rid){ RECORDS=RECORDS.filter(r=>r.id!==rid); if(SB) SB.from("billing_records").delete().eq("id",rid).then(()=>{}); render(); }
function redownload(rid){
  const r=loadRecords().find(x=>x.id===rid); if(!r) return;
  const header=r.columns.map(c=>c.l).concat(r.totals.hasCost?["Cost","GP"]:[]);
  const aoa=[[r.service+" — Billing"],["Total billing (SGD):",r.totals.amount],[],header];
  r.lines.forEach(o=>{ const row=r.columns.map(c=>o[c.k]); if(r.totals.hasCost) row.push(o.cost,o.cost!=null?round2(o.amount-o.cost):null); aoa.push(row); });
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(aoa),"Billing");
  XLSX.writeFile(wb, `${r.service.replace(/[^\w]+/g,"_")}_${r.month}.xlsx`);
}

/* ---------------- DASHBOARD ---------------- */
let dashCut="month";
function renderDashboard(v){
  const recs=loadRecords();
  if(!recs.length){ v.innerHTML=`<div class="card"><div class="empty"><div class="big">▚</div>
    Nothing to show yet.<br>Process a service and <b>Save to records</b> — totals appear here by month, quarter, year and customer.</div></div>`; return; }

  // headline totals
  let amount=0,gp=0,ships=0,hasCost=false;
  recs.forEach(r=>{ amount+=r.totals.amount; ships+=r.totals.shipments; if(r.totals.hasCost){gp+=r.totals.gp;hasCost=true;} });

  let h=`<div class="metrics" style="margin-bottom:20px">
    <div class="metric accent"><div class="lbl">Total billing</div><div class="val">${money(round2(amount))}</div></div>
    <div class="metric"><div class="lbl">Shipments</div><div class="val">${ships.toLocaleString()}</div></div>
    <div class="metric"><div class="lbl">Saved runs</div><div class="val">${recs.length}</div></div>
    ${hasCost?`<div class="metric good"><div class="lbl">Gross profit</div><div class="val">${money(round2(gp))} · ${amount?(gp/amount*100).toFixed(1):0}%</div></div>`:""}
  </div>`;

  // by service
  const bySvc={};
  recs.forEach(r=>{ bySvc[r.service]=bySvc[r.service]||{amount:0,ships:0,gp:0,hasCost:false,runs:0};
    const b=bySvc[r.service]; b.amount+=r.totals.amount; b.ships+=r.totals.shipments; b.runs++;
    if(r.totals.hasCost){b.gp+=r.totals.gp;b.hasCost=true;} });
  h+=`<div class="card"><div class="step">Overview</div><h3>By service</h3>
      <div class="tbl-scroll" style="max-height:none"><table><thead><tr><th>Service</th><th class="num">Runs</th>
      <th class="num">Billing</th><th class="num">Shipments</th><th class="num">GP</th><th class="num">GP %</th></tr></thead><tbody>`;
  Object.entries(bySvc).sort((a,b)=>b[1].amount-a[1].amount).forEach(([s,b])=>{
    h+=`<tr><td>${esc(s)}</td><td class="num">${b.runs}</td><td class="num">${money(round2(b.amount))}</td>
      <td class="num">${b.ships}</td><td class="num">${b.hasCost?money(round2(b.gp)):"—"}</td>
      <td class="num">${b.hasCost&&b.amount?(b.gp/b.amount*100).toFixed(1)+"%":"—"}</td></tr>`;
  });
  h+=`</tbody></table></div></div>`;

  // time cut
  h+=`<div class="card"><div class="flexhead"><div><div class="step">Trend</div><h3>By period</h3></div>
      <div class="filters" style="margin:0">
        <button class="${dashCut==='month'?'':'subtle'} sm" onclick="setCut('month')">Month</button>
        <button class="${dashCut==='q'?'':'subtle'} sm" onclick="setCut('q')">Quarter</button>
        <button class="${dashCut==='y'?'':'subtle'} sm" onclick="setCut('y')">Year</button>
      </div></div>
      <div class="tbl-scroll" style="max-height:none"><table><thead><tr><th>Period</th><th class="num">Billing</th>
      <th class="num">Shipments</th><th class="num">GP</th></tr></thead><tbody>`;
  const byPeriod={};
  recs.forEach(r=>{ const p= dashCut==='month'?periodOf(r.month+"-01").m : dashCut==='q'?periodOf(r.month+"-01").q : periodOf(r.month+"-01").y;
    byPeriod[p]=byPeriod[p]||{amount:0,ships:0,gp:0,hasCost:false}; const b=byPeriod[p];
    b.amount+=r.totals.amount; b.ships+=r.totals.shipments; if(r.totals.hasCost){b.gp+=r.totals.gp;b.hasCost=true;} });
  Object.entries(byPeriod).sort().forEach(([p,b])=>{
    h+=`<tr><td>${esc(p)}</td><td class="num">${money(round2(b.amount))}</td><td class="num">${b.ships}</td>
      <td class="num">${b.hasCost?money(round2(b.gp)):"—"}</td></tr>`;
  });
  h+=`</tbody></table></div></div>`;

  // by customer
  const byCust={};
  recs.forEach(r=>(r.byCustomer||[]).forEach(c=>{ byCust[c.customer]=byCust[c.customer]||{amount:0,ships:0,gp:0,hasCost:false};
    const b=byCust[c.customer]; b.amount+=c.amount; b.ships+=c.shipments; if(c.gp!=null){b.gp+=c.gp;b.hasCost=true;} }));
  h+=`<div class="card"><div class="step">Customers</div><h3>By customer</h3>
      <div class="tbl-scroll" style="max-height:none"><table><thead><tr><th>Customer</th><th class="num">Shipments</th>
      <th class="num">Billing</th><th class="num">GP</th></tr></thead><tbody>`;
  Object.entries(byCust).sort((a,b)=>b[1].amount-a[1].amount).forEach(([c,b])=>{
    h+=`<tr><td>${esc(c)}</td><td class="num">${b.ships}</td><td class="num">${money(round2(b.amount))}</td>
      <td class="num">${b.hasCost?money(round2(b.gp)):"—"}</td></tr>`;
  });
  h+=`</tbody></table></div></div>`;

  v.innerHTML=h;
}
function setCut(c){ dashCut=c; render(); }

/* ---------- boot ---------- */

/* ============================================================
   SUPABASE LAYER — auth · shared data · realtime sync
   (appended to the reused calculation engine above)
   ============================================================ */
const $ = id => document.getElementById(id);

/* ---- record <-> db row mapping ---- */
function recordToDb(rec){
  return { id:rec.id, service_id:rec.serviceId, service:rec.service, month:rec.month,
    totals:rec.totals, by_customer:rec.byCustomer, lines:rec.lines, columns:rec.columns,
    currency:rec.currency, review_count:rec.reviewCount, files:rec.files,
    saved_by:rec.savedBy||CURRENT_USER, saved_at:rec.savedAt };
}
function dbToRecord(row){
  return { id:row.id, serviceId:row.service_id, service:row.service, month:row.month,
    totals:row.totals||{}, byCustomer:row.by_customer||[], lines:row.lines||[], columns:row.columns||[],
    currency:row.currency||"$", reviewCount:row.review_count||0, files:row.files||[],
    savedBy:row.saved_by, savedAt:row.saved_at };
}
function insertRecordDB(rec){
  if(!SB) return;
  SB.from("billing_records").insert(recordToDb(rec)).then(({error})=>{ if(error) console.error("record save failed", error); });
}

/* ---- initial load of shared data ---- */
async function loadAllData(){
  const {data:rc,error:e1}=await SB.from("rate_cards").select("*");
  if(e1) console.error("load rate_cards", e1);
  RATES={}; (rc||[]).forEach(r=>{ RATES[r.service_id+":"+r.card_id]=r.rows; });
  const {data:br,error:e2}=await SB.from("billing_records").select("*").order("saved_at",{ascending:false});
  if(e2) console.error("load billing_records", e2);
  RECORDS=(br||[]).map(dbToRecord);
}

/* ---- realtime ---- */
let RT=null;
function subscribeRealtime(){
  if(RT) return;
  RT=SB.channel("shipx-rt")
    .on("postgres_changes",{event:"*",schema:"public",table:"rate_cards"}, applyRateChange)
    .on("postgres_changes",{event:"*",schema:"public",table:"billing_records"}, applyRecordChange)
    .subscribe();
}
function applyRateChange(payload){
  const row = payload.new || payload.old; if(!row) return;
  const k = row.service_id+":"+row.card_id, svcId=row.service_id, cardId=row.card_id;
  if(payload.eventType==="DELETE") delete RATES[k]; else RATES[k]=row.rows;
  if(row.updated_by===CURRENT_USER) return;           // ignore our own echo
  const st=state[svcId];
  if(st){
    if(SVC[svcId] && SVC[svcId].rateCards && st.rateCards && st.rateCards[cardId]) st.rateCards[cardId].rows=(row.rows||[]).map(r=>({...r}));
    else if(st.rate) st.rate.rows=(row.rows||[]).map(r=>({...r}));
  }
  if(route.view==="service" && route.id===svcId){
    if(SVC[svcId].rateCards){ SVC[svcId].rateCards.forEach(c=>{ if($("rateTbl_"+svcId+"_"+c.id)) renderRateTableN(svcId,c.id); }); }
    else if($("rateTbl_"+svcId)) renderRateTable(svcId);
    if(st && st.result) rerun(svcId);
    toast(svcId, "Rate card updated by "+(row.updated_by||"a teammate")+" — refreshed.", "info");
  }
}
function applyRecordChange(payload){
  if(payload.eventType==="INSERT"){
    const rec=dbToRecord(payload.new);
    if(!RECORDS.some(r=>r.id===rec.id)) RECORDS.unshift(rec);
    RECORDS.sort((a,b)=>String(b.savedAt||"").localeCompare(String(a.savedAt||"")));
  } else if(payload.eventType==="DELETE"){
    RECORDS=RECORDS.filter(r=>r.id!==(payload.old&&payload.old.id));
  } else if(payload.eventType==="UPDATE"){
    const rec=dbToRecord(payload.new); RECORDS=RECORDS.map(r=>r.id===rec.id?rec:r);
  }
  if(route.view==="dashboard"||route.view==="records") render();
}

/* ---- auth screens ---- */
function showLogin(){ $("appRoot").style.display="none"; $("login").style.display="flex"; }
function showApp(){ $("login").style.display="none"; const a=$("appRoot"); a.style.display="flex"; a.style.width="100%"; }

let APP_STARTED=false;
async function startApp(session){
  if(APP_STARTED) return; APP_STARTED=true;
  CURRENT_USER=session.user.email;
  $("whoami").textContent=CURRENT_USER;
  for(const k in state) delete state[k];   // re-read shared rates fresh
  await loadAllData();
  subscribeRealtime();
  showApp();
  route={view:"dashboard",id:null};
  render();
}

function initAuth(){
  const cfg=window.SHIPX_CONFIG||{};
  if(!cfg.SUPABASE_URL || String(cfg.SUPABASE_URL).indexOf("PASTE_")===0){
    $("loginNote").innerHTML='✓ <b>App deployed.</b> Last step: create a Supabase project, then paste its URL &amp; anon key into <code>config.js</code> and redeploy to turn on login &amp; shared data. See <code>README.md</code>.';
    $("email").disabled=true; $("password").disabled=true; $("loginBtn").disabled=true; $("loginBtn").textContent="Waiting for setup";
    return;
  }
  if(!window.supabase){ $("loginNote").textContent="Could not load Supabase library (check your connection)."; return; }
  SB=window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY);

  SB.auth.onAuthStateChange((event,session)=>{
    if(session) startApp(session);
    else { APP_STARTED=false; showLogin(); }
  });
  SB.auth.getSession().then(({data})=>{ if(data && data.session) startApp(data.session); });

  $("loginForm").addEventListener("submit", async e=>{
    e.preventDefault();
    const email=$("email").value.trim(), password=$("password").value, err=$("loginErr");
    err.classList.remove("show");
    $("loginBtn").disabled=true; $("loginBtn").textContent="Signing in…";
    const {error}=await SB.auth.signInWithPassword({email,password});
    $("loginBtn").disabled=false; $("loginBtn").textContent="Sign in";
    if(error){ err.textContent=error.message||"Sign in failed."; err.classList.add("show"); }
  });
  $("logoutBtn").addEventListener("click", async ()=>{ if(SB) await SB.auth.signOut(); location.reload(); });
}
document.addEventListener("DOMContentLoaded", initAuth);
