import { spawnSync } from "node:child_process";
import fs from "node:fs";

const builder = process.env.DEALCHECK_BUILDER;
if (!builder) throw new Error("DEALCHECK_BUILDER is required");

const base = { address:"Test Property", asking:285000, offer:250000, arv:360000, rent:2850, repairs:42000, downPayment:0.2, interestRate:0.0725, loanYears:30 };
const cases = [
  {name:"Blank first-use state", data:{}, check:o=>o.verdict.startsWith("STOP")&&o.filled==="0 of 5"&&o.statuses.every(x=>x==="WAIT")},
  {name:"Weak sample deal", data:base, check:o=>o.verdict.startsWith("NO")&&o.wholesaleResult<12000&&o.safeOffer<o.offer},
  {name:"Strong rental", data:{...base,asking:190000,offer:170000,arv:230000,rent:3200,repairs:5000,interestRate:0.055}, check:o=>o.rentalStatus==="PASS"&&o.best.includes("RENT")&&o.expected>300},
  {name:"Strong flip", data:{...base,asking:180000,offer:150000,arv:330000,rent:900,repairs:25000}, check:o=>o.flipStatus==="PASS"&&o.noFormulaErrors},
  {name:"Strong wholesale", data:{...base,asking:175000,offer:150000,arv:320000,rent:500,repairs:25000}, check:o=>o.wholesaleStatus==="PASS"&&o.best.includes("CONTRACT")&&o.expected>=12000&&o.cash===0},
  {name:"Missing ARV", data:{...base,arv:null}, check:o=>o.verdict.startsWith("STOP")&&o.filled==="4 of 5"},
  {name:"Zero rent non-rental", data:{...base,rent:0,offer:150000,arv:320000,repairs:25000}, check:o=>!o.verdict.startsWith("STOP")&&o.rentalStatus==="FAIL"},
  {name:"Very high repairs", data:{...base,repairs:250000}, check:o=>o.verdict.startsWith("NO")&&o.statuses.every(x=>x==="FAIL")},
  {name:"Zero-interest loan", data:{...base,interestRate:0}, check:o=>o.noFormulaErrors&&Number.isFinite(o.cash)},
  {name:"All-cash-equivalent down payment", data:{...base,downPayment:1}, check:o=>o.noFormulaErrors&&o.debtService===0},
  {name:"Offer exactly at wholesale MAO", data:{...base,offer:218970.87378640776}, check:o=>Math.abs(o.wholesaleResult-12000)<0.02&&o.wholesaleStatus==="PASS"},
  {name:"Offer one dollar above wholesale MAO", data:{...base,offer:218971.87378640776}, check:o=>o.wholesaleResult<12000&&o.wholesaleStatus!=="PASS"},
];

const table = ndjson => JSON.parse(ndjson).values;
const mapRows = values => Object.fromEntries(values.map(r=>[r[0],r[1]]));
const results=[];
for(const tc of cases){
  const run=spawnSync(process.execPath,[builder],{encoding:"utf8",env:{...process.env,DEALCHECK_TEST_MODE:"1",DEALCHECK_SCENARIO:JSON.stringify(tc.data)}});
  if(run.status!==0){results.push({name:tc.name,pass:false,error:run.stderr||run.stdout});continue;}
  const raw=JSON.parse(run.stdout);
  const desk=table(raw.inspect), financing=mapRows(table(raw.audit.financing)), wholesale=mapRows(table(raw.audit.wholesale));
  const o={
    verdict:desk[6][4], best:desk[7][4], score:desk[8][4], safeOffer:desk[9][4], cash:desk[10][4], expected:desk[11][4],
    filled:desk[17][1], offer:tc.data.offer??0, statuses:[desk[21][7],desk[22][7],desk[23][7]], rentalStatus:desk[21][7],flipStatus:desk[22][7],wholesaleStatus:desk[23][7],
    wholesaleResult:desk[23][4], debtService:financing["Annual debt service"], supportedAssignment:wholesale["Supported assignment fee at your offer"],
    noFormulaErrors:raw.errors.includes("matched 0 entries"),
  };
  let pass=false,error=null; try{pass=Boolean(tc.check(o));}catch(e){error=e.message;}
  results.push({name:tc.name,pass,error,observed:o});
}

// Independent reconciliation for the weak sample deal.
const s=base;
const loan=s.offer*(1-s.downPayment);
const monthlyRate=s.interestRate/12;
const payment=loan*(monthlyRate*Math.pow(1+monthlyRate,s.loanYears*12))/(Math.pow(1+monthlyRate,s.loanYears*12)-1);
const gross=s.rent*12;
const effective=gross*(1-0.07);
const noi=effective-4200-3000-175*12-effective*(0.08+0.06+0.05);
const rehab=s.repairs*1.15;
const flipProject=s.offer+s.offer*0.03+rehab+(loan*0.01+1800)+(550*6)+(payment*6);
const flipProfit=s.arv-s.arv*0.08-flipProject;
const supportedAssignment=(s.arv-s.offer-rehab-45000-s.arv*0.08-s.offer*0.03)/(1+0.03);

const weakRun=spawnSync(process.execPath,[builder],{encoding:"utf8",env:{...process.env,DEALCHECK_TEST_MODE:"1",DEALCHECK_SCENARIO:JSON.stringify(base)}});
const weakRaw=JSON.parse(weakRun.stdout);
const rentalRows=mapRows(table(weakRaw.audit.rental)), flipRows=mapRows(table(weakRaw.audit.flip)), wholesaleRows=mapRows(table(weakRaw.audit.wholesale));
const reconciliations=[
  {name:"Monthly payment reconciliation",actual:mapRows(table(weakRaw.audit.financing))["Monthly principal & interest"],expected:payment,tolerance:0.02},
  {name:"Rental NOI reconciliation",actual:rentalRows["Net operating income (NOI)"],expected:noi,tolerance:0.02},
  {name:"Flip profit reconciliation",actual:flipRows["Net flip profit"],expected:flipProfit,tolerance:0.02},
  {name:"Wholesale supported-fee reconciliation",actual:wholesaleRows["Supported assignment fee at your offer"],expected:supportedAssignment,tolerance:0.02},
].map(x=>({...x,pass:Math.abs(x.actual-x.expected)<=x.tolerance}));

const report={scenarioTests:results,reconciliations,scenarioPassed:results.filter(x=>x.pass).length,scenarioTotal:results.length,reconciliationPassed:reconciliations.filter(x=>x.pass).length,reconciliationTotal:reconciliations.length};
fs.mkdirSync("outputs/dealcheck-pro",{recursive:true});
fs.writeFileSync("outputs/dealcheck-pro/test-report.json",JSON.stringify(report,null,2));
console.log(JSON.stringify(report,null,2));
if(report.scenarioPassed!==report.scenarioTotal||report.reconciliationPassed!==report.reconciliationTotal) process.exit(1);
