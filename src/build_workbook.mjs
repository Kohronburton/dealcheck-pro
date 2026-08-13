import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const wb = Workbook.create();
const names = ["Start Here","Quick Analysis","Property Input","Financing","Rental","Flip","Wholesale","Checks"];
for (const n of names) wb.worksheets.add(n);

const navy="#12233F", teal="#00A6A6", light="#EAF2F8", pale="#E8F7F6", gold="#F4B942", red="#C0392B", green="#18864B", gray="#667085";
const money='"$"#,##0;[Red]("$"#,##0);-';
const pct='0.0%;[Red](0.0%);-';
const title=(s,range,text)=>{s.getRange(range).merge();s.getRange(range).values=[[text]];s.getRange(range).format={fill:navy,font:{bold:true,color:"#FFFFFF",size:18},verticalAlignment:"center"};};
const section=(s,range,text)=>{s.getRange(range).merge();s.getRange(range).values=[[text]];s.getRange(range).format={fill:teal,font:{bold:true,color:"#FFFFFF"}};};
const setup=(s)=>{s.showGridLines=false;};
for(const n of names) setup(wb.worksheets.getItem(n));

const start=wb.worksheets.getItem("Start Here");
title(start,"A1:H2","DealCheck Pro | Fast Decision MVP");
start.getRange("A4:H4").merge(); start.getRange("A4:H4").values=[["Enter the property once. Compare Rental, Flip, and Wholesale. Know the risk and the safest price."]];
start.getRange("A4:H4").format={fill:pale,font:{bold:true,color:navy,size:12},wrapText:true};
section(start,"A6:H6","3-STEP WORKFLOW");
start.getRange("A8:H10").values=[
  ["1","Enter inputs","Use Quick Analysis for the minimum inputs, then Property Input for detail.",null,null,null,null,null],
  ["2","Review strategies","Rental, Flip, and Wholesale update automatically.",null,null,null,null,null],
  ["3","Check the model","Open Checks before relying on the result.",null,null,null,null,null]
];
start.getRange("A8:A10").format={fill:teal,font:{bold:true,color:"#FFFFFF",size:14},horizontalAlignment:"center"};
start.getRange("B8:B10").format.font={bold:true,color:navy}; start.getRange("C8:H10").merge(true); start.getRange("C8:H10").format.wrapText=true;
section(start,"A12:H12","COLOR LEGEND");
start.getRange("A14:H17").values=[["Blue text","Editable input","Black text","Formula","Green text","Cross-sheet formula","Yellow fill","Needs attention"],["GO","Meets target","REVIEW","Verify assumptions","NO-GO","Fails hard stop","Version","0.1.0"],[null,null,null,null,null,null,null,null],["Disclaimer","Educational estimates only; not legal, tax, lending, appraisal, financial, or investment advice. Verify all data independently.",null,null,null,null,null,null]];
start.getRange("A14:A14").format.font={color:"#0000FF"}; start.getRange("E14:E14").format.font={color:"#008000"}; start.getRange("G14:H14").format.fill="#FFF2CC";
start.getRange("B17:H17").merge(); start.getRange("B17:H17").format={wrapText:true,font:{italic:true,color:gray}};
start.getRange("A1:H17").format.rowHeight=22; start.getRange("A:A").format.columnWidth=13; start.getRange("B:B").format.columnWidth=22; start.getRange("C:H").format.columnWidth=15;

const input=wb.worksheets.getItem("Property Input");
title(input,"A1:D2","Property & Buy Box Inputs"); section(input,"A4:D4","PROPERTY");
const rows=[
 ["Property address","125 Sample Ave, Miami, FL","Text","Sample property"],
 ["Asking price",285000,"Currency","Seller ask"],
 ["Proposed purchase price",250000,"Currency","Your expected contract price"],
 ["After-repair value (ARV)",360000,"Currency","Verify with comparable sales"],
 ["Monthly market rent",2850,"Currency","Verify with rental comps"],
 ["Other monthly income",0,"Currency","Parking, laundry, etc."],
 ["Rehab budget",42000,"Currency","Before contingency"],
 ["Rehab contingency",0.15,"Percent","Recommended buffer"],
 ["Annual property taxes",4200,"Currency","Current or projected"],
 ["Annual insurance",3000,"Currency","Obtain a quote"],
 ["Monthly HOA",0,"Currency","If applicable"],
 ["Vacancy rate",0.07,"Percent","Expected economic vacancy"],
 ["Management rate",0.08,"Percent","Percent of collected income"],
 ["Maintenance rate",0.06,"Percent","Percent of collected income"],
 ["CapEx rate",0.05,"Percent","Long-term replacements"],
 ["Monthly utilities/other",175,"Currency","Owner-paid operating costs"],
 ["Buyer closing costs",0.03,"Percent","Percent of purchase price"],
 ["Selling costs",0.08,"Percent","Agent, closing, concessions"],
 ["Holding period (months)",6,"Months","Flip holding period"],
 ["Monthly holding costs",550,"Currency","Utilities, lawn, security, etc."],
 ["Minimum monthly cash flow",300,"Currency","Buy Box"],
 ["Minimum cash-on-cash",0.08,"Percent","Buy Box"],
 ["Minimum DSCR",1.2,"Ratio","Buy Box"],
 ["Minimum flip profit",40000,"Currency","Buy Box"],
 ["Target wholesale assignment fee",12000,"Currency","Your desired fee"],
 ["End-buyer target profit",45000,"Currency","Editable buyer requirement"]
];
input.getRange("A5:D30").values=rows;
input.getRange("A5:A30").format.font={bold:true,color:navy}; input.getRange("B5:B30").format={fill:"#EFF6FF",font:{color:"#0000FF"}};
input.getRange("D5:D30").format={font:{color:gray},wrapText:true}; input.getRange("B6:B11").format.numberFormat=money; input.getRange("B13:B15").format.numberFormat=money; input.getRange("B20:B20").format.numberFormat=money; input.getRange("B24:B25").format.numberFormat=money; input.getRange("B28:B30").format.numberFormat=money;
input.getRange("B12:B12").format.numberFormat=pct; input.getRange("B16:B18").format.numberFormat=pct; input.getRange("B21:B21").format.numberFormat=pct; input.getRange("B9:B9").format.numberFormat=pct; input.getRange("B22:B22").format.numberFormat=pct;
input.freezePanes.freezeRows(4); input.getRange("A:A").format.columnWidth=30; input.getRange("B:B").format.columnWidth=18; input.getRange("C:C").format.columnWidth=13; input.getRange("D:D").format.columnWidth=34;

const fin=wb.worksheets.getItem("Financing"); title(fin,"A1:D2","Financing & Acquisition Cash"); section(fin,"A4:D4","LOAN INPUTS");
fin.getRange("A5:D12").values=[["Financing type","Amortizing","Input","Cash or Amortizing"],["Down payment",0.2,"Input","Percent of purchase price"],["Annual interest rate",0.0725,"Input","Nominal annual rate"],["Loan term (years)",30,"Input","Amortization term"],["Loan points",0.01,"Input","Percent of loan amount"],["Other loan fees",1800,"Input","Dollar amount"],["Acquisition reserves",7500,"Input","Initial cash reserve"],["Monthly payment",null,"Formula","Principal and interest"]];
fin.getRange("B5:B11").format={fill:"#EFF6FF",font:{color:"#0000FF"}}; fin.getRange("B6:B7").format.numberFormat=pct; fin.getRange("B9:B9").format.numberFormat=pct; fin.getRange("B10:B11").format.numberFormat=money;
section(fin,"A14:D14","CALCULATED CAPITAL"); fin.getRange("A15:A21").values=[["Purchase price"],["Down payment dollars"],["Loan amount"],["Monthly principal & interest"],["Annual debt service"],["Loan points dollars"],["Total acquisition cash"]];
fin.getRange("B15:B21").formulas=[["='Property Input'!B7"],["=IF(B5=\"Cash\",B15,B15*B6)"],["=MAX(0,B15-B16)"],["=IF(OR(B5=\"Cash\",B17=0),0,IF(B7=0,B17/(B8*12),-PMT(B7/12,B8*12,B17)))"],["=B18*12"],["=B17*B9"],["=B16+('Property Input'!B7*'Property Input'!B21)+B20+B10+B11"]];
fin.getRange("B12").formulas=[["=B18"]]; fin.getRange("B12:B21").format.numberFormat=money; fin.getRange("B15:B21").format.font={color:"#008000"};
fin.getRange("B5").dataValidation={rule:{type:"list",values:["Cash","Amortizing"]}}; fin.getRange("A:A").format.columnWidth=31; fin.getRange("B:B").format.columnWidth=19; fin.getRange("C:C").format.columnWidth=14; fin.getRange("D:D").format.columnWidth=32;

const rental=wb.worksheets.getItem("Rental"); title(rental,"A1:D2","Rental Analysis"); section(rental,"A4:D4","INCOME & OPERATIONS");
rental.getRange("A5:A23").values=[["Gross scheduled income"],["Vacancy loss"],["Other income"],["Effective gross income"],["Property taxes"],["Insurance"],["HOA"],["Utilities/other"],["Management"],["Maintenance"],["Capital expenditures"],["Total operating expenses"],["Net operating income (NOI)"],["Annual debt service"],["Annual pre-tax cash flow"],["Monthly cash flow"],["Cap rate"],["DSCR"],["Cash-on-cash return"]];
rental.getRange("B5:B23").formulas=[
 ["='Property Input'!B9*12"],["=-B5*'Property Input'!B16"],["='Property Input'!B10*12"],["=SUM(B5:B7)"],["=-'Property Input'!B13"],["=-'Property Input'!B14"],["=-'Property Input'!B15*12"],["=-'Property Input'!B20*12"],["=-B8*'Property Input'!B17"],["=-B8*'Property Input'!B18"],["=-B8*'Property Input'!B19"],["=SUM(B9:B15)"],["=B8+B16"],["=-'Financing'!B19"],["=B17+B18"],["=B19/12"],["=IFERROR(B17/'Property Input'!B7,0)"],["=IF('Financing'!B19=0,0,B17/'Financing'!B19)"],["=IFERROR(B19/('Financing'!B21+'Property Input'!B11*(1+'Property Input'!B12)),0)"]];
rental.getRange("B5:B20").format.numberFormat=money; rental.getRange("B21:B23").format.numberFormat=pct; rental.getRange("B22").format.numberFormat="0.00x"; rental.getRange("B5:B23").format.font={color:"#008000"};
section(rental,"A25:D25","BUY BOX RESULT"); rental.getRange("A26:A28").values=[["Monthly cash flow target"],["Cash-on-cash target"],["DSCR target"]]; rental.getRange("B26:B28").formulas=[["=IF(B20>='Property Input'!B25,\"PASS\",\"FAIL\")"],["=IF(B23>='Property Input'!B26,\"PASS\",\"FAIL\")"],["=IF(OR('Financing'!B19=0,B22>='Property Input'!B27),\"PASS\",\"FAIL\")"]];
rental.getRange("A:A").format.columnWidth=32; rental.getRange("B:B").format.columnWidth=20; rental.getRange("C:D").format.columnWidth=18;

const flip=wb.worksheets.getItem("Flip"); title(flip,"A1:D2","Fix & Flip Analysis"); section(flip,"A4:D4","PROJECT ECONOMICS");
flip.getRange("A5:A20").values=[["Purchase price"],["Buyer closing costs"],["Rehab budget"],["Rehab contingency"],["Total rehab"],["Loan points & fees"],["Holding costs"],["Estimated loan interest"],["Total project cost"],["Expected sale price (ARV)"],["Selling costs"],["Net flip profit"],["ROI on acquisition cash + rehab"],["Break-even sale price"],["Maximum offer for target profit"],["Profit with 20% rehab overrun"]];
flip.getRange("B5:B20").formulas=[["='Property Input'!B7"],["=B5*'Property Input'!B21"],["='Property Input'!B11"],["=B7*'Property Input'!B12"],["=SUM(B7:B8)"],["='Financing'!B20+'Financing'!B10"],["='Property Input'!B22*'Property Input'!B23"],["='Financing'!B18*'Property Input'!B23"],["=SUM(B5:B12)"],["='Property Input'!B8"],["=-B14*'Property Input'!B22"],["=B14+B15-B13"],["=IFERROR(B16/('Financing'!B21+B9),0)"],["=IFERROR(B13/(1-'Property Input'!B22),0)"],["=MAX(0,('Property Input'!B8*(1-'Property Input'!B22))-'Property Input'!B11*(1+'Property Input'!B12)-B10-B11-B12-'Property Input'!B28)"],["=B16-('Property Input'!B11*0.2)"]];
flip.getRange("B5:B16").format.numberFormat=money; flip.getRange("B17").format.numberFormat=pct; flip.getRange("B18:B20").format.numberFormat=money; flip.getRange("B5:B20").format.font={color:"#008000"};
section(flip,"A22:D22","VERDICT"); flip.getRange("A23:A24").values=[["Meets profit target"],["ARV covers project cost"]]; flip.getRange("B23:B24").formulas=[["=IF(B16>='Property Input'!B28,\"PASS\",\"FAIL\")"],["=IF(B14>B13,\"PASS\",\"FAIL\")"]]; flip.getRange("A:A").format.columnWidth=35; flip.getRange("B:B").format.columnWidth=20; flip.getRange("C:D").format.columnWidth=18;

const wh=wb.worksheets.getItem("Wholesale"); title(wh,"A1:D2","Wholesale MAO & Buyer Spread"); section(wh,"A4:D4","BUYER-DRIVEN OFFER");
wh.getRange("A5:A15").values=[["ARV"],["Repairs incl. contingency"],["Buyer target profit"],["Buyer selling costs"],["Buyer closing costs"],["Target assignment fee"],["Maximum seller contract price"],["Buyer purchase price"],["Buyer total project cost"],["Buyer projected profit"],["Negotiation room vs proposed price"]];
wh.getRange("B5:B15").formulas=[["='Property Input'!B8"],["='Property Input'!B11*(1+'Property Input'!B12)"],["='Property Input'!B30"],["=B5*'Property Input'!B22"],["='Property Input'!B7*'Property Input'!B21"],["='Property Input'!B29"],["=MAX(0,B5-B6-B7-B8-B9-B10)"],["=B11+B10"],["=SUM(B6,B8,B9,B12)"],["=B5-B13"],["=B11-'Property Input'!B7"]];
wh.getRange("B5:B15").format.numberFormat=money; wh.getRange("B5:B15").format.font={color:"#008000"}; section(wh,"A17:D17","VERDICT"); wh.getRange("A18:A19").values=[["Assignment fee supported"],["Proposed price is at/below MAO"]]; wh.getRange("B18:B19").formulas=[["=IF(B14>=B7,\"PASS\",\"FAIL\")"],["=IF('Property Input'!B7<=B11,\"PASS\",\"FAIL\")"]]; wh.getRange("A:A").format.columnWidth=35; wh.getRange("B:B").format.columnWidth=20; wh.getRange("C:D").format.columnWidth=18;

const quick=wb.worksheets.getItem("Quick Analysis"); title(quick,"A1:H2","Quick Analysis | One Property, Three Strategies");
quick.getRange("A4:H4").merge(); quick.getRange("A4:H4").values=[["Expected result based on the current sample inputs. Replace blue inputs on Property Input and Financing."]]; quick.getRange("A4:H4").format={fill:pale,font:{color:navy,bold:true}};
quick.getRange("A6:B12").values=[["Property",null],["Proposed price",null],["Cash required",null],["Rental cash flow",null],["Flip profit",null],["Wholesale MAO",null],["Model status",null]];
quick.getRange("B6:B12").formulas=[["='Property Input'!B5"],["='Property Input'!B7"],["='Financing'!B21"],["='Rental'!B20"],["='Flip'!B16"],["='Wholesale'!B11"],["='Checks'!B5"]];
quick.getRange("B7:B11").format.numberFormat=money; quick.getRange("A6:A12").format={fill:light,font:{bold:true,color:navy}}; quick.getRange("B6:B12").format.font={color:"#008000",bold:true};
quick.getRange("D6:H6").values=[["Strategy","Key result","Target","Status","What it means"]]; quick.getRange("D6:H6").format={fill:navy,font:{bold:true,color:"#FFFFFF"}};
quick.getRange("D7:H9").values=[["Rental",null,null,null,"Monthly cash flow vs your Buy Box"],["Flip",null,null,null,"Expected profit vs your Buy Box"],["Wholesale",null,null,null,"Proposed price vs buyer-driven MAO"]];
quick.getRange("E7:G9").formulas=[["='Rental'!B20","='Property Input'!B25","=IF(E7>=F7,\"GO\",\"REVIEW\")"],["='Flip'!B16","='Property Input'!B28","=IF(E8>=F8,\"GO\",\"REVIEW\")"],["='Wholesale'!B11","='Property Input'!B7","=IF(F9<=E9,\"GO\",\"NO-GO\")"]];
quick.getRange("E7:F9").format.numberFormat=money; quick.getRange("D7:D9").format.font={bold:true,color:navy}; quick.getRange("G7:G9").format.font={bold:true};
quick.getRange("A:A").format.columnWidth=22; quick.getRange("B:B").format.columnWidth=24; quick.getRange("C:C").format.columnWidth=3; quick.getRange("D:D").format.columnWidth=16; quick.getRange("E:G").format.columnWidth=16; quick.getRange("H:H").format.columnWidth=32;

const checks=wb.worksheets.getItem("Checks"); title(checks,"A1:G2","Model Checks");
checks.getRange("A4:G4").values=[["Check","Actual","Expected","Difference","Tolerance","Status","Fix hint"]]; checks.getRange("A4:G4").format={fill:navy,font:{bold:true,color:"#FFFFFF"}};
checks.getRange("A5:A11").values=[["Overall model status"],["Purchase price entered"],["ARV entered"],["Financing ties"],["Rental NOI ties"],["Flip project cost positive"],["Wholesale spread nonnegative"]];
checks.getRange("B5:B11").formulas=[["=IF(COUNTIF(F6:F11,\"FAIL\")=0,\"OK\",\"REVIEW\")"],["='Property Input'!B7"],["='Property Input'!B8"],["='Financing'!B16+'Financing'!B17"],["='Rental'!B8+'Rental'!B16"],["='Flip'!B13"],["='Wholesale'!B14"]];
checks.getRange("C6:C11").formulas=[["=0"],["=0"],["='Financing'!B15"],["='Rental'!B17"],["=0"],["=0"]];
checks.getRange("D6:D11").formulas=[["=IF(B6>0,0,1)"],["=IF(B7>0,0,1)"],["=B8-C8"],["=B9-C9"],["=IF(B10>0,0,1)"],["=IF(B11>=0,0,1)"]]; checks.getRange("E6:E11").values=[[0],[0],[0.01],[0.01],[0],[0]];
checks.getRange("F6:F11").formulas=[["=IF(D6<=E6,\"OK\",\"FAIL\")"],["=IF(D7<=E7,\"OK\",\"FAIL\")"],["=IF(ABS(D8)<=E8,\"OK\",\"FAIL\")"],["=IF(ABS(D9)<=E9,\"OK\",\"FAIL\")"],["=IF(D10<=E10,\"OK\",\"FAIL\")"],["=IF(D11<=E11,\"OK\",\"FAIL\")"]];
checks.getRange("G6:G11").values=[["Enter a proposed purchase price"],["Enter a verified ARV"],["Review down payment and loan amount"],["Review income and expense schedule"],["Review project costs"],["Reduce fee/price or improve buyer economics"]];
checks.getRange("B6:E11").format.numberFormat="0.00"; checks.getRange("A:A").format.columnWidth=29; checks.getRange("B:F").format.columnWidth=15; checks.getRange("G:G").format.columnWidth=36; checks.getRange("G5:G11").format.wrapText=true;

for(const n of ["Rental","Flip","Wholesale","Checks"]){const s=wb.worksheets.getItem(n);s.freezePanes.freezeRows(4);s.getRange("A4:D30").format.borders={preset:"inside",style:"thin",color:"#D7DEE8"};}
for(const n of ["Rental","Flip","Wholesale"]){const s=wb.worksheets.getItem(n);s.getRange("B23:B28").conditionalFormats.add("containsText",{text:"PASS",format:{fill:"#D9EAD3",font:{color:green,bold:true}}});s.getRange("B23:B28").conditionalFormats.add("containsText",{text:"FAIL",format:{fill:"#F4CCCC",font:{color:red,bold:true}}});}
quick.getRange("G7:G9").conditionalFormats.add("containsText",{text:"GO",format:{fill:"#D9EAD3",font:{color:green,bold:true}}}); quick.getRange("G7:G9").conditionalFormats.add("containsText",{text:"REVIEW",format:{fill:"#FFF2CC",font:{color:"#8A5A00",bold:true}}}); quick.getRange("G7:G9").conditionalFormats.add("containsText",{text:"NO-GO",format:{fill:"#F4CCCC",font:{color:red,bold:true}}});
checks.getRange("F6:F11").conditionalFormats.add("containsText",{text:"OK",format:{fill:"#D9EAD3",font:{color:green,bold:true}}}); checks.getRange("F6:F11").conditionalFormats.add("containsText",{text:"FAIL",format:{fill:"#F4CCCC",font:{color:red,bold:true}}});

const outDir="/workspace/scratch/2ed71f09e1c4/outputs/dealcheck-pro"; await fs.mkdir(outDir,{recursive:true});
for(const n of names){const png=await wb.render({sheetName:n,autoCrop:"all",scale:1,format:"png"});await fs.writeFile(`${outDir}/${n.replaceAll(" ","-").toLowerCase()}.png`,new Uint8Array(await png.arrayBuffer()));}
const inspect=await wb.inspect({kind:"table",range:"Quick Analysis!A1:H12",include:"values,formulas",tableMaxRows:14,tableMaxCols:8,maxChars:8000});
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:100},summary:"final formula error scan"});
await fs.writeFile(`${outDir}/verification.txt`,inspect.ndjson+"\n\n"+errors.ndjson);
const xlsx=await SpreadsheetFile.exportXlsx(wb); await xlsx.save(`${outDir}/DealCheck-Pro-Sprint-1-MVP.xlsx`);
console.log(JSON.stringify({output:`${outDir}/DealCheck-Pro-Sprint-1-MVP.xlsx`,inspect:inspect.ndjson,errors:errors.ndjson},null,2));
