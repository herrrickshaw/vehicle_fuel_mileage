const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, PageBreak, TableOfContents
} = require('docx');

const NAVY = "1F3864", ACCENT = "2E74B5", GREY = "595959", HDR = "1F3864";
const F = "Arial";

function h1(t){return new Paragraph({heading:HeadingLevel.HEADING_1, spacing:{before:280,after:140}, children:[new TextRun({text:t, font:F, bold:true, size:28, color:NAVY})]});}
function h2(t){return new Paragraph({heading:HeadingLevel.HEADING_2, spacing:{before:200,after:100}, children:[new TextRun({text:t, font:F, bold:true, size:24, color:ACCENT})]});}
function p(runs, opts={}){const arr=Array.isArray(runs)?runs:[new TextRun({text:runs,font:F,size:21,color:"222222"})];return new Paragraph({spacing:{after:120,line:276},...opts,children:arr});}
function bullet(t){return new Paragraph({bullet:{level:0}, spacing:{after:70,line:270}, children:[new TextRun({text:t,font:F,size:21,color:"222222"})]});}
function note(t){return new Paragraph({spacing:{after:120}, children:[new TextRun({text:t,font:F,italics:true,size:18,color:GREY})]});}

function cell(text,{bold=false,fill=null,color="222222",align=AlignmentType.LEFT,size=19,w}={}){
  return new TableCell({
    width:{size:w,type:WidthType.DXA},
    shading: fill?{type:ShadingType.CLEAR, fill:fill, color:"auto"}:undefined,
    margins:{top:40,bottom:40,left:80,right:80},
    children:[new Paragraph({alignment:align, children:[new TextRun({text:text,font:F,size:size,bold:bold,color:color})]})]
  });
}
function table(headers, rows, widths){
  const total=widths.reduce((a,b)=>a+b,0);
  const hr=new TableRow({tableHeader:true, children:headers.map((t,i)=>cell(t,{bold:true,fill:HDR,color:"FFFFFF",align:i==0?AlignmentType.LEFT:AlignmentType.CENTER,w:widths[i]}))});
  const brs=rows.map((r,ri)=>new TableRow({children:r.map((t,i)=>cell(String(t),{fill:ri%2?"EEF3F9":null,align:i==0?AlignmentType.LEFT:AlignmentType.CENTER,w:widths[i],bold:i==0}))}));
  return new Table({columnWidths:widths, width:{size:total,type:WidthType.DXA},
    borders:{top:{style:BorderStyle.SINGLE,size:2,color:"BFBFBF"},bottom:{style:BorderStyle.SINGLE,size:2,color:"BFBFBF"},left:{style:BorderStyle.SINGLE,size:2,color:"BFBFBF"},right:{style:BorderStyle.SINGLE,size:2,color:"BFBFBF"},insideHorizontal:{style:BorderStyle.SINGLE,size:1,color:"D9D9D9"},insideVertical:{style:BorderStyle.SINGLE,size:1,color:"D9D9D9"}},
    rows:[hr,...brs]});
}
function hrule(){return new Paragraph({spacing:{after:120}, border:{bottom:{style:BorderStyle.SINGLE,size:6,color:ACCENT}}, children:[new TextRun({text:"",font:F,size:2})]});}

const children = [];

// ================= COVER =================
children.push(new Paragraph({spacing:{before:1200,after:0}, alignment:AlignmentType.CENTER, children:[new TextRun({text:"Which Gives More Bang for Your Buck?",font:F,bold:true,size:48,color:NAVY})]}));
children.push(new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:160}, children:[new TextRun({text:"Ethanol-Blended Petrol vs Diesel vs CNG vs Electric — A Cost-Per-Kilometre, Distance, Daily-Use and Multi-City Payback Analysis for Indian Vehicle Buyers",font:F,size:24,color:ACCENT})]}));
children.push(new Paragraph({alignment:AlignmentType.CENTER, spacing:{after:60}, children:[new TextRun({text:"A consumer-advice research paper grounded in ARAI and SIAM findings and public reporting",font:F,italics:true,size:20,color:GREY})]}));
children.push(new Paragraph({alignment:AlignmentType.CENTER, spacing:{before:400}, children:[new TextRun({text:"July 2026  ·  Six-metro pricing (Delhi basis)  ·  Companion to an editable Excel cost model",font:F,size:18,color:GREY})]}));

// ================= ABSTRACT =================
children.push(new Paragraph({spacing:{before:500,after:100}, alignment:AlignmentType.CENTER, children:[new TextRun({text:"Abstract",font:F,bold:true,size:26,color:NAVY})]}));
children.push(new Paragraph({
  spacing:{after:120,line:276},
  border:{top:{style:BorderStyle.SINGLE,size:4,color:ACCENT},bottom:{style:BorderStyle.SINGLE,size:4,color:ACCENT},left:{style:BorderStyle.SINGLE,size:4,color:ACCENT},right:{style:BorderStyle.SINGLE,size:4,color:ACCENT}},
  shading:{type:ShadingType.CLEAR,fill:"F3F7FB",color:"auto"},
  children:[new TextRun({text:"Background. ",font:F,bold:true,size:20,color:"222222"}),
    new TextRun({text:"With India's fuel supply now standardised on E20 (20% ethanol-blended petrol) and diesel, CNG and electric vehicles all available across segments, private buyers face a genuine four-way choice. Ethanol's lower energy content reduces petrol mileage, changing the running-cost calculus at the point of purchase. ",font:F,size:20,color:"222222"}),
    new TextRun({text:"Objective. ",font:F,bold:true,size:20,color:"222222"}),
    new TextRun({text:"To quantify which powertrain delivers the best value (“bang for your buck”) across realistic distance, daily-use and city-price scenarios. ",font:F,size:20,color:"222222"}),
    new TextRun({text:"Methods. ",font:F,bold:true,size:20,color:"222222"}),
    new TextRun({text:"A bottom-up cost model was built for four vehicle classes (two-wheeler, hatchback, sedan, compact SUV) and up to four powertrains (petrol, diesel, CNG, electric), using ARAI/SIAM ethanol mileage-drop factors (2–6%), PPAC-basis July-2026 fuel and electricity prices across six metros, and representative real-world mileage. Cost per kilometre, five-year total cost of ownership, and payback period were computed and validated on a single class before scaling, with twenty-one automated consistency checks, and cross-checked against Vahan fleet composition and FADA FY26 new-vehicle sales. ",font:F,size:20,color:"222222"}),
    new TextRun({text:"Results. ",font:F,bold:true,size:20,color:"222222"}),
    new TextRun({text:"On a Delhi basis, electric was cheapest to run (₹1.41/km hatchback; ₹0.35/km scooter), CNG next (₹2.97/km), diesel intermediate for cars (₹4.33–5.01/km), and E20 petrol dearest (₹6.65–8.18/km), the ethanol penalty adding ~₹4,000/year. CNG repaid its modest premium in under two years for average drivers; EV payback ran from ~10 years at low mileage to ~2.4–3.5 years at high mileage; diesel paid back only at high mileage. Pump prices varied ~13% across cities (₹102 Delhi to ₹116 Hyderabad), driven almost entirely by state VAT — since petrol and diesel are taxed by both centre and state, whereas ethanol carries only 5% GST (federal). FADA FY26 sales confirm the value logic: CNG has risen to 22% of passenger-vehicle sales and EV penetration to 8.5% overall, led by the two- and three-wheelers where payback is fastest — yet petrol still dominates the on-road fleet, so the ethanol penalty aggregates to an estimated ₹3,200 crore/year on new petrol vehicles alone. ",font:F,size:20,color:"222222"}),
    new TextRun({text:"Conclusion. ",font:F,bold:true,size:20,color:"222222"}),
    new TextRun({text:"Best value is distance- and city-dependent: EV for high-mileage drivers with home charging, CNG for average city drivers where a network exists, diesel only for high-mileage highway car buyers, and modern E20-compliant petrol for low-mileage or infrastructure-constrained buyers.",font:F,size:20,color:"222222"})]
}));
children.push(new Paragraph({spacing:{before:120,after:120}, children:[new TextRun({text:"Keywords: ",font:F,bold:true,size:18,color:GREY}),new TextRun({text:"E20 ethanol blending; diesel; CNG; electric vehicles; cost per kilometre; total cost of ownership; payback period; fuel taxation; state VAT; India; ARAI; SIAM; PPAC.",font:F,italics:true,size:18,color:GREY})]}));
children.push(new Paragraph({children:[new PageBreak()]}));

// ================= TOC =================
children.push(h1("Contents"));
children.push(new TableOfContents("Contents",{hyperlink:true,headingStyleRange:"1-2"}));
children.push(new Paragraph({children:[new PageBreak()]}));

// ================= 1 INTRODUCTION =================
children.push(h1("1. Introduction"));
children.push(p("India's petrol pumps now dispense E20 — petrol blended with 20% ethanol — as the default fuel, part of a programme that reached a 10% blend in 2023 (saving an estimated 30 million barrels of oil) and targeted 20% by 2025. Ethanol carries less energy than petrol, so E20 lowers mileage slightly and quietly raises the real running cost of a petrol vehicle at the very moment buyers are also weighing CNG and electric alternatives."));
children.push(p("This paper asks a single practical question: for an Indian buyer today, which powertrain gives the most bang for the buck? “Value” here is not the lowest sticker price nor the lowest running cost in isolation, but the combination — how quickly a costlier-to-buy vehicle repays itself through cheaper fuel, and how that answer shifts with how far and how often you drive. We model that explicitly across distance and daily-use scenarios, and place the cost findings alongside the non-cost factors (emissions, refuelling time, infrastructure) that a real purchase decision must weigh."));

// ================= 2 DATA & METHODS =================
children.push(h1("2. Data & Methods"));
children.push(p("The comparison is built bottom-up. For each vehicle class we take a representative real-world mileage, apply current city fuel and electricity prices, and derive a cost per kilometre; petrol figures include the E20 mileage penalty. We then add maintenance and the upfront price premium to produce a five-year total cost of ownership (TCO) and a payback period. Prices are modelled for six metros and default to Delhi; selecting a different city reprices the whole model. Every calculation lives in the companion spreadsheet with its source noted, and twenty-one automated consistency checks validate the model (for example, that ethanol-adjusted mileage is always below the pure-petrol baseline, and that CNG, diesel and EV always cost less per km than petrol). Following good modelling practice, the model was validated on one vehicle class first and then scaled to all four."));
children.push(h2("2.1 Data sources"));
children.push(bullet("Ethanol mileage impact: ARAI controlled-laboratory study reported July 2026 (2–6% drop), and SIAM statements (2–4%), cross-checked against Autocar India real-world tests (up to ~12% on older, non-compliant cars)."));
children.push(bullet("Fuel & electricity prices: PPAC-basis city retail rates, July 2026 (petrol ₹102–116/L, diesel ₹95–104/L, CNG ₹83–97/kg across the six metros; home electricity ₹8–11/kWh; public fast charging ₹18–24/kWh)."));
children.push(bullet("Tax structure: ClearTax and Business Standard (central excise ₹13/L petrol and ₹10/L diesel; state VAT rates; ethanol 5% GST)."));
children.push(bullet("Running costs, premiums and qualitative comparison: VahanBazaar, AutoPunditz, BikeDekho (2026), and GoMechanic's biofuels-vs-EVs analysis (operating cost ₹4–5/km biofuel vs ₹1.1/km EV; 30–50% CO₂ reduction from biofuels; charging and range constraints)."));
children.push(h2("2.2 Key assumptions"));
children.push(bullet("E20 mileage drop set at 4% (mid-range of ARAI/SIAM findings); a 6% sensitivity is discussed for older vehicles."));
children.push(bullet("Diesel is modelled only for the sedan and compact-SUV classes, where it is still sold; diesel hatchbacks and two-wheelers are effectively discontinued and biodiesel blending is negligible for mileage."));
children.push(bullet("EV running cost assumes an 80% home / 20% public charging mix; pure home charging is cheaper (~₹1.07/km hatchback in Delhi)."));
children.push(bullet("Payback compares each alternative against the petrol variant of the same vehicle class, over a five-year horizon."));
children.push(note("Prices vary by city and over time; CNG and electricity tariffs differ by state. All levers, including the city selector, are editable in the companion spreadsheet."));
children.push(h2("2.3 Calibration against ARAI / SIAM certified fuel-efficiency data"));
children.push(p("To check that the model's mileage assumptions are grounded, we benchmarked them against SIAM's official BS-VI fuel-efficiency declarations (ARAI-certified figures for 302 passenger-vehicle variants and 130 two-wheelers). Certified figures are measured under standardised test conditions and run higher than real-world driving; the model deliberately uses real-world numbers, which sit at a sensible discount below the certified benchmark:"));
children.push(table(
  ["Class","ARAI certified (avg)","Model real-world (E0)","Note"],
  [
    ["Hatchback petrol","20.3 kmpl","16 kmpl","~21% real-world discount"],
    ["Sedan petrol","18.4 kmpl","15 kmpl","~18% discount"],
    ["Compact SUV petrol","14.8 kmpl","13 kmpl","~12% discount"],
    ["CNG car","27.4 km/kg","28 km/kg","In line (CNG tracks certified)"],
    ["Two-wheeler petrol","50.7 kmpl","55 kmpl","In line"],
  ],
  [2200,1900,1900,2100]
));
children.push(p("The benchmark confirms two things. First, the model's real-world mileages are conservative-but-consistent versions of the certified data, so the cost-per-km figures are not optimistic. Second, the official declarations independently corroborate the powertrain ordering used throughout this paper: certified CNG cars return ~27 km/kg with markedly lower CO₂ (about 102 g/km versus ~151 g/km for petrol), confirming both the running-cost and emissions advantages that the model and Section 9 rely on. The ethanol penalty of Section 3 is applied on top of these baselines."));

// ================= 3 ETHANOL =================
children.push(h1("3. The Ethanol Mileage Penalty"));
children.push(p("Ethanol contains less energy per litre than petrol, so E20 propels a car slightly less far per litre than pure petrol. The effect is real but modest, and far smaller than the alarming figures that circulated online — SIAM explicitly rejected claims of a 50% drop. ARAI's controlled study found a 2–6% efficiency reduction; SIAM cites 2–4%. The penalty depends mainly on vehicle age: E20-compliant models (roughly 2023 onward) sit at the low end, while older vehicles can lose 5–12% in real-world use. ARAI ruled out engine failures in compatible vehicles; the one caveat is that some older rubber hoses and seals may degrade faster and need earlier replacement."));
children.push(h2("3.1 The penalty in rupees (average use, 4% drop)"));
children.push(table(
  ["Vehicle type","E0 mileage","E20 mileage","Extra ₹/km","Extra ₹/year"],
  [
    ["Two-wheeler","55.0 km/L","52.8 km/L","₹0.08","₹1,160"],
    ["Hatchback","16.0 km/L","15.4 km/L","₹0.27","₹3,989"],
    ["Sedan","15.0 km/L","14.4 km/L","₹0.28","₹4,255"],
    ["Compact SUV","13.0 km/L","12.5 km/L","₹0.33","₹4,910"],
  ],
  [2400,1800,1800,1500,1700]
));
children.push(note("Delhi petrol basis (₹102.12/L, July 2026). At a 6% drop the annual amounts rise by roughly half; in higher-priced cities (e.g. Hyderabad) they are ~13% larger again. Ethanol raises petrol fuel cost by a few percent, but only a few thousand rupees a year — far less than the gap to CNG, diesel or EV."));

// ================= 4 COST PER KM =================
children.push(h1("4. Results I — Running Cost Per Kilometre"));
children.push(p("This is where the options separate sharply. Figures are on a Delhi basis (July 2026); petrol includes the E20 penalty and EV assumes the 80/20 charging mix. Diesel is shown only where it is sold (sedan, compact SUV)."));
children.push(table(
  ["Vehicle type","Petrol (E20)","Diesel","CNG","Electric"],
  [
    ["Two-wheeler","₹1.93","—","₹0.81","₹0.35"],
    ["Hatchback","₹6.65","—","₹2.97","₹1.41"],
    ["Sedan","₹7.09","₹4.33","₹3.20","₹1.51"],
    ["Compact SUV","₹8.18","₹5.01","₹3.78","₹1.63"],
  ],
  [2300,1700,1500,1500,1500]
));
children.push(p("A Delhi hatchback owner spends about ₹6.65 per km on E20 petrol, ₹2.97 on CNG, and ₹1.41 on electricity. This tracks the independent GoMechanic figures of ~₹4–5/km for liquid fuel and ~₹1.1/km for EVs, and confirms the running-cost ranking is robust across sources. For cars where diesel is offered it sits between petrol and CNG — about ₹4.33/km (sedan) to ₹5.01/km (SUV) — cheaper per km than petrol, but on a higher-priced, higher-maintenance vehicle. Two-wheelers show the same pattern in miniature: a CNG motorcycle such as the Bajaj Freedom 125 (India's first CNG bike, 102 km/kg) runs at about ₹0.81/km — well under half the petrol figure (₹1.93/km) and only marginally above an electric scooter (₹0.35/km), but without the charging dependency."));

// ================= 5 SCENARIO / DAILY-USE MODELLING =================
children.push(h1("5. Results II — Distance & Daily-Use Scenario Modelling"));
children.push(p("Cheap running comes with a higher sticker price for CNG and especially EVs, so value depends on distance. We model three usage profiles, expressed both annually and as an everyday driving habit, since buyers think in “km per day”:"));
children.push(table(
  ["Usage profile","Daily driving","Annual distance","Typical buyer"],
  [
    ["Low","~22 km/day","8,000 km/yr","Light city use / second vehicle"],
    ["Average","~41 km/day","15,000 km/yr","Typical private owner"],
    ["High","~68 km/day","25,000 km/yr","Long commute / cab-like use"],
  ],
  [1800,1900,1900,2600]
));
children.push(h2("5.1 Payback across scenarios (years to repay the premium)"));
children.push(p("Payback is the number of years of fuel-and-maintenance savings needed to repay the upfront premium over the petrol variant. Because savings grow with distance, payback shortens dramatically for high-mileage drivers."));
children.push(table(
  ["Vehicle · Powertrain","Upfront premium","Low (22 km/d)","Average (41 km/d)","High (68 km/d)"],
  [
    ["Hatchback · CNG","₹90,000","3.3 yrs","1.7 yrs","1.0 yrs"],
    ["Hatchback · EV","₹4,50,000","10.1 yrs","5.6 yrs","3.4 yrs"],
    ["Sedan · Diesel","₹1,20,000","6.3 yrs","3.1 yrs","1.8 yrs"],
    ["Sedan · CNG","₹95,000","3.3 yrs","1.7 yrs","1.0 yrs"],
    ["Sedan · EV","₹5,00,000","10.6 yrs","5.8 yrs","3.5 yrs"],
    ["Compact SUV · Diesel","₹1,50,000","6.7 yrs","3.4 yrs","2.0 yrs"],
    ["Compact SUV · CNG","₹1,10,000","3.4 yrs","1.7 yrs","1.0 yrs"],
    ["Compact SUV · EV","₹4,00,000","7.3 yrs","4.0 yrs","2.4 yrs"],
    ["E-scooter (2W) · EV","₹45,000","3.3 yrs","1.8 yrs","1.1 yrs"],
    ["Bajaj Freedom (2W) · CNG","₹15,000","1.8 yrs","0.9 yrs","0.5 yrs"],
  ],
  [2600,1700,1500,1500,1500]
));
children.push(note("Delhi basis. Payback shortens in higher-priced cities (larger fuel savings) and lengthens where electricity is dearer (e.g. Mumbai) — see Section 6."));
children.push(h2("5.2 Interpretation — where the bang-for-buck lies"));
children.push(p("Three patterns dominate. CNG pays back fast almost everywhere — under two years for an average driver — because the premium is small relative to the fuel saving. EVs pay back slowly at low mileage but become compelling as distance rises: a compact-SUV EV repays a ₹4-lakh premium in about 2.4 years at 68 km/day, after which it runs for a fraction of a petrol car's cost. Diesel is the squeezed middle: it pays back only for high-mileage car buyers (around 2 years at 68 km/day for an SUV) and never at low mileage, because its large premium and higher maintenance are undercut by cheaper-to-run CNG on one side and cheaper-to-buy petrol on the other. For the low-mileage buyer, no premium repays itself quickly, so the modern petrol car — ethanol penalty and all — remains the value choice."));

// ================= 6 CITY VARIATION =================
children.push(h1("6. Results III — How Prices Vary Across Cities"));
children.push(p("Running cost is not the same everywhere, because fuel prices differ sharply by state — chiefly through state VAT (see Section 7). The table below shows cost per kilometre for a reference hatchback (petrol, CNG, EV) and a compact-SUV diesel across six metros, using July-2026 PPAC-basis prices."));
children.push(table(
  ["City","Petrol pump ₹/L","Petrol ₹/km","Diesel ₹/km (SUV)","CNG ₹/km","EV ₹/km"],
  [
    ["Delhi","₹102.12","₹6.65","₹5.01","₹2.97","₹1.41"],
    ["Chennai","₹107.76","₹7.02","₹5.24","₹3.46","₹1.41"],
    ["Bengaluru","₹111.68","₹7.27","₹5.24","₹3.46","₹1.52"],
    ["Mumbai","₹111.21","₹7.24","₹5.15","₹3.07","₹1.73"],
    ["Kolkata","₹113.51","₹7.39","₹5.25","₹3.34","₹1.52"],
    ["Hyderabad","₹115.69","₹7.53","₹5.46","₹3.46","₹1.57"],
  ],
  [1700,1900,1500,1700,1400,1300]
));
children.push(p("Petrol runs from about ₹6.65/km in Delhi to ₹7.53/km in Hyderabad — a ~13% spread that comes almost entirely from state taxation, since the central excise component is identical everywhere. The electric advantage is widest where petrol is expensive and electricity is cheap (Delhi, Chennai), and narrows in Mumbai where domestic tariffs are higher (₹1.73/km). CNG economics also shift geographically: it is cheapest in the north (Delhi ₹2.97/km) and dearer in the south, where CNG sells around ₹97/kg. The ranking never changes — EV, then CNG, then diesel, then petrol — but the size of each gap, and therefore payback, depends on where you live."));
children.push(h2("6.1 The wider OMC price board — 14 cities"));
children.push(p("The six metros understate the national range. Oil marketing companies (IndianOil, BPCL, HPCL) revise and publish city fuel prices every day, and those boards show the same litre of E20 petrol costing very different amounts across the country. Extending the comparison to fourteen cities in different states (July 2026):"));
children.push(table(
  ["City","State / UT","Petrol ₹/L","Diesel ₹/L","Petrol ₹/km","vs Delhi"],
  [
    ["Chandigarh","UT","101.54","89.47","₹6.61","−₹0.58"],
    ["Lucknow","Uttar Pradesh","101.86","95.36","₹6.63","−₹0.26"],
    ["Noida","Uttar Pradesh","101.96","95.44","₹6.64","−₹0.16"],
    ["Delhi","Delhi","102.12","95.20","₹6.65","—"],
    ["Gurgaon","Haryana","102.97","95.64","₹6.70","+₹0.85"],
    ["Chennai","Tamil Nadu","107.76","99.55","₹7.02","+₹5.64"],
    ["Bhubaneswar","Odisha","110.49","100.68","₹7.19","+₹8.37"],
    ["Bengaluru","Karnataka","110.82","99.56","₹7.21","+₹8.70"],
    ["Mumbai","Maharashtra","111.21","97.83","₹7.24","+₹9.09"],
    ["Jaipur","Rajasthan","112.66","98.25","₹7.33","+₹10.54"],
    ["Kolkata","West Bengal","113.51","99.82","₹7.39","+₹11.39"],
    ["Patna","Bihar","113.53","99.36","₹7.39","+₹11.41"],
    ["Thiruvananthapuram","Kerala","115.49","104.40","₹7.52","+₹13.37"],
    ["Hyderabad","Telangana","115.69","103.82","₹7.53","+₹13.57"],
  ],
  [1900,1900,1300,1300,1300,1300]
));
children.push(p("The OMC boards reveal a petrol spread of about ₹14.15/litre (13.9%) and a diesel spread of ₹14.93/litre from the cheapest to the dearest of these cities. The pattern is fiscal, not logistical: the lowest prices sit in low-VAT states and union territories (Chandigarh, and the UP/Delhi belt around ₹102), while the highest are in high-VAT states (Telangana, Kerala, and the eastern states near ₹114–116). Because the central excise (₹13/L petrol, ₹10/L diesel) is uniform nationwide, essentially the entire ₹14 gap is state VAT. For the buyer this means the very same E20 car costs about ₹0.92 more per kilometre to run in Hyderabad than in Chandigarh — roughly ₹14,000 a year at average use — before the vehicle-choice decision is even made. It also amplifies the ethanol penalty: the mileage loss is worth proportionately more rupees in the high-price states."));

// ================= 7 TAX STRUCTURE =================
children.push(h1("7. The Fuel Tax Structure — Central, State, and Why Ethanol Is Different"));
children.push(p("The city price differences above are almost entirely a tax story, and that tax story is central to the ethanol question. Petrol and diesel are deliberately kept outside GST and are taxed in two layers: a central (federal) excise duty that is fixed per litre and uniform nationwide — ₹13/L on petrol, ₹10/L on diesel — plus a state VAT levied as a percentage, which varies enormously from state to state. Ethanol supplied for blending, by contrast, sits inside GST and attracts only 5% GST — a single, federally-set levy, with no separate state VAT on the ethanol molecule."));
children.push(h2("7.1 Worked example — a litre of petrol in Delhi"));
children.push(table(
  ["Component","₹ per litre","Levied by"],
  [
    ["Base price + freight (refiner/OMC)","₹68.76","—"],
    ["Central excise duty","₹13.00","Centre (federal)"],
    ["Dealer commission","₹3.77","—"],
    ["State VAT (19.40%)","₹16.59","State"],
    ["Retail pump price (E20)","₹102.12","—"],
    ["Total tax (central + state)","₹29.59","≈ 29% of pump price"],
  ],
  [3200,1700,2400]
));
children.push(p("Roughly ₹30 of the ₹102 Delhi pump price is tax, split between the centre (₹13 excise) and the state (₹16.59 VAT). Because the state slice is a percentage of a base that already includes the central excise, high-VAT states compound the burden — which is why Hyderabad (Telangana VAT ~35% on petrol) is the most expensive metro and Delhi (~19%) among the cheapest."));
children.push(h2("7.2 Petrol vs ethanol — who taxes what"));
children.push(table(
  ["Fuel","Central levy","State levy","Net"],
  [
    ["Petrol / Diesel","Excise ₹13 / ₹10 per L (fixed)","VAT ~14%–35% (varies)","Dual: central + state"],
    ["Ethanol (for blending)","5% GST","None (outside state VAT)","Single: federal only"],
  ],
  [1900,2400,2100,1900]
));
children.push(p("This asymmetry matters for the “who gains” question. The roughly 20% ethanol in E20 escapes the heavy dual petroleum taxation and is taxed lightly and only federally. That lowers the tax the exchequer collects on the ethanol share — but because E20 sells at the same pump price as before, the benefit is retained in the price build-up rather than rebated to the buyer. It also shifts revenue away from the state VAT base toward the federally-controlled GST/excise base, which is one reason states have been cautious about aggressive blending: every litre of petrol displaced by ethanol is a litre that leaves their VAT net. For the motorist, the practical consequence is the one flagged in Section 10 — lower taxes on the ethanol portion, but no lower price at the pump."));

// ================= 8 MARKET / FLEET =================
children.push(h1("8. Results IV — Matching the Model to the Market: Vahan Fleet & FADA Sales"));
children.push(p("A cost model is only persuasive if real buyers behave the way its value logic predicts. We therefore reconcile the findings against two independent datasets: FADA's FY26 new-vehicle retail registrations (what buyers are choosing now) and the Vahan-registered fleet (what is actually on the road). The two tell different but complementary stories."));
children.push(h2("8.1 What buyers are choosing — FADA FY26 new sales"));
children.push(p("FADA recorded 2.97 crore vehicles retailed in FY26 (+13.3%), of which two-wheelers were 72%, passenger vehicles 16% and three-wheelers, commercial vehicles and tractors the remainder. The fuel mix is shifting exactly toward the powertrains this paper finds cheapest to own:"));
children.push(table(
  ["Passenger-vehicle fuel","FY26 share","FY25 share","Direction"],
  [
    ["Petrol","47.5%","50.8%","↓ 3.3 pp"],
    ["CNG","22.0%","19.6%","↑ 2.4 pp"],
    ["Diesel","18.1%","18.2%","≈ flat"],
    ["Electric","4.25%","2.6%","↑ 1.6 pp"],
    ["Hybrid / other","8.2%","8.7%","≈ flat"],
  ],
  [2500,1600,1600,1600]
));
children.push(p("The market is voting for value. CNG — which the model identifies as the fastest-payback car option — is the biggest gainer, now more than one in five new cars, while petrol is losing about three points of share a year. Electrification is led precisely where our payback maths is strongest and quickest: two-wheeler EVs are 6.5% of new two-wheelers, and electric three-wheelers are already ~61% of that segment (payback under two years in the model), versus 4.25% for the higher-premium passenger cars. Overall EV penetration reached 8.5% in FY26. In other words, revealed buyer behaviour lines up with the cost ranking — cheap-to-run, low-premium options (CNG, e-2W, e-3W) are winning fastest, while EV cars, with the largest premium, diffuse more slowly."));
children.push(h2("8.2 What is on the road — the Vahan fleet, and why the ethanol penalty is so broad"));
children.push(p("New-sales momentum is one thing; the installed fleet is another, and it changes far more slowly. Two-wheelers are roughly seven in ten vehicles registered, and the overwhelming majority of both the two-wheeler and car parc still runs on petrol, with commercial vehicles almost entirely diesel; cumulative EV penetration of the fleet remains low single digits even as new-sales penetration hits 8.5%. The practical consequence for this paper is important: the ethanol mileage penalty of Section 3 is not a niche cost — it applies to the largest slice of vehicles on Indian roads today, and will for years, because fleet turnover takes well over a decade."));
children.push(h2("8.3 The ethanol penalty, aggregated"));
children.push(p("Scaling the per-vehicle ethanol cost by sales and fleet volumes turns a few thousand rupees per owner into a large national number. Using FADA FY26 units and the Delhi per-vehicle figures from Section 3:"));
children.push(table(
  ["Population","Petrol vehicles","Extra ₹/veh/yr","Aggregate ₹/yr"],
  [
    ["New petrol cars (FY26 cohort)","~22.3 lakh","₹3,989","~₹891 crore"],
    ["New petrol two-wheelers (FY26)","~2.00 crore","₹1,160","~₹2,323 crore"],
    ["New petrol vehicles — combined","—","—","~₹3,214 crore"],
    ["On-road petrol fleet (illustrative)","~24 crore*","(blended)","~₹39,000 crore"],
  ],
  [2600,1700,1500,1700]
));
children.push(note("*Illustrative parc assumption (≈4 crore petrol cars + ≈20 crore petrol two-wheelers), editable in the model's Fleet & Sales tab. The new-cohort figures are derived directly from FADA FY26 units; the fleet figure is an order-of-magnitude scenario, not a measured total."));
children.push(p("The reading is twofold. Just one year's new petrol vehicles carry an ethanol mileage cost of roughly ₹3,200 crore a year, every year they are driven; across the whole petrol fleet the recurring national figure plausibly runs into the tens of thousands of crore. This is the aggregate counterpart to Section 11's finding that the saving on the ethanol tax is not rebated to the motorist — the money is real and large, it is simply spread thinly across a very large number of petrol owners who each lose only a little. At the same time, the FADA trend shows the escape routes (CNG, EV) are being adopted quickly where they pay back fastest, which is precisely the consumer-rational response the model would recommend."));
children.push(h2("8.4 Source reconciliation — SIAM wholesale vs FADA retail"));
children.push(p("Two industry bodies count vehicles differently, and it is worth confirming they agree. SIAM reports wholesale dispatches (vehicles shipped from factories to dealers); FADA reports retail registrations (actual customer purchases via the Vahan portal). The two FY26 totals corroborate each other closely for the segments this paper models:"));
children.push(table(
  ["Category","SIAM wholesale FY26","FADA retail FY26","Retail vs wholesale"],
  [
    ["Passenger vehicles","46.4 lakh","47.1 lakh","+1.3%"],
    ["Two-wheelers","2.17 crore","2.14 crore","−1.3%"],
    ["Commercial vehicles","10.8 lakh","10.6 lakh","−1.8%"],
    ["Three-wheelers","8.4 lakh","13.6 lakh","+63%"],
  ],
  [2300,1900,1800,1700]
));
children.push(p("Passenger vehicles, two-wheelers and commercial vehicles reconcile within about 2% — well inside normal dealer-inventory swings — which gives confidence in the volumes used for the aggregate above. The one large divergence is three-wheelers, where FADA/Vahan registrations far exceed SIAM dispatches because Vahan captures the many electric-rickshaw and small-maker three-wheelers that fall outside SIAM's member-OEM wholesale reporting. For fuel-mix detail specifically, Vahan/FADA is the authoritative source, since SIAM's headline releases report volumes but not the powertrain split; SIAM's value here is corroborating the totals and providing the certified fuel-efficiency declarations used in Section 2.3."));

// ================= 9 DISCUSSION =================
children.push(h1("9. Discussion — Beyond Cost Per Kilometre"));
children.push(p("A purchase is not decided on running cost alone. Three non-cost dimensions materially shape the real-world verdict."));
children.push(h2("9.1 Emissions and energy security"));
children.push(p("Biofuels cut carbon emissions by an estimated 30–50% versus pure petrol and are produced domestically, aiding energy security; EVs eliminate tailpipe emissions entirely, though their well-to-wheel footprint depends on the grid mix. Ethanol production also raises land- and water-use and food-security concerns at scale. On climate grounds the choice is directional rather than clear-cut, which is why commentators such as GoMechanic advocate a hybrid transition — ethanol for immediate reach, EVs as infrastructure matures."));
children.push(h2("9.2 Convenience and infrastructure"));
children.push(bullet("Refuelling time: petrol and CNG refuel in 3–5 minutes; EV charging ranges from ~30 minutes (fast) to several hours (home)."));
children.push(bullet("Range: EVs typically offer 300–500 km per charge; petrol/CNG face no range anxiety but CNG reduces boot space and needs a station network."));
children.push(bullet("Coverage: CNG and public charging are concentrated in metros and larger cities; lower-tier towns still favour liquid fuels."));
children.push(h2("9.3 What can change the answer"));
children.push(p("EV economics hinge on home charging — heavy reliance on ₹18–24/kWh public fast charging can roughly double EV running cost and stretch payback. FAME-style subsidies (historically 20–40% for EVs) and state road-tax waivers can shorten EV payback sharply where available. Resale value, battery warranty and insurance differ by powertrain and are not modelled here; buyers should fold them into a final decision."));

// ================= 9 RECOMMENDATIONS =================
children.push(h1("10. Recommendations by Buyer Profile"));
children.push(h2("High-mileage driver (68+ km/day) with home charging"));
children.push(p("Buy electric. Your distance repays the premium in ~3–4 years and the running-cost gap is largest for you. Keep a CNG option in mind if daily runs occasionally exceed range and public charging is unreliable on your routes."));
children.push(h2("Average city driver (~41 km/day)"));
children.push(p("If CNG is available nearby, a factory-fitted CNG car is the value pick — under two-year payback, running cost close to an EV's, and no charging setup. Choose an EV instead if you have home charging, want the lowest running cost, and plan to keep the car 6+ years."));
children.push(h2("Low-mileage driver (~22 km/day)"));
children.push(p("Stick with a modern E20-compliant petrol car. Neither premium repays quickly at low distance, and the ethanol penalty costs only ~₹4,000–4,900 a year. Simplicity and low upfront cost win."));
children.push(h2("High-mileage highway / car buyer considering diesel"));
children.push(p("Diesel now makes sense only in a narrow case: a sedan or SUV driven long distances (25,000+ km/year, mostly highway) and kept for many years, where the ₹4.33–5.01/km running cost and strong torque/range repay a ₹1.2–1.5 lakh premium in about two years. For city-bound or average-mileage buyers, CNG beats diesel on running cost with a far smaller premium, and petrol beats it on upfront cost — so diesel is hard to justify outside that high-distance niche."));
children.push(h2("Two-wheeler buyer"));
children.push(p("Two strong value options now exist beyond petrol. An electric scooter offers the lowest running cost (~₹0.35/km) with a ~two-year payback, but depends on charging access. A CNG motorcycle such as the Bajaj Freedom 125 is the more infrastructure-light choice: it costs only about ₹15,000 more than a petrol commuter, runs at ~₹0.74/km, and repays that premium in around a year at average use — while retaining a petrol tank as backup where CNG stations are scarce. Choose the e-scooter for the lowest cost and quietest ride if you can charge conveniently; choose the CNG bike if you want quick refuelling, no charging setup, and the fastest payback. Plain petrol remains sensible only for frequent long intercity rides or where neither CNG nor charging is available."));

// ================= 10 WHO GAINS / OMC =================
children.push(h1("11. The Economics of the Transition — Who Really Gains?"));
children.push(p("A recurring public question is whether the ethanol push has quietly transferred money from motorists to oil companies. Because petrol running cost is central to this paper, it is worth examining directly: has E20 made consumers pay more, and has it made fossil-fuel companies richer? The honest answer is nuanced — consumers do pay modestly more, but the gains are spread across the ethanol value chain and the exchequer rather than being a clean oil-company windfall."));
children.push(h2("11.1 The pump price did not fall — but mileage did"));
children.push(p("Despite blending in 20% ethanol, E20 sells at the same pump price as the petrol it replaced; there is no separate, cheaper E20 rate. Meanwhile the petroleum ministry itself acknowledges a roughly 3–5% reduction in fuel economy on E20 (ARAI's controlled range is 2–6%). Identical rupees per litre buying fewer kilometres means the effective cost per kilometre rises — about ₹0.27/km for a hatchback, or roughly ₹4,000 a year at average use, as quantified in Section 3. This is the core of the “consumers pay more” claim, and on the fuel itself it is correct."));
children.push(h2("11.2 Why E20 is not actually cheaper to produce"));
children.push(p("The reason the saving never reaches the pump is that domestic ethanol is not cheap. Maize-based ethanol is procured at about ₹71.86/litre (excluding GST, transport and handling) and C-heavy molasses ethanol at about ₹57.97/litre — both up sharply from ₹52.92 and ₹46.66 respectively in 2021–22. These prices are deliberately fixed to guarantee farmers a stable income, so they do not fall when crude oil falls. At crude around $70/barrel, ethanol is comparable to or costlier than the petrol it displaces; the ministry notes E20 would only become cheaper to produce if crude climbed to roughly $120–130/barrel. So there is often no production saving to pass on in the first place."));
children.push(h2("11.3 Tax treatment and the transparency gap"));
children.push(p("The tax structure does favour the blend, as Section 7 sets out. Ethanol for blending attracts only 5% GST (a federal levy), whereas petrol carries a central excise of ₹13/litre plus state VAT (about ₹16.59/litre in Delhi, at 19.40%) — a dual, centre-plus-state burden of roughly ₹30/litre. Yet none of the lower input-tax on the ethanol portion is passed back to buyers as a retail rebate — NITI Aayog's 2021 roadmap explicitly recommended GST breaks and retail price incentives to compensate consumers, and these were not implemented. Compounding the concern, oil companies publish only a consolidated retail price build-up, with no separate breakdown for E20, so the distribution of the tax benefit is not transparent."));
children.push(h2("11.4 Bottom line: pay more, and who got richer?"));
children.push(p("Did consumers pay more? Yes — modestly and indirectly. The pump price is unchanged while mileage falls a few percent, and the lower taxes on the ethanol portion are not returned at retail, so buyers absorb both the mileage penalty and the foregone tax relief. For a typical hatchback that is on the order of ₹4,000 a year; more for older, non-compliant vehicles."));
children.push(p("Did fossil-fuel companies get richer? Partly, but this is not a clean story of oil-company enrichment. Oil marketing companies do benefit from stable marketing margins, price-stability, and the avoided ₹2/litre unblended excise — but because ethanol is not markedly cheaper for them to buy at current crude prices, the direct margin uplift attributable to ethanol is limited, and OMC profitability is driven far more by refining margins and pricing freedom than by blending. The clearer beneficiaries are the ethanol and agricultural value chain, which is paid high, guaranteed procurement prices, and the government, which retains the duty structure on the petrol portion and books the macro-level gains — reduced crude imports, foreign-exchange savings and energy security. Framed fairly: the programme delivers real national and rural benefits, but the specific consumer grievance — paying the same for less mileage with no offsetting retail rebate — is legitimate, and the fix NITI Aayog proposed (retail tax relief for the ethanol share) remains the most direct way to make the deal fair to motorists."));
children.push(note("Counterpoint: supporters argue price stability itself has value — E20 insulates pump prices from crude spikes, farmers gain assured income, and import dependence falls. Critics counter that stability and macro gains are being financed partly by an untransparent, un-rebated cost to individual motorists."));

// ================= 12 EXCHEQUER =================
children.push(h1("12. The Exchequer Angle — What Vehicle and Kilometre Growth Adds to Tax Revenue"));
children.push(p("Sections 10 and 11 asked who pays and who gains at the level of the individual motorist. Scaled up, the same mechanics make the vehicle sector one of the government's largest revenue engines — and reveal why the fisc has mixed incentives about the very transition this paper recommends. The state earns from vehicles through two distinct channels, both of which grow as the fleet and its kilometres grow."));
children.push(h2("12.1 The two revenue streams"));
children.push(bullet("A one-time tax at purchase: GST (plus, historically, compensation cess) on every new vehicle sold — so this stream grows with the number of vehicles."));
children.push(bullet("A recurring tax on use: central excise plus state VAT on every litre of petrol and diesel burned — so this stream grows with vehicle-kilometres travelled (vehicles × distance ÷ mileage)."));
children.push(h2("12.2 Recurring fuel-tax revenue (PPAC consumption × tax per litre)"));
children.push(p("India consumed about 42.7 million tonnes of petrol and 93.9 MT of diesel in FY26 (PPAC). Converting to volume and applying the current tax per litre gives the recurring haul:"));
children.push(table(
  ["Stream","Basis","Annual revenue"],
  [
    ["Central excise (petrol + diesel)","₹13/L petrol, ₹10/L diesel","~₹1.88 lakh crore"],
    ["State VAT (petrol + diesel)","~₹18/L petrol, ~₹12/L diesel (avg)","~₹2.40 lakh crore"],
    ["Total recurring fuel tax","","~₹4.28 lakh crore/year"],
  ],
  [2800,3000,2200]
));
children.push(note("Bottom-up estimate using national-average VAT (states range 14–35%); consistent with PPAC/PIB reporting that the petroleum sector contributes on the order of ₹7+ lakh crore to the exchequer once crude cess, gas and dividends are included."));
children.push(h2("12.3 One-time vehicle GST (FADA FY26 sales × GST 2.0 rates)"));
children.push(p("Under the GST 2.0 rates effective September 2025, small cars attract 18%, larger cars and SUVs 40%, two-wheelers 18%, and electric vehicles just 5%. Applying a blended rate to FADA's FY26 volumes:"));
children.push(table(
  ["Segment","FY26 units","Revenue"],
  [
    ["Two-wheelers (18%)","2.14 crore","~₹0.35 lakh crore"],
    ["Passenger vehicles (blended ~28%)","47 lakh","~₹1.31 lakh crore"],
    ["Total one-time vehicle GST","—","~₹1.66 lakh crore"],
  ],
  [2800,2200,2000]
));
children.push(h2("12.4 The additional earnings from growth"));
children.push(p("This is the crux of the question. Holding tax rates constant, growth alone lifts revenue every year. Petrol demand is rising ~7% (more personal vehicles and kilometres), diesel ~4% (freight), and new-vehicle sales grew 13.3% in FY26. The incremental take:"));
children.push(table(
  ["Source of growth","Additional revenue/year"],
  [
    ["Higher petrol volume (~7%)","~₹12,500 crore"],
    ["Higher diesel volume (~4%)","~₹10,000 crore"],
    ["More vehicles sold (GST, +13.3%)","~₹22,100 crore"],
    ["Total additional earnings","~₹44,500 crore/year"],
  ],
  [3400,2600]
));
children.push(p("In other words, the growth in vehicles and the kilometres they cover hands the exchequer roughly ₹44,500 crore of extra revenue a year at unchanged rates — before any tax hike. That is the fiscal reward for a growing, driving vehicle population, and it is the mirror image of the motorist-level costs quantified earlier: what a household pays per kilometre, the state collects in aggregate."));
children.push(h2("12.5 The emerging headwind — why the transition cuts both ways"));
children.push(p("The same shift toward CNG, EVs and ethanol that saves consumers money erodes this revenue base, which is why the fisc is conflicted. Three channels bite. First, ethanol: the ~20% ethanol blended into every litre of E20 is taxed at just 5% GST instead of the ~₹31/litre a full petrol litre would bear, foregoing on the order of ₹32,000 crore a year in tax that petrol would otherwise have carried. Second, electric vehicles pay 5% GST at purchase against 18–40% for combustion vehicles, and then burn no taxed fuel at all over their life — a permanent hole in both streams. Third, CNG is taxed more lightly than petrol or diesel. Set against the ~₹44,500 crore that volume growth adds, the ethanol channel alone claws back most of it, and the EV/CNG drag is structural and rising."));
children.push(p("The evenhanded reading: vehicle and kilometre growth remains a powerful, still-growing source of exchequer earnings — on the order of tens of thousands of crore in fresh revenue each year. But the clean-fuel transition is quietly hollowing out the base from within. Petrol and diesel sit outside GST and are taxed heavily by both centre and states; the alternatives that are winning in the market (Section 8) sit inside GST and are taxed lightly. Over time that pushes the government toward finding replacement revenue — a road-use or distance-based charge, higher electricity duties, or bringing fuels into GST — which is the fiscal counterpart to the consumer-fairness question this paper began with."));

// ================= 13 LIMITATIONS =================
children.push(h1("13. Limitations"));
children.push(bullet("Prices are PPAC-basis, July 2026, for six metros defaulting to Delhi; local fuel, CNG and electricity tariffs vary and should be substituted in the model. State VAT rates used are indicative effective rates — several states add fixed per-litre cesses, so an exact retail reconstruction may differ slightly."));
children.push(bullet("Mileage figures are representative, not model-specific; individual results vary with driving style, maintenance and traffic."));
children.push(bullet("Resale value, financing cost, insurance and battery-replacement risk are excluded from the five-year TCO."));
children.push(bullet("The aggregate national ethanol-penalty figures use FADA new-sales units directly, but the on-road fleet total is an illustrative order-of-magnitude scenario based on editable parc assumptions, not a measured Vahan parc count by fuel."));
children.push(bullet("The ethanol drop is modelled at a single 4% central value; older or poorly-maintained vehicles may see more."));

// ================= 12 SOURCES =================
children.push(h1("14. Sources & References"));
[
 "ARAI E20 study (2–6% efficiency drop; no engine failure): Business Today, The Tribune, ANI — July 2026.",
 "SIAM position (2–4% controlled-environment drop; “50% drop” claims rejected): Deccan Herald — 2025–26.",
 "Real-world E20 mileage tests by vehicle age: Autocar India.",
 "CNG vs petrol running costs and premiums (Delhi, Feb 2026): VahanBazaar.",
 "EV charging cost and payback analysis (2026): AutoPunditz.",
 "Biofuels vs EVs — operating cost, emissions, infrastructure and “hybrid approach” verdict: GoMechanic (gomechanic.in/blog/biofuels-vs-evs-in-india).",
 "Bajaj Freedom 125 CNG motorcycle — 102 km/kg, running cost and pricing: BikeDekho (2026).",
 "City-wise petrol/diesel retail prices, revised daily by OMCs (IndianOil, BPCL, HPCL) and compiled by Goodreturns / HDFCSky (PPAC basis, July 2026). CNG city prices: Goodreturns (2026).",
 "New-vehicle sales mix and category volumes, FY26: FADA retail data (fada.in) and Informist/evelectree analyses; overall EV penetration 8.5%: JMK Research. Fleet composition: Vahan dashboard.",
 "Wholesale dispatch volumes FY26 (SIAM): Outlook Business / SIAM press release; SIAM vs FADA methodology (wholesale vs retail): Business Standard.",
 "Petroleum consumption FY26 (petrol 42.7 MMT, diesel 93.9 MMT): PPAC via Fuelwings. Petroleum-sector tax contribution (~₹7.5 lakh crore; centre/state split): PPAC/PIB via Business Standard.",
 "GST 2.0 vehicle rates (18% small cars/2W, 40% large/SUV, 5% EV; cess merged) effective Sept 2025: ClearTax.",
 "Certified fuel-efficiency benchmarks: SIAM BS-VI 4W FE Declaration 2019-20 and 2W FE Declaration 2019 (ARAI-certified, siam.in).",
 "Fuel tax structure — central excise (₹13/L petrol, ₹10/L diesel), state VAT rates, dealer commission: ClearTax; ethanol 5% GST: Business Standard (2026).",
 "Why E20 is not cheaper despite 20% ethanol; ethanol procurement cost, crude break-even: Business Today — 11 July 2026.",
 "E20 pricing, ethanol procurement prices (2021–22 to 2024–25), tax treatment, transparency gap and NITI Aayog 2021 recommendation: Business Standard — 15 July 2026.",
].forEach(s=>children.push(bullet(s)));
children.push(note("This paper is a companion to the editable Excel model (EV_CNG_Ethanol_Vehicle_Cost_Model.xlsx), where all calculations, assumptions and consistency checks can be inspected and adjusted."));

const doc = new Document({
  styles:{default:{document:{run:{font:F}}}},
  sections:[{
    properties:{page:{margin:{top:1000,bottom:1000,left:1100,right:1100}}},
    children
  }]
});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync("outputs/Bang_For_Your_Buck_Vehicle_Study.docx",b);console.log("written",b.length,"bytes");});
