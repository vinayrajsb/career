const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,AlignmentType,LevelFormat,BorderStyle,TabStopType,ExternalHyperlink}=require('docx');
const ACC="1F4E79", F="Calibri";
const t=(text,o={})=>new TextRun({text,font:F,size:o.size||19,bold:o.bold,italics:o.it,color:o.color,highlight:o.hl});
const p=(runs,o={})=>new Paragraph({children:Array.isArray(runs)?runs:[runs],spacing:{after:o.after??60,before:o.before??0},alignment:o.align,tabStops:o.tabs});
const H=(s)=>new Paragraph({children:[t(s.toUpperCase(),{bold:true,size:21,color:ACC})],spacing:{before:140,after:60},border:{bottom:{style:BorderStyle.SINGLE,size:6,color:ACC,space:2}}});
const b=(runs)=>new Paragraph({children:Array.isArray(runs)?runs:[t(runs)],numbering:{reference:"b",level:0},spacing:{after:40}});
const RT=[{type:TabStopType.RIGHT,position:10546}];
const role=(title,dates)=>p([t(title,{bold:true}),t("\t"+dates,{color:"555555"})],{tabs:RT,after:20,before:80});
const org=(s)=>p(t(s,{it:true,color:"444444"}),{after:40});
const ph=(s)=>t(s,{hl:"yellow"});
const kids=[
 p(t("VINAYRAJ BANGERA",{bold:true,size:36,color:ACC}),{align:AlignmentType.CENTER,after:20}),
 p(t("Aviation Maintenance Leader  |  Engine Condition Trend Monitoring  |  Aviation Data & Predictive Maintenance",{size:21,bold:true}),{align:AlignmentType.CENTER,after:40}),
 p([t("Abu Dhabi, UAE  |  +971 50 228 9259  |  vinayraj.bangera@gmail.com  |  "),new ExternalHyperlink({link:"https://www.linkedin.com/in/vinayraj-bangera/",children:[new TextRun({text:"linkedin.com/in/vinayraj-bangera",font:F,size:19,color:ACC})]})],{align:AlignmentType.CENTER,after:60}),
 H("Profile"),
 p(t("Licensed AW139 engineer and maintenance supervisor with 15 years on helicopters, the last four in Abu Dhabi. Supervise a 90-person maintenance team across line, phase and outstation operations for a 14-aircraft fleet. FlightSafety-certified in Engine Condition Trend Monitoring. Years on the hangar floor have shown me how much value sits in aviation data, and I want to apply my engineering expertise to aviation data, predictive maintenance and digital product roles where that value can be put to best use."),{after:60}),
 H("Core strengths"),
 b([t("Engine health & trend monitoring: ",{bold:true}),t("ECTM (PT6C-67C), parameter trending, exceedance and fault diagnosis")]),
 b([t("Data-minded maintenance: ",{bold:true}),t("see the value of aviation data for safer, earlier and better-planned maintenance")]),
 b([t("Leadership & operations: ",{bold:true}),t("90-person team, multi-base operations, spares and facilities, heavy inspections")]),
 b([t("Quality & compliance: ",{bold:true}),t("ISO internal auditor (TÜV NORD), DGCA and EASA regulatory environment")]),
 H("Experience"),
 role("Maintenance Supervisor (Line, Phase & Outstation), AW139","Jul 2024 – Present"),
 org("Global Aerospace and Logistics, Abu Dhabi"),
 b("Supervise a 90-person maintenance team across line, phase and outstation operations for a 14-aircraft AW139 fleet."),
 b("Run outstation bases end to end: facilities, spare parts provisioning and expanding base capability."),
 b("Carry out regular audits and random inspections to keep bases compliant and safety standards consistent."),
  role("Aircraft Maintenance Engineer, AW139","Jan 2022 – Jul 2024"),
 org("Global Aerospace and Logistics, Abu Dhabi"),
 b("Carried out phase and line maintenance on a 14-aircraft AW139 fleet with a strong focus on safety."),
 b("Worked on more than three 4-yearly / 8,000-hour inspections and led two of them."),
 role("Licensed Aircraft Maintenance Engineer (DGCA B1.3)","Apr 2012 – Jan 2022"),
 org("Heligo Charters Pvt. Ltd., Mumbai"),
 b("Progressed from aircraft technician to certifying engineer over a decade of line and major maintenance."),
 b("Certified aircraft for offshore, ferry and VIP operations, mostly at outstations, working independently away from main base."),
 role("Aircraft Apprentice","Feb 2011 – Apr 2012"),
 p(t("Raymond Ltd. Aviation Division, Mumbai: line and flight line maintenance on Eurocopter AS355N and Bell 206 L3.",{it:true,color:"444444"}),{after:40}),
 H("Licences & certifications"),
 b("DGCA ICAO Type II B1.3 licence, type rated AW139 (PT6C-67C)"),
 b("EASA B1.3 Basic licence"),
 b("AW139 and PT6C-67C type training, Leonardo Helicopters Training Centre, Italy"),
 b("Engine Condition Trend Monitoring, FlightSafety International"),
 b("ISO Internal Auditor, TÜV NORD"),
 H("Education"),
 role("B.Sc. Industrial Science, Thakur Institute of Aviation Technology","2011 – 2014"),
 p(t("Business mathematics and statistics, quality assurance, inventory management, entrepreneurship development, manufacturing processes.",{color:"444444"}),{after:40}),
 role("Aircraft Maintenance Engineering, Thakur Institute of Aviation Technology","2008 – 2011"),
 p(t("Specialisation in jet engines and helicopter airframes.",{color:"444444"}),{after:40}),
 p([t("Languages: ",{bold:true}),t("English (fluent), Hindi (native)   |   "),t("Interests: ",{bold:true}),t("Freediving, scuba diving, hiking, squash, reading, music")],{before:80}),
];
const doc=new Document({numbering:{config:[{reference:"b",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.LEFT,style:{paragraph:{indent:{left:360,hanging:220}}}}]}]},
 sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:560,bottom:560,left:680,right:680}}},children:kids}]});
Packer.toBuffer(doc).then(buf=>fs.writeFileSync("CV_Vinayraj_Bangera_Aviation_Data.docx",buf));
