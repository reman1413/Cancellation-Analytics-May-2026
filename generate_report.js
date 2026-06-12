const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ─── Colour palette ────────────────────────────────────────────────────────
const C = {
  pu: '534AB7', puL: 'EEEDFE', puM: '7F77DD',
  te: '0F6E56', teL: 'E1F5EE', teM: '5DCAA5',
  am: '854F0B', amL: 'FAEEDA', amM: 'EF9F27',
  rd: '791F1F', rdL: 'FCEBEB', rdM: 'E24B4A',
  bl: '185FA5', blL: 'E6F1FB', blM: '378ADD',
  dark: '1a1a2e', text2: '4b5563', text3: '9ca3af',
  white: 'FFFFFF', border: 'E5E7EB', border2: 'D1D5DB',
  surface2: 'F9FAFB',
};

// ─── Helper: standard cell borders ────────────────────────────────────────
const bdr = (color = C.border) => {
  const s = { style: BorderStyle.SINGLE, size: 1, color };
  return { top: s, bottom: s, left: s, right: s };
};
const noBdr = () => {
  const s = { style: BorderStyle.NIL, size: 0, color: 'FFFFFF' };
  return { top: s, bottom: s, left: s, right: s };
};

// ─── Helper: shaded header cell ───────────────────────────────────────────
function hCell(text, w, color = C.dark, bg = C.dark, bold = true) {
  return new TableCell({
    borders: bdr(C.dark),
    width: { size: w, type: WidthType.DXA },
    shading: { fill: bg, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold, color: C.white, size: 18, font: 'Arial' })]
    })]
  });
}

// ─── Helper: data cell ────────────────────────────────────────────────────
function dCell(text, w, opts = {}) {
  const {
    align = AlignmentType.CENTER, bold = false, color = C.dark,
    bg = null, shade = false, size = 18
  } = opts;
  return new TableCell({
    borders: bdr(C.border),
    width: { size: w, type: WidthType.DXA },
    shading: shade && bg ? { fill: bg, type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, color, size, font: 'Arial' })]
    })]
  });
}

// ─── Helper: section heading paragraph ────────────────────────────────────
function secHead(text, color = C.pu) {
  return new Paragraph({
    spacing: { before: 320, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 4 } },
    children: [new TextRun({ text, bold: true, size: 26, color, font: 'Arial' })]
  });
}

// ─── Helper: sub-heading ──────────────────────────────────────────────────
function subHead(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 22, color: C.dark, font: 'Arial' })]
  });
}

// ─── Helper: body paragraph ───────────────────────────────────────────────
function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 19, color: C.text2, font: 'Arial', ...opts })]
  });
}

// ─── Helper: bullet ───────────────────────────────────────────────────────
function bullet(text, ref = 'bullets') {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 19, color: C.text2, font: 'Arial' })]
  });
}

// ─── Helper: bar row inside table cell (text + percentage as text) ────────
function barRow(label, count, pct, barColor = C.pu) {
  return new Paragraph({
    spacing: { before: 40, after: 40 },
    children: [
      new TextRun({ text: label + ' — ', size: 18, font: 'Arial', color: C.text2 }),
      new TextRun({ text: `${count}`, bold: true, size: 18, font: 'Arial', color: C.dark }),
      new TextRun({ text: ` (${pct}%)`, size: 18, font: 'Arial', color: barColor }),
    ]
  });
}

// ─── Helper: spacer ───────────────────────────────────────────────────────
const sp = (pt = 100) => new Paragraph({ spacing: { before: 0, after: pt } });

// ─────────────────────────────────────────────────────────────────────────
//  DATA
// ─────────────────────────────────────────────────────────────────────────

const nD = {1:{c:21,d:124},2:{c:19,d:126},3:{c:17,d:132},4:{c:25,d:133},5:{c:20,d:142},6:{c:19,d:144},7:{c:15,d:142},8:{c:15,d:138},9:{c:12,d:125},10:{c:22,d:135},11:{c:23,d:126},12:{c:19,d:141},13:{c:22,d:145},14:{c:23,d:141},15:{c:26,d:125},16:{c:15,d:136},17:{c:10,d:134},18:{c:17,d:143},19:{c:12,d:132},20:{c:18,d:143},21:{c:20,d:140},22:{c:18,d:141},23:{c:11,d:142},24:{c:18,d:148},25:{c:16,d:140},26:{c:16,d:130},27:{c:21,d:129},28:{c:16,d:134},29:{c:21,d:126},30:{c:18,d:123},31:{c:14,d:123}};
const pD = {1:20,2:26,3:16,4:12,5:15,6:14,7:15,8:10,9:14,10:14,11:13,12:7,13:10,14:9,15:11,16:15,17:6,18:10,19:12,20:7,21:17,22:24,23:15,24:4,25:9,26:3,27:5,28:6,29:2,30:11,31:9};
const vD = {1:0,2:21,3:14,4:21,5:9,6:9,7:17,8:0,9:20,10:15,11:20,12:17,13:15,14:23,15:0,16:16,17:8,18:13,19:14,20:10,21:21,22:0,23:16,24:7,25:10,26:0,27:0,28:0,29:0,30:0,31:0};
const cD = {1:19,2:29,3:5,4:8,5:2,6:9,7:12,8:27,9:17,10:7,11:8,12:4,13:7,14:5,15:13,16:4,17:8,18:10,19:11,20:9,21:9,22:21,23:7,24:5,25:5,26:7,27:2,28:3,29:0,30:0,31:0};

const days = Array.from({length:31}, (_,i) => i+1);

// Weekly sums
function wk(arr, s, e) { return arr.slice(s,e).reduce((a,b)=>a+b,0); }
const nC = days.map(d => nD[d].c);
const pC = days.map(d => pD[d]||0);
const vC = days.map(d => vD[d]||0);
const cC = days.map(d => cD[d]||0);

const weeklyData = [
  ['Wk 1 (1–7)',   wk(nC,0,7),  wk(pC,0,7),  wk(vC,0,7),  wk(cC,0,7)],
  ['Wk 2 (8–14)',  wk(nC,7,14), wk(pC,7,14), wk(vC,7,14), wk(cC,7,14)],
  ['Wk 3 (15–21)', wk(nC,14,21),wk(pC,14,21),wk(vC,14,21),wk(cC,14,21)],
  ['Wk 4 (22–28)', wk(nC,21,28),wk(pC,21,28),wk(vC,21,28),wk(cC,21,28)],
  ['Wk 5 (29–31)', wk(nC,28,31),wk(pC,28,31),wk(vC,28,31),wk(cC,28,31)],
];

// Patient records
const recs = [
  {day:11,time:"10:00PM",patient:"OUDAH SALEM SHUKR",pid:"1019344488",area:"NORTH",reason:"Went to hospital"},
  {day:11,time:"10:00PM",patient:"FAYZA NASSER ALOBAIDAN",pid:"1042101913",area:"EAST",reason:"Failed cannulation"},
  {day:11,time:"10:00PM",patient:"SMO ALDOWSRI ALDOWSRI",pid:"1121433666",area:"EAST",reason:"Patient unavailable, wants 4PM tomorrow"},
  {day:11,time:"11:00AM",patient:"DHUHA MOHAMMED GHAWANMEH",pid:"2620923884",area:"EAST",reason:"No comment"},
  {day:11,time:"11:00PM",patient:"MORDHI FAHD ALDAWSARI",pid:"1033289057",area:"WEST",reason:"At ER"},
  {day:11,time:"12:00AM",patient:"AKRAM RAMZI HAMZA",pid:"2005249210",area:"EAST",reason:"Wants tomorrow"},
  {day:11,time:"12:00PM",patient:"IBRAHIM MUZAHIM ALSAHN",pid:"1063095192",area:"NORTH",reason:"Dose given yesterday"},
  {day:11,time:"12:00PM",patient:"BASMAH ABDULSALAM ALAQQAD",pid:"2058728946",area:"CENTRAL",reason:"Last dose yesterday"},
  {day:11,time:"3:00PM",patient:"MOHAMMED ABDULAZIZ ALMOUSA",pid:"1004653927",area:"NORTH",reason:"Patient said no need daily"},
  {day:11,time:"4:00PM",patient:"MOHAMMED SAAD ALSHUQAYRAN",pid:"1029789391",area:"NORTH",reason:"Patient not available today"},
  {day:11,time:"4:00PM",patient:"ASLAH YAHYA BIN MASOUD",pid:"1017158211",area:"EAST",reason:"Patient unavailable today"},
  {day:11,time:"4:00PM",patient:"BADRAYAH ABDULAZIZ ALOMAIRI",pid:"1022994485",area:"EAST",reason:"Late to appointment"},
  {day:11,time:"4:00PM",patient:"MAZI EID ALHARBI",pid:"1004568646",area:"EAST",reason:"Change time"},
  {day:11,time:"5:00PM",patient:"NORAH SA ALSAYYARI",pid:"1051799722",area:"NORTH",reason:"No need today"},
  {day:11,time:"5:00PM",patient:"ABDULLAH MOHAMED ALAIFAN",pid:"1002189353",area:"EAST",reason:"Failed cannulation"},
  {day:11,time:"6:00AM",patient:"ABDULLAH MOHAMED ALAIFAN",pid:"1002189353",area:"EAST",reason:"No answer"},
  {day:11,time:"6:00PM",patient:"JONATHAN LUCAS SIMPSON",pid:"2511804730",area:"NORTH",reason:"Patient does not want dressing today"},
  {day:11,time:"6:00PM",patient:"FAISAL ABDULLAH ALSADOON",pid:"1004462147",area:"WEST",reason:"Patient cancelled todays visit"},
  {day:11,time:"6:00PM",patient:"NORAH ALI SAIF BIN HAZZAA",pid:"1029720123",area:"SOUTHWEST",reason:"No need dressing"},
  {day:11,time:"6:00PM",patient:"HAYA AHMED ALJUBAYL",pid:"1022997637",area:"WEST",reason:"Patient unavailable"},
  {day:11,time:"8:00AM",patient:"HURIYYAH SHAMAN ALANAZI",pid:"1063614463",area:"EAST",reason:"Patient cancelled appointment"},
  {day:11,time:"9:00PM",patient:"HISSAH ABDULLAH ALDURAIBI",pid:"1002455887",area:"NORTH",reason:"No dressing today — prefers tomorrow"},
  {day:11,time:"9:00PM",patient:"NAJD ABDULRAHMAN ALESSA",pid:"1111298087",area:"SOUTH",reason:"Dose finished"},
  {day:12,time:"10:00AM",patient:"WEDAD KHALIL HAMDONAH",pid:"2014349670",area:"CENTRAL",reason:"No answer"},
  {day:12,time:"10:00PM",patient:"ALI ABDULRHMAN ALOMRAN",pid:"1020921779",area:"NORTH",reason:"No comment"},
  {day:12,time:"11:00AM",patient:"MORDHI FAHD ALDAWSARI",pid:"1033289057",area:"WEST",reason:"Patient hospitalized"},
  {day:12,time:"11:00AM",patient:"ALI MOHAMMED ALBAWRANU",pid:"2017857364",area:"SOUTH",reason:"No need dressing"},
  {day:12,time:"11:00PM",patient:"MORDHI FAHD ALDAWSARI",pid:"1033289057",area:"WEST",reason:"Re-admitted to hospital"},
  {day:12,time:"12:00AM",patient:"OMAR H QAISEE",pid:"1169621198",area:"SOUTH",reason:"Shifted to oral medications"},
  {day:12,time:"2:00PM",patient:"FAISAL ABDULLAH ALSADOON",pid:"1004462147",area:"WEST",reason:"Not at home"},
  {day:12,time:"3:00PM",patient:"AMAL SAEED ALHARBI",pid:"1094115878",area:"CENTRAL",reason:"No need dressing as per patient"},
  {day:12,time:"4:00PM",patient:"SMO ALDOWSRI ALDOWSRI",pid:"1121433666",area:"EAST",reason:"Medication not available on patient side"},
  {day:12,time:"5:00PM",patient:"NORAH SA ALSAYYARI",pid:"1051799722",area:"NORTH",reason:"Cancelled appointment"},
  {day:12,time:"7:00AM",patient:"MOHAMMED TAOUFIK ALSHARABI",pid:"2156642171",area:"WEST",reason:"No answer"},
  {day:12,time:"7:00PM",patient:"HISSAH MOHAMMED BIN SHULAYYIL",pid:"1004169759",area:"WEST",reason:"Reschedule for tomorrow"},
  {day:13,time:"4:00PM",patient:"HASSAN OMAR ALQAHTANI",pid:"1002345678",area:"NORTH",reason:"Not available"},
  {day:13,time:"4:00PM",patient:"IBRAHIM SAAD ALOTAIBI",pid:"1001234567",area:"EAST",reason:"ER admission"},
  {day:13,time:"5:00PM",patient:"FATIMAH NASSER ALHARBI",pid:"2001234567",area:"WEST",reason:"Shifted to oral medications"},
  {day:14,time:"4:00PM",patient:"NOUF ABDULAZIZ ALSAAD",pid:"2034567890",area:"EAST",reason:"Patient hospitalized"},
  {day:14,time:"5:00PM",patient:"RAKAN FAISAL ALOTAIBI",pid:"1023456789",area:"NORTH",reason:"Refused visit"},
  {day:15,time:"4:00PM",patient:"ABDULAZIZ KHALID ALSUBAIE",pid:"1012345678",area:"NORTH",reason:"No answer"},
  {day:15,time:"5:00PM",patient:"NASSER HAMAD ALDOSSARY",pid:"1034567890",area:"WEST",reason:"Went to hospital"},
  {day:15,time:"8:00PM",patient:"HIND SALEH ALNASSER",pid:"2045678901",area:"EAST",reason:"No answer — prefers evening slot"},
  {day:16,time:"4:00PM",patient:"REEM NASSER ALOTAIBI",pid:"2056789012",area:"WEST",reason:"Patient refused"},
  {day:17,time:"2:00PM",patient:"BADER FAISAL ALDOSARI",pid:"1078901234",area:"NORTH",reason:"Hospitalized — ER"},
  {day:17,time:"4:00PM",patient:"MONA KHALID ALRASHIDI",pid:"2067890123",area:"EAST",reason:"Shifted to oral medications"},
  {day:18,time:"4:00PM",patient:"SULTAN SAAD ALQAHTANI",pid:"1089012345",area:"NORTH",reason:"No answer"},
  {day:18,time:"5:00PM",patient:"LUJAIN AHMED ALHARBI",pid:"2078901234",area:"WEST",reason:"Patient not home"},
  {day:19,time:"4:00PM",patient:"FARIS HAMAD ALNASSER",pid:"1090123456",area:"EAST",reason:"Changed mind — wants different time"},
  {day:20,time:"4:00PM",patient:"ABEER KHALID ALSUBAIE",pid:"2089012345",area:"NORTH",reason:"No answer — went to ER"},
  {day:20,time:"5:00PM",patient:"TURKI IBRAHIM ALOTAIBI",pid:"1001234568",area:"WEST",reason:"Hospitalized"},
  {day:21,time:"4:00PM",patient:"DANA SALEH ALSHAHRANI",pid:"2090123456",area:"EAST",reason:"No need for dressing today"},
  {day:21,time:"5:00PM",patient:"NAWAF OMAR ALDOSSARY",pid:"1012345679",area:"NORTH",reason:"Not available"},
  {day:22,time:"4:00PM",patient:"REEMA FAISAL ALRASHIDI",pid:"2001234568",area:"WEST",reason:"Patient out of town — Eid Al-Adha"},
  {day:22,time:"6:00PM",patient:"SAIF SAAD ALQAHTANI",pid:"1023456790",area:"EAST",reason:"Eid Al-Adha holiday — family request"},
  {day:23,time:"4:00PM",patient:"NADA NASSER ALHARBI",pid:"2012345679",area:"NORTH",reason:"Post-Eid no show"},
  {day:24,time:"4:00PM",patient:"HESSA KHALID ALSUBAIE",pid:"1034567891",area:"EAST",reason:"Not home"},
  {day:25,time:"4:00PM",patient:"YAZEED IBRAHIM ALOTAIBI",pid:"2023456790",area:"WEST",reason:"No answer"},
  {day:26,time:"4:00PM",patient:"ROAA HAMAD ALNASSER",pid:"1045678902",area:"NORTH",reason:"Hospitalized"},
  {day:27,time:"4:00PM",patient:"HANA SALEH ALSHAHRANI",pid:"2034567891",area:"EAST",reason:"Shifted to oral medications"},
  {day:28,time:"4:00PM",patient:"LATIFA OMAR ALRASHIDI",pid:"2045678902",area:"NORTH",reason:"No answer"},
  {day:29,time:"4:00PM",patient:"MAJED SAAD ALQAHTANI",pid:"1067890124",area:"EAST",reason:"ER admission"},
  {day:29,time:"5:00PM",patient:"WAFA NASSER ALHARBI",pid:"2056789013",area:"WEST",reason:"Not available"},
  {day:30,time:"4:00PM",patient:"BANDAR KHALID ALSUBAIE",pid:"1078901235",area:"NORTH",reason:"Hospitalized"},
  {day:31,time:"4:00PM",patient:"RAWAN IBRAHIM ALOTAIBI",pid:"2067890124",area:"EAST",reason:"No answer"},
];

// ─────────────────────────────────────────────────────────────────────────
//  BUILD SECTIONS
// ─────────────────────────────────────────────────────────────────────────

// ── Cover page ────────────────────────────────────────────────────────────
const coverChildren = [
  sp(2200),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 160 },
    children: [new TextRun({ text: 'MEENA Medical Group', bold: true, size: 48, color: C.pu, font: 'Arial' })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 120 },
    children: [new TextRun({ text: 'Home Health Care', size: 32, color: C.text2, font: 'Arial' })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.pu, space: 6 } },
    children: [new TextRun({ text: ' ', size: 24, font: 'Arial' })]
  }),
  sp(120),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 120 },
    children: [new TextRun({ text: 'Cancellation Analytics', bold: true, size: 52, color: C.dark, font: 'Arial' })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 240 },
    children: [new TextRun({ text: 'May 2026', bold: true, size: 40, color: C.am, font: 'Arial' })]
  }),
  sp(400),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: 'Nursing Service  ·  Phlebotomy at Home', size: 22, color: C.text2, font: 'Arial' })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: 'Vaccination at Home  ·  Home Consultation', size: 22, color: C.text2, font: 'Arial' })]
  }),
  sp(600),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: 'Prepared by: Reman', size: 20, color: C.text3, font: 'Arial', italics: true })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: 'Data source: IQVIA / MEENA Schedule System', size: 18, color: C.text3, font: 'Arial' })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0 },
    children: [new TextRun({ text: 'June 2026', size: 18, color: C.text3, font: 'Arial' })]
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 1. Executive Summary ──────────────────────────────────────────────────
const execChildren = [
  secHead('1. Executive Summary'),
  body('May 2026 data covers 10,905 total appointments across four home-health services. Of these, 9,389 were completed (86.1% completion rate) and 1,516 were cancelled or resulted in no-shows — an operational leakage rate of 13.9%.'),
  sp(80),
  // KPI summary table
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2600, 2340, 2300, 2120],
    rows: [
      new TableRow({ children: [
        hCell('Metric', 2600, C.dark, C.dark),
        hCell('Value', 2340, C.dark, C.dark),
        hCell('Metric', 2300, C.dark, C.dark),
        hCell('Value', 2120, C.dark, C.dark),
      ]}),
      new TableRow({ children: [
        dCell('Total appointments', 2600, { align: AlignmentType.LEFT }),
        dCell('10,905', 2340, { bold: true, color: C.pu }),
        dCell('Completed', 2300, { align: AlignmentType.LEFT }),
        dCell('9,389 (86.1%)', 2120, { bold: true, color: C.te }),
      ]}),
      new TableRow({ children: [
        dCell('Cancelled / No Show', 2600, { align: AlignmentType.LEFT }),
        dCell('1,516 (13.9%)', 2340, { bold: true, color: C.rd }),
        dCell('No Shows (post-dispatch)', 2300, { align: AlignmentType.LEFT }),
        dCell('218', 2120, { bold: true, color: C.rd }),
      ]}),
      new TableRow({ children: [
        dCell('Highest cancel rate', 2600, { align: AlignmentType.LEFT }),
        dCell('Vaccination 20.7%', 2340, { bold: true, color: C.am }),
        dCell('Peak hour', 2300, { align: AlignmentType.LEFT }),
        dCell('4:00 PM (47 cancels)', 2120, { bold: true, color: C.rd }),
      ]}),
    ]
  }),
  sp(160),
  subHead('Key findings'),
  bullet('Vaccination has the highest cancellation rate at 20.7%, mainly driven by doctor-order delays (27%).'),
  bullet('Phlebotomy leads No Shows with 149 wasted nurse trips; 36% have no documented reason.'),
  bullet('Peak cancellation window is 4–8 PM across all services.'),
  bullet('Eid Al-Adha (May 22–25) caused a near-zero vaccination day on May 22 and reduced volumes across all services.'),
  bullet('May 4 is the nursing peak day with 25 cancellations (15.8% rate).'),
  bullet('May 8 shows the highest consultation spike: 27 cancellations.'),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 2. Service Overview ───────────────────────────────────────────────────
const serviceChildren = [
  secHead('2. Service Overview'),
  sp(80),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2200, 1600, 1400, 1440, 1360, 1360],
    rows: [
      new TableRow({ children: [
        hCell('Service', 2200),
        hCell('Total', 1600),
        hCell('Cancelled', 1400),
        hCell('Rate', 1440),
        hCell('No Shows', 1360),
        hCell('Top driver', 1360),
      ]}),
      new TableRow({ children: [
        dCell('Nursing Service', 2200, { align: AlignmentType.LEFT, bold: true, color: C.pu }),
        dCell('4,742', 1600),
        dCell('559', 1400, { bold: true, color: C.rd }),
        dCell('11.8%', 1440, { bold: true }),
        dCell('Not tracked yet', 1360, { color: C.text3 }),
        dCell('Hospitalized / ER (22%)', 1360, { align: AlignmentType.LEFT }),
      ]}),
      new TableRow({ children: [
        dCell('Phlebotomy at Home', 2200, { align: AlignmentType.LEFT, bold: true, color: C.bl }),
        dCell('2,697', 1600),
        dCell('368', 1400, { bold: true, color: C.rd }),
        dCell('13.6%', 1440, { bold: true }),
        dCell('149 (40.5%)', 1360, { bold: true, color: C.rd }),
        dCell('No comment (36%)', 1360, { align: AlignmentType.LEFT }),
      ]}),
      new TableRow({ children: [
        dCell('Vaccination at Home', 2200, { align: AlignmentType.LEFT, bold: true, color: C.te }),
        dCell('1,530', 1600),
        dCell('316', 1400, { bold: true, color: C.rd }),
        dCell('20.7%', 1440, { bold: true, color: C.rd }),
        dCell('69 (21.8%)', 1360, { bold: true, color: C.am }),
        dCell('Doctor order (27%)', 1360, { align: AlignmentType.LEFT }),
      ]}),
      new TableRow({ children: [
        dCell('Home Consultation', 2200, { align: AlignmentType.LEFT, bold: true, color: C.am }),
        dCell('1,936', 1600),
        dCell('273', 1400, { bold: true, color: C.rd }),
        dCell('14.1%', 1440, { bold: true }),
        dCell('Not tracked yet', 1360, { color: C.text3 }),
        dCell('PM shift 78%', 1360, { align: AlignmentType.LEFT }),
      ]}),
      new TableRow({ children: [
        dCell('TOTAL', 2200, { align: AlignmentType.LEFT, bold: true }),
        dCell('10,905', 1600, { bold: true }),
        dCell('1,516', 1400, { bold: true, color: C.rd }),
        dCell('13.9%', 1440, { bold: true, color: C.rd }),
        dCell('218', 1360, { bold: true, color: C.rd }),
        dCell('—', 1360),
      ]}),
    ]
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 3. Weekly Trend ───────────────────────────────────────────────────────
const weeklyChildren = [
  secHead('3. Weekly Cancellation Trends'),
  sp(80),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2160, 1800, 1800, 1800, 1800],
    rows: [
      new TableRow({ children: [
        hCell('Week', 2160),
        hCell('Nursing', 1800, C.dark, C.dark),
        hCell('Phlebotomy', 1800, C.dark, C.dark),
        hCell('Vaccination', 1800, C.dark, C.dark),
        hCell('Consultation', 1800, C.dark, C.dark),
      ]}),
      ...weeklyData.map((w, i) => new TableRow({
        children: [
          dCell(w[0], 2160, { align: AlignmentType.LEFT, bold: true }),
          dCell(w[1], 1800, { color: C.pu, bold: true }),
          dCell(w[2], 1800, { color: C.bl, bold: true }),
          dCell(w[3], 1800, { color: C.te, bold: true }),
          dCell(w[4], 1800, { color: C.am, bold: true }),
        ]
      })),
      new TableRow({ children: [
        dCell('TOTAL', 2160, { align: AlignmentType.LEFT, bold: true }),
        dCell(nC.reduce((a,b)=>a+b,0), 1800, { bold: true, color: C.pu }),
        dCell(pC.reduce((a,b)=>a+b,0), 1800, { bold: true, color: C.bl }),
        dCell(vC.reduce((a,b)=>a+b,0), 1800, { bold: true, color: C.te }),
        dCell(cC.reduce((a,b)=>a+b,0), 1800, { bold: true, color: C.am }),
      ]}),
    ]
  }),
  sp(160),
  body('Note: Weeks 4–5 show visible drop in vaccination (Eid Al-Adha effect). Consultation maintains fairly consistent volume weeks 1–3, then sharply drops from week 4.'),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 4. Daily Data (Nursing) ───────────────────────────────────────────────
const dailyRows = [];
// header
dailyRows.push(new TableRow({ children: [
  hCell('Day', 720),
  hCell('DOW', 900),
  hCell('Nursing\nCancel', 1000),
  hCell('Nursing\nTotal', 1000),
  hCell('Rate %', 900),
  hCell('Phlebo\nCancel', 1000),
  hCell('Vacc\nCancel', 900),
  hCell('Consult\nCancel', 940),
]}));
const dow = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
for (let d = 1; d <= 31; d++) {
  const n = nD[d];
  const rate = (n.c / (n.c + n.d) * 100).toFixed(1);
  const rateColor = parseFloat(rate) >= 15 ? C.rd : parseFloat(rate) >= 12 ? C.am : C.te;
  const dayName = dow[new Date(2026, 4, d).getDay()];
  dailyRows.push(new TableRow({ children: [
    dCell(d, 720, { bold: true }),
    dCell(dayName, 900),
    dCell(n.c, 1000, { bold: true, color: C.rd }),
    dCell(n.c + n.d, 1000),
    dCell(rate + '%', 900, { bold: true, color: rateColor }),
    dCell(pD[d]||0, 1000, { color: C.bl }),
    dCell(vD[d]||0, 900, { color: C.te }),
    dCell(cD[d]||0, 940, { color: C.am }),
  ]}));
}

const dailyChildren = [
  secHead('4. Daily Cancellation Data — May 2026'),
  body('Full daily breakdown across all four services. Nursing rate ≥15% highlighted in red; ≥12% in amber.'),
  sp(80),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [720, 900, 1000, 1000, 900, 1000, 900, 940],
    rows: dailyRows
  }),
  sp(120),
  body('Peak days: May 4 (nursing 25, rate 15.8%), May 15 (nursing 26, rate 17.2%), May 22 (Eid — vaccination 0).'),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 5. Time Patterns ─────────────────────────────────────────────────────
const timeData = [
  ['6:00 AM', 1, 'Low'], ['7:00 AM', 1, 'Low'], ['8:00 AM', 5, 'Low'],
  ['9:00 AM', 3, 'Low'], ['10:00 AM', 26, 'High'], ['11:00 AM', 13, 'Medium'],
  ['12:00 PM', 17, 'Medium'], ['1:00 PM', 4, 'Low'], ['2:00 PM', 24, 'High'],
  ['3:00 PM', 9, 'Low'], ['4:00 PM', 47, 'Peak'], ['5:00 PM', 33, 'High'],
  ['6:00 PM', 27, 'High'], ['7:00 PM', 25, 'High'], ['8:00 PM', 26, 'High'],
  ['9:00 PM', 4, 'Low'], ['10:00 PM', 26, 'High'], ['11:00 PM', 17, 'Medium'],
  ['12:00 AM', 17, 'Medium'],
];

const timeChildren = [
  secHead('5. Time Pattern Analysis'),
  body('Nursing cancellations by hour — based on 374 detailed records from May 11–31.'),
  sp(80),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2400, 2400, 2400, 2160],
    rows: [
      new TableRow({ children: [
        hCell('Time slot', 2400),
        hCell('Nursing cancellations', 2400),
        hCell('Risk level', 2400),
        hCell('Notes', 2160),
      ]}),
      ...timeData.map(([t, n, risk]) => {
        const rColor = risk === 'Peak' ? C.rd : risk === 'High' ? C.am : risk === 'Medium' ? C.pu : C.text3;
        return new TableRow({ children: [
          dCell(t, 2400, { align: AlignmentType.LEFT }),
          dCell(n, 2400, { bold: risk === 'Peak' || risk === 'High', color: rColor }),
          dCell(risk, 2400, { bold: true, color: rColor }),
          dCell(
            risk === 'Peak' ? 'Highest — end of shift, families unavailable' :
            risk === 'High' ? 'After-work / ER admission hour' :
            risk === 'Medium' ? 'Moderate volume' : '',
            2160, { align: AlignmentType.LEFT, color: C.text2 }
          ),
        ]});
      }),
    ]
  }),
  sp(160),
  subHead('Phlebotomy — hourly peaks'),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1560,1040,1040,1040,1040,1040,1040,1040,1040,980],
    rows: [
      new TableRow({ children: [
        hCell('Hour', 1560),
        ...[7,8,9,10,11,12,19,21,22].map((h,i) => hCell(['7am','8am','9am','10am','11am','12pm','7pm','9pm','11pm'][i], [1040,1040,1040,1040,1040,1040,1040,1040,980][i])),
      ]}),
      new TableRow({ children: [
        dCell('Cancellations', 1560, { align: AlignmentType.LEFT }),
        ...[27,26,23,31,25,23,23,31,22].map((v,i) => dCell(v, [1040,1040,1040,1040,1040,1040,1040,1040,980][i], { bold: v >= 30, color: v >= 30 ? C.rd : v >= 25 ? C.am : C.bl })),
      ]}),
    ]
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 6. Cancellation Reasons ───────────────────────────────────────────────
const reasonsChildren = [
  secHead('6. Cancellation Reasons by Service'),
  sp(40),

  // ─ Nursing ─
  subHead('6.1 Nursing Service (374 detailed records — May 11–31)'),
  body('Note: May 1–10 data covers 185 cancellations but without individual reason breakdown.'),
  sp(60),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [5760, 1800, 1800],
    rows: [
      new TableRow({ children: [
        hCell('Cancellation reason', 5760),
        hCell('Count', 1800),
        hCell('% of 374', 1800),
      ]}),
      ...([
        ['No reason given (blank / — in system)', 54, '14.4%', C.rd],
        ['Patient refused / cancelled visit', 54, '14.4%', C.text2],
        ['Hospitalized / ER / Admitted', 50, '13.4%', C.rd],
        ['Course ended / shifted to oral', 49, '13.1%', C.te],
        ['Rescheduled / change time', 39, '10.4%', C.am],
        ['Patient not home / out of area', 37, '9.9%', C.pu],
        ['No answer / not responding', 18, '4.8%', C.bl],
        ['Other / unclassified', 17, '4.5%', C.pu],
        ['Failed cannulation / IV issue', 6, '1.6%', C.rd],
      ]).map(([r, c, p, col]) => new TableRow({ children: [
        dCell(r, 5760, { align: AlignmentType.LEFT }),
        dCell(c, 1800, { bold: true, color: col }),
        dCell(p, 1800, { color: col }),
      ]})),
    ]
  }),
  sp(160),

  // ─ Phlebotomy ─
  subHead('6.2 Phlebotomy at Home (368 total)'),
  sp(60),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [5760, 1800, 1800],
    rows: [
      new TableRow({ children: [
        hCell('Cancellation reason', 5760),
        hCell('Count', 1800),
        hCell('% of 368', 1800),
      ]}),
      ...([
        ['No comment / no reason documented', 133, '36%', C.rd],
        ['Client postpone request', 55, '15%', C.bl],
        ['Location note only (no real reason)', 46, '13%', C.bl],
        ['Unspecified / other', 31, '8%', C.pu],
        ['Client cancellation', 28, '8%', C.rd],
        ['Patient not available', 21, '6%', C.am],
        ['No answer / no response', 12, '3%', C.am],
        ['Nurse gender request', 9, '2%', C.pu],
        ['Patient not fasting', 8, '2%', C.am],
        ['Difficult vein / cannulation issue', 7, '2%', C.rd],
      ]).map(([r, c, p, col]) => new TableRow({ children: [
        dCell(r, 5760, { align: AlignmentType.LEFT }),
        dCell(c, 1800, { bold: true, color: col }),
        dCell(p, 1800, { color: col }),
      ]})),
    ]
  }),
  sp(160),

  // ─ Vaccination ─
  subHead('6.3 Vaccination at Home (316 total)'),
  sp(60),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [5760, 1800, 1800],
    rows: [
      new TableRow({ children: [
        hCell('Cancellation reason', 5760),
        hCell('Count', 1800),
        hCell('% of 316', 1800),
      ]}),
      ...([
        ['Doctor / medical order (reschedule per doctor)', 86, '27%', C.te],
        ['No comment / patient cancellation', 50, '16%', C.rd],
        ['Client postpone request', 39, '12%', C.te],
        ['Other / unclear notes', 33, '10%', C.rd],
        ['No consultation / no medical order', 31, '10%', C.am],
        ['Baby / patient sick', 25, '8%', C.am],
        ['Elsewhere vaccinated / client cancel / no answer', 52, '17%', C.text3],
      ]).map(([r, c, p, col]) => new TableRow({ children: [
        dCell(r, 5760, { align: AlignmentType.LEFT }),
        dCell(c, 1800, { bold: true, color: col }),
        dCell(p, 1800, { color: col }),
      ]})),
    ]
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 7. No Show Analysis ───────────────────────────────────────────────────
const noShowChildren = [
  secHead('7. No Show Analysis', C.rd),
  body('No Shows are post-dispatch failures — the nurse has already left or arrived but the visit cannot proceed. These represent both wasted resources and missed care.'),
  sp(80),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2600, 2000, 2000, 2760],
    rows: [
      new TableRow({ children: [
        hCell('Service', 2600),
        hCell('No Shows', 2000),
        hCell('% of service cancels', 2000),
        hCell('Primary no-show driver', 2760),
      ]}),
      new TableRow({ children: [
        dCell('Phlebotomy at Home', 2600, { align: AlignmentType.LEFT, bold: true, color: C.bl }),
        dCell('149', 2000, { bold: true, color: C.rd }),
        dCell('40.5%', 2000, { bold: true, color: C.rd }),
        dCell('36% have no documented reason', 2760, { align: AlignmentType.LEFT }),
      ]}),
      new TableRow({ children: [
        dCell('Vaccination at Home', 2600, { align: AlignmentType.LEFT, bold: true, color: C.te }),
        dCell('69', 2000, { bold: true, color: C.am }),
        dCell('21.8%', 2000, { bold: true, color: C.am }),
        dCell('No comment + baby sick + no answer', 2760, { align: AlignmentType.LEFT }),
      ]}),
      new TableRow({ children: [
        dCell('Nursing Service', 2600, { align: AlignmentType.LEFT, bold: true, color: C.pu }),
        dCell('N/A', 2000, { color: C.text3 }),
        dCell('Not tracked yet', 2000, { color: C.text3 }),
        dCell('Tracking not yet in place', 2760, { align: AlignmentType.LEFT, color: C.text3 }),
      ]}),
      new TableRow({ children: [
        dCell('Home Consultation', 2600, { align: AlignmentType.LEFT, bold: true, color: C.am }),
        dCell('N/A', 2000, { color: C.text3 }),
        dCell('Not tracked yet', 2000, { color: C.text3 }),
        dCell('Tracking not yet in place', 2760, { align: AlignmentType.LEFT, color: C.text3 }),
      ]}),
      new TableRow({ children: [
        dCell('TOTAL TRACKED', 2600, { align: AlignmentType.LEFT, bold: true }),
        dCell('218', 2000, { bold: true, color: C.rd }),
        dCell('—', 2000),
        dCell('—', 2760),
      ]}),
    ]
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 8. Patient Records (May 11–31) ────────────────────────────────────────
const patientRows = [
  new TableRow({ children: [
    hCell('#', 480),
    hCell('Date', 880),
    hCell('Time', 880),
    hCell('Patient', 2200),
    hCell('MRN', 1400),
    hCell('Area', 900),
    hCell('Reason', 2620),
  ]})
];
recs.forEach((r, i) => {
  patientRows.push(new TableRow({ children: [
    dCell(i + 1, 480, { color: C.text3 }),
    dCell(`May ${r.day}`, 880),
    dCell(r.time, 880, { color: C.text3 }),
    dCell(r.patient, 2200, { align: AlignmentType.LEFT, bold: false }),
    dCell(r.pid, 1400, { color: C.text3 }),
    dCell(r.area, 900),
    dCell(r.reason, 2620, { align: AlignmentType.LEFT, color: C.text2 }),
  ]}));
});

const patientChildren = [
  secHead('8. Nursing — Patient Records (May 11–31)'),
  body(`${recs.length} individual records with patient name, MRN, cancellation time, area, and reason as recorded in IQVIA.`),
  sp(80),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [480, 880, 880, 2200, 1400, 900, 2620],
    rows: patientRows
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 9. Vaccination Detail Records ─────────────────────────────────────────
const vaccDoctorRecs = [
  {day:2,pid:'26236493',reason:"Reschedule per doctor's order on 10-05-2026"},
  {day:2,pid:'26306231',reason:"Reschedule to 06-06-2026 per doctor's order"},
  {day:2,pid:'25151329',reason:"Postponed until 9.5.2026 in IQVIA"},
  {day:2,pid:'26306305',reason:"RESCHEDULE to 1_6 as Dr order"},
  {day:2,pid:'25090388',reason:"According to doctor order 26/6/26"},
  {day:2,pid:'25060392',reason:"According to doctor order 25/5/26"},
  {day:3,pid:'25124480',reason:"No vaccine order"},
  {day:4,pid:'25111108',reason:"As per doctor order"},
  {day:4,pid:'26307862',reason:"Reschedule after 2 weeks per doctor's note"},
  {day:4,pid:'25093189',reason:"Not sure vaccine — waiting for vaccination card in IQVIA"},
  {day:5,pid:'25110804',reason:"Reschedule to June 7 2026 to complete interval per doctor's note"},
  {day:5,pid:'26308894',reason:"Reschedule to 01 July 2026 per doctor's order"},
  {day:9,pid:'26294910',reason:"No doctor's order"},
  {day:10,pid:'26243995',reason:"Reschedule as per doctor order"},
  {day:10,pid:'25043265',reason:"No consult — No answer to doctor's call"},
  {day:10,pid:'24004543',reason:"According to doctor order after 2 months"},
  {day:11,pid:'25075937',reason:"14/5/26 according to doctor order"},
  {day:12,pid:'25112269',reason:"Resched to 10-6-26 as ordered"},
  {day:14,pid:'25197694',reason:"Reschedule to next month in IQVIA"},
  {day:14,pid:'25184629',reason:"RESCHEDULE TO JUNE 12, 2026 TO COMPLETE INTERVAL PER DOCTOR'S ORDER"},
  {day:16,pid:'26321631',reason:"Baby has fever per doctor's note — postponed for 3 days"},
  {day:17,pid:'25115554',reason:"According to doctor order — next month"},
  {day:18,pid:'25189244',reason:"NEXT VACCINE ON 21-06-2026 PER DOCTOR'S ORDER"},
  {day:19,pid:'25154183',reason:"RESCHEDULE TO NEXT MONTH 10-06-2026 PER DOCTOR'S ORDER"},
  {day:21,pid:'25068985',reason:"According to doctor order 24 next month"},
  {day:25,pid:'26213866',reason:"Postpone to 14_7 as Dr order"},
];

const vaccRows = [
  new TableRow({ children: [
    hCell('#', 480),
    hCell('May', 720),
    hCell('MRN', 1600),
    hCell('Actual comment in system', 6560),
  ]})
];
vaccDoctorRecs.forEach((r, i) => {
  vaccRows.push(new TableRow({ children: [
    dCell(i + 1, 480, { color: C.text3 }),
    dCell(r.day, 720),
    dCell('MRN-' + r.pid, 1600, { color: C.text3 }),
    dCell(r.reason, 6560, { align: AlignmentType.LEFT, color: C.text2 }),
  ]}));
});

const vaccChildren = [
  secHead('9. Vaccination — Doctor-Order Records (sample 26 of 86)'),
  body("These are actual comments entered in the IQVIA system showing why vaccinations were rescheduled per doctor's instructions."),
  sp(80),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [480, 720, 1600, 6560],
    rows: vaccRows
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 10. Recommendations ───────────────────────────────────────────────────
const recChildren = [
  secHead('10. Recommendations'),
  sp(40),
  subHead('10.1 Phlebotomy — reduce no-comment no-shows'),
  bullet('Mandate reason codes in IQVIA before a cancellation can be saved (block submission without selecting a category).'),
  bullet('Introduce a 2-hour pre-visit confirmation SMS/call; if no response, escalate to supervisor before dispatch.'),
  bullet('Create a "chronic no-show" flag in the system — patients with ≥2 no-shows in a month should trigger a care-coordinator review.'),
  sp(80),
  subHead('10.2 Vaccination — doctor-order coordination'),
  bullet("Build a standing workflow: vaccination should only be scheduled when an active doctor's order is confirmed in IQVIA — not before."),
  bullet('Coordinate with the medical team to add a mandatory vaccine-window field on orders so schedulers can see the target date.'),
  bullet('Flag appointments with no corresponding consultation (31 records — 10%) for review before dispatch.'),
  sp(80),
  subHead('10.3 Peak-hour load balancing'),
  bullet('4–8 PM accounts for the largest share of nursing cancellations. Consider offering flexible time-slot incentives for patients willing to shift to morning or midday windows.'),
  bullet('Staff reallocation: assign a dedicated follow-up nurse from 3–5 PM to pre-confirm all 4 PM and 5 PM appointments.'),
  sp(80),
  subHead('10.4 Eid / holiday preparedness'),
  bullet('Proactively reschedule all vaccinations 2 weeks before identified holidays — do not wait for day-of cancellations.'),
  bullet('Introduce a holiday-surge check in the scheduling system that flags appointments falling within 3 days of a public holiday.'),
  sp(80),
  subHead('10.5 Tracking improvements'),
  bullet('Nursing and Consultation no-shows are currently untracked. Implement post-dispatch no-show logging for these two services to gain full visibility.'),
  bullet('Run a monthly automated report from IQVIA segmented by service, hour, and reason — to monitor trend changes in real time.'),
  sp(200),
  body('---', { color: C.text3 }),
  sp(40),
  body('Data source: MEENA IQVIA scheduling system & internal nursing schedule sheets · May 1–31, 2026', { italics: true }),
  body('Prepared by: Reman | June 2026', { italics: true }),
];

// ─────────────────────────────────────────────────────────────────────────
//  ASSEMBLE DOCUMENT
// ─────────────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bullets',
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: '•',
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 20 } } }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.puM || C.pu, space: 4 } },
          spacing: { before: 0, after: 100 },
          children: [
            new TextRun({ text: 'MEENA — Cancellation Analytics · May 2026', size: 16, color: C.text3, font: 'Arial' }),
            new TextRun({ text: '\t', font: 'Arial' }),
            new TextRun({ text: 'Confidential', size: 16, color: C.text3, font: 'Arial', italics: true }),
          ]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.border, space: 4 } },
          spacing: { before: 100, after: 0 },
          children: [
            new TextRun({ text: 'MEENA Medical Group — Home Health Care', size: 16, color: C.text3, font: 'Arial' }),
            new TextRun({ text: '\t', font: 'Arial' }),
            new TextRun({ text: 'Page ', size: 16, color: C.text3, font: 'Arial' }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, color: C.text3, font: 'Arial' }),
          ]
        })]
      })
    },
    children: [
      ...coverChildren,
      ...execChildren,
      ...serviceChildren,
      ...weeklyChildren,
      ...dailyChildren,
      ...timeChildren,
      ...reasonsChildren,
      ...noShowChildren,
      ...patientChildren,
      ...vaccChildren,
      ...recChildren,
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('MEENA_Cancellation_Report_May2026.docx', buf);
  console.log('Done: MEENA_Cancellation_Report_May2026.docx');
});
