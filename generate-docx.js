const docx = require('docx');
const fs = require('fs');

const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, TabStopPosition, TabStopType,
  ShadingType, Table, TableRow, TableCell, WidthType,
  TableBorders, VerticalAlign, ImageRun
} = docx;

const GOLD = '9E7B4F';
const TEAL = '2E8B82';
const DARK = '1A1A2E';
const MED = '444444';
const LIGHT = '777777';
const WHITE = 'FFFFFF';
const BG_LIGHT = 'F5F3EF';
const BG_ACCENT = 'F0EBE3';

function heading(text) {
  return new Paragraph({
    spacing: { before: 360, after: 160 },
    border: { bottom: { color: GOLD, size: 6, style: BorderStyle.SINGLE, space: 6 } },
    children: [
      new TextRun({ text: text.toUpperCase(), font: 'Trebuchet MS', size: 22, bold: true, color: DARK, characterSpacing: 120 })
    ]
  });
}

function bulletItem(text, opts = {}) {
  const runs = [];
  const parts = text.split(/(\*\*.*?\*\*)/g);
  parts.forEach(p => {
    if (p.startsWith('**') && p.endsWith('**')) {
      runs.push(new TextRun({ text: p.slice(2, -2), font: 'Calibri', size: 21, bold: true, color: DARK }));
    } else {
      runs.push(new TextRun({ text: p, font: 'Calibri', size: 21, color: MED }));
    }
  });
  return new Paragraph({
    spacing: { after: 60 },
    bullet: { level: 0 },
    children: runs,
    ...opts
  });
}

function bodyText(text, opts = {}) {
  const runs = [];
  const parts = text.split(/(\*\*.*?\*\*)/g);
  parts.forEach(p => {
    if (p.startsWith('**') && p.endsWith('**')) {
      runs.push(new TextRun({ text: p.slice(2, -2), font: 'Calibri', size: 21, bold: true, color: DARK }));
    } else {
      runs.push(new TextRun({ text: p, font: 'Calibri', size: 21, color: MED }));
    }
  });
  return new Paragraph({ spacing: { after: 100 }, children: runs, ...opts });
}

function spacer(h = 120) {
  return new Paragraph({ spacing: { after: h }, children: [] });
}

function expBlock(date, role, org, bullets) {
  const items = [
    new Paragraph({
      spacing: { after: 20 },
      children: [
        new TextRun({ text: date, font: 'Calibri', size: 19, color: TEAL, bold: true })
      ]
    }),
    new Paragraph({
      spacing: { after: 30 },
      children: [
        new TextRun({ text: role, font: 'Trebuchet MS', size: 23, bold: true, color: DARK })
      ]
    }),
    new Paragraph({
      spacing: { after: 80 },
      children: [
        new TextRun({ text: org, font: 'Calibri', size: 21, italics: true, color: LIGHT })
      ]
    }),
    ...bullets.map(b => bulletItem(b))
  ];
  return items;
}

function skillCell(text) {
  return new TableCell({
    width: { size: 33, type: WidthType.PERCENTAGE },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    shading: { type: ShadingType.SOLID, color: BG_LIGHT },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
      left: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
      right: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' }
    },
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text, font: 'Calibri', size: 20, color: DARK })]
      })
    ]
  });
}

function compCard(title, items) {
  return new TableCell({
    width: { size: 33, type: WidthType.PERCENTAGE },
    margins: { top: 120, bottom: 120, left: 140, right: 140 },
    shading: { type: ShadingType.SOLID, color: BG_LIGHT },
    verticalAlign: VerticalAlign.TOP,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: GOLD },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
      left: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
      right: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' }
    },
    children: [
      new Paragraph({
        spacing: { after: 100 },
        children: [new TextRun({ text: title, font: 'Trebuchet MS', size: 20, bold: true, color: DARK })]
      }),
      ...items.map(item => new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: '  ▸  ', font: 'Calibri', size: 18, color: GOLD }),
          new TextRun({ text: item, font: 'Calibri', size: 19, color: MED })
        ]
      }))
    ]
  });
}

function certPill(text) {
  return new Paragraph({
    spacing: { after: 60 },
    children: [
      new TextRun({ text: '  ●  ', font: 'Calibri', size: 18, color: TEAL }),
      new TextRun({ text, font: 'Calibri', size: 21, color: MED })
    ]
  });
}

const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: 'Calibri', size: 21, color: MED }
      },
      listParagraph: {
        run: { font: 'Calibri', size: 21 }
      }
    }
  },
  numbering: {
    config: [{
      reference: 'bullet-list',
      levels: [{
        level: 0,
        format: docx.LevelFormat.BULLET,
        text: '▸',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 200 } } }
      }]
    }]
  },
  sections: [{
    properties: {
      page: {
        margin: { top: 800, bottom: 700, left: 900, right: 900 }
      }
    },
    children: [
      // ═══ HEADER ═══
      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: 'REMAN MOHAMMED BADAWOOD', font: 'Trebuchet MS', size: 36, bold: true, color: DARK, characterSpacing: 80 })
        ]
      }),
      new Paragraph({
        spacing: { after: 80 },
        border: { bottom: { color: GOLD, size: 8, style: BorderStyle.SINGLE, space: 10 } },
        children: [
          new TextRun({ text: 'Healthcare Digital Transformation & Audit Analytics Leader', font: 'Calibri', size: 24, color: GOLD, bold: true })
        ]
      }),
      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: 'AI-Driven Analytics  |  Risk Assessment  |  Compliance Monitoring  |  Process Improvement', font: 'Calibri', size: 19, italics: true, color: LIGHT })
        ]
      }),
      new Paragraph({
        spacing: { after: 30 },
        children: [
          new TextRun({ text: 'Riyadh, Saudi Arabia   ·   reman.m.b@hotmail.com   ·   SCFHS Licensed Registered Nurse', font: 'Calibri', size: 19, color: LIGHT })
        ]
      }),

      spacer(160),

      // ═══ SUMMARY ═══
      heading('Professional Summary'),
      bodyText('Healthcare operations and quality leader with **10+ years** of progressive experience spanning critical care, home healthcare operations, leadership, compliance monitoring, audit analytics, and **digital transformation**.'),
      bodyText('Demonstrated success in leading large-scale healthcare operations, overseeing **112+ clinical staff** across multiple service lines, developing audit frameworks, monitoring compliance, identifying operational risks, and building **AI-assisted analytics tools** to support strategic decision-making.'),
      bodyText('Skilled in bridging healthcare domain expertise with business operations, data analytics, and digital innovation to improve workflows, strengthen internal controls, and drive measurable organizational performance.'),

      spacer(100),

      // ═══ COMPETENCIES ═══
      heading('Core Competencies'),
      spacer(60),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              compCard('Operations, Audit & Risk', [
                'Operational Audit',
                'Compliance Monitoring',
                'Risk Assessment',
                'Root Cause Analysis',
                'Corrective Action Planning',
                'Incident Investigation',
                'Process Improvement'
              ]),
              compCard('Data Analytics & Reporting', [
                'KPI Dashboards',
                'Data Cleaning & Trend Analysis',
                'Performance Reporting',
                'Operational Analytics',
                'Risk Indicators (KRIs)',
                'Data Storytelling'
              ]),
              compCard('Digital Transformation', [
                'AI-assisted Automation',
                'Dashboard Development',
                'Product Requirement Analysis',
                'Workflow Optimization',
                'Internal Digital Tools'
              ])
            ]
          })
        ]
      }),

      spacer(100),

      // ═══ EXPERIENCE ═══
      heading('Professional Experience'),

      ...expBlock('2025 — Present', 'Head Nurse — Operations, Compliance & Analytics Lead', 'Meena Health Group  ·  Riyadh', [
        'Oversee operational performance and compliance monitoring for **112 nurses** across **7 specialized services**',
        'Designed KPI dashboards and analytics reports for performance tracking and decision support',
        'Built operational audit systems to monitor staff readiness, documentation quality, compliance, and performance',
        'Identified workflow bottlenecks and operational risks, implementing corrective actions to improve service quality',
        'Analyzed complaints, incidents, and service trends to identify root causes and support risk mitigation',
        'Led accreditation readiness and policy implementation across multiple service lines',
        'Built internal AI-assisted digital solutions including operational dashboards, LMS, and nurse management systems'
      ]),
      spacer(100),

      ...expBlock('2023 — 2025', 'Acting Head Nurse — Home Healthcare Operations', 'King Khalid University Hospital  ·  Riyadh', [
        'Supervised home healthcare operations and staff performance',
        'Monitored compliance with documentation and quality standards',
        'Supported workflow optimization and service quality improvement initiatives'
      ]),
      spacer(100),

      ...expBlock('2020 — 2023', 'Clinical Operations / Charge Nurse — Home Healthcare', 'King Fahad Medical City  ·  Riyadh', [
        'Delivered specialized home healthcare services',
        'Ensured clinical documentation compliance and service quality',
        'Participated in operational monitoring and quality initiatives'
      ]),
      spacer(100),

      ...expBlock('2016 — 2020', 'Critical Care Nurse (ICU / RRT / PICU)', 'Security Forces Hospital  ·  Riyadh', [
        'Managed critically ill patients in high-risk clinical environments',
        'Developed strong risk assessment, escalation, and crisis management skills',
        'Strengthened patient safety and incident response competencies'
      ]),

      spacer(100),

      // ═══ PROJECTS ═══
      heading('Audit & Analytics Projects'),

      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: 'Nurse Audit & Compliance System', font: 'Trebuchet MS', size: 22, bold: true, color: DARK })
        ]
      }),
      bodyText('Developed a digital audit framework for monitoring compliance, documentation quality, staff readiness, and corrective action tracking.'),
      spacer(60),

      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: 'Operational KPI Dashboard', font: 'Trebuchet MS', size: 22, bold: true, color: DARK })
        ]
      }),
      bodyText('Built multi-service dashboards tracking cancellation trends, no-show trends, workload distribution, performance metrics, and operational risk indicators.'),
      spacer(60),

      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: 'HomeCare Pro', font: 'Trebuchet MS', size: 22, bold: true, color: DARK })
        ]
      }),
      bodyText('Developed an internal nurse management platform featuring readiness checklists, compliance tracking, performance analytics, and leadership reporting.'),
      spacer(60),

      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: 'Learning Hub LMS', font: 'Trebuchet MS', size: 22, bold: true, color: DARK })
        ]
      }),
      bodyText('Built a learning management platform with training compliance tracking, analytics, certificates, and engagement monitoring.'),

      spacer(100),

      // ═══ SKILLS ═══
      heading('Technical Skills'),
      spacer(60),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({ children: [skillCell('Excel — Advanced'), skillCell('Chart.js — Advanced'), skillCell('HTML / CSS — Advanced')] }),
          new TableRow({ children: [skillCell('Supabase — Advanced'), skillCell('MS Office 365 — Advanced'), skillCell('Google Sheets API — Advanced')] }),
          new TableRow({ children: [skillCell('Python — Basic'), skillCell('SQL — Learning'), skillCell('Canva Pro — Advanced')] }),
        ]
      }),

      spacer(100),

      // ═══ EDUCATION ═══
      heading('Education'),

      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: 'Bachelor of Science in Nursing (BSN)', font: 'Trebuchet MS', size: 22, bold: true, color: DARK })
        ]
      }),
      new Paragraph({
        spacing: { after: 20 },
        children: [
          new TextRun({ text: 'Princess Nourah Bint Abdulrahman University', font: 'Calibri', size: 21, color: TEAL })
        ]
      }),
      bodyText('GPA: **4.4 / 5**'),
      spacer(80),

      new Paragraph({
        spacing: { after: 40 },
        children: [
          new TextRun({ text: 'Higher Diploma in Artificial Intelligence', font: 'Trebuchet MS', size: 22, bold: true, color: DARK })
        ]
      }),
      new Paragraph({
        spacing: { after: 20 },
        children: [
          new TextRun({ text: '2025 — 2026  (In Progress)', font: 'Calibri', size: 21, color: TEAL })
        ]
      }),

      spacer(100),

      // ═══ CERTIFICATIONS ═══
      heading('Certifications'),
      certPill('IIWCC — University of Toronto Wound Care'),
      certPill('Healthcare Risk Management'),
      certPill('Change Management'),
      certPill('PMP Preparation Course'),
      certPill('Nursing Leadership Diploma'),
      certPill('BLS  /  ACLS  /  PALS'),

      spacer(100),

      // ═══ LANGUAGES ═══
      heading('Languages'),
      new Table({
        width: { size: 60, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                shading: { type: ShadingType.SOLID, color: BG_LIGHT },
                borders: {
                  top: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
                  bottom: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
                  left: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
                  right: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' }
                },
                children: [
                  new Paragraph({
                    children: [new TextRun({ text: 'Arabic', font: 'Trebuchet MS', size: 22, bold: true, color: DARK })]
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: 'Native', font: 'Calibri', size: 19, color: LIGHT })]
                  })
                ]
              }),
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                margins: { top: 80, bottom: 80, left: 120, right: 120 },
                shading: { type: ShadingType.SOLID, color: BG_LIGHT },
                borders: {
                  top: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
                  bottom: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
                  left: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
                  right: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' }
                },
                children: [
                  new Paragraph({
                    children: [new TextRun({ text: 'English', font: 'Trebuchet MS', size: 22, bold: true, color: DARK })]
                  }),
                  new Paragraph({
                    children: [new TextRun({ text: 'Professional Working Proficiency', font: 'Calibri', size: 19, color: LIGHT })]
                  })
                ]
              })
            ]
          })
        ]
      }),

      spacer(200),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('Reman_Badawood_CV.docx', buffer);
  console.log('Done — Reman_Badawood_CV.docx created');
});
