/**
 * NSS Report Generator
 * Reads nss_retrieved.json → writes NSS_Report.docx
 *
 * Usage:
 *   node generate_report.js [path/to/nss_retrieved.json]
 *
 * Output: NSS_Report.docx
 */

const fs   = require("fs");
const path = require("path");

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumberElement, PageBreak,
  LevelFormat, TabStopType, TabStopPosition,
} = require("docx");

// ── INPUT ─────────────────────────────────────────────────────────────────
const jsonPath = process.argv[2] || "nss_retrieved.json";

// ── SAMPLE DATA (used when JSON not found) ───────────────────────────────
// Replace with real JSON output from rag_retrieve.py
const SAMPLE_DATA = {
  "The teaching on my course": [
    { course_code: "CS 135",  review_text: "Lectures were well-structured and engaging throughout the term.", nss_labels: "Teaching", similarity: 0.821 },
    { course_code: "MATH 137",review_text: "Professor explained concepts clearly with real-world examples.",  nss_labels: "Teaching", similarity: 0.794 },
    { course_code: "ECON 101",review_text: "Course content was relevant but delivery could be improved.",    nss_labels: "Teaching", similarity: 0.768 },
  ],
  "Learning opportunities": [
    { course_code: "STAT 230",review_text: "Limited hands-on projects; more applied work would help learning.", nss_labels: "Learning", similarity: 0.810 },
    { course_code: "CS 240",  review_text: "Assignments challenged me to think independently.",                 nss_labels: "Learning", similarity: 0.788 },
    { course_code: "ECON 201",review_text: "Group projects fostered collaborative skills.",                    nss_labels: "Learning", similarity: 0.762 },
  ],
  "Assessment and feedback": [
    { course_code: "ECON 101",review_text: "ez 95 w testbank — exam was straightforward if you prepped.", nss_labels: "Assessment", similarity: 0.800 },
    { course_code: "CS 135",  review_text: "Feedback on assignments was timely and constructive.",        nss_labels: "Assessment", similarity: 0.779 },
    { course_code: "MATH 237",review_text: "Midterm weighting felt unbalanced compared to final.",        nss_labels: "Assessment", similarity: 0.751 },
  ],
  "Academic support": [
    { course_code: "STAT 230",review_text: "Office hours were rarely held; TAs were more helpful.",        nss_labels: "Support",    similarity: 0.793 },
    { course_code: "CS 241",  review_text: "Piazza forum was responsive and a great learning resource.",   nss_labels: "Support",    similarity: 0.774 },
    { course_code: "ECON 102",review_text: "Academic advisor helped navigate course selection effectively.", nss_labels: "Support",   similarity: 0.750 },
  ],
  "Organisation and management": [
    { course_code: "PD 1",    review_text: "fk pd — poorly managed with unclear expectations.",           nss_labels: "Organisation", similarity: 0.835 },
    { course_code: "CS 240",  review_text: "Schedule conflicts between labs and lectures were frustrating.", nss_labels: "Organisation", similarity: 0.799 },
    { course_code: "ECON 201",review_text: "Course outline was detailed and followed consistently.",        nss_labels: "Organisation", similarity: 0.772 },
  ],
  "Learning resources": [
    { course_code: "MATH 237",review_text: "Textbook was expensive and barely used in lectures.",          nss_labels: "Resources",  similarity: 0.804 },
    { course_code: "CS 135",  review_text: "Course notes were comprehensive and freely available online.", nss_labels: "Resources",  similarity: 0.781 },
    { course_code: "STAT 230",review_text: "Lab computers were outdated but software licences were fine.", nss_labels: "Resources",  similarity: 0.755 },
  ],
  "Student voice": [
    { course_code: "ECON 101",review_text: "Mid-term surveys were collected but no changes were made.",   nss_labels: "Voice",      similarity: 0.788 },
    { course_code: "CS 240",  review_text: "Student feedback was acknowledged in the next lecture.",      nss_labels: "Voice",      similarity: 0.764 },
    { course_code: "PD 1",    review_text: "No mechanism to raise concerns with the programme team.",     nss_labels: "Voice",      similarity: 0.741 },
  ],
  "Student union": [
    { course_code: "ENGL 108W",review_text: "Union events were well advertised but poorly attended.",     nss_labels: "Union",      similarity: 0.762 },
    { course_code: "ECON 201", review_text: "Union rep visited twice; impact on curriculum was unclear.", nss_labels: "Union",      similarity: 0.744 },
    { course_code: "MATH 137", review_text: "Unaware of union activities during the academic term.",      nss_labels: "Union",      similarity: 0.721 },
  ],
  "Mental wellbeing": [
    { course_code: "PD 1",    review_text: "Workload spikes before deadlines caused significant stress.", nss_labels: "Wellbeing",  similarity: 0.819 },
    { course_code: "CS 135",  review_text: "No wellness check-ins from instructor during heavy weeks.",  nss_labels: "Wellbeing",  similarity: 0.796 },
    { course_code: "STAT 230",review_text: "Support services existed but were difficult to access.",     nss_labels: "Wellbeing",  similarity: 0.770 },
  ],
  "Freedom of expression": [
    { course_code: "ECON 101",review_text: "Discussion sections felt open; dissenting views were respected.", nss_labels: "Expression", similarity: 0.776 },
    { course_code: "ENGL 108W",review_text: "Creative writing assignments allowed genuine self-expression.",  nss_labels: "Expression", similarity: 0.758 },
    { course_code: "CS 240",   review_text: "Classroom culture was inclusive and respectful.",               nss_labels: "Expression", similarity: 0.735 },
  ],
  "Academic staff and support": [
    { course_code: "STAT 230",review_text: "Meh — instructor was knowledgeable but not engaging.",          nss_labels: "Staff",      similarity: 0.772 },
    { course_code: "ECON 101",review_text: "Topics covered were relevant to industry needs.",               nss_labels: "Staff",      similarity: 0.758 },
    { course_code: "CS 135",  review_text: "Staff were approachable and responded quickly to emails.",      nss_labels: "Staff",      similarity: 0.742 },
  ],
  "Covid-19 pandemic": [
    { course_code: "PD 1",   review_text: "fk pd — online delivery during covid felt pointless.",          nss_labels: "Covid-19",   similarity: 0.861 },
    { course_code: "CS 240", review_text: "Like the hybrid model actually — more flexible schedule.",       nss_labels: "Covid-19",   similarity: 0.838 },
    { course_code: "ECON 101",review_text: "Remote labs lacked the hands-on aspect of in-person sessions.", nss_labels: "Covid-19",  similarity: 0.812 },
  ],
};

// ── LOAD JSON ─────────────────────────────────────────────────────────────
let data;
if (fs.existsSync(jsonPath)) {
  data = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
  console.log(`✅ Loaded ${jsonPath}`);
} else {
  console.log(`⚠️  ${jsonPath} not found — using sample data`);
  data = SAMPLE_DATA;
}

const themes = Object.keys(data);

// ── HELPERS ───────────────────────────────────────────────────────────────
const BLUE   = "1F4E79";
const LBLUE  = "D5E8F0";
const GREY   = "F2F2F2";
const WHITE  = "FFFFFF";
const DARK   = "2C2C2C";

const border  = (color = "CCCCCC") => ({
  style: BorderStyle.SINGLE, size: 1, color,
});
const allBorders = (color = "CCCCCC") => ({
  top: border(color), bottom: border(color),
  left: border(color), right: border(color),
});

function simBadgeColor(sim) {
  if (sim >= 0.82) return "1A7A4A"; // green
  if (sim >= 0.75) return "B8860B"; // amber
  return "C0392B";                   // red
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial", size: 32, bold: true, color: WHITE })],
    shading: { fill: BLUE, type: ShadingType.CLEAR },
    spacing: { before: 360, after: 200 },
    indent: { left: 180 },
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: BLUE })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } },
    spacing: { before: 320, after: 120 },
  });
}

function bodyPara(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 100 },
    children: [new TextRun({
      text, font: "Arial", size: 22,
      color: DARK, ...opts,
    })],
  });
}

function spacer(after = 160) {
  return new Paragraph({ spacing: { after }, children: [] });
}

// ── THEME SECTION ─────────────────────────────────────────────────────────
function themeSection(theme, chunks) {
  const rows = chunks.map((c, i) => {
    const simColor = simBadgeColor(c.similarity);
    return new TableRow({
      children: [
        // # column
        new TableCell({
          borders: allBorders(),
          width: { size: 400, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          shading: { fill: i % 2 === 0 ? GREY : WHITE, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: String(i + 1), font: "Arial", size: 20, bold: true, color: BLUE })],
          })],
        }),
        // Course
        new TableCell({
          borders: allBorders(),
          width: { size: 1200, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          shading: { fill: i % 2 === 0 ? GREY : WHITE, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: c.course_code, font: "Arial", size: 20, bold: true, color: DARK })],
          })],
        }),
        // Review
        new TableCell({
          borders: allBorders(),
          width: { size: 6560, type: WidthType.DXA },
          shading: { fill: i % 2 === 0 ? GREY : WHITE, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: String(c.review_text), font: "Arial", size: 20, color: DARK })],
          })],
        }),
        // Similarity
        new TableCell({
          borders: allBorders(),
          width: { size: 1200, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          shading: { fill: i % 2 === 0 ? GREY : WHITE, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              text: c.similarity.toFixed(3),
              font: "Arial", size: 20, bold: true, color: simColor,
            })],
          })],
        }),
      ],
    });
  });

  // Header row
  const headerRow = new TableRow({
    tableHeader: true,
    children: ["#", "Course", "Student Review", "Sim Score"].map((label, i) => {
      const widths = [400, 1200, 6560, 1200];
      return new TableCell({
        borders: allBorders(BLUE),
        width: { size: widths[i], type: WidthType.DXA },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: i >= 3 ? AlignmentType.CENTER : AlignmentType.LEFT,
          children: [new TextRun({ text: label, font: "Arial", size: 20, bold: true, color: WHITE })],
        })],
      });
    }),
  });

  const avgSim = chunks.reduce((s, c) => s + c.similarity, 0) / chunks.length;

  return [
    heading2(theme),
    bodyPara(
      `Average relevance score: ${avgSim.toFixed(3)} | ${chunks.length} representative reviews retrieved`,
      { color: "666666", italics: true }
    ),
    spacer(100),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [400, 1200, 6560, 1200],
      rows: [headerRow, ...rows],
    }),
    spacer(200),
  ];
}

// ── TITLE PAGE ────────────────────────────────────────────────────────────
const titleSection = [
  spacer(1440),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 240 },
    children: [new TextRun({
      text: "National Student Survey", font: "Arial",
      size: 64, bold: true, color: BLUE,
    })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
    children: [new TextRun({
      text: "University of Waterloo", font: "Arial",
      size: 40, color: DARK,
    })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 480 },
    children: [new TextRun({
      text: "Student Experience Report", font: "Arial",
      size: 36, italics: true, color: "555555",
    })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE, space: 1 } },
    spacing: { after: 480 },
    children: [],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
    children: [new TextRun({
      text: `Generated: ${new Date().toLocaleDateString("en-GB", { year: "numeric", month: "long", day: "numeric" })}`,
      font: "Arial", size: 24, color: "777777",
    })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
    children: [new TextRun({
      text: `Themes covered: ${themes.length} NSS categories`,
      font: "Arial", size: 24, color: "777777",
    })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 },
    children: [new TextRun({
      text: `Reviews per theme: ${data[themes[0]]?.length || 3} (top by cosine similarity)`,
      font: "Arial", size: 24, color: "777777",
    })],
  }),
  new Paragraph({
    children: [new PageBreak()],
  }),
];

// ── INTRO ─────────────────────────────────────────────────────────────────
const introSection = [
  heading1("Report Overview"),
  spacer(120),
  bodyPara(
    "This report presents student feedback for the University of Waterloo, organised " +
    "according to the National Student Survey (NSS) framework. For each of the " +
    `${themes.length} NSS themes, the top ${data[themes[0]]?.length || 3} most semantically ` +
    "relevant student reviews were retrieved using a RAG pipeline (FAISS index + BGE-small embeddings, " +
    "re-ranked by cosine similarity)."
  ),
  spacer(80),
  bodyPara("How to read the similarity score:", { bold: true }),
  bodyPara("  \u2022  0.82+ (green)  — highly relevant match"),
  bodyPara("  \u2022  0.75\u20130.81 (amber) — good relevance"),
  bodyPara("  \u2022  below 0.75 (red) — lower relevance; interpret with care"),
  spacer(200),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── THEME SECTIONS ────────────────────────────────────────────────────────
const themeSections = themes.flatMap((theme, i) => {
  const chunks = data[theme];
  if (!chunks || chunks.length === 0) return [];
  const section = themeSection(theme, chunks);
  // page break after every theme except the last
  if (i < themes.length - 1) {
    section.push(new Paragraph({ children: [new PageBreak()] }));
  }
  return section;
});

// ── HEADER / FOOTER ───────────────────────────────────────────────────────
const header = new Header({
  children: [
    new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 1 } },
      spacing: { after: 80 },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({ text: "NSS Report — University of Waterloo", font: "Arial", size: 18, color: BLUE, bold: true }),
        new TextRun({ text: "\t", font: "Arial", size: 18 }),
        new TextRun({ text: "Confidential", font: "Arial", size: 18, italics: true, color: "888888" }),
      ],
    }),
  ],
});

const footer = new Footer({
  children: [
    new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC", space: 1 } },
      spacing: { before: 80 },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({ text: `Generated ${new Date().getFullYear()} | RAG-powered`, font: "Arial", size: 16, color: "888888" }),
        new TextRun({ text: "\t", font: "Arial", size: 16 }),
        new TextRun({ text: "Page ", font: "Arial", size: 16, color: "888888" }),
        new PageNumberElement(),
      ],
    }),
  ],
});

// ── BUILD DOC ─────────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: WHITE },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: BLUE },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1 },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      },
    },
    headers: { default: header },
    footers: { default: footer },
    children: [
      ...titleSection,
      ...introSection,
      ...themeSections,
    ],
  }],
});

// ── WRITE ─────────────────────────────────────────────────────────────────
const outPath = "NSS_Report.docx";
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(outPath, buf);
  const kb = (buf.length / 1024).toFixed(1);
  console.log(`\n✅ Report written → ${outPath}  (${kb} KB)`);
  console.log(`   Themes: ${themes.length} | Reviews/theme: ${data[themes[0]]?.length || 3}`);
}).catch(err => {
  console.error("❌ Error:", err.message);
  process.exit(1);
});
