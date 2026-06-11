// backend/utils/docxBuilder.js
// Converts plain report text into a formatted .docx Buffer.

import {
  Document, Packer, Paragraph, TextRun, LevelFormat,
  AlignmentType, BorderStyle, Header, Footer, TabStopType, SimpleField,
} from "docx";

const NAVY = "1F3864";
const BLUE = "2E75B6";
const DGRAY = "444444";
const MGRAY = "666666";
const sp = (b, a) => ({ before: b * 20, after: a * 20 });

function classifyLine(line) {
  const t = line.trim();
  if (!t) return null;
  if (/^SECTION\s+\d+\s*[—–-]/.test(t)) return { type: "section", text: t };
  if (/^#{1,2}\s/.test(t) || /^\d+\.\d+\s/.test(t)) return { type: "h2", text: t.replace(/^#+\s*/, "") };
  if (/^[•\-\*]\s/.test(t)) return { type: "bullet", text: t.replace(/^[•\-\*]\s*/, "") };
  return { type: "body", text: t };
}

function buildParagraph(classified) {
  if (!classified) return new Paragraph({ children: [new TextRun({ text: "", size: 4 })] });
  const { type, text } = classified;
  if (type === "section")
    return new Paragraph({
      spacing: sp(10, 4),
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 1 } },
      children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: NAVY })],
    });
  if (type === "h2")
    return new Paragraph({
      spacing: sp(6, 2),
      children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: BLUE })],
    });
  if (type === "bullet")
    return new Paragraph({
      numbering: { reference: "bullets", level: 0 },
      spacing: sp(1, 1),
      children: [new TextRun({ text, font: "Arial", size: 20, color: DGRAY })],
    });
  return new Paragraph({
    spacing: sp(2, 2),
    children: [new TextRun({ text, font: "Arial", size: 20, color: DGRAY })],
  });
}

export async function buildReportDocx({ reportText, name, dept, week, manager }) {
  const coverChildren = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: sp(0, 8),
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE, space: 1 } },
      children: [new TextRun({ text: "ESDS SOFTWARE SOLUTION PVT. LTD.", font: "Arial", size: 18, bold: true, color: MGRAY, allCaps: true })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: sp(4, 4),
      children: [new TextRun({ text: "WEEKLY WORK REPORT", font: "Arial", size: 44, bold: true, color: NAVY })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: sp(0, 20),
      children: [new TextRun({ text: `${name}  ·  ${dept}  ·  Week Ending: ${week}  ·  Manager: ${manager}`, font: "Arial", size: 20, color: DGRAY })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: sp(0, 24),
      children: [new TextRun({ text: `Generated: ${new Date().toLocaleString("en-IN")}`, font: "Arial", size: 16, color: MGRAY, italics: true })],
    }),
  ];

  const bodyChildren = reportText.split("\n").map((line) => buildParagraph(classifyLine(line)));

  const doc = new Document({
    numbering: {
      config: [{
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 600, hanging: 300 } } } }],
      }],
    },
    sections: [{
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 1 } },
            tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
            children: [
              new TextRun({ text: `Weekly Report — ${name}  |  ${dept}  |  ${week}`, font: "Arial", size: 16, color: MGRAY }),
              new TextRun({ text: "\tESDS Software Solution Pvt. Ltd.", font: "Arial", size: 16, color: MGRAY }),
            ],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 1 } },
            tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
            children: [
              new TextRun({ text: "Confidential — Internal Use Only", font: "Arial", size: 16, color: MGRAY }),
              new TextRun({ text: "\tPage ", font: "Arial", size: 16, color: MGRAY }),
              new SimpleField("PAGE", { font: "Arial", size: 16, color: MGRAY }),
            ],
          })],
        }),
      },
      children: [...coverChildren, ...bodyChildren],
    }],
  });

  return Packer.toBuffer(doc);
}
