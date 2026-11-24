// Shared docx layout helpers used by build_handout.js and build_workbook.js.
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, LevelFormat, AlignmentType, BorderStyle,
} = require("docx");

const FONT = "Microsoft JhengHei";
const W = 9638; // A4 content width (DXA) with 0.79" margins

const p = (text, opts = {}) =>
  new Paragraph({ spacing: { after: 100, line: 300 }, ...opts, children: [new TextRun({ text, font: FONT, ...(opts.run || {}) })] });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const bl = (t) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 60, line: 290 }, children: [new TextRun({ text: t, font: FONT })] });
const note = (t) => new Paragraph({
  spacing: { before: 60, after: 140 }, indent: { left: 200 },
  border: { left: { style: BorderStyle.SINGLE, size: 12, color: "1F4E79", space: 8 } },
  children: [new TextRun({ text: t, font: FONT, size: 20, color: "444444" })],
});
const code = (t) => t.split("\n").map((l) => new Paragraph({
  spacing: { after: 0, line: 252 }, shading: { type: ShadingType.CLEAR, fill: "F2F4F7" },
  children: [new TextRun({ text: l || " ", font: "Consolas", size: 18 })],
}));

function table(widths, header, rows) {
  const cell = (t, w, head) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: head ? { type: ShadingType.CLEAR, fill: "1F4E79" } : undefined,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    children: [new Paragraph({ children: [new TextRun({ text: String(t), font: FONT, size: 19, bold: head, color: head ? "FFFFFF" : "000000" })] })],
  });
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: header.map((t, i) => cell(t, widths[i], true)) }),
      ...rows.map((r) => new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, widths[i], false)) })),
    ],
  });
}
const gap = () => new Paragraph({ spacing: { after: 120 }, children: [] });

function makeDoc(title, children) {
  return new Document({
    creator: "CRM Training",
    title,
    styles: {
      default: { document: { run: { font: FONT, size: 21 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 32, bold: true, color: "1F4E79" },
          paragraph: { spacing: { before: 360, after: 160 }, keepNext: true, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 25, bold: true, color: "2E2E2E" },
          paragraph: { spacing: { before: 220, after: 100 }, keepNext: true, outlineLevel: 1 } },
      ],
    },
    numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } },
      children,
    }],
  });
}

function save(doc, out) {
  fs.mkdirSync(path.dirname(out), { recursive: true });
  return Packer.toBuffer(doc).then((b) => { fs.writeFileSync(out, b); console.log("wrote", out); });
}

module.exports = { FONT, W, p, h1, h2, bl, note, code, table, gap, makeDoc, save };
