import {
  Document, Paragraph, Table, TableRow, TableCell,
  TextRun, AlignmentType, WidthType, BorderStyle,
  VerticalAlign, Packer
} from "docx";
import { saveAs } from "file-saver";

const font = "Nikosh";
const sz = 18;

function cell(text, opts = {}) {
  return new TableCell({
    rowSpan: opts.rowSpan,
    columnSpan: opts.colSpan,
    verticalAlign: VerticalAlign.CENTER,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    children: [
      new Paragraph({
        alignment: opts.align || AlignmentType.CENTER,
        children: [
          new TextRun({
            text: String(text || ""),
            font,
            size: opts.size || sz,
            bold: opts.bold,
          }),
        ],
      }),
    ],
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : opts.right ? AlignmentType.RIGHT : opts.justify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
    indent: opts.indent ? { firstLine: 480 } : undefined,
    children: [
      new TextRun({
        text: String(text || ""),
        font,
        size: opts.size || sz,
        bold: opts.bold,
        underline: opts.underline ? {} : undefined,
      }),
    ],
  });
}

export async function generateDocx({ UNIONS, currentRows, smarak, dateInfo, masYear, letterBody, toBangla, sumField }) {
  const toEn = (s) => String(s).replace(/[০-৯]/g, (d) => "০১২৩৪৫৬৭৮৯".indexOf(d));

  // Header rows for the big table
  const hRow1 = new TableRow({
    children: [
      cell("ইউনিয়নের নাম", { rowSpan: 3, bold: true }),
      cell("জন্ম নিবন্ধনের তথ্য", { colSpan: 5, bold: true }),
      cell("মৃত্যু নিবন্ধনের তথ্য", { colSpan: 5, bold: true }),
      cell("মোট", { rowSpan: 3, bold: true }),
    ],
  });

  const hRow2 = new TableRow({
    children: [
      cell("মাসিক\nটার্গেট", { rowSpan: 2, bold: true }),
      cell("৪৫ দিনের মধ্যে", { colSpan: 2, bold: true }),
      cell("৪৬-৩৬৫ দিনের মধ্যে", { colSpan: 2, bold: true }),
      cell("মাসিক\nটার্গেট", { rowSpan: 2, bold: true }),
      cell("৪৫ দিনের মধ্যে", { colSpan: 2, bold: true }),
      cell("৪৬-৩৬৫ দিনের মধ্যে", { colSpan: 2, bold: true }),
    ],
  });

  const hRow3 = new TableRow({
    children: ["আজ", "এ পর্যন্ত", "আজ", "এ পর্যন্ত", "আজ", "এ পর্যন্ত", "আজ", "এ পর্যন্ত"].map(t => cell(t, { bold: true })),
  });

  const dataRows = UNIONS.map((union, i) => {
    const r = currentRows[i] || {};
    const jTotal = (parseInt(toEn(r.j45Today || "0")) || 0) + (parseInt(toEn(r.j46Today || "0")) || 0) +
      (parseInt(toEn(r.m45Today || "0")) || 0) + (parseInt(toEn(r.m46Today || "0")) || 0);
    return new TableRow({
      children: [
        cell(union.name, { align: AlignmentType.LEFT }),
        cell(toBangla(union.jT)),
        cell(r.j45Today || ""), cell(r.j45Cum || ""),
        cell(r.j46Today || ""), cell(r.j46Cum || ""),
        cell(toBangla(union.mT)),
        cell(r.m45Today || ""), cell(r.m45Cum || ""),
        cell(r.m46Today || ""), cell(r.m46Cum || ""),
        cell(jTotal > 0 ? toBangla(jTotal) : ""),
      ],
    });
  });

  const totalRow = new TableRow({
    children: [
      cell("সর্বমোট", { bold: true }),
      cell(toBangla(UNIONS.reduce((a, u) => a + u.jT, 0)), { bold: true }),
      cell(toBangla(sumField("j45Today")), { bold: true }),
      cell(toBangla(sumField("j45Cum")), { bold: true }),
      cell(toBangla(sumField("j46Today")), { bold: true }),
      cell(toBangla(sumField("j46Cum")), { bold: true }),
      cell(toBangla(UNIONS.reduce((a, u) => a + u.mT, 0)), { bold: true }),
      cell(toBangla(sumField("m45Today")), { bold: true }),
      cell(toBangla(sumField("m45Cum")), { bold: true }),
      cell(toBangla(sumField("m46Today")), { bold: true }),
      cell(toBangla(sumField("m46Cum")), { bold: true }),
      cell(toBangla(sumField("j45Today") + sumField("j46Today") + sumField("m45Today") + sumField("m46Today")), { bold: true }),
    ],
  });

  const mainTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [hRow1, hRow2, hRow3, ...dataRows, totalRow],
  });

  // Signature table
  const sigTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1 } },
            children: [
              para("মোঃ আশরাফুল আলম রাসেল", { bold: true, size: sz }),
              para("উপজেলা নির্বাহী অফিসার", { size: sz }),
              para("জীবননগর, চুয়াডাঙ্গা", { size: sz }),
              para("০২৪৭৭৭৮৯৭০১", { size: sz }),
            ],
          }),
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1 } },
            children: [para("জেলা প্রশাসক", { size: sz }), para("চুয়াডাঙ্গা", { size: sz })],
          }),
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 1 } },
            children: [para("দুঃআঃ উপ-পরিচালক, স্থানীয় সরকার, চুয়াডাঙ্গা", { size: sz })],
          }),
        ],
      }),
    ],
  });

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: { top: 720, bottom: 720, left: 1080, right: 1080 },
        },
      },
      children: [
        para("গণপ্রজাতন্ত্রী বাংলাদেশ সরকার", { bold: true, center: true, size: 26 }),
        para("উপজেলা নির্বাহী অফিসারের কার্যালয়", { center: true, size: sz }),
        para("জীবননগর, চুয়াডাঙ্গা", { center: true, size: sz }),
        para("www.jibannagor.chuadanga.gov.bd", { center: true, size: 14 }),
        para(""),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE }, insideH: { style: BorderStyle.NONE }, insideV: { style: BorderStyle.NONE } },
          rows: [
            new TableRow({
              children: [
                new TableCell({ borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, children: [para(smarak, { size: sz })] }),
                new TableCell({ borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, children: [para(`তারিখ: ${dateInfo.bangla}`, { right: true, size: sz })] }),
              ],
            }),
          ],
        }),
        para(""),
        new Paragraph({
          children: [
            new TextRun({ text: "বিষয়: ", font, size: sz, bold: true }),
            new TextRun({ text: "জন্ম ও মৃত্যু নিবন্ধনের লক্ষমাত্রা অনুযায়ী প্রতিদিনের তথ্য প্রেরণ", font, size: sz }),
          ],
        }),
        new Paragraph({
          children: [
            new TextRun({ text: "সূত্র : ", font, size: sz, bold: true }),
            new TextRun({ text: "জেলা প্রশাসক, চুয়াডাঙ্গা মহোদয়ের ২১/০৯/২০২২ তারিখের ০৫.৪৪.১৮০০.১০৬.১৪.০০৮.২২-৫০৩(৮২) নং স্মারক", font, size: sz }),
          ],
        }),
        para(""),
        para(letterBody, { justify: true, indent: true, size: sz }),
        para(""),
        para("উপজেলার নাম: জীবননগর", { bold: true, center: true, underline: true, size: sz }),
        para(""),
        mainTable,
        para(""),
        sigTable,
      ],
    }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `nibondhon_${dateInfo.english || "report"}.docx`);
}
