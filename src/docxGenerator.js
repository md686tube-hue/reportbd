import {
  Document, Paragraph, Table, TableRow, TableCell,
  TextRun, AlignmentType, WidthType, BorderStyle,
  VerticalAlign, Packer, PageOrientation, convertInchesToTwip,
  ShadingType,
} from "docx";
import { saveAs } from "file-saver";

const FONT = "Nikosh";
const SZ = 18;      // 9pt
const SZ_HD = 20;   // 10pt header
const SZ_TITLE = 28; // 14pt title

const BORDER = { style: BorderStyle.SINGLE, size: 4, color: "333333" };
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const ALL_BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };

function run(text, opts = {}) {
  return new TextRun({
    text: String(text || ""),
    font: FONT,
    size: opts.size || SZ,
    bold: opts.bold || false,
    underline: opts.underline ? {} : undefined,
    color: opts.color,
  });
}

function para(children, opts = {}) {
  const runs = Array.isArray(children)
    ? children
    : [run(children, opts)];
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER
      : opts.right ? AlignmentType.RIGHT
      : opts.justify ? AlignmentType.JUSTIFIED
      : AlignmentType.LEFT,
    indent: opts.indent ? { firstLine: 480 } : undefined,
    spacing: opts.spacing || { before: 0, after: 40 },
    children: runs,
  });
}

function cell(text, opts = {}) {
  const shading = opts.shading ? { type: ShadingType.CLEAR, fill: opts.shading } : undefined;
  return new TableCell({
    rowSpan: opts.rowSpan,
    columnSpan: opts.colSpan,
    verticalAlign: VerticalAlign.CENTER,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading,
    borders: opts.borders || ALL_BORDERS,
    margins: { top: 40, bottom: 40, left: 60, right: 60 },
    children: [
      new Paragraph({
        alignment: opts.align || AlignmentType.CENTER,
        spacing: { before: 0, after: 0 },
        children: [run(text, { size: opts.size || SZ, bold: opts.bold, color: opts.color })],
      }),
    ],
  });
}

export async function generateDocx({ UNIONS, currentRows, smarak, dateInfo, masYear, letterBody, toBangla, sumCum, sumField }) {

  const toEn = (s) => parseInt(String(s).replace(/[০-৯]/g, d => "০১২৩৪৫৬৭৮৯".indexOf(d))) || 0;
  const getCumVal = (r, cumField, todayField) => r[cumField] || r[todayField] || "";

  // ===== Table header =====
  const HDR_FILL = "1a5276";
  const HDR_FILL2 = "117a65";
  const HDR_COLOR = "FFFFFF";

  const hRow1 = new TableRow({ tableHeader: true, children: [
    cell("ইউনিয়নের নাম", { rowSpan: 3, bold: true, size: SZ_HD, shading: HDR_FILL, color: HDR_COLOR }),
    cell("জন্ম নিবন্ধনের তথ্য", { colSpan: 6, bold: true, size: SZ_HD, shading: HDR_FILL, color: HDR_COLOR }),
    cell("মৃত্যু নিবন্ধনের তথ্য", { colSpan: 6, bold: true, size: SZ_HD, shading: HDR_FILL2, color: HDR_COLOR }),
  ]});

  const hRow2 = new TableRow({ tableHeader: true, children: [
    cell("মাসিক\nটার্গেট", { rowSpan: 2, bold: true, size: SZ, shading: HDR_FILL, color: HDR_COLOR }),
    cell("৪৫ দিনের মধ্যে", { colSpan: 2, bold: true, size: SZ, shading: HDR_FILL, color: HDR_COLOR }),
    cell("৪৬-৩৬৫ দিনের মধ্যে", { colSpan: 2, bold: true, size: SZ, shading: HDR_FILL, color: HDR_COLOR }),
    cell("মোট", { rowSpan: 2, bold: true, size: SZ, shading: HDR_FILL, color: HDR_COLOR }),
    cell("মাসিক\nটার্গেট", { rowSpan: 2, bold: true, size: SZ, shading: HDR_FILL2, color: HDR_COLOR }),
    cell("৪৫ দিনের মধ্যে", { colSpan: 2, bold: true, size: SZ, shading: HDR_FILL2, color: HDR_COLOR }),
    cell("৪৬-৩৬৫ দিনের মধ্যে", { colSpan: 2, bold: true, size: SZ, shading: HDR_FILL2, color: HDR_COLOR }),
    cell("মোট", { rowSpan: 2, bold: true, size: SZ, shading: HDR_FILL2, color: HDR_COLOR }),
  ]});

  const hRow3 = new TableRow({ tableHeader: true, children: [
    cell("আজ", { bold: true, size: SZ, shading: HDR_FILL, color: HDR_COLOR }),
    cell("এ পর্যন্ত", { bold: true, size: SZ, shading: HDR_FILL, color: HDR_COLOR }),
    cell("আজ", { bold: true, size: SZ, shading: HDR_FILL, color: HDR_COLOR }),
    cell("এ পর্যন্ত", { bold: true, size: SZ, shading: HDR_FILL, color: HDR_COLOR }),
    cell("আজ", { bold: true, size: SZ, shading: HDR_FILL2, color: HDR_COLOR }),
    cell("এ পর্যন্ত", { bold: true, size: SZ, shading: HDR_FILL2, color: HDR_COLOR }),
    cell("আজ", { bold: true, size: SZ, shading: HDR_FILL2, color: HDR_COLOR }),
    cell("এ পর্যন্ত", { bold: true, size: SZ, shading: HDR_FILL2, color: HDR_COLOR }),
  ]});

  const dataRows = UNIONS.map((union, i) => {
    const r = currentRows[i] || {};
    const j45c = getCumVal(r, "j45Cum", "j45Today");
    const j46c = getCumVal(r, "j46Cum", "j46Today");
    const m45c = getCumVal(r, "m45Cum", "m45Today");
    const m46c = getCumVal(r, "m46Cum", "m46Today");
    const jTotal = toEn(j45c) + toEn(j46c);
    const mTotal = toEn(m45c) + toEn(m46c);
    return new TableRow({ children: [
      cell(union.name, { align: AlignmentType.CENTER, size: SZ }),
      cell(toBangla(union.jT), { size: SZ }),
      cell(r.j45Today || "", { size: SZ }),
      cell(j45c, { size: SZ }),
      cell(r.j46Today || "", { size: SZ }),
      cell(j46c, { size: SZ }),
      cell(jTotal > 0 ? toBangla(jTotal) : "", { size: SZ, bold: true }),
      cell(toBangla(union.mT), { size: SZ }),
      cell(r.m45Today || "", { size: SZ }),
      cell(m45c, { size: SZ }),
      cell(r.m46Today || "", { size: SZ }),
      cell(m46c, { size: SZ }),
      cell(mTotal > 0 ? toBangla(mTotal) : "", { size: SZ, bold: true }),
    ]});
  });

  const j45CumTotal = UNIONS.reduce((a, _, i) => { const r = currentRows[i] || {}; return a + toEn(r.j45Cum || r.j45Today || 0); }, 0);
  const j46CumTotal = UNIONS.reduce((a, _, i) => { const r = currentRows[i] || {}; return a + toEn(r.j46Cum || r.j46Today || 0); }, 0);
  const m45CumTotal = UNIONS.reduce((a, _, i) => { const r = currentRows[i] || {}; return a + toEn(r.m45Cum || r.m45Today || 0); }, 0);
  const m46CumTotal = UNIONS.reduce((a, _, i) => { const r = currentRows[i] || {}; return a + toEn(r.m46Cum || r.m46Today || 0); }, 0);

  const totalRow = new TableRow({ children: [
    cell("সর্বমোট", { bold: true, size: SZ }),
    cell(toBangla(UNIONS.reduce((a, u) => a + u.jT, 0)), { bold: true, size: SZ }),
    cell(toBangla(UNIONS.reduce((a, _, i) => a + toEn((currentRows[i] || {}).j45Today || 0), 0)), { bold: true, size: SZ }),
    cell(toBangla(j45CumTotal), { bold: true, size: SZ }),
    cell(toBangla(UNIONS.reduce((a, _, i) => a + toEn((currentRows[i] || {}).j46Today || 0), 0)), { bold: true, size: SZ }),
    cell(toBangla(j46CumTotal), { bold: true, size: SZ }),
    cell(toBangla(j45CumTotal + j46CumTotal), { bold: true, size: SZ }),
    cell(toBangla(UNIONS.reduce((a, u) => a + u.mT, 0)), { bold: true, size: SZ }),
    cell(toBangla(UNIONS.reduce((a, _, i) => a + toEn((currentRows[i] || {}).m45Today || 0), 0)), { bold: true, size: SZ }),
    cell(toBangla(m45CumTotal), { bold: true, size: SZ }),
    cell(toBangla(UNIONS.reduce((a, _, i) => a + toEn((currentRows[i] || {}).m46Today || 0), 0)), { bold: true, size: SZ }),
    cell(toBangla(m46CumTotal), { bold: true, size: SZ }),
    cell(toBangla(m45CumTotal + m46CumTotal), { bold: true, size: SZ }),
  ]});

  const mainTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [hRow1, hRow2, hRow3, ...dataRows, totalRow],
  });

  // ===== Smarak row (borderless table) =====
  const smarakTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDERS,
    rows: [new TableRow({ children: [
      new TableCell({
        borders: NO_BORDERS,
        children: [para(smarak, { size: SZ })],
      }),
      new TableCell({
        borders: NO_BORDERS,
        children: [
          para(`তারিখ: ${dateInfo.bangla}`, { right: true, size: SZ, spacing: { before: 0, after: 0 } }),
          para(dateInfo.english, { right: true, size: SZ, spacing: { before: 0, after: 0 } }),
        ],
      }),
    ]})],
  });

  // ===== Signature section =====
  const sigTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: NO_BORDERS,
    rows: [
      // Top row: নাম ডানে
      new TableRow({ children: [
        new TableCell({ borders: NO_BORDERS, children: [para("")] }),
        new TableCell({ borders: NO_BORDERS, children: [
          para("", { spacing: { before: 200, after: 0 } }),
          para("মোঃ আশরাফুল আলম রাসেল", { center: true, bold: true, size: SZ, spacing: { before: 0, after: 0 } }),
          para("উপজেলা নির্বাহী অফিসার", { center: true, size: SZ, spacing: { before: 0, after: 0 } }),
          para("জীবননগর, চুয়াডাঙ্গা।", { center: true, size: SZ, spacing: { before: 0, after: 0 } }),
          para("☎ ০২৪৭৭৭৮৯৭০১", { center: true, size: SZ, spacing: { before: 0, after: 0 } }),
          para("unojibannagar@mopa.gov.bd", { center: true, size: SZ, spacing: { before: 0, after: 0 } }),
        ]}),
      ]}),
      // Bottom row: বামে জেলা প্রশাসক
      new TableRow({ children: [
        new TableCell({ borders: NO_BORDERS, children: [
          para("জেলা প্রশাসক", { size: SZ, spacing: { before: 0, after: 0 } }),
          para("চুয়াডাঙ্গা।", { size: SZ, spacing: { before: 0, after: 0 } }),
          para("দুঃআঃ উপ-পরিচালক, স্থানীয় সরকার, চুয়াডাঙ্গা।", { size: SZ, spacing: { before: 0, after: 0 } }),
        ]},
        new TableCell({ borders: NO_BORDERS, children: [para("")] }),
      ]}),
    ],
  });

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { orientation: PageOrientation.LANDSCAPE, width: convertInchesToTwip(11.69), height: convertInchesToTwip(8.27) },
          margin: { top: convertInchesToTwip(0.5), bottom: convertInchesToTwip(0.5), left: convertInchesToTwip(0.6), right: convertInchesToTwip(0.6) },
        },
      },
      children: [
        para("গণপ্রজাতন্ত্রী বাংলাদেশ সরকার", { center: true, bold: true, size: SZ_TITLE, spacing: { before: 0, after: 40 } }),
        para("উপজেলা নির্বাহী অফিসারের কার্যালয়", { center: true, size: SZ_HD, spacing: { before: 0, after: 40 } }),
        para("জীবননগর, চুয়াডাঙ্গা", { center: true, size: SZ_HD, spacing: { before: 0, after: 40 } }),
        para("www.jibannagar.chuadanga.gov.bd", { center: true, size: 16, spacing: { before: 0, after: 80 } }),
        smarakTable,
        para(""),
        new Paragraph({
          spacing: { before: 0, after: 40 },
          children: [run("বিষয়: ", { bold: true, size: SZ }), run("জন্ম ও মৃত্যু নিবন্ধনের লক্ষমাত্রা অনুযায়ী প্রতিদিনের তথ্য প্রেরণ", { size: SZ })],
        }),
        new Paragraph({
          spacing: { before: 0, after: 40 },
          children: [run("সূত্র : ", { bold: true, size: SZ }), run("জেলা প্রশাসক, চুয়াডাঙ্গা মহোদয়ের ২১/০৯/২০২২ তারিখের ০৫.৪৪.১৮০০.১০৬.১৪.০০৮.২২-৫০৩(৮২) নং স্মারক", { size: SZ })],
        }),
        para(""),
        para(letterBody, { justify: true, indent: true, size: SZ, spacing: { before: 0, after: 120 } }),
        para("উপজেলার নাম: জীবননগর", { center: true, bold: true, underline: true, size: SZ_HD, spacing: { before: 0, after: 80 } }),
        mainTable,
        para(""),
        sigTable,
      ],
    }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `nibondhon_${dateInfo.english || "report"}.docx`);
}
