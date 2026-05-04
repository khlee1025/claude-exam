/**
 * Confluence MD → Word 보고서 생성기 (Samsung Style)
 * Usage:
 *   node make_report.js <input.md|folder> <output.docx> [title]
 */
"use strict";
const fs   = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, TableOfContents,
  LevelFormat,
} = require("docx");

// ─── 색상 ──────────────────────────────────────────
const C = {
  blue:      "1428A0",
  lightBlue: "E8EBF8",
  midBlue:   "C7CDE8",
  rowAlt:    "F4F6FD",
  gray:      "6B7280",
  red:       "DC2626",
  green:     "16A34A",
  white:     "FFFFFF",
  dark:      "1F2937",
};

// ─── 유틸 ──────────────────────────────────────────
function border(color = C.midBlue, size = 4) {
  const b = { style: BorderStyle.SINGLE, size, color };
  return { top: b, bottom: b, left: b, right: b };
}

function parseInline(text, size = 22, defaultColor = C.dark) {
  const runs = [];
  const re   = /(\*\*(.+?)\*\*|\*(.+?)\*|`([^`]+)`)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last)
      runs.push(new TextRun({ text: text.slice(last, m.index), size, font: "맑은 고딕", color: defaultColor }));
    if      (m[2]) runs.push(new TextRun({ text: m[2], bold: true,    size, font: "맑은 고딕", color: defaultColor }));
    else if (m[3]) runs.push(new TextRun({ text: m[3], italics: true, size, font: "맑은 고딕", color: defaultColor }));
    else if (m[4]) runs.push(new TextRun({ text: m[4], font: "Consolas", size: size - 2,
      shading: { type: ShadingType.CLEAR, fill: "F3F4F6" }, color: "B91C1C" }));
    last = m.index + m[0].length;
  }
  if (last < text.length)
    runs.push(new TextRun({ text: text.slice(last), size, font: "맑은 고딕", color: defaultColor }));
  return runs.length ? runs : [new TextRun({ text, size, font: "맑은 고딕", color: defaultColor })];
}

// ─── 헤딩 ──────────────────────────────────────────
const HEADING_CFG = {
  1: { level: HeadingLevel.HEADING_1, size: 36, color: C.blue,  before: 400, after: 200, outline: 0 },
  2: { level: HeadingLevel.HEADING_2, size: 30, color: C.blue,  before: 300, after: 160, outline: 1 },
  3: { level: HeadingLevel.HEADING_3, size: 26, color: "333333",before: 240, after: 120, outline: 2 },
  4: { level: HeadingLevel.HEADING_4, size: 24, color: "555555",before: 180, after:  80, outline: 3 },
};
function makeHeading(text, depth) {
  const c = HEADING_CFG[Math.min(depth, 4)] || HEADING_CFG[4];
  return new Paragraph({
    heading: c.level,
    spacing: { before: c.before, after: c.after },
    children: [new TextRun({ text, bold: true, size: c.size, color: c.color, font: "맑은 고딕" })],
  });
}

// ─── 표 ────────────────────────────────────────────
function makeTable(headers, rows) {
  const cols   = Math.max(headers.length, ...rows.map(r => r.length), 1);
  const totalW = 9026;
  const colW   = Math.floor(totalW / cols);
  const colWs  = Array(cols).fill(colW);

  const hdrRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      width:   { size: colWs[i], type: WidthType.DXA },
      borders: border(C.blue),
      shading: { type: ShadingType.CLEAR, fill: C.blue },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: h, bold: true, color: C.white, size: 20, font: "맑은 고딕" })] })],
    })),
  });

  const dataRows = rows.map((row, ri) =>
    new TableRow({
      children: Array(cols).fill(null).map((_, ci) => {
        const val   = (row[ci] || "").trim();
        const isRed = /지연|초과|실패|오류|경고/.test(val);
        const isGrn = /정상|완료|성공/.test(val);
        return new TableCell({
          width:   { size: colWs[ci], type: WidthType.DXA },
          borders: border(C.midBlue),
          shading: { type: ShadingType.CLEAR, fill: ri % 2 === 0 ? C.white : C.rowAlt },
          margins: { top: 80, bottom: 80, left: 160, right: 160 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ children: [new TextRun({
            text: val, size: 20, font: "맑은 고딕",
            color: isRed ? C.red : isGrn ? C.green : C.dark,
            bold: isRed || isGrn,
          })] })],
        });
      }),
    })
  );

  return new Table({ width: { size: totalW, type: WidthType.DXA }, columnWidths: colWs, rows: [hdrRow, ...dataRows] });
}

// ─── 콜아웃 (이미지 분석) ─────────────────────────
function makeCallout(text) {
  const clean = text.replace(/\*\*\[이미지[^\]]*\]\*\*\s*/g, "").replace(/\*\*/g, "").trim();
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    indent:  { left: 360 },
    border:  { left: { style: BorderStyle.THICK, size: 20, color: C.blue, space: 8 } },
    shading: { type: ShadingType.CLEAR, fill: C.lightBlue },
    children: [
      new TextRun({ text: "🔍 이미지 분석  ", bold: true, size: 20, font: "맑은 고딕", color: C.blue }),
      new TextRun({ text: clean, size: 20, font: "맑은 고딕", color: "2D3A8A", italics: true }),
    ],
  });
}

// ─── 페이지 출처 배너 ─────────────────────────────
function makeSourceBanner(pageTitle) {
  return new Paragraph({
    spacing: { before: 0, after: 160 },
    shading: { type: ShadingType.CLEAR, fill: C.lightBlue },
    border:  { left: { style: BorderStyle.THICK, size: 16, color: C.blue, space: 6 } },
    indent:  { left: 160 },
    children: [new TextRun({ text: `📄  ${pageTitle}`, size: 19, font: "맑은 고딕", color: C.gray, italics: true })],
  });
}

// ─── MD → children 변환 ───────────────────────────
function parseMD(md) {
  const children = [];
  const lines    = md.split("\n");
  let i = 0, inCode = false, codeLines = [];
  let thdrs = null, trows = [];

  function flushTable() {
    if (!thdrs) return;
    children.push(makeTable(thdrs, trows));
    children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
    thdrs = null; trows = [];
  }

  while (i < lines.length) {
    const line = lines[i], s = line.trim();

    // 코드 블록
    if (s.startsWith("```")) {
      if (!inCode) { inCode = true; codeLines = []; i++; continue; }
      inCode = false;
      children.push(new Paragraph({
        spacing: { before: 80, after: 80 },
        shading: { type: ShadingType.CLEAR, fill: "F3F4F6" },
        border:  border(C.midBlue, 2),
        indent:  { left: 240 },
        children: [new TextRun({ text: codeLines.join("\n"), font: "Consolas", size: 18, color: "1F2937" })],
      }));
      i++; continue;
    }
    if (inCode) { codeLines.push(line); i++; continue; }

    // 헤딩
    const hm = s.match(/^(#{1,4})\s+(.+)/);
    if (hm) { flushTable(); children.push(makeHeading(hm[2], hm[1].length)); i++; continue; }

    // 표
    if (s.startsWith("|")) {
      const cells = s.split("|").slice(1, -1).map(c => c.trim());
      if (cells.every(c => /^[-: ]+$/.test(c))) { i++; continue; }
      if (!thdrs) thdrs = cells; else trows.push(cells);
      i++; continue;
    } else { flushTable(); }

    // 이미지 분석 콜아웃
    if (s.startsWith("**[이미지") || s.startsWith("**[Vision")) {
      children.push(makeCallout(s)); i++; continue;
    }

    // 글머리
    const bm = s.match(/^[-*+]\s+(.+)/);
    if (bm) {
      children.push(new Paragraph({
        spacing: { before: 40, after: 40 },
        indent:  { left: 480, hanging: 240 },
        children: [new TextRun({ text: "•  ", size: 22, font: "맑은 고딕", color: C.blue, bold: true }), ...parseInline(bm[1])],
      }));
      i++; continue;
    }

    // 번호 목록
    const nm = s.match(/^(\d+)\.\s+(.+)/);
    if (nm) {
      children.push(new Paragraph({
        spacing: { before: 40, after: 40 },
        indent:  { left: 480, hanging: 280 },
        children: [new TextRun({ text: `${nm[1]}.  `, size: 22, font: "맑은 고딕", color: C.blue, bold: true }), ...parseInline(nm[2])],
      }));
      i++; continue;
    }

    // 인용
    const qm = s.match(/^>\s*(.*)/);
    if (qm) {
      children.push(new Paragraph({
        spacing: { before: 80, after: 80 }, indent: { left: 480 },
        border:  { left: { style: BorderStyle.SINGLE, size: 16, color: C.blue, space: 8 } },
        shading: { type: ShadingType.CLEAR, fill: C.lightBlue },
        children: parseInline(qm[1], 20),
      }));
      i++; continue;
    }

    // HR
    if (/^-{3,}$/.test(s)) {
      children.push(new Paragraph({
        border:  { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.midBlue } },
        spacing: { before: 120, after: 120 },
        children: [],
      }));
      i++; continue;
    }

    // 빈 줄
    if (!s) {
      children.push(new Paragraph({ spacing: { before: 20, after: 20 }, children: [] }));
      i++; continue;
    }

    // 일반 문단
    children.push(new Paragraph({ spacing: { before: 60, after: 60 }, children: parseInline(s) }));
    i++;
  }
  flushTable();
  return children;
}

// ─── 커버 페이지 ──────────────────────────────────
function makeCover(title, subtitle, date) {
  return [
    new Paragraph({ spacing: { before: 1800 }, children: [] }),
    new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 40, color: C.blue } },
      spacing: { after: 800 }, children: [],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 240 },
      children: [new TextRun({ text: title, bold: true, size: 72, color: C.blue, font: "맑은 고딕" })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 600 },
      children: [new TextRun({ text: subtitle, size: 32, color: C.gray, font: "맑은 고딕" })],
    }),
    new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.midBlue } },
      spacing: { after: 400 }, children: [],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 200 },
      children: [new TextRun({ text: date, size: 26, color: C.gray, font: "맑은 고딕" })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Samsung Confidential", size: 22, color: C.red, bold: true, font: "맑은 고딕" })],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ─── 헤더 / 푸터 ──────────────────────────────────
function makeHeader(title) {
  return new Header({ children: [new Paragraph({
    border:  { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.blue, space: 1 } },
    spacing: { after: 100 },
    tabStops: [{ type: "right", position: 9026 }],
    children: [
      new TextRun({ text: title,                   size: 18, font: "맑은 고딕", color: C.blue, bold: true }),
      new TextRun({ text: "\tSamsung Confidential", size: 18, font: "맑은 고딕", color: C.gray }),
    ],
  })] });
}

function makeFooter(date) {
  return new Footer({ children: [new Paragraph({
    border:    { top: { style: BorderStyle.SINGLE, size: 4, color: C.midBlue, space: 1 } },
    spacing:   { before: 80 },
    alignment: AlignmentType.CENTER,
    children: [
      new TextRun({ text: `${date}  ·  `, size: 18, font: "맑은 고딕", color: C.gray }),
      new TextRun({ children: [PageNumber.CURRENT],     size: 18, font: "맑은 고딕", color: C.gray }),
      new TextRun({ text: " / ",                         size: 18, font: "맑은 고딕", color: C.gray }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: "맑은 고딕", color: C.gray }),
    ],
  })] });
}

// ─── 페이지 설정 ──────────────────────────────────
const PAGE_PROPS = {
  page: {
    size:   { width: 11906, height: 16838 },          // A4
    margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 },
  },
};

// ─── 문서 스타일 ──────────────────────────────────
function docStyles() {
  return {
    default: { document: { run: { font: "맑은 고딕", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "맑은 고딕", color: C.blue },
        paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "맑은 고딕", color: C.blue },
        paragraph: { spacing: { before: 300, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "맑은 고딕", color: "333333" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2 } },
    ],
  };
}

// ─── 소스 로드 (파일 or 폴더) ────────────────────
function loadSources(src) {
  const stat = fs.statSync(src);
  if (stat.isFile()) {
    return [{ title: path.basename(src, ".md"), md: fs.readFileSync(src, "utf8") }];
  }
  // 폴더: .md 파일 정렬 로드 (보고서/report 파일 제외)
  return fs.readdirSync(src)
    .filter(f => f.endsWith(".md") && !/보고서|report/i.test(f))
    .sort()
    .map(f => ({
      title: f.replace(".md", ""),
      md:    fs.readFileSync(path.join(src, f), "utf8"),
    }));
}

// ─── 메인 ─────────────────────────────────────────
async function main() {
  const [,, src, outPath, titleArg] = process.argv;
  if (!src || !outPath) {
    console.error("Usage: node make_report.js <file.md|folder> <output.docx> [title]");
    process.exit(1);
  }

  const sources  = loadSources(src);
  const title    = titleArg || path.basename(src, ".md") || "Confluence 보고서";
  const date     = new Date().toLocaleDateString("ko-KR", { year: "numeric", month: "long", day: "numeric" });
  const subtitle = `Confluence 자동 수집 보고서  ·  총 ${sources.length}개 페이지`;

  console.log(`📄 소스 ${sources.length}개 로드 완료`);

  // 본문 children 구성
  const bodyChildren = [];
  sources.forEach((s, idx) => {
    if (idx > 0) bodyChildren.push(new Paragraph({ children: [new PageBreak()] }));
    if (sources.length > 1) bodyChildren.push(makeSourceBanner(s.title));
    bodyChildren.push(...parseMD(s.md));
  });

  const doc = new Document({
    styles:   docStyles(),
    sections: [
      // ── 섹션1: 커버 + 목차 ──
      {
        properties: { page: PAGE_PROPS.page },
        children: [
          ...makeCover(title, subtitle, date),
          makeHeading("목  차", 1),
          new TableOfContents("목차", { hyperlink: true, headingStyleRange: "1-3" }),
          new Paragraph({ children: [new PageBreak()] }),
        ],
      },
      // 섹션2: 본문
      {
        properties: { page: PAGE_PROPS.page },
        headers:    { default: makeHeader(title)  },
        footers:    { default: makeFooter(date)   },
        children:   bodyChildren,
      },
    ],
  });

  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buf);
  console.log(`✅ Word 보고서 생성 완료: ${outPath}`);
}

main().catch(e => { console.error("❌", e.message); process.exit(1); });
