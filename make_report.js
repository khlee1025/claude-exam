/**
 * Confluence MD → Word 보고서 생성기
 * docx-js 기반, Samsung 스타일
 * Usage: node make_report.js <input.md> <output.docx> [title]
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, TableOfContents,
  LevelFormat, ExternalHyperlink, Bookmark,
} = require("docx");

// ── 색상 상수 ─────────────────────────────────────
const SAMSUNG_BLUE  = "1428A0";
const LIGHT_BLUE    = "E8EBF8";
const MID_BLUE      = "D0D6F0";
const HEADER_BG     = "1428A0";
const ROW_ALT       = "F4F6FD";
const GRAY_TEXT     = "6B7280";
const RED_TEXT      = "DC2626";
const BORDER_COLOR  = "C7CDE8";

// ── 셀 보더 ───────────────────────────────────────
function cellBorder(color = BORDER_COLOR) {
  const b = { style: BorderStyle.SINGLE, size: 4, color };
  return { top: b, bottom: b, left: b, right: b };
}

// ── 인라인 마크다운 파싱 (**bold**, *italic*, `code`) ──
function parseInline(text, baseSize = 22) {
  const runs = [];
  const re = /(\*\*(.+?)\*\*|\*(.+?)\*|`(.+?)`)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) runs.push(new TextRun({ text: text.slice(last, m.index), size: baseSize, font: "맑은 고딕" }));
    if (m[2]) runs.push(new TextRun({ text: m[2], bold: true, size: baseSize, font: "맑은 고딕" }));
    else if (m[3]) runs.push(new TextRun({ text: m[3], italics: true, size: baseSize, font: "맑은 고딕" }));
    else if (m[4]) runs.push(new TextRun({ text: m[4], font: "Consolas", size: baseSize - 2, shading: { type: ShadingType.CLEAR, fill: "F3F4F6" } }));
    last = m.index + m[0].length;
  }
  if (last < text.length) runs.push(new TextRun({ text: text.slice(last), size: baseSize, font: "맑은 고딕" }));
  return runs.length ? runs : [new TextRun({ text, size: baseSize, font: "맑은 고딕" })];
}

// ── 헤딩 단락 ──────────────────────────────────────
function makeHeading(text, level) {
  const configs = {
    1: { hlevel: HeadingLevel.HEADING_1, size: 36, color: SAMSUNG_BLUE, before: 360, after: 180, outline: 0 },
    2: { hlevel: HeadingLevel.HEADING_2, size: 30, color: SAMSUNG_BLUE, before: 280, after: 140, outline: 1 },
    3: { hlevel: HeadingLevel.HEADING_3, size: 26, color: "333333",     before: 220, after: 100, outline: 2 },
    4: { hlevel: HeadingLevel.HEADING_4, size: 24, color: "555555",     before: 180, after:  80, outline: 3 },
  };
  const c = configs[level] || configs[4];
  return new Paragraph({
    heading: c.hlevel,
    spacing: { before: c.before, after: c.after },
    children: [new TextRun({ text, bold: true, size: c.size, color: c.color, font: "맑은 고딕" })],
  });
}

// ── 일반 단락 ──────────────────────────────────────
function makeParagraph(text, opts = {}) {
  const { size = 22, spacing = { before: 60, after: 60 }, indent } = opts;
  return new Paragraph({ spacing, indent, children: parseInline(text, size) });
}

// ── 콜아웃 박스 (이미지 분석 결과용) ────────────────
function makeCallout(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    indent: { left: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 20, color: SAMSUNG_BLUE, space: 8 } },
    shading: { type: ShadingType.CLEAR, fill: LIGHT_BLUE },
    children: [new TextRun({ text, size: 20, italics: true, font: "맑은 고딕", color: "2D3A8A" })],
  });
}

// ── 표 생성 ────────────────────────────────────────
function makeTable(headers, rows) {
  const cols = Math.max(headers.length, ...rows.map(r => r.length));
  const totalW = 9026; // A4 content width DXA
  const colW = Math.floor(totalW / cols);
  const colWidths = Array(cols).fill(colW);

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      width: { size: colWidths[i], type: WidthType.DXA },
      borders: cellBorder("FFFFFF"),
      shading: { type: ShadingType.CLEAR, fill: HEADER_BG },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: "FFFFFF", size: 20, font: "맑은 고딕" })] })],
    })),
  });

  const dataRows = rows.map((row, ri) =>
    new TableRow({
      children: Array(cols).fill(null).map((_, ci) => {
        const val = row[ci] || "";
        const isRed = /지연|초과|실패|오류/.test(val);
        const isBold = /정상|완료/.test(val);
        return new TableCell({
          width: { size: colWidths[ci], type: WidthType.DXA },
          borders: cellBorder(BORDER_COLOR),
          shading: { type: ShadingType.CLEAR, fill: ri % 2 === 0 ? "FFFFFF" : ROW_ALT },
          margins: { top: 80, bottom: 80, left: 160, right: 160 },
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({ children: [new TextRun({ text: val, size: 20, font: "맑은 고딕", color: isRed ? RED_TEXT : "222222", bold: isBold })] })],
        });
      }),
    })
  );

  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows],
  });
}

// ── MD 파싱 → docx children 배열 ──────────────────
function parseMD(md) {
  const children = [];
  const lines = md.split("\n");
  let i = 0;
  let tableHeaders = null, tableRows = [];
  let inCode = false, codeLines = [];

  function flushTable() {
    if (!tableHeaders) return;
    children.push(makeTable(tableHeaders, tableRows));
    children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
    tableHeaders = null; tableRows = [];
  }

  while (i < lines.length) {
    const line = lines[i];
    const s = line.trim();

    // 코드 블록
    if (s.startsWith("```")) {
      if (!inCode) { inCode = true; codeLines = []; i++; continue; }
      else {
        inCode = false;
        const code = codeLines.join("\n");
        children.push(new Paragraph({
          spacing: { before: 80, after: 80 },
          shading: { type: ShadingType.CLEAR, fill: "F3F4F6" },
          border: { top: cellBorder(BORDER_COLOR).top, bottom: cellBorder(BORDER_COLOR).bottom, left: cellBorder(BORDER_COLOR).left, right: cellBorder(BORDER_COLOR).right },
          children: [new TextRun({ text: code, font: "Consolas", size: 18, color: "333333" })],
        }));
        i++; continue;
      }
    }
    if (inCode) { codeLines.push(line); i++; continue; }

    // 헤딩
    const hm = s.match(/^(#{1,4})\s+(.+)/);
    if (hm) {
      flushTable();
      children.push(makeHeading(hm[2], hm[1].length));
      i++; continue;
    }

    // 표
    if (s.startsWith("|")) {
      const cells = s.split("|").slice(1, -1).map(c => c.trim());
      if (cells.every(c => /^[-:]+$/.test(c))) { i++; continue; } // 구분선 행
      if (!tableHeaders) { tableHeaders = cells; }
      else { tableRows.push(cells); }
      i++; continue;
    } else { flushTable(); }

    // 이미지 분석 결과 콜아웃
    if (s.startsWith("**[이미지") || s.startsWith("**[Vision")) {
      children.push(makeCallout(s.replace(/\*\*/g, "")));
      i++; continue;
    }

    // 글머리 기호
    const bm = s.match(/^[-*+]\s+(.+)/);
    if (bm) {
      children.push(new Paragraph({
        spacing: { before: 40, after: 40 },
        indent: { left: 480, hanging: 240 },
        children: [new TextRun({ text: "•  ", size: 22, font: "맑은 고딕", color: SAMSUNG_BLUE }), ...parseInline(bm[1])],
      }));
      i++; continue;
    }

    // 번호 목록
    const nm = s.match(/^(\d+)\.\s+(.+)/);
    if (nm) {
      children.push(new Paragraph({
        spacing: { before: 40, after: 40 },
        indent: { left: 480, hanging: 240 },
        children: [new TextRun({ text: `${nm[1]}.  `, bold: true, size: 22, font: "맑은 고딕", color: SAMSUNG_BLUE }), ...parseInline(nm[2])],
      }));
      i++; continue;
    }

    // HR
    if (/^---+$/.test(s)) {
      children.push(new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: MID_BLUE } },
        spacing: { before: 120, after: 120 },
        children: [],
      }));
      i++; continue;
    }

    // 인용/블록쿼트
    const qm = s.match(/^>\s*(.*)/);
    if (qm) {
      children.push(new Paragraph({
        spacing: { before: 80, after: 80 },
        indent: { left: 480 },
        border: { left: { style: BorderStyle.SINGLE, size: 16, color: SAMSUNG_BLUE, space: 8 } },
        children: parseInline(qm[1], 20),
      }));
      i++; continue;
    }

    // 빈 줄
    if (!s) {
      children.push(new Paragraph({ spacing: { before: 20, after: 20 }, children: [] }));
      i++; continue;
    }

    // 일반 문단
    children.push(makeParagraph(s));
    i++;
  }
  flushTable();
  return children;
}

// ── 커버 페이지 ────────────────────────────────────
function makeCoverPage(title, subtitle, date) {
  return [
    new Paragraph({ spacing: { before: 2000, after: 200 }, children: [] }),
    // 상단 삼성 블루 구분선
    new Paragraph({
      border: { top: { style: BorderStyle.SINGLE, size: 30, color: SAMSUNG_BLUE } },
      spacing: { before: 0, after: 600 },
      children: [],
    }),
    // 제목
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: title, bold: true, size: 64, color: SAMSUNG_BLUE, font: "맑은 고딕" })],
    }),
    // 부제목
    ...(subtitle ? [new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 600 },
      children: [new TextRun({ text: subtitle, size: 32, color: GRAY_TEXT, font: "맑은 고딕" })],
    })] : []),
    // 구분선
    new Paragraph({
      border: { bottom: { style: BorderStyle.SINGLE, size: 10, color: MID_BLUE } },
      spacing: { before: 0, after: 400 },
      children: [],
    }),
    // 날짜/작성 정보
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      children: [new TextRun({ text: `작성일: ${date}`, size: 24, color: GRAY_TEXT, font: "맑은 고딕" })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0 },
      children: [new TextRun({ text: "Samsung Confidential", size: 22, color: RED_TEXT, bold: true, font: "맑은 고딕" })],
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ── 메인 ───────────────────────────────────────────
async function main() {
  const [,, mdPath, outPath, titleArg] = process.argv;
  if (!mdPath || !outPath) {
    console.error("Usage: node make_report.js <input.md> <output.docx> [title]");
    process.exit(1);
  }

  const md = fs.readFileSync(mdPath, "utf8");
  const title = titleArg || path.basename(mdPath, ".md");
  const date = new Date().toLocaleDateString("ko-KR", { year:"numeric", month:"long", day:"numeric" });

  const bodyChildren = parseMD(md);

  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: "맑은 고딕", size: 22 } },
      },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 36, bold: true, font: "맑은 고딕", color: SAMSUNG_BLUE },
          paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 30, bold: true, font: "맑은 고딕", color: SAMSUNG_BLUE },
          paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 26, bold: true, font: "맑은 고딕", color: "333333" },
          paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 2 } },
      ],
    },
    numbering: {
      config: [
        { reference: "bullets", levels: [
          { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 480, hanging: 240 } } } },
        ]},
      ],
    },
    sections: [
      // ── 섹션 1: 커버 + 목차 ──
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 },
          },
        },
        children: [
          ...makeCoverPage(title, "Confluence 자동 수집 보고서", date),
          // 목차
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: "목  차", bold: true, size: 36, color: SAMSUNG_BLUE, font: "맑은 고딕" })],
          }),
          new TableOfContents("목차", { hyperlink: true, headingStyleRange: "1-3" }),
          new Paragraph({ children: [new PageBreak()] }),
        ],
      },
      // ── 섹션 2: 본문 ──
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: SAMSUNG_BLUE, space: 1 } },
                spacing: { before: 0, after: 120 },
                children: [
                  new TextRun({ text: title, size: 18, font: "맑은 고딕", color: SAMSUNG_BLUE, bold: true }),
                  new TextRun({ text: "\tSamsung Confidential", size: 18, font: "맑은 고딕", color: GRAY_TEXT }),
                ],
                tabStops: [{ type: "right", position: 9026 }],
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                border: { top: { style: BorderStyle.SINGLE, size: 4, color: MID_BLUE, space: 1 } },
                spacing: { before: 80, after: 0 },
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({ text: `${date}  ·  `, size: 18, font: "맑은 고딕", color: GRAY_TEXT }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "맑은 고딕", color: GRAY_TEXT }),
                  new TextRun({ text: " / ", size: 18, font: "맑은 고딕", color: GRAY_TEXT }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: "맑은 고딕", color: GRAY_TEXT }),
                ],
              }),
            ],
          }),
        },
        children: bodyChildren,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buffer);
  console.log(`✅ 완료: ${outPath}`);
}

main().catch(e => { console.error("❌", e.message); process.exit(1); });

// ── 폴더 모드: 여러 MD → 하나의 Word ─────────────────
async function mainFolder() {
  const [,, folderPath, outPath, titleArg] = process.argv;
  if (!fs.existsSync(folderPath) || !fs.statSync(folderPath).isDirectory()) return false;

  const mdFiles = fs.readdirSync(folderPath)
    .filter(f => f.endsWith(".md") && !f.includes("보고서") && !f.toLowerCase().includes("report"))
    .sort();

  if (!mdFiles.length) { console.error("MD 파일이 없습니다."); process.exit(1); }

  const title = titleArg || path.basename(folderPath) + " 보고서";
  const date = new Date().toLocaleDateString("ko-KR", { year:"numeric", month:"long", day:"numeric" });

  const allChildren = [];
  for (let fi = 0; fi < mdFiles.length; fi++) {
    const fname = mdFiles[fi];
    const md = fs.readFileSync(path.join(folderPath, fname), "utf8");
    const pageTitle = fname.replace(".md", "");
    // 페이지 구분 헤딩
    if (fi > 0) allChildren.push(new Paragraph({ children: [new PageBreak()] }));
    allChildren.push(new Paragraph({
      spacing: { before: 0, after: 200 },
      shading: { type: ShadingType.CLEAR, fill: LIGHT_BLUE },
      children: [new TextRun({ text: `📄 ${pageTitle}`, size: 20, color: GRAY_TEXT, font: "맑은 고딕", italics: true })],
    }));
    allChildren.push(...parseMD(md));
  }
  return { title, date, allChildren };
}
