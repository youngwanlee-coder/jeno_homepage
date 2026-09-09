/**
 * Sweet Balance 문서 공용 양식 라이브러리
 * 브랜드: Primary Green #007C5E / White / 맑은 고딕
 * 사용: const sb = require('../_lib/sbdoc')({ landscape: true });
 */
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, HeadingLevel, PageBreak, PageOrientation, VerticalAlign,
  BorderStyle, ImageRun, Header, Footer, PageNumber, LevelFormat, convertMillimetersToTwip,
} = require('docx');

const GREEN = '007C5E';
const WHITE = 'FFFFFF';
const DARK  = '212121';
const GREY  = 'BFBFBF';
const LGREY = 'F2F2F2';
const MUTED = '595959';
const FONT  = '맑은 고딕';
const P_W = 11906;   // A4 세로 폭
const P_H = 16838;   // A4 세로 높이

const noB = {
  top:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, bottom:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
  left:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, right:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
  insideHorizontal:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, insideVertical:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
};
const gB = {
  top:{style:BorderStyle.SINGLE,size:4,color:GREY}, bottom:{style:BorderStyle.SINGLE,size:4,color:GREY},
  left:{style:BorderStyle.SINGLE,size:4,color:GREY}, right:{style:BorderStyle.SINGLE,size:4,color:GREY},
  insideHorizontal:{style:BorderStyle.SINGLE,size:4,color:GREY}, insideVertical:{style:BorderStyle.SINGLE,size:4,color:GREY},
};

module.exports = function sb({ landscape = false, marginMm = 20, baseSize = 20 } = {}) {
  const PAGE_W = landscape ? P_H : P_W;   // 렌더 기준 폭
  const PAGE_H = landscape ? P_W : P_H;   // 렌더 기준 높이
  const M = convertMillimetersToTwip(marginMm);
  const BODY_W = PAGE_W - M * 2;
  const logo = n => fs.readFileSync(path.join(__dirname, n));

  const txt = (t, o = {}) => new TextRun({
    text: t, font: FONT, size: o.size ?? baseSize,
    bold: o.bold ?? false, color: o.color ?? DARK,
  });

  function cellP(t, o = {}) {
    const ls = String(t).split('\n');
    return ls.map((l, i) => new Paragraph({
      alignment: o.align ?? AlignmentType.LEFT,
      spacing: { before: i === 0 ? 30 : 0, after: i === ls.length - 1 ? 30 : 0, line: 250 },
      children: [txt(l, o)],
    }));
  }

  function cell(t, { w, fill, bold, color, align, valign, size, colSpan, rowSpan } = {}) {
    return new TableCell({
      width: { size: w, type: WidthType.DXA },
      shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
      verticalAlign: valign ?? VerticalAlign.CENTER,
      columnSpan: colSpan, rowSpan,
      margins: { top: 50, bottom: 50, left: 90, right: 90 },
      children: cellP(t, { bold, color, align, size }),
    });
  }

  /**
   * table(head, rows, widths, opts)
   *  opts.center     중앙정렬 열 인덱스 (기본 [0])
   *  opts.boldCols   굵게 표시할 열
   *  opts.fillCols   회색 배경(기입란) 열
   *  opts.h          행 최소 높이
   *  opts.sizes      열별 글자 크기
   */
  function table(head, rows, widths, opts = {}) {
    const centerCols = opts.center ?? [0];
    const boldCols = opts.boldCols ?? [0];
    return new Table({
      width: { size: BODY_W, type: WidthType.DXA }, columnWidths: widths, borders: gB,
      rows: [
        ...(head ? [new TableRow({
          tableHeader: true,
          children: head.map((h, i) => cell(h, { w: widths[i], fill: GREEN, bold: true, color: WHITE, align: AlignmentType.CENTER })),
        })] : []),
        ...rows.map(r => new TableRow({
          height: opts.h ? { value: opts.h, rule: 'atLeast' } : undefined,
          children: r.map((c, i) => cell(c, {
            w: widths[i],
            align: centerCols.includes(i) ? AlignmentType.CENTER : AlignmentType.LEFT,
            bold: boldCols.includes(i),
            size: opts.sizes ? opts.sizes[i] : undefined,
            fill: (opts.fillCols || []).includes(i) ? LGREY : undefined,
          })),
        })),
      ],
    });
  }

  /** 라벨(녹색) + 값 2열 정보표 */
  function infoTable(rows, labelW = 2600, { fillValue = false } = {}) {
    const vW = BODY_W - labelW;
    return new Table({
      width: { size: BODY_W, type: WidthType.DXA }, columnWidths: [labelW, vW], borders: gB,
      rows: rows.map(([k, v]) => new TableRow({
        children: [cell(k, { w: labelW, fill: GREEN, bold: true, color: WHITE }),
                   cell(v, { w: vW, fill: fillValue ? LGREY : undefined })],
      })),
    });
  }

  /** 점검·검사표: 마지막 2열을 기입란(결과/판정)으로 고정 */
  function checkTable(head, rows, widths, opts = {}) {
    const nFill = opts.fillFrom ?? head.length - 2;
    return table(head, rows, widths, {
      center: opts.center ?? [0, head.length - 2, head.length - 1],
      boldCols: opts.boldCols ?? [0, 1],
      fillCols: Array.from({ length: head.length - nFill }, (_, i) => nFill + i),
      h: opts.h ?? 560,
      sizes: opts.sizes,
    });
  }

  const h1 = t => new Paragraph({
    heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 180 },
    shading: { type: ShadingType.CLEAR, fill: GREEN, color: 'auto' },
    children: [new TextRun({ text: '  ' + t, font: FONT, size: landscape ? 25 : 26, bold: true, color: WHITE })],
  });
  const h2 = t => new Paragraph({
    heading: HeadingLevel.HEADING_2, spacing: { before: 260, after: 130 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GREEN } },
    children: [new TextRun({ text: t, font: FONT, size: landscape ? 22 : 23, bold: true, color: GREEN })],
  });
  const body = (t, o = {}) => new Paragraph({ spacing: { before: 40, after: 80, line: 295 }, children: [txt(t, o)] });
  const bullet = t => new Paragraph({ numbering: { reference: 'b1', level: 0 }, spacing: { before: 20, after: 55, line: 295 }, children: [txt(t)] });
  const numbered = ref => t => new Paragraph({ numbering: { reference: ref, level: 0 }, spacing: { before: 20, after: 55, line: 295 }, children: [txt(t)] });
  const spacer = (h = 120) => new Paragraph({ spacing: { after: h }, children: [] });
  const cap = t => new Paragraph({ spacing: { before: 50, after: 90 }, children: [txt(t, { size: 17, color: MUTED })] });
  const pageBreak = () => new Paragraph({ children: [new PageBreak()] });
  const note = t => new Paragraph({ spacing: { before: 80, after: 100, line: 295 }, children: [txt(t, { bold: true, color: GREEN })] });
  const blanks = (n, cols) => Array.from({ length: n }, (_, i) => cols.map(c => typeof c === 'function' ? c(i) : c));

  /** 표지 — Primary Green 전면 + 흰색 로고 */
  function cover({ title, subEn, equip, equipSub, lines = [] }) {
    const kids = [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: landscape ? 520 : 640 }, children: [
        new ImageRun({ type: 'png', data: logo('logo_white.png'), transformation: { width: landscape ? 190 : 200, height: landscape ? 100 : 105 } })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 110 }, children: [
        new TextRun({ text: title, font: FONT, size: landscape ? 58 : 60, bold: true, color: WHITE })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: landscape ? 400 : 480 }, children: [
        new TextRun({ text: subEn, font: FONT, size: 26, color: WHITE })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 90 }, children: [
        new TextRun({ text: equip, font: FONT, size: 32, bold: true, color: WHITE })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: landscape ? 600 : 780 }, children: [
        new TextRun({ text: equipSub, font: FONT, size: 24, color: WHITE })] }),
      ...lines.map(l => new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 50 }, children: [
        new TextRun({ text: l, font: FONT, size: 21, color: WHITE })] })),
    ];
    return new Table({
      width: { size: PAGE_W, type: WidthType.DXA }, columnWidths: [PAGE_W], borders: noB,
      rows: [new TableRow({ height: { value: PAGE_H, rule: 'exact' }, children: [new TableCell({
        width: { size: PAGE_W, type: WidthType.DXA },
        shading: { type: ShadingType.CLEAR, fill: GREEN, color: 'auto' },
        verticalAlign: VerticalAlign.CENTER, margins: { top: 0, bottom: 0, left: 900, right: 900 },
        children: kids })] })],
    });
  }

  /** 문서 생성 및 저장 */
  function render({ file, title, description, headerText, cover: cv, children }) {
    const pageSpec = { width: P_W, height: P_H, orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT };
    const hdr = new Header({ children: [new Paragraph({
      alignment: AlignmentType.RIGHT, spacing: { after: 50 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GREEN } },
      children: [
        new TextRun({ text: headerText + '        ', font: FONT, size: 16, color: MUTED }),
        new ImageRun({ type: 'png', data: logo('logo_green.png'), transformation: { width: 70, height: 37 } })] })] });
    const ftr = new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: ['- ', PageNumber.CURRENT, ' -'], font: FONT, size: 16, color: MUTED })] })] });
    const lvl = fmt => ({ level: 0, format: fmt, text: fmt === LevelFormat.BULLET ? '•' : '%1.',
      alignment: fmt === LevelFormat.BULLET ? AlignmentType.LEFT : AlignmentType.START,
      style: { paragraph: { indent: { left: 420, hanging: 240 } }, run: { color: GREEN, bold: fmt !== LevelFormat.BULLET, font: FONT } } });

    const doc = new Document({
      creator: 'Sweet Balance 공무팀', title, description,
      styles: { default: { document: { run: { font: FONT, size: baseSize, color: DARK } } } },
      numbering: { config: [
        { reference: 'b1', levels: [lvl(LevelFormat.BULLET)] },
        { reference: 'n1', levels: [lvl(LevelFormat.DECIMAL)] },
        { reference: 'n2', levels: [lvl(LevelFormat.DECIMAL)] },
        { reference: 'n3', levels: [lvl(LevelFormat.DECIMAL)] },
      ] },
      sections: [
        { properties: { page: { size: pageSpec, margin: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0 } } },
          children: [cover(cv)] },
        { properties: { page: { size: pageSpec, margin: { top: M, right: M, bottom: M, left: M,
            header: convertMillimetersToTwip(8), footer: convertMillimetersToTwip(8) } } },
          headers: { default: hdr }, footers: { default: ftr }, children },
      ],
    });
    return Packer.toBuffer(doc).then(b => { fs.writeFileSync(file, b); console.log('written', file, b.length); });
  }

  return { GREEN, WHITE, DARK, LGREY, MUTED, FONT, BODY_W, AlignmentType, VerticalAlign,
    txt, cell, table, infoTable, checkTable, h1, h2, body, bullet, spacer, cap, note,
    pageBreak, blanks, cover, render,
    n1: numbered('n1'), n2: numbered('n2'), n3: numbered('n3') };
};
