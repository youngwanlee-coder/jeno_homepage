const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, HeadingLevel, PageBreak, PageOrientation, VerticalAlign,
  BorderStyle, ImageRun, Header, Footer, PageNumber, LevelFormat, convertMillimetersToTwip,
} = require('docx');
const D = require('./data.js');

const GREEN = '007C5E';
const WHITE = 'FFFFFF';
const DARK  = '212121';
const GREY  = 'BFBFBF';
const LGREY = 'F2F2F2';
const FONT  = '맑은 고딕';

const PAGE_W = 11906, PAGE_H = 16838;
const MARGIN = convertMillimetersToTwip(20);
const BODY_W = PAGE_W - MARGIN * 2; // 9638

const noBorders = {
  top:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, bottom:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
  left:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, right:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
  insideHorizontal:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}, insideVertical:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},
};
const gridBorders = {
  top:{style:BorderStyle.SINGLE,size:4,color:GREY}, bottom:{style:BorderStyle.SINGLE,size:4,color:GREY},
  left:{style:BorderStyle.SINGLE,size:4,color:GREY}, right:{style:BorderStyle.SINGLE,size:4,color:GREY},
  insideHorizontal:{style:BorderStyle.SINGLE,size:4,color:GREY}, insideVertical:{style:BorderStyle.SINGLE,size:4,color:GREY},
};

const txt = (text, o={}) => new TextRun({
  text, font: FONT, size: o.size ?? 20, bold: o.bold ?? false,
  color: o.color ?? DARK,
});

function cellP(text, o={}) {
  const lines = String(text).split('\n');
  return lines.map((l, i) => new Paragraph({
    alignment: o.align ?? AlignmentType.LEFT,
    spacing: { before: i === 0 ? 40 : 0, after: i === lines.length - 1 ? 40 : 0, line: 260 },
    children: [txt(l, o)],
  }));
}

function cell(text, { w, fill, bold, color, align, valign } = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    shading: fill ? { type: ShadingType.CLEAR, fill, color: 'auto' } : undefined,
    verticalAlign: valign ?? VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: cellP(text, { bold, color, align }),
  });
}

// 2열 정보 테이블 (라벨 녹색)
function infoTable(rows, labelW = 2400) {
  const valW = BODY_W - labelW;
  return new Table({
    width: { size: BODY_W, type: WidthType.DXA },
    columnWidths: [labelW, valW],
    borders: gridBorders,
    rows: rows.map(([k, v]) => new TableRow({
      children: [
        cell(k, { w: labelW, fill: GREEN, bold: true, color: WHITE }),
        cell(v, { w: valW }),
      ],
    })),
  });
}

// 헤더 있는 일반 테이블
function headTable(head, rows, widths) {
  return new Table({
    width: { size: BODY_W, type: WidthType.DXA },
    columnWidths: widths,
    borders: gridBorders,
    rows: [
      new TableRow({
        tableHeader: true,
        children: head.map((h, i) => cell(h, { w: widths[i], fill: GREEN, bold: true, color: WHITE, align: AlignmentType.CENTER })),
      }),
      ...rows.map(r => new TableRow({
        children: r.map((c, i) => cell(c, {
          w: widths[i],
          align: i === 0 ? AlignmentType.CENTER : AlignmentType.LEFT,
          bold: i === 0,
        })),
      })),
    ],
  });
}

// 요구사항 테이블: ID / 구분 / 요구사항 / 검증
const REQ_W = [1000, 800, 5838, 2000];
function reqTable(rows) {
  return new Table({
    width: { size: BODY_W, type: WidthType.DXA },
    columnWidths: REQ_W,
    borders: gridBorders,
    rows: [
      new TableRow({
        tableHeader: true,
        children: ['요구사항 ID', '구분', '요구사항 내용', '검증방법'].map((h, i) =>
          cell(h, { w: REQ_W[i], fill: GREEN, bold: true, color: WHITE, align: AlignmentType.CENTER })),
      }),
      ...rows.map(r => new TableRow({
        children: [
          cell(r[0], { w: REQ_W[0], bold: true, align: AlignmentType.CENTER }),
          cell(r[1], { w: REQ_W[1], bold: true, align: AlignmentType.CENTER,
            fill: r[1] === 'M' ? undefined : LGREY }),
          cell(r[2], { w: REQ_W[2] }),
          cell(r[3], { w: REQ_W[3], align: AlignmentType.CENTER }),
        ],
      })),
    ],
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    shading: { type: ShadingType.CLEAR, fill: GREEN, color: 'auto' },
    children: [new TextRun({ text: '  ' + text, font: FONT, size: 26, bold: true, color: WHITE })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GREEN } },
    children: [new TextRun({ text, font: FONT, size: 23, bold: true, color: GREEN })],
  });
}
const body = (text, o={}) => new Paragraph({
  spacing: { before: 40, after: 80, line: 300 },
  children: [txt(text, o)],
});
const bullet = (text) => new Paragraph({
  numbering: { reference: 'sb-bullet', level: 0 },
  spacing: { before: 20, after: 60, line: 300 },
  children: [txt(text)],
});
const numRef = (ref) => (text) => new Paragraph({
  numbering: { reference: ref, level: 0 },
  spacing: { before: 20, after: 60, line: 300 },
  children: [txt(text)],
});
const num = numRef('sb-num');
const num2 = numRef('sb-num2');
const spacer = (h = 120) => new Paragraph({ spacing: { after: h }, children: [] });
const caption = (text) => new Paragraph({
  spacing: { before: 60, after: 100 },
  children: [txt(text, { size: 18, color: '595959' })],
});

/* ---------- 표지 ---------- */
const coverCell = new TableCell({
  width: { size: PAGE_W, type: WidthType.DXA },
  shading: { type: ShadingType.CLEAR, fill: GREEN, color: 'auto' },
  verticalAlign: VerticalAlign.CENTER,
  margins: { top: 0, bottom: 0, left: 900, right: 900 },
  children: [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 700 }, children: [
      new ImageRun({ type: 'png', data: fs.readFileSync('logo_white.png'),
        transformation: { width: 200, height: 105 } }),
    ]}),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [
      new TextRun({ text: '사용자 요구사항 명세서', font: FONT, size: 64, bold: true, color: WHITE }),
    ]}),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 500 }, children: [
      new TextRun({ text: 'User Requirement Specification (URS)', font: FONT, size: 28, color: WHITE }),
    ]}),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 }, children: [
      new TextRun({ text: '자동용기포장기 (컵실링기)', font: FONT, size: 34, bold: true, color: WHITE }),
    ]}),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 900 }, children: [
      new TextRun({ text: '모델 HW-103RT300  ·  1식(SET)  ·  컵과일 포장라인', font: FONT, size: 26, color: WHITE }),
    ]}),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [
      new TextRun({ text: '문서번호  URS-PKG-2026-001   |   Rev.0', font: FONT, size: 22, color: WHITE }),
    ]}),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [
      new TextRun({ text: '2026. 09. 09.', font: FONT, size: 22, color: WHITE }),
    ]}),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun({ text: '공무팀 (생산지원)', font: FONT, size: 22, color: WHITE }),
    ]}),
  ],
});
const coverTable = new Table({
  width: { size: PAGE_W, type: WidthType.DXA },
  columnWidths: [PAGE_W],
  borders: noBorders,
  rows: [new TableRow({ height: { value: PAGE_H, rule: 'exact' }, children: [coverCell] })],
});

/* ---------- 본문 ---------- */
const bodyChildren = [];
const P = bodyChildren.push.bind(bodyChildren);

P(h1('문서 정보'));
P(infoTable(D.meta, 2400));
P(spacer(240));
P(h2('개정 이력'));
P(headTable(D.revision.head, D.revision.rows, [700, 1400, 5238, 1150, 1150]));

P(h1('1.  개요'));
P(h2('1.1  목적'));
P(body('컵과일 포장라인의 실링·컷팅·배출 공정 자동화를 위해 신규 구매하는 자동용기포장기(컵실링기, 모델 HW-103RT300)에 대하여 당사가 요구하는 성능·기능·위생·안전·검증 요건을 명확히 규정한다. 본 문서는 공급자 제작·납품의 기준이 되며, 설비 인수(검수) 판정의 근거 문서로 사용한다.'));
P(h2('1.2  적용 범위'));
P(bullet('대상 설비: 자동용기포장기(컵실링기) HW-103RT300 본체 1식'));
P(bullet('포함 옵션: 석션 배출장치(4구) + 진공펌프, 1.5M 정렬컨베이어(4구→1구) + 스토퍼, 아이마크 장치'));
P(bullet('적용 공정: 컵과일 충전 완료품의 리드지 실링 → 트리밍 → 석션 배출 → 1열 정렬 반출'));
P(bullet('제외 범위: 컵 충전기, 라벨러, 금속검출기 등 전·후공정 설비 및 유틸리티 1차 공급(전기·Air·콤프레샤)은 당사 부담'));
P(h2('1.3  요구사항 표기 규칙'));
P(infoTable([
  ['구분 M', '필수(Mandatory) — 미충족 시 인수 불가. 시정 완료 후 재검증 대상'],
  ['구분 D', '권장(Desirable) — 미충족 시 대안 협의 가능. 협의 결과는 문서로 확정'],
  ['검증 FAT', '공장검사(Factory Acceptance Test) — 출하 전 공급자 공장에서 확인'],
  ['검증 IQ', '설치확인(Installation Qualification) — 설치 완료 상태 확인'],
  ['검증 OQ', '운전확인(Operational Qualification) — 시운전 중 기능·안전 확인'],
  ['검증 PQ', '성능확인(Performance Qualification) — 실제 생산조건 확인'],
  ['검증 DOC', '문서 확인 — 제출 서류·성적서·도면으로 확인'],
], 1800));
P(h2('1.4  관련 법규 및 기준'));
D.laws.forEach(l => P(bullet(l)));

P(h1('2.  설비 개요'));
P(h2('2.1  도입 배경'));
P(bullet('현재 컵과일 포장라인의 실링·컷팅·배출 공정이 수동 의존으로 라인 병목 및 인력 투입 과다 발생'));
P(bullet('연속 생산이 가능한 로터리 4구 방식 도입으로 대량 포장 대응 및 라인 병목 해소'));
P(bullet('실링·컷팅·배출 자동화로 생산성 향상 및 인력 투입 최소화'));
P(bullet('아이마크·석션 배출장치 등 옵션 무상/포함 구성으로 추가 설치비 부담 최소화'));
P(h2('2.2  확정 사양 요약 (공급자 제시)'));
P(headTable(D.specSummary.head, D.specSummary.rows, [2400, 7238]));
P(caption('※ 본 표는 견적서·공급계약서 상 제시 사양이며, 4장 요구사항이 상세 기준으로 우선한다.'));
P(h2('2.3  공정 흐름'));
P(body('컵 충전 완료품 투입  →  용기 정위치 정렬  →  리드지 공급(아이마크 정위치 검출)  →  가열 실링(4구 동시)  →  트리밍(1~2열)  →  스크랩 회수  →  석션 배출(4구)  →  정렬컨베이어(4구→1구) 반출  →  후공정', { bold: true }));

P(new Paragraph({ children: [new PageBreak()] }));

P(h1('3.  대상 제품 및 자재 규격'));
P(headTable(D.product.head, D.product.rows, [1800, 4838, 3000]));
P(caption('※ 리드지 외형·인쇄 규격은 첨부 제품 도면(리드지 도면) 기준. 이형 리드지·인쇄 로고가 있어 아이마크 정위치 실링이 필수 요건이다.'));
P(spacer(160));
P(body('[핵심 관리 포인트] 컵과일은 가열 살균 공정이 없는 비가열 냉장 제품이므로, 설비의 세척성(위생설계)과 실링 밀봉성이 제품 안전성에 직결된다. 또한 냉장 상태 용기의 플랜지 결로가 실링 불량의 최대 원인이 되므로 관련 대책을 필수 요구사항으로 규정한다.', { bold: true, color: GREEN }));

P(new Paragraph({ children: [new PageBreak()] }));

P(h1('4.  요구사항 명세'));
D.reqs.forEach((g, i) => {
  P(h2(g.title));
  P(reqTable(g.rows));
  if (i < D.reqs.length - 1) P(spacer(200));
});

P(new Paragraph({ children: [new PageBreak()] }));

P(h1('5.  검증 및 인수 계획'));
P(headTable(D.verify.head, D.verify.rows, [1300, 1400, 3738, 3200]));
P(spacer(200));
P(h2('5.1  판정 및 시정 절차'));
P(num('당사 검증 담당(공무·생산·품질)이 각 단계별 체크시트로 확인 후 합격/부적합 판정'));
P(num('부적합 항목은 부적합 리스트(Punch List)에 기록하고 공급자에게 서면 통보'));
P(num('공급자 시정 완료 후 해당 항목 재검증, 재검증 합격 시 다음 단계 진행'));
P(num('필수(M) 항목 전 항목 합격 및 문서 제출 완료 시 검수조서 서명 → 최종 인수'));
P(h2('5.2  검증 참여 및 책임'));
P(infoTable([
  ['공무팀', '검증 총괄, 유틸리티·전기·안전 항목 확인, IQ/OQ 주관'],
  ['생산팀', '조작성·생산성 확인, PQ 주관, 작업자 교육 수료'],
  ['품질팀', '재질·위생·밀봉성 확인, 관련 성적서 검토, HACCP 문서 반영'],
  ['공급자', 'FAT 실시 및 성적서 제출, 설치·시운전, 교육, 부적합 시정'],
], 1800));

P(new Paragraph({ children: [new PageBreak()] }));

P(h1('6.  계약 · 납기 조건'));
P(headTable(D.contract.head, D.contract.rows, [2000, 7638]));
P(caption('※ 견적서 상 일반 공급가 44,200,000원 대비 우대기업 적용가 40,000,000원으로 계약 체결 (비용 절감 4,200,000원).'));
P(caption('※ 입금계좌: IBK기업은행 504-019013-01-011 (예금주: 하이퍼박(주))'));

P(h1('7.  확인 필요 사항 (Open Issues)'));
P(body('아래 항목은 계약 사양에 확정 반영되지 않았거나 상호 확인이 필요한 사항으로, 제작 착수 전 서면 확정한다.'));
D.open.forEach(o => P(num2(o)));

P(new Paragraph({ children: [new PageBreak()] }));

P(h1('8.  승인'));
P(body('본 사용자 요구사항 명세서(URS)의 내용을 검토·승인하며, 공급자는 본 문서를 제작·납품의 기준으로 준수한다.'));
P(spacer(200));
P(new Table({
  width: { size: BODY_W, type: WidthType.DXA },
  columnWidths: [1900, 1900, 1900, 1900, 2038],
  borders: gridBorders,
  rows: [
    new TableRow({
      tableHeader: true,
      children: ['구분', '작성', '검토', '검토', '승인'].map((h, i) =>
        cell(h, { w: [1900,1900,1900,1900,2038][i], fill: GREEN, bold: true, color: WHITE, align: AlignmentType.CENTER })),
    }),
    new TableRow({
      children: ['부서', '공무팀', '생산팀', '품질팀', '공장장 / 대표'].map((h, i) =>
        cell(h, { w: [1900,1900,1900,1900,2038][i], align: AlignmentType.CENTER })),
    }),
    new TableRow({
      height: { value: 1400, rule: 'atLeast' },
      children: ['성명 / 서명', '', '', '', ''].map((h, i) =>
        cell(h, { w: [1900,1900,1900,1900,2038][i], align: AlignmentType.CENTER })),
    }),
    new TableRow({
      children: ['일자', '', '', '', ''].map((h, i) =>
        cell(h, { w: [1900,1900,1900,1900,2038][i], align: AlignmentType.CENTER })),
    }),
  ],
}));
P(spacer(300));
P(h2('첨부 서류'));
P(bullet('견적서 1부 (하이퍼박(주), 2026.7.14.)'));
P(bullet('공급계약서 1부 (2026.7.15. 체결)'));
P(bullet('제품 도면 1부 (본체 도면 및 리드지 도면)'));

/* ---------- 문서 조립 ---------- */
const bodyHeader = new Header({
  children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { after: 60 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GREEN } },
    children: [
      new TextRun({ text: 'URS-PKG-2026-001  |  자동용기포장기(컵실링기) HW-103RT300        ', font: FONT, size: 16, color: '595959' }),
      new ImageRun({ type: 'png', data: fs.readFileSync('logo_green.png'),
        transformation: { width: 76, height: 40 } }),
    ],
  })],
});
const bodyFooter = new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ children: ['- ', PageNumber.CURRENT, ' -'], font: FONT, size: 16, color: '595959' })],
  })],
});

const doc = new Document({
  creator: 'Sweet Balance 공무팀',
  title: '자동용기포장기(컵실링기) HW-103RT300 사용자 요구사항 명세서(URS)',
  description: 'URS-PKG-2026-001 Rev.0',
  styles: {
    default: {
      document: { run: { font: FONT, size: 20, color: DARK } },
    },
  },
  numbering: {
    config: [
      { reference: 'sb-bullet', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 400, hanging: 220 } }, run: { color: GREEN, font: FONT } } }] },
      { reference: 'sb-num', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
        alignment: AlignmentType.START,
        style: { paragraph: { indent: { left: 420, hanging: 240 } }, run: { color: GREEN, bold: true, font: FONT } } }] },
      { reference: 'sb-num2', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
        alignment: AlignmentType.START,
        style: { paragraph: { indent: { left: 420, hanging: 240 } }, run: { color: GREEN, bold: true, font: FONT } } }] },
    ],
  },
  sections: [
    {
      properties: {
        page: { size: { width: PAGE_W, height: PAGE_H, orientation: PageOrientation.PORTRAIT },
          margin: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0 } },
      },
      children: [coverTable],
    },
    {
      properties: {
        page: { size: { width: PAGE_W, height: PAGE_H, orientation: PageOrientation.PORTRAIT },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN,
            header: convertMillimetersToTwip(10), footer: convertMillimetersToTwip(10) } },
      },
      headers: { default: bodyHeader },
      footers: { default: bodyFooter },
      children: bodyChildren,
    },
  ],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync('URS_자동용기포장기_컵실링기_HW-103RT300_Rev0.docx', b);
  console.log('written', b.length);
});
