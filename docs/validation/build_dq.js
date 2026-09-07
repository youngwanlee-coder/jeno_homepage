const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, BorderStyle, HeadingLevel, PageBreak, ImageRun,
  VerticalAlign, Header, Footer, PageNumber, LevelFormat, convertMillimetersToTwip,
} = require('docx');

const ASSETS = '/root/.claude/skills/synced/97ee583e-2ae6-41fa-8e18-9c1031af2ac1_f2d3a598-1ea9-4785-ba2d-4051e2cdd41d/sweet-balance-brand/assets';
const GREEN = '007C5E';
const WHITE = 'FFFFFF';
const DARK = '212121';
const LIGHT = 'F2F7F5';
const FONT = '맑은 고딕';
const TW = 9700;

const NB = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: NB, bottom: NB, left: NB, right: NB, insideHorizontal: NB, insideVertical: NB };

function txt(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: opts.size || 20, bold: !!opts.bold, color: opts.color || DARK });
}

function h1(text) {
  return new Paragraph({
    spacing: { before: 320, after: 140 },
    shading: { type: ShadingType.CLEAR, fill: GREEN, color: 'auto' },
    children: [new TextRun({ text: '  ' + text, font: FONT, size: 26, bold: true, color: WHITE })],
    heading: HeadingLevel.HEADING_1,
  });
}

function h2(text) {
  return new Paragraph({
    spacing: { before: 220, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GREEN } },
    children: [new TextRun({ text, font: FONT, size: 23, bold: true, color: GREEN })],
    heading: HeadingLevel.HEADING_2,
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: opts.before || 0, after: opts.after === undefined ? 80 : opts.after },
    alignment: opts.align,
    children: [txt(text, opts)],
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'sb-bullet', level },
    spacing: { after: 40 },
    children: [txt(text)],
  });
}

function cell(children, opts = {}) {
  return new TableCell({
    width: { size: opts.w, type: WidthType.DXA },
    shading: opts.fill ? { type: ShadingType.CLEAR, fill: opts.fill, color: 'auto' } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    columnSpan: opts.span,
    children: children,
  });
}

function tcell(text, opts = {}) {
  return cell([new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { after: 0 },
    children: [new TextRun({ text: String(text), font: FONT, size: opts.size || 18, bold: !!opts.bold, color: opts.color || DARK })],
  })], opts);
}

function table(headers, rows, widths, opts = {}) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => tcell(h, { w: widths[i], fill: GREEN, color: WHITE, bold: true, align: AlignmentType.CENTER })),
  });
  const bodyRows = rows.map((r) => new TableRow({
    children: r.map((c, i) => tcell(c, {
      w: widths[i],
      align: (opts.center || []).includes(i) ? AlignmentType.CENTER : AlignmentType.LEFT,
    })),
  }));
  return new Table({
    columnWidths: widths,
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: GREEN },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: GREEN },
      left: { style: BorderStyle.SINGLE, size: 4, color: GREEN },
      right: { style: BorderStyle.SINGLE, size: 4, color: GREEN },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'BFD8D0' },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: 'BFD8D0' },
    },
    rows: [headerRow, ...bodyRows],
  });
}

// 점검 체크리스트: No / 점검항목 / 요구기준(URS) / 확인방법 / 판정 / 비고
const CK_W = [520, 1780, 3600, 1300, 800, 1700];
function checklist(prefix, items) {
  return table(
    ['No', '점검 항목', '요구 기준 (URS)', '확인 방법', '판정', '비고 / 실측값'],
    items.map((it, i) => [`${prefix}-${i + 1}`, it[0], it[1], it[2], '', '']),
    CK_W,
    { center: [0, 4] }
  );
}

const spacer = (h = 120) => new Paragraph({ spacing: { after: h }, children: [txt('')] });

/* ─────────────────── 표지 ─────────────────── */
const coverBar = new Table({
  columnWidths: [TW],
  width: { size: TW, type: WidthType.DXA },
  borders: noBorders,
  rows: [new TableRow({
    height: { value: convertMillimetersToTwip(285), rule: 'atLeast' },
    children: [new TableCell({
      width: { size: TW, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: GREEN, color: 'auto' },
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 400, bottom: 400, left: 400, right: 400 },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 700 },
          children: [new ImageRun({
            type: 'png',
            data: fs.readFileSync(`${ASSETS}/logo_white.png`),
            transformation: { width: 200, height: 200 * 0.28 },
          })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 160 },
          children: [new TextRun({ text: '설계적격성평가 (DQ)', font: FONT, size: 64, bold: true, color: WHITE })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: 'Design Qualification', font: FONT, size: 28, color: WHITE })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 500, after: 60 },
          children: [new TextRun({ text: '하이퍼박(주) 로타리 자동 포장기', font: FONT, size: 34, bold: true, color: WHITE })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 900 },
          children: [new TextRun({ text: 'HW-104 Series  |  MAP · 실링 포장 설비', font: FONT, size: 24, color: WHITE })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 40 },
          children: [new TextRun({ text: '문서번호 : SB-VAL-DQ-2026-001   |   Rev. 0', font: FONT, size: 22, color: WHITE })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: '작성 : 공무팀        작성일 : 2026. __. __.', font: FONT, size: 22, color: WHITE })],
        }),
      ],
    })],
  })],
});

/* ─────────────────── 본문 ─────────────────── */
const body = [];

// 문서 정보
body.push(h1('문서 정보'));
body.push(table(
  ['구분', '내용', '구분', '내용'],
  [
    ['문서번호', 'SB-VAL-DQ-2026-001', '개정번호', 'Rev. 0'],
    ['설비명', '로타리 자동 포장기', '설비관리번호', 'PK-___'],
    ['제조사', '하이퍼박(주) (HYPERVAC)', '모델명', 'HW-104 ___ (해당 모델 기재)'],
    ['설치 장소', '____ 공장 / 포장실', '용도', '용기 실링 · MAP 가스치환 포장'],
    ['평가 구분', '신규 도입 설비 DQ', '평가 일자', '2026. __. __.'],
  ],
  [1300, 3550, 1300, 3550]
));

body.push(spacer());
body.push(h2('개정 이력'));
body.push(table(
  ['Rev.', '개정일', '개정 내용', '작성', '승인'],
  [['0', '', '최초 작성', '', ''], ['', '', '', '', ''], ['', '', '', '', '']],
  [800, 1500, 4500, 1450, 1450],
  { center: [0, 1, 3, 4] }
));

body.push(spacer());
body.push(h2('검토 · 승인'));
body.push(table(
  ['구분', '작성', '검토', '검토', '승인'],
  [
    ['부서', '공무팀', '생산팀', '품질팀(QA)', '공장장'],
    ['성명', '', '', '', ''],
    ['서명 / 일자', '', '', '', ''],
  ],
  [1500, 2050, 2050, 2050, 2050],
  { center: [1, 2, 3, 4] }
));

body.push(new Paragraph({ children: [new PageBreak()] }));

// 1. 목적
body.push(h1('1. 목적 (Purpose)'));
body.push(bullet('본 문서는 하이퍼박(주) 로타리 자동 포장기(이하 "본 설비")의 도입에 앞서, 공급업체가 제시한 설계 및 사양이 당사 사용자요구사항(URS), 식품위생 관련 법규, 산업안전 기준에 적합한지 사전 검증하기 위함이다.'));
body.push(bullet('설계 단계에서 위생·안전·품질 리스크를 사전 제거하여, 설치 후 발생하는 재작업 및 추가 비용을 원천 차단한다.'));
body.push(bullet('DQ 승인 후 발주를 진행하며, 본 문서는 후속 IQ(설치)·OQ(운전)·PQ(성능) 검증의 기준 문서로 활용한다.'));

// 2. 적용범위
body.push(h1('2. 적용 범위 (Scope)'));
body.push(bullet('대상 설비 : 하이퍼박(주) 로타리 자동 포장기 본체 및 부속 설비'));
body.push(bullet('포함 범위 : 용기 공급부 → 충전부 → 가스치환(MAP)부 → 실링부 → 배출부, 제어반, 유틸리티 접속부, 안전장치, 부속 금형(몰드)'));
body.push(bullet('제외 범위 : 상류 계량기 · 하류 금속검출기 / 중량선별기 등 별도 설비 (연계 인터페이스만 검토)'));
body.push(bullet('평가 시점 : 발주 전 설계 도면 및 사양서 접수 완료 시점'));

// 3. 관련 법규
body.push(h1('3. 관련 법규 및 기준'));
body.push(table(
  ['구분', '법규 · 기준', '주요 확인 사항'],
  [
    ['식품 위생', '식품위생법 및 시행규칙 (시설기준)', '제조·가공 시설의 위생 요건 충족'],
    ['식품 위생', '식품 및 축산물 안전관리인증기준(HACCP) 고시 — 선행요건 「제조·가공 시설·설비 관리」', '세척·소독 용이성, 이물 혼입 방지 구조'],
    ['식품 위생', '식품공전 「기구 및 용기·포장의 기준·규격」', '식품 접촉 재질의 적합성 (재질증명서)'],
    ['산업 안전', '산업안전보건법 제80조 (유해·위험 기계의 방호조치)', '회전·협착 위험부 방호조치 이행'],
    ['산업 안전', '산업안전보건법 제84조 (안전인증) / KCs 표시', '해당 대상 설비 여부 및 인증서 확보'],
    ['산업 안전', '산업안전보건기준에 관한 규칙 (원동기·회전축 등 위험 방지, 정비 시 운전정지)', '가드·인터록·LOTO 적용 가능 구조'],
    ['전기', '전기설비기술기준 / 한국전기설비규정(KEC)', '접지, 누전차단기, 방수·방진 등급'],
    ['가스', '고압가스 안전관리법 (MAP용 N₂ · CO₂ 사용 시)', '용기·배관 안전, 산소결핍 질식 예방 대책'],
    ['설비', '기계설비법 (해당 시)', '유틸리티 배관·공조 연계 적합성'],
  ],
  [1100, 4300, 4300]
));

// 4. 용어
body.push(h1('4. 용어 정의'));
body.push(table(
  ['용어', '정의'],
  [
    ['DQ (Design Qualification)', '설비의 설계·사양이 사용자요구사항 및 법규에 적합한지 발주 전 검증하는 활동'],
    ['IQ / OQ / PQ', '설치적격성 / 운전적격성 / 성능적격성 평가 — DQ 이후 순차 수행'],
    ['URS (User Requirement Spec.)', '당사가 설비에 요구하는 성능·위생·안전 요구사항 명세'],
    ['FAT / SAT', '공장 출하 검사(Factory Acceptance Test) / 현장 설치 검사(Site Acceptance Test)'],
    ['MAP', '가스치환포장(Modified Atmosphere Packaging) — N₂ · CO₂ 등으로 포장 내 공기를 치환'],
    ['잔존산소 (Residual O₂)', 'MAP 포장 완료 후 포장 내부에 남은 산소 농도(%)'],
    ['CCP', 'HACCP 중요관리점 — 실링 온도·시간 등 관리 파라미터 연계 대상'],
  ],
  [2400, 7300]
));

// 5. 절차 및 담당
body.push(h1('5. 평가 절차 및 담당'));
body.push(table(
  ['순서', '절차', '주요 활동', '담당', '산출물'],
  [
    ['1', 'URS 확정', '요구 능력·위생·안전 요구사항 정의', '공무 / 생산 / 품질', 'URS'],
    ['2', '사양서 접수', '공급업체 도면 · 사양서 · 견적 접수', '공무', '도면, 사양서'],
    ['3', 'DQ 평가', '본 체크리스트에 따른 항목별 적합성 판정', '공무 주관, 생산·품질 합동', '본 DQ 문서'],
    ['4', '부적합 조치', '부적합 항목에 대한 설계 변경 요구 및 회신 확인', '공무 ↔ 공급업체', '조치 회신서'],
    ['5', 'DQ 승인', '종합 판정 후 승인, 발주 진행', '공장장', '승인 DQ'],
    ['6', '후속 검증', 'FAT → IQ → OQ → PQ 수행', '공무 / 품질', 'FAT·IQ·OQ·PQ 보고서'],
  ],
  [700, 1500, 3900, 1900, 1700],
  { center: [0] }
));

body.push(new Paragraph({ children: [new PageBreak()] }));

// 6. 설비 개요 및 주요 사양
body.push(h1('6. 설비 개요 및 주요 사양 대비'));
body.push(p('※ "URS 요구사양"은 당사 요구값, "제시 사양"은 공급업체 사양서 기재값을 기입하고 적합 여부를 판정한다.', { size: 18 }));
body.push(table(
  ['항목', 'URS 요구사양', '공급업체 제시 사양', '판정'],
  [
    ['포장 방식', '로타리 인덱스 방식 (용기 실링 + MAP)', '', ''],
    ['스테이션 구성', '용기 공급 → 충전 → 가스치환 → 실링 → 배출', '', ''],
    ['생산 능력', '____ pack/min 이상 (___ cycle/min × ___ cavity)', '', ''],
    ['대응 용기', '____ × ____ × ____ mm (재질 PP / PET)', '', ''],
    ['금형(몰드)', '____ cavity, 교체 소요시간 30분 이내', '', ''],
    ['실링 필름', '폭 ____ mm, 외경 ____ mm, 인쇄 무늬 보정 기능', '', ''],
    ['실링 온도 범위', '____ ~ ____ ℃, 설정값 ±3℃ 이내 제어', '', ''],
    ['MAP 가스', 'N₂ / CO₂, 잔존산소 ____ % 이하', '', ''],
    ['외형 치수', 'W ____ × D ____ × H ____ mm (설치 가능 범위 내)', '', ''],
    ['총 중량', '____ kg 이하 (바닥 허용하중 이내)', '', ''],
    ['전원', 'AC 380V 3Φ 60Hz, ____ kW 이하', '', ''],
    ['압축공기', '0.5 ~ 0.7 MPa, ____ NL/min', '', ''],
    ['주요 재질', '식품 접촉부 STS304 이상', '', ''],
    ['제어', 'PLC + 한글 터치 HMI, 제품별 레시피 저장', '', ''],
    ['소음', '작업자 위치 85 dB(A) 이하', '', ''],
    ['납기 / 보증', '발주 후 ____ 주 / 무상보증 ____ 개월', '', ''],
  ],
  [1600, 4100, 3000, 1000],
  { center: [3] }
));

body.push(new Paragraph({ children: [new PageBreak()] }));

// 7. 체크리스트
body.push(h1('7. DQ 점검 체크리스트'));
body.push(p('판정 : 적합(P) / 부적합(F) / 해당없음(N.A.)   —   부적합 항목은 9항 부적합 조치표에 등재하여 관리한다.', { size: 18 }));

body.push(h2('7.1 설치 조건 및 유틸리티'));
body.push(checklist('A', [
  ['설치 공간', '외형 치수 및 유지보수 공간(전후좌우 최소 800mm) 확보, 기존 라인 동선 간섭 없음', '배치도 검토'],
  ['바닥 하중 / 고정', '설비 중량이 바닥 허용하중 이내, 앵커 고정 및 레벨 조정(어저스터) 가능', '구조 검토'],
  ['반입 경로', '출입구·통로 폭 및 층고 통과 가능, 필요 시 분해 반입 방안 제시', '현장 실측'],
  ['전원 사양', 'AC 380V 3Φ 60Hz, 소비전력이 기존 분전반 여유용량 이내 (부하계산 첨부)', '전기 도면'],
  ['압축공기', '요구 압력·유량 명시, 제유·제습 필터(식품용) 적용', '사양서'],
  ['MAP 가스 공급', 'N₂ / CO₂ 순도 및 공급압력 명시, 배관 재질 STS, 감압·차단밸브 위치 적정', '배관 도면'],
  ['급배수', '세척수 급수 및 배수 트렌치 연결 방안, 설비 하부 물고임 없음', '도면 검토'],
  ['방수 등급', '세척 환경 고려 IP65 이상(제어반·모터·센서), 고압세척 가능 여부 명시', '사양서'],
  ['소음 · 진동', '작업자 위치 85 dB(A) 이하, 진동 저감 방안', '사양서'],
]));

body.push(spacer());
body.push(h2('7.2 위생 설계 (HACCP 선행요건 연계)'));
body.push(checklist('B', [
  ['식품 접촉부 재질', 'STS304 이상 (부식 우려부 STS316), 식품공전 기준 적합 — 재질증명서 제출', '재질증명서'],
  ['비접촉 구조부 재질', 'STS304 또는 동등 이상, 도장 마감 배제(박리 이물 리스크)', '사양서'],
  ['표면 조도 · 마감', '접촉면 Ra 0.8㎛ 이하 연마 마감, 오염 잔류 없는 표면', '사양서'],
  ['용접부', '연속 용접 및 그라인딩 마감, 틈새(Crevice) 및 점용접 배제', '도면 / 실물'],
  ['데드스페이스', '세척수 고임 없는 경사 설계, 세척 사각지대 최소화', '도면 검토'],
  ['분해 세척성', '접촉부 무공구(Tool-less) 분해 가능, 1회 세척 소요시간 명시', '사양서'],
  ['윤활유', '식품등급(NSF H1) 윤활유 사용, 접촉 가능부 누유 방지 구조', '사양서'],
  ['이물 관리', '풀림방지 체결, 유리·경질 플라스틱 배제, 탈락 우려 부품 리스트 제출', '부품 리스트'],
  ['방충 · 방서', '개구부 밀폐, 커버·캐노피 적용으로 외부 이물 유입 차단', '도면 검토'],
  ['컨베이어 · 벨트', '식품용 인증(FDA / EU) 벨트, 청소 용이 구조 및 분해 가능', '사양서'],
]));

body.push(new Paragraph({ children: [new PageBreak()] }));

body.push(h2('7.3 기계 · 구조 설계'));
body.push(checklist('C', [
  ['로타리 인덱스 구조', '스테이션 수 및 배치가 공정 순서(충전→치환→실링→배출)에 부합, 인덱스 정지 정밀도', '도면 검토'],
  ['생산 능력', 'URS 요구 능력 이상, 연속 가동 기준 실사용 능력(가동률 반영) 제시', '사양서 / FAT'],
  ['금형(몰드)', '당사 용기 전 규격 대응, 교체 30분 이내, 추가 금형 제작 확장성 및 단가 명시', '도면 / 견적'],
  ['용기 공급부', '용기 정렬·낙하 불량 방지 구조, 용기 미투입 시 자동 감지', '도면 검토'],
  ['필름 공급부', '필름 폭·외경 대응, 인쇄 무늬 보정(마크 센서), 필름 소진·끊김 알람', '사양서'],
  ['실링부', '몰드 평행도 및 가압력 조절, 구간별 히터 온도 개별 제어', '사양서'],
  ['배출 · 연계', '후공정(금속검출기·중량선별기) 연결 높이 및 인터페이스 정합', '배치도'],
  ['불량 배출', '미충전·미실링 제품 자동 감지 및 배출 또는 정지 기능', '사양서'],
  ['유지보수성', '소모품(히터·실링고무·센서) 접근성, 교체 공구 및 소요시간 명시', '도면 검토'],
]));

body.push(spacer());
body.push(h2('7.4 포장 품질 설계'));
body.push(checklist('D', [
  ['실링 강도', '실링부 인장강도 ____ N/15mm 이상, 전 둘레 균일', 'FAT 시험'],
  ['밀봉성 (Leak)', '진공 침수법 검사 시 누출 0%, 검사 조건 제시', 'FAT 시험'],
  ['잔존산소', 'MAP 포장 후 잔존산소 ____ % 이하 (제품 규격 기준)', 'FAT 측정'],
  ['가스 치환 재현성', '목표 가스 비율 대비 ±____ % 이내, 연속 생산 시 편차 유지', 'FAT 측정'],
  ['실링 온도 편차', '몰드 표면 온도 설정값 ±3℃ 이내, 위치별 편차 제시', '온도 측정'],
  ['유통기한 인쇄', '잉크젯 프린터 연동, 인쇄 위치 정확도 및 판독성 확보', 'FAT 확인'],
  ['외관 품질', '실링부 주름·오염·필름 밀림·용기 변형 없음', 'FAT 확인'],
]));

body.push(spacer());
body.push(h2('7.5 제어 · 계측 설계'));
body.push(checklist('E', [
  ['제어반 구성', 'PLC 제조사·모델 명시, 국내 부품 수급 및 A/S 가능', '사양서'],
  ['HMI', '한글 지원 터치스크린, 제품별 레시피(파라미터) 저장 및 호출', '사양서'],
  ['파라미터 보안', '관리자 / 작업자 권한 분리, 실링 온도·시간 임의 변경 방지 (CCP 관리 연계)', '사양서'],
  ['기록 · 데이터', '실링 온도·시간, 가스 농도 기록 저장 및 출력(USB·프린터) 가능, 저장 기간 명시', '사양서'],
  ['계측기 정확도', '온도센서 ±1℃ 이내, 탈착·교정 가능 구조 및 교정성적서 제출', '성적서'],
  ['알람 · 인터록', '온도 이상, 가스압 저하, 필름 소진, 도어 열림 시 경보 및 자동 정지', '사양서'],
  ['상위 연동', '후공정 · MES 통신 연동 가능성 (해당 시 프로토콜 명시)', '사양서'],
]));

body.push(new Paragraph({ children: [new PageBreak()] }));

body.push(h2('7.6 안전 설계 (산업안전보건법 연계)'));
body.push(checklist('F', [
  ['안전인증', '산업안전보건법 상 방호조치 이행 및 해당 시 KCs 인증서(또는 CE 동등 자료) 제출', '인증서'],
  ['비상정지', '조작부 및 양측에 비상정지 버튼, 작동 시 전 구동부 즉시 정지 및 자동 복귀 금지', '도면 / 실물'],
  ['안전가드', '회전·구동부 고정식 가드 설치, 공구 없이 개방 불가', '도면 검토'],
  ['인터록', '점검 도어·커버 개방 시 인터록 스위치로 즉시 정지, 우회(Bypass) 불가 구조', '도면 검토'],
  ['협착 방지', '실링 몰드 상하 작동부 손 삽입 불가 구조 또는 2중 안전장치(양수조작·광전 센서)', '도면 검토'],
  ['고온부 방호', '히터·실링몰드 접촉 위험부 단열 커버 및 「고온주의」 경고표지 부착', '도면 / 실물'],
  ['전기 안전', '접지 시공(KEC 기준), 누전차단기, 제어반 방수·방유, 배선 보호관 처리', '전기 도면'],
  ['가스 질식 예방', 'N₂ 사용 시 누출 감지 및 실내 산소농도 저하 대책(환기·산소농도계) 검토', '리스크 검토'],
  ['LOTO', '정비 시 전원·공기압 차단 및 잠금(Lock Out / Tag Out) 적용 가능 구조', '도면 검토'],
  ['과열 방지', '히터부 온도퓨즈·서모스탯 등 2중 과열 방지 장치', '사양서'],
  ['안전 표지', '한글 안전·조작 표지 부착, 한글 취급설명서 제공', '실물 / 문서'],
]));

body.push(spacer());
body.push(h2('7.7 문서 · 공급 조건'));
body.push(checklist('G', [
  ['도면', '외형도, 유틸리티 접속도, 전기 결선도, 공기압 회로도 제출', '도면'],
  ['매뉴얼', '한글 취급설명서 · 정비 설명서 · 부품도(Parts List) 제출', '문서'],
  ['증명서', '식품 접촉 재질증명서, 계측기 교정성적서, 안전인증서 제출', '문서'],
  ['FAT / SAT', '출하 전 FAT 및 설치 후 SAT 실시 계획 제시, 당사 입회 및 판정 기준 합의', '계획서'],
  ['예비품', '소모품(히터·실링고무·센서·벨트) 리스트, 단가, 국내 재고 보유 여부', '리스트'],
  ['교육', '운전자 · 정비자 교육 계획(시간·인원) 및 교육 자료 제공', '계획서'],
  ['A/S 조건', '무상 보증기간 ____ 개월, 장애 대응 기준(유선 4시간 내 / 방문 24시간 내)', '계약 조건'],
  ['납기 · 일정', '발주 후 제작 ____ 주, 설치·시운전 일정 및 생산 중단 최소화 방안', '일정표'],
]));

body.push(new Paragraph({ children: [new PageBreak()] }));

// 8. 판정 기준
body.push(h1('8. 판정 기준'));
body.push(table(
  ['판정', '기호', '기준'],
  [
    ['적합', 'P', '요구 기준을 충족하며 근거 자료(도면·사양서·증명서)로 확인됨'],
    ['부적합', 'F', '요구 기준 미충족 또는 근거 자료 미제출 — 설계 변경 또는 보완 필요'],
    ['해당없음', 'N.A.', '본 설비 구성상 해당되지 않는 항목 (사유를 비고란에 기재)'],
  ],
  [1200, 900, 7600],
  { center: [0, 1] }
));
body.push(spacer());
body.push(bullet('전 항목이 「적합」 또는 「해당없음」인 경우 DQ 승인 후 발주를 진행한다.'));
body.push(bullet('「부적합」 항목이 있는 경우, 조치 완료 및 재확인 후 승인한다. 단, 안전·위생 관련 부적합은 조치 완료 전 발주를 보류한다.'));

// 9. 부적합 조치
body.push(h1('9. 부적합 사항 및 조치 (Deviation)'));
body.push(table(
  ['No', '항목 No', '부적합 내용', '조치 요구 사항', '조치 기한', '확인'],
  [['1', '', '', '', '', ''], ['2', '', '', '', '', ''], ['3', '', '', '', '', ''], ['4', '', '', '', '', ''], ['5', '', '', '', '', '']],
  [520, 900, 3200, 3200, 1080, 800],
  { center: [0, 1, 4, 5] }
));

// 10. 종합 판정
body.push(h1('10. 종합 판정'));
body.push(table(
  ['구분', '점검 항목 수', '적합', '부적합', '해당없음'],
  [
    ['A. 설치 조건 및 유틸리티', '9', '', '', ''],
    ['B. 위생 설계', '10', '', '', ''],
    ['C. 기계 · 구조 설계', '9', '', '', ''],
    ['D. 포장 품질 설계', '7', '', '', ''],
    ['E. 제어 · 계측 설계', '7', '', '', ''],
    ['F. 안전 설계', '11', '', '', ''],
    ['G. 문서 · 공급 조건', '8', '', '', ''],
    ['합    계', '61', '', '', ''],
  ],
  [3300, 1600, 1600, 1600, 1600],
  { center: [1, 2, 3, 4] }
));
body.push(spacer());
body.push(h2('종합 결론'));
body.push(table(
  ['판정', '내용'],
  [
    ['□ 승인 (Approved)', '전 항목 적합 — 설계 적격, 발주 진행'],
    ['□ 조건부 승인', '경미한 부적합 조치를 조건으로 발주 진행 (조치 사항 : ____________________ )'],
    ['□ 부적합 (Rejected)', '설계 변경 후 재평가 필요'],
  ],
  [2400, 7300]
));
body.push(spacer(200));
body.push(p('종합 의견', { bold: true }));
body.push(table(
  ['의견'],
  [[''], [''], ['']],
  [9700]
));

// 11. 첨부
body.push(h1('11. 첨부 자료'));
body.push(table(
  ['No', '첨부 자료', '첨부 여부'],
  [
    ['1', '사용자요구사항서 (URS)', '□'],
    ['2', '공급업체 사양서 및 견적서', '□'],
    ['3', '설비 외형도 · 배치도 · 유틸리티 접속도', '□'],
    ['4', '전기 결선도 · 공기압 회로도', '□'],
    ['5', '식품 접촉부 재질증명서', '□'],
    ['6', '안전인증서 (KCs 등)', '□'],
    ['7', '계측기 교정성적서', '□'],
    ['8', 'FAT / SAT 계획서', '□'],
    ['9', '예비품 리스트 및 A/S 조건', '□'],
  ],
  [700, 7700, 1300],
  { center: [0, 2] }
));

/* ─────────────────── 문서 조립 ─────────────────── */
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: FONT, size: 20, color: DARK }, paragraph: { spacing: { line: 288 } } },
    },
  },
  numbering: {
    config: [{
      reference: 'sb-bullet',
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: '▪', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 340, hanging: 200 } } } },
        { level: 1, format: LevelFormat.BULLET, text: '–', alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 700, hanging: 200 } } } },
      ],
    }],
  },
  sections: [
    {
      properties: {
        page: { margin: { top: 0, right: 1080, bottom: 0, left: 1080, header: 0, footer: 0 } },
      },
      children: [coverBar],
    },
    {
      properties: {
        page: { margin: { top: 1200, right: 1080, bottom: 1080, left: 1080, header: 560, footer: 460 } },
      },
      headers: {
        default: new Header({
          children: [new Table({
            columnWidths: [6200, 3500],
            width: { size: 9700, type: WidthType.DXA },
            borders: { ...noBorders, bottom: { style: BorderStyle.SINGLE, size: 6, color: GREEN } },
            rows: [new TableRow({
              children: [
                cell([new Paragraph({ spacing: { after: 60 }, children: [txt('설계적격성평가(DQ) — 하이퍼박 로타리 자동 포장기', { size: 17, color: GREEN, bold: true })] })], { w: 6200 }),
                cell([new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  spacing: { after: 60 },
                  children: [new ImageRun({
                    type: 'png',
                    data: fs.readFileSync(`${ASSETS}/logo_green.png`),
                    transformation: { width: 108, height: 108 * 0.28 },
                  })],
                })], { w: 3500 }),
              ],
            })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              txt('SB-VAL-DQ-2026-001  Rev.0        ', { size: 16, color: '707070' }),
              new TextRun({ children: [PageNumber.CURRENT, ' / ', PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: '707070' }),
            ],
          })],
        }),
      },
      children: body,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || 'DQ.docx', buf);
  console.log('written', (process.argv[2] || 'DQ.docx'), buf.length);
});
