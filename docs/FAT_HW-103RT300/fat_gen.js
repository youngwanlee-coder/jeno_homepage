const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  ShadingType, AlignmentType, HeadingLevel, PageBreak, PageOrientation, VerticalAlign,
  BorderStyle, ImageRun, Header, Footer, PageNumber, LevelFormat, convertMillimetersToTwip,
} = require('docx');
const D = require('./fat_data.js');

const GREEN='007C5E', WHITE='FFFFFF', DARK='212121', GREY='BFBFBF', LGREY='F2F2F2', FONT='맑은 고딕';
// A4 가로: docx-js는 세로 치수를 받아 내부에서 스왑한다
const P_W=11906, P_H=16838;
const LAND_W=16838;                      // 실제 렌더 폭
const M=convertMillimetersToTwip(15);    // 850
const BODY_W=LAND_W-M*2;                 // 15138

const noB={top:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},bottom:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},left:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},right:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},insideHorizontal:{style:BorderStyle.NONE,size:0,color:'FFFFFF'},insideVertical:{style:BorderStyle.NONE,size:0,color:'FFFFFF'}};
const gB={top:{style:BorderStyle.SINGLE,size:4,color:GREY},bottom:{style:BorderStyle.SINGLE,size:4,color:GREY},left:{style:BorderStyle.SINGLE,size:4,color:GREY},right:{style:BorderStyle.SINGLE,size:4,color:GREY},insideHorizontal:{style:BorderStyle.SINGLE,size:4,color:GREY},insideVertical:{style:BorderStyle.SINGLE,size:4,color:GREY}};

const txt=(t,o={})=>new TextRun({text:t,font:FONT,size:o.size??19,bold:o.bold??false,color:o.color??DARK});
function cellP(t,o={}){
  const ls=String(t).split('\n');
  return ls.map((l,i)=>new Paragraph({alignment:o.align??AlignmentType.LEFT,
    spacing:{before:i===0?30:0,after:i===ls.length-1?30:0,line:250},children:[txt(l,o)]}));
}
function cell(t,{w,fill,bold,color,align,valign,rowSpan,colSpan}={}){
  return new TableCell({width:{size:w,type:WidthType.DXA},
    shading:fill?{type:ShadingType.CLEAR,fill,color:'auto'}:undefined,
    verticalAlign:valign??VerticalAlign.CENTER,
    rowSpan,columnSpan:colSpan,
    margins:{top:50,bottom:50,left:90,right:90},
    children:cellP(t,{bold,color,align})});
}
function table(head,rows,widths,opts={}){
  const centerCols=opts.center??[0];
  return new Table({width:{size:BODY_W,type:WidthType.DXA},columnWidths:widths,borders:gB,
    rows:[
      ...(head?[new TableRow({tableHeader:true,children:head.map((h,i)=>cell(h,{w:widths[i],fill:GREEN,bold:true,color:WHITE,align:AlignmentType.CENTER}))})]:[]),
      ...rows.map(r=>new TableRow({height:opts.h?{value:opts.h,rule:'atLeast'}:undefined,
        children:r.map((c,i)=>cell(c,{w:widths[i],
          align:centerCols.includes(i)?AlignmentType.CENTER:AlignmentType.LEFT,
          bold:opts.boldCols?opts.boldCols.includes(i):centerCols.includes(i)&&i===0,
          fill:opts.fillCols&&opts.fillCols.includes(i)?LGREY:undefined}))}))]});
}
function infoTable(rows,labelW=2600){
  const vW=BODY_W-labelW;
  return new Table({width:{size:BODY_W,type:WidthType.DXA},columnWidths:[labelW,vW],borders:gB,
    rows:rows.map(([k,v])=>new TableRow({children:[
      cell(k,{w:labelW,fill:GREEN,bold:true,color:WHITE}),cell(v,{w:vW})]}))});
}
// 검사표: No. / 검사 항목 / 검사 방법 / 합격 기준 / URS ID / 실측·결과 / 판정
const IW=[750,2450,3750,3750,900,2438,1100];
function inspTable(rows){
  return new Table({width:{size:BODY_W,type:WidthType.DXA},columnWidths:IW,borders:gB,
    rows:[
      new TableRow({tableHeader:true,children:['No.','검사 항목','검사 방법','합격 기준','URS ID','실측 · 결과 기록','판정'].map((h,i)=>cell(h,{w:IW[i],fill:GREEN,bold:true,color:WHITE,align:AlignmentType.CENTER}))}),
      ...rows.map(r=>new TableRow({height:{value:600,rule:'atLeast'},children:[
        cell(r[0],{w:IW[0],bold:true,align:AlignmentType.CENTER}),
        cell(r[1],{w:IW[1],bold:true}),
        cell(r[2],{w:IW[2]}),
        cell(r[3],{w:IW[3]}),
        cell(r[4],{w:IW[4],align:AlignmentType.CENTER,size:17}),
        cell('',{w:IW[5],fill:LGREY}),
        cell('□ 적합\n□ 부적합',{w:IW[6],align:AlignmentType.CENTER,fill:LGREY,size:17}),
      ]}))]});
}
function h1(t){return new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:320,after:180},
  shading:{type:ShadingType.CLEAR,fill:GREEN,color:'auto'},
  children:[new TextRun({text:'  '+t,font:FONT,size:25,bold:true,color:WHITE})]});}
function h2(t){return new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:260,after:130},
  border:{bottom:{style:BorderStyle.SINGLE,size:6,color:GREEN}},
  children:[new TextRun({text:t,font:FONT,size:22,bold:true,color:GREEN})]});}
const body=(t,o={})=>new Paragraph({spacing:{before:40,after:80,line:290},children:[txt(t,o)]});
const bullet=t=>new Paragraph({numbering:{reference:'b1',level:0},spacing:{before:20,after:50,line:290},children:[txt(t)]});
const numRef=r=>t=>new Paragraph({numbering:{reference:r,level:0},spacing:{before:20,after:50,line:290},children:[txt(t)]});
const n1=numRef('n1'), n2=numRef('n2');
const spacer=(h=120)=>new Paragraph({spacing:{after:h},children:[]});
const cap=t=>new Paragraph({spacing:{before:50,after:90},children:[txt(t,{size:17,color:'595959'})]});
const blanks=(n,cols)=>Array.from({length:n},(_,i)=>cols.map((c,j)=>typeof c==='function'?c(i):c));

/* ---------- 표지 (가로 전면 녹색) ---------- */
const cover=new Table({width:{size:LAND_W,type:WidthType.DXA},columnWidths:[LAND_W],borders:noB,
  rows:[new TableRow({height:{value:P_W,rule:'exact'},children:[new TableCell({
    width:{size:LAND_W,type:WidthType.DXA},
    shading:{type:ShadingType.CLEAR,fill:GREEN,color:'auto'},
    verticalAlign:VerticalAlign.CENTER,margins:{top:0,bottom:0,left:900,right:900},
    children:[
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:560},children:[
        new ImageRun({type:'png',data:fs.readFileSync('logo_white.png'),transformation:{width:190,height:100}})]}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:110},children:[
        new TextRun({text:'공장검사 프로토콜 및 성적서',font:FONT,size:60,bold:true,color:WHITE})]}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:420},children:[
        new TextRun({text:'Factory Acceptance Test (FAT) Protocol & Report',font:FONT,size:26,color:WHITE})]}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:90},children:[
        new TextRun({text:'자동용기포장기 (컵실링기)',font:FONT,size:32,bold:true,color:WHITE})]}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:640},children:[
        new TextRun({text:'모델 HW-103RT300  ·  1식(SET)  ·  컵과일 포장라인',font:FONT,size:24,color:WHITE})]}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:50},children:[
        new TextRun({text:'문서번호  FAT-PKG-2026-001   |   Rev.0   |   근거 URS-PKG-2026-001',font:FONT,size:21,color:WHITE})]}),
      new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:50},children:[
        new TextRun({text:'2026. 09. 09.',font:FONT,size:21,color:WHITE})]}),
      new Paragraph({alignment:AlignmentType.CENTER,children:[
        new TextRun({text:'공무팀 (생산지원)',font:FONT,size:21,color:WHITE})]}),
    ]})]})]});

/* ---------- 본문 ---------- */
const C=[]; const A=C.push.bind(C);

A(h1('문서 정보'));
A(infoTable(D.meta,2600));
A(spacer(200));
A(h2('개정 이력'));
A(table(D.revision.head,D.revision.rows,[900,1800,8400,2019,2019],{center:[0,1,3,4]}));

A(h1('1.  개요'));
A(h2('1.1  목적'));
A(body('자동용기포장기(컵실링기) HW-103RT300의 출하 전 공장검사(FAT)를 실시하여, URS-PKG-2026-001에 규정된 요구사항의 충족 여부를 확인하고 출하 승인 가부를 판정한다. 본 문서는 검사 프로토콜(검사 방법·합격 기준)과 성적서(실측 기록·판정)를 겸하며, 검사 현장에서 그대로 기입하여 검수 증빙으로 보관한다.'));
A(h2('1.2  적용 범위 및 검사 항목 수'));
A(bullet('대상: 자동용기포장기(컵실링기) HW-103RT300 본체 1식 및 계약 포함 옵션 전량'));
A(bullet('검사 항목: 총 76항목 (A~H 8개 그룹) — URS 요구사항 71건 중 FAT 단계 검증 대상'));
A(bullet('실물 시험: 당사 실제 용기(PP 외경 95 mm) 및 실제 인쇄 리드지 100컵 기준'));
A(bullet('FAT 미실시 항목은 부속서 A(검증 이관 항목)에 따라 IQ / OQ / PQ / SAT / DOC 단계로 이관'));
A(h2('1.3  판정 기준'));
A(infoTable([
  ['항목별 판정','합격 기준 충족 시 「적합」, 미충족 시 「부적합」. 해당 없음은 「N/A」로 기재하고 사유를 비고에 명시'],
  ['합격 (출하 승인)','전 항목 적합 — 출하 승인'],
  ['조건부 합격','경미 부적합만 존재하고 공급자가 출하 전 시정 후 사진·성적서로 증빙 가능한 경우. 부적합 리스트(6장)에 시정 기한 명기'],
  ['불합격 (재검사)','B그룹(안전장치) 1건 이상 부적합, 또는 E·F그룹(실링 성능·아이마크) 부적합, 또는 경미하지 않은 부적합 3건 이상 — 시정 후 재검사 실시'],
  ['재검사 비용','공급자 귀책 부적합에 의한 재검사 발생 시 소요 비용은 공급자 부담'],
],3200));
A(cap('※ B그룹(안전장치)과 E·F그룹(실링 성능·아이마크 정위치)은 제품 안전성 및 라인 가동에 직결되므로 조건부 합격 대상에서 제외한다.'));

A(new Paragraph({children:[new PageBreak()]}));

A(h1('2.  검사 일정 및 참석자'));
A(table(D.schedule.head,D.schedule.rows,[2600,4969,2600,4969],{center:[],boldCols:[0,2],fillCols:[1,3]}, ));
A(spacer(200));
A(h1('3.  검사 준비물 확인'));
A(body('검사 당일 아래 준비물이 모두 확보되지 않으면 해당 항목은 검사 불가로 처리하고, 재검사 일정을 협의한다.'));
A(table(D.prep.head,D.prep.rows,[1100,7200,2900,2000,1938],{center:[0,3,4],fillCols:[4]}));

A(new Paragraph({children:[new PageBreak()]}));

A(h1('4.  검사 항목 및 결과'));
A(cap('※ 「실측 · 결과 기록」란에는 측정값·관찰 내용을 반드시 기입한다. 「적합」 표기만으로는 검수 증빙이 되지 않는다.'));
D.groups.forEach((g,i)=>{
  A(h2(g.title));
  A(inspTable(g.rows));
  if(i<D.groups.length-1) A(spacer(180));
});

A(new Paragraph({children:[new PageBreak()]}));

A(h1('5.  실측 기록지'));
A(h2('5.1  실링 온도 측정 기록  (관련: E-01 / PR-07)'));
A(body('설정 온도: ________ ℃    측정 기기: ____________________    측정자: __________'));
A(table(['측정점','측정 위치','설정값 (℃)','실측값 (℃)','편차 (℃)','판정','비고'],
  blanks(10,[i=>String(i+1),'','','','','□ 적합  □ 부적합','']),
  [1200,3200,2200,2200,2000,2400,1938],{center:[0,2,3,4,5],fillCols:[1,2,3,4,5,6],h:420}));
A(cap('합격 기준: 전 측정점이 설정값 ±3 ℃ 이내'));

A(h2('5.2  사이클 속도 측정 기록  (관련: C-06 / PR-01)'));
A(table(['회차','측정 시간','카운트 (회)','환산 (cycle/min)','판정','비고'],
  blanks(3,[i=>String(i+1),'60초','','','□ 적합  □ 부적합','']).concat([['평균','—','','','□ 적합  □ 부적합','']]),
  [1200,2400,2600,3000,2400,3538],{center:[0,1,2,3,4],fillCols:[2,3,4,5],h:420}));
A(cap('합격 기준: 평균 9 cycle/min 이상 (4구 기준 36컵/min 이상)'));

A(new Paragraph({children:[new PageBreak()]}));

A(h1('5.3  아이마크 정위치 측정 기록  (관련: F-02 · F-03 / PR-06)'));
A(body('측정 기기: 버니어캘리퍼스 / 각도 게이지    기준: 로고 중심 편차 ±1.5 mm 이내, 회전 각도 편차 ±3° 이내'));
A(table(['컵 No.','로고 중심 편차 X (mm)','로고 중심 편차 Y (mm)','회전 각도 편차 (°)','판정','컵 No.','로고 중심 편차 X (mm)','로고 중심 편차 Y (mm)','회전 각도 편차 (°)','판정'],
  blanks(10,[i=>String(i+1),'','','','□ 적합\n□ 부적합',i=>String(i+11),'','','','□ 적합\n□ 부적합']),
  [1000,1880,1880,1700,1110,1000,1880,1880,1700,1108],{center:[0,1,2,3,4,5,6,7,8,9],fillCols:[1,2,3,4,6,7,8,9],h:400}));
A(cap('20컵 전량이 기준을 충족해야 F-02 · F-03 적합. 1컵이라도 초과 시 부적합 처리하고 원인(센서 감도·필름 인쇄 피치·정렬 정밀도)을 부적합 리스트에 기재한다.'));

A(h2('5.4  실링 100컵 집계  (관련: E-02 ~ E-05 · E-09 / PR-02 · PR-04 · PR-05)'));
A(table(['확인 항목','검사 수량','불량 수','불량률 (%)','판정','불량 내용 · 원인'],
  [['미실링 (4구 중 누락)','100','','','□ 적합  □ 부적합',''],
   ['컵 변형 · 플랜지 눌림','100','','','□ 적합  □ 부적합',''],
   ['실링 외관 (주름 · 기포 · 타흔)','100','','','□ 적합  □ 부적합',''],
   ['트리밍 버 · 필름 잔재','100','','','□ 적합  □ 부적합',''],
   ['이형 형상 재현 불량','100','','','□ 적합  □ 부적합',''],
   ['합계 (중복 제외)','100','','','□ 적합  □ 부적합','']],
  [4500,1700,1700,1800,2400,3038],{center:[1,2,3,4],boldCols:[0],fillCols:[2,3,4,5],h:420}));
A(cap('합격 기준: 합계 불량률 0.5% 이하 — 100컵 기준 부적합 0건'));

A(new Paragraph({children:[new PageBreak()]}));

A(h1('5.5  누설 · 개봉성 시험 기록  (관련: E-06 ~ E-08 / PR-05)'));
A(table(['시험 종류','시험 조건','검사 수량','불량 수','판정','관찰 내용'],
  [['도립 누액 시험','내용물 충전 후 5분 도립','20','','□ 적합  □ 부적합',''],
   ['감압 누설 시험','감압 챔버 (조건: ____________)','10','','□ 적합  □ 부적합',''],
   ['실링 강도 (개봉성)','OPEN 탭 수동 개봉','10','','□ 적합  □ 부적합','']],
  [2800,4200,1600,1600,2400,2538],{center:[2,3,4],boldCols:[0],fillCols:[3,4,5],h:480}));
A(cap('합격 기준: 도립 누액 0건 / 감압 누기 0건 / 개봉 시 리드지 찢김·용기 파손 없이 탭에서 균일 박리'));

A(h2('5.6  소음 측정 기록  (관련: C-03 / FN-09)'));
A(table(['측정 위치','측정 조건','측정값 dB(A)','판정','비고'],
  [['전면','설비 외곽 1 m / 높이 1.5 m','','□ 적합  □ 부적합',''],
   ['후면','설비 외곽 1 m / 높이 1.5 m','','□ 적합  □ 부적합',''],
   ['좌측','설비 외곽 1 m / 높이 1.5 m','','□ 적합  □ 부적합',''],
   ['우측','설비 외곽 1 m / 높이 1.5 m','','□ 적합  □ 부적합',''],
   ['진공펌프 근접','펌프 1 m','','□ 적합  □ 부적합',''],
   ['최대값','—','','□ 적합  □ 부적합','']],
  [2400,4200,2400,2400,3738],{center:[2,3],boldCols:[0],fillCols:[2,3,4],h:420}));
A(cap('합격 기준: 최대값 80 dB(A) 이하. 초과 시 원인(진공펌프·에어 배기)과 저감 대책을 부적합 리스트에 기재'));

A(h2('5.7  결로 조건 실링 시험 기록  (관련: E-10 / PR-11)'));
A(table(['항목','내용 / 측정값','항목','내용 / 측정값'],
  [['용기 표면 온도 (투입 시)','','검사장 온·습도',''],
   ['결로 대책 작동 여부','□ 작동  □ 미작동  (방식: ____________)','실링 불량 수 / 20컵',''],
   ['판정','□ 적합  □ 부적합','관찰 내용','']],
  [3000,4569,3000,4569],{center:[],boldCols:[0,2],fillCols:[1,3],h:480}));
A(cap('합격 기준: 냉장(0~10 ℃) 용기 20컵 실링 시 결로에 의한 실링 불량 0건. 냉장 용기는 당사가 보온박스로 지참한다.'));

A(new Paragraph({children:[new PageBreak()]}));

A(h1('6.  부적합 리스트 (Punch List)'));
A(body('부적합 판정 항목은 아래에 전량 기재하고, 공급자 확인 서명을 받는다. 시정 완료 후 재검증 결과를 기입하여 마감한다.'));
A(table(['No.','관련 검사번호','부적합 내용','시정 요구사항','시정 기한','재검증 방법','시정 확인'],
  blanks(8,[i=>String(i+1),'','','','','','□ 완료']),
  [750,1700,4000,4000,1500,1900,1288],{center:[0,4,6],fillCols:[1,2,3,4,5,6],h:560}));
A(cap('※ 경미 부적합(서류 보완, 표지 부착, 도색 등)과 중대 부적합(안전장치, 실링 성능, 아이마크 정위치)을 구분하여 「부적합 내용」란 앞에 [경미] / [중대]로 표기한다.'));

A(new Paragraph({children:[new PageBreak()]}));

A(h1('7.  종합 판정'));
A(h2('7.1  그룹별 집계'));
A(table(['그룹','검사 분야','검사 항목 수','적합','부적합','N/A','비고'],
  D.summary.map(s=>[s[0],s[1],s[2],'','','','']).concat([['합계','—','76','','','','']]),
  [1000,4200,2000,1700,1700,1700,2838],{center:[0,2,3,4,5],boldCols:[0,1],fillCols:[3,4,5,6],h:420}));
A(spacer(200));
A(h2('7.2  판정'));
A(body('□  합격 (출하 승인)          □  조건부 합격 (부적합 리스트 시정 조건)          □  불합격 (시정 후 재검사)',{bold:true,size:22}));
A(spacer(120));
A(body('종합 의견 :  ______________________________________________________________________________________________________________________________'));
A(body('              ______________________________________________________________________________________________________________________________'));
A(body('재검사 필요 시 예정일 :  __________________          출하 예정일 :  __________________'));
A(spacer(240));
A(new Paragraph({children:[new PageBreak()]}));
A(h2('7.3  검사 확인 서명'));
A(new Table({width:{size:BODY_W,type:WidthType.DXA},columnWidths:[2200,2600,2600,2600,2600,2538],borders:gB,
  rows:[
    new TableRow({tableHeader:true,children:['구분','당사 공무팀','당사 생산팀','당사 품질팀','공급자 기술책임','공급자 대표'].map((h,i)=>cell(h,{w:[2200,2600,2600,2600,2600,2538][i],fill:GREEN,bold:true,color:WHITE,align:AlignmentType.CENTER}))}),
    new TableRow({children:['성명','','','','',''].map((h,i)=>cell(h,{w:[2200,2600,2600,2600,2600,2538][i],align:AlignmentType.CENTER,bold:i===0}))}),
    new TableRow({height:{value:1300,rule:'atLeast'},children:['서명','','','','',''].map((h,i)=>cell(h,{w:[2200,2600,2600,2600,2600,2538][i],align:AlignmentType.CENTER,bold:i===0}))}),
    new TableRow({children:['일자','','','','',''].map((h,i)=>cell(h,{w:[2200,2600,2600,2600,2600,2538][i],align:AlignmentType.CENTER,bold:i===0}))}),
  ]}));

A(spacer(300));

A(h1('부속서 A.  검증 이관 항목 (FAT 미실시)'));
A(body('아래 URS 요구사항은 공장 조건에서 검증이 불가하거나 현장·양산 조건 확인이 필요하여 FAT 대상에서 제외한다. 각 항목은 지정된 단계에서 검증하고, 최종 인수(검수) 전 전 항목 완료를 확인한다.'));
A(table(D.deferred.head,D.deferred.rows,[3200,4600,4600,2738],{center:[3],boldCols:[0]}));
A(spacer(240));
A(h2('참고 — 검증 단계 구성'));
A(body('FAT (출하 전, 본 문서)  →  IQ (설치확인)  →  OQ (운전확인)  →  PQ (성능확인, 3배치 × 500컵)  →  SAT · 교육  →  검수조서 서명 (최종 인수)',{bold:true}));
A(cap('※ 필수(M) 항목 전 항목 합격 및 문서 제출 완료 시 검수조서에 서명하며, 잔금 지급은 인수 완료 후 계약 조건에 따라 진행한다.'));

/* ---------- 조립 ---------- */
const hdr=new Header({children:[new Paragraph({alignment:AlignmentType.RIGHT,spacing:{after:50},
  border:{bottom:{style:BorderStyle.SINGLE,size:6,color:GREEN}},
  children:[
    new TextRun({text:'FAT-PKG-2026-001  |  자동용기포장기(컵실링기) HW-103RT300  |  공장검사 프로토콜 및 성적서        ',font:FONT,size:16,color:'595959'}),
    new ImageRun({type:'png',data:fs.readFileSync('logo_green.png'),transformation:{width:70,height:37}})]})]});
const ftr=new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,
  children:[new TextRun({children:['- ',PageNumber.CURRENT,' -'],font:FONT,size:16,color:'595959'})]})]});

const landscape={page:{size:{width:P_W,height:P_H,orientation:PageOrientation.LANDSCAPE}}};

const doc=new Document({
  creator:'Sweet Balance 공무팀',
  title:'자동용기포장기(컵실링기) HW-103RT300 공장검사 프로토콜 및 성적서(FAT)',
  description:'FAT-PKG-2026-001 Rev.0',
  styles:{default:{document:{run:{font:FONT,size:19,color:DARK}}}},
  numbering:{config:[
    {reference:'b1',levels:[{level:0,format:LevelFormat.BULLET,text:'•',alignment:AlignmentType.LEFT,
      style:{paragraph:{indent:{left:400,hanging:220}},run:{color:GREEN,font:FONT}}}]},
    {reference:'n1',levels:[{level:0,format:LevelFormat.DECIMAL,text:'%1.',alignment:AlignmentType.START,
      style:{paragraph:{indent:{left:420,hanging:240}},run:{color:GREEN,bold:true,font:FONT}}}]},
    {reference:'n2',levels:[{level:0,format:LevelFormat.DECIMAL,text:'%1.',alignment:AlignmentType.START,
      style:{paragraph:{indent:{left:420,hanging:240}},run:{color:GREEN,bold:true,font:FONT}}}]},
  ]},
  sections:[
    {properties:{page:{...landscape.page,margin:{top:0,right:0,bottom:0,left:0,header:0,footer:0}}},children:[cover]},
    {properties:{page:{...landscape.page,margin:{top:M,right:M,bottom:M,left:M,
      header:convertMillimetersToTwip(8),footer:convertMillimetersToTwip(8)}}},
     headers:{default:hdr},footers:{default:ftr},children:C},
  ],
});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync('FAT_자동용기포장기_컵실링기_HW-103RT300_Rev0.docx',b);console.log('written',b.length);});
