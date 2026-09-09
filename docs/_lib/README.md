# Sweet Balance 문서 공용 라이브러리

컵실링기 HW-103RT300 검증 문서 세트가 공용으로 사용하는 브랜드 양식 모듈.

| 파일 | 역할 |
|---|---|
| `sbdoc.js` | docx 생성 — 표지(Primary Green 전면 + 흰색 로고), 헤더·푸터, 표·정보표·점검표, 제목 스타일 |
| `md.js` | markdown 생성 헬퍼 |
| `logo_green.png` / `logo_white.png` | 브랜드 로고 (밝은 배경 / 어두운 배경) |

## 사용법

```js
const S = require('../_lib/sbdoc')({ landscape: true, marginMm: 15, baseSize: 19 });
const { h1, h2, body, bullet, table, infoTable, blanks, pageBreak } = S;
const C = [];
C.push(h1('1. 개요'));
C.push(table(['항목','내용'], [['a','b']], [3000, 12138]));
S.render({ file: 'out.docx', title: '...', headerText: '...', cover: {...}, children: C });
```

- `landscape: false` → A4 세로 (검수조서 등 결재 문서)
- `landscape: true` → A4 가로 (점검표·체크리스트)
- `BODY_W` = 본문 사용 가능 폭(twip). 표의 컬럼 폭 합계는 반드시 `BODY_W`와 일치해야 한다.

## 주의

- 표 컬럼 폭 합계 ≠ BODY_W 이면 Word/한글에서 표가 어긋난다. 가변 폭은 `fit()` 헬퍼로 잔여 폭을 분배한다.
- `table()`은 헤더가 있는 표, `infoTable()`은 라벨(녹색)+값 2열 표.
- 페이지를 정확히 채운 뒤 `pageBreak()`를 두면 빈 페이지가 생긴다. 렌더 후 빈 페이지를 확인할 것.

## 렌더 검증

```bash
soffice --headless --convert-to pdf out.docx
pdftoppm -jpeg -r 80 out.pdf page
```
