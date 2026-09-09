# URS — 자동용기포장기(컵실링기) HW-103RT300

컵과일 포장라인 자동용기포장기(컵실링기) 신규 구매에 대한 사용자 요구사항 명세서(URS).

| 파일 | 용도 |
|---|---|
| `URS_자동용기포장기_컵실링기_HW-103RT300_Rev0.docx` | 배포·결재용 (Sweet Balance 브랜드 양식, 13p) |
| `URS_자동용기포장기_컵실링기_HW-103RT300_Rev0.md` | 텍스트 원본 (검토·수정용) |
| `data.js` | 요구사항 데이터 (단일 원본) |
| `gen.js` | docx 생성 스크립트 |
| `md.js` | markdown 생성 스크립트 |

## 개요

| 항목 | 내용 |
|---|---|
| 문서번호 | URS-PKG-2026-001 (Rev.0) |
| 대상설비 | 자동용기포장기(컵실링기) HW-103RT300 / 1식 |
| 공급업체 | 하이퍼박(주) |
| 요구사항 수 | 71건 (필수 M 60건 / 권장 D 11건) |
| 근거 자료 | 설비 구매 품의(컵과일 포장라인), 견적서(2026.7.14.), 공급계약서(2026.7.15.), 제품 도면 |

## 수정 방법

요구사항은 `data.js` 한 곳만 수정하고 두 스크립트를 재실행하면 docx·md가 동시에 갱신된다.

```bash
npm install docx
node gen.js   # docx 생성
node md.js    # markdown 생성
```

렌더링 확인:

```bash
soffice --headless --convert-to pdf URS_*.docx
pdftoppm -jpeg -r 90 URS_*.pdf page
```
