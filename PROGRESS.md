# 진행 상황 (어휘 워크북 자동 생성기)

> 마지막 업데이트: 2026-07-23

## 프로젝트 개요
엑셀(단어/뜻/예문/동의어/반의어)을 입력받아 DAY(유닛)별 어휘 워크북을
HTML로 조립하고 헤드리스 Chrome으로 PDF까지 뽑는 도구.

- 핵심 스크립트: `build_workbook.py`
- 실행 예: `python build_workbook.py <입력.xlsx> --split 10 --answer --out output`
  - `--split N` : **분권** — 권당 N개 유닛으로 나눠 여러 권 생성(나머지<N는 마지막 권에 합침). 파일명에 `_N권_` 붙음. 0이면 단권.
  - `--from/--to` : 단권 생성 시 DAY 범위 (분권 미사용 시)
  - `--answer` : 정답 버전도 함께 생성
  - `--out` : 출력 폴더 (파일명은 엑셀 안의 `book_name` 열 값으로 자동 생성)
- **분권 정책: 10유닛/권** 확정(2026-07-23). 400쪽 단권이 제본·휴대 불가라 전 교재 분권.
  누적 복습(N-2·N-1)은 항상 전체 목록에서 가져오므로 권 경계에서도 안 깨짐(예: 2권 첫 DAY도 1권 마지막 2개 DAY 복습 포함).
- 유닛 페이지 구성(DAY당 ~7쪽): ①단어 목록 ②**암기 노트(딸기케이크식 4단 자가시험, 빈칸·단어 직접 쓰기)** ③4회 쓰기 ④누적 복습 시험(N-2, N-1) ⑤예문 빈칸 ⑥반의어 고르기 ⑦동의어 고르기
  - 암기 노트는 단어목록 바로 뒤에 배치해 '외우는 단계'를 강제. 행 높이는 단어 수에 맞춰 자동(하단 로고 여백 확보). standalone 양식은 `templates/`에도 유지.

## 작업 환경 셋업 (PC마다 필요)
> `.claude`(메모리)는 구글 드라이브에 심볼릭 링크로 연결돼 PC 간 동기화됨. 단, **Google Drive Desktop이 켜져 있어야** 함.

각 PC에서 이어 작업하려면:
1. **Python + openpyxl** 설치 (`winget install Python.Python.3.12` → `python -m pip install openpyxl`)
2. **Chrome**(또는 Edge) 설치 — PDF 렌더용
3. **폰트 Pretendard** 설치 (Thin~Black)
4. **GitHub 재인증** — 이 레포 clone 후 push하려면 자격증명 필요
   - `git`은 있으나 `gh`(GitHub CLI)는 없을 수 있음. device flow 또는 credential manager 사용.

## 데이터 (입력 엑셀) — `excel/` 폴더
원본 11종을 레포에 포함(공개). 표준 헤더: `book_name, unit_name, word_no, english, meaning, synonym(s), antonym(s), example` (동의어/반의어는 단수·복수 둘 다 인식).

| 파일 | book_name(내부) | 단어 수 | 예상 DAY |
|---|---|---|---|
| 능률VOCA 초등 기본 | | 406 | ~21 |
| 능률VOCA 초등 필수 | | 550 | ~28 |
| 능률VOCA 중등 기본 | 주니어 능률 VOAC 기본 | 1200 | 60 |
| 능률VOCA 중등 필수 | | 1128 | ~57 |
| Bricks Vocabulary 300 | Bricks Vocabulary 300 | 320 | 16 |
| Bricks Vocabulary 900 | | 999 | 50 |
| Bricks Vocabulary 1500 | | 300 | 15 |
| Bricks Vocabulary 2300 | | 375 | ~19 |
| Bricks Vocabulary 3100 | | 750 | ~38 |
| Bricks Vocabulary 3900 | | 900 | 45 |
| Bricks Vocabulary 4800 | | 900 | 45 |

- `Bricks Vocabulary 300.xlsx`는 원래 헤더가 달라(`Unit Name/Number/Word...`) 표준 스키마로 정규화함(320단어, book_name 추가).

## 교재별 진행 (하나씩 순차 처리)
- **능률VOCA 중등 기본** (= "주니어 능률 VOAC 기본", 1200단어 / DAY 01~60)
  - ✅ DAY 1~3 샘플 생성 → 레이아웃 승인 완료. 결과물: `output_new/능률VOCA 중등 기본/`
  - ✅ 전량 생성 완료. 입력: `excel/능률VOCA 중등 기본.xlsx`. **10유닛/권 분권 → 6권**(DAY 1-10/11-20/…/51-60).
- **능률VOCA 초등 기본** (406단어 / 32유닛, DAY 01~39 중 32개)
  - ✅ 전량 생성 완료. 입력: `excel/능률VOCA 초등 기본.xlsx`. **10유닛/권 분권 → 3권**(10·10·12).
- **능률VOCA 중등 필수** (1200단어 / 50유닛, DAY 01~50)
  - ✅ 전량 생성 완료. 입력: `excel/능률VOCA 중등 필수.xlsx`. **10유닛/권 분권 → 5권**(10씩).
  - 데이터 정리 이력:
    - book_name 잔여 1줄(DAY 32 'set up' → `7`)을 `능률VOCA 중등 필수`로 교정.
    - **DAY 27·28·29(72단어)가 원본에 누락**돼 있었음 → 사용자 제공 PDF(`Downloads/능률VOCA 중등 필수 [2025] DAY 27.pdf`, 3페이지=27·28·29)에서
      단어·뜻 추출(품사태그 유지, 숙어는 태그 제거), 예문·동의어·반의어는 기존 스타일로 보완 생성 후 삽입. DAY번호순 재정렬. → 1128→1200단어.
- **능률VOCA 초등 필수** (550단어 / 32유닛)
  - ✅ 전량 생성 완료. 입력: `excel/능률VOCA 초등 필수.xlsx`. **10유닛/권 분권 → 3권**(10·10·12).
  - DAY 5·10·15·20·25·30·35 결번은 교재의 누적 TEST(복습) DAY로 정상 구조(데이터 손실 아님).
- **→ 능률VOCA 시리즈 4종(초등 기본·초등 필수·중등 기본·중등 필수) 전량 완료.**
- **Bricks Vocabulary 시리즈 7종** — ✅ 전량 생성 완료(2026-07-23). 10유닛/권 분권(15유닛 책은 자투리 합쳐 단권).
  | 교재 | 단어/유닛 | 권수·쪽수 | 비고 |
  |---|---|---|---|
  | 300 | 320 / 40(8단어) | 4권 · 77~80p | 저학년, 유닛 잘게 |
  | 900 | 300 / 15 | 1권 · 117p | |
  | 1500 | 300 / 15 | 1권 · 117p | Unit 16~30 라벨 |
  | 2300 | 374 / 15 | 1권 · 147p | Unit 16~30 라벨 |
  | 3100 | 749 / 30 | 3권 · ~110p | |
  | 3900 | 900 / 30 | 3권 · ~127p | |
  | 4800 | 900 / 30 | 3권 · ~130p | book_name 'Bricks Vocabulary 30' 1줄 교정함 |
- **→ 전 교재(능률 4종 + Bricks 7종 = 11종) 전량 생성 완료.**
  - 참고: 구글드라이브 원본 `능률VOCA 중등 필수.xlsx`는 아직 옛 book_name·DAY27~29 누락 상태. 레포 사본만 정리됨.

**결과물 출력 규칙:** `output_new/<교재이름>/` 아래에 넣는다. 분권 파일명: `<book_name>_N권_DAY{a}-{b}.pdf`(+`_정답`).

## 다음에 할 일
1. ✅ 능률VOCA 4종 + Bricks 7종 = **전 11종 전량 생성·분권·푸시 완료.**
2. (선택) 4회 쓰기 vs 암기 노트 중복 정리 여부 검토 — 분량 부담되면 4회 쓰기 제거 고려.
3. (선택) 구글드라이브 원본 엑셀들도 레포 정리본과 동기화(중등 필수 등).

## 참고
- 입력 엑셀 헤더: `book_name, unit_name, word_no, english, meaning, synonyms, antonyms, example`
- 폰트: Pretendard, 로고: `source/logo.png` (스크립트가 base64로 임베드, 실행 위치 무관)
- PDF 변환은 Chrome/Edge 헤드리스 사용 (`--print-to-pdf`).
- 페이지 수 검증: PDF 바이트에서 `/Type /Page` 개수 카운트. 레이아웃 확인: `--screenshot` + `--window-size=794,3400`.
