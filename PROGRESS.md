# 진행 상황 (어휘 워크북 자동 생성기)

> 마지막 업데이트: 2026-07-24

---
## 🔖 다음 세션 인수인계 (직장 PC에서 이어서) — 먼저 읽기

**바로 할 일: Bricks 1500 보강** (엑셀에 Unit 16-30만 있고 **Unit 1-15 누락**. PDF 있음)

이 대화 없이도 아래 순서만 따라 하면 됨. (Bricks 900을 이 방식으로 이미 완료함 — 그대로 복제)

1. **최신 pull**: `git pull` (build_workbook.py, import_pdf.py, PDFs/, excel/ 다 들어있음)
2. **PDF에서 Unit 1-15 추출** (english+뜻만, 품사 자동 제거):
   ```
   python import_pdf.py "PDFs/Bricks Vocabulary 1500.pdf" --book "Bricks Vocabulary 1500" --unit-label Unit --from-unit 1 --to-unit 15 --out "scratch/b1500_1_15.xlsx"
   ```
   → `scratch/b1500_1_15.xlsx`(스켈레톤) + `.review.txt`(검수용) 생성. 유닛당 20단어 × 15 = 300 확인.
3. **예문·동의어·반의어 생성**(★수작업/AI 파트): 각 단어에 예문(짧고 표제어 포함, 100%), 동의/반의는 확실한 것만.
   기존 Bricks 형식과 맞춤: 뜻은 품사 없이 plain(`아이, 어린이`), 동의/반의 헤더는 `synonym`/`antonym`(단수).
4. **엑셀 병합**: `excel/Bricks Vocabulary 1500.xlsx`의 기존 Unit 16-30은 유지, 앞에 Unit 1-15 추가 → 유닛 번호순 정렬(1→30).
   (Bricks 900 때 쓴 병합 스크립트 패턴: 기존 행 읽고 해당 유닛 제외 후, 새 행과 합쳐 유닛·word_no 순 정렬해 시트 재작성.)
5. **워크북 생성**: 동의/반의 적으니 merge-choice, 30유닛이니 15DAY/권:
   ```
   python build_workbook.py "excel/Bricks Vocabulary 1500.xlsx" --split 15 --merge-choice --answer --out "output_new/Bricks Vocabulary 1500"
   ```
   → 구 파일(권 없는 `_DAY16-30`) 삭제: `find "출력폴더" -maxdepth 1 -type f ! -name '*권*' -delete`
6. **커밋·푸시**. PROGRESS 이 표 갱신.
   - **2300도 동일**(Unit 1-15 누락, `PDFs/Bricks Vocabulary 2300.pdf` 있음). book_name `Bricks Vocabulary 2300`.

### build_workbook.py 옵션 요약
- `--split N` 권당 N유닛(자투리는 마지막 권). `--answer` 정답본. 페이지 번호(하단 중앙)는 **자동**.
- `--merge-choice` 동의어+반의어를 한 페이지로(동의/반의 적은 교재). 
- `--merge-practice` 예문빈칸+동의+반의를 한 페이지로(단어 아주 적은 교재, 예: Bricks 300 8단어/유닛).

### import_pdf.py (PDF→스켈레톤 엑셀)
- "찐 실력자 양성소" 형식 PDF 공통. footer의 `<book_name> [- DAY|Unit] N`으로 유닛번호 인식.
- 뜻 품사(n./v./adj.…) 자동 제거, 여러 품사 줄 합치고 중복 제거. 영단어=한글 없는 줄.
- ⚠️ 예문·동의/반의는 안 만듦(수작업). english는 원문 그대로(`mother[mom]` 등 유지).

### 교재별 상태 (2026-07-24)
| 교재 | 상태 | 비고 |
|---|---|---|
| 능률 중등기본/초등기본/중등필수/초등필수 | ✅ 완료(구버전) | **페이지번호 없음** → 재생성 필요(아래) |
| Bricks 3100/3900/4800 | ✅ 완료(구버전) | **페이지번호 없음** → 재생성 필요 |
| Bricks 300 | ✅ 완료 | 2권, merge-practice, 페이지번호 O |
| Bricks 900 | ✅ 완료 | 2권, Unit16-30 보강, merge-choice, 페이지번호 O |
| 뜯어먹는 1200 | ✅ 완료 | 3권(20DAY), merge-choice, 페이지번호 O |
| **Bricks 1500** | ⬜ **다음 작업** | Unit1-15 누락. PDF 있음 |
| **Bricks 2300** | ⬜ 대기 | Unit1-15 누락. PDF 있음 |

**아직 페이지번호 없는 완료본**(능률 4종 + Bricks 3100/3900/4800): 페이지번호는 재생성해야 붙음.
합치기 추천 — 능률 초등기본은 merge-choice, 나머지는 분리. 분권/합치기 최종안은 아래 '분권 재검토' 참고. 시간 될 때 일괄 재생성.

### 새 교재 임포트 대기 (PDF는 `PDFs/`에 있음, 대형 — 예문 생성 오래 걸림)
능률 고등 기본 · 수능 필수 · 수능 고난도 · 어원편 고등 · 중등 고난도 · 뜯어먹는 1800 · 뜯어먹는 영숙어 1000
- 모두 "찐 실력자" 형식이면 `import_pdf.py`로 추출 가능(unit-label은 DAY/Unit 확인). 능률류 뜻은 `[명]`형일 수 있으니 첫 유닛 review.txt로 형식 확인 후 맞추기.

---

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
  | 300 | 320 / 40(8단어) | **2권 · 117·120p** | 유닛당 8단어라 `--merge-practice`(예문+동의+반의 1페이지)+20DAY/권 |
  | 900 | **600 / 30** | **2권 · 105p** | PDF에서 **Unit 16~30 보강**(엑셀 15유닛만 있었음). 품사(n./v.) 제거, 예문·동의/반의 생성. merge-choice·15DAY/권 |
  | 1500 | 300 / 15 | 1권 · 117p | Unit 16~30 라벨 |
  | 2300 | 374 / 15 | 1권 · 147p | Unit 16~30 라벨 |
  | 3100 | 749 / 30 | 3권 · ~110p | |
  | 3900 | 900 / 30 | 3권 · ~127p | |
  | 4800 | 900 / 30 | 3권 · ~130p | book_name 'Bricks Vocabulary 30' 1줄 교정함 |
- **→ 전 교재(능률 4종 + Bricks 7종 = 11종) 전량 생성 완료.**
- **뜯어먹는 중학 기본 영단어 1200** (1200단어 / 60 DAY, 유닛당 20단어)
  - ✅ 전량 생성 완료(2026-07-24). **20 DAY/권 → 3권**(DAY 1-20/21-40/41-60, 각 ~138쪽).
    동의/반의 적어(11%/19%) `--merge-choice`로 합본, 페이지 번호(하단 중앙) 적용.
  - **PDF→xlsx 워크플로 첫 적용**: 사용자 제공 PDF(`PDFs/뜯어먹는 중학 기본 영단어 1200.pdf`)에서 english+뜻 자동 추출(파서: `찐 실력자…` 헤더/`… - DAY N` footer 스킵, 영단어/한글 교대 파싱).
    예문 100%·동의어·반의어는 기존 스타일로 보완 생성(기초어라 동의/반의 sparse: 11%/19%). english 대괄호(`mother[mom]`)는 원문 그대로 유지.
  - 입력: `excel/뜯어먹는 중학 기본 영단어 1200.xlsx`.
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
