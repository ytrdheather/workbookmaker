# 진행 상황 (어휘 워크북 자동 생성기)

> 마지막 업데이트: 2026-07-22

## 프로젝트 개요
엑셀(단어/뜻/예문/동의어/반의어)을 입력받아 DAY(유닛)별 어휘 워크북을
HTML로 조립하고 헤드리스 Chrome으로 PDF까지 뽑는 도구.

- 핵심 스크립트: `build_workbook.py`
- 실행 예: `python build_workbook.py <입력.xlsx> --from 1 --to 60 --answer --out output`
  - `--from/--to` : 생성할 DAY 범위
  - `--answer` : 정답 버전도 함께 생성
- 유닛 페이지 구성: ①단어 목록 ②4회 쓰기 ③누적 복습 시험(N-2, N-1) ④예문 빈칸 ⑤반의어 고르기 ⑥동의어 고르기

## 지금까지 한 것
- 전체 파이프라인 완성 (엑셀 → HTML → PDF, 학생용/정답용).
- 교재 데이터: "주니어 능률 VOAC 기본", 총 1200단어 / DAY 01~60 (하루 20단어).
- 현재 **DAY 1~3만 샘플 생성**해서 레이아웃 검수 (`output/`, `output_new/`).
- **예문 빈칸 페이지 레이아웃 확정**: 영어 문장이 먼저 나오고, 문장 뒤에 뜻을 회색으로 표기.
  줄 간격을 넉넉히(줄 높이 1.7, 항목 간격 17px) 늘려 아이들이 빈칸에 쓸 공간 확보. 20문항이 A4 한 장에 들어감.

## 다음에 할 일 (직장에서 이어서)
1. **원본 입력 엑셀 파일이 필요함.** 원래 위치에서 사라져서, 현재 DAY1-3 재생성은
   기존 출력 HTML의 단어표에서 데이터를 역추출해 임시로 진행했음.
   전체 DAY 1~60을 다시 뽑으려면 원본 1200단어 엑셀(book_name/unit_name/word_no/
   english/meaning/synonyms/antonyms/example 열)이 있어야 함. → **엑셀 파일 챙겨오기.**
2. 엑셀 확보 후 `--from 1 --to 60 --answer` 로 전량 생성.
3. `output` vs `output_new` 두 폴더 정리 (최종본 하나로 통일).

## 참고
- 입력 엑셀 헤더: `book_name, unit_name, word_no, english, meaning, synonyms, antonyms, example`
- 폰트: Pretendard(설치 필요), 로고: `source/logo.png`
- PDF 변환은 Chrome/Edge 헤드리스 사용 (`--print-to-pdf`).
