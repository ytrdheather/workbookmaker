# -*- coding: utf-8 -*-
"""찐 실력자 양성소 형식 PDF -> 엑셀 스키마 추출기 (재사용).

지원 형식: 페이지마다 [헤더] + (영단어 / 뜻) 반복 + [footer].
  - footer 예: "뜯어먹는 중학 기본 영단어 1200 - DAY 01"  또는  "Bricks Vocabulary 900 16"
  - 뜻의 품사(n./v./adj./adv. 등)는 자동 제거(있으면), 여러 품사 줄은 합치고 중복 제거.
  - 영단어=한글 없는 줄, 뜻=한글 있는 줄 (능률처럼 한 단어에 뜻 여러 줄이어도 OK).

사용:
  python import_pdf.py "PDFs/Bricks Vocabulary 1500.pdf" --book "Bricks Vocabulary 1500" \\
         --unit-label Unit --from-unit 1 --to-unit 15 --out "scratch_pdf/bricks1500_1_15.xlsx"

출력: english + meaning만 채운 엑셀(스키마 동일). synonym/antonym/example은 빈칸 -> 이후 사람이/AI가 보강.
검수용 텍스트도 같은 이름(.review.txt)으로 저장.

⚠️ 예문·동의어·반의어는 이 스크립트가 만들지 않는다. 추출 후 별도로 생성해 채운다(수작업/AI).
"""
import fitz, re, argparse, os
import openpyxl

HEADER = "book_name,unit_name,word_no,english,meaning,synonym,antonym,example".split(",")
POS = re.compile(r"\b(n|v|adj|adv|prep|conj|pron|int|aux|art)\.\s*", re.I)


def has_hangul(s):
    return bool(re.search(r"[가-힣]", s))


def clean_meaning(mlines):
    seen = []
    for ml in mlines:
        m = POS.sub("", ml).strip(" ;,\t")
        if m and m not in seen:
            seen.append(m)
    return ", ".join(seen)


def parse(pdf_path, book_name):
    """-> {unit_int: [(english, meaning), ...]}"""
    d = fitz.open(pdf_path)
    # footer의 유닛 번호: "<book> [- DAY|Unit] N"
    foot = re.compile(re.escape(book_name) + r"\s*-?\s*(?:DAY|Unit)?\s*(\d+)", re.I)

    def is_noise(l):
        return (not l or l.startswith("PAGE") or "찐 실력자" in l or l.startswith("031.")
                or l.startswith(book_name) or l.startswith("Name") or "Show and Prove" in l)

    result = {}
    for pg in d:
        t = pg.get_text()
        m = foot.search(t)
        unit = int(m.group(1)) if m else None
        if unit is None:
            continue
        lines = [l.strip().replace("\xa0", " ").strip() for l in t.splitlines()]
        content = [l for l in lines if not is_noise(l)]
        entries = []
        for l in content:
            if has_hangul(l):
                if entries:
                    entries[-1][1].append(l)
            else:
                entries.append([l, []])
        result[unit] = [(e, clean_meaning(ml)) for e, ml in entries]
    return result


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--book", required=True, help="book_name (footer/출력에 사용)")
    ap.add_argument("--unit-label", default="DAY", help="유닛 라벨: DAY 또는 Unit")
    ap.add_argument("--from-unit", type=int, default=1)
    ap.add_argument("--to-unit", type=int, default=999)
    ap.add_argument("--out", required=True, help="출력 xlsx 경로")
    args = ap.parse_args()

    data = parse(args.pdf, args.book)
    units = sorted(u for u in data if args.from_unit <= u <= args.to_unit)

    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(HEADER)
    review = []
    for u in units:
        review.append(f"===== {args.unit_label} {u:02d} ({len(data[u])} words) =====")
        for i, (eng, mean) in enumerate(data[u], 1):
            ws.append([args.book, f"{args.unit_label} {u:02d}", i, eng, mean, None, None, None])
            review.append(f"{i:2d}. {eng:<18} | {mean}")
        review.append("")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    wb.save(args.out)
    open(args.out + ".review.txt", "w", encoding="utf-8").write("\n".join(review))

    counts = {u: len(data[u]) for u in units}
    print("추출 유닛:", units[0] if units else None, "~", units[-1] if units else None,
          "| 유닛수", len(units), "| 총 단어", sum(counts.values()))
    print("유닛당 단어수:", sorted(set(counts.values())))
    print("저장:", args.out, "/", args.out + ".review.txt")


if __name__ == "__main__":
    main()
