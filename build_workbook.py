# -*- coding: utf-8 -*-
"""
어휘 워크북 자동 생성기
- 입력: 엑셀 (book_name, unit_name, word_no, english, meaning, synonyms, antonyms, example)
- 출력: 유닛(DAY)별로 규칙에 맞는 워크북 HTML -> Chrome 헤드리스로 PDF

유닛 페이지 규칙:
  1. 단어표 (english / 의미 / 예문 / 반의어 / 동의어)   * 품사 열 없음
  2. 4회 쓰기 (english 3번 + 뜻 1번), 순서 랜덤
  [ N-2 유닛 복습 시험 ]  (존재할 때만, 매번 다른 버전)
  [ N-1 유닛 복습 시험 ]  (존재할 때만, 매번 다른 버전)
  5. 예문 빈칸 (의미 + 빈칸 예문), 순서 랜덤
  6. 반의어 고르기 (4지선다)
  7. 동의어 고르기 (4지선다)
"""
import openpyxl
import random
import html
import re
import os
import sys
import subprocess
import argparse

# ---------------------------------------------------------------- 데이터 로드
def load_words(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(c).strip().lower() if c else "" for c in rows[0]]
    idx = {name: header.index(name) for name in header}

    def col(r, *names):
        for n in names:
            if n in idx and idx[n] < len(r):
                v = r[idx[n]]
                return ("" if v is None else str(v)).strip()
        return ""

    books = {}   # book_name -> list of unit dicts (in order)
    for r in rows[1:]:
        if r is None or all(c is None for c in r):
            continue
        book = col(r, "book_name") or "Workbook"
        unit = col(r, "unit_name") or "UNIT"
        w = {
            "no": col(r, "word_no"),
            "english": col(r, "english"),
            "meaning": col(r, "meaning"),
            "synonyms": col(r, "synonyms", "synonym"),
            "antonyms": col(r, "antonyms", "antonym"),
            "example": col(r, "example"),
        }
        if not w["english"]:
            continue
        books.setdefault(book, {})
        books[book].setdefault(unit, [])
        books[book][unit].append(w)

    # dict 순서 유지 -> 유닛 리스트로
    result = {}
    for book, units in books.items():
        result[book] = [(uname, words) for uname, words in units.items()]
    return result


def split_multi(val):
    """'repair, mend' -> ['repair','mend']"""
    if not val:
        return []
    parts = re.split(r"[,/;]", val)
    return [p.strip() for p in parts if p.strip()]


# ---------------------------------------------------------------- HTML 조각
def esc(s):
    return html.escape(s or "")


def blank_out(example, english):
    """예문에서 대상 단어를 빈칸으로. 못 찾으면 문장 뒤에 빈칸 표시."""
    if not example:
        return "_________"
    pat = re.compile(re.escape(english), re.IGNORECASE)
    if pat.search(example):
        return pat.sub('<span class="blank"></span>', example, count=1)
    return esc(example) + ' ( <span class="blank"></span> )'


# ---------------------------------------------------------------- 페이지: 단어표
def page_wordlist(day_name, words):
    rows = []
    for w in words:
        rows.append(f"""
      <tr>
        <td class="c-no">{esc(w['no'])}</td>
        <td class="c-eng">{esc(w['english'])}</td>
        <td class="c-mean">{esc(w['meaning'])}</td>
        <td class="c-ex">{esc(w['example'])}</td>
        <td class="c-ant">{esc(w['antonyms'] or '-')}</td>
        <td class="c-syn">{esc(w['synonyms'] or '-')}</td>
      </tr>""")
    return f"""
  <section class="page">
    {page_head(day_name, "단어 목록", "Word List")}
    <table class="wl">
      <thead>
        <tr><th>No.</th><th>단어</th><th>의미</th><th>예문</th><th>반의어</th><th>동의어</th></tr>
      </thead>
      <tbody>{''.join(rows)}</tbody>
    </table>
  </section>"""


# ---------------------------------------------------------------- 페이지: 4회 쓰기
def page_writing(day_name, words, seed):
    ws = words[:]
    random.Random(seed).shuffle(ws)
    rows = []
    for i, w in enumerate(ws, 1):
        rows.append(f"""
      <tr>
        <td class="w-no">{i}</td>
        <td class="w-eng">{esc(w['english'])}</td>
        <td class="w-mean">{esc(w['meaning'])}</td>
        <td class="w-write"><span class="wl3"></span><span class="wl3"></span><span class="wl3"></span></td>
      </tr>""")
    return f"""
  <section class="page">
    {page_head(day_name, "쓰기 연습", "Write &amp; Remember")}
    <p class="guide">뜻과 단어를 확인하고, 오른쪽 줄에 <b>영어 단어를 3번</b> 따라 쓰세요. <b>따라 쓸 때는 반드시 큰 소리로 단어를 읽으세요.</b></p>
    <table class="wr">
      <thead>
        <tr><th class="wh-no">No.</th><th class="wh-eng">단어</th><th class="wh-mean">뜻</th><th>3번 쓰기</th></tr>
      </thead>
      <tbody>{''.join(rows)}</tbody>
    </table>
  </section>"""


# ---------------------------------------------------------------- 페이지: 예문 빈칸
def page_fillblank(day_name, words, seed, answer=False):
    ws = words[:]
    random.Random(seed).shuffle(ws)
    # 단어 은행(랜덤)
    bank = [w["english"] for w in ws]
    random.Random(seed + 1).shuffle(bank)
    items = []
    for i, w in enumerate(ws, 1):
        sent = blank_out(w["example"], w["english"])
        ans = f'<span class="ans">{esc(w["english"])}</span>' if answer else ""
        items.append(f"""
      <li>
        <span class="fb-sent">{sent}</span> {ans}
        <span class="fb-mean">({esc(w['meaning'])})</span>
      </li>""")
    bank_html = "".join(f'<span class="chip">{esc(b)}</span>' for b in bank)
    return f"""
  <section class="page">
    {page_head(day_name, "예문 빈칸 채우기", "Fill in the Blank")}
    <p class="guide">뜻을 참고하여 빈칸에 알맞은 단어를 <b>단어 은행</b>에서 골라 쓰세요.</p>
    <div class="bank">{bank_html}</div>
    <ol class="fb">{''.join(items)}</ol>
  </section>"""


# ---------------------------------------------------------------- 페이지: 복습 시험
def page_review(cur_day, target_day_name, words, seed, answer=False):
    """이전 유닛 단어↔의미 시험 (섞어서). seed로 매번 다른 버전."""
    rnd = random.Random(seed)
    ws = words[:]
    rnd.shuffle(ws)
    items = []
    for i, w in enumerate(ws, 1):
        ask_meaning = rnd.random() < 0.5  # True: 단어 주고 뜻 쓰기 / False: 뜻 주고 단어 쓰기
        if ask_meaning:
            q = esc(w["english"])
            a = esc(w["meaning"])
            qcls = "rv-eng"
        else:
            q = esc(w["meaning"])
            a = esc(w["english"])
            qcls = "rv-mean"
        ansline = f'<span class="ans">{a}</span>' if answer else ''
        items.append(f"""
      <tr>
        <td class="rv-no">{i}</td>
        <td class="{qcls}">{q}</td>
        <td class="rv-a">{ansline}</td>
      </tr>""")
    return f"""
  <section class="page">
    {page_head(cur_day, f"누적 복습 시험 &middot; {esc(target_day_name)}", "Review Test")}
    <p class="guide">단어는 뜻을, 뜻은 단어를 쓰세요.</p>
    <table class="rv">
      <tbody>{''.join(items)}</tbody>
    </table>
  </section>"""


# ---------------------------------------------------------------- 페이지: 반의어/동의어 고르기
def page_choice(day_name, words, kind, seed, answer=False):
    """kind: 'antonyms' or 'synonyms'. 해당 값 있는 단어만 출제."""
    rnd = random.Random(seed)
    label = "반의어" if kind == "antonyms" else "동의어"
    en_label = "Antonym" if kind == "antonyms" else "Synonym"
    pool = [w for w in words if split_multi(w[kind])]
    all_english = [w["english"] for w in words]
    rnd.shuffle(pool)
    items = []
    for i, w in enumerate(pool, 1):
        corrects = split_multi(w[kind])
        correct = corrects[0]
        # 오답: 같은 유닛의 다른 english (정답/문제어 제외)
        avoid = set([w["english"], correct] + corrects)
        distract_pool = [e for e in all_english if e not in avoid]
        rnd.shuffle(distract_pool)
        options = [correct] + distract_pool[:3]
        # 보기가 4개 안되면 채우기 skip
        rnd.shuffle(options)
        letters = "①②③④"
        opt_html = "".join(
            f'<span class="opt {"opt-ans" if (answer and o==correct) else ""}">{letters[j]} {esc(o)}</span>'
            for j, o in enumerate(options)
        )
        items.append(f"""
      <li>
        <div class="ch-q"><b>{esc(w['english'])}</b> <span class="ch-mean">({esc(w['meaning'])})</span></div>
        <div class="ch-opts">{opt_html}</div>
      </li>""")
    if not items:
        items.append('<li class="empty">이 유닛에는 출제할 단어가 부족합니다.</li>')
    return f"""
  <section class="page">
    {page_head(day_name, f"{label} 고르기", f"Choose the {en_label}")}
    <p class="guide">각 단어의 <b>{label}</b>를 보기에서 고르세요.</p>
    <ol class="ch">{''.join(items)}</ol>
  </section>"""


# ---------------------------------------------------------------- 공통 헤더
def page_head(day_name, title_ko, title_en):
    return f"""
    <div class="phead">
      <div class="ph-day">{esc(day_name)}</div>
      <div class="ph-title"><span class="pt-ko">{title_ko}</span><span class="pt-en">{title_en}</span></div>
      <div class="ph-name">학습한 날짜 : <span class="nline"></span></div>
    </div>"""


# ---------------------------------------------------------------- CSS
CSS = """
:root{
  --teal:#32bfb6; --teal-lt:#94e1da; --ink:#2f302f; --sand:#dfc0aa;
  --teal-bg:#eafaf8; --teal-bg2:#f4fbfa; --sand-bg:#f7efe6;
  --line:#d9d9d6; --muted:#8a8f8d;
}
@page { size: A4; margin: 13mm 11mm 13mm 11mm; }
* { box-sizing: border-box; }
body { font-family:'Pretendard','Malgun Gothic',sans-serif; color:var(--ink); margin:0; font-size:13px; }
.page { page-break-after: always; }
.page:last-child { page-break-after: auto; }

.phead { display:flex; align-items:flex-end; justify-content:space-between;
  border-bottom:3px solid var(--teal); padding-bottom:6px; margin-bottom:11px; }
.ph-day { font-size:18px; font-weight:800; color:var(--teal); letter-spacing:.5px; }
.ph-title { text-align:center; }
.pt-ko { display:block; font-size:19px; font-weight:800; color:var(--ink); }
.pt-en { display:block; font-size:10px; color:var(--sand); letter-spacing:1.5px; text-transform:uppercase; font-weight:700; }
.ph-name { font-size:12px; color:var(--muted); }
.nline { display:inline-block; width:120px; border-bottom:1px solid var(--sand); margin-left:4px; }

/* 페이지 로고 (매 페이지 우하단 반복) */
.pagelogo { position:fixed; bottom:1mm; right:2mm; width:30mm; height:auto; opacity:.95; }
.guide { background:var(--teal-bg); border-left:4px solid var(--teal); padding:7px 11px; margin:0 0 12px;
  font-size:12.5px; color:var(--ink); border-radius:3px; }
.guide b { color:var(--teal); }

/* 단어표 */
table.wl { width:100%; border-collapse:collapse; }
table.wl th { background:var(--teal); color:#fff; font-size:12px; padding:8px 6px; font-weight:700; }
table.wl td { border:1px solid var(--line); padding:7px 7px; vertical-align:top; font-size:12.5px; }
table.wl tr:nth-child(even) td { background:var(--teal-bg2); }
.c-no { text-align:center; color:var(--muted); width:26px; }
.c-eng { font-weight:800; color:var(--ink); width:82px; }
.c-mean { width:150px; }
.c-ex { color:#55605d; font-style:italic; }
.c-ant, .c-syn { width:80px; color:#6b6f6d; text-align:center; }

/* 쓰기 (뜻 먼저 -> 가로 쓰기줄) */
table.wr { width:100%; border-collapse:collapse; }
table.wr th { background:var(--teal); color:#fff; padding:8px; font-size:12px; font-weight:700; }
table.wr th.wh-no{ width:38px; } table.wr th.wh-eng{ width:120px; } table.wr th.wh-mean{ width:190px; }
table.wr td { border-bottom:1px solid var(--line); height:38px; vertical-align:middle; }
.w-no { text-align:center; color:var(--muted); font-size:12px; }
.w-eng { font-weight:800; padding-left:12px; color:var(--teal); font-size:14px; border-left:1px dashed var(--teal-lt); }
.w-mean { padding-left:10px; color:var(--ink); font-size:12.5px; border-left:1px dashed var(--teal-lt); }
.w-write { padding:0 10px; white-space:nowrap; border-left:1px dashed var(--teal-lt); }
.wl3 { display:inline-block; width:29%; border-bottom:1.4px solid var(--line); margin:0 1.5%; height:20px; }

/* 예문 빈칸 */
.bank { border:1.5px dashed var(--teal-lt); background:var(--teal-bg); border-radius:6px; padding:8px 10px;
  margin-bottom:11px; line-height:1.95; }
.chip { display:inline-block; background:#fff; border:1px solid var(--teal-lt); border-radius:13px;
  padding:2px 11px; margin:2px 3px; font-size:12px; font-weight:600; color:var(--ink); }
ol.fb { margin:0; padding-left:22px; }
ol.fb li { margin-bottom:17px; line-height:1.7; }
.fb-sent { font-size:13px; }
.fb-mean { font-size:11px; color:var(--muted); margin-left:6px; }
.blank { display:inline-block; min-width:120px; border-bottom:1.6px solid var(--ink); }
.ans { color:#c47a3d; font-weight:800; margin-left:4px; }

/* 복습 시험 */
table.rv { width:100%; border-collapse:collapse; }
table.rv td { border-bottom:1px solid var(--line); padding:9px 7px; font-size:13px; }
.rv-no { width:28px; text-align:center; color:var(--muted); }
.rv-eng { font-weight:800; width:180px; color:var(--teal); }
.rv-mean { width:180px; }
.rv-a { }
.wline { display:inline-block; width:90%; border-bottom:1px solid var(--line); }

/* 고르기 */
ol.ch { margin:0; padding-left:22px; }
ol.ch li { margin-bottom:7px; }
.ch-q { font-size:13px; margin-bottom:3px; }
.ch-q b { color:var(--teal); }
.ch-mean { color:var(--muted); font-size:11px; font-weight:400; }
.ch-opts { display:flex; gap:7px; flex-wrap:wrap; }
.opt { border:1px solid var(--line); border-radius:5px; padding:4px 12px; font-size:12.5px; }
.opt-ans { background:var(--sand-bg); border-color:var(--sand); color:#b06a2c; font-weight:800; }
.empty { color:var(--muted); }
"""


# ---------------------------------------------------------------- 유닛 페이지 조립
def build_unit_pages(units, i, answer=False):
    """units: [(name, words)]. i: 0-based 현재 유닛 index."""
    name, words = units[i]
    base = (i + 1) * 1000
    pages = []
    pages.append(page_wordlist(name, words))
    pages.append(page_writing(name, words, seed=base + 1))
    # 누적 복습: N-2, N-1
    if i - 2 >= 0:
        tname, twords = units[i - 2]
        pages.append(page_review(name, tname, twords, seed=base + 200 + i, answer=answer))
    if i - 1 >= 0:
        tname, twords = units[i - 1]
        pages.append(page_review(name, tname, twords, seed=base + 300 + i, answer=answer))
    pages.append(page_fillblank(name, words, seed=base + 4, answer=answer))
    pages.append(page_choice(name, words, "antonyms", seed=base + 5, answer=answer))
    pages.append(page_choice(name, words, "synonyms", seed=base + 6, answer=answer))
    return pages


def font_face_css():
    """설치된 Pretendard(ttf)를 file:// 로 임베드해 헤드리스 Chrome에서 확실히 로드."""
    dirs = [
        os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Windows\Fonts"),
        r"C:\Windows\Fonts",
    ]
    weights = {"Regular": 400, "Medium": 500, "SemiBold": 600, "Bold": 700, "ExtraBold": 800}
    rules = []
    for name, wt in weights.items():
        for d in dirs:
            fp = os.path.join(d, f"Pretendard-{name}.ttf")
            if os.path.exists(fp):
                url = "file:///" + fp.replace("\\", "/")
                rules.append(
                    "@font-face{font-family:'Pretendard';font-style:normal;"
                    f"font-weight:{wt};src:url('{url}') format('truetype');}}"
                )
                break
    return "\n".join(rules)


def logo_datauri():
    """source/logo.png -> data URI. 없으면 빈 문자열."""
    import base64
    here = os.path.dirname(os.path.abspath(__file__))
    fp = os.path.join(here, "source", "logo.png")
    if not os.path.exists(fp):
        return ""
    with open(fp, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return "data:image/png;base64," + b64


def build_html(book_name, units, day_from, day_to, answer=False):
    body = []
    for i in range(day_from - 1, min(day_to, len(units))):
        body.extend(build_unit_pages(units, i, answer=answer))
    title = f"{book_name} 워크북" + (" (정답)" if answer else "")
    logo = logo_datauri()
    logo_html = f'<img class="pagelogo" src="{logo}" alt="logo">' if logo else ""
    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<title>{esc(title)}</title><style>{font_face_css()}
{CSS}</style></head>
<body>{logo_html}{''.join(body)}</body></html>"""


# ---------------------------------------------------------------- PDF 변환
def find_chrome():
    cands = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for c in cands:
        if os.path.exists(c):
            return c
    return None


def html_to_pdf(html_path, pdf_path):
    chrome = find_chrome()
    if not chrome:
        raise RuntimeError("Chrome/Edge를 찾을 수 없습니다.")
    import tempfile, shutil
    url = "file:///" + os.path.abspath(html_path).replace("\\", "/")
    out = os.path.abspath(pdf_path).replace("\\", "/")
    profile = tempfile.mkdtemp(prefix="vocawb_")  # 새 프로필 -> 실행 중인 Chrome과 독립
    try:
        r = subprocess.run([
            chrome, "--headless=new", "--disable-gpu", "--no-pdf-header-footer",
            f"--user-data-dir={profile}",
            f"--print-to-pdf={out}", url
        ], timeout=180, capture_output=True, encoding="utf-8", errors="replace")
        log = (r.stderr or "") + (r.stdout or "")
        if "written to file" not in log:
            raise RuntimeError(
                "PDF 생성 실패 (파일이 열려 있어 잠겨있을 수 있습니다 — 뷰어를 닫고 다시 시도하세요).\n"
                + log[-400:])
    finally:
        shutil.rmtree(profile, ignore_errors=True)


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--from", dest="dfrom", type=int, default=1)
    ap.add_argument("--to", dest="dto", type=int, default=3)
    ap.add_argument("--out", default="output")
    ap.add_argument("--answer", action="store_true", help="정답 버전도 생성")
    args = ap.parse_args()

    books = load_words(args.xlsx)
    os.makedirs(args.out, exist_ok=True)
    for book_name, units in books.items():
        safe = re.sub(r"[\\/:*?\"<>|]", "_", book_name)
        # 학생용
        h = build_html(book_name, units, args.dfrom, args.dto, answer=False)
        hp = os.path.join(args.out, f"{safe}_DAY{args.dfrom}-{args.dto}.html")
        with open(hp, "w", encoding="utf-8") as f:
            f.write(h)
        pp = hp[:-5] + ".pdf"
        html_to_pdf(hp, pp)
        print("생성:", pp)
        if args.answer:
            ha = build_html(book_name, units, args.dfrom, args.dto, answer=True)
            hap = os.path.join(args.out, f"{safe}_DAY{args.dfrom}-{args.dto}_정답.html")
            with open(hap, "w", encoding="utf-8") as f:
                f.write(ha)
            ppa = hap[:-5] + ".pdf"
            html_to_pdf(hap, ppa)
            print("생성:", ppa)


if __name__ == "__main__":
    main()
