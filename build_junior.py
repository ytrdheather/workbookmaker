# -*- coding: utf-8 -*-
"""찐 예비 실력자 시리즈 — 초등 저학년(1~3학년) 어휘 워크북 생성기.

기존 `build_workbook.py`(찐 실력자, 중등 이상)와 별개의 템플릿이다.
중등용 지면(누적 복습 30문항·동의어/반의어 고르기·예문 빈칸)은 저학년에 맞지 않고,
초등 기초어는 동의어/반의어가 거의 없어 고르기 지면 자체가 성립하지 않는다.

기존 코드는 건드리지 않고 아래만 가져다 쓴다(엑셀 읽기·PDF 변환·폰트·로고·분권).
따라서 기존 13종 교재의 출력에는 아무 영향이 없다.

DAY당 4장 (유닛당 8~11단어 기준):
  1. 단어 만나기   — 단어·뜻 + 소리내어 읽기 3회 체크
  2. 따라 쓰기     — 4선 위 회색 글씨 덧쓰기 + 빈 4선에 4번 더 쓰기
  3. 확인하기      — 선 잇기 + 3지선다(뜻 보고 영어 고르기)
  4. 놀이로 익히기 — 낱말 찾기 + 첫 글자 힌트 채우기

사용:
  python build_junior.py "excel/초등영어 단어가 된다_1.xlsx" --split 12 --answer --out "output_new/..."
"""
import argparse
import os
import random
import re
import string
import sys

import build_workbook as B          # 엑셀 읽기 / PDF 변환 / 폰트 / 로고 / 분권 재사용
from build_workbook import esc

GRID = 11          # 낱말 찾기 격자 크기
TRACE_REPEAT = 2     # 따라쓰기(회색 글씨) 반복 횟수
WRITE_SLOTS = 4      # 따라쓰기 뒤 빈칸에 더 쓰는 횟수
TRACE_COLS = 3       # 한 줄에 놓는 칸 수(적을수록 칸이 넓어 글자가 커진다)
TRACE_PER_PAGE = 5   # 따라쓰기 한 장에 담는 단어 수
MATCH_GROUP = 5      # 선 잇기 한 묶음의 문항 수
REVIEW_EVERY = 4     # 몇 DAY마다 누적 복습을 넣는가
REVIEW_SPAN = 2      # 복습 한 장이 다루는 DAY 수


# ---------------------------------------------------------------- 공통 조각
def page_head(day_name, title_ko, title_en):
    return f"""
    <div class="phead">
      <div class="ph-day">{esc(day_name)}</div>
      <div class="ph-title"><span class="pt-ko">{title_ko}</span><span class="pt-en">{title_en}</span></div>
      <div class="ph-name">이름 : <span class="nline"></span></div>
    </div>"""


def alpha_only(w):
    """낱말 찾기·첫 글자 채우기에 쓸 알파벳만 남긴 형태(공백·기호 제거)."""
    return re.sub(r"[^A-Za-z]", "", w)


# ---------------------------------------------------------------- 1. 단어 만나기
def page_meet(day_name, words):
    """단어·뜻을 크게 보여주고, 소리내어 3회 읽으며 체크하게 한다."""
    rows = []
    has_ex = any(w.get("example") for w in words)
    for i, w in enumerate(words, 1):
        boxes = "".join('<span class="rd-box"></span>' for _ in range(3))
        span = ' rowspan="2"' if w.get("example") else ""
        rows.append(
            f'<tr class="mt-a"><td class="mt-no"{span}>{i}</td>'
            f'<td class="mt-eng">{esc(w["english"])}</td>'
            f'<td class="mt-mean">{esc(w["meaning"])}</td>'
            f'<td class="mt-read"{span}>{boxes}</td></tr>')
        if w.get("example"):
            # 뜻만으로는 쓰임을 알기 어려운 단어(arm=팔/무장시키다)를 문장으로 잡아 준다.
            rows.append(f'<tr class="mt-b"><td class="mt-ex" colspan="2">'
                        f'{esc(w["example"])}</td></tr>')
    guide = ("단어와 <b>문장</b>을 <b>소리 내어</b> 읽어요. 한 번 읽을 때마다 <b>□에 색칠</b>해요."
             if has_ex else
             "단어를 <b>소리 내어</b> 읽어요. 한 번 읽을 때마다 <b>□에 색칠</b>해요.")
    return f"""
  <section class="page">
    {page_head(day_name, "단어 만나기", "Meet the Words")}
    <p class="guide">{guide} 세 번 다 읽으면 다음 장으로!</p>
    <table class="mt">
      <colgroup><col style="width:34px"><col style="width:40%"><col><col style="width:96px"></colgroup>
      <thead><tr><th>No</th><th>단어</th><th>뜻</th><th>읽기 3번</th></tr></thead>
      <tbody>{''.join(rows)}</tbody>
    </table>
  </section>"""


# ---------------------------------------------------------------- 2. 따라 쓰기
def _fourline(content_html, cls=""):
    """알파벳 4선지 한 줄. 윗선·중간점선·기준선·아랫선 위에 글자를 얹는다."""
    return (f'<div class="fl {cls}">'
            f'<span class="fl-l1"></span><span class="fl-l2"></span>'
            f'<span class="fl-l3"></span><span class="fl-l4"></span>'
            f'<span class="fl-tx">{content_html}</span></div>')


def page_trace(day_name, words):
    """회색 글씨를 덧쓴 뒤, 빈 4선에 스스로 4번 더 쓴다. -> 페이지 리스트

    한 줄에 3칸씩 두 줄(덧쓰기 2 + 빈칸 4). 한 줄에 6칸을 몰아넣으면 칸이 좁아
    글자가 작아지므로, 칸 수를 줄이고 대신 지면을 나눈다(저학년은 크게 써야 한다)."""
    return [_trace_page(day_name, words[s:s + TRACE_PER_PAGE], s, ci)
            for ci, s in enumerate(range(0, len(words), TRACE_PER_PAGE))]


def _trace_page(day_name, words, offset, ci):
    blocks = []
    for i, w in enumerate(words, offset + 1):
        eng = esc(w["english"])
        ghost = f'<span class="tr-ghost">{eng}</span>'
        r1 = "".join(_fourline(ghost) for _ in range(TRACE_REPEAT))
        r1 += "".join(_fourline("") for _ in range(TRACE_COLS - TRACE_REPEAT))
        r2 = "".join(_fourline("") for _ in range(TRACE_COLS))
        blocks.append(f"""
      <div class="tr-item">
        <div class="tr-head"><span class="tr-no">{i}</span>
          <span class="tr-eng">{eng}</span>
          <span class="tr-mean">{esc(w['meaning'])}</span></div>
        <div class="tr-rows">{r1}</div>
        <div class="tr-rows">{r2}</div>
      </div>""")
    return f"""
  <section class="page">
    {page_head(day_name, f"따라 쓰기{' (계속)' if ci else ''}", "Trace &amp; Write")}
    <p class="guide"><b>연한 글씨</b> 위에 덧써요. 그 다음 <b>빈 줄에 혼자서</b> 써 봐요.
      쓰면서 소리 내어 읽으면 더 잘 외워져요.</p>
    <div class="tr-wrap">{''.join(blocks)}</div>
  </section>"""


# ---------------------------------------------------------------- 3. 확인하기
def page_check(day_name, words, seed, answer=False):
    """선 잇기 + 3지선다. 단어가 적어 한 장에 둘 다 넣는다."""
    rnd = random.Random(seed)
    # --- 선 잇기: 5개씩 묶음으로. 10개를 한 번에 이으면 저학년에겐 선이 너무 얽힌다.
    groups = []
    for g in range(0, len(words), MATCH_GROUP):
        chunk = words[g:g + MATCH_GROUP]
        right = list(chunk)
        rnd.shuffle(right)
        ml, mr = [], []
        for w in chunk:
            tag = f'<span class="ans">{right.index(w) + 1}</span>' if answer else ""
            ml.append(f'<li><span class="mt-dot"></span>{esc(w["english"])}{tag}</li>')
        for w in right:
            mr.append(f'<li><span class="mt-dot"></span>{esc(w["meaning"])}</li>')
        groups.append(f'<div class="match">'
                      f'<ul class="mt-l">{"".join(ml)}</ul>'
                      f'<ul class="mt-r">{"".join(mr)}</ul></div>')

    # --- 3지선다: 뜻을 주고 영어를 고른다. 오답은 같은 DAY의 다른 단어에서.
    quiz = []
    pool = [w["english"] for w in words]
    order = list(words)
    rnd.shuffle(order)
    for i, w in enumerate(order, 1):
        others = [e for e in pool if e != w["english"]]
        rnd.shuffle(others)
        opts = [w["english"]] + others[:2]
        rnd.shuffle(opts)
        cells = []
        for k, o in enumerate(opts, 1):
            hit = answer and o == w["english"]
            cells.append(f'<span class="op{" op-a" if hit else ""}">'
                         f'<b>{k}</b> {esc(o)}</span>')
        quiz.append(f'<li><span class="q-mean">{esc(w["meaning"])}</span>'
                    f'<span class="q-ops">{"".join(cells)}</span></li>')

    return f"""
  <section class="page">
    {page_head(day_name, "확인하기", "Check It")}
    <div class="sec-t">1. 알맞은 것끼리 <b>선으로 이어요</b>.</div>
    <div class="match-wrap">{''.join(groups)}</div>
    <div class="sec-t">2. 뜻에 맞는 <b>영어 단어</b>를 골라 번호를 쓰세요.</div>
    <ol class="quiz">{''.join(quiz)}</ol>
  </section>"""


# ---------------------------------------------------------------- 4. 놀이로 익히기
def _wordsearch(words, rnd, size=GRID):
    """가로(→)·세로(↓)로만 배치하는 낱말 찾기. 못 넣은 단어는 돌려준다."""
    grid = [[None] * size for _ in range(size)]
    placed, failed = [], []
    for w in sorted((alpha_only(x["english"]).upper() for x in words),
                    key=len, reverse=True):
        if not w or len(w) > size:
            failed.append(w)
            continue
        spots = []
        for r in range(size):
            for c in range(size):
                for dr, dc in ((0, 1), (1, 0)):
                    er, ec = r + dr * (len(w) - 1), c + dc * (len(w) - 1)
                    if er >= size or ec >= size:
                        continue
                    if all(grid[r + dr * k][c + dc * k] in (None, w[k])
                           for k in range(len(w))):
                        spots.append((r, c, dr, dc))
        if not spots:
            failed.append(w)
            continue
        r, c, dr, dc = rnd.choice(spots)
        for k, ch in enumerate(w):
            grid[r + dr * k][c + dc * k] = ch
        placed.append((w, r, c, dr, dc))
    return grid, placed, failed


def page_play(day_name, words, seed, answer=False):
    """낱말 찾기 + 첫 글자 힌트 채우기."""
    rnd = random.Random(seed)
    grid, placed, failed = _wordsearch(words, rnd)
    hit = {(r + dr * k, c + dc * k)
           for w, r, c, dr, dc in placed for k in range(len(w))}
    cells = []
    for r, row in enumerate(grid):
        for c, ch in enumerate(row):
            mark = " ws-a" if answer and (r, c) in hit else ""
            cells.append(f'<span class="ws-c{mark}">{ch or rnd.choice(string.ascii_uppercase)}</span>')
    found = ", ".join(esc(w) for w, *_ in placed)

    # 첫 글자 힌트 — a____ 처럼 첫 글자만 주고 나머지 칸 수를 알려 준다
    items = []
    order = list(words)
    rnd.shuffle(order)
    for i, w in enumerate(order, 1):
        raw = w["english"]
        a = alpha_only(raw)
        if not a:
            continue
        rest = ("".join(f'<span class="hb-c">{c}</span>' for c in a[1:])
                if answer else "".join('<span class="hb-c"></span>' for _ in a[1:]))
        items.append(f'<li><span class="hb-mean">{esc(w["meaning"])}</span>'
                     f'<span class="hb-w"><span class="hb-c hb-1">{esc(a[0])}</span>{rest}</span></li>')

    note = ("" if not failed else
            f'<p class="ws-note">※ 격자에 넣지 못한 단어: {esc(", ".join(failed))}</p>')
    return f"""
  <section class="page">
    {page_head(day_name, "놀이로 익히기", "Word Play")}
    <div class="sec-t">1. 숨어 있는 단어를 찾아 <b>동그라미</b> 하세요. (→ 가로, ↓ 세로)</div>
    <div class="ws">{''.join(cells)}</div>
    <p class="ws-list"><b>찾을 단어</b> &nbsp; {found}</p>{note}
    <div class="sec-t">2. <b>첫 글자</b>를 보고 나머지 칸을 채우세요.</div>
    <ol class="hb">{''.join(items)}</ol>
  </section>"""


# ---------------------------------------------------------------- 누적 복습
def page_review(label, words, seed, answer=False):
    """2개 DAY를 묶은 복습. 앞쪽은 영어를 보고 뜻을, 뒤쪽은 뜻을 보고 영어를 쓴다.

    뜻→영어에 단어 은행을 주면 보기 수와 문항 수가 같아 옮겨 적기·소거법으로 풀려
    시험이 되지 않는다. 대신 첫 글자와 글자 수만 알려 준다."""
    rnd = random.Random(seed)
    order = list(words)
    rnd.shuffle(order)
    half = (len(order) + 1) // 2
    to_ko, to_en = order[:half], order[half:]

    ko = []
    for w in to_ko:
        fill = f'<span class="rv-w">{esc(w["meaning"])}</span>' if answer else ""
        ko.append(f'<li><span class="rv-q">{esc(w["english"])}</span>'
                  f'<span class="rv-line">{fill}</span></li>')

    en = []
    for w in to_en:
        a = alpha_only(w["english"])
        hint = ("%s %s" % (a[0], "_" * max(len(a) - 1, 1))) if a else ""
        ghost = (f'<span class="rv-a">{esc(w["english"])}</span>') if answer else ""
        en.append(f'<li><span class="rv-q">{esc(w["meaning"])}</span>'
                  f'<span class="rv-hint">{esc(hint)}</span>'
                  f'{_fourline(ghost, "rv-fl")}</li>')

    return f"""
  <section class="page">
    {page_head(label, "누적 복습", "Review Test")}
    <div class="sec-t">1. 영어를 보고 <b>뜻</b>을 쓰세요.</div>
    <ol class="rv rv-ko">{''.join(ko)}</ol>
    <div class="sec-t">2. 뜻을 보고 <b>영어</b>를 쓰세요.
      <span class="sec-n">첫 글자와 글자 수가 힌트예요.</span></div>
    <ol class="rv rv-en">{''.join(en)}</ol>
  </section>"""


# ---------------------------------------------------------------- CSS
CSS = """
:root{
  --teal:#32bfb6; --teal-lt:#94e1da; --ink:#2f302f; --sand:#dfc0aa;
  --teal-bg:#eafaf8; --teal-bg2:#f4fbfa; --sand-bg:#f7efe6;
  --line:#d9d9d6; --muted:#8a8f8d; --ghost:#c9cfce;
}
@page { size: A4; margin: 13mm 11mm 13mm 11mm; }
* { box-sizing: border-box; }
body { font-family:'Pretendard','Malgun Gothic',sans-serif; color:var(--ink); margin:0; font-size:14px; }
.page { page-break-after: always; position:relative; min-height:262mm;
  padding-bottom:40px; }   /* 우하단 로고 자리 */
/* 지면을 .unit 으로 감싸므로 :last-child 는 '유닛의 마지막 장'을 뜻하게 된다.
   그대로 두면 유닛마다 페이지 나눔이 빠져 다음 유닛이 같은 장에 이어붙는다. */
.unit:last-child .page:last-child { page-break-after: auto; }

.phead { display:flex; align-items:flex-end; justify-content:space-between;
  border-bottom:3px solid var(--teal); padding-bottom:6px; margin-bottom:12px; }
.ph-day { font-size:20px; font-weight:800; color:var(--teal); letter-spacing:.5px; }
.ph-title { text-align:center; }
.pt-ko { display:block; font-size:21px; font-weight:800; }
.pt-en { display:block; font-size:10px; color:var(--sand); letter-spacing:1.5px;
  text-transform:uppercase; font-weight:700; }
.ph-name { font-size:12px; color:var(--muted); }
.nline { display:inline-block; width:96px; border-bottom:1px solid var(--line); }

.guide { background:var(--teal-bg); border-left:4px solid var(--teal);
  padding:9px 12px; margin:0 0 14px; font-size:13.5px; line-height:1.6; border-radius:0 6px 6px 0; }
.sec-t { font-size:15px; font-weight:800; margin:16px 0 9px; color:var(--ink); }
.sec-t:first-of-type { margin-top:4px; }
.pagelogo { position:fixed; right:0; bottom:0; width:118px; opacity:.95; }

/* 1. 단어 만나기 */
.mt { width:100%; border-collapse:collapse; }
.mt th { background:var(--teal); color:#fff; font-size:12.5px; font-weight:700;
  padding:8px 6px; border:1px solid var(--teal); }
.mt td { border:1px solid var(--line); padding:0 8px; height:var(--mth); vertical-align:middle; }
/* 문장이 있으면 단어행+문장행이 한 칸을 이룬다(테두리를 안 끊는다) */
.mt tr.mt-a td { border-bottom:none; }
.mt tr.mt-b td { border-top:none; height:var(--mtxh); }
.mt-ex { font-size:15px; color:#4a7f7a; font-style:italic; padding-left:16px !important; }
.mt-no { text-align:center; color:var(--muted); font-size:13px; }
.mt-eng { font-size:21px; font-weight:700; color:var(--teal); letter-spacing:.3px; }
.mt-mean { font-size:15px; }
.mt-read { text-align:center; white-space:nowrap; }
.rd-box { display:inline-block; width:20px; height:20px; margin:0 3px;
  border:1.6px solid var(--teal-lt); border-radius:5px; vertical-align:middle; }

/* 2. 따라 쓰기 — 알파벳 4선지 */
.tr-wrap { display:flex; flex-direction:column; gap:var(--trgap); }
.tr-item { border:1px solid var(--line); border-radius:7px; padding:5px 9px 7px; }
.tr-head { display:flex; align-items:baseline; gap:8px; margin-bottom:5px; }
.tr-no { display:inline-block; min-width:18px; height:18px; line-height:18px; text-align:center;
  background:var(--teal); color:#fff; border-radius:50%; font-size:11px; font-weight:700; }
.tr-eng { font-size:15px; font-weight:800; color:var(--teal); }
.tr-mean { font-size:12px; color:var(--muted); }
/* 6칸(덧쓰기 2 + 빈칸 4)을 한 줄에. 접히면 지면을 넘기므로 nowrap. */
.tr-rows { display:flex; flex-wrap:nowrap; gap:8px; margin-top:5px; }
.fl { position:relative; height:var(--flh); flex:1 1 0; min-width:0; }
.fl-l1,.fl-l2,.fl-l3,.fl-l4 { position:absolute; left:0; right:0; }
.fl-l1 { top:0; border-top:1px solid var(--teal-lt); }
.fl-l2 { top:33.3%; border-top:1px dashed var(--teal-lt); }
.fl-l3 { top:66.6%; border-top:1.6px solid var(--teal); }
.fl-l4 { bottom:0; border-top:1px solid var(--teal-lt); }
.fl-tx { position:absolute; left:5px; right:2px; top:0; overflow:hidden; height:66.6%; display:flex; align-items:flex-end;
  font-size:var(--flf); font-weight:600; line-height:1; letter-spacing:1px; }
.tr-ghost { color:var(--ghost); }

/* 3. 확인하기 — 선 잇기 */
.match-wrap { display:flex; flex-direction:column; gap:9px; }
.match { display:flex; justify-content:space-between; gap:26px;
  border:1px solid var(--line); border-radius:7px; padding:9px 16px; }
.match ul { list-style:none; margin:0; padding:0; flex:1; }
.match li { display:flex; align-items:center; gap:9px; height:var(--mh); font-size:15px; }
.mt-l li { justify-content:flex-start; }
.mt-r li { justify-content:flex-start; }
.mt-dot { width:8px; height:8px; border-radius:50%; background:var(--teal); flex:none; }
.mt-l .mt-dot { order:2; margin-left:auto; }   /* 왼쪽 단은 글씨 뒤(가운데 쪽) */
.ans { color:#e2564d; font-weight:800; margin-left:6px; font-size:13px; }

/* 3. 확인하기 — 3지선다 */
.quiz { margin:0; padding-left:20px; }
.quiz li { margin-bottom:var(--qgap); font-size:14px; }
.q-mean { display:inline-block; min-width:132px; font-weight:600; }
.q-ops { display:inline-block; }
.op { display:inline-block; border:1px solid var(--line); border-radius:14px;
  padding:3px 11px; margin-right:7px; font-size:13.5px; }
.op b { color:var(--teal); margin-right:3px; }
.op-a { border-color:#e2564d; color:#e2564d; font-weight:700; }
.op-a b { color:#e2564d; }

/* 4. 놀이로 익히기 — 낱말 찾기 */
.ws { display:grid; grid-template-columns:repeat(var(--gs), 1fr); gap:2px;
  width:var(--wsw); margin:0 auto; }
.ws-c { aspect-ratio:1/1; display:flex; align-items:center; justify-content:center;
  border:1px solid var(--line); border-radius:4px; font-size:15px; font-weight:600;
  background:#fff; }
.ws-a { background:var(--teal-bg); border-color:var(--teal); color:var(--teal); }
.ws-list { text-align:center; font-size:13px; margin:9px 0 0; color:var(--ink); }
.ws-note { text-align:center; font-size:11px; color:var(--muted); margin:4px 0 0; }

/* 누적 복습 */
.rv { margin:0; padding-left:22px; }
.rv li { display:flex; align-items:center; gap:12px; margin-bottom:var(--rvgap); font-size:14px; }
.rv-q { min-width:118px; font-weight:600; }
.rv-ko .rv-q { color:var(--teal); font-size:16px; }
.rv-line { flex:1; border-bottom:1px solid var(--line); height:24px; padding-left:8px;
  display:flex; align-items:center; }
.rv-w { color:#e2564d; font-weight:600; font-size:13.5px; }
.rv-fl { flex:1; height:var(--rvflh) !important; }
.rv-a { color:#e2564d; }
.rv-hint { min-width:78px; color:var(--muted); font-size:13px; letter-spacing:1px; }
.sec-n { font-weight:400; font-size:12px; color:var(--muted); margin-left:6px; }

/* 4. 첫 글자 힌트 */
.hb { margin:0; padding-left:20px; }
.hb li { display:flex; align-items:center; gap:12px; margin-bottom:var(--hgap); font-size:14px; }
.hb-mean { min-width:132px; font-weight:600; }
.hb-w { display:flex; gap:4px; }
.hb-c { width:24px; height:30px; border:1px solid var(--line); border-radius:4px;
  display:flex; align-items:center; justify-content:center; font-size:15px; font-weight:600;
  color:#e2564d; }
.hb-1 { background:var(--teal-bg); border-color:var(--teal); color:var(--teal); font-weight:800; }
"""


# ---------------------------------------------------------------- 지면 조립
def build_unit_pages(day_name, words, seed, answer=False):
    return ([page_meet(day_name, words)]
            + page_trace(day_name, words)
            + [page_check(day_name, words, seed + 1, answer),
               page_play(day_name, words, seed + 2, answer)])


def density(words):
    """단어 수에 맞춰 행 높이·글자를 정한다(유닛당 8~11단어)."""
    n = max(len(words), 1)
    longest = max((len(alpha_only(w["english"])) for w in words), default=6)
    # 따라쓰기: 한 단어가 한 줄, 그 안에 6칸. 칸 폭은 (본문폭-여백)/6 ≈ 112px 고정이므로
    # 긴 단어는 글자만 줄인다. 행 높이는 단어 수에 맞춰 지면을 꽉 채운다.
    # 따라쓰기는 한 장에 5단어 × 2줄이라 행을 크게 잡을 수 있다.
    # 칸 폭 = (본문폭 - 여백)/3 ≈ 230px 이므로 긴 단어도 글자를 키울 수 있다.
    return {
        "--flh": "%dpx" % 44,
        "--trgap": "%dpx" % 11,
        "--flf": "%dpx" % (28 if longest <= 8 else (24 if longest <= 11 else 20)),
        "--mh": "%dpx" % (38 if n <= 9 else (34 if n <= 10 else 29)),
        "--qgap": "%dpx" % (17 if n <= 9 else (14 if n <= 10 else 9)),
        "--hgap": "%dpx" % (9 if n <= 8 else 7),
        "--mth": "%dpx" % (34 if n <= 10 else 31),
        "--mtxh": "%dpx" % (28 if n <= 10 else 25),
        "--rvgap": "%dpx" % (10 if n <= 18 else 7),
        "--rvflh": "%dpx" % (34 if n <= 18 else 30),
        "--gs": str(GRID),
        # 낱말 찾기 격자와 첫글자 목록이 한 장을 나눠 쓰므로, 단어가 많으면 격자를 줄인다.
        "--wsw": "%dmm" % (118 if n <= 8 else (112 if n <= 9 else (106 if n <= 10 else 100))),
    }


def _wrap(pages, words):
    var = ";".join(f"{k}:{v}" for k, v in density(words).items())
    return f'<div class="unit" style="{var}">{"".join(pages)}</div>'


def build_html(book_name, units, day_from, day_to, answer=False, title_suffix=""):
    body = []
    lo, hi = day_from - 1, min(day_to, len(units))
    for i in range(lo, hi):
        name, words = units[i]
        body.append(_wrap(build_unit_pages(name, words, (i + 1) * 1000, answer), words))
        # 4DAY가 끝날 때마다 그 4DAY를 2유닛씩 나눠 복습한다(모든 유닛이 한 번씩 복습됨).
        if (i - lo + 1) % REVIEW_EVERY == 0:
            for s in range(i - REVIEW_EVERY + 1, i + 1, REVIEW_SPAN):
                grp = units[s:s + REVIEW_SPAN]
                merged = [w for _, ws in grp for w in ws]
                label = "%s - %s" % (grp[0][0], grp[-1][0].replace("DAY ", ""))
                body.append(_wrap([page_review(label, merged, (i + 1) * 7000 + s, answer)],
                                  merged))
    title = f"{book_name}{title_suffix} 워크북" + (" (정답)" if answer else "")
    logo = B.logo_datauri()
    logo_html = f'<img class="pagelogo" src="{logo}" alt="logo">' if logo else ""
    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<title>{esc(title)}</title><style>{B.font_face_css()}
{CSS}</style></head>
<body>{logo_html}{''.join(body)}</body></html>"""


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="찐 예비 실력자 — 초등 저학년 워크북 생성기")
    ap.add_argument("xlsx")
    ap.add_argument("--out", default="output_junior")
    ap.add_argument("--split", type=int, default=0, help="권당 DAY 수(0=한 권)")
    ap.add_argument("--answer", action="store_true", help="정답본도 생성")
    ap.add_argument("--from-day", type=int, default=1)
    ap.add_argument("--to-day", type=int, default=999)
    args = ap.parse_args()

    books = B.load_words(args.xlsx)
    os.makedirs(args.out, exist_ok=True)
    for book_name, units in books.items():
        units = units[args.from_day - 1:args.to_day]
        total = len(units)
        ranges = (B.volume_ranges(total, max(1, -(-total // args.split)))
                  if args.split else [(1, total)])
        for vi, (a, b) in enumerate(ranges, 1):
            tag = "" if len(ranges) == 1 else f"_{vi}권_DAY{a}-{b}"
            for ans in ([False, True] if args.answer else [False]):
                safe = re.sub(r'[\\/:*?"<>|]', "_", book_name)
                base = f"{safe}{tag}" + ("_정답" if ans else "")
                hp = os.path.join(args.out, base + ".html")
                pp = os.path.join(args.out, base + ".pdf")
                open(hp, "w", encoding="utf-8").write(
                    build_html(book_name, units, a, b, answer=ans, title_suffix=tag))
                B.html_to_pdf(hp, pp)
                B.stamp_page_numbers(pp)
                print("생성:", pp)


if __name__ == "__main__":
    main()
