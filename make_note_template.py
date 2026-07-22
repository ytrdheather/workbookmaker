# -*- coding: utf-8 -*-
"""단어 암기 노트 (빈 템플릿) 생성기

딸기케이크식 4단 자가시험 암기법을 인쇄용 A4 한 장으로 옮긴 재사용 양식.
- 단어를 미리 인쇄하지 않는다 → 학생이 직접 쓴다(쓰는 것 자체가 스펠링 연습).
- 세로 4단 + 열 사이 점선(접어서 가리고 자가 시험).
- 하단에 '오답만 다시' 정리 박스.
- 정답판/채점칸 없음(자습·자가시험 전용). 엑셀 데이터 불필요.

실행:  python make_note_template.py
출력:  templates/단어암기노트_양식.html / .pdf
"""
import os
from build_workbook import font_face_css, logo_datauri, html_to_pdf

CSS = """
@page { size: A4; margin: 12mm 11mm 12mm 11mm; }
* { box-sizing: border-box; }
body { font-family:'Pretendard','Malgun Gothic',sans-serif; color:#2f302f; margin:0; font-size:13px; }

.phead { display:flex; align-items:flex-end; justify-content:space-between;
  border-bottom:3px solid #32bfb6; padding-bottom:6px; margin-bottom:9px; }
.ph-title { text-align:center; }
.pt-ko { display:block; font-size:19px; font-weight:800; }
.pt-en { display:block; font-size:9px; color:#dfc0aa; letter-spacing:1.5px; text-transform:uppercase; font-weight:700; }
.ph-meta { font-size:11px; color:#8a8f8d; text-align:right; line-height:1.9; }
.nline { display:inline-block; width:78px; border-bottom:1px solid #dfc0aa; margin-left:4px; }
.ph-left { font-size:11px; color:#8a8f8d; line-height:1.9; }

.guide { background:#eafaf8; border-left:4px solid #32bfb6; padding:7px 11px; margin:0 0 10px;
  font-size:11.5px; line-height:1.65; border-radius:3px; }
.guide b { color:#0f6e56; }
.guide .cpen { color:#c0392b; }

table.nt { width:100%; border-collapse:collapse; table-layout:fixed; }
table.nt th { background:#32bfb6; color:#fff; font-size:11px; font-weight:700; padding:6px 4px; }
table.nt col.c-no { width:26px; } table.nt col.c-eng { width:23%; }
table.nt col.c-spell { width:27%; } table.nt col.c-mark { width:34px; }
.fold { border-left:1px dashed #94e1da; }
.th-sub { font-weight:400; font-size:8.5px; opacity:.9; }
table.nt td { border-bottom:1px solid #e3e3e0; }
table.nt tr:nth-child(even) td { background:#f4fbfa; }
.c-no { text-align:center; color:#8a8f8d; font-size:11px; }
.c-mark { text-align:center; color:#b4b2a9; }

.oadap { margin-top:12px; border:1.5px dashed #dfc0aa; background:#f7efe6; border-radius:6px; padding:9px 11px; }
.oadap .oa-t { font-size:11.5px; font-weight:800; color:#b06a2c; margin-bottom:8px; }
.oadap .oa-line { border-bottom:1px solid #d9c3ab; height:20px; margin-bottom:6px; }
.oadap .oa-line:last-child { margin-bottom:0; }
.pagelogo { position:fixed; bottom:1mm; right:2mm; width:28mm; height:auto; opacity:.95; }
"""


def build_html(show_logo=True, rows=20, row_h=31, oa_lines=5):
    head_cells = (
        '<th class="c-no">No</th>'
        '<th class="fold">① 단어</th>'
        '<th class="fold">② 뜻 쓰기<br><span class="th-sub">단어 보고 · 외워서</span></th>'
        '<th class="fold">③ 스펠링 쓰기<br><span class="th-sub">단어 접어 가리고 · 뜻만 보고</span></th>'
        '<th class="fold">모름</th>'
    )
    body_rows = []
    for i in range(1, rows + 1):
        body_rows.append(
            f'<tr><td class="c-no">{i}</td>'
            f'<td class="fold"></td>'
            f'<td class="fold"></td>'
            f'<td class="fold"></td>'
            f'<td class="fold c-mark">□</td></tr>'
        )
    oa_html = "".join('<div class="oa-line"></div>' for _ in range(oa_lines))
    logo = logo_datauri() if show_logo else ""
    logo_html = f'<img class="pagelogo" src="{logo}" alt="logo">' if logo else ""
    guide = (
        '<div class="guide">'
        '<b>① 단어</b> — 단어장을 보고 영어 단어를 옮겨 씁니다. &nbsp;'
        '<b>② 뜻</b> — 단어를 보고 뜻을 <b>외워서</b> 씁니다 '
        '(모르면 <span class="cpen">다른 색 펜</span>으로 다시). &nbsp;'
        '<b>③ 스펠링</b> — 오른쪽 점선을 <b>접어 단어를 가리고</b>, 뜻만 보고 스펠링을 씁니다. &nbsp;'
        '<b>④ 모름 □</b> 체크 → 아래 <b>오답 정리</b> 칸에 <b>모르는 것만</b> 다시.'
        '</div>'
    )
    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<title>단어 암기 노트 양식</title><style>{font_face_css()}
{CSS}
table.nt td {{ height:{row_h}px; }}</style></head>
<body>{logo_html}
  <div class="phead">
    <div class="ph-left">이름 :<span class="nline"></span><br>날짜 :<span class="nline"></span></div>
    <div class="ph-title"><span class="pt-ko">단어 암기 노트</span>
      <span class="pt-en">Memorize &amp; Self-Test</span></div>
    <div class="ph-meta">교재 :<span class="nline"></span><br>범위 :<span class="nline"></span></div>
  </div>
  {guide}
  <table class="nt">
    <colgroup><col class="c-no"><col class="c-eng"><col><col class="c-spell"><col class="c-mark"></colgroup>
    <thead><tr>{head_cells}</tr></thead>
    <tbody>{''.join(body_rows)}</tbody>
  </table>
  <div class="oadap">
    <div class="oa-t">⑤ 오답 정리 — 틀린 단어만 다시 (단어 + 뜻)</div>
    {oa_html}
  </div>
</body></html>"""


def main():
    out = "templates"
    os.makedirs(out, exist_ok=True)
    # (stem, 로고여부, 행 수, 행 높이(px), 오답줄 수)
    variants = [
        ("단어암기노트_양식_20", True, 20, 31, 5),
        ("단어암기노트_양식_20_로고없음", False, 20, 31, 5),
        ("단어암기노트_양식_30", True, 30, 23, 4),
        ("단어암기노트_양식_30_로고없음", False, 30, 23, 4),
    ]
    for stem, show_logo, rows, row_h, oa_lines in variants:
        hp = os.path.join(out, stem + ".html")
        with open(hp, "w", encoding="utf-8") as f:
            f.write(build_html(show_logo=show_logo, rows=rows, row_h=row_h, oa_lines=oa_lines))
        pp = hp[:-5] + ".pdf"
        html_to_pdf(hp, pp)
        print("생성:", pp)


if __name__ == "__main__":
    main()
