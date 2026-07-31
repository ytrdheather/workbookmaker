# -*- coding: utf-8 -*-
"""제본용(무선제본·A4 양면 컬러) PDF 후처리.

생성된 워크북 PDF를 다시 만들지 않고, PDF 단계에서만 손본다:
  1. 짝수 쪽(펼쳤을 때 왼쪽 면)의 로고를 좌하단으로 옮긴다.
     — 안 그러면 로고가 제본 골(책등)로 들어가 잘린다.
  2. 지면을 축소해 안쪽(제본 쪽) 여백을 확보하고, 홀짝을 번갈아 배치한다.
     — 홀수 쪽은 골이 왼쪽, 짝수 쪽은 골이 오른쪽.

벡터 그대로 옮기므로 화질 손실이 없다. 원본 PDF는 건드리지 않는다.

기본값은 A4 무선제본 기준(안쪽 20mm · 바깥 14mm). 본문 폭 188mm가 176.7mm가 되며
축소율은 약 94%다. 세로는 17mm 여유가 생겨 아래쪽 여백으로 간다.

⚠️ 단어 목록이 30단어인 교재(중등 고난도·Bricks 3900)는 축소하면 본문이 8pt 아래로
내려간다. 그 교재는 축소 전에 지면을 먼저 손볼 것.

사용:
  python make_print_ready.py "output_new/초등영어 단어가 된다" --out "print_ready"
  python make_print_ready.py "...\\어떤.pdf" --inner 20 --outer 14 --out print_ready
"""
import argparse
import glob
import os
import sys

import fitz

MM = 72 / 25.4          # 1mm = 2.8346pt


def mm(v):
    return v * MM


def find_content_box(doc):
    """본문이 실제로 차지하는 좌우 범위(pt).

    표본만 훑으면 삐져나간 지면을 놓쳐 여백이 1mm쯤 모자란다. 전 쪽을 잰다."""
    lo, hi = None, None
    for p in doc:
        xs = [b[0] for b in p.get_text("blocks") if b[4].strip()]
        xe = [b[2] for b in p.get_text("blocks") if b[4].strip()]
        for im in p.get_images(full=True):
            try:
                r = p.get_image_bbox(im)
            except Exception:
                continue
            xs.append(r.x0)
            xe.append(r.x1)
        # 쪽번호는 가운데라 좌우 범위에 영향을 주지 않는다
        if xs:
            lo = min(xs) if lo is None else min(lo, min(xs))
            hi = max(xe) if hi is None else max(hi, max(xe))
    return lo, hi


LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "source", "logo.png")


def move_logos(doc, page_w):
    """짝수 쪽(1-based)의 로고를 좌우 대칭 위치로 옮긴다. -> 옮긴 개수

    PDF 안의 이미지를 추출해 다시 넣으면 알파 채널이 빠져 배경이 검게 나온다.
    원본 `source/logo.png`를 그대로 쓴다."""
    if not os.path.exists(LOGO):
        print("  ! source/logo.png 가 없어 로고를 옮기지 않습니다.")
        return 0
    moved = 0
    for i, page in enumerate(doc, 1):
        if i % 2 == 1:
            continue                      # 홀수 쪽은 골이 왼쪽 -> 로고는 그대로
        for im in page.get_images(full=True):
            try:
                r = page.get_image_bbox(im)
            except Exception:
                continue
            if r.width > page_w * 0.5 or r.width <= 0:
                continue                  # 배경 이미지 등은 건드리지 않는다
            # 원래 자리를 흰색으로 덮고(가장자리 자국이 남지 않게 1mm 넉넉히),
            # 페이지 중심 기준 반대쪽에 원본 파일로 다시 넣는다.
            pad = mm(1)
            page.draw_rect(fitz.Rect(r.x0 - pad, r.y0 - pad, r.x1 + pad, r.y1 + pad),
                           color=None, fill=(1, 1, 1), overlay=True)
            page.insert_image(fitz.Rect(page_w - r.x1, r.y0, page_w - r.x0, r.y1),
                              filename=LOGO, overlay=True)
            moved += 1
            break                          # 지면당 로고 하나
    return moved


def impose(src_doc, inner, outer, top, keep_scale=None):
    """축소 + 홀짝 배치. -> 새 문서"""
    out = fitz.open()
    for i, page in enumerate(src_doc, 1):
        pw, ph = page.rect.width, page.rect.height
        lo, hi = impose.box
        avail = pw - mm(inner) - mm(outer)
        s = keep_scale if keep_scale else avail / (hi - lo)
        np = out.new_page(width=pw, height=ph)
        # 본문 왼쪽 끝(lo)이 놓일 자리: 홀수 쪽은 안쪽 여백만큼, 짝수 쪽은 바깥 여백만큼
        left = mm(inner) if i % 2 == 1 else mm(outer)
        dx = left - lo * s
        dy = mm(top) - impose.top * s
        np.show_pdf_page(fitz.Rect(dx, dy, dx + pw * s, dy + ph * s), src_doc, i - 1)
    return out


def process(path, args):
    doc = fitz.open(path)
    pw = doc[0].rect.width
    lo, hi = find_content_box(doc)
    if lo is None:
        print("  건너뜀(본문을 찾지 못함):", os.path.basename(path))
        return
    # 본문 위쪽 끝
    tops = []
    for pg in doc:
        ys = [b[1] for b in pg.get_text("blocks") if b[4].strip()]
        if ys:
            tops.append(min(ys))
    impose.box = (lo, hi)
    impose.top = min(tops) if tops else mm(13)

    moved = move_logos(doc, pw)
    tmp = path + ".tmp"
    doc.save(tmp)
    doc.close()

    src = fitz.open(tmp)
    out = impose(src, args.inner, args.outer, args.top)
    os.makedirs(args.out, exist_ok=True)
    dst = os.path.join(args.out, os.path.basename(path))
    out.save(dst, deflate=True)
    out.close()
    src.close()
    os.remove(tmp)

    scale = (pw - mm(args.inner) - mm(args.outer)) / (hi - lo)
    chk = fitz.open(dst)
    l1, h1 = find_content_box(chk)
    print("  %-38s 축소 %.1f%% | 로고 %d개 이동 | 첫 쪽 여백 좌 %.1f / 우 %.1fmm"
          % (os.path.basename(path)[:38], scale * 100, moved,
             l1 / MM, (pw - h1) / MM))
    chk.close()


def main():
    ap = argparse.ArgumentParser(description="무선제본용 PDF 후처리(축소·홀짝 배치·로고 이동)")
    ap.add_argument("target", help="PDF 파일 또는 폴더")
    ap.add_argument("--out", default="print_ready", help="출력 폴더")
    ap.add_argument("--inner", type=float, default=20, help="안쪽(제본 쪽) 여백 mm")
    ap.add_argument("--outer", type=float, default=14, help="바깥쪽 여백 mm")
    ap.add_argument("--top", type=float, default=18, help="위쪽 여백 mm")
    args = ap.parse_args()

    files = ([args.target] if args.target.lower().endswith(".pdf")
             else sorted(glob.glob(os.path.join(args.target, "*.pdf"))))
    if not files:
        sys.exit("PDF를 찾지 못했습니다: " + args.target)
    print("안쪽 %gmm / 바깥 %gmm / 위 %gmm" % (args.inner, args.outer, args.top))
    for f in files:
        process(f, args)


if __name__ == "__main__":
    main()
