# -*- coding: utf-8 -*-
"""뜯어먹는 1800: 묶음별 예문 데이터를 모아 최종 엑셀로 만든다.

`ex_*.py`의 DATA(dict: english -> (example, synonyms, antonyms))를 전부 합쳐
`all.xlsx`(추출 스켈레톤)에 채운다. 아직 안 쓴 DAY는 빈칸으로 두고 진행률만 보고.

검사:
  - 표제어가 예문에 그대로 들어 있는가 (없으면 워크북 빈칸이 문장 끝에 붙어버림)
  - 엑셀에 없는 여분 단어가 데이터에 있는가 (오타)
  - 예문 중복 (다른 단어에 같은 문장을 쓰면 학습 효과가 떨어짐)
"""
import importlib.util
import glob
import os
import re
import sys
import collections
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "all.xlsx")
OUT = os.path.join(HERE, "뜯어먹는 중학 영단어 1800.xlsx")


def load_data():
    """ex_*.py 의 DATA를 합친다."""
    data, where = {}, {}
    mods = sorted(glob.glob(os.path.join(HERE, "ex_*.py")))
    for path in mods:
        if not os.path.exists(path):
            continue
        name = os.path.basename(path)[:-3]
        spec = importlib.util.spec_from_file_location(name, path)
        mod = importlib.util.module_from_spec(spec)
        sys.modules[name] = mod
        spec.loader.exec_module(mod)
        for k, v in getattr(mod, "DATA", {}).items():
            if k in data:
                print("  !! 중복 정의:", k, "(", where[k], "/", name, ")")
            data[k] = v
            where[k] = name
    return data


def primary(eng):
    """표제어 대표형 — teenager[teen] -> teenager, dialog(ue) -> dialog."""
    return re.split(r"[\[(]", eng)[0].strip()


def main():
    data = load_data()
    wb = openpyxl.load_workbook(SRC)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2))
    words = {r[3].value for r in rows}

    done_units, noword, filled = collections.Counter(), [], 0
    for r in rows:
        eng = r[3].value
        if eng not in data:
            continue
        ex, syn, ant = data[eng]
        r[5].value = syn or None
        r[6].value = ant or None
        r[7].value = ex
        filled += 1
        done_units[r[1].value] += 1
        if not re.search(r"\b" + re.escape(primary(eng)) + r"\b", ex, re.I):
            noword.append((eng, ex))

    extra = sorted(set(data) - words)
    dup = [(x, n) for x, n in collections.Counter(
        d[0] for k, d in data.items()).items() if n > 1]

    total = len(rows)
    print("진행: %d / %d 단어 (%.0f%%) | 완료 DAY %d개" % (
        filled, total, filled / total * 100,
        sum(1 for u, n in done_units.items() if n == 30)))
    part = [u for u, n in sorted(done_units.items()) if n != 30]
    if part:
        print("  !! 30개가 안 되는 DAY:", part)
    if extra:
        print("  !! 엑셀에 없는 단어(오타?):", extra)
    if noword:
        print("  !! 표제어가 예문에 없음:")
        for e, x in noword:
            print("     ", e, "|", x)
    if dup:
        print("  !! 예문 중복:")
        for x, n in dup:
            print("      (%d회) %s" % (n, x))
    if extra or noword or dup or part:
        sys.exit(1)

    syn_n = sum(1 for r in rows if r[5].value)
    ant_n = sum(1 for r in rows if r[6].value)
    print("동의어 %d(%.0f%%) / 반의어 %d(%.0f%%)  * 1200은 11%%/19%%" % (
        syn_n, syn_n / filled * 100, ant_n, ant_n / filled * 100))
    wb.save(OUT)
    print("저장:", OUT)


if __name__ == "__main__":
    main()
