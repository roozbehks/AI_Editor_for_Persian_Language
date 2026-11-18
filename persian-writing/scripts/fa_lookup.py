#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fa_lookup.py — جست‌وجو در پیکرهٔ متنیِ منابع (پرونده‌های txt و md).

وقتی قاعده‌ای در راهنما نیست یا دلیل و نمونهٔ بیشتر می‌خواهید، به جای خواندنِ کل منابع
فقط همان تکه را پیدا کنید. نیم‌فاصله و فاصله، «ی/ک» فارسی و عربی، و «هٔ/ۀ/ه» یکی گرفته می‌شود.

  python3 fa_lookup.py "می‌باشد"                 در پوشهٔ corpus (کنار پوشهٔ scripts) یا متغیر FA_CORPUS
  python3 fa_lookup.py "حذف فعل" -n 12 --ctx 200
  python3 fa_lookup.py "گرته" --files             فقط نام پرونده‌ها و شمارِ هر کدام
  python3 fa_lookup.py "تنوین" --name فرهنگستان   فقط پرونده‌هایی که نامشان این را دارد
  python3 fa_lookup.py "الف" "ب"                  سطرهایی که هر دو را نزدیک هم دارند
"""
import argparse
import os
import re
import sys

ZW = '‌'


def pattern(q):
    out = []
    for ch in q.strip():
        if ch in ' ' + ZW:
            out.append('[' + ZW + r'\s]*')
        elif ch in 'یيى':
            out.append('[یيىئ]')
        elif ch in 'کك':
            out.append('[کك]')
        elif ch in 'هۀ':
            out.append('[هۀة]ٔ?')
        elif ch == 'ٔ':
            continue
        else:
            out.append(re.escape(ch) + '[ً-ْ]?')
    return ''.join(out)


def main(argv=None):
    here = os.path.dirname(os.path.abspath(__file__))
    default = os.environ.get('FA_CORPUS') or os.path.join(os.path.dirname(here), 'corpus')
    ap = argparse.ArgumentParser(description='جست‌وجو در پیکرهٔ منابع')
    ap.add_argument('query', nargs='+', help='یک یا چند عبارت (همه باید نزدیک هم بیایند)')
    ap.add_argument('--dir', default=default, help='پوشهٔ پیکره')
    ap.add_argument('-n', type=int, default=8, help='بیشینهٔ نتیجه')
    ap.add_argument('--per-file', type=int, default=2, help='بیشینهٔ نتیجه از هر پرونده')
    ap.add_argument('--ctx', type=int, default=160, help='نویسه‌های پیش و پس')
    ap.add_argument('--name', help='فقط پرونده‌هایی که این در نامشان هست')
    ap.add_argument('--files', action='store_true', help='فقط نام پرونده و شمار')
    a = ap.parse_args(argv)
    if not os.path.isdir(a.dir):
        sys.exit('پوشهٔ پیکره پیدا نشد: ' + a.dir + ' (با --dir یا متغیر FA_CORPUS بدهید)')
    rx = [re.compile(pattern(q)) for q in a.query]
    hits = []
    for root, _, names in os.walk(a.dir):
        for name in names:
            if not name.lower().endswith(('.txt', '.md')):
                continue
            path = os.path.join(root, name)
            rel = os.path.relpath(path, a.dir)
            if a.name and a.name not in rel:
                continue
            try:
                text = open(path, encoding='utf-8', errors='ignore').read()
            except OSError:
                continue
            found = []
            for m in rx[0].finditer(text):
                s, e = max(0, m.start() - a.ctx), min(len(text), m.end() + a.ctx)
                ctx = text[s:e]
                if all(r.search(ctx) for r in rx[1:]):
                    found.append(re.sub(r'\s+', ' ', ctx).strip())
            if found:
                hits.append((len(found), rel, found))
    hits.sort(key=lambda h: (-h[0], h[1]))
    total = sum(h[0] for h in hits)
    print('{0} مورد در {1} پرونده'.format(total, len(hits)))
    if a.files:
        for n, rel, _ in hits[:60]:
            print('{0:>5}  {1}'.format(n, rel))
        return 0
    shown = 0
    for n, rel, found in hits:
        print('\n■ {0}  ({1} مورد)'.format(rel, n))
        for ctx in found[:a.per_file]:
            print('  …' + ctx + '…')
            shown += 1
        if shown >= a.n:
            break
    return 0


if __name__ == '__main__':
    sys.exit(main())
