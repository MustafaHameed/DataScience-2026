"""Confirm Appendix E has exactly one answer per checkpoint question.

Checkpoint questions and their answers live in different files, so adding a
question to a chapter and forgetting the answer produces no LaTeX error and
no visible gap -- the reader just finds the numbering stops early. This
walks both and reports any chapter where the counts or titles disagree.

Exit status is 0 only when every chapter matches.
"""
import re, os

ROOT = r'F:\Documents\Github\DataScience-2026\CCNA'
PARTS = os.path.join(ROOT, 'parts')
MASTER = os.path.join(ROOT, 'CCNA_v2_Handout.tex')

# questions, per chapter, in master order
qcount, n = [], 0
for stem in re.findall(r'\\input\{parts/([^}]+)\}',
                       open(MASTER, encoding='utf-8').read()):
    p = os.path.join(PARTS, stem + '.tex')
    if not os.path.exists(p):
        continue
    body = open(p, encoding='utf-8').read()
    if not re.search(r'^\\chapter\{', body, re.M):
        continue
    n += 1
    m = re.search(r'\\begin\{checkpoint\}(.*?)\\end\{checkpoint\}', body, re.S)
    if m:
        title = re.search(r'\\chapter\{(.+?)\}', body).group(1)
        qcount.append((n, title, len(re.findall(r'\\item\s', m.group(1)))))

# answers, per subsection, in Appendix E
ans = open(os.path.join(PARTS, 'E_answers.tex'), encoding='utf-8').read()
blocks = re.findall(
    r'\\subsection\*\{Chapter (\d+) --- (.+?)\}\s*'
    r'\\begin\{tightnum\}(.*?)\\end\{tightnum\}', ans, re.S)
acount = [(int(a), b, len(re.findall(r'\\item\s', c))) for a, b, c in blocks]

print(f'{len(qcount)} chapters with checkpoints, {len(acount)} answer blocks')
qd = {c: (t, k) for c, t, k in qcount}
ad = {c: (t, k) for c, t, k in acount}

bad = 0
for c in sorted(set(qd) | set(ad)):
    q = qd.get(c)
    a = ad.get(c)
    if not q:
        print(f'  ch{c}: answers with no questions'); bad += 1
    elif not a:
        print(f'  ch{c} {q[0]}: NO ANSWERS'); bad += 1
    elif q[1] != a[1]:
        print(f'  ch{c} {q[0]}: {q[1]} questions vs {a[1]} answers'); bad += 1
    elif q[0].strip() != a[0].strip():
        print(f'  ch{c}: title mismatch "{q[0]}" vs "{a[0]}"'); bad += 1

print(f'total questions {sum(k for _,_,k in qcount)}, '
      f'total answers {sum(k for _,_,k in acount)}')
print('MISMATCHES:', bad if bad else 'none')
raise SystemExit(1 if bad else 0)
