"""Prove that every CCNA v2.0 exam topic is covered by a chapter.

Reads the twenty-nine numbered topics straight out of Cisco's own published
exam-topics PDF -- the source of record, not a copy typed into this script --
then greps the \\bp{} tags out of parts/*.tex and reports:

  * topics claimed by no chapter        (a gap; the build should not ship)
  * tags naming a topic Cisco does not  (a typo in a chapter)
  * the full topic -> chapter mapping   (used to write Appendix D)

Run:  python tools/check_coverage.py
Exit: 0 if every topic is covered, 1 otherwise.
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
BLUEPRINT = os.path.join(ROOT, '200-301_CCNA_v2.0_Exam_Topics_PDF.pdf')
PARTS = os.path.join(ROOT, 'parts')


def read_blueprint():
    """Return {topic: text} for the top-level topics, and {domain: (weight, name)}."""
    try:
        import fitz
    except ImportError:
        sys.exit('PyMuPDF is required: pip install pymupdf')
    doc = fitz.open(BLUEPRINT)
    raw = ''.join(page.get_text() for page in doc)
    raw = re.sub(r'\d{4} Cisco Systems, Inc\..*?Page \d+', '', raw, flags=re.S)
    flat = re.sub(r'\s+', ' ', raw)

    domains = {}
    for weight, num, name in re.findall(
            r'(\d+)% ([1-5])\.0 ([A-Za-z ,&]+?)(?= \d\.\d )', flat):
        domains[num] = (weight + '%', name.strip())

    # A top-level topic runs from its number to the next number of any depth.
    topics = {}
    for m in re.finditer(
            r'(?<![.\d])([1-5]\.[1-9])\s+(.*?)(?=\s[1-5]\.[1-9](?:\.[a-z])?\s|\s\d+%\s|$)',
            flat):
        num, text = m.group(1), m.group(2).strip()
        if num not in topics:
            topics[num] = re.sub(r'\s+', ' ', text)
    return domains, topics


def read_tags():
    """Return {topic: [chapter titles]} from the \\bp{} tags in parts/."""
    claimed = {}
    for fn in sorted(os.listdir(PARTS)):
        if not fn.endswith('.tex'):
            continue
        text = io.open(os.path.join(PARTS, fn), encoding='utf-8').read()
        title = re.search(r'\\chapter\{(.+?)\}', text)
        title = title.group(1) if title else fn
        for tag in re.findall(r'\\bp\{([^}]*)\}', text):
            for topic in [t.strip() for t in tag.split(',') if t.strip()]:
                claimed.setdefault(topic, []).append((fn, title))
    return claimed


def main():
    domains, topics = read_blueprint()
    claimed = read_tags()

    print('Blueprint: %d domains, %d top-level topics' % (len(domains), len(topics)))
    for d in sorted(domains):
        w, name = domains[d]
        print('  %s.0  %-5s %s' % (d, w, name))
    print()

    missing = [t for t in sorted(topics) if t not in claimed]
    unknown = [t for t in sorted(claimed) if t not in topics]

    print('%-6s %-46s %s' % ('TOPIC', 'CHAPTER(S) CLAIMING IT', 'STATUS'))
    print('-' * 78)
    for t in sorted(topics, key=lambda s: (int(s.split('.')[0]), int(s.split('.')[1]))):
        who = claimed.get(t, [])
        names = '; '.join(sorted({title for _, title in who})) if who else '--'
        print('%-6s %-46s %s' % (t, names[:46], 'ok' if who else 'MISSING'))

    print()
    if unknown:
        print('Tags naming a topic that is not in the blueprint (typo?):')
        for t in unknown:
            for fn, _ in claimed[t]:
                print('  %s  in %s' % (t, fn))
        print()

    if missing:
        print('UNCOVERED TOPICS (%d):' % len(missing))
        for t in missing:
            print('  %-5s %s' % (t, topics[t][:70]))
        return 1

    print('All %d topics are covered.' % len(topics))
    return 0


if __name__ == '__main__':
    sys.exit(main())
