"""Scan the slide pipeline's string literals for a sibling course's prose.

This pipeline is a fork. The retarget replaced *names* -- "Analysis of
Algorithms" became "Database Systems" everywhere in one
sed -- but it could not replace *descriptions*, because a description of
another course does not contain that course's name.

That is not hypothetical. Three blocks of Virtual Systems and Services prose
survived into the Algorithms course's overview.py: a slide about hypervisors
and GPUs, and a card telling students to "bring a laptop with virtualization
enabled". They survived because they never used a word the retarget searched
for, and only a rendered contact sheet found them. This course inherited the
Algorithms copy and promptly shipped a title slide reading "Design, Proof and
Measurement" for exactly the same reason.

So this looks for the *vocabulary* of the sibling courses rather than their
names, in every string literal long enough to be prose.

    python scanprose.py          0 suspicious literals, or a list and exit 1
"""
import argparse
import ast
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

FILES = ["overview.py", "dsdeck.py", "build_slides.py", "texparse.py",
         "qa.py", "icons.py", "pptxfx.py", "measure.py", "mathrender.py",
         os.path.join("figures", "extract_figures.py")]

# The vocabulary of the other six courses in this repository. A hit is not
# automatically wrong -- it is a literal somebody must look at.
OTHER = re.compile(
    r"(?i)\b("
    # Analysis of Algorithms -- the course this pipeline was forked from,
    # and therefore by far the most likely source of a survivor.
    r"asymptotic\w*|recurrence\w*|theorem\w*|lemma\w*|corollar\w+|"
    r"pseudocode|loop invariant|big[- ]?o\b|master theorem|amorti[sz]ed|"
    r"quicksort|merge ?sort|insertion sort|radix sort|NP-complete\w*|"
    r"proof\b|proofs\b|QED|divide[- ]and[- ]conquer|greedy algorithm|"
    r"dynamic programming|shortest path\w*|spanning tree|union[- ]find|"
    # Virtual Systems and Services
    r"hypervis\w*|virtuali[sz]\w*|virtual machine|guest OS|VDI|tenant|"
    r"vmware|xen|kubernetes|"
    # Data Science / Machine Learning / AI
    r"neural|gradient|regression|classifier|dataframe|"
    r"heuristic search|knowledge base|ontology|"
    # Advanced Research Methodology
    r"questionnaire|respondent|literature review|"
    # E-Commerce Application Development -- the course THIS pipeline was
    # forked from, and therefore now the most likely source of a survivor.
    # Added when Database Systems was forked from it, for the same reason
    # the Algorithms list was added when E-Commerce was forked from that:
    # the retarget replaces names, never descriptions.
    # "shop" bare, not just "storefront": the leak this gate was extended for
    # was the single word "shop" in a sentence about a college, and a
    # vocabulary that misses the thing that prompted it is decoration.
    r"shops?\b|storefront|shopping (?:basket|cart)|checkout|opencart|"
    r"admin panel|click path|breadcrumb|coupon|voucher|"
    r"payment gateway|merchant|chargeback|PCI[- ]DSS|"
    r"catalogue page|product variant|geo zone|SEO|sitemap|"
    r"docker compose|container\w* image|"
    # Bare words, added when overview.py was found carrying six E-Commerce
    # cards -- "A catalogue, and the tables under it", "A basket, and where
    # its state lives", "A gateway, and what happens when it fails" -- and a
    # closing card naming "What E-Commerce Is". The phrases above did not
    # match any of them, because the fork used the words alone.
    r"e-commerce|baskets?|gateways?|catalogues?|refunds?|shipping|"
    r"sales tax"
    r")\b")

# Literals that may mention one of those anyway, because they are this
# course's own subject matter or a library's actual name.
ALLOW = re.compile(
    r"(?i)dockerfile|terraform|DockerLexer|TerraformLexer|hcl|lexer|"
    r"kernel density|"
    # This course's own files legitimately explain what was REMOVED from the
    # E-Commerce copy, and those docstrings are the documentation of the fork.
    #
    # These are PHRASES, not bare words, and deliberately so. An earlier
    # draft allowed bare "e-commerce", which suppressed every real leak that
    # happened to name the course it leaked from -- i.e. most of them. An
    # allow-list entry must be narrower than the thing it excuses.
    r"algorithms course|e-commerce course|e-commerce build|"
    r"the siblings|sibling course|"
    # This handout names the E-Commerce course as a downstream dependant;
    # these are the exact places it does so, and the ordinary English uses
    # of two flagged words.
    r"E-Commerce Application Development|catalogue or a bookshop|"
    r"proofs are written|Proof: ")

# A file whose whole subject is the other courses. Its lines are expected to
# use their vocabulary; flagging all of them would train readers to ignore
# the gate, which is the one outcome a gate cannot survive.
EXEMPT_TEX = {"F_boundary.tex"}


def literals(src: str):
    """(line, text) for every string literal, via the parser.

    An earlier version matched "..." with a regular expression, which cannot
    tell a string from a comment and so reported the comments in this very
    pipeline explaining what was removed from the fork. ast sees exactly what
    Python sees: real literals, including docstrings, and nothing else.
    """
    for node in ast.walk(ast.parse(src)):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            if len(node.value) >= 12:
                yield node.lineno, node.value


def tex_lines(path: str):
    """(line, text) for the prose of a .tex file, comments removed.

    The slide pipeline is not the only place a fork leaks. The FIRST draft of
    this course's Chapter 1 described a college's three CSV files and then
    said "the SHOP now holds two different answers" -- E-Commerce vocabulary,
    in the handout, where the Python gate could not see it. One sentence, and
    nothing else in the build would ever have reported it.

    A % that is not \\% starts a comment; the rest of the line is the author
    talking to the next author and is exempt, exactly as a Python comment is.
    """
    for i, raw in enumerate(io.open(path, encoding="utf-8"), 1):
        line = re.sub(r"(?<!\\)%.*$", "", raw).strip()
        if len(line) >= 12:
            yield i, line


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--tex", action="store_true",
                    help="also scan the handout's own prose in parts/")
    a = ap.parse_args()

    hits = 0
    for rel in FILES:
        path = os.path.join(HERE, rel)
        if not os.path.exists(path):
            continue
        src = io.open(path, encoding="utf-8").read()
        for line, lit in literals(src):
            if OTHER.search(lit) and not ALLOW.search(lit):
                flat = " ".join(lit.split())
                print(f"  {rel}:{line}: {flat[:84]}")
                hits += 1

    if a.tex:
        parts = os.path.join(os.path.dirname(HERE), "parts")
        for name in sorted(os.listdir(parts)) if os.path.isdir(parts) else []:
            if not name.endswith(".tex") or name in EXEMPT_TEX:
                continue
            for line, text in tex_lines(os.path.join(parts, name)):
                if OTHER.search(text) and not ALLOW.search(text):
                    print(f"  parts/{name}:{line}: {text[:84]}")
                    hits += 1

    print(f"\n  {hits} suspicious literal(s)")
    return 1 if hits else 0


if __name__ == "__main__":
    sys.exit(main())
