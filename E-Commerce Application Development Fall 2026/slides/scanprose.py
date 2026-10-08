"""Scan the slide pipeline's string literals for a sibling course's prose.

This pipeline is a fork. The retarget replaced *names* -- "Analysis of
Algorithms" became "E-Commerce Application Development" everywhere in one
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
    r"questionnaire|respondent|literature review"
    r")\b")

# Literals that may mention one of those anyway, because they are this
# course's own subject matter or a library's actual name.
ALLOW = re.compile(
    r"(?i)dockerfile|terraform|DockerLexer|TerraformLexer|hcl|lexer|"
    r"kernel density|"
    # This course's own files legitimately explain what was REMOVED from the
    # algorithms copy, and those comments are the documentation of the fork.
    r"algorithms course|the siblings|sibling course")


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


def main() -> int:
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
    print(f"\n  {hits} suspicious literal(s)")
    return 1 if hits else 0


if __name__ == "__main__":
    sys.exit(main())
