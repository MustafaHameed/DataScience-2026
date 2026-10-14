"""Fill a university.db with reproducible data, at any scale.

Every measured chapter from 13 on needs enough rows for a difference to be
visible, and needs the SAME rows every time so that a figure in the handout
can be reproduced. This builds the v1 schema and populates it from a fixed
random seed, so that build(path, 4000) produces the identical database on
every machine.

    python seed.py university.db              4,000 students (the default)
    python seed.py big.db --students 100000   for Chapter 21
    python seed.py t.db --students 500 --seed 7

As a module:

    from seed import build
    cx = build("u.db", students=4000)        returns an open connection

Proportions are fixed, so the shape of the data does not change with scale:
about 5 sections per student per semester, 8 departments, 6 courses per
department, 2-3 sections per course per semester. That is the shape Chapter 9
predicted -- enrolment is by far the largest table -- and Chapter 21's index
measurements depend on it.
"""
from __future__ import annotations

import argparse
import os
import random
import sqlite3

HERE = os.path.dirname(os.path.abspath(__file__))
V1 = os.path.join(HERE, "v1_create.sql")

DEPTS = [("CS", "Computer Science"), ("IT", "Information Technology"),
         ("SE", "Software Engineering"), ("DS", "Data Science"),
         ("MA", "Mathematics"), ("PH", "Physics"),
         ("EN", "English"), ("MG", "Management")]
PROGRAMMES = ["BSIT", "BSCS", "BSSE", "BSDS"]
SEMESTERS = ["FALL2025", "SPRING2026", "FALL2026"]
GRADES = ["A", "B", "C", "D", "F"]
GRADE_W = [18, 32, 28, 14, 8]          # a plausible, not a flattering, spread
CITIES = ["Bahawalpur", "Multan", "Lahore", "Rahim Yar Khan",
          "Bahawalnagar", "Lodhran", "Vehari", "Karachi"]
FIRST = ["Ayesha", "Bilal", "Hira", "Usman", "Fatima", "Ali", "Zainab",
         "Hamza", "Maryam", "Omar", "Sana", "Ahmed", "Iqra", "Hassan",
         "Khadija", "Saad", "Amna", "Talha", "Rabia", "Fahad"]
LAST = ["Khan", "Ahmed", "Nadeem", "Tariq", "Malik", "Hussain", "Iqbal",
        "Raza", "Shah", "Butt", "Chaudhry", "Qureshi", "Siddiqui", "Akhtar"]


def build(path: str, students: int = 4000, seed: int = 2026,
          fk: bool = True) -> sqlite3.Connection:
    """Create path from v1_create.sql and fill it. Returns the connection."""
    if os.path.exists(path):
        os.remove(path)
    rng = random.Random(seed)
    cx = sqlite3.connect(path)
    with open(V1, encoding="utf-8") as fh:
        cx.executescript(fh.read())
    if fk:
        cx.execute("PRAGMA foreign_keys = ON")

    cx.executemany("INSERT INTO department VALUES (?,?)", DEPTS)

    # instructors: about one per 40 students, at least 2 per department
    n_ins = max(len(DEPTS) * 2, students // 40)
    ins = []
    for i in range(n_ins):
        d = DEPTS[i % len(DEPTS)][0]
        ins.append((f"E-{i + 100:04d}",
                    f"{rng.choice(FIRST)} {rng.choice(LAST)}",
                    rng.choice(["Lecturer", "Assistant Professor",
                                "Associate Professor", "Professor"]), d))
    cx.executemany("INSERT INTO instructor VALUES (?,?,?,?)", ins)

    # courses: 6 per department
    courses = []
    for d, _ in DEPTS:
        for k in range(6):
            courses.append((f"{d}-{100 + k * 50 + rng.randint(0, 9)}",
                            f"{d} Course {k + 1}", rng.choice([2, 3, 3, 4]),
                            d))
    # course codes must be unique; the random digit can collide
    seen, uniq = set(), []
    for c in courses:
        if c[0] not in seen:
            seen.add(c[0])
            uniq.append(c)
    courses = uniq
    cx.executemany("INSERT INTO course VALUES (?,?,?,?)", courses)
    codes = [c[0] for c in courses]

    # prerequisites: a sparse DAG -- a course may require one or two
    # courses from its own department with a lower number
    pre = set()
    for c in courses:
        same = [x for x in codes if x[:2] == c[0][:2] and x < c[0]]
        for r in rng.sample(same, k=min(len(same), rng.choice([0, 0, 1, 2]))):
            pre.add((c[0], r))
    cx.executemany("INSERT INTO prerequisite VALUES (?,?)", sorted(pre))

    # sections: sized so that about 5 enrolments per student per semester
    # fit at roughly 40 students a section
    per_sem = max(1, (students * 5) // 40)
    sections = []
    for sem in SEMESTERS:
        for j in range(per_sem):
            code = codes[j % len(codes)]
            no = j // len(codes) + 1
            emp = rng.choice(ins)[0] if rng.random() > 0.04 else None
            sections.append((code, sem, no, f"R{rng.randint(1, 40)}", emp))
    cx.executemany("INSERT INTO section VALUES (?,?,?,?,?)", sections)
    by_sem: dict[str, list[tuple]] = {s: [] for s in SEMESTERS}
    for s in sections:
        by_sem[s[1]].append(s)

    # students
    stu = []
    for i in range(students):
        prog = rng.choice(PROGRAMMES)
        year = rng.choice([2023, 2024, 2025])
        roll = f"{year}-{prog[2:]}-{i:05d}"
        stu.append((roll, f"{rng.choice(FIRST)} {rng.choice(LAST)}", prog,
                    rng.randint(1, 8),
                    None if rng.random() < 0.12 else
                    f"{rng.randint(31000, 36999)}-{rng.randint(1000000, 9999999)}-{rng.randint(1, 9)}",
                    f"House {rng.randint(1, 400)}",
                    f"Street {rng.randint(1, 60)}",
                    rng.choice(CITIES)))
    cx.executemany("INSERT INTO student VALUES (?,?,?,?,?,?,?,?)", stu)

    # phones: 0-3 each
    ph = set()
    for s in stu:
        for _ in range(rng.choice([0, 1, 1, 1, 2, 2, 3])):
            ph.add((s[0], f"03{rng.randint(0, 4)}{rng.randint(0, 9)}-"
                          f"{rng.randint(1000000, 9999999)}"))
    cx.executemany("INSERT INTO student_phone VALUES (?,?)", sorted(ph))

    # enrolments: each student takes about 5 sections in each semester,
    # never the same section twice (the composite key would refuse it).
    # The current semester has no grades yet -- null, as Chapter 7 said.
    enr = []
    for s in stu:
        for sem in SEMESTERS:
            pool = by_sem[sem]
            for sec in rng.sample(pool, k=min(len(pool), rng.choice([4, 5, 5, 6]))):
                g = None if sem == "FALL2026" else \
                    rng.choices(GRADES, GRADE_W)[0]
                enr.append((s[0], sec[0], sec[1], sec[2], g))
    cx.executemany("INSERT INTO enrolment VALUES (?,?,?,?,?)", enr)

    cx.execute("PRAGMA user_version = 1")
    cx.commit()
    return cx


def counts(cx: sqlite3.Connection) -> dict[str, int]:
    tabs = ["department", "instructor", "course", "prerequisite", "section",
            "student", "student_phone", "enrolment"]
    return {t: cx.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
            for t in tabs}


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("path")
    ap.add_argument("--students", type=int, default=4000)
    ap.add_argument("--seed", type=int, default=2026)
    a = ap.parse_args()
    cx = build(a.path, a.students, a.seed)
    for t, n in counts(cx).items():
        print(f"  {t:<15}{n:>10,}")
    cx.close()
    print(f"\n  {a.path}: {os.path.getsize(a.path):,} bytes")


if __name__ == "__main__":
    main()
