-- ====================================================================
-- university.db  --  version 1: the schema as Chapter 9's seven rules
--                    produce it from the Chapter 6 diagram.
--
-- Every table is STRICT and every primary key is NOT NULL, for the two
-- reasons Chapter 8 measured: without STRICT an INTEGER column accepts
-- 'hello', and without NOT NULL a TEXT PRIMARY KEY stores a null.
--
-- Each foreign key carries a referential action and a reason. CASCADE
-- appears exactly twice, both times to a weak entity or an owned value.
--
--   sqlite3 university.db ".read code/schema/v1_create.sql"
-- ====================================================================

-- Per CONNECTION, and ignored inside a transaction. Chapter 8.
PRAGMA foreign_keys = ON;

-- --- Rule 1: strong entities ----------------------------------------

CREATE TABLE department (
  dept_code TEXT PRIMARY KEY NOT NULL,
  name      TEXT NOT NULL
) STRICT;

CREATE TABLE instructor (
  emp_id      TEXT PRIMARY KEY NOT NULL,
  name        TEXT NOT NULL,
  designation TEXT,
  dept_code   TEXT NOT NULL
              REFERENCES department(dept_code)
              ON DELETE RESTRICT        -- a department with staff stays
) STRICT;

CREATE TABLE student (
  roll_no     TEXT    PRIMARY KEY NOT NULL,
  name        TEXT    NOT NULL,
  programme   TEXT    NOT NULL,
  semester    INTEGER NOT NULL CHECK (semester BETWEEN 1 AND 8),
  cnic        TEXT    UNIQUE,           -- alternate key; may be absent
  -- Rule 2: the composite attribute, stored as its parts. One column
  -- called "address" cannot answer "how many students live in
  -- Bahawalpur" -- Chapter 5.
  addr_house  TEXT,
  addr_street TEXT,
  addr_city   TEXT
) STRICT;

-- --- Rule 5: 1:N, foreign key on the MANY side ----------------------

CREATE TABLE course (
  course_code  TEXT    PRIMARY KEY NOT NULL,
  title        TEXT    NOT NULL,
  credit_hours INTEGER NOT NULL CHECK (credit_hours BETWEEN 1 AND 4),
  dept_code    TEXT    NOT NULL        -- total participation: Chapter 6
               REFERENCES department(dept_code)
               ON DELETE RESTRICT       -- a department with courses stays
) STRICT;

-- --- Rule 3: the multi-valued attribute -----------------------------
-- BOTH columns are in the primary key. roll_no alone would give each
-- student one number, which is where we started; no key at all and the
-- same number goes in twice, undeletably -- Chapter 7.

CREATE TABLE student_phone (
  roll_no TEXT NOT NULL
          REFERENCES student(roll_no)
          ON DELETE CASCADE,            -- owned value: cannot outlive it
  phone   TEXT NOT NULL,
  PRIMARY KEY (roll_no, phone)
) STRICT;

-- --- Rule 7: the weak entity ----------------------------------------
-- course_code is three things at once: part of the primary key, a
-- foreign key, and NOT NULL. All three come off the double line in
-- Figure 6.1.

CREATE TABLE section (
  course_code TEXT    NOT NULL
              REFERENCES course(course_code)
              ON DELETE CASCADE,        -- weak entity: the one correct one
  semester    TEXT    NOT NULL,
  section_no  INTEGER NOT NULL CHECK (section_no > 0),
  room        TEXT,
  emp_id      TEXT                      -- nullable: partial participation
              REFERENCES instructor(emp_id)
              ON DELETE SET NULL,       -- the section survives a resignation
  PRIMARY KEY (course_code, semester, section_no)
) STRICT;

-- --- Rule 6: the M:N relationships ----------------------------------
-- The primary key is the combination, and it is what forbids a student
-- appearing in the same section twice.

CREATE TABLE enrolment (
  roll_no     TEXT    NOT NULL
              REFERENCES student(roll_no)
              ON DELETE RESTRICT,       -- never lose a grade history
  course_code TEXT    NOT NULL,
  semester    TEXT    NOT NULL,
  section_no  INTEGER NOT NULL,
  grade       TEXT    CHECK (grade IN ('A','B','C','D','F')),
  PRIMARY KEY (roll_no, course_code, semester, section_no),
  -- A three-column foreign key cannot be declared inline; it needs its
  -- own clause, with the columns in the parent key's order.
  FOREIGN KEY (course_code, semester, section_no)
    REFERENCES section(course_code, semester, section_no)
    ON DELETE RESTRICT
) STRICT;

-- The recursive M:N. Rule 6 unchanged: it takes course_code twice,
-- under the role names from the ER diagram. Without those names there
-- would be two columns of the same name and no way to refer to either.

CREATE TABLE prerequisite (
  course_code TEXT NOT NULL
              REFERENCES course(course_code)
              ON DELETE CASCADE,
  requires    TEXT NOT NULL
              REFERENCES course(course_code)
              ON DELETE RESTRICT,       -- a required course is not deleted
  PRIMARY KEY (course_code, requires),
  CHECK (course_code <> requires)       -- nothing is its own prerequisite
) STRICT;
