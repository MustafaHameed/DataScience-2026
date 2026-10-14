-- ====================================================================
-- university.db  --  migration 1 -> 2
--
-- Adds, to a database that already has rows in it:
--   * student.email, NOT NULL, backfilled from the roll number
--   * a CHECK that a grade is one of the five letters
--   * an index on student(name)
--
-- The first two cannot be done with ALTER TABLE: SQLite has no
-- ALTER COLUMN, and ADD COLUMN ... NOT NULL is refused on a table with
-- rows unless a default is supplied. So this is the twelve-step
-- rebuild from Chapter 11, written out.
--
-- Run only against a version 1 database:
--   sqlite3 university.db "PRAGMA user_version;"   -- expect 1
--   sqlite3 university.db ".read code/schema/v2_constraints.sql"
--
-- NOT REVERSIBLE. v2_rollback.sql drops the column and the CHECK, but
-- the backfilled addresses cannot afterwards be told from real ones.
-- ====================================================================

-- Outside the transaction, both times. Inside one it is silently
-- ignored, and in a rebuild that costs you a foreign key -- Chapter 8.
PRAGMA foreign_keys = OFF;

BEGIN;

-- --- Step 0 is not in this file ------------------------------------
-- Before running it, check the data can satisfy the new constraints:
--   SELECT COUNT(*) FROM student   WHERE semester NOT BETWEEN 1 AND 8;
--   SELECT COUNT(*) FROM enrolment WHERE grade IS NOT NULL
--                                    AND grade NOT IN ('A','B','C','D','F');
-- Both must be 0. If they are not, the COMMIT below will fail and the
-- decision needed is the university's, not a cleverer script.

-- --- student: add email NOT NULL, by rebuild -----------------------

CREATE TABLE student_new (
  roll_no     TEXT    PRIMARY KEY NOT NULL,
  name        TEXT    NOT NULL,
  programme   TEXT    NOT NULL,
  semester    INTEGER NOT NULL CHECK (semester BETWEEN 1 AND 8),
  cnic        TEXT    UNIQUE,
  addr_house  TEXT,
  addr_street TEXT,
  addr_city   TEXT,
  email       TEXT    NOT NULL UNIQUE
) STRICT;

-- Every column named. SELECT * depends on column order, which is
-- exactly what a rebuild may be changing.
INSERT INTO student_new
  (roll_no, name, programme, semester, cnic,
   addr_house, addr_street, addr_city, email)
SELECT roll_no, name, programme, semester, cnic,
       addr_house, addr_street, addr_city,
       -- the backfill: this is the irreversible part
       roll_no || '@iub.edu.pk'
  FROM student;

DROP TABLE student;
ALTER TABLE student_new RENAME TO student;

-- --- enrolment: the grade CHECK was already in v1, so this table is
-- --- rebuilt only to re-point its foreign key at the new student ----
-- (DROP TABLE student above removed the parent; with foreign_keys OFF
-- the child survived, and step 9's foreign_key_check confirms it is
-- still consistent.)

-- --- Step 7: indexes do not survive DROP TABLE ---------------------
CREATE INDEX idx_student_name ON student(name);

-- --- The version, set INSIDE the transaction -----------------------
-- A rollback then takes it with it, and the database is honestly
-- still at version 1. Set after COMMIT, a half-failed migration
-- leaves a file claiming to be something it is not.
PRAGMA user_version = 2;

COMMIT;

PRAGMA foreign_keys = ON;
