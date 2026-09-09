"""Pre-commit PHI guard. Blocks the commit if any STAGED text file contains
something that looks like member data. Git history is permanent, so this
runs before the commit exists.

Patterns (any hit blocks):
  - UHC authorization numbers   A + 9 digits           (e.g. A3xxxxxxxx)
  - WellSky client IDs          10-digit numbers
  - Medicaid IDs                 12-digit numbers
  - date-of-birth labels        DOB / date of birth followed by a date
  - SSN shape                   ddd-dd-dddd

Known-fake test fixtures are allow-listed by file below.
"""
import re
import subprocess
import sys

ALLOW_FILES = {
    "test_compose.py",          # FAKE extract (fictional auth #, ids)
    "test_pipeline_wiring.py",  # same fixture
}
RULES = [
    ("auth number", re.compile(r"\bA3\d{8}\b")),
    ("client id", re.compile(r"(?<![\d.])\d{10}(?![\d.])")),
    ("medicaid id", re.compile(r"(?<![\d.])\d{12}(?![\d.])")),
    ("dob", re.compile(r"\b(?:dob|date of birth)\b[\s:]*\d{1,4}[/-]\d{1,2}[/-]\d{1,4}", re.I)),
    ("ssn", re.compile(r"\b\d{3}-\d{2}-\d{4}\b")),
]


def staged_files():
    out = subprocess.run(["git", "diff", "--cached", "--name-only", "--diff-filter=ACM"],
                         capture_output=True, text=True, check=True).stdout
    return [l.strip() for l in out.splitlines() if l.strip()]


def staged_content(path):
    r = subprocess.run(["git", "show", f":{path}"], capture_output=True)
    return r.stdout.decode("utf-8", errors="replace")


def main():
    bad = []
    for path in staged_files():
        if path.rsplit("/", 1)[-1] in ALLOW_FILES:
            continue
        text = staged_content(path)
        for name, rx in RULES:
            n = len(rx.findall(text))
            if n:
                bad.append(f"  {path}: {n} x {name}")
    if bad:
        print("COMMIT BLOCKED - possible PHI in staged files:")
        print("\n".join(bad))
        print("Remove the data (or add the file to .gitignore) and try again.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
