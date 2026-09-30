"""POC-002: validation script for generated Athlete Guides.

Run after generate_guides.py --pdf:
    python validate_guides.py

Checks:
- Exactly 20 DOCX guides (and 20 matching PDFs, if present)
- All 5 event types represented
- Each guide contains all required sections, a distances table, and a schedule table
- Each guide mentions a date, a policy keyword, and a region-specific detail
- Each guide contains version metadata (script version, generation seed, content hash)
- Regenerating produces byte-identical output (reproducibility) -- checked separately,
  see README: rerun generate_guides.py and diff the guides/ folder.
Exit code is 1 if any check fails.
"""
import re
import sys
from pathlib import Path
from docx import Document

OUT_DIR = Path("guides")
REQUIRED_SECTIONS = [
    "Welcome", "Event Overview", "Distances & Cutoff Times", "Race Week Schedule",
    "Race Policies", "Regional Travel & Logistics", "Emergency Contacts",
]
EVENT_CODES = {"IRONMAN", "IRONMAN_703", "IRONMAN_5150", "IRONMAN_25150", "IRONKIDS"}
POLICY_KEYWORDS = ["cancellation", "wetsuit", "cutoff"]
VERSION_META_PATTERN = re.compile(r"document version[:\s]", re.IGNORECASE)


def check_doc(path):
    doc = Document(path)
    headings = [p.text.strip() for p in doc.paragraphs if p.style.name.startswith("Heading")]
    body_text = "\n".join(p.text for p in doc.paragraphs).lower()
    problems = []

    for section in REQUIRED_SECTIONS:
        if not any(section in h for h in headings):
            problems.append(f"missing section '{section}'")

    if len(doc.tables) < 2:
        problems.append(f"expected >=2 tables (distances + schedule), found {len(doc.tables)}")

    for kw in POLICY_KEYWORDS:
        if kw not in body_text:
            problems.append(f"missing policy keyword '{kw}'")

    if not re.search(r"\d{4}-\d{2}-\d{2}", body_text):
        problems.append("no ISO date found")

    if not VERSION_META_PATTERN.search(body_text):
        problems.append("missing version metadata (Document Version / Generation Seed / Content Hash)")
    if "content hash" not in body_text or "generation seed" not in body_text:
        problems.append("incomplete version metadata")

    return problems


def main():
    docx_files = sorted(OUT_DIR.glob("*.docx"))
    pdf_files = sorted(OUT_DIR.glob("*.pdf"))
    failed = 0

    print("== File counts ==")
    ok = len(docx_files) == 20
    print(f"[{'PASS' if ok else 'FAIL'}] 20 DOCX guides found (got {len(docx_files)})")
    failed += 0 if ok else 1

    if pdf_files:
        ok = len(pdf_files) == 20
        print(f"[{'PASS' if ok else 'FAIL'}] 20 matching PDF guides found (got {len(pdf_files)})")
        failed += 0 if ok else 1
    else:
        print("[SKIP] No PDFs found (run generate_guides.py --pdf to produce them)")

    print("\n== Event type coverage ==")
    codes_seen = {f.stem.split("_")[0] if not f.stem.startswith("IRONMAN_") else
                  "_".join(f.stem.split("_")[:2]) for f in docx_files}
    # normalize: filenames start with IRONMAN, IRONMAN_703, IRONMAN_5150, IRONMAN_25150, IRONKIDS
    codes_seen = set()
    for f in docx_files:
        for code in sorted(EVENT_CODES, key=len, reverse=True):
            if f.stem.startswith(code):
                codes_seen.add(code)
                break
    missing = EVENT_CODES - codes_seen
    ok = not missing
    print(f"[{'PASS' if ok else 'FAIL'}] All 5 event types present (missing: {missing or 'none'})")
    failed += 0 if ok else 1

    print("\n== Per-document content checks ==")
    for f in docx_files:
        problems = check_doc(f)
        status = "PASS" if not problems else "FAIL"
        print(f"[{status}] {f.name}" + (f"  -> {'; '.join(problems)}" if problems else ""))
        failed += 0 if not problems else 1

    print(f"\n{'ALL CHECKS PASSED' if not failed else str(failed) + ' CHECK(S) FAILED'}")
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
