# POC-002: Athlete Guide Document Generator

Generates 20 synthetic Athlete Guide documents (DOCX, optionally converted to PDF)
covering 5 IRONMAN-family event types with realistic sections, tables, dates,
policies, and regional variations.

## Run order
1. `python3 -m venv venv && source venv/bin/activate`   (Windows: `venv\Scripts\activate`)
2. `pip install -r requirements.txt`
3. `python generate_guides.py --pdf`
   (drop `--pdf` to generate DOCX only, no LibreOffice required)
4. `python validate_guides.py`

PDF conversion uses LibreOffice headless (`soffice --headless --convert-to pdf`).
Install LibreOffice if `--pdf` reports it's missing (e.g. `sudo apt install libreoffice` /
download from libreoffice.org on Windows/Mac).

## Output
`guides/` — 20 `.docx` files (and 20 `.pdf` if `--pdf` used), one per event type x region
combination, e.g. `IRONMAN_Frankfurt_2027-02-14.docx`.

## Event types covered (5)
IRONMAN (full), IRONMAN 70.3 (half), IRONMAN 5150 (Olympic), IRONMAN 25150 (Sprint), IRONKIDS

## Each guide includes
- Document Version / Generation Seed / Content Hash line under the title (version metadata),
  repeated in the footer along with a unique Doc ID
- Welcome, Event Overview, Distances & Cutoff Times (table), Race Week Schedule (table),
  Race Policies (cancellation, wetsuit, weather, cutoffs), Regional Travel & Logistics,
  Emergency Contacts
- Regional variations: units (metric/imperial), currency, wetsuit rule by water temperature,
  local emergency number, language, travel note

## Reproducibility
Generation uses a fixed seed (`SEED = 42` in `generate_guides.py`) and no non-deterministic
content, so rerunning `generate_guides.py` produces byte-identical DOCX text every time.
Bump `SCRIPT_VERSION` at the top of the file whenever the template or content logic changes —
this shows up in each guide's version metadata and content hash.

## Validation
`validate_guides.py` checks: exactly 20 DOCX (+20 PDF if present), all 5 event types present,
every document has all 7 required sections, at least 2 tables, policy keywords, a valid date,
and version metadata (Document Version, Generation Seed, Content Hash).
