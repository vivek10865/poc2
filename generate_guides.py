"""POC-002: Athlete Guide document generator.

Generates 20 synthetic Athlete Guide documents (.docx, converted to .pdf) covering
five IRONMAN-family event types, each with realistic sections, tables, dates,
policies, and regional variations (units, wetsuit rules, cutoff times, emergency
numbers, language notes).

Usage:
    python generate_guides.py            # writes .docx files to ./guides
    python generate_guides.py --pdf      # also converts each to PDF via LibreOffice
"""
import argparse
import hashlib
import random
import subprocess
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

SEED = 42
OUT_DIR = Path("guides")
REFERENCE_DATE = date(2026, 9, 1)
SCRIPT_VERSION = "1.0.0"  # bump this when the guide template/content logic changes

# 5 IRONMAN-family event types, each with its own distances / cutoff structure
EVENT_TYPES = [
    dict(code="IRONMAN", label="IRONMAN (Full Distance)",
         swim_km=3.8, bike_km=180.2, run_km=42.2, cutoff_h=17.0,
         swim_cutoff_h=2.3, bike_cutoff_h=10.5),
    dict(code="IRONMAN_703", label="IRONMAN 70.3 (Half Distance)",
         swim_km=1.9, bike_km=90.0, run_km=21.1, cutoff_h=8.5,
         swim_cutoff_h=1.2, bike_cutoff_h=5.5),
    dict(code="IRONMAN_5150", label="IRONMAN 5150 (Olympic Distance)",
         swim_km=1.5, bike_km=40.0, run_km=10.0, cutoff_h=4.0,
         swim_cutoff_h=0.75, bike_cutoff_h=2.5),
    dict(code="IRONMAN_25150", label="IRONMAN 25150 (Sprint Distance)",
         swim_km=0.75, bike_km=20.0, run_km=5.0, cutoff_h=2.0,
         swim_cutoff_h=0.4, bike_cutoff_h=1.25),
    dict(code="IRONKIDS", label="IRONKIDS Fun Run",
         swim_km=0.0, bike_km=0.0, run_km=2.0, cutoff_h=1.0,
         swim_cutoff_h=0.0, bike_cutoff_h=0.0),
]

# 5 regions with regional variations
REGIONS = [
    dict(name="North America", cities=["Chattanooga, USA", "Mont-Tremblant, Canada"],
         units="imperial", currency="USD", emergency="911",
         water_temp_c=22, language="English", travel_note="Domestic flights connect through Atlanta or Toronto."),
    dict(name="Europe", cities=["Frankfurt, Germany", "Nice, France"],
         units="metric", currency="EUR", emergency="112",
         water_temp_c=19, language="English / local language", travel_note="Well served by regional rail from major EU hubs."),
    dict(name="Asia-Pacific", cities=["Cairns, Australia", "Taupo, New Zealand"],
         units="metric", currency="AUD", emergency="000",
         water_temp_c=24, language="English", travel_note="International arrivals via Brisbane or Auckland recommended."),
    dict(name="Latin America", cities=["Cozumel, Mexico", "Florianopolis, Brazil"],
         units="metric", currency="USD", emergency="911",
         water_temp_c=27, language="Spanish / Portuguese / English", travel_note="Airport transfers should be booked 48h in advance."),
    dict(name="Middle East & Africa", cities=["Dubai, UAE", "Cape Town, South Africa"],
         units="metric", currency="USD", emergency="999",
         water_temp_c=25, language="English / Arabic", travel_note="Visa-on-arrival available for most nationalities; check current rules."),
]

STYLE = dict(heading_color="1F4E79")


def wetsuit_rule(water_temp_c):
    if water_temp_c <= 24.5:
        return "Wetsuits permitted and recommended."
    elif water_temp_c <= 28.9:
        return "Wetsuits optional; athletes using one are ineligible for age-group awards above 24.5C per race rules."
    else:
        return "Wetsuits not permitted (water temperature exceeds the non-wetsuit threshold)."


def build_guide(event, region, city, event_date, rng):
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)

    title = doc.add_heading(f"{event['label']} — {city}", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph(f"Official Athlete Guide  |  Race Day: {event_date.strftime('%A, %d %B %Y')}")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 0. Document metadata (version, generation seed, content hash for reproducibility)
    doc_id = f"{event['code']}-{city}-{event_date.isoformat()}"
    content_hash = hashlib.sha256(f"{doc_id}|seed={SEED}|v={SCRIPT_VERSION}".encode()).hexdigest()[:12]
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta_run = meta.add_run(
        f"Document Version: {SCRIPT_VERSION}  |  Generation Seed: {SEED}  |  "
        f"Content Hash: {content_hash}  |  Generated: {REFERENCE_DATE.isoformat()}"
    )
    meta_run.italic = True
    meta_run.font.size = Pt(9)

    # 1. Welcome
    doc.add_heading("1. Welcome", level=1)
    doc.add_paragraph(
        f"Welcome to {event['label']} in {city}. This guide contains everything you need to know "
        f"for race week: schedule, distances, cutoffs, policies, and regional logistics for the "
        f"{region['name']} region. Please read it carefully before race day."
    )

    # 2. Event overview
    doc.add_heading("2. Event Overview", level=1)
    doc.add_paragraph(f"Event type: {event['label']}")
    doc.add_paragraph(f"Location: {city} ({region['name']})")
    doc.add_paragraph(f"Race date: {event_date.isoformat()}")
    doc.add_paragraph(f"Registration opens: {(event_date - timedelta(days=365)).isoformat()}")
    doc.add_paragraph(f"Registration closes: {(event_date - timedelta(days=14)).isoformat()}")
    doc.add_paragraph(f"Athlete check-in window: {(event_date - timedelta(days=2)).isoformat()} "
                       f"to {(event_date - timedelta(days=1)).isoformat()}")

    # 3. Distances & cutoffs table
    doc.add_heading("3. Distances & Cutoff Times", level=1)
    unit = "mi / ft" if region["units"] == "imperial" else "km"

    def conv(km):
        return round(km * 0.621371, 2) if region["units"] == "imperial" else km

    table = doc.add_table(rows=1, cols=3)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text = "Discipline", f"Distance ({unit})", "Cutoff Time"
    rows_data = [
        ("Swim", conv(event["swim_km"]), f"{event['swim_cutoff_h']:.2f} h" if event["swim_cutoff_h"] else "N/A"),
        ("Bike", conv(event["bike_km"]), f"{event['bike_cutoff_h']:.2f} h" if event["bike_cutoff_h"] else "N/A"),
        ("Run", conv(event["run_km"]), f"{event['cutoff_h']:.2f} h (overall)" if event["cutoff_h"] else "N/A"),
    ]
    for label, dist, cutoff in rows_data:
        cells = table.add_row().cells
        cells[0].text, cells[1].text, cells[2].text = label, str(dist), cutoff

    # 4. Schedule table
    doc.add_heading("4. Race Week Schedule", level=1)
    sched = doc.add_table(rows=1, cols=2)
    sched.style = "Light Grid Accent 1"
    sched.rows[0].cells[0].text, sched.rows[0].cells[1].text = "Date", "Activity"
    schedule_items = [
        (event_date - timedelta(days=2), "Athlete check-in & expo opens"),
        (event_date - timedelta(days=1), "Mandatory athlete briefing, bike/gear check-in"),
        (event_date, f"Race day — {event['label']} start"),
        (event_date, "Awards ceremony (evening)"),
        (event_date + timedelta(days=1), "Gear bag pickup deadline"),
    ]
    for d, activity in schedule_items:
        cells = sched.add_row().cells
        cells[0].text, cells[1].text = d.strftime("%d %b %Y"), activity

    # 5. Race policies
    doc.add_heading("5. Race Policies", level=1)
    doc.add_paragraph(
        f"Cancellation & refunds: Entries cancelled more than 30 days before race day receive a "
        f"50% credit toward a future event; no refunds within 30 days. Fees quoted in {region['currency']}."
    )
    doc.add_paragraph(
        f"Wetsuit rule: expected water temperature is approximately {region['water_temp_c']}C. "
        + wetsuit_rule(region["water_temp_c"])
    )
    doc.add_paragraph(
        "Weather / force majeure: organizers may modify, shorten, or neutralize any race segment "
        "for athlete safety. No refunds are issued for weather-related changes."
    )
    doc.add_paragraph(
        "Missed cutoffs: athletes not reaching a checkpoint within the published cutoff time will be "
        "withdrawn from the race by race officials and must surrender their timing chip."
    )

    # 6. Regional information
    doc.add_heading("6. Regional Travel & Logistics", level=1)
    doc.add_paragraph(f"Primary language(s) on course: {region['language']}")
    doc.add_paragraph(f"Emergency services number: {region['emergency']}")
    doc.add_paragraph(f"Travel note: {region['travel_note']}")
    doc.add_paragraph(f"Units used on course signage: {'miles/feet' if region['units']=='imperial' else 'kilometers/meters'}")

    # 7. Emergency contacts
    doc.add_heading("7. Emergency Contacts", level=1)
    doc.add_paragraph(f"Local emergency services: {region['emergency']}")
    doc.add_paragraph(f"Race medical tent: staffed from 2 hours before start until 2 hours after cutoff.")
    doc.add_paragraph("Athlete hotline: printed on the back of your race bib and in the mobile app.")

    doc.add_paragraph()
    footer = doc.add_paragraph(
        f"This is a synthetic document generated for POC-002 testing purposes only. "
        f"Generated {REFERENCE_DATE.isoformat()} | Script v{SCRIPT_VERSION} | Seed {SEED} | "
        f"Doc ID: {doc_id} | Hash: {content_hash}"
    )
    footer.runs[0].italic = True
    return doc


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", action="store_true", help="also convert each guide to PDF via LibreOffice")
    args = ap.parse_args()

    rng = random.Random(SEED)
    OUT_DIR.mkdir(exist_ok=True)

    # Build 20 guides: each of the 5 event types appears 4 times, once per region-ish spread
    assignments = []
    region_cycle = REGIONS * 2  # 10 slots, cycle to cover all 5 event types x 4 = 20
    idx = 0
    base_date = date(2027, 1, 10)
    for event in EVENT_TYPES:
        for n in range(4):
            region = REGIONS[(idx) % len(REGIONS)]
            city = region["cities"][n % len(region["cities"])]
            event_date = base_date + timedelta(days=idx * 11)
            event_date += timedelta(days=(6 - event_date.weekday()) % 7)  # Sunday
            assignments.append((event, region, city, event_date))
            idx += 1

    filenames = []
    for event, region, city, event_date in assignments:
        doc = build_guide(event, region, city, event_date, rng)
        safe_city = city.split(",")[0].replace(" ", "_")
        fname = f"{event['code']}_{safe_city}_{event_date.isoformat()}.docx"
        path = OUT_DIR / fname
        doc.save(path)
        filenames.append(path)
    print(f"Generated {len(filenames)} DOCX athlete guides in ./{OUT_DIR}")

    if args.pdf:
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            print("LibreOffice not found; skipping PDF conversion. Install libreoffice or drop --pdf.")
            return
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(OUT_DIR)] +
            [str(p) for p in filenames],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.STDOUT,
        )
        print(f"Converted {len(filenames)} guides to PDF in ./{OUT_DIR}")


if __name__ == "__main__":
    main()
