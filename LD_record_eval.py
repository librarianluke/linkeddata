# LD_record_eval
# Author Luke Low
# Version: 2.2
# Description : Fetches WorldCat records for OCLC numbers from a spreadsheet, identifies records with no linked data, counts subject fields and  saves results to Excel.
# Dependencies: bookops-worldcat (MIT), openpyxl (MIT)

import time
import threading
import xml.etree.ElementTree as ET
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from bookops_worldcat import WorldcatAccessToken, MetadataSession

# excel location input and output
BASE_DIR = Path(r"c:/")
EXCEL_FILE = BASE_DIR / "oclc_numbers.xlsx"
OUTPUT_FILE = BASE_DIR / f"oclc_results_{datetime.now():%Y-%m-%d_%H-%M-%S}.xlsx"
MAX_WORKERS = 8  # number of records to fetch at once

# config file with key and secret key
token = WorldcatAccessToken(
    key="",
    secret="",
    scopes="WorldCatMetadataAPI",
)


ns = {"m": "http://www.loc.gov/MARC21/slim"}

# Each thread gets its own session, sharing the one token
thread_data = threading.local()
all_sessions = []
sessions_lock = threading.Lock()

# Get a per-thread MetadataSession, created on first use
def get_session():
    if not hasattr(thread_data, "session"):
        thread_data.session = MetadataSession(authorization=token)
        with sessions_lock:
            all_sessions.append(thread_data.session)
    return thread_data.session


# get record from WorldCat, eval LD and count 650s
def process(record):
    oclc, genre = record
    session = get_session()


    # fetch a record and timeit, also handle errors
    start = time.perf_counter()
    try:
        response = session.bib_get(oclc)
    except Exception as e:
        elapsed = time.perf_counter() - start
        print(f"{oclc}: ERROR after {elapsed:.2f}s - {e}")
        return {"oclc": oclc, "genre": genre, "error": str(e), "time": elapsed}
    elapsed = time.perf_counter() - start

    # Parse the MARC record and set up counters
    root = ET.fromstring(response.content)
    ld_tags = []
    subject_counts = Counter()

    # Loop over every MARC datafield and get fields
    for field in root.findall(".//m:datafield", ns):
        tag = field.get("tag")

        # Count 6XX fields, excluding French headings (ind2=6 or $2 rvm)
        if tag and tag.isdigit() and 600 <= int(tag) <= 699:
            is_french = field.get("ind2") == "6" or any(
                sub.get("code") == "2" and sub.text and sub.text.strip().lower() == "rvm"
                for sub in field.findall("m:subfield", ns)
            )
            if not is_french:
                subject_counts[tag] += 1

        # Check for WorldCat Entity links
        for sub in field.findall("m:subfield", ns):
            if sub.text and "id.oclc.org" in sub.text:
                ld_tags.append(tag)
                break

    # summary of finding LD, 6xx counts and  low 650 flags
    ld_text = f"LD was found in {','.join(ld_tags)}" if ld_tags else "No LD found"
    total_6xx = sum(subject_counts.values())
    breakdown = ", ".join(f"{t}: {n}" for t, n in sorted(subject_counts.items()))
    breakdown_text = f" ({breakdown})" if breakdown else ""
    low_650 = subject_counts["650"] < 6
    flag = " ⚠ LOW 650 COUNT" if low_650 else ""

    print(f"{oclc} [{genre}]: {ld_text} | 6XX fields: {total_6xx}{breakdown_text} ({elapsed:.2f}s){flag}")

    # store results for excel to parse
    return {
        "oclc": oclc,
        "genre": genre,
        "ld_tags": ",".join(ld_tags),
        "subject_counts": subject_counts,
        "total_6xx": total_6xx,
        "low_650": low_650,
        "time": elapsed,
        "error": "",
    }


# Read OCLC numbers (column A) and genre (column B), skipping blanks and header rows
wb = load_workbook(EXCEL_FILE, read_only=True)
ws = wb.active
records = []
for row in ws.iter_rows(min_col=1, max_col=2, values_only=True):
    value = row[0]
    genre = row[1] if len(row) > 1 and row[1] is not None else ""
    if value is None:
        continue
    number = str(value).strip().removesuffix(".0")
    if number.isdigit():
        records.append((number, str(genre).strip()))
wb.close()

# summary of what was found
print(f"Found {len(records)} OCLC numbers, fetching {MAX_WORKERS} at a time\n")

# process records and report total time
total_start = time.perf_counter()
with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
    results = list(executor.map(process, records))
total = time.perf_counter() - total_start

# close session
for s in all_sessions:
    s.close()

print(f"\nDone. Total time: {total:.2f}s")

# Write results to Excel
all_tags = sorted({t for r in results for t in r.get("subject_counts", {})})

# Setup workbook, and its headers
out_wb = Workbook()
out_ws = out_wb.active
out_ws.title = "Results"

# Create header as bold
headers = ["OCLC Number", "Genre", "LD Found", "LD Tags", "Total 6XX"] + all_tags + \
          ["Low 650 Flag", "Time (s)", "Error"]
out_ws.append(headers)
for cell in out_ws[1]:
    cell.font = Font(bold=True)

# Make flagged records red f
flag_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

# Write the output row by row
for r in results:
    if r["error"]:
        row = [r["oclc"], r["genre"], "", "", ""] + [""] * len(all_tags) + \
              ["", round(r["time"], 2), r["error"]]
        out_ws.append(row)
        continue

    counts = r["subject_counts"]
    row = [
        r["oclc"],
        r["genre"],
        "Yes" if r["ld_tags"] else "No",
        r["ld_tags"],
        r["total_6xx"],
    ] + [counts.get(t, 0) for t in all_tags] + [
        "LOW" if r["low_650"] else "",
        round(r["time"], 2),
        "",
    ]
    out_ws.append(row)

    if r["low_650"]:
        for cell in out_ws[out_ws.max_row]:
            cell.fill = flag_fill

# Freeze row one
out_ws.freeze_panes = "A2"

# Set max width of columns
for col in out_ws.columns:
    width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
    out_ws.column_dimensions[col[0].column_letter].width = min(width + 2, 80)

# Write the file to disk
out_wb.save(OUTPUT_FILE)

# Displaylocation of output file
print(f"Results saved to {OUTPUT_FILE}")
