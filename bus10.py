
import os
import zipfile
from datetime import datetime

import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ---------- CONFIG ----------
GTFS_ZIP     = "gtfs.zip"
EXTRACT_DIR  = "gtfs_data"
OUTPUT_DIR   = "bus10"

# Font sizes
TITLE_PT     = 28   # Stop name
DATE_PT      = 12   # Date line
DIR_PT       = 18   # Direction header
TIME_PT      = 12   # Table text

# Service-flag color
FLAG_COLOR   = RGBColor(200, 0, 0)  # red for D/Š/S
# ----------------------------

def extract_gtfs():
    """Unzip GTFS feed if not already extracted."""
    if not os.path.isdir(EXTRACT_DIR):
        print(f"Extracting {GTFS_ZIP} → {EXTRACT_DIR}")
        with zipfile.ZipFile(GTFS_ZIP, "r") as z:
            z.extractall(EXTRACT_DIR)

def load_gtfs():
    extract_gtfs()
    routes     = pd.read_csv(f"{EXTRACT_DIR}/routes.txt",    dtype=str)
    trips      = pd.read_csv(f"{EXTRACT_DIR}/trips.txt",     dtype=str)
    stop_times = pd.read_csv(f"{EXTRACT_DIR}/stop_times.txt",dtype=str)
    calendar   = pd.read_csv(f"{EXTRACT_DIR}/calendar.txt",  dtype=int)
    stops      = pd.read_csv(f"{EXTRACT_DIR}/stops.txt",     dtype=str)
    return routes, trips, stop_times, calendar, stops

def compute_flags(row):
    flags = []
    if any(row[wd] for wd in ["monday","tuesday","wednesday","thursday","friday"]):
        flags.append("D")
    if row["saturday"]:
        flags.append("Š")
    if row["sunday"]:
        flags.append("S")
    return "".join(flags)

def build_by_direction():
    routes, trips, stop_times, cal, stops = load_gtfs()

    # Only route 211 (not 211A)
    routes = routes[routes["route_short_name"] == "211"]
    trip_ids = trips[trips["route_id"].isin(routes["route_id"])]["trip_id"]
    stps = stop_times[stop_times["trip_id"].isin(trip_ids)]

    # Merge calendar to get service days
    tc = trips.merge(cal, on="service_id", how="left")[["trip_id","trip_headsign",
        "monday","tuesday","wednesday","thursday","friday","saturday","sunday"]]
    stps = stps.merge(tc, on="trip_id", how="left")
    stps = stps.merge(stops[["stop_id","stop_name"]], on="stop_id", how="left")

    # Precompute flags and time
    stps["flags"] = stps.apply(compute_flags, axis=1)
    stps["time"]  = stps["departure_time"].str.slice(0,5)

    # Group: { stop_name: { headsign: [ {time,flags}, ... ] } }
    schedule = {}
    for (stop, hs), grp in stps.groupby(["stop_name","trip_headsign"]):
        entries = (
            grp[["time","flags"]]
            .drop_duplicates()
            .sort_values("time")
            .to_dict("records")
        )
        schedule.setdefault(stop, {})[hs] = entries

    return schedule

def make_docx(stop, headsign, entries, out_path):
    doc = Document()
    doc.styles["Normal"].font.name = "Arial"

    # Title
    p1 = doc.add_paragraph()
    r1 = p1.add_run(stop)
    r1.bold = True
    r1.font.size = Pt(TITLE_PT)
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Date
    p_date = doc.add_paragraph()
    r_date = p_date.add_run(f"Dated: {datetime.now().strftime('%Y-%m-%d')}")
    r_date.font.size = Pt(DATE_PT)
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # Direction header
    p2 = doc.add_paragraph()
    r2 = p2.add_run(f"Towards {headsign}")
    r2.bold = True
    r2.font.size = Pt(DIR_PT)
    p2.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Table of times
    table = doc.add_table(rows=1, cols=2)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Time"
    hdr_cells[1].text = "Days"
    for cell in hdr_cells:
        for run in cell.paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(TIME_PT)

    for ent in entries:
        row_cells = table.add_row().cells
        row_cells[0].text = ent["time"]
        run = row_cells[1].paragraphs[0].add_run(ent["flags"])
        run.font.color = FLAG_COLOR
        run.font.size = Pt(TIME_PT)

    doc.save(out_path)

def main():
    schedule = build_by_direction()
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    for stop, dirs in schedule.items():
        for hs, entries in dirs.items():
            safe_stop = stop.replace(" ", "_")
            safe_dir  = hs.replace(" ", "_")
            fname = f"{safe_stop}_{safe_dir}.docx"
            path  = os.path.join(OUTPUT_DIR, fname)
            make_docx(stop, hs, entries, path)
            print("Created:", fname)

    print(f"\nDone! Check the '{OUTPUT_DIR}' folder for your two stickers per stop.")

if __name__ == "__main__":
    main()
