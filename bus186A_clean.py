# bus186A_clean.py
import os
import zipfile
import argparse
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from collections import defaultdict
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

GTFS_ZIP = "gtfs.zip"
EXTRACT_DIR = "gtfs_data"
ROUTE_REGEX = r"^\s*186\s*-?\s*A\b"  # tolerant for 186A, 186 A, 186-A

# ----------------- GTFS extraction/reading -----------------
def extract_gtfs():
    if not os.path.exists(EXTRACT_DIR):
        if not os.path.exists(GTFS_ZIP):
            raise FileNotFoundError(f"gtfs.zip not found in {os.getcwd()}")
        with zipfile.ZipFile(GTFS_ZIP, 'r') as zip_ref:
            zip_ref.extractall(EXTRACT_DIR)

def read_table(filename):
    path = os.path.join(EXTRACT_DIR, filename)
    if not os.path.exists(path):
        raise FileNotFoundError(f"{filename} not found under {EXTRACT_DIR}")
    return pd.read_csv(path, dtype=str, keep_default_na=False)

def load_data():
    extract_gtfs()
    routes = read_table("routes.txt")
    stops = read_table("stops.txt")
    trips = read_table("trips.txt")
    stop_times = read_table("stop_times.txt")

    # calendar: prefer calendar.txt, fallback to calendar_dates -> create basic calendar if only dates present
    calendar = None
    cal_path = os.path.join(EXTRACT_DIR, "calendar.txt")
    cal_dates_path = os.path.join(EXTRACT_DIR, "calendar_dates.txt")
    if os.path.exists(cal_path):
        calendar = pd.read_csv(cal_path, dtype=str, keep_default_na=False)
    elif os.path.exists(cal_dates_path):
        # create a simple calendar from calendar_dates (not perfect, but helps avoid crashes)
        cd = pd.read_csv(cal_dates_path, dtype=str, keep_default_na=False)
        # make a very small calendar-like frame: service_id + placeholder weekdays (treat presence as all days)
        service_ids = cd['service_id'].unique()
        calendar = pd.DataFrame({
            'service_id': service_ids,
            'monday': [1]*len(service_ids),
            'tuesday': [1]*len(service_ids),
            'wednesday': [1]*len(service_ids),
            'thursday': [1]*len(service_ids),
            'friday': [1]*len(service_ids),
            'saturday': [0]*len(service_ids),
            'sunday': [0]*len(service_ids),
        })
    else:
        # minimal fallback - empty calendar (will produce no flags)
        calendar = pd.DataFrame(columns=['service_id','monday','tuesday','wednesday','thursday','friday','saturday','sunday'])

    # Add stop sequence sanity
    if 'stop_sequence' not in stop_times.columns:
        stop_times['stop_sequence'] = 0

    return routes, stops, trips, stop_times, calendar

# ----------------- Schedule extraction for 186A -----------------
def get_route_186a_data():
    routes, stops, trips, stop_times, calendar = load_data()

    # normalize columns to str
    for col in ['route_short_name', 'route_long_name']:
        if col in routes.columns:
            routes[col] = routes[col].astype(str)

    stops['stop_name'] = stops['stop_name'].astype(str)

    # find matching routes by regex in short or long name
    mask_short = routes.get('route_short_name', pd.Series()).str.contains(ROUTE_REGEX, case=False, regex=True, na=False)
    long_col = [c for c in routes.columns if 'long' in c.lower()]
    mask_long = False
    if long_col:
        mask_long = routes[long_col[0]].astype(str).str.contains(ROUTE_REGEX, case=False, regex=True, na=False)

    route_matches = routes[mask_short | mask_long]
    if route_matches.empty:
        print(f"⚠️ No routes matched regex {ROUTE_REGEX}. Found route_short_name examples:")
        if 'route_short_name' in routes.columns:
            print(list(routes['route_short_name'].unique())[:40])
        return {}

    route_ids = route_matches['route_id'].unique().tolist()
    trips = trips[trips['route_id'].isin(route_ids)]

    enriched = stop_times.merge(trips[['trip_id','route_id','service_id']], on='trip_id', how='inner')
    enriched = enriched.merge(routes[['route_id','route_short_name']], on='route_id', how='left')
    enriched = enriched.merge(stops[['stop_id','stop_name']], on='stop_id', how='left')
    enriched = enriched.merge(calendar, on='service_id', how='left')

    # compute service flags D/Š/S
    def compute_flags(row):
        flags = []
        try:
            # calendar columns may be strings '1' or '0'
            def is_true(val):
                return str(val).strip() not in ("", "0", "False", "false", "nan")
            if any(is_true(row.get(day, 0)) for day in ['monday','tuesday','wednesday','thursday','friday']):
                flags.append('D')
            if is_true(row.get('saturday', 0)): flags.append('Š')
            if is_true(row.get('sunday', 0)): flags.append('S')
        except Exception:
            pass
        return ''.join(flags)

    enriched['service_flags'] = enriched.apply(compute_flags, axis=1)

    # Use a small custom mapping to help infer directions if available (keeps your knowledge)
    CUSTOM_DIRECTION_MAP = {
        "Stotis": ["Katiliai", "Šiaudinė"],
        "Katiliai": ["Stotis", "Stadionas"],
        "Šiaudinė": ["Stotis", "Stadionas"],
        "Stadionas": ["Stotis", "Šiaudinė"],
    }
    terminal_names = list({t for v in CUSTOM_DIRECTION_MAP.values() for t in v})

    # Group by trip to determine direction
    stop_data = defaultdict(list)  # stop_name -> list of (route_short_name, {time:set(flags)})
    for trip_id, trip_group in enriched.groupby('trip_id'):
        trip_group = trip_group.sort_values(by='stop_sequence').reset_index(drop=True)
        stop_names_seq = trip_group['stop_name'].astype(str).tolist()
        if not stop_names_seq:
            continue

        # try to infer direction label
        direction_label = None
        first_stop = stop_names_seq[0]
        if first_stop in CUSTOM_DIRECTION_MAP:
            for term in CUSTOM_DIRECTION_MAP[first_stop]:
                if term in stop_names_seq:
                    direction_label = f"to_{term.replace(' ','_')}"
                    break
        if not direction_label:
            for term in terminal_names:
                if term in stop_names_seq:
                    direction_label = f"to_{term.replace(' ','_')}"
                    break
        if not direction_label:
            # last fallback: use last stop name
            last_name = trip_group.iloc[-1].get('stop_name')
            if last_name:
                direction_label = f"to_{str(last_name).replace(' ','_')}"

        if not direction_label:
            continue

        route_short = trip_group.iloc[0].get('route_short_name')
        for _, row in trip_group.iterrows():
            arrival = str(row.get('arrival_time','')).strip()
            if not arrival:
                continue
            try:
                parts = arrival.split(':')
                hour = int(parts[0]); minute = int(parts[1])
                time = f"{hour:02}:{minute:02}"
            except Exception:
                continue
            stop_name = row.get('stop_name')
            flags = row.get('service_flags','') or ''
            # We'll collect times under a dict per stop + direction per route
            stop_data[(stop_name, direction_label)].append((route_short, time, flags))

    # convert to final structure: stop -> direction -> [(route, {time:set(flags)})...]
    final = defaultdict(lambda: defaultdict(list))
    for (stop_name, direction_label), items in stop_data.items():
        grouped = defaultdict(lambda: defaultdict(set))
        for route, time, flags in items:
            for f in flags or ['']:
                if f != '':
                    grouped[route][time].add(f)
                else:
                    grouped[route][time]  # ensure time exists (empty set)
        schedule = [(route, grouped[route]) for route in grouped]
        final[stop_name][direction_label] = schedule

    return final

# ----------------- DOCX helpers -----------------
def set_cell_background(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

def set_vertical_alignment(cell, alignment="center"):
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), alignment)
    tcPr.append(vAlign)

def export_schedule_to_docx(stop, direction, schedule, prefix="186A"):
    safe_stop = str(stop).replace(' ','_').replace('/','_')
    safe_dir = str(direction).replace(' ','_').replace('/','_')
    filename = f"Sticker_{prefix}_{safe_stop}_{safe_dir}.docx"
    doc = Document()
    try:
        section = doc.sections[-1]
        section.orientation = 1
        new_w, new_h = section.page_height, section.page_width
        section.page_width = new_w; section.page_height = new_h
    except Exception:
        pass

    title = doc.add_paragraph()
    run = title.add_run(f"{stop} → {direction} ({prefix})")
    run.bold = True; run.font.size = Pt(28)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    date_par = doc.add_paragraph()
    date_par.add_run(f"nuo {datetime.now().strftime('%Y %m %d')}").bold = True
    date_par.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    doc.add_paragraph("|Vilniaus AS:")

    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    try:
        table.columns[0].width = Inches(0.7)
        table.columns[1].width = Inches(6.5)
    except Exception:
        pass

    for route, time_flag_dict in schedule:
        row = table.add_row().cells
        p = row[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(str(route).upper())
        r.font.size = Pt(10); r.bold = True; r.font.color.rgb = RGBColor(255,255,255)
        set_vertical_alignment(row[0])
        set_cell_background(row[0], "001F5B")

        sorted_times = sorted(time_flag_dict.items())
        formatted = [f"{time} {''.join(sorted(flags))}".strip() for time, flags in sorted_times]
        row[1].paragraphs[0].add_run(" ".join(formatted)).font.size = Pt(10)

        spacer = table.add_row().cells
        spacer[0].text = ""; spacer[1].text = ""

    doc.add_paragraph()
    doc.add_paragraph("D - darbo dienomis, Š - šeštadieniais, S - sekmadieniais")
    footer = doc.add_paragraph()
    footer.add_run("Aktualūs grafikai ir naujienos ").bold = True
    link = footer.add_run("WWW.VRAP.LT")
    link.bold = True; link.font.color.rgb = RGBColor(0,0,255); link.underline = True
    footer.add_run(" arba TRAFI programėlėje")

    doc.save(filename)
    print(f"✅ Saved {filename}")
    return filename

# ----------------- GUI -----------------
def create_gui(stop_data):
    # stop_data: {stop: {direction: schedule}}
    root = tk.Tk()
    root.title("Route 186A Timetable")
    root.geometry("900x600")
    ttk.Label(root, text="🚌 Route 186A Bus Timetable", font=("Segoe UI", 16, "bold")).pack(pady=8)

    top_frame = ttk.Frame(root); top_frame.pack(pady=4)
    ttk.Label(top_frame, text="Stop:").pack(side=tk.LEFT, padx=6)
    stop_var = tk.StringVar()
    stops_list = sorted(stop_data.keys())
    stop_menu = ttk.Combobox(top_frame, textvariable=stop_var, values=stops_list, width=50, state="readonly")
    stop_menu.pack(side=tk.LEFT)
    if stops_list:
        stop_var.set(stops_list[0])

    tree = ttk.Treeview(root, columns=("Direction","Route","Times"), show="headings", height=20)
    tree.heading("Direction", text="Direction"); tree.heading("Route", text="Route"); tree.heading("Times", text="Arrival Times")
    tree.column("Direction", width=180); tree.column("Route", width=80); tree.column("Times", width=560)
    tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def show_schedule():
        sel = stop_var.get()
        tree.delete(*tree.get_children())
        if not sel:
            return
        for direction, schedule in stop_data.get(sel, {}).items():
            for route, time_dict in schedule:
                formatted = ", ".join([f"{t}{''.join(sorted(flags))}" for t,flags in sorted(time_dict.items())])
                tree.insert("", "end", values=(direction, str(route).upper(), formatted))

    def export_current():
        sel = stop_var.get()
        if not sel:
            messagebox.showerror("No stop", "Please select a stop first")
            return
        for direction, schedule in stop_data.get(sel, {}).items():
            export_schedule_to_docx(sel, direction, schedule)

    btn_frame = ttk.Frame(root); btn_frame.pack(pady=4)
    ttk.Button(btn_frame, text="Show Timetable", command=show_schedule).pack(side=tk.LEFT, padx=6)
    ttk.Button(btn_frame, text="Export Selected Stop (DOCX)", command=export_current).pack(side=tk.LEFT, padx=6)

    root.mainloop()

# ----------------- Main / Batch run -----------------
def main(batch=False):
    print("Starting 186A extraction...")
    data = get_route_186a_data()
    if not data:
        print("No data found for 186A. Exiting.")
        return

    print(f"Found {len(data)} stops with directional schedules.")
    if batch:
        out_files = []
        for stop, dirs in data.items():
            for direction, sched in dirs.items():
                fname = export_schedule_to_docx(stop, direction, sched, prefix="186A")
                out_files.append(fname)
        # zip results
        zipname = "186A_Stickers.zip"
        with zipfile.ZipFile(zipname, "w") as zf:
            for f in out_files:
                zf.write(f)
        print(f"✅ Zipped {len(out_files)} files into {zipname}")
    else:
        create_gui(data)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate stickers for route 186A (GTFS feed).")
    parser.add_argument("--batch", action="store_true", help="Export all stickers to DOCX and ZIP them (no GUI).")
    args = parser.parse_args()
    try:
        main(batch=args.batch)
    except Exception as e:
        # surface a helpful error message
        print("ERROR:", e)
        try:
            messagebox.showerror("Error", str(e))
        except Exception:
            pass
