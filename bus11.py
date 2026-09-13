import os
import zipfile
import pandas as pd
from datetime import datetime
from collections import defaultdict
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

GTFS_ZIP = "gtfs.zip"
EXTRACT_DIR = "gtfs_data"

def extract_gtfs():
    if not os.path.exists(EXTRACT_DIR):
        with zipfile.ZipFile(GTFS_ZIP, 'r') as zip_ref:
            zip_ref.extractall(EXTRACT_DIR)

def load_data():
    extract_gtfs()
    routes = pd.read_csv(f"{EXTRACT_DIR}/routes.txt")
    stops = pd.read_csv(f"{EXTRACT_DIR}/stops.txt")
    trips = pd.read_csv(f"{EXTRACT_DIR}/trips.txt")
    stop_times = pd.read_csv(f"{EXTRACT_DIR}/stop_times.txt")
    calendar = pd.read_csv(f"{EXTRACT_DIR}/calendar.txt")

    stop_times_sorted = stop_times.sort_values(by=['trip_id', 'stop_sequence'])
    first_stops = stop_times_sorted.groupby('trip_id').first().reset_index()[['trip_id', 'stop_id']]
    last_stops = stop_times_sorted.groupby('trip_id').last().reset_index()[['trip_id', 'stop_id']]
    first_stops.columns = ['trip_id', 'first_stop_id']
    last_stops.columns = ['trip_id', 'last_stop_id']

    trips = trips.merge(first_stops, on='trip_id').merge(last_stops, on='trip_id')
    trips = trips.merge(stops[['stop_id', 'stop_name']], left_on='last_stop_id', right_on='stop_id', how='left')
    trips = trips.rename(columns={'stop_name': 'direction_stop_name'})

    return routes, stops, trips, stop_times, calendar

def get_schedule_for_route(route_regex):
    """
    route_regex: a regex string used to match route_short_name or route_long_name (case-insensitive).
    Returns final_data: { stop_name: { direction_label: [(route_short_name, {time: set(flags)})...] } }
    """
    routes, stops, trips, stop_times, calendar = load_data()

    routes['route_short_name'] = routes['route_short_name'].astype(str)
    stops['stop_name'] = stops['stop_name'].astype(str)

    # match either short or long name using the provided regex
    mask_short = routes['route_short_name'].str.contains(route_regex, case=False, regex=True, na=False)
    long_col = routes.columns[ routes.columns.str.contains('long', case=False) ]
    if len(long_col) > 0:
        route_long_name_col = long_col[0]
    else:
        route_long_name_col = None

    mask_long = False
    if route_long_name_col:
        mask_long = routes[route_long_name_col].astype(str).str.contains(route_regex, case=False, regex=True, na=False)

    filtered_routes = routes[mask_short | mask_long]
    if filtered_routes.empty:
        print(f"⚠️ No routes matched regex: {route_regex}")
        return {}

    route_ids = filtered_routes['route_id'].unique()
    trips = trips[trips['route_id'].isin(route_ids)]

    enriched = stop_times.merge(trips[['trip_id', 'route_id', 'service_id']], on='trip_id', how='left')
    enriched = enriched.merge(calendar, on='service_id', how='left')
    enriched = enriched.merge(routes[['route_id', 'route_short_name']], on='route_id', how='left')
    enriched = enriched.merge(stops[['stop_id', 'stop_name']], on='stop_id', how='left')

    # compute service flags D / Š / S
    def compute_flags(row):
        flags = []
        try:
            if any(row.get(day, 0) for day in ['monday', 'tuesday', 'wednesday', 'thursday', 'friday']):
                flags.append('D')
            if row.get('saturday', 0):
                flags.append('Š')
            if row.get('sunday', 0):
                flags.append('S')
        except Exception:
            pass
        return ''.join(flags)

    enriched['service_flags'] = enriched.apply(compute_flags, axis=1)

    # Determine common terminals to help infer direction (fallback)
    terminal_lookup = stop_times.merge(trips[['trip_id', 'route_id']], on='trip_id')
    terminal_lookup = terminal_lookup.sort_values(by=['trip_id', 'stop_sequence'])
    trip_ends = terminal_lookup.groupby('trip_id').last()
    terminal_stop_ids = trip_ends['stop_id'].value_counts().nlargest(2).index.tolist()

    # CUSTOM_DIRECTION_MAP copied from your original mapping (keeps your domain knowledge)
    CUSTOM_DIRECTION_MAP = {
        "13-asis kilometras": ["Stotis", "Stadionas"],
        "16-asis kilometras": ["1-ieji Rūko Rūdninkai", "V. Sirokomlės muziejus"],
        "Laibiškių": ["Stotis", "Stadionas"],
        "Lazdynėliai": ["Stotis", "Stadionas"],
        "Vaduvos st.": ["Stotis", "Stadionas"],
        "Stotis": ["Katiliai", "Šiaudinė"],
        "Savanorių pr.": ["Stotis", "Stadionas"],
        "Katiliai": ["Stotis", "Stadionas"],
        "Šiaudinė": ["Stotis", "Stadionas"],
        "Stadionas": ["Stotis", "Šiaudinė"],
        "1-ieji Rūko Rūdninkai": ["Stotis", "Stadionas"],
        "V. Sirokomlės muziejus": ["Stotis", "Stadionas"],
        "Liepkalnio": ["Stotis", "Stadionas"],
        "Salininkai": ["Stotis", "Stadionas"],
        "Drogžiai": ["Stotis", "Stadionas"]
    }
    terminal_names = list(set([term for pairs in CUSTOM_DIRECTION_MAP.values() for term in pairs]))

    stop_data = defaultdict(lambda: defaultdict(list))

    # group by trip and infer direction using CUSTOM_DIRECTION_MAP or terminal_names
    for trip_id, trip_group in enriched.groupby('trip_id'):
        trip_group = trip_group.sort_values('stop_sequence').reset_index(drop=True)
        stop_names_seq = trip_group['stop_name'].astype(str).tolist()
        if not stop_names_seq:
            continue

        direction_label = None

        # First try: if the trip's first stop is in CUSTOM_DIRECTION_MAP, use that mapping
        first_stop = stop_names_seq[0]
        if first_stop in CUSTOM_DIRECTION_MAP:
            # pick the first mapped terminal that exists in the trip stops
            for label in CUSTOM_DIRECTION_MAP[first_stop]:
                if label in stop_names_seq:
                    direction_label = f"to_{label.replace(' ', '_')}"
                    break

        # Fallback: check terminal names anywhere in sequence
        if not direction_label:
            for term in terminal_names:
                if term in stop_names_seq:
                    direction_label = f"to_{term.replace(' ', '_')}"
                    break

        # Last fallback: use most common terminal stop ids (if any)
        if not direction_label and terminal_stop_ids:
            last_stop_id = trip_group.iloc[-1].get('stop_id')
            # map id -> name if available
            last_name = None
            try:
                last_name = trip_group.iloc[-1].get('stop_name')
            except Exception:
                last_name = None
            if last_name:
                direction_label = f"to_{str(last_name).replace(' ', '_')}"

        if not direction_label:
            # can't infer direction for this trip: skip it
            continue

        for _, row in trip_group.iterrows():
            stop_name = row.get('stop_name')
            route = row.get('route_short_name')
            arrival = str(row.get('arrival_time', '')).strip()
            if not arrival or pd.isna(arrival):
                continue
            try:
                parts = arrival.split(':')
                hour = int(parts[0])
                minute = int(parts[1])
                time = f"{hour:02}:{minute:02}"
            except Exception:
                # skip malformed times
                continue
            flags = row.get('service_flags', '') or ''
            for flag in flags:
                stop_data[stop_name][direction_label].append((route, time, flag))
            if not flags:
                # still append without any flag so time appears
                stop_data[stop_name][direction_label].append((route, time, ''))

    # convert into the grouped format you used earlier
    final_data = defaultdict(lambda: defaultdict(list))
    for stop, dir_data in stop_data.items():
        for direction, values in dir_data.items():
            grouped = defaultdict(lambda: defaultdict(set))
            for route, time, flag in values:
                grouped[route][time].add(flag)
            schedule = [(route, grouped[route]) for route in grouped]
            final_data[stop][direction] = schedule

    return final_data

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

def export_schedule_to_docx(stop, direction, schedule):
    safe_stop = stop.replace(' ', '_').replace('/', '_')
    safe_dir = direction.replace(' ', '_').replace('/', '_')
    filename = f"Sticker_186A_{safe_stop}_{safe_dir}.docx"
    doc = Document()
    # try to set landscape
    try:
        section = doc.sections[-1]
        section.orientation = 1  # Landscape (word uses constants; this usually works)
        new_width, new_height = section.page_height, section.page_width
        section.page_width = new_width
        section.page_height = new_height
    except Exception:
        pass

    title = doc.add_paragraph()
    run = title.add_run(f"{stop} → {direction} (186A)")
    run.bold = True
    run.font.size = Pt(36)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    date_par = doc.add_paragraph()
    date_run = date_par.add_run(f"nuo {datetime.now().strftime('%Y %m %d')}")
    date_run.bold = True
    date_run.font.size = Pt(12)
    date_run.font.color.rgb = RGBColor(255, 0, 0)
    date_par.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    doc.add_paragraph("|Vilniaus AS:")

    table = doc.add_table(rows=0, cols=2)
    table.autofit = False
    # approximate widths (docx may not strictly enforce these)
    try:
        table.columns[0].width = Inches(0.6)
        table.columns[1].width = Inches(5.8)
    except Exception:
        pass

    for route, time_flag_dict in schedule:
        row = table.add_row().cells
        p = row[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(str(route).upper())
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(255, 255, 255)
        run.bold = True
        set_vertical_alignment(row[0])
        set_cell_background(row[0], "001F5B")

        sorted_times = sorted(time_flag_dict.items())
        formatted = [f"{time} {''.join(sorted(flags))}".strip() for time, flags in sorted_times]
        row[1].paragraphs[0].add_run(" ".join(formatted)).font.size = Pt(10)

        spacer = table.add_row().cells
        spacer[0].text = ""
        spacer[1].text = ""

    doc.add_paragraph()
    doc.add_paragraph("D - darbo dienomis, Š - šeštadieniais, S - sekmadieniais")

    footer = doc.add_paragraph()
    footer.add_run("Aktualūs grafikai ir naujienos ").bold = True
    link = footer.add_run("WWW.VRAP.LT")
    link.bold = True
    link.font.color.rgb = RGBColor(0, 0, 255)
    link.underline = True
    footer.add_run(" arba TRAFI programėlėje")

    doc.save(filename)
    print(f"✅ Saved {filename}")

# ---- Run for route 186A (tolerant regex: matches "186A", "186 A", "186-A") ----
if __name__ == "__main__":
    # tolerant regex for 186A variants
    route_regex = r"^\s*186\s*-?\s*A\b"
    print("Starting extraction and schedule generation for route regex:", route_regex)
    data = get_schedule_for_route(route_regex)
    print(f"Found {len(data)} stops with directional schedules.")
    # export each direction's sticker
    exported_files = []
    for stop, directions in data.items():
        for direction, sched in directions.items():
            export_schedule_to_docx(stop, direction, sched)
            exported_files.append(f"Sticker_186A_{stop.replace(' ', '_')}_{direction.replace(' ', '_')}.docx")

    # create a zip with all generated 186A stickers
    zip_name = "186A_Directional_Stickers.zip"
    with zipfile.ZipFile(zip_name, "w") as zipf:
        for file in os.listdir():
            if file.startswith("Sticker_186A_") and file.endswith(".docx"):
                zipf.write(file)
    print(f"✅ All stickers zipped into {zip_name}")
