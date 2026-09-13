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

DIRECTION_LABELS = {
    "Šiaudinė": "to_Šiaudinė",
    "Katiliai": "to_Katiliai"
}

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

def get_schedule_for_route(route_prefix):
    routes, stops, trips, stop_times, calendar = load_data()

    routes['route_short_name'] = routes['route_short_name'].astype(str)
    stops['stop_name'] = stops['stop_name'].astype(str)

    filtered_routes = routes[routes['route_short_name'].str.startswith(route_prefix)]
    route_ids = filtered_routes['route_id'].unique()
    trips = trips[trips['route_id'].isin(route_ids)]

    enriched = stop_times.merge(trips[['trip_id', 'route_id', 'service_id']], on='trip_id')
    enriched = enriched.merge(calendar, on='service_id', how='left')
    enriched = enriched.merge(routes[['route_id', 'route_short_name']], on='route_id', how='left')
    enriched = enriched.merge(stops[['stop_id', 'stop_name']], on='stop_id', how='left')

    enriched['service_flags'] = enriched.apply(lambda row: ''.join(
        [d for d, active in zip(['D', 'Š', 'S'], [
            any(row[day] for day in ['monday', 'tuesday', 'wednesday', 'thursday', 'friday']),
            row['saturday'], row['sunday']]) if active]
    ), axis=1)

        # Automatically determine the two most common terminal stops for direction inference
    terminal_lookup = stop_times.merge(trips[['trip_id', 'route_id']], on='trip_id')
    terminal_lookup = terminal_lookup.sort_values(by=['trip_id', 'stop_sequence'])
    trip_ends = terminal_lookup.groupby('trip_id').last()
    terminal_stop_ids = trip_ends['stop_id'].value_counts().nlargest(2).index.tolist()
    terminal_names = ['Šiaudinė', 'Katiliai', '1-ieji Rūko Rūdninkai', 'V. Sirokomlės muziejus']

    stop_data = defaultdict(lambda: defaultdict(list))

    for trip_id, trip_group in enriched.groupby('trip_id'):
        trip_group = trip_group.sort_values('stop_sequence').reset_index(drop=True)
        stop_names_seq = trip_group['stop_name'].tolist()

        # Determine which direction this trip is going
        direction_label = None
        if terminal_names[0] in stop_names_seq:
            direction_label = f'to_{terminal_names[0].replace(" ", "_")}'
        elif terminal_names[1] in stop_names_seq:
            direction_label = f'to_{terminal_names[1].replace(" ", "_")}'

        if not direction_label:
            continue

        for i, row in trip_group.iterrows():
            stop_name = row['stop_name']
            route = row['route_short_name']
            try:
                hour, minute, *_ = map(int, row['arrival_time'].split(':'))
                time = f"{hour:02}:{minute:02}"
                for flag in row['service_flags']:
                    stop_data[stop_name][direction_label].append((route, time, flag))
            except:
                continue

    final_data = defaultdict(lambda: defaultdict(list))
    for stop, dir_data in stop_data.items():
        for direction, values in dir_data.items():
            grouped = defaultdict(lambda: defaultdict(set))
            for route, time, flag in values:
                grouped[route][time].add(flag)
            schedule = [(route, grouped[route]) for route in grouped]
            final_data[stop][direction] = schedule

    return final_data

def export_schedule_to_docx(stop, direction, schedule):
    filename = f"Sticker_{stop.replace(' ', '_')}_{direction.replace(' ', '_')}.docx"
    doc = Document()

    title = doc.add_paragraph()
    run = title.add_run(f"{stop} → {direction}")
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
    table.columns[0].width = Inches(0.6)
    table.columns[1].width = Inches(5.8)

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
        formatted = [f"{time} {''.join(sorted(flags))}" for time, flags in sorted_times]
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

# --- Run for route 211 only ---
route_211_data = get_schedule_for_route("211")
for stop, directions in route_211_data.items():
    for direction, sched in directions.items():
        export_schedule_to_docx(stop, direction, sched)

# --- Create ZIP of all generated stickers ---
with zipfile.ZipFile("Directional_Stickers_Final.zip", "w") as zipf:
    for file in os.listdir():
        if file.startswith("Sticker_") and file.endswith(".docx"):
            zipf.write(file)
