import json
import re
from datetime import date

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import gspread
from google.oauth2.service_account import Credentials

# =========================================================================
# CONFIG
#
# "Running Shrimp Farms - KMN" is a VIEW-ONLY, kiosk/TV style display app.
# It shares the same Google Sheet and "Customer List.xlsx" as the other
# apps in this family (data-entry app, full manager app, Marketing
# Manager & Technician view). It performs NO writes to the Google Sheet.
#
# What it shows:
#   1) A carousel (one farm per slide) of every Customer + Farm that is
#      currently "Running" (i.e. NOT every pond on that farm is at Full
#      Harvest) -- Customer Name, Farm Name with Code, and that farm's
#      Pond Layout. Each slide's background is tinted with a color for
#      that farm's Zone (from Customer List.xlsx), and each pond box
#      keeps the same status colors used across this app family (blue =
#      Running, yellow = Partial H, green = Full H, gray = Soon to be).
#      Behind that zone tint, each slide also shows a blurred satellite
#      snapshot of the farm's actual map location (pulled from the same
#      Locations Google Sheet the "Farm Map" app uses). Farms with no
#      matching location simply fall back to the zone-tint-only
#      background.
#      CHANGED: the display now OPENS IN FULL SCREEN by default, has a
#      Zone selector (All Zones / one Zone), and the slides move ONLY
#      with the Back / Next buttons (or swipe / arrow keys) -- there is no
#      automatic rotation any more. It is also phone-friendly.
#   2) CHANGED: the old "Harvest Updates" line was removed. In its place
#      a control bar shows the Active Farms count for the selected Zone
#      ("Active Farms in Zone X: N"), or "All Running Farms: N" when no
#      Zone is selected, between the Back and Next buttons.
#   3) A "⛶ Full Screen" / "Exit Full Screen" button in the top bar.
#   4) A "Zone wise Running Farms - Live Display" section (below the
#      carousel) -- a plain, non-rotating table, grouped by Zone, listing
#      every currently Running farm with its Vannamei Ponds and Monodon
#      Ponds laid out side by side, each pond box using the same status
#      colors, DOC Today values, Issues and WQ Special Cases shown
#      elsewhere in this app family.
#
# All of the carousel/zone/full-screen behaviour runs client-side in a
# single self-contained HTML/CSS/JS component
# (streamlit.components.v1.html) -- Python only computes the data once
# per page load/refresh. The Zone wise section below it is plain
# server-rendered HTML (st.markdown), not part of that component.
# =========================================================================
st.set_page_config(page_title="Running Shrimp Farms - KMN", layout="wide", page_icon="🎡")

CUSTOMER_FILE = "Customer List.xlsx"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Must match app.py's COLUMN_ORDER exactly, since all apps in this family
# read/write the same Google Sheet.
COLUMN_ORDER = [
    "Timestamp", "Customer", "Farm Name with Code", "Zone", "Area",
    "Pond Number", "Date", "Species Culture", "Cycle Type",
    "DOC", "Density", "Feed Per Day", "ABW",
    "Expect Harvest (KG)", "Survival QTY",
    "Issues", "Water Color", "Grade", "Remark", "Technician",
    "Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW",
    "Harvest Date 2", "Harvest Type 2", "Harvest KG 2", "Harvest ABW 2",
    "Deleted",
]

# ---- Display timing (seconds).
# NOTE: CAROUSEL_SECONDS and HARVEST_UPDATE_SECONDS are no longer used
# (slides now move only with Back / Next; the Harvest Updates line was
# removed). They are left here untouched.
CAROUSEL_SECONDS = 4          # (unused) how long each farm's Pond Layout slide stayed on screen
HARVEST_UPDATE_SECONDS = 5    # (unused) how long each Harvest Updates item stayed on screen
DATA_REFRESH_SECONDS = 2000    # how often the whole page reloads to pull fresh sheet data

# ---- Zone color palette -- cycles if there are more zones than colors.
ZONE_PALETTE = [
    "#2563eb", "#dc2626", "#059669", "#d97706", "#7c3aed",
    "#0891b2", "#db2777", "#65a30d", "#ea580c", "#4f46e5",
]

_CUSTOMER_CODE_COLUMN_CANDIDATES = [
    "Customer Code", "Customer ID", "Customer Code with Code", "Code", "Cust Code",
]

# ---- Farm location lookup (same public "Locations" Google Sheet the
# "Farm Map" app in this family reads -- Customer ID / Customer Name /
# Farm Name / Location, where Location is either "lat, lon" or a WKT
# Polygon string). Used ONLY to fetch a static satellite snapshot for the
# blurred slide background below; nothing else about this app changes.
LOCATIONS_SHEET_ID = "1v2qTD5iUtdjFTixt9VZ1vM0dZPnyEVz4AYHtILVJi0A"
LOCATIONS_GID = "0"
LOCATIONS_CSV_URL = (
    f"https://docs.google.com/spreadsheets/d/{LOCATIONS_SHEET_ID}"
    f"/export?format=csv&gid={LOCATIONS_GID}"
)
# How wide an area (in degrees) to capture around each farm's point for
# the background snapshot -- small enough to stay zoomed in on the farm,
# large enough that panning/precision differences still land inside frame.
MAP_BBOX_SPAN_DEG = 0.005
MAP_IMAGE_WIDTH = 900
MAP_IMAGE_HEIGHT = 600
MAP_IMAGE_SIZE = f"{MAP_IMAGE_WIDTH},{MAP_IMAGE_HEIGHT}"
# Extra padding (in degrees) added around a farm's own polygon boundary so
# the outline isn't cropped flush against the image edge.
MAP_POLYGON_PADDING_DEG = 0.002

st.markdown("<h1 style='text-align: center;'>Shrimp FarmFlow - KMN</h1>", unsafe_allow_html=True)
st.subheader("🎡 Running Farms — Live Display")
st.markdown("---")

# =========================================================================
# GOOGLE SHEETS BACKEND -- read-only, same pattern as the Marketing
# Manager view app. This app never writes to the Google Sheet.
# =========================================================================
def _gsheet_configured():
    return "gcp_service_account" in st.secrets and "gsheet" in st.secrets and "sheet_id" in st.secrets["gsheet"]

@st.cache_resource(show_spinner=False)
def get_worksheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet_id = st.secrets["gsheet"]["sheet_id"]
    worksheet_name = st.secrets["gsheet"].get("worksheet_name", "WaterQualityData")
    sh = client.open_by_key(sheet_id)
    return sh.worksheet(worksheet_name)

def load_data():
    """Always reads fresh from the Google Sheet (no caching), so the
    display always shows the latest saved records. Soft-deleted rows
    (Deleted = Yes) are filtered out, same as the other apps."""
    ws = get_worksheet()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for c in COLUMN_ORDER:
        if c not in df.columns:
            df[c] = ""
    if "Harvest Status" not in df.columns:
        df["Harvest Status"] = ""
    if "Harvest Status 2" not in df.columns:
        df["Harvest Status 2"] = ""
    # "WQ Special Cases" already exists as its own column in the Sheet
    # (outside COLUMN_ORDER, same pattern as Harvest Status), kept here so
    # the Pond Layout carousel below can flag a pond whose latest record
    # has text in this column, same as the Marketing Manager app.
    if "WQ Special Cases" not in df.columns:
        df["WQ Special Cases"] = ""
    if len(df) > 0:
        df = df[COLUMN_ORDER + ["Harvest Status", "Harvest Status 2", "WQ Special Cases"]]
    df = df.astype(str).replace("nan", "")
    if "Deleted" in df.columns:
        is_deleted = df["Deleted"].astype(str).str.strip().str.lower().isin(["yes", "true", "1"])
        df = df[~is_deleted].reset_index(drop=True)
    return df

if not _gsheet_configured():
    st.error("❌ Google Sheets is not configured yet. This app needs the same "
             "`.streamlit/secrets.toml` (the `[gcp_service_account]` and `[gsheet]` sections) "
             "used by the other apps in this family.")
    st.stop()

try:
    get_worksheet()
except Exception as e:
    st.error(f"❌ Could not connect to the Google Sheet. Check your secrets and sharing settings.\n\n{e}")
    st.stop()

# =========================================================================
# LOAD CUSTOMER LIST (for Zone + Customer Code lookups)
# =========================================================================
@st.cache_data
def load_customer_data():
    return pd.read_excel(CUSTOMER_FILE)

try:
    customer_df = load_customer_data()
except Exception as e:
    st.error(f"❌ Could not load '{CUSTOMER_FILE}'. Make sure it's in the app folder. ({e})")
    st.stop()

REQUIRED_COLS = ["Customer Name", "Farm Name with Code", "Zone", "Area"]
missing_cols = [c for c in REQUIRED_COLS if c not in customer_df.columns]
if missing_cols:
    st.error(f"❌ 'Customer List.xlsx' is missing required column(s): {', '.join(missing_cols)}")
    st.stop()

for _col in REQUIRED_COLS:
    customer_df[_col] = customer_df[_col].apply(
        lambda v: "" if pd.isna(v) else (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))
    )

# =========================================================================
# LOAD FARM LOCATIONS (for the blurred slide background only)
#
# Ported from the "Farm Map" app's own loader/parser so a farm's point
# (or polygon centroid) can be turned into a small satellite snapshot.
# This is read-only, best-effort: if the sheet can't be reached, farms
# just fall back to the existing zone-tint-only background -- nothing
# else in this app is affected.
# =========================================================================
@st.cache_data(ttl=300, show_spinner=False)
def load_farm_locations():
    try:
        loc_df = pd.read_csv(LOCATIONS_CSV_URL)
        loc_df.columns = [c.strip() for c in loc_df.columns]
        return loc_df
    except Exception:
        return pd.DataFrame(columns=["Customer ID", "Customer Name", "Farm Name", "Location"])

def parse_location(location):
    """
    Parses the Locations sheet's Location cell in either of two formats:
      - "lat, lon"                                    -> plain point
      - "Polygon ((lon lat, lon lat, ...))"            -> WKT polygon ring
    Returns (lat, lon, polygon):
      - lat, lon: the point, or the polygon's centroid, to center the
        snapshot on -- or (None, None) if the value can't be parsed.
      - polygon: list of (lat, lon) tuples for the ring, if the value was
        a WKT polygon; otherwise None. Used to zoom tight to the farm's
        actual boundary and to draw its outline on the snapshot.
    """
    if not isinstance(location, str):
        return None, None, None
    location = location.strip()

    if location.lower().startswith("polygon"):
        coords_match = re.search(r"\(\(([^)]+)\)\)", location)
        if not coords_match:
            return None, None, None
        points = []
        for pair in coords_match.group(1).split(","):
            parts = pair.strip().split()
            if len(parts) != 2:
                continue
            try:
                lon, lat = float(parts[0]), float(parts[1])
                points.append((lat, lon))
            except ValueError:
                continue
        if not points:
            return None, None, None
        avg_lat = sum(p[0] for p in points) / len(points)
        avg_lon = sum(p[1] for p in points) / len(points)
        return avg_lat, avg_lon, points

    match = re.match(r"\s*(-?\d+\.?\d*)\s*,\s*(-?\d+\.?\d*)\s*", location)
    if not match:
        return None, None, None
    return float(match.group(1)), float(match.group(2)), None

def build_farm_location_lookup():
    """Returns {customer_id_code (upper): (lat, lon, polygon_or_None)},
    built once per (cached) Locations sheet load."""
    loc_df = load_farm_locations()
    lookup = {}
    if loc_df.empty or "Location" not in loc_df.columns or "Customer ID" not in loc_df.columns:
        return lookup
    for _, row in loc_df.iterrows():
        code = str(row.get("Customer ID", "")).strip().upper()
        if not code:
            continue
        lat, lon, polygon = parse_location(row.get("Location", ""))
        if lat is not None and lon is not None:
            lookup[code] = (lat, lon, polygon)
    return lookup

def build_map_image_url(lat, lon, polygon=None):
    """A single static satellite snapshot (no Leaflet/JS map needed) via
    ArcGIS's MapServer 'export' endpoint. Used as a plain <img> so it can
    be styled with a CSS filter client-side.

    When a farm has an actual polygon boundary, the snapshot is zoomed to
    that polygon's own bounding box (plus a little padding) instead of a
    fixed-size box around its centroid -- so a large farm doesn't get
    cropped and a small one doesn't drown in unrelated surroundings.
    Falls back to the fixed MAP_BBOX_SPAN_DEG box for plain point
    locations. Returns (image_url, bbox) where bbox is
    (min_lon, min_lat, max_lon, max_lat), needed later to draw the
    polygon outline in the exact right place on top of the image.
    """
    if polygon:
        lats = [p[0] for p in polygon]
        lons = [p[1] for p in polygon]
        lat_span = max(lats) - min(lats)
        lon_span = max(lons) - min(lons)
        # Pad generously (40% of the shape's own extent on each side, with
        # a floor for tiny/thin polygons) so the boundary sits comfortably
        # inside the frame with breathing room, instead of touching --
        # or nearly filling -- the image edges.
        pad_lat = max(lat_span * 0.9, MAP_POLYGON_PADDING_DEG)
        pad_lon = max(lon_span * 0.9, MAP_POLYGON_PADDING_DEG)
        min_lat, max_lat = min(lats) - pad_lat, max(lats) + pad_lat
        min_lon, max_lon = min(lons) - pad_lon, max(lons) + pad_lon
    else:
        min_lon, max_lon = lon - MAP_BBOX_SPAN_DEG, lon + MAP_BBOX_SPAN_DEG
        min_lat, max_lat = lat - MAP_BBOX_SPAN_DEG, lat + MAP_BBOX_SPAN_DEG

    url = (
        "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/export"
        f"?bbox={min_lon},{min_lat},{max_lon},{max_lat}&bboxSR=4326&size={MAP_IMAGE_SIZE}"
        "&format=png32&transparent=false&f=image"
    )
    return url, (min_lon, min_lat, max_lon, max_lat)

# =========================================================================
# HELPERS -- ported from the Marketing Manager / full manager app's Pond
# Layout logic, kept self-contained here.
# =========================================================================
def _farm_zone(customer, farm):
    match = customer_df[
        (customer_df["Customer Name"] == customer) & (customer_df["Farm Name with Code"] == farm)
    ]
    if len(match) > 0:
        return str(match.iloc[0].get("Zone", "")).strip()
    return ""

def build_zone_colors(zones):
    zones = sorted({z for z in zones if z})
    return {z: ZONE_PALETTE[i % len(ZONE_PALETTE)] for i, z in enumerate(zones)}

def _pond_status(prow, has_partial_history):
    h_type = (str(prow.get("Harvest Type 2", "")).strip() or str(prow.get("Harvest Type", "")).strip()).lower()
    if "full" in h_type:
        return "Full H"
    elif "partial" in h_type or has_partial_history:
        return "Partial H"
    elif str(prow.get("Cycle Type", "")).strip() == "Soon to be":
        return "Soon to be"
    return "Running"

def _pond_color(status):
    return {
        "Partial H": "#f59e0b",
        "Full H": "#22c55e",
        "Soon to be": "#94a3b8",
    }.get(status, "#3b82f6")

def _species_letter(species):
    s = str(species).strip().lower()
    if "vannamei" in s:
        return "V"
    elif "monodon" in s:
        return "M"
    return ""

def _doc_today(row):
    if str(row.get("Cycle Type") or "").strip() == "Soon to be":
        return 0
    parsed = pd.to_datetime(row.get("Date"), errors="coerce")
    if pd.isna(parsed):
        return None
    try:
        doc_num = int(float(row.get("DOC")))
    except (TypeError, ValueError):
        return None
    t2 = str(row.get("Harvest Type 2", "")).strip().lower()
    t1 = str(row.get("Harvest Type", "")).strip().lower()
    full_date_str = ""
    if "full" in t2:
        full_date_str = str(row.get("Harvest Date 2", "")).strip()
    elif "full" in t1:
        full_date_str = str(row.get("Harvest Date", "")).strip()
    if full_date_str:
        full_date = pd.to_datetime(full_date_str, errors="coerce")
        if pd.notna(full_date):
            return doc_num + (full_date - parsed).days
    return doc_num + (pd.Timestamp(date.today()) - parsed).days

def _customer_code_lookup():
    lookup = {}
    for _, row in customer_df.drop_duplicates(subset=["Customer Name", "Farm Name with Code"]).iterrows():
        code = ""
        for cand in _CUSTOMER_CODE_COLUMN_CANDIDATES:
            if cand in customer_df.columns:
                val = str(row.get(cand, "")).strip()
                if val and val.lower() != "nan":
                    code = val
                    break
        lookup[(row["Customer Name"], row["Farm Name with Code"])] = code
    return lookup

# =========================================================================
# BUILD "RUNNING FARMS" DATA FOR THE CAROUSEL
#
# A farm counts as "Running" for this display when NOT every pond it has
# a saved record for is currently at Full Harvest (same rule the Last
# Visit Date Report uses elsewhere in this app family) -- individual
# ponds within a running farm can still show Partial H / Full H / Soon
# to be, they just aren't ALL at Full H yet.
# =========================================================================
def build_running_farms(df):
    required = {"Customer", "Farm Name with Code", "Pond Number", "Date", "Harvest Type", "Harvest Type 2"}
    if len(df) == 0 or not required.issubset(df.columns):
        return []

    work = df.copy()
    work["_ParsedDate"] = pd.to_datetime(work["Date"], errors="coerce")
    latest_per_pond = (
        work.dropna(subset=["_ParsedDate"])
        .sort_values("_ParsedDate")
        .groupby(["Customer", "Farm Name with Code", "Pond Number"], as_index=False)
        .last()
    )
    if len(latest_per_pond) == 0:
        return []

    partial_history = (
        work.assign(_HasPartial=(
            work.get("Harvest Type", pd.Series("", index=work.index)).astype(str).str.lower().str.contains("partial")
            | work.get("Harvest Type 2", pd.Series("", index=work.index)).astype(str).str.lower().str.contains("partial")
        ))
        .groupby(["Customer", "Farm Name with Code", "Pond Number"])["_HasPartial"]
        .any()
    )

    latest_per_pond["_HasPartial"] = latest_per_pond.apply(
        lambda r: bool(partial_history.get((r["Customer"], r["Farm Name with Code"], r["Pond Number"]), False)),
        axis=1,
    )
    latest_per_pond["_Status"] = latest_per_pond.apply(lambda r: _pond_status(r, r["_HasPartial"]), axis=1)
    latest_per_pond["_DocToday"] = latest_per_pond.apply(_doc_today, axis=1)

    zones_seen = [
        _farm_zone(c, f)
        for c, f in latest_per_pond[["Customer", "Farm Name with Code"]].drop_duplicates().itertuples(index=False)
    ]
    zone_colors = build_zone_colors(zones_seen)

    # Code + location lookups, used only to attach an (optional) blurred
    # satellite background image per farm below.
    code_lookup = _customer_code_lookup()
    farm_location_lookup = build_farm_location_lookup()

    farms = []
    for (customer, farm), group in latest_per_pond.groupby(["Customer", "Farm Name with Code"]):
        total_ponds = group["Pond Number"].nunique()
        full_h_ponds = (group["_Status"] == "Full H").sum()
        if total_ponds > 0 and full_h_ponds >= total_ponds:
            continue  # every pond on this farm is Full H -- not "Running"

        zone = _farm_zone(customer, farm)
        ponds = []
        for _, prow in group.sort_values("Pond Number").iterrows():
            status = prow["_Status"]
            doc_val = prow["_DocToday"]
            if status == "Full H":
                display = "H"
            elif status == "Soon to be":
                display = "-"
            else:
                display = str(doc_val) if doc_val is not None else "-"
            ponds.append({
                "pond_no": str(prow.get("Pond Number", "")),
                "status": status,
                "display": display,
                "color": _pond_color(status),
                # Same fields the Marketing Manager app's Pond Layout
                # cards show: species letter (V/M), a WQ Special Cases
                # flag/note, and that pond's latest Issues text.
                "species": _species_letter(prow.get("Species Culture", "")),
                "wq_special": str(prow.get("WQ Special Cases", "")).strip(),
                "issues": str(prow.get("Issues", "")).strip(),
            })

        # Resolve this farm's map background image, if a matching
        # location exists. Failure here (no code, no match, bad coords)
        # just leaves map_image empty and the slide falls back to the
        # existing zone-tint-only background. When the location is a
        # polygon, also keep the polygon points + the exact bbox used for
        # the snapshot so the JS side can draw the boundary outline in
        # the right spot on top of the image.
        code = code_lookup.get((customer, farm), "")
        map_image, map_bbox, map_polygon = "", None, None
        if code:
            loc = farm_location_lookup.get(code.upper())
            if loc:
                lat, lon, polygon = loc
                map_image, bbox = build_map_image_url(lat, lon, polygon)
                map_bbox = list(bbox)
                map_polygon = [[p[0], p[1]] for p in polygon] if polygon else None

        farms.append({
            "customer": str(customer),
            "farm": str(farm),
            "zone": zone,
            "zone_color": zone_colors.get(zone, "#2563eb"),
            "ponds": ponds,
            "map_image": map_image,
            "map_bbox": map_bbox,
            "map_polygon": map_polygon,
        })

    farms.sort(key=lambda f: (f["zone"], f["customer"], f["farm"]))
    return farms

# =========================================================================
# BUILD "ZONE WISE RUNNING FARMS" DATA (for the plain table section
# below the carousel).
#
# Same "Running" rule as build_running_farms() above (a farm qualifies
# when NOT every pond it has a saved record for is at Full Harvest), but
# instead of one combined Pond Layout per farm, each farm's latest-per-
# pond records are split into a Vannamei Ponds list and a Monodon Ponds
# list (by Species Culture), so the table below can show them side by
# side. Each pond card below is ported directly from the Marketing
# Manager app's own Pond Layout cards (box color, DOC Today / Started on
# / Full H / Soon to be, Stocking Density, L.V.D, Feed/Day, ABW,
# Expecting Harvest / Harvest Weight, Total Harvest KG, Issues, species
# letter, and the WQ Special Cases icon + caption) rather than a
# simplified box, so this section looks and behaves the same as that
# reference.
# =========================================================================
def _farm_zone_zw(customer, farm):
    """Same lookup as _farm_zone() above, used ONLY by this Zone wise
    section (the carousel keeps using the original _farm_zone()
    untouched). Matches Customer Name / Farm Name with Code with
    whitespace trimmed and case-insensitively, so a farm whose Google
    Sheet spelling differs from 'Customer List.xlsx' only by case or a
    stray space still resolves to its Zone here instead of showing up
    with no zone."""
    customer_norm = str(customer).strip().lower()
    farm_norm = str(farm).strip().lower()
    match = customer_df[
        (customer_df["Customer Name"].astype(str).str.strip().str.lower() == customer_norm)
        & (customer_df["Farm Name with Code"].astype(str).str.strip().str.lower() == farm_norm)
    ]
    if len(match) > 0:
        return str(match.iloc[0].get("Zone", "")).strip()
    return ""

def _zw_parse_harvest_kg(raw_value):
    """Same combined-harvest parser as the Marketing Manager app's Pond
    Layout section: a Harvest KG value entered as a combined total across
    several ponds harvested together, e.g. '2000 (2)' (2000 kg split
    across 2 ponds), is turned into that pond's PER-POND share
    (2000 / 2 = 1000). A plain numeric value is returned as-is."""
    s = str(raw_value).strip()
    if not s:
        return float("nan")
    m = re.match(r"^([\d,]+(?:\.\d+)?)\s*\(\s*(\d+)\s*\)\s*$", s)
    if m:
        total = pd.to_numeric(m.group(1).replace(",", ""), errors="coerce")
        count = pd.to_numeric(m.group(2), errors="coerce")
        if pd.notna(total) and pd.notna(count) and count > 0:
            return total / count
        return float("nan")
    return pd.to_numeric(s.replace(",", ""), errors="coerce")

def _zw_harvest_kg_sum_row(row):
    """Same per-row harvest-KG summing rule as the Marketing Manager
    app's Pond Layout section: adds Harvest KG (slot 1) whenever that
    slot's own Harvest Type is filled in, and Harvest KG 2 (slot 2)
    whenever ITS Harvest Type 2 is filled in -- a KG value with no Type
    text is skipped."""
    row_total = 0.0
    t1 = str(row.get("Harvest Type", "")).strip()
    kg1 = _zw_parse_harvest_kg(row.get("Harvest KG", ""))
    if t1 and pd.notna(kg1):
        row_total += kg1
    t2 = str(row.get("Harvest Type 2", "")).strip()
    kg2 = _zw_parse_harvest_kg(row.get("Harvest KG 2", ""))
    if t2 and pd.notna(kg2):
        row_total += kg2
    return row_total

def build_zone_wise_running_farms(df):
    required = {"Customer", "Farm Name with Code", "Pond Number", "Date", "Harvest Type", "Harvest Type 2"}
    if len(df) == 0 or not required.issubset(df.columns):
        return []

    work = df.copy()
    work["_ParsedDate"] = pd.to_datetime(work["Date"], errors="coerce")
    latest_per_pond = (
        work.dropna(subset=["_ParsedDate"])
        .sort_values("_ParsedDate")
        .groupby(["Customer", "Farm Name with Code", "Pond Number"], as_index=False)
        .last()
    )
    if len(latest_per_pond) == 0:
        return []

    partial_history = (
        work.assign(_HasPartial=(
            work.get("Harvest Type", pd.Series("", index=work.index)).astype(str).str.lower().str.contains("partial")
            | work.get("Harvest Type 2", pd.Series("", index=work.index)).astype(str).str.lower().str.contains("partial")
        ))
        .groupby(["Customer", "Farm Name with Code", "Pond Number"])["_HasPartial"]
        .any()
    )

    latest_per_pond["_HasPartial"] = latest_per_pond.apply(
        lambda r: bool(partial_history.get((r["Customer"], r["Farm Name with Code"], r["Pond Number"]), False)),
        axis=1,
    )
    latest_per_pond["_Status"] = latest_per_pond.apply(lambda r: _pond_status(r, r["_HasPartial"]), axis=1)
    latest_per_pond["_DocToday"] = latest_per_pond.apply(_doc_today, axis=1)

    # Total Harvest KG per pond -- summed across EVERY saved record for
    # that pond (not just its latest one), same rule + parser as the
    # Marketing Manager app's Pond Layout "Total: X KG" line, so a pond
    # that had one or more Partial H harvests and then later a Full H
    # harvest gets both added together here too.
    total_harvest_kg_by_pond = (
        work.assign(_HarvestKGRow=work.apply(_zw_harvest_kg_sum_row, axis=1))
        .groupby(["Customer", "Farm Name with Code", "Pond Number"])["_HarvestKGRow"]
        .sum()
    )

    farms = []
    for (customer, farm), group in latest_per_pond.groupby(["Customer", "Farm Name with Code"]):
        total_ponds = group["Pond Number"].nunique()
        full_h_ponds = (group["_Status"] == "Full H").sum()
        if total_ponds > 0 and full_h_ponds >= total_ponds:
            continue  # every pond on this farm is Full H -- not "Running"

        zone = _farm_zone_zw(customer, farm)
        vannamei_ponds, monodon_ponds = [], []
        for _, prow in group.sort_values("Pond Number").iterrows():
            pond_no = prow.get("Pond Number", "")
            pond = {
                "pond_no": str(pond_no),
                "status": prow["_Status"],
                "doc_today": prow["_DocToday"],
                "density": prow.get("Density", ""),
                "lvd_date": prow.get("Date", ""),
                "feed_per_day": prow.get("Feed Per Day", ""),
                "abw": prow.get("ABW", ""),
                "expect_harvest_kg": prow.get("Expect Harvest (KG)", ""),
                "harvest_type": prow.get("Harvest Type", ""),
                "harvest_type2": prow.get("Harvest Type 2", ""),
                "harvest_date": prow.get("Harvest Date", ""),
                "harvest_date2": prow.get("Harvest Date 2", ""),
                "harvest_kg": prow.get("Harvest KG", ""),
                "harvest_kg2": prow.get("Harvest KG 2", ""),
                "total_harvest_kg": total_harvest_kg_by_pond.get((customer, farm, pond_no), 0),
                "issues": str(prow.get("Issues", "")).strip(),
                "wq_special": str(prow.get("WQ Special Cases", "")).strip(),
                "species": _species_letter(prow.get("Species Culture", "")),
            }
            if pond["species"] == "V":
                vannamei_ponds.append(pond)
            elif pond["species"] == "M":
                monodon_ponds.append(pond)
            # Ponds whose Species Culture is neither Vannamei nor Monodon
            # (blank/unrecognized) simply aren't shown in either column,
            # same as elsewhere in this app family.

        farms.append({
            "customer": str(customer),
            "farm": str(farm),
            "zone": zone,
            "vannamei_ponds": vannamei_ponds,
            "monodon_ponds": monodon_ponds,
        })

    farms.sort(key=lambda f: (f["zone"], f["customer"], f["farm"]))
    return farms

# =========================================================================
# BUILD HARVEST UPDATES ITEMS
#
# NOTE: the Harvest Updates line was removed from the display (replaced by
# the Active Farms count bar). These helper functions are left here
# untouched but are no longer called.
# =========================================================================
def _customer_code_lookup_for_ticker():
    return _customer_code_lookup()

def _harvest_label(h_type):
    t = str(h_type).strip().lower()
    if "full" in t:
        return "Full H"
    elif "partial" in t:
        return "Partial H"
    return str(h_type).strip() or "Harvest"

def _parse_harvest_kg(raw):
    s = str(raw).strip()
    if not s:
        return None
    m = re.match(r"^([\d,]+(?:\.\d+)?)\s*\(\s*(\d+)\s*\)\s*$", s)
    if m:
        total = pd.to_numeric(m.group(1).replace(",", ""), errors="coerce")
        count = pd.to_numeric(m.group(2), errors="coerce")
        if pd.notna(total) and pd.notna(count) and count > 0:
            return total / count
        return None
    val = pd.to_numeric(s.replace(",", ""), errors="coerce")
    return None if pd.isna(val) else val

def build_harvest_ticker(df, limit=40):
    required = {"Customer", "Farm Name with Code", "Pond Number", "Harvest Date", "Harvest Type",
                "Harvest Date 2", "Harvest Type 2"}
    if len(df) == 0 or not required.issubset(df.columns):
        return []

    code_lookup = _customer_code_lookup_for_ticker()
    events = []

    for _, row in df.iterrows():
        status1 = str(row.get("Harvest Status", "")).strip().upper()
        status2 = str(row.get("Harvest Status 2", "")).strip().upper()
        slots = []
        if (str(row.get("Harvest Date", "")).strip() or str(row.get("Harvest Type", "")).strip()) and status1 != "H":
            slots.append(("Harvest Date", "Harvest Type", "Harvest KG", "Harvest ABW"))
        if (str(row.get("Harvest Date 2", "")).strip() or str(row.get("Harvest Type 2", "")).strip()) and status2 != "H":
            slots.append(("Harvest Date 2", "Harvest Type 2", "Harvest KG 2", "Harvest ABW 2"))

        for date_col, type_col, kg_col, abw_col in slots:
            h_type = row.get(type_col, "")
            if not str(h_type).strip():
                continue
            h_date_raw = row.get(date_col, "")
            h_date = pd.to_datetime(h_date_raw, errors="coerce")
            kg_val = _parse_harvest_kg(row.get(kg_col, ""))
            abw_val = row.get(abw_col, "")

            customer = str(row.get("Customer", "")).strip()
            farm = str(row.get("Farm Name with Code", "")).strip()
            code = code_lookup.get((customer, farm), "")
            code_part = f" -{code}" if code else ""
            kg_part = f"{kg_val:,.0f}" if kg_val is not None else "-"
            abw_part = str(abw_val).strip() or "-"
            date_part = h_date.strftime("%Y/%m/%d") if pd.notna(h_date) else str(h_date_raw).strip()

            text = (
                f"{customer}{code_part} {farm} {_harvest_label(h_type)} from "
                f"Pond No {row.get('Pond Number', '')} - {kg_part} KG with {abw_part} ABW at {date_part}"
            )
            events.append((h_date if pd.notna(h_date) else pd.Timestamp.min, text))

    events.sort(key=lambda e: e[0], reverse=True)
    return [text for _, text in events[:limit]]

# =========================================================================
# LOAD + COMPUTE
# =========================================================================
if st.button("🔄 Refresh Now"):
    st.rerun()

df = load_data()
running_farms = build_running_farms(df)

st.caption(
    f"{len(running_farms)} running farm(s) • opens in Full Screen • choose a Zone, then use Back / Next "
    f"(or swipe) to move between farms • page auto-refreshes every {DATA_REFRESH_SECONDS // 60} min"
)

# =========================================================================
# RENDER -- self-contained HTML/CSS/JS component. The Zone selector, the
# Back / Next slide transitions, the Active Farms count and the Full
# Screen toggle all run entirely client-side; Python only supplies the
# data as JSON once per page load.
# =========================================================================
_HTML_TEMPLATE = """
<div id="kmn-wrap">
  <style>
    #kmn-wrap { font-family: 'Segoe UI', Tahoma, sans-serif; color:#1e293b; display:flex; flex-direction:column; }

    /* ---- top toolbar = Zone selector + Full Screen button */
    #kmn-toolbar {
      display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap;
      background:#0f172a; border-radius:12px; padding:8px 12px; margin-bottom:10px; flex-shrink:0;
    }
    .kmn-zone-label { display:flex; align-items:center; gap:8px; color:#e2e8f0; font-size:.9rem; font-weight:700; }
    #kmn-zone-select {
      background:#1e293b; color:#f8fafc; border:1px solid rgba(255,255,255,.35); border-radius:8px;
      padding:8px 10px; font-size:1rem; min-height:40px; max-width:60vw;
    }
    .kmn-fs-btn {
      background:rgba(255,255,255,.1); color:#fff; border:1px solid rgba(255,255,255,.35);
      border-radius:8px; padding:8px 14px; font-size:.85rem; cursor:pointer; min-height:40px;
    }
    .kmn-fs-btn:hover { background:rgba(255,255,255,.2); }

    #kmn-carousel {
      position: relative; width: 100%; height: 560px; overflow: hidden; touch-action: pan-y;
      border-radius: 16px; background: radial-gradient(circle at 50% 35%, #1e293b, #0f172a);
      box-shadow: 0 8px 30px rgba(0,0,0,.28); flex-shrink: 0;
    }
    .kmn-slide {
      position: absolute; inset: 0; overflow: hidden;
      transform: translateX(100%); transition: transform .6s ease-in-out;
      pointer-events: none;
    }
    .kmn-slide.active { transform: translateX(0); pointer-events: auto; z-index: 2; }
    .kmn-slide.leaving { transform: translateX(-100%); z-index: 1; pointer-events: none; }
    /* used when going Back, so the incoming slide enters from the left */
    .kmn-slide.from-left { transform: translateX(-100%); transition: none; }

    /* ---- blurred farm-location snapshot sitting behind the zone tint +
       content of each slide. Sized slightly larger than the slide
       (inset:-20px) so the blur's soft edge never shows a lighter halo
       at the slide's border. Slides with no matching location simply
       never get this element (see JS below). */
    .kmn-slide-mapbg {
      position: absolute; inset: -20px; width: calc(100% + 40px); height: calc(100% + 40px);
      object-fit: cover; filter: brightness(.85) saturate(1.15);
      z-index: 0;
    }
    /* ---- the zone-colored gradient, drawn as its own layer on top of
       the blurred map so the map shows through. */
    .kmn-slide-tint { position: absolute; inset: 0; z-index: 1; }
    /* ---- wraps the farm header + pond grid so it always sits above both
       background layers. overflow-y:auto so a farm with many ponds can be
       scrolled on a phone. */
    .kmn-slide-content {
      position: relative; z-index: 2; width: 100%; height: 100%;
      display: flex; flex-direction: column; align-items: center; justify-content: flex-start;
      padding: 24px 12px 30px; box-sizing: border-box; overflow-y: auto; scrollbar-width: thin;
      -webkit-overflow-scrolling: touch;
    }

    .kmn-slide-header { text-align: center; margin-bottom: 18px; }
    .kmn-zone-badge {
      display: inline-block; padding: 5px 16px; border-radius: 999px; font-size: .8rem;
      font-weight: 700; color: #fff; letter-spacing: .03em; margin-bottom: 8px;
    }
    .kmn-farm-name { font-size: 1.75rem; font-weight: 800; color: #f8fafc; line-height: 1.2; }
    .kmn-customer-name { font-size: 1.1rem; color: #cbd5e1; margin-top: 2px; }
    .kmn-pond-grid {
      display: flex; flex-wrap: wrap; justify-content: center; gap: 16px;
      margin-top: 16px; max-width: 1100px;
    }
    /* ---- wraps each pond box + its optional WQ Special Cases caption
       (shown below the box, same as the Marketing Manager app). */
    .kmn-pond-cell { display: flex; flex-direction: column; align-items: center; max-width: 132px; }
    .kmn-pond-box {
      position: relative;
      width: 132px; min-height: 112px; border-radius: 12px; border: 2px solid rgba(255,255,255,.18);
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      box-shadow: 0 3px 10px rgba(0,0,0,.25);
    }
    .kmn-pond-label { font-size: .72rem; color: rgba(15,23,42,.75); font-weight: 600; }
    .kmn-pond-doc { font-size: 1.45rem; font-weight: 800; color: #0f172a; margin: 2px 0; }
    .kmn-pond-status { font-size: .68rem; font-weight: 700; color: rgba(15,23,42,.8); }
    /* ---- species letter (V/M), WQ Special Cases icon + caption, and
       Issues text -- same info the Marketing Manager app's Pond Layout
       cards show, scaled down to fit this carousel's smaller box. */
    .kmn-pond-species { font-size: .68rem; font-weight: 700; color: rgba(15,23,42,.85); margin-top: 2px; }
    .kmn-pond-wq-icon { position: absolute; top: 2px; right: 4px; font-size: .95rem; line-height: 1; }
    .kmn-pond-issues {
      font-size: .62rem; font-weight: 700; color: #b91c1c; text-align: center;
      margin-top: 4px; padding: 0 4px; line-height: 1.2;
    }
    .kmn-pond-wq-caption {
      font-size: .65rem; color: #fbbf24; text-align: center; margin-top: 3px; line-height: 1.2;
    }
    /* ---- small, sharp "exact location" thumbnail shown after the Pond
       Layout grid -- distinct from the dimmed full-slide background
       image above. */
    .kmn-location-thumb-wrap { margin-top: 14px; display: flex; flex-direction: column; align-items: center; }
    .kmn-location-thumb {
      width: 220px; height: 150px; object-fit: cover; border-radius: 10px;
      border: 2px solid rgba(255,255,255,.35); box-shadow: 0 4px 14px rgba(0,0,0,.35);
    }
    .kmn-location-caption { margin-top: 5px; font-size: .75rem; color: #e2e8f0; font-weight: 600; }
    .kmn-empty { color: #94a3b8; font-size: 1.2rem; margin-top: 60px; text-align: center; }
    .kmn-dots { position: absolute; bottom: 8px; left: 0; right: 0; display: flex; justify-content: center; gap: 7px; z-index: 3; pointer-events: none; }
    .kmn-dot { width: 8px; height: 8px; border-radius: 50%; background: rgba(255,255,255,.28); transition: background .3s; }
    .kmn-dot.active { background: #fff; }

    /* ---- bottom control bar (replaces the old Harvest Updates line):
       Back button | Active Farms count | Next button */
    #kmn-controls {
      margin-top: 10px; background: #0f172a; border-radius: 12px; padding: 8px 12px;
      display: flex; align-items: center; justify-content: space-between; gap: 10px; flex-shrink: 0;
    }
    .kmn-nav-btn {
      background: #2563eb; color: #fff; border: none; border-radius: 10px; padding: 12px 18px;
      font-size: 1rem; font-weight: 700; cursor: pointer; min-width: 92px; touch-action: manipulation;
    }
    .kmn-nav-btn:disabled { background: #475569; opacity: .6; cursor: default; }
    #kmn-active-info { flex: 1; text-align: center; min-width: 0; }
    #kmn-active-count { color: #fbbf24; font-weight: 800; font-size: .95rem; }
    #kmn-slide-counter { color: #94a3b8; font-size: .75rem; font-weight: 600; margin-top: 2px; }

    /* ---- Full Screen mode (now the default on open): expands the whole
       component to fill the browser window. */
    #kmn-wrap.kmn-fullscreen-mode {
       position: fixed; inset: 0; z-index: 999999;
       background: #0b1220; padding: 10px; box-sizing: border-box; align-items: center;
    }
    #kmn-wrap.kmn-fullscreen-mode #kmn-toolbar,
    #kmn-wrap.kmn-fullscreen-mode #kmn-controls { width: 100%; max-width: 1600px; box-sizing: border-box; }
    #kmn-wrap.kmn-fullscreen-mode #kmn-carousel {
       width: 100%; max-width: 1600px; height: auto; flex: 1 1 auto; min-height: 0;
    }
    /* Centers each slide's content vertically in Full Screen, but stays
       scrollable when it is taller than the screen. */
    #kmn-wrap.kmn-fullscreen-mode .kmn-slide-content > :first-child { margin-top: auto; }
    #kmn-wrap.kmn-fullscreen-mode .kmn-slide-content > :last-child { margin-bottom: auto; }

    /* ---- phone-friendly sizing */
    @media (max-width: 600px) {
      .kmn-farm-name { font-size: 1.3rem; }
      .kmn-customer-name { font-size: .95rem; }
      .kmn-pond-grid { gap: 10px; }
      .kmn-location-thumb { width: 180px; height: 120px; }
      #kmn-carousel { height: 520px; }
      #kmn-zone-select { max-width: 46vw; }
      .kmn-nav-btn { padding: 12px 12px; min-width: 78px; }
    }
  </style>

  <div id="kmn-toolbar">
    <label class="kmn-zone-label">📍 Zone
      <select id="kmn-zone-select"></select>
    </label>
    <button id="kmn-fullscreen-btn" class="kmn-fs-btn" title="Toggle full screen">⛶ Full Screen</button>
  </div>

  <div id="kmn-carousel">
    <div id="kmn-slides"></div>
    <div class="kmn-dots" id="kmn-dots"></div>
  </div>

  <div id="kmn-controls">
    <button id="kmn-prev" class="kmn-nav-btn">◀ Back</button>
    <div id="kmn-active-info">
      <div id="kmn-active-count"></div>
      <div id="kmn-slide-counter"></div>
    </div>
    <button id="kmn-next" class="kmn-nav-btn">Next ▶</button>
  </div>

  <script>
    (function () {
      const farms = __FARMS_JSON__;
      const dataRefreshSeconds = __DATA_REFRESH_SECONDS__;
      const ALL = '__all__', NONE = '__none__';

      const slidesEl = document.getElementById('kmn-slides');
      const dotsEl = document.getElementById('kmn-dots');
      const carouselEl = document.getElementById('kmn-carousel');
      const zoneSelect = document.getElementById('kmn-zone-select');
      const countEl = document.getElementById('kmn-active-count');
      const counterEl = document.getElementById('kmn-slide-counter');
      const prevBtn = document.getElementById('kmn-prev');
      const nextBtn = document.getElementById('kmn-next');
      const fsButton = document.getElementById('kmn-fullscreen-btn');
      const wrapEl = document.getElementById('kmn-wrap');

      let visible = farms.slice();
      let selectedZone = ALL;
      let current = 0;
      let isFullscreen = false;
      let fsFrameEl = null;
      try { fsFrameEl = window.frameElement; } catch (e) { fsFrameEl = null; }

      function escapeHtml(v) {
        return String(v == null ? '' : v)
          .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      }
      // Remembers the chosen Zone / farm across the periodic page refresh.
      function store(k, v) { try { sessionStorage.setItem(k, v); } catch (e) {} }
      function recall(k) { try { return sessionStorage.getItem(k); } catch (e) { return null; } }

      function zoneKey(f) { return f.zone ? f.zone : NONE; }
      function zoneLabel(k) { return k === NONE ? 'No Zone' : 'Zone ' + k; }

      // ---- Each zone gets its own tinted slide background (in addition
      // to the zone badge), so different zones are visually distinct.
      function zoneSlideBackground(zoneColor) {
        return 'linear-gradient(165deg, ' + zoneColor + '55 0%, #0f172a 62%)';
      }

      function buildZoneOptions() {
        const keys = [];
        farms.forEach(function (f) { const k = zoneKey(f); if (keys.indexOf(k) === -1) keys.push(k); });
        zoneSelect.innerHTML = '<option value="' + ALL + '">All Zones</option>'
          + keys.map(function (k) {
            return '<option value="' + escapeHtml(k) + '">' + escapeHtml(zoneLabel(k)) + '</option>';
          }).join('');
      }

      function renderSlides() {
        if (!visible.length) {
          slidesEl.innerHTML = '<div class="kmn-slide active"><div class="kmn-empty">No running farms found.</div></div>';
          dotsEl.innerHTML = '';
          return;
        }
        slidesEl.innerHTML = visible.map(function (f, i) {
          const ponds = f.ponds.map(function (p) {
            const wqIconHtml = p.wq_special
              ? '<div class="kmn-pond-wq-icon" title="WQ Special Case">🫨</div>' : '';
            const speciesHtml = p.species
              ? '<div class="kmn-pond-species">' + escapeHtml(p.species) + '</div>' : '';
            const issuesHtml = p.issues
              ? '<div class="kmn-pond-issues">' + escapeHtml(p.issues) + '</div>' : '';
            const wqCaptionHtml = p.wq_special
              ? '<div class="kmn-pond-wq-caption">🫨 ' + escapeHtml(p.wq_special) + '</div>' : '';
            return '<div class="kmn-pond-cell">'
              + '<div class="kmn-pond-box" style="background:' + p.color + '; border-color:' + f.zone_color + ';">'
              + wqIconHtml
              + '<div class="kmn-pond-label">Pond ' + escapeHtml(p.pond_no) + '</div>'
              + '<div class="kmn-pond-doc">' + escapeHtml(p.display) + '</div>'
              + '<div class="kmn-pond-status">' + escapeHtml(p.status) + '</div>'
              + speciesHtml
              + issuesHtml
              + '</div>'
              + wqCaptionHtml
              + '</div>';
          }).join('');

          // Only add the blurred map <img> when this farm actually
          // resolved a location -- farms with no match keep exactly the
          // zone-tint-only background. onerror hides it gracefully if
          // the snapshot URL ever fails to load.
          const mapBgHtml = f.map_image
            ? '<img class="kmn-slide-mapbg" src="' + f.map_image + '" alt="" onerror="this.remove();" />'
            : '';

          // A small, sharp "exact location" thumbnail shown after the
          // Pond Layout grid -- separate from the dimmed full-slide
          // background above, and only added when a location was found.
          const locationThumbHtml = f.map_image
            ? '<div class="kmn-location-thumb-wrap">'
              + '<img class="kmn-location-thumb" src="' + f.map_image + '" alt="Farm location" onerror="this.parentElement.remove();" />'
              + '<div class="kmn-location-caption">📍 Farm Location</div>'
              + '</div>'
            : '';

          return '<div class="kmn-slide' + (i === current ? ' active' : '') + '" data-index="' + i + '">'
            + mapBgHtml
            + '<div class="kmn-slide-tint" style="background:' + zoneSlideBackground(f.zone_color) + ';"></div>'
            + '<div class="kmn-slide-content">'
            + '<div class="kmn-slide-header">'
            + '<div class="kmn-zone-badge" style="background:' + f.zone_color + ';">Zone ' + escapeHtml(f.zone || '-') + '</div>'
            + '<div class="kmn-farm-name">' + escapeHtml(f.farm) + '</div>'
            + '<div class="kmn-customer-name">' + escapeHtml(f.customer) + '</div>'
            + '</div>'
            + '<div class="kmn-pond-grid">' + ponds + '</div>'
            + locationThumbHtml
            + '</div>'
            + '</div>';
        }).join('');

        // Dots only when there are few enough to fit; the "Farm x of y"
        // counter below always shows.
        dotsEl.innerHTML = visible.length <= 15
          ? visible.map(function (_, i) {
              return '<div class="kmn-dot' + (i === current ? ' active' : '') + '"></div>';
            }).join('')
          : '';
      }

      // ---- Active Farms count (replaces the old Harvest Updates line).
      function updateInfo() {
        const n = visible.length;
        countEl.textContent = selectedZone === ALL
          ? 'All Running Farms: ' + n
          : 'Active Farms in ' + zoneLabel(selectedZone) + ': ' + n;
        counterEl.textContent = n ? 'Farm ' + (current + 1) + ' of ' + n : '';
        prevBtn.disabled = nextBtn.disabled = n < 2;
      }

      function applyZone(zone, startIndex) {
        selectedZone = zone;
        zoneSelect.value = zone;
        visible = zone === ALL ? farms.slice() : farms.filter(function (f) { return zoneKey(f) === zone; });
        current = Math.min(Math.max(startIndex || 0, 0), Math.max(visible.length - 1, 0));
        renderSlides();
        updateInfo();
        store('kmn_zone', zone);
        store('kmn_idx', String(current));
      }

      // ---- Slide transition. dir = +1 (Next: the incoming slide moves
      // in from the right, the outgoing one leaves to the left) or -1
      // (Back: the incoming slide enters from the left). No fade/opacity.
      function goTo(newIndex, dir) {
        const slides = slidesEl.querySelectorAll('.kmn-slide');
        const dots = dotsEl.querySelectorAll('.kmn-dot');
        slides.forEach(function (s, i) {
          if (i === newIndex) {
            if (dir < 0) {
              s.classList.remove('active', 'leaving');
              s.classList.add('from-left');
              void s.offsetWidth; // force reflow so the transition below actually runs
              s.classList.remove('from-left');
              s.classList.add('active');
            } else {
              s.classList.remove('leaving', 'from-left');
              s.classList.add('active');
            }
          } else if (i === current) {
            s.classList.remove('active');
            if (dir > 0) {
              s.classList.add('leaving');
              setTimeout(function () { s.classList.remove('leaving'); }, 650);
            }
          } else {
            s.classList.remove('active', 'leaving', 'from-left');
          }
        });
        dots.forEach(function (d, i) { d.classList.toggle('active', i === newIndex); });
        current = newIndex;
        updateInfo();
        store('kmn_idx', String(current));
      }

      function step(dir) {
        const n = visible.length;
        if (n < 2) return;
        goTo((current + dir + n) % n, dir);
      }

      zoneSelect.addEventListener('change', function () { applyZone(this.value, 0); });
      prevBtn.addEventListener('click', function () { step(-1); });
      nextBtn.addEventListener('click', function () { step(1); });

      // ---- Swipe left/right on touch screens = Next / Back.
      let touchX = null, touchY = null;
      carouselEl.addEventListener('touchstart', function (e) {
        touchX = e.touches[0].clientX; touchY = e.touches[0].clientY;
      }, { passive: true });
      carouselEl.addEventListener('touchend', function (e) {
        if (touchX === null) return;
        const dx = e.changedTouches[0].clientX - touchX;
        const dy = e.changedTouches[0].clientY - touchY;
        touchX = touchY = null;
        if (Math.abs(dx) > 50 && Math.abs(dx) > Math.abs(dy) * 1.5) step(dx < 0 ? 1 : -1);
      }, { passive: true });

      // ---- Full Screen: expands the whole component to fill the
      // browser window. Resizes the actual Streamlit component iframe
      // (same-origin, via window.frameElement) so it behaves like a true
      // full-screen kiosk view; also makes a best-effort attempt at the
      // browser's native Fullscreen API (browsers only allow that after
      // a tap/click, so on first open the CSS-based full screen applies).
      function enterFullscreen() {
        isFullscreen = true;
        wrapEl.classList.add('kmn-fullscreen-mode');
        if (fsFrameEl) {
          fsFrameEl.style.position = 'fixed';
          fsFrameEl.style.top = '0';
          fsFrameEl.style.left = '0';
          fsFrameEl.style.width = '100vw';
          fsFrameEl.style.height = '100vh';
          fsFrameEl.style.height = '100dvh'; // ignored by browsers that don't know dvh (keeps 100vh)
          fsFrameEl.style.zIndex = '999999';
          fsFrameEl.style.border = 'none';
        }
        fsButton.textContent = '⛶ Exit Full Screen';
        try {
          if (document.documentElement.requestFullscreen) {
            document.documentElement.requestFullscreen().catch(function () {});
          }
        } catch (e) { /* ignored -- CSS-based fallback above still applies */ }
      }

      function exitFullscreen() {
        isFullscreen = false;
        wrapEl.classList.remove('kmn-fullscreen-mode');
        if (fsFrameEl) {
          fsFrameEl.style.position = '';
          fsFrameEl.style.top = '';
          fsFrameEl.style.left = '';
          fsFrameEl.style.width = '';
          fsFrameEl.style.height = '';
          fsFrameEl.style.zIndex = '';
          fsFrameEl.style.border = '';
        }
        fsButton.textContent = '⛶ Full Screen';
        try {
          if (document.fullscreenElement && document.exitFullscreen) {
            document.exitFullscreen().catch(function () {});
          }
        } catch (e) { /* ignored */ }
      }

      fsButton.addEventListener('click', function () {
        if (isFullscreen) { exitFullscreen(); } else { enterFullscreen(); }
      });
      document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape' && isFullscreen) exitFullscreen();
        else if (e.key === 'ArrowRight') step(1);
        else if (e.key === 'ArrowLeft') step(-1);
      });

      // ---- Start-up: restore the last Zone / farm (if any), then open
      // in Full Screen by default.
      buildZoneOptions();
      const savedZone = recall('kmn_zone');
      const zoneOk = savedZone && Array.prototype.some.call(zoneSelect.options, function (o) { return o.value === savedZone; });
      applyZone(zoneOk ? savedZone : ALL, zoneOk ? parseInt(recall('kmn_idx') || '0', 10) : 0);
      enterFullscreen();

      // Periodically reload the whole app so it pulls fresh data from the
      // Google Sheet (the chosen Zone / farm is restored afterwards); the
      // manual "Refresh Now" button above still works at any time.
      setInterval(function () {
        try { window.parent.location.reload(); } catch (e) { window.location.reload(); }
      }, dataRefreshSeconds * 1000);
    })();
  </script>
</div>
"""

_html = (
    _HTML_TEMPLATE
    .replace("__FARMS_JSON__", json.dumps(running_farms))
    .replace("__DATA_REFRESH_SECONDS__", json.dumps(DATA_REFRESH_SECONDS))
)

components.html(_html, height=720, scrolling=False)

# =========================================================================
# "Zone wise Running Farms - Live Display" section.
#
# A plain (non-rotating) table below the carousel, grouped by Zone.
# Columns: Customer Name + Farm Name with Code | Vannamei Ponds | Monodon
# Ponds. Each pond is rendered as a small box using the same status
# colors as the carousel above, showing that pond's DOC Today value, and
# -- ported from the Marketing Manager app's Pond Layout cards -- a WQ
# Special Cases icon/caption and an Issues line when either is present on
# that pond's latest saved record.
# =========================================================================
def _escape_html_zw(v):
    return str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def _zw_pond_box_color(status):
    """Same box-color mapping as the Marketing Manager app's Pond Layout
    section's _pond_box_color(): yellow = Partial H, green = Full H,
    gray = Soon to be, blue = Running (default)."""
    return {
        "Partial H": "#fff3cd",
        "Full H": "#d4edda",
        "Soon to be": "#e2e2e2",
    }.get(status, "#eaf4ff")

def _render_zone_wise_pond_box(p):
    """Renders one pond card using the EXACT same layout, fields and
    rules as the Marketing Manager app's Pond Layout section (box color,
    DOC Today + 'Started on' date / Full H + Harvest Date / Soon to be,
    Total Harvest KG, Expecting Harvest / Harvest Weight, Stocking
    Density, L.V.D, Feed/Day, ABW, Issues, species letter, and the WQ
    Special Cases icon + caption) -- ported field-for-field from that
    reference rather than a simplified box."""
    status = p["status"]
    box_color = _zw_pond_box_color(status)

    wq_special_val = p["wq_special"]
    wq_icon_html = (
        "<div style='position:absolute;top:2px;right:4px;font-size:1.2rem;line-height:1;' "
        "title='WQ Special Case'>🫨</div>"
        if wq_special_val else ""
    )
    wq_caption_html = (
        f"<div style='font-size:0.8rem;color:#b45309;text-align:center;max-width:190px;"
        f"margin-top:2px;'>🫨 {_escape_html_zw(wq_special_val)}</div>"
        if wq_special_val else ""
    )

    # ---- Stocking Density, L.V.D (this pond's own saved Date), Feed/Day,
    # ABW -- same detail lines as the reference's pond cards.
    density_val = pd.to_numeric(p.get("density", ""), errors="coerce")
    density_str = f"{density_val:,.0f}" if pd.notna(density_val) else "-"
    lvd_str = _escape_html_zw(str(p.get("lvd_date", "")).strip() or "-")
    feed_day_str = _escape_html_zw(p.get("feed_per_day", "") or "-")
    abw_str = _escape_html_zw(p.get("abw", "") or "-")
    extra_details_html = (
        "<div style='font-size:0.85rem;color:#333;text-align:left;width:100%;"
        "padding:0 8px;margin-top:4px;line-height:1.4;'>"
        f"<div>Stocking Density - {density_str}</div>"
        f"<div>L.V.D - {lvd_str}</div>"
        f"<div>Feed/Day - {feed_day_str} &nbsp;|&nbsp; ABW - {abw_str}</div>"
        "</div>"
    )

    # ---- Issues (this pond's latest saved record), shown at the bottom
    # of the card in red -- same as the reference.
    issues_val = str(p.get("issues", "")).strip()
    issues_html = (
        "<div style='margin-top:auto;width:100%;text-align:center;font-size:0.85rem;"
        "font-weight:bold;border-top:1px dashed #bbb;padding-top:3px;'>"
        f"<span style='color:red;'>{_escape_html_zw(issues_val)}</span></div>"
        if issues_val and issues_val.lower() != "nan" else ""
    )

    total_kg = p.get("total_harvest_kg", 0) or 0

    if status == "Full H":
        # Full H: "Full H" + its Harvest Date, plus Total Harvest KG
        # (all harvests summed for this pond) instead of DOC Today.
        h_date = str(p.get("harvest_date2", "")).strip() or str(p.get("harvest_date", "")).strip()
        h_date = _escape_html_zw(h_date or "-")
        total_kg_html = (
            f"<div style='font-size:0.75rem;color:#333;'>Total: {total_kg:,.2f} KG</div>" if total_kg else ""
        )
        box_middle_html = (
            "<div style='font-size:1.2rem;font-weight:bold;color:red;'>Full H</div>"
            f"<div style='font-size:0.75rem;color:#333;'>Harvest Date - {h_date}</div>"
            f"{total_kg_html}"
        )
    elif status == "Soon to be":
        box_middle_html = "<div style='font-size:1.1rem;font-weight:bold;color:#555;'>Soon to be</div>"
    else:
        # Running / Partial H: DOC Today (red, large) + "Started on
        # <date>" below it; Partial H additionally shows Total Harvest
        # KG (Running has no harvest yet, so this stays blank for it).
        if status == "Partial H":
            total_kg_html = (
                f"<div style='font-size:0.7rem;color:#333;'>Total: {total_kg:,.2f} KG</div>" if total_kg else ""
            )
        else:
            total_kg_html = ""

        doc_today_raw = p.get("doc_today", "")
        doc_today_val = _escape_html_zw(doc_today_raw if doc_today_raw is not None else "-")
        try:
            started_date = (
                pd.Timestamp(date.today()) - pd.Timedelta(days=int(float(doc_today_raw)))
            ).strftime("%Y-%m-%d")
            started_label = f"Started on {started_date}"
        except (TypeError, ValueError):
            started_label = "Started on ---"
        box_middle_html = (
            f"<div style='font-size:1.4rem;font-weight:bold;color:red;'>{doc_today_val}</div>"
            f"<div style='font-size:0.7rem;color:#777;'>{_escape_html_zw(started_label)}</div>"
            f"{total_kg_html}"
        )

    # ---- Expecting Harvest (KG) / Harvest Weight line -- same label
    # switch and "2nd slot wins" rule as the reference.
    if status == "Full H":
        t2_lower = str(p.get("harvest_type2", "")).strip().lower()
        kg2 = _zw_parse_harvest_kg(p.get("harvest_kg2", ""))
        kg1 = _zw_parse_harvest_kg(p.get("harvest_kg", ""))
        harvest_kg_val = kg2 if ("full" in t2_lower and pd.notna(kg2)) else kg1
        expect_label = "Harvest Weight"
        expect_val = f"{harvest_kg_val:,.2f} KG" if pd.notna(harvest_kg_val) else "-"
    elif status == "Soon to be":
        expect_label = "Expecting Harvest"
        expect_val = "-"
    else:
        expect_label = "Expecting Harvest"
        expect_kg = pd.to_numeric(p.get("expect_harvest_kg", ""), errors="coerce")
        expect_val = f"{expect_kg:,.2f} KG" if pd.notna(expect_kg) else "-"

    expect_html = (
        "<div style='font-size:0.85rem;color:#333;text-align:center;width:100%;margin-top:4px;"
        "border-top:1px dashed #bbb;padding-top:3px;'>"
        f"<b>{expect_label}:</b> {_escape_html_zw(expect_val)}</div>"
    )

    species_label = p.get("species", "")
    species_html = (
        f"<div style='font-size:0.75rem;font-weight:bold;color:#444;margin-top:2px;'>{_escape_html_zw(species_label)}</div>"
        if species_label else ""
    )

    # ---- Card shell: same 210px x 175px card, border and padding as the
    # reference's Pond Layout cards.
    return (
        "<div style='display:inline-flex;flex-direction:column;align-items:center;margin:6px;vertical-align:top;'>"
        "<div style='position:relative;width:210px;min-height:175px;border:2px solid #333;"
        "border-radius:6px;display:flex;flex-direction:column;align-items:center;"
        f"justify-content:flex-start;padding:8px 0;background:{box_color};'>"
        f"{wq_icon_html}"
        f"<div style='font-size:0.8rem;color:#555;'>Pond {_escape_html_zw(p['pond_no'])}</div>"
        f"{box_middle_html}"
        f"{expect_html}"
        f"{extra_details_html}"
        f"{issues_html}"
        "</div>"
        f"{species_html}"
        f"{wq_caption_html}"
        "</div>"
    )

def render_zone_wise_section(zone_wise_farms):
    if not zone_wise_farms:
        st.info("No running farms found.")
        return

    rows_html = ""
    current_zone = None
    for f in zone_wise_farms:
        if f["zone"] != current_zone:
            current_zone = f["zone"]
            # A farm whose Customer Name / Farm Name with Code couldn't be
            # matched against 'Customer List.xlsx' (see _farm_zone_zw()
            # above) has no Zone -- labelled "Unassigned" here instead of
            # a blank header, so it's still grouped and visible rather
            # than silently disappearing from the table.
            zone_label = current_zone if current_zone else "Unassigned"
            rows_html += (
                "<tr><td colspan='3' style='background:#1e293b;color:#f8fafc;font-weight:800;"
                f"font-size:.85rem;padding:8px 14px;'>Zone: {_escape_html_zw(zone_label)}</td></tr>"
            )
        vannamei_html = (
            "".join(_render_zone_wise_pond_box(p) for p in f["vannamei_ponds"])
            or "<span style='color:#94a3b8;font-size:.85rem;'>—</span>"
        )
        monodon_html = (
            "".join(_render_zone_wise_pond_box(p) for p in f["monodon_ponds"])
            or "<span style='color:#94a3b8;font-size:.85rem;'>—</span>"
        )
        rows_html += (
            "<tr>"
            "<td style='padding:12px 14px;border-bottom:1px solid #e2e8f0;vertical-align:top;white-space:nowrap;'>"
            f"<div style='font-weight:700;color:#0f172a;font-size:.95rem;'>{_escape_html_zw(f['farm'])}</div>"
            f"<div style='font-size:.82rem;color:#475569;'>{_escape_html_zw(f['customer'])}</div>"
            "</td>"
            f"<td style='padding:12px 14px;border-bottom:1px solid #e2e8f0;'>{vannamei_html}</td>"
            f"<td style='padding:12px 14px;border-bottom:1px solid #e2e8f0;'>{monodon_html}</td>"
            "</tr>"
        )

    table_html = (
        "<div style='overflow-x:auto;width:100%;'>"
        "<table style='width:100%;border-collapse:collapse;font-family:\"Segoe UI\", Tahoma, sans-serif;'>"
        "<thead><tr style='background:#0f172a;color:#f8fafc;'>"
        "<th style='padding:10px 14px;text-align:left;font-size:.85rem;'>Customer Name / Farm Name with Code</th>"
        "<th style='padding:10px 14px;text-align:left;font-size:.85rem;'>Vannamei Ponds</th>"
        "<th style='padding:10px 14px;text-align:left;font-size:.85rem;'>Monodon Ponds</th>"
        "</tr></thead>"
        f"<tbody>{rows_html}</tbody>"
        "</table></div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)

st.markdown("---")
st.subheader("🗺️ Zone wise Running Farms — Live Display")
zone_wise_farms = build_zone_wise_running_farms(df)
render_zone_wise_section(zone_wise_farms)
st.caption(
    f"{len(zone_wise_farms)} running farm(s) shown, grouped by Zone (farms with no matching Zone in "
    "'Customer List.xlsx' are grouped under 'Zone: Unassigned') • V/M columns list only that farm's "
    "Vannamei / Monodon ponds • box color = pond status (blue = Running, yellow = Partial H, "
    "green = Full H, gray = Soon to be) • each card shows DOC Today / Started on date, Expecting "
    "Harvest or Harvest Weight, Stocking Density, L.V.D, Feed/Day, ABW, Issues and WQ Special Cases — "
    "same fields as the Marketing Manager app's Pond Layout cards"
)

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>KMN Aqua Services - Running Farms Live Display "
    "(read only, auto-updating)</p>",
    unsafe_allow_html=True,
)
