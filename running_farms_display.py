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
#      The display ALWAYS opens (and stays) in Full Screen -- there is no
#      toggle / exit button any more, since this is a kiosk-only view.
#      There's a Zone selector (All Zones / one Zone) and the slides move
#      ONLY with the Back / Next buttons (or swipe / arrow keys) -- no
#      automatic rotation. It is also phone-friendly. Each slide scrolls
#      internally when its content is taller than the screen.
#   2) A control bar shows the Active Farms count for the selected Zone
#      ("Active Farms in Zone X: N"), or "All Running Farms: N" when no
#      Zone is selected, between the Back and Next buttons.
#      After choosing a Zone, a second dropdown lists each running farm as
#      "Customer name - Farm name" and jumps straight to it.
#
# CHANGED (this revision):
#   - Removed the "Zone wise Running Farms - Live Display" table section
#     that used to sit below the carousel. Only the carousel ("Running
#     Farms — Live Display") remains.
#   - Full Screen is no longer a toggle: there is no "⛶ Full Screen" /
#     "Exit Full Screen" button any more. The display applies Full Screen
#     immediately on open and stays there -- this is the only mode.
#   - The blurred satellite background image / small "Farm Location"
#     thumbnail stayed / is back (Locations Google Sheet + ArcGIS
#     snapshot lookup, same as before).
#   - Added an "L.V.D" line on each slide -- the LATEST date among that
#     farm's own ponds' individual L.V.D dates (each pond's own most
#     recently saved "Date"), same field the Marketing Manager app's Pond
#     Layout cards call "L.V.D", just rolled up to one date per farm here.
#   - Added a small table per slide (Pond No / Feed Per Day / ABW /
#     Expecting Harvest), one row per pond, using each pond's latest
#     saved record -- same fields and "2nd harvest slot wins" Expecting
#     Harvest rule as the Marketing Manager app's Pond Layout cards. Ponds
#     whose latest record is Full Harvest are left OUT of this table.
#   - Added a "Last Feed Purchased Date" + "Last Feed Order" block at the
#     bottom of each slide (each feed item on its own line), pulled from
#     the same Sales Details Google Sheet + "FEED" item-prefix rule the
#     Marketing Manager app uses, filtered to that farm's Customer Code.
#   - Each slide's content area scrolls (the extra rows/table/feed block
#     can make a slide taller than the screen) -- swipe/scroll down on a
#     slide to see everything.
#
# All of the carousel/zone/full-screen behaviour runs client-side in a
# single self-contained HTML/CSS/JS component
# (streamlit.components.v1.html) -- Python only computes the data once
# per page load/refresh.
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

# ---- Second Google Sheet -- Sales Details. Same spreadsheet the
# Marketing Manager app reads, used here ONLY to build each farm's
# "Last Feed Purchased Date" / "Last Feed Order" line at the bottom of
# its slide (same "FEED" item-prefix rule, filtered by that farm's
# Customer Code). The same service account must also be shared (as
# Viewer or Editor) on this sheet.
SALES_SHEET_ID = "1S3csAE-E_hN8vstuHR0KkeAN7yCVQTFe4AkEVlw4vQw"
SALES_COLUMN_ORDER = ["Date", "Item No.", "Item Description", "Customer Code", "Quantity"]

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
    # Merge spelling variants of the same Customer / Farm name (extra or
    # non-breaking spaces, different capitalisation) into ONE spelling (the
    # most frequent one), so a single farm is not split into several
    # "duplicate" farms with different pond layouts.
    if len(df) > 0:
        for _c in ("Customer", "Farm Name with Code"):
            cleaned = df[_c].map(lambda v: re.sub(r"\s+", " ", str(v).replace("\u00a0", " ")).strip())
            keys = cleaned.str.lower()
            canon = cleaned.groupby(keys).agg(lambda x: x.value_counts().idxmax())
            df[_c] = keys.map(canon)
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
# LOAD SALES DETAILS (for each farm's "Last Feed Purchased Date" / "Last
# Feed Order" line only)
#
# Ported from the Marketing Manager app's own loader -- same spreadsheet,
# read-only, best-effort: if the sheet can't be reached, every farm just
# falls back to showing "-" for these two lines, nothing else in this app
# is affected.
# =========================================================================
@st.cache_resource(show_spinner=False)
def get_sales_worksheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sh = client.open_by_key(SALES_SHEET_ID)
    worksheet_name = st.secrets.get("gsheet", {}).get("sales_worksheet_name", "")
    if worksheet_name:
        return sh.worksheet(worksheet_name)
    return sh.sheet1

def load_sales_data():
    """Always reads fresh from the Sales Details Google Sheet (no
    caching), same pattern as load_data() above."""
    ws = get_sales_worksheet()
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for c in SALES_COLUMN_ORDER:
        if c not in df.columns:
            df[c] = ""
    if len(df) > 0:
        df = df[SALES_COLUMN_ORDER]
    return df

def build_last_feed_order_lookup(sales_df):
    """Returns {customer_code (lower): (last_feed_date, [\"Item - Qty\", ...])}.

    Same rule as the Marketing Manager app's "Last Feed Order" line: only
    rows whose \"Item No.\" starts with \"FEED\" count, the most recent
    Date among those wins, and every item purchased on that exact date is
    listed (one entry per Item Description, quantities summed). Uses ALL
    sales rows regardless of any Settle flag, same as that app.
    """
    lookup = {}
    if sales_df.empty or "Customer Code" not in sales_df.columns:
        return lookup

    work = sales_df.copy()
    work["Quantity"] = pd.to_numeric(work["Quantity"], errors="coerce").fillna(0)
    work["_CodeKey"] = work["Customer Code"].astype(str).str.strip().str.lower()
    is_feed = work["Item No."].astype(str).str.strip().str.upper().str.startswith("FEED")
    work = work[is_feed & (work["_CodeKey"] != "")]
    if work.empty:
        return lookup

    for code_key, group in work.groupby("_CodeKey"):
        last_date = group["Date"].max()
        items = (
            group[group["Date"] == last_date]
            .groupby("Item Description")["Quantity"]
            .sum()
        )
        item_lines = [f"{item} - {qty:,.0f}" for item, qty in items.items()]
        lookup[code_key] = (str(last_date), item_lines)
    return lookup

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
def _norm_key(v):
    """Lower-cases and collapses extra / non-breaking spaces so names that
    only differ by case or hidden spacing still match."""
    return re.sub(r"\s+", " ", str(v).replace("\u00a0", " ")).strip().lower()

def _farm_zone(customer, farm):
    """Finds a farm's Zone in 'Customer List.xlsx'. Tries the names first
    (ignoring case / odd spacing / rows with a blank Zone), then falls back
    to the Customer ID (e.g. C00876) if it appears in the sheet's names."""
    cust_n, farm_n = _norm_key(customer), _norm_key(farm)
    cust_col = customer_df["Customer Name"].map(_norm_key)
    farm_col = customer_df["Farm Name with Code"].map(_norm_key)
    zones = customer_df["Zone"].astype(str).str.strip()
    has_zone = zones.ne("") & zones.str.lower().ne("nan")

    # 1) match by names
    for mask in ((cust_col == cust_n) & (farm_col == farm_n), (farm_col == farm_n)):
        m = customer_df[mask & has_zone]
        if len(m) > 0:
            return str(m.iloc[0]["Zone"]).strip()

    # 2) fallback: match by Customer ID / code found in the names
    for cand in _CUSTOMER_CODE_COLUMN_CANDIDATES:
        if cand in customer_df.columns:
            ids = customer_df[cand].astype(str).str.strip().str.upper()
            for code in re.findall(r"[A-Za-z]\d{3,}", f"{customer} {farm}"):
                m = customer_df[(ids == code.upper()) & has_zone]
                if len(m) > 0:
                    return str(m.iloc[0]["Zone"]).strip()
            break
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
def _parse_harvest_kg_value(raw_value):
    """Same combined-harvest parser used across this app family: a
    Harvest KG value entered as a combined total across several ponds
    harvested together, e.g. '2000 (2)' (2000 kg split across 2 ponds),
    is turned into that pond's PER-POND share (2000 / 2 = 1000). A plain
    numeric value is returned as-is."""
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

def _expecting_harvest_display(prow, status):
    """Same label/value switch used by the Marketing Manager app's Pond
    Layout cards: Full H ponds show their actual Harvest Weight (2nd
    harvest slot wins when it says 'Full'), Soon to be ponds show '-',
    everything else shows the pond's own Expect Harvest (KG) estimate."""
    if status == "Full H":
        t2 = str(prow.get("Harvest Type 2", "")).strip().lower()
        kg2 = _parse_harvest_kg_value(prow.get("Harvest KG 2", ""))
        kg1 = _parse_harvest_kg_value(prow.get("Harvest KG", ""))
        val = kg2 if ("full" in t2 and pd.notna(kg2)) else kg1
        return f"{val:,.2f} KG" if pd.notna(val) else "-"
    elif status == "Soon to be":
        return "-"
    else:
        val = pd.to_numeric(prow.get("Expect Harvest (KG)", ""), errors="coerce")
        return f"{val:,.2f} KG" if pd.notna(val) else "-"

def build_running_farms(df, last_feed_lookup=None):
    last_feed_lookup = last_feed_lookup or {}
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

    # Customer Code lookup, used both for the map background image below
    # and to look up each farm's "Last Feed Purchased Date" / "Last Feed
    # Order" in last_feed_lookup below.
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
        feed_table = []
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
            # ---- Small table row: Pond No / Feed Per Day / ABW /
            # Expecting Harvest -- one row per pond, same fields the
            # Marketing Manager app's Pond Layout cards show. Ponds
            # already at Full Harvest are left OUT of this table (they
            # have no more feed/ABW/expecting-harvest to track).
            if status != "Full H":
                feed_table.append({
                    "pond_no": str(prow.get("Pond Number", "")),
                    "feed_per_day": str(prow.get("Feed Per Day", "")).strip() or "-",
                    "abw": str(prow.get("ABW", "")).strip() or "-",
                    "expect_harvest": _expecting_harvest_display(prow, status),
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

        # ---- L.V.D for the farm = the LATEST date among this farm's own
        # ponds' individual L.V.D dates (each pond's own most recently
        # saved "Date" -- the same field the Marketing Manager app's Pond
        # Layout cards label "L.V.D" per pond), rolled up to one date.
        _pond_dates = pd.to_datetime(group["Date"], errors="coerce").dropna()
        lvd = _pond_dates.max().strftime("%Y-%m-%d") if len(_pond_dates) > 0 else "-"

        # ---- Last Feed Purchased Date / Last Feed Order, looked up by
        # this farm's Customer Code from the Sales Details sheet. Falls
        # back to "-" / no items when there's no code, no match, or the
        # sheet couldn't be reached.
        last_feed_date, last_feed_items = last_feed_lookup.get(code.strip().lower(), ("-", []))

        farms.append({
            "customer": str(customer),
            "farm": str(farm),
            "zone": zone,
            "zone_color": zone_colors.get(zone, "#2563eb"),
            "ponds": ponds,
            "map_image": map_image,
            "map_bbox": map_bbox,
            "map_polygon": map_polygon,
            "lvd": lvd,
            "feed_table": feed_table,
            "last_feed_date": last_feed_date if last_feed_date and last_feed_date != "nan" else "-",
            "last_feed_items": last_feed_items,
        })

    farms.sort(key=lambda f: (f["zone"], f["customer"], f["farm"]))
    return farms

# =========================================================================
# LOAD + COMPUTE
# =========================================================================
if st.button("🔄 Refresh Now"):
    st.rerun()

df = load_data()

# Sales Details is read best-effort -- if that sheet can't be reached
# (not shared with the service account, wrong ID, etc.) every farm just
# falls back to showing "-" for Last Feed Purchased Date / Last Feed
# Order instead of breaking the rest of the display.
try:
    sales_df = load_sales_data()
    last_feed_lookup = build_last_feed_order_lookup(sales_df)
except Exception:
    last_feed_lookup = {}

running_farms = build_running_farms(df, last_feed_lookup)

st.caption(
    f"{len(running_farms)} running farm(s) • opens in Full Screen • choose a Zone and a Customer / Farm, then use Back / Next "
    f"(or swipe) to move between farms • page auto-refreshes every {DATA_REFRESH_SECONDS // 60} min"
)

# =========================================================================
# RENDER -- self-contained HTML/CSS/JS component. The Zone selector, the
# Back / Next slide transitions and the Active Farms count all run
# entirely client-side; Python only supplies the data as JSON once per
# page load. The display always opens (and stays) in Full Screen mode --
# there is no toggle / exit button.
# =========================================================================
_HTML_TEMPLATE = """
<div id="kmn-wrap">
  <style>
    #kmn-wrap { font-family: 'Segoe UI', Tahoma, sans-serif; color:#1e293b; display:flex; flex-direction:column; }

    /* ---- top toolbar = Zone / Customer selectors */
    #kmn-toolbar {
      display:flex; align-items:center; justify-content:space-between; gap:10px; flex-wrap:wrap;
      background:#0f172a; border-radius:12px; padding:8px 12px; margin-bottom:10px; flex-shrink:0;
    }
    .kmn-zone-label { display:flex; align-items:center; gap:8px; color:#e2e8f0; font-size:.9rem; font-weight:700; }
    /* holds the Zone + Running Customers dropdowns */
    .kmn-filters { display:flex; align-items:center; gap:10px 16px; flex-wrap:wrap; }
    #kmn-zone-select, #kmn-customer-select {
      background:#1e293b; color:#f8fafc; border:1px solid rgba(255,255,255,.35); border-radius:8px;
      padding:8px 10px; font-size:1rem; min-height:40px; max-width:60vw;
    }

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
    /* ---- L.V.D line: the latest date among this farm's own ponds'
       individual L.V.D dates, shown just under the customer name. */
    .kmn-lvd-line { font-size: .85rem; color: #fbbf24; font-weight: 700; margin-top: 6px; }
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
    /* ---- small Pond No / Feed Per Day / ABW / Expecting Harvest table
       shown below the Pond Layout grid. */
    .kmn-feed-table-wrap {
      margin-top: 20px; width: 100%; max-width: 640px; background: rgba(15,23,42,.55);
      border-radius: 12px; padding: 12px 14px; box-sizing: border-box;
      border: 1px solid rgba(255,255,255,.12);
    }
    .kmn-feed-table-title {
      font-size: .85rem; font-weight: 800; color: #f8fafc; margin-bottom: 8px; text-align: center;
    }
    .kmn-feed-table { width: 100%; border-collapse: collapse; font-size: .8rem; }
    .kmn-feed-table th {
      text-align: left; padding: 5px 8px; color: #94a3b8; font-weight: 700;
      border-bottom: 1px solid rgba(255,255,255,.18); white-space: nowrap;
    }
    .kmn-feed-table td {
      text-align: left; padding: 5px 8px; color: #e2e8f0;
      border-bottom: 1px solid rgba(255,255,255,.08); white-space: nowrap;
    }
    /* ---- Last Feed Purchased Date / Last Feed Order block at the very
       bottom of each slide -- each feed item shown on its own line. */
    .kmn-last-feed-wrap {
      margin-top: 16px; width: 100%; max-width: 640px; background: rgba(15,23,42,.55);
      border-radius: 12px; padding: 12px 14px; box-sizing: border-box;
      border: 1px solid rgba(255,255,255,.12); text-align: center;
    }
    .kmn-last-feed-date { font-size: .85rem; font-weight: 800; color: #f8fafc; }
    .kmn-last-feed-title { font-size: .85rem; font-weight: 800; color: #f8fafc; margin-top: 8px; }
    .kmn-last-feed-item { font-size: .8rem; color: #e2e8f0; margin-top: 3px; }
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

    /* ---- bottom control bar: Back button | Active Farms count | Next button */
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

    /* ---- Full Screen: this is now the ONLY mode -- applied immediately
       on load. The component always fills the browser window; there is
       no toggle button and no exit. */
    #kmn-wrap.kmn-fullscreen-mode {
       position: fixed; inset: 0; z-index: 999999;
       min-height: 100vh; min-height: 100dvh; /* fallback in case the outer iframe never actually got resized */
       /* extra bottom padding keeps the Back / Next bar above the floating badges that
          Streamlit Cloud draws over the bottom-right corner of the page */
       background: #0b1220; padding: 10px 10px 68px; box-sizing: border-box; align-items: center;
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
      .kmn-feed-table-wrap, .kmn-last-feed-wrap { max-width: 100%; }
      .kmn-location-thumb { width: 180px; height: 120px; }
      #kmn-carousel { height: 520px; }
      /* Zone + Customer dropdowns share one row on phones */
      #kmn-zone-select, #kmn-customer-select { max-width: none; width: 100%; min-width: 0; flex: 1; }
      .kmn-filters { flex: 1 1 100%; flex-wrap: nowrap; }
      .kmn-zone-label { flex: 1 1 0; min-width: 0; }
      .kmn-lbl { display: none; }
      .kmn-nav-btn { padding: 12px 12px; min-width: 78px; }
    }
  </style>

  <div id="kmn-toolbar">
    <div class="kmn-filters">
      <label class="kmn-zone-label">📍 <span class="kmn-lbl">Zone</span>
        <select id="kmn-zone-select" title="Zone"></select>
      </label>
      <label class="kmn-zone-label">👤 <span class="kmn-lbl">Customer / Farm</span>
        <select id="kmn-customer-select" title="Customer name with farm name"></select>
      </label>
    </div>
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
      const customerSelect = document.getElementById('kmn-customer-select');
      const countEl = document.getElementById('kmn-active-count');
      const counterEl = document.getElementById('kmn-slide-counter');
      const prevBtn = document.getElementById('kmn-prev');
      const nextBtn = document.getElementById('kmn-next');
      const wrapEl = document.getElementById('kmn-wrap');

      let visible = farms.slice();
      let selectedZone = ALL;
      let current = 0;
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

      // ---- Customer / Farm dropdown: one entry per running farm in the
      // selected Zone, shown as "Customer name — Farm name" (grouped by Zone
      // when "All Zones" is selected). Picking one jumps straight to that farm.
      function buildFarmOptions() {
        let html = '', openZone = null;
        visible.forEach(function (f, i) {
          if (selectedZone === ALL && zoneKey(f) !== openZone) {
            if (openZone !== null) html += '</optgroup>';
            openZone = zoneKey(f);
            html += '<optgroup label="' + escapeHtml(zoneLabel(openZone)) + '">';
          }
          html += '<option value="' + i + '">' + escapeHtml(f.customer + ' — ' + f.farm) + '</option>';
        });
        if (openZone !== null) html += '</optgroup>';
        customerSelect.innerHTML = html;
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

          // ---- Small Pond No / Feed Per Day / ABW / Expecting Harvest
          // table, one row per pond (Full Harvest ponds excluded).
          const feedRows = (f.feed_table || []).map(function (r) {
            return '<tr>'
              + '<td>' + escapeHtml(r.pond_no) + '</td>'
              + '<td>' + escapeHtml(r.feed_per_day) + '</td>'
              + '<td>' + escapeHtml(r.abw) + '</td>'
              + '<td>' + escapeHtml(r.expect_harvest) + '</td>'
              + '</tr>';
          }).join('');
          const feedTableHtml = feedRows
            ? '<div class="kmn-feed-table-wrap">'
              + '<div class="kmn-feed-table-title">Feed / ABW / Expecting Harvest</div>'
              + '<table class="kmn-feed-table"><thead><tr>'
              + '<th>Pond No</th><th>Feed Per Day</th><th>ABW</th><th>Expecting Harvest</th>'
              + '</tr></thead><tbody>' + feedRows + '</tbody></table>'
              + '</div>'
            : '';

          // ---- Last Feed Purchased Date + Last Feed Order, each item on
          // its own line, at the very bottom of the slide.
          const lastFeedItemsHtml = (f.last_feed_items || []).length
            ? f.last_feed_items.map(function (item) {
                return '<div class="kmn-last-feed-item">' + escapeHtml(item) + '</div>';
              }).join('')
            : '<div class="kmn-last-feed-item">-</div>';
          const lastFeedHtml =
            '<div class="kmn-last-feed-wrap">'
            + '<div class="kmn-last-feed-date">Last Feed Purchased Date: ' + escapeHtml(f.last_feed_date || '-') + '</div>'
            + '<div class="kmn-last-feed-title">Last Feed Order</div>'
            + lastFeedItemsHtml
            + '</div>';

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
            + '<div class="kmn-lvd-line">L.V.D: ' + escapeHtml(f.lvd || '-') + '</div>'
            + '</div>'
            + '<div class="kmn-pond-grid">' + ponds + '</div>'
            + feedTableHtml
            + lastFeedHtml
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

      // ---- Active Farms count.
      function updateInfo() {
        const n = visible.length;
        countEl.textContent = selectedZone === ALL
          ? 'All Running Farms: ' + n
          : 'Active Farms in ' + zoneLabel(selectedZone) + ': ' + n;
        if (n) customerSelect.value = String(current); // dropdown always shows the farm on screen
        counterEl.textContent = n ? 'Farm ' + (current + 1) + ' of ' + n : '';
        prevBtn.disabled = nextBtn.disabled = n < 2;
      }

      function applyZone(zone, startIndex) {
        selectedZone = zone;
        zoneSelect.value = zone;
        visible = zone === ALL ? farms.slice() : farms.filter(function (f) { return zoneKey(f) === zone; });
        current = Math.min(Math.max(startIndex || 0, 0), Math.max(visible.length - 1, 0));
        buildFarmOptions();
        renderSlides();
        updateInfo();
        store('kmn_zone', zone);
        store('kmn_idx', String(current));
      }

      // Moves a slide back to its parked spot (off-screen right) without animating.
      function parkRight(s) {
        s.style.transition = 'none';
        s.classList.remove('active', 'leaving', 'from-left');
        void s.offsetWidth;
        s.style.transition = '';
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
              // After it has slid out to the left, park it back on the right
              // INSTANTLY (no transition) so it never sweeps back across the
              // screen behind the new slide.
              setTimeout(function () {
                if (!s.classList.contains('active')) parkRight(s);
              }, 650);
            }
          } else {
            parkRight(s);
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
      // Picking a "Customer — Farm" entry jumps straight to that farm's slide.
      customerSelect.addEventListener('change', function () {
        const idx = parseInt(this.value, 10);
        if (isNaN(idx) || idx === current) return;
        goTo(idx, idx > current ? 1 : -1);
      });
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

      document.addEventListener('keydown', function (e) {
        if (e.key === 'ArrowRight') step(1);
        else if (e.key === 'ArrowLeft') step(-1);
      });

      // ---- Full Screen is now the ONLY mode: applied immediately on
      // load, with no toggle button and no exit -- the whole component
      // fills the browser window (and the actual Streamlit component
      // iframe, same-origin, via window.frameElement) the instant the
      // page opens. A best-effort attempt at the browser's native
      // Fullscreen API is made too (browsers usually only allow that
      // after a tap/click, so the CSS-based full screen above is what
      // actually applies on open). Re-applied on resize/orientation
      // change and a few times right after load, since on some mobile
      // browsers window.frameElement isn't available yet on the very
      // first call, or the browser's own address-bar show/hide changes
      // the visible viewport height without re-firing our code -- both
      // of which could otherwise leave the component looking like it
      // "lost" its full-screen sizing (e.g. after tapping Next, a taller
      // farm's content sitting outside what actually got resized).
      function applyFullscreen() {
        wrapEl.classList.add('kmn-fullscreen-mode');
        if (!fsFrameEl) {
          try { fsFrameEl = window.frameElement; } catch (e) { fsFrameEl = null; }
        }
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
        try {
          if (document.documentElement.requestFullscreen) {
            document.documentElement.requestFullscreen().catch(function () {});
          }
        } catch (e) { /* ignored -- CSS-based fallback above still applies */ }
      }

      window.addEventListener('resize', applyFullscreen);
      window.addEventListener('orientationchange', applyFullscreen);
      // A few retries shortly after load -- covers mobile browsers where
      // window.frameElement isn't reachable yet on the very first call.
      [300, 1000, 2500].forEach(function (delay) { setTimeout(applyFullscreen, delay); });

      // ---- Start-up: restore the last Zone / farm (if any), then apply
      // Full Screen instantly.
      buildZoneOptions();
      const savedZone = recall('kmn_zone');
      const zoneOk = savedZone && Array.prototype.some.call(zoneSelect.options, function (o) { return o.value === savedZone; });
      applyZone(zoneOk ? savedZone : ALL, zoneOk ? parseInt(recall('kmn_idx') || '0', 10) : 0);
      applyFullscreen();

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

components.html(_html, height=760, scrolling=True)

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>KMN Aqua Services - Running Farms Live Display "
    "(read only, auto-updating)</p>",
    unsafe_allow_html=True,
)
