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
#   1) An auto-rotating carousel (circular "zoom in" transition, one
#      farm per slide, changing every CAROUSEL_SECONDS) of every
#      Customer + Farm that is currently "Running" (i.e. NOT every pond
#      on that farm is at Full Harvest) -- Customer Name, Farm Name with
#      Code, and that farm's Pond Layout. Each slide is themed with a
#      color for that farm's Zone (from Customer List.xlsx). Each pond
#      card uses the SAME detailed Pond Layout design as the Marketing
#      Manager & Technician view app (status colors, WQ Special Case
#      icon, Stocking Density / L.V.D / Feed per Day / ABW, Total
#      Harvest KG, Expecting Harvest / Harvest Weight, and Issues).
#   2) A "Harvest Updates" panel that shows recent Harvest Details ONE AT
#      A TIME, changing every HARVEST_UPDATE_SECONDS, e.g.:
#      "Wasantha Pathmakumara -C00876 Madurankuliya Farm Partial H from
#       Pond No 1 - 1100 KG with 10 ABW at 2026/09/19"
#   3) A "Stay" tick box that freezes BOTH the carousel and the Harvest
#      Updates panel in place (and pauses the page's periodic data
#      refresh) so a viewer can read one slide/item without it moving on.
#   4) A "⛶ Full Screen" button that expands the whole display (Pond
#      Layout carousel + Harvest Updates panel) to fill the entire
#      browser window -- handy for a TV / kiosk screen.
#
# All of the carousel/rotation/full-screen behaviour runs client-side in
# a single self-contained HTML/CSS/JS component
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

# ---- Display timing (seconds) -- tweak these to taste.
CAROUSEL_SECONDS = 4          # how long each farm's Pond Layout slide stays on screen
HARVEST_UPDATE_SECONDS = 4    # how long each Harvest Updates item stays on screen
DATA_REFRESH_SECONDS = 300    # how often the whole page reloads to pull fresh sheet data

# ---- Zone color palette -- cycles if there are more zones than colors.
ZONE_PALETTE = [
    "#2563eb", "#dc2626", "#059669", "#d97706", "#7c3aed",
    "#0891b2", "#db2777", "#65a30d", "#ea580c", "#4f46e5",
]

_CUSTOMER_CODE_COLUMN_CANDIDATES = [
    "Customer Code", "Customer ID", "Customer Code with Code", "Code", "Cust Code",
]

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
    # (outside COLUMN_ORDER, same as "Harvest Status") -- kept here so the
    # Pond Layout cards below can show the same sad-face icon/label used
    # by the Marketing Manager & Technician view app.
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
# HELPERS -- ported from the Marketing Manager view app's Pond Layout
# logic, kept self-contained here so this display uses the exact same
# per-pond fields and rules (status, colors, WQ Special Case icon,
# Stocking Density / L.V.D / Feed per Day / ABW, Total Harvest KG,
# Expecting Harvest / Harvest Weight, Issues).
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
    # Same palette as the Marketing Manager & Technician view app's
    # Pond Layout: light yellow / green / gray boxes, default light blue.
    return {
        "Partial H": "#fff3cd",
        "Full H": "#d4edda",
        "Soon to be": "#e2e2e2",
    }.get(status, "#eaf4ff")

def _species_letter(species_culture):
    s = str(species_culture).strip().lower()
    if "vannamei" in s:
        return "V"
    elif "monodon" in s:
        return "M"
    return ""

def _parse_pond_harvest_kg(raw_value):
    """Same combined-harvest parser used elsewhere in this app family:
    'A (2000)' -> Full total. '2000 (2)' -> per-pond share (2000 / 2).
    Plain numbers are returned as-is."""
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

def _harvest_kg_sum_row(row):
    """Per-row contribution to a pond's running Total Harvest KG -- both
    harvest slots count whenever that slot's own Harvest Type is filled
    in, same rule used by the Marketing Manager & Technician view app."""
    row_total = 0.0
    t1 = str(row.get("Harvest Type", "")).strip()
    kg1 = _parse_pond_harvest_kg(row.get("Harvest KG", ""))
    if t1 and pd.notna(kg1):
        row_total += kg1
    t2 = str(row.get("Harvest Type 2", "")).strip()
    kg2 = _parse_pond_harvest_kg(row.get("Harvest KG 2", ""))
    if t2 and pd.notna(kg2):
        row_total += kg2
    return row_total

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

# =========================================================================
# BUILD "RUNNING FARMS" DATA FOR THE CAROUSEL
#
# A farm counts as "Running" for this display when NOT every pond it has
# a saved record for is currently at Full Harvest (same rule the Last
# Visit Date Report uses elsewhere in this app family) -- individual
# ponds within a running farm can still show Partial H / Full H / Soon
# to be, they just aren't ALL at Full H yet.
#
# Each pond now carries the SAME fields as the Marketing Manager &
# Technician view app's Pond Layout cards, so this display can render
# identical-looking pond boxes.
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

    # Running Total Harvest KG per pond -- summed across EVERY saved
    # record for that pond (not just its latest one), same as the
    # Marketing Manager & Technician view app's Pond Layout cards.
    total_harvest_kg_by_pond = (
        work.assign(_HarvestKGRow=work.apply(_harvest_kg_sum_row, axis=1))
        .groupby(["Customer", "Farm Name with Code", "Pond Number"])["_HarvestKGRow"]
        .sum()
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
            pond_no = prow.get("Pond Number", "")

            total_kg = total_harvest_kg_by_pond.get((customer, farm, pond_no), 0) or 0
            total_kg_str = f"{total_kg:,.2f}" if total_kg else ""

            density_val = pd.to_numeric(prow.get("Density", ""), errors="coerce")
            density_str = f"{density_val:,.0f}" if pd.notna(density_val) else "-"
            lvd_str = str(prow.get("Date", "")).strip() or "-"
            feed_day_str = str(prow.get("Feed Per Day", "")).strip() or "-"
            abw_str = str(prow.get("ABW", "")).strip() or "-"

            issues_str = str(prow.get("Issues", "")).strip()
            if issues_str.lower() == "nan":
                issues_str = ""
            wq_special_str = str(prow.get("WQ Special Cases", "")).strip()
            if wq_special_str.lower() == "nan":
                wq_special_str = ""

            species_letter = _species_letter(prow.get("Species Culture", ""))

            doc_today_str = ""
            started_label = ""
            harvest_date_str = ""

            if status == "Full H":
                harvest_date_str = (
                    str(prow.get("Harvest Date 2", "")).strip() or str(prow.get("Harvest Date", "")).strip() or "-"
                )
                t2_expect = str(prow.get("Harvest Type 2", "")).strip().lower()
                kg2_expect = _parse_pond_harvest_kg(prow.get("Harvest KG 2", ""))
                kg1_expect = _parse_pond_harvest_kg(prow.get("Harvest KG", ""))
                harvest_kg_val = kg2_expect if ("full" in t2_expect and pd.notna(kg2_expect)) else kg1_expect
                expect_label = "Harvest Weight"
                expect_val = f"{harvest_kg_val:,.2f} KG" if pd.notna(harvest_kg_val) else "-"
            elif status == "Soon to be":
                expect_label = "Expecting Harvest"
                expect_val = "-"
            else:
                expect_label = "Expecting Harvest"
                expect_kg = pd.to_numeric(prow.get("Expect Harvest (KG)", ""), errors="coerce")
                expect_val = f"{expect_kg:,.2f} KG" if pd.notna(expect_kg) else "-"

            if status not in ("Full H", "Soon to be"):
                doc_today_str = str(doc_val) if doc_val is not None else "-"
                try:
                    started_date = (
                        pd.Timestamp(date.today()) - pd.Timedelta(days=int(float(doc_val)))
                    ).strftime("%Y-%m-%d")
                    started_label = f"Started on {started_date}"
                except (TypeError, ValueError):
                    started_label = "Started on ---"

            ponds.append({
                "pond_no": str(pond_no),
                "status": status,
                "color": _pond_color(status),
                "species_letter": species_letter,
                "wq_special": wq_special_str,
                "issues": issues_str,
                "doc_today": doc_today_str,
                "started_label": started_label,
                "harvest_date": harvest_date_str,
                "total_harvest_kg": total_kg_str,
                "expect_label": expect_label,
                "expect_val": expect_val,
                "density": density_str,
                "lvd": lvd_str,
                "feed_day": feed_day_str,
                "abw": abw_str,
            })

        farms.append({
            "customer": str(customer),
            "farm": str(farm),
            "zone": zone,
            "zone_color": zone_colors.get(zone, "#2563eb"),
            "ponds": ponds,
        })

    farms.sort(key=lambda f: (f["zone"], f["customer"], f["farm"]))
    return farms

# =========================================================================
# BUILD HARVEST UPDATES ITEMS
#
# Splits each saved row into up to two harvest events (1st slot / 2nd
# slot, same convention used by "All Harvest Details" elsewhere in this
# app family), formatted as one line each, most recent first. These are
# shown ONE AT A TIME (rotating every HARVEST_UPDATE_SECONDS) rather than
# as a scrolling ticker.
# =========================================================================
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

    code_lookup = _customer_code_lookup()
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
ticker_items = build_harvest_ticker(df)

st.caption(
    f"{len(running_farms)} running farm(s) shown • Pond Layout rotates every {CAROUSEL_SECONDS}s • "
    f"Harvest Updates rotate every {HARVEST_UPDATE_SECONDS}s • "
    f"page auto-refreshes every {DATA_REFRESH_SECONDS // 60} min (paused while 'Stay' is ticked)"
)

# =========================================================================
# RENDER -- self-contained HTML/CSS/JS component. All animation (the
# circular zoom carousel transition for Pond Layout, the one-by-one
# Harvest Updates rotation, the "Stay" freeze toggle, and the Full Screen
# toggle) runs entirely client-side; Python only supplies the data as
# JSON once per page load.
# =========================================================================
_HTML_TEMPLATE = """
<div id="kmn-wrap">
  <style>
    #kmn-wrap { font-family: 'Segoe UI', Tahoma, sans-serif; color:#1e293b; display:flex; flex-direction:column; }
    #kmn-carousel {
      position: relative; width: 100%; height: 620px; overflow: hidden;
      border-radius: 16px; background: radial-gradient(circle at 50% 35%, #1e293b, #0f172a);
      box-shadow: 0 8px 30px rgba(0,0,0,.28); flex-shrink: 0;
    }
    .kmn-slide {
      position: absolute; inset: 0; display: flex; flex-direction: column;
      align-items: center; justify-content: flex-start; padding: 34px 20px 10px;
      opacity: 0; transform: scale(.8); transition: opacity 1s ease, transform 1s ease;
      pointer-events: none;
    }
    .kmn-slide.active { opacity: 1; transform: scale(1); pointer-events: auto; z-index: 2; }
    .kmn-slide-header { text-align: center; margin-bottom: 14px; flex-shrink: 0; }
    .kmn-zone-badge {
      display: inline-block; padding: 5px 16px; border-radius: 999px; font-size: .8rem;
      font-weight: 700; color: #fff; letter-spacing: .03em; margin-bottom: 8px;
    }
    .kmn-farm-name { font-size: 1.75rem; font-weight: 800; color: #f8fafc; line-height: 1.2; }
    .kmn-customer-name { font-size: 1.1rem; color: #cbd5e1; margin-top: 2px; }

    /* ---- Pond Layout cards: same design as the Marketing Manager &
       Technician view app's Pond Layout section. */
    .kmn-pond-grid {
      display: flex; flex-wrap: wrap; justify-content: center; gap: 16px;
      width: 100%; max-width: 1150px; flex: 1; overflow-y: auto; padding: 4px 6px 12px;
    }
    .kmn-pond-card { display: flex; flex-direction: column; align-items: center; margin: 4px; }
    .kmn-pond-box {
      position: relative; width: 200px; min-height: 168px; border: 2px solid #1e293b;
      border-radius: 8px; display: flex; flex-direction: column; align-items: center;
      justify-content: flex-start; padding: 8px 0; box-shadow: 0 3px 10px rgba(0,0,0,.25);
    }
    .kmn-pond-label { font-size: .78rem; color: #555; }
    .kmn-status-full { font-size: 1.15rem; font-weight: 800; color: red; }
    .kmn-status-soon { font-size: 1.05rem; font-weight: 800; color: #555; }
    .kmn-doc-today { font-size: 1.35rem; font-weight: 800; color: red; margin: 1px 0; }
    .kmn-started { font-size: .68rem; color: #777; }
    .kmn-subline { font-size: .7rem; color: #333; }
    .kmn-expect {
      font-size: .78rem; color: #333; text-align: center; width: 100%; margin-top: 4px;
      border-top: 1px dashed #bbb; padding-top: 3px;
    }
    .kmn-extra {
      font-size: .74rem; color: #333; text-align: left; width: 100%; padding: 0 8px;
      margin-top: 4px; line-height: 1.35;
    }
    .kmn-issues {
      margin-top: auto; width: 100%; text-align: center; font-size: .78rem; font-weight: 700;
      border-top: 1px dashed #bbb; padding-top: 3px; color: red;
    }
    .kmn-wq-icon { position: absolute; top: 2px; right: 4px; font-size: 1.2rem; line-height: 1; }
    .kmn-wq-text { font-size: .78rem; color: #fbbf24; text-align: center; max-width: 190px; margin-top: 2px; }
    .kmn-species { font-size: .72rem; font-weight: 700; color: #e2e8f0; margin-top: 2px; }

    .kmn-empty { color: #94a3b8; font-size: 1.2rem; margin-top: 60px; text-align: center; }
    .kmn-dots { position: absolute; bottom: 14px; left: 0; right: 0; display: flex; justify-content: center; gap: 7px; z-index: 3; }
    .kmn-dot { width: 8px; height: 8px; border-radius: 50%; background: rgba(255,255,255,.28); transition: background .3s; }
    .kmn-dot.active { background: #fff; }

    .kmn-fs-btn {
      position: absolute; top: 10px; right: 14px; z-index: 6;
      background: rgba(15,23,42,.65); color: #fff; border: 1px solid rgba(255,255,255,.35);
      border-radius: 8px; padding: 6px 12px; font-size: .8rem; cursor: pointer;
    }
    .kmn-fs-btn:hover { background: rgba(15,23,42,.9); }

    /* ---- Harvest Updates: one item at a time, rotates every
       HARVEST_UPDATE_SECONDS, instead of a scrolling ticker. */
    #kmn-harvest-panel {
      margin-top: 14px; background: #0f172a; border-radius: 12px; padding: 12px 16px;
      display: flex; align-items: center; gap: 16px; flex-shrink: 0;
    }
    .kmn-harvest-header { color: #fbbf24; font-weight: 800; font-size: .85rem; white-space: nowrap; }
    #kmn-harvest-counter { color: #94a3b8; font-weight: 600; font-size: .75rem; margin-left: 6px; }
    #kmn-harvest-viewport { flex: 1; overflow: hidden; position: relative; min-height: 26px; }
    #kmn-harvest-text {
      color: #e2e8f0; font-size: .95rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
      display: block; opacity: 0; transform: translateY(4px); transition: opacity .4s ease, transform .4s ease;
    }
    #kmn-harvest-text.kmn-fade-in { opacity: 1; transform: translateY(0); }
    .kmn-stay-toggle {
      display: flex; align-items: center; gap: 6px; color: #e2e8f0; font-size: .9rem;
      white-space: nowrap; cursor: pointer; user-select: none;
    }
    .kmn-stay-toggle input { width: 16px; height: 16px; cursor: pointer; }

    /* ---- Full Screen mode: expands the whole component (Pond Layout
       carousel + Harvest Updates) to fill the entire browser window. */
    #kmn-wrap.kmn-fullscreen-mode {
      position: fixed; inset: 0; width: 100vw; height: 100vh; z-index: 999999;
      background: #0b1220; padding: 14px; box-sizing: border-box;
    }
    #kmn-wrap.kmn-fullscreen-mode #kmn-carousel { flex: 1; height: auto; }
  </style>

  <div id="kmn-carousel">
    <div id="kmn-slides"></div>
    <div class="kmn-dots" id="kmn-dots"></div>
    <button id="kmn-fullscreen-btn" class="kmn-fs-btn" title="Toggle full screen">⛶ Full Screen</button>
  </div>

  <div id="kmn-harvest-panel">
    <span class="kmn-harvest-header">🌾 HARVEST UPDATES<span id="kmn-harvest-counter"></span></span>
    <div id="kmn-harvest-viewport">
      <span id="kmn-harvest-text"></span>
    </div>
    <label class="kmn-stay-toggle">
      <input type="checkbox" id="kmn-stay-toggle"> Stay (freeze)
    </label>
  </div>

  <script>
    (function () {
      const farms = __FARMS_JSON__;
      const ticker = __TICKER_JSON__;
      const carouselSeconds = __CAROUSEL_SECONDS__;
      const harvestUpdateSeconds = __HARVEST_UPDATE_SECONDS__;
      const dataRefreshSeconds = __DATA_REFRESH_SECONDS__;

      const slidesEl = document.getElementById('kmn-slides');
      const dotsEl = document.getElementById('kmn-dots');
      const harvestTextEl = document.getElementById('kmn-harvest-text');
      const harvestCounterEl = document.getElementById('kmn-harvest-counter');
      const stayCheckbox = document.getElementById('kmn-stay-toggle');
      const fsButton = document.getElementById('kmn-fullscreen-btn');
      const wrapEl = document.getElementById('kmn-wrap');

      let current = 0;
      let paused = false;
      let carouselTimer = null;
      let harvestIndex = 0;
      let harvestTimer = null;
      let isFullscreen = false;
      let fsFrameEl = null;
      try { fsFrameEl = window.frameElement; } catch (e) { fsFrameEl = null; }

      function escapeHtml(v) {
        return String(v == null ? '' : v)
          .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      }

      // ---- Pond Layout card markup -- mirrors the Marketing Manager &
      // Technician view app's Pond Layout cards field-for-field.
      function buildPondBoxHtml(p, zoneColor) {
        const wqIcon = p.wq_special ? '<div class="kmn-wq-icon" title="WQ Special Case">🫨</div>' : '';
        const wqText = p.wq_special ? '<div class="kmn-wq-text">🫨 ' + escapeHtml(p.wq_special) + '</div>' : '';
        const speciesHtml = p.species_letter ? '<div class="kmn-species">' + escapeHtml(p.species_letter) + '</div>' : '';
        const issuesHtml = p.issues ? '<div class="kmn-issues">' + escapeHtml(p.issues) + '</div>' : '';

        let middleHtml;
        if (p.status === 'Full H') {
          const totalHtml = p.total_harvest_kg
            ? '<div class="kmn-subline">Total: ' + escapeHtml(p.total_harvest_kg) + ' KG</div>' : '';
          middleHtml = '<div class="kmn-status-full">Full H</div>'
            + '<div class="kmn-subline">Harvest Date - ' + escapeHtml(p.harvest_date || '-') + '</div>'
            + totalHtml;
        } else if (p.status === 'Soon to be') {
          middleHtml = '<div class="kmn-status-soon">Soon to be</div>';
        } else {
          const totalHtml = (p.status === 'Partial H' && p.total_harvest_kg)
            ? '<div class="kmn-subline">Total: ' + escapeHtml(p.total_harvest_kg) + ' KG</div>' : '';
          middleHtml = '<div class="kmn-doc-today">' + escapeHtml(p.doc_today || '-') + '</div>'
            + '<div class="kmn-started">' + escapeHtml(p.started_label || '') + '</div>'
            + totalHtml;
        }

        const expectHtml = '<div class="kmn-expect"><b>' + escapeHtml(p.expect_label) + ':</b> '
          + escapeHtml(p.expect_val) + '</div>';
        const extraHtml = '<div class="kmn-extra">'
          + '<div>Stocking Density - ' + escapeHtml(p.density) + '</div>'
          + '<div>L.V.D - ' + escapeHtml(p.lvd) + '</div>'
          + '<div>Feed/Day - ' + escapeHtml(p.feed_day) + ' &nbsp;|&nbsp; ABW - ' + escapeHtml(p.abw) + '</div>'
          + '</div>';

        return '<div class="kmn-pond-card">'
          + '<div class="kmn-pond-box" style="background:' + p.color + '; border-color:' + zoneColor + ';">'
          + wqIcon
          + '<div class="kmn-pond-label">Pond ' + escapeHtml(p.pond_no) + '</div>'
          + middleHtml
          + expectHtml
          + extraHtml
          + issuesHtml
          + '</div>'
          + speciesHtml
          + wqText
          + '</div>';
      }

      function renderSlides() {
        if (!farms.length) {
          slidesEl.innerHTML = '<div class="kmn-slide active"><div class="kmn-empty">No running farms found.</div></div>';
          return;
        }
        slidesEl.innerHTML = farms.map(function (f, i) {
          const ponds = f.ponds.map(function (p) { return buildPondBoxHtml(p, f.zone_color); }).join('');
          return '<div class="kmn-slide' + (i === 0 ? ' active' : '') + '" data-index="' + i + '">'
            + '<div class="kmn-slide-header">'
            + '<div class="kmn-zone-badge" style="background:' + f.zone_color + ';">Zone ' + escapeHtml(f.zone || '-') + '</div>'
            + '<div class="kmn-farm-name">' + escapeHtml(f.farm) + '</div>'
            + '<div class="kmn-customer-name">' + escapeHtml(f.customer) + '</div>'
            + '</div>'
            + '<div class="kmn-pond-grid">' + ponds + '</div>'
            + '</div>';
        }).join('');

        dotsEl.innerHTML = farms.map(function (_, i) {
          return '<div class="kmn-dot' + (i === 0 ? ' active' : '') + '"></div>';
        }).join('');
      }

      function goTo(index) {
        const slides = slidesEl.querySelectorAll('.kmn-slide');
        const dots = dotsEl.querySelectorAll('.kmn-dot');
        slides.forEach(function (s, i) { s.classList.toggle('active', i === index); });
        dots.forEach(function (d, i) { d.classList.toggle('active', i === index); });
        current = index;
      }

      function nextSlide() {
        if (paused || farms.length < 2) return;
        goTo((current + 1) % farms.length);
      }

      function startCarousel() {
        if (carouselTimer) clearInterval(carouselTimer);
        carouselTimer = setInterval(nextSlide, carouselSeconds * 1000);
      }

      // ---- Harvest Updates: shows one item at a time, fading in, and
      // advances to the next item every harvestUpdateSeconds.
      function renderHarvestItem() {
        if (!ticker.length) {
          harvestTextEl.textContent = 'No harvest activity recorded yet.';
          harvestCounterEl.textContent = '';
          harvestTextEl.classList.add('kmn-fade-in');
          return;
        }
        harvestTextEl.classList.remove('kmn-fade-in');
        void harvestTextEl.offsetWidth; // restart the fade-in transition
        harvestTextEl.textContent = ticker[harvestIndex];
        harvestCounterEl.textContent = ' (' + (harvestIndex + 1) + ' / ' + ticker.length + ')';
        requestAnimationFrame(function () { harvestTextEl.classList.add('kmn-fade-in'); });
      }

      function nextHarvestItem() {
        if (paused || ticker.length < 2) return;
        harvestIndex = (harvestIndex + 1) % ticker.length;
        renderHarvestItem();
      }

      function startHarvestRotation() {
        if (harvestTimer) clearInterval(harvestTimer);
        harvestTimer = setInterval(nextHarvestItem, harvestUpdateSeconds * 1000);
      }

      // ---- "Stay" freezes both the Pond Layout carousel and the
      // Harvest Updates rotation (and the periodic page refresh below).
      stayCheckbox.addEventListener('change', function () {
        paused = this.checked;
      });

      // ---- Full Screen: expands the whole component to fill the
      // browser window. Resizes the actual Streamlit component iframe
      // (same-origin, via window.frameElement) so it behaves like a true
      // full-screen kiosk view; also makes a best-effort attempt at the
      // browser's native Fullscreen API.
      function enterFullscreen() {
        isFullscreen = true;
        wrapEl.classList.add('kmn-fullscreen-mode');
        if (fsFrameEl) {
          fsFrameEl.style.position = 'fixed';
          fsFrameEl.style.top = '0';
          fsFrameEl.style.left = '0';
          fsFrameEl.style.width = '100vw';
          fsFrameEl.style.height = '100vh';
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
      });

      renderSlides();
      renderHarvestItem();
      startCarousel();
      startHarvestRotation();

      // Periodically reload the whole app so it pulls fresh data from the
      // Google Sheet. Skipped while "Stay" is ticked so a frozen screen
      // doesn't jump mid-review; the manual "Refresh Now" button above
      // still works at any time.
      setInterval(function () {
        if (!paused) {
          try { window.parent.location.reload(); } catch (e) { window.location.reload(); }
        }
      }, dataRefreshSeconds * 1000);
    })();
  </script>
</div>
"""

_html = (
    _HTML_TEMPLATE
    .replace("__FARMS_JSON__", json.dumps(running_farms))
    .replace("__TICKER_JSON__", json.dumps(ticker_items))
    .replace("__CAROUSEL_SECONDS__", json.dumps(CAROUSEL_SECONDS))
    .replace("__HARVEST_UPDATE_SECONDS__", json.dumps(HARVEST_UPDATE_SECONDS))
    .replace("__DATA_REFRESH_SECONDS__", json.dumps(DATA_REFRESH_SECONDS))
)

components.html(_html, height=760, scrolling=False)

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>KMN Aqua Services - Running Farms Live Display "
    "(read only, auto-updating)</p>",
    unsafe_allow_html=True,
)
