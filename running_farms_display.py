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
#      color for that farm's Zone (from Customer List.xlsx), and each
#      pond box keeps the same status colors used across this app family
#      (blue = Running, yellow = Partial H, green = Full H, gray = Soon
#      to be).
#   2) A continuously scrolling ticker bar along the bottom listing
#      recent Harvest Details one after another, e.g.:
#      "Wasantha Pathmakumara -C00876 Madurankuliya Farm Partial H from
#       Pond No 1 - 1100 KG with 10 ABW at 2026/09/19"
#   3) A "Stay" tick box that freezes BOTH the carousel and the ticker in
#      place (and pauses the page's periodic data refresh) so a viewer
#      can read one slide/item without it moving on.
#
# All of the carousel/ticker animation runs client-side in a single
# self-contained HTML/CSS/JS component (streamlit.components.v1.html) --
# Python only computes the data once per page load/refresh.
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
CAROUSEL_SECONDS = 4          # how long each farm slide stays on screen
TICKER_LOOP_SECONDS = 45      # how long one full pass of the ticker text takes
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
    if len(df) > 0:
        df = df[COLUMN_ORDER + ["Harvest Status", "Harvest Status 2"]]
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
# BUILD HARVEST TICKER ITEMS
#
# Splits each saved row into up to two harvest events (1st slot / 2nd
# slot, same convention used by "All Harvest Details" elsewhere in this
# app family), formatted as one ticker line each, most recent first.
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
    f"{len(running_farms)} running farm(s) shown • carousel rotates every {CAROUSEL_SECONDS}s • "
    f"page auto-refreshes every {DATA_REFRESH_SECONDS // 60} min (paused while 'Stay' is ticked)"
)

# =========================================================================
# RENDER -- self-contained HTML/CSS/JS component. All animation (the
# circular zoom carousel transition, the scrolling ticker, and the
# "Stay" freeze toggle) runs entirely client-side; Python only supplies
# the data as JSON once per page load.
# =========================================================================
_HTML_TEMPLATE = """
<div id="kmn-wrap">
  <style>
    #kmn-wrap { font-family: 'Segoe UI', Tahoma, sans-serif; color:#1e293b; }
    #kmn-carousel {
      position: relative; width: 100%; height: 560px; overflow: hidden;
      border-radius: 16px; background: radial-gradient(circle at 50% 35%, #1e293b, #0f172a);
      box-shadow: 0 8px 30px rgba(0,0,0,.28);
    }
    .kmn-slide {
      position: absolute; inset: 0; display: flex; flex-direction: column;
      align-items: center; justify-content: flex-start; padding: 34px 20px 10px;
      opacity: 0; transform: scale(.8); transition: opacity 1s ease, transform 1s ease;
      pointer-events: none;
    }
    .kmn-slide.active { opacity: 1; transform: scale(1); pointer-events: auto; z-index: 2; }
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
    .kmn-pond-box {
      width: 132px; min-height: 112px; border-radius: 12px; border: 2px solid rgba(255,255,255,.18);
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      box-shadow: 0 3px 10px rgba(0,0,0,.25);
    }
    .kmn-pond-label { font-size: .72rem; color: rgba(15,23,42,.75); font-weight: 600; }
    .kmn-pond-doc { font-size: 1.45rem; font-weight: 800; color: #0f172a; margin: 2px 0; }
    .kmn-pond-status { font-size: .68rem; font-weight: 700; color: rgba(15,23,42,.8); }
    .kmn-empty { color: #94a3b8; font-size: 1.2rem; margin-top: 60px; text-align: center; }
    .kmn-dots { position: absolute; bottom: 14px; left: 0; right: 0; display: flex; justify-content: center; gap: 7px; z-index: 3; }
    .kmn-dot { width: 8px; height: 8px; border-radius: 50%; background: rgba(255,255,255,.28); transition: background .3s; }
    .kmn-dot.active { background: #fff; }

    #kmn-ticker-bar {
      margin-top: 14px; background: #0f172a; border-radius: 12px; padding: 10px 16px;
      display: flex; align-items: center; gap: 16px;
    }
    .kmn-ticker-label { color: #fbbf24; font-weight: 800; font-size: .85rem; white-space: nowrap; }
    .kmn-ticker-viewport { flex: 1; overflow: hidden; white-space: nowrap; position: relative; height: 26px; }
    .kmn-ticker-track {
      display: inline-block; white-space: nowrap; position: absolute; top: 0; will-change: transform;
      animation-name: kmnTickerScroll; animation-timing-function: linear; animation-iteration-count: infinite;
      color: #e2e8f0; font-size: .95rem;
    }
    @keyframes kmnTickerScroll {
      from { transform: translateX(100%); }
      to   { transform: translateX(-100%); }
    }
    .kmn-stay-toggle {
      display: flex; align-items: center; gap: 6px; color: #e2e8f0; font-size: .9rem;
      white-space: nowrap; cursor: pointer; user-select: none;
    }
    .kmn-stay-toggle input { width: 16px; height: 16px; cursor: pointer; }
  </style>

  <div id="kmn-carousel">
    <div id="kmn-slides"></div>
    <div class="kmn-dots" id="kmn-dots"></div>
  </div>

  <div id="kmn-ticker-bar">
    <span class="kmn-ticker-label">🌾 HARVEST UPDATES</span>
    <div class="kmn-ticker-viewport" id="kmn-ticker-viewport">
      <span class="kmn-ticker-track" id="kmn-ticker-track"></span>
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
      const tickerLoopSeconds = __TICKER_LOOP_SECONDS__;
      const dataRefreshSeconds = __DATA_REFRESH_SECONDS__;

      const slidesEl = document.getElementById('kmn-slides');
      const dotsEl = document.getElementById('kmn-dots');
      const tickerTrack = document.getElementById('kmn-ticker-track');
      const stayCheckbox = document.getElementById('kmn-stay-toggle');

      let current = 0;
      let paused = false;
      let carouselTimer = null;

      function escapeHtml(v) {
        return String(v == null ? '' : v)
          .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      }

      function renderSlides() {
        if (!farms.length) {
          slidesEl.innerHTML = '<div class="kmn-slide active"><div class="kmn-empty">No running farms found.</div></div>';
          return;
        }
        slidesEl.innerHTML = farms.map(function (f, i) {
          const ponds = f.ponds.map(function (p) {
            return '<div class="kmn-pond-box" style="background:' + p.color + '; border-color:' + f.zone_color + ';">'
              + '<div class="kmn-pond-label">Pond ' + escapeHtml(p.pond_no) + '</div>'
              + '<div class="kmn-pond-doc">' + escapeHtml(p.display) + '</div>'
              + '<div class="kmn-pond-status">' + escapeHtml(p.status) + '</div>'
              + '</div>';
          }).join('');
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

      function renderTicker() {
        if (!ticker.length) {
          tickerTrack.textContent = 'No harvest activity recorded yet.';
          tickerTrack.style.animation = 'none';
          return;
        }
        tickerTrack.textContent = ticker.join('     •     ');
        tickerTrack.style.animationDuration = tickerLoopSeconds + 's';
      }

      stayCheckbox.addEventListener('change', function () {
        paused = this.checked;
        tickerTrack.style.animationPlayState = paused ? 'paused' : 'running';
      });

      renderSlides();
      renderTicker();
      startCarousel();

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
    .replace("__TICKER_LOOP_SECONDS__", json.dumps(TICKER_LOOP_SECONDS))
    .replace("__DATA_REFRESH_SECONDS__", json.dumps(DATA_REFRESH_SECONDS))
)

components.html(_html, height=680, scrolling=False)

st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: gray;'>KMN Aqua Services - Running Farms Live Display "
    "(read only, auto-updating)</p>",
    unsafe_allow_html=True,
)
