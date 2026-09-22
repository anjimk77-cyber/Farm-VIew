# Running Shrimp Farms - KMN (Live Display)

A read-only, kiosk/TV-style Streamlit app for KMN Aqua Services. It shares
the same Google Sheet and `Customer List.xlsx` as your other Shrimp
FarmFlow apps (data-entry app, full manager app, Marketing Manager &
Technician view) and never writes to the Sheet.

**What it shows**
- An auto-rotating carousel, one farm per slide, with a circular
  "zoom in" transition every 4 seconds — Customer Name, Farm Name with
  Code, and that farm's Pond Layout, only for farms that are currently
  **Running** (not every pond at Full Harvest). Each slide is tinted with
  a color for that farm's **Zone**.
- A continuously scrolling ticker bar at the bottom listing recent
  Harvest Details one after another, e.g.:
  `Wasantha Pathmakumara -C00876 Madurankuliya Farm Partial H from Pond No 1 - 1100 KG with 10 ABW at 2026/09/19`
- A **Stay** tick box that freezes the carousel and ticker in place (and
  pauses the page's periodic auto-refresh) so a viewer can read one
  slide/item without it moving.

## Files in this folder

```
running-farms-app/
├── running_farms_display.py   # the app
├── requirements.txt           # Python dependencies
├── Customer List.xlsx         # ← you add this (same file as your other apps)
└── .streamlit/
    └── secrets.toml           # ← you create this locally / in Streamlit Cloud (not committed)
```

## 1. Add your Customer List

Copy the same `Customer List.xlsx` used by your other apps into this
folder, next to `running_farms_display.py`. It must have at least these
columns: `Customer Name`, `Farm Name with Code`, `Zone`, `Area` (a
`Customer Code` column, if present, is used to prefix ticker lines).

## 2. Google Sheets access (same service account as your other apps)

This app reuses the **same** Google Cloud service account and the
**same** Google Sheet (`WaterQualityData` tab) as your data-entry /
manager / marketing-manager apps — no new sheet or credentials needed.

Create `.streamlit/secrets.toml` (locally, or in Streamlit Cloud's
"Secrets" settings — never commit this file):

```toml
[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "your-service-account@your-project.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."

[gsheet]
sheet_id = "your-google-sheet-id"
worksheet_name = "WaterQualityData"
```

Use the exact same values as your other apps' `secrets.toml` — this is
the same sheet, so the same service account already has access.

## 3. Push to GitHub

```bash
cd running-farms-app
git init
git add running_farms_display.py requirements.txt README.md "Customer List.xlsx"
git commit -m "Add Running Shrimp Farms live display app"
git branch -M main
git remote add origin https://github.com/<your-username>/running-shrimp-farms-kmn.git
git push -u origin main
```

`secrets.toml` should **not** be committed — add a `.gitignore`:

```
.streamlit/secrets.toml
```

## 4. Deploy on Streamlit Community Cloud

1. Go to https://share.streamlit.io → **New app**.
2. Pick your GitHub repo and branch, and set **Main file path** to
   `running_farms_display.py`.
3. Under **Advanced settings → Secrets**, paste the same content as your
   local `secrets.toml` (step 2 above).
4. Click **Deploy**. You'll get a URL like
   `https://running-shrimp-farms-kmn.streamlit.app` — that's the link to
   put on a TV/kiosk screen (open it in a browser and use fullscreen /
   kiosk mode).

## Adjusting the display

Near the top of `running_farms_display.py`:

```python
CAROUSEL_SECONDS = 4          # how long each farm slide stays on screen
TICKER_LOOP_SECONDS = 45      # how long one full pass of the ticker text takes
DATA_REFRESH_SECONDS = 300    # how often the page reloads to pull fresh sheet data
ZONE_PALETTE = [...]          # colors cycled across Zones
```

Pond box colors follow the same convention as your other apps: blue =
Running, yellow = Partial H, green = Full H, gray = Soon to be.
