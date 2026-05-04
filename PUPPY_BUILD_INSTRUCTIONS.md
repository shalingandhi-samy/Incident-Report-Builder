# 🐕 RCA Incident Report Builder — Complete Build Instructions for Code Puppy

**Tool Name:** RCA Incident Report Builder (Two-Phase Workflow)  
**Version:** 3.0  
**Last Updated:** April 30, 2026  
**Original Authors:** Erica Wilson & Jose Porchas (built with Code Puppy)

> **For Code Puppy:** This package contains ALL source files ready to deploy. Read this document for context, then copy the `rca_app/` folder to your workspace and follow the Quick Start.

---

## 🎯 What This Tool Does

A **Safety Incident Root Cause Analysis (RCA)** web app with a two-phase workflow:

- **Phase 1 (with associate present):** Collect initial incident data → generate a **WMW-738 Excel form** for associate + manager signatures
- **Phase 2 (manager returns later):** Complete full RCA → generate an **RCA PowerPoint deck**

**Key Features:**
- Incidents saved locally as JSON (Phase 1 + Phase 2 data persists)
- Load saved incidents to continue work later
- Smart 5 Whys suggestions based on keywords in incident description
- Auto-calculated due dates based on countermeasure pyramid level
- Reenactment photo and CCTV screenshot uploads (placed in 2x2 grid on slides)
- Email integration via Outlook COM automation
- Walmart brand colors throughout

---

## 📦 Prerequisites — Template Files (User Must Provide)

These two files must be placed in the **workspace root** (one level ABOVE `rca_app/`):

1. `Blank RCA Deck.pptx` — RCA PowerPoint template with exactly **7 slides**
2. `WMW-738   new template.xlsx` — WMW-738 Excel template (note: **three spaces** in filename)

⚠️ **CRITICAL:** These templates come from the user's organization. The generators fill specific cell addresses and table positions. Do NOT create generic replacements — ask the user to provide the originals.

---

## 🏗️ Project Structure

```
workspace/
├── Blank RCA Deck.pptx                  ← User provides
├── WMW-738   new template.xlsx          ← User provides (3 spaces in name!)
└── rca_app/
    ├── main_v2.py                       ← FastAPI application
    ├── models.py                        ← Pydantic models & constants
    ├── pptx_generator.py                ← PowerPoint generation
    ├── xlsx_generator.py                ← Excel (WMW-738) generation
    ├── templates/
    │   ├── index_v2.html                ← Main form (~1400 lines)
    │   ├── phase1_success.html          ← Phase 1 completion page
    │   ├── pyramid_viewer.html          ← Countermeasure pyramid reference
    │   └── success.html                 ← RCA generation success + email
    ├── static/
    │   └── countermeasure_pyramid.png   ← Pyramid image (included)
    ├── incidents/                       ← Auto-created: saved JSON files
    └── outputs/                         ← Auto-created: generated PPTX/XLSX
```

---

## 🐍 Quick Start

```bash
# 1. Create venv in workspace root
uv venv .venv
.venv\Scripts\activate   # Windows

# 2. Install dependencies
uv pip install fastapi uvicorn jinja2 python-multipart pydantic python-pptx openpyxl pillow \
  --index-url https://pypi.ci.artifacts.walmart.com/artifactory/api/pypi/external-pypi/simple \
  --allow-insecure-host pypi.ci.artifacts.walmart.com

# 3. Ensure template files are in workspace root
# - Blank RCA Deck.pptx
# - WMW-738   new template.xlsx

# 4. Run the app
cd rca_app
python main_v2.py
# Opens on http://127.0.0.1:8000
```

---

## 🎨 Walmart Brand Colors

All UI uses these Tailwind custom colors:

| Token | Hex | Usage |
|---|---|---|
| `wm-blue` | `#0053e2` | Primary — header, buttons, links |
| `wm-blue-hover` | `#0046c7` | Blue hover state |
| `wm-spark` | `#ffc220` | Accent — badges, back buttons |
| `wm-spark-dark` | `#995213` | Warning text |
| `wm-green` | `#2a8703` | Success, Phase 2 header |
| `wm-red` | `#ea1100` | STOP box, errors |
| `wm-gray-160` | `#333333` | Body text, footer |

---

## 📄 File-by-File Overview

### 1. `models.py`
Pydantic models for all form data:
- `FiveWhysCategory` — one 4M row with why1–why5 + root_cause
- `FiveWhys` — problem_statement + four category fields (material, machine, method, human)
- `Countermeasure` — root_cause, countermeasure, owner, due_date, cost, pyramid_level (1-6)
- `ObjectDetails` — specifics, item_name, size, shape, weight, distance_reach
- `RCAIncident` — master model combining all fields

**Constants:**
- `PYRAMID_DAYS` — maps level 1-6 to days (120, 90, 25, 5, 5, 3)
- `PYRAMID_LABELS` — display labels for each level
- `INCIDENT_TYPES` — 7 incident type options
- `INCIDENT_CATEGORIES` — 3 category options (Material Handling, Struck By/Against, Slip/Trip/Fall)

### 2. `main_v2.py`
FastAPI application with these endpoints:

| Method | Path | Description |
|---|---|---|
| GET | `/` | Main form — passes incident_types, categories, pyramid_labels, saved_incidents |
| GET | `/pyramid` | Countermeasure pyramid reference page |
| GET | `/incident/{id}` | Returns saved incident JSON |
| DELETE | `/incident/{id}` | Deletes saved incident file |
| GET | `/reprint-738/{id}` | Regenerates WMW-738 Excel from saved JSON |
| POST | `/save-phase1` | Saves Phase 1 JSON + generates WMW-738 |
| POST | `/generate-rca` | Builds full RCA, generates PPTX, **saves ALL Phase 2 data** |
| POST | `/send-email` | Opens Outlook via PowerShell COM with PPTX attached |
| GET | `/incidents` | Lists all saved incidents |
| GET | `/download/{filename}` | Downloads file from outputs/ |

**Key Behaviors:**
- Incident ID format: `YYYYMMDD_HHMMSS_xxxx` (timestamp + 4-char UUID)
- `/generate-rca` saves ALL Phase 2 fields back to JSON so work persists
- Email body: optional custom message + "This RCA was developed with the assistance of Code Puppy and has been reviewed and validated before distribution."

### 3. `xlsx_generator.py`
Fills `WMW-738   new template.xlsx` with Phase 1 data using openpyxl.

**Cell mappings:**
```
E2  = site_location
E5  = assoc_hire_date
L5  = assoc_department
L6  = assoc_schedule
B11 = incident date (MM/DD/YYYY)
F11 = incident time (HH:MM AM/PM)
I11 = reported date
K11 = reported time
N11 = reporting_manager
C13 = assoc_last_name
K13 = assoc_first_name
P13 = assoc_mi
E14 = incident_types (joined with ", ")
J14 = doing_normal_duties
M14 = if_not_normal_duties
E15 = incident_location
E16 = equipment_make
J16 = equipment_model
N16 = equipment_asset_id
A20 = how_incident_occurred
A23 = injury_description
A26 = objects_involved
A29 = safety_accountability
A34 = behavior_conditions
A37 = prevention
E40 = witness_names
```

### 4. `pptx_generator.py`
Fills `Blank RCA Deck.pptx` with full RCA data using python-pptx.

**Slide mapping:**
- **Slide 1:** Incident Overview (location, job description, incident type checkboxes, dates, work status)
- **Slide 2:** Associate & Incident Details (work history, incident category checkboxes, object details)
- **Slide 3:** Incident Description (injury description, investigation account)
- **Slide 4:** Reenactment Photos (2x2 grid, up to 4 images)
- **Slide 5:** CCTV Screenshots (2x2 grid, up to 4 images)
- **Slide 6:** 5 Whys / 4Ms Matrix (problem statement + 4 rows × 7 cols table)
- **Slide 7:** Countermeasures (5 rows × 6 cols table)

### 5. `templates/index_v2.html`
The main form (~1400 lines). Key sections:

**Phase 1 Sections** (visible initially, hidden in Phase 2):
- Site & Incident Information
- Associate Information
- Incident Classification (checkboxes)
- Associate's Statement (for WMW-738)
- Safety Accountability Matrix History
- Phase 1 Submit Button

**Phase 2 Sections** (hidden initially, shown when loading saved incident):
- Phase 2 Header (green banner with "Reprint WMW-738" button)
- Editable Summary of Phase 1 Data
- **Incident Account (Based on Investigation)** — NOT "Manager's Account"
  - Label: "📝 Incident Account (Based on Investigation)"
  - Warning: "This is NOT the associate's statement. Document the facts based on your investigation — witness statements, CCTV review, and evidence collected."
- Object Details
- Associate Work History
- Incident Category (checkboxes)
- 5 Whys / 4Ms Matrix with smart suggestions
- Reenactment Photos Upload
- CCTV Screenshots Upload
- Countermeasures (neutral guidance, NO "Goal of 4 or higher" messaging)
- Generate RCA Button + Back to Phase 1 Button

**5 Whys Smart Suggestion System:**
- Triggers on typing in investigation account field (500ms debounce)
- Minimum 20 characters before analysis runs
- Scans text for keywords in 4 categories (material, machine, method, human)
- Highlights top 2 matching rows in yellow
- Shows leading questions in a yellow box with note: "💡 Based on keywords in your description:"
- Categories have extensive keyword lists (body mechanics verbs in all tenses, warehouse tasks, etc.)

**Countermeasures Section:**
- NO goal messaging (removed per AP team feedback)
- Neutral guidance: "Select the countermeasure level that best addresses each root cause. Due dates auto-calculate based on level."
- Plain gray reference bar showing all 6 levels with days
- All 4 rows in ONE gray container (not individual cards)
- Pyramid level select auto-calculates due date

**Phase Switching:**
- CSS classes: `.phase-2-active .phase-1-section { display: none; }` etc.
- `loadSavedIncident()` adds `phase-2-active` class to `#mainContent`
- Loads ALL Phase 2 fields when incident is loaded (5 Whys, countermeasures, etc.)

### 6. `templates/phase1_success.html`
Shown after Phase 1 completes:
- Success card with download WMW-738 button
- STOP box (red gradient) with 5 steps: Print → Associate signs → Manager signs → Turn in to AP → Return for RCA
- Return card with "Complete RCA Now" and "Start New Incident" buttons

### 7. `templates/success.html`
Shown after RCA PowerPoint generated:
- Success icon + download PowerPoint button
- Email form (recipients + optional message + "Open in Outlook" button)
- "Create Another Report" link

### 8. `templates/pyramid_viewer.html`
Reference page for countermeasure pyramid:
- Header with close button
- Info banner explaining the pyramid
- Pyramid image from `/static/countermeasure_pyramid.png`
- Quick reference table (Level, Name, Description, Due Date)

---

## ⚙️ Key Behaviors to Get Exactly Right

### 1. Phase 2 Data Persistence
- `/generate-rca` saves ALL Phase 2 fields back to the incident JSON
- `loadSavedIncident()` populates ALL Phase 2 fields when loading
- This allows users to: start Phase 2 → generate RCA → come back later → all work is still there

### 2. "Incident Account (Based on Investigation)" — NOT "Manager's Account"
- Per AP team feedback: avoid "manager's account" wording (implies opinion)
- Use: "Incident Account (Based on Investigation)"
- Guidance text emphasizes facts from investigation (witness statements, CCTV, evidence)

### 3. No "Goal of 4 or Higher" Messaging
- Per AP team feedback: don't push users toward a specific pyramid level
- Let the data lead them to the appropriate countermeasure level
- Removed green banner with goal messaging
- Removed green highlight on Level 4 in reference bar

### 4. 5 Whys Suggestion Box — Trimmed Down
- Simple note: "💡 Based on keywords in your description:"
- No lengthy 4M explanation paragraph
- Just the numbered questions

### 5. Countermeasures: One Box, Not Four
- All 4 countermeasure rows in ONE `bg-gray-50 rounded-lg` container
- Rows separated by `border-b border-gray-200` (no border on last row)
- Do NOT give each row its own individual card

### 6. Back to WMW-738 Button
- Appears ONLY at the bottom of Phase 2 content (left side of action row)
- Spark yellow background, calls `backToPhase1()`
- There is NO back button in the header

### 7. Email Body
- Optional custom message first (if provided)
- Then ONLY: "This RCA was developed with the assistance of Code Puppy and has been reviewed and validated before distribution."
- No "Hi,", no "Please find attached", no "Best regards"

### 8. Image Uploads
- Form must be `enctype="multipart/form-data"`
- Up to 4 reenactment photos, up to 4 CCTV screenshots
- Images placed in 2×2 grid with aspect-ratio preservation
- Temp upload dir cleaned up after PPTX generation

---

## 📎 Source Files Included in This Package

All files in `rca_app/` are ready to deploy:
- `main_v2.py`
- `models.py`
- `pptx_generator.py`
- `xlsx_generator.py`
- `templates/index_v2.html`
- `templates/phase1_success.html`
- `templates/pyramid_viewer.html`
- `templates/success.html`
- `static/countermeasure_pyramid.png`

**User must provide:**
- `Blank RCA Deck.pptx` (place in workspace root)
- `WMW-738   new template.xlsx` (place in workspace root — note 3 spaces!)

---

## 🚀 Deployment Checklist

1. ☐ Copy `rca_app/` folder to workspace
2. ☐ Get template files from user: `Blank RCA Deck.pptx` and `WMW-738   new template.xlsx`
3. ☐ Place templates in workspace root (one level above `rca_app/`)
4. ☐ Create venv: `uv venv .venv`
5. ☐ Install dependencies (see Quick Start)
6. ☐ Run: `cd rca_app && python main_v2.py`
7. ☐ Open browser to `http://127.0.0.1:8000`
8. ☐ Test Phase 1 workflow (save, generate WMW-738)
9. ☐ Test Phase 2 workflow (load incident, complete RCA, generate PPTX)
10. ☐ Test persistence (load incident, verify all Phase 2 data restored)

---

*Happy building! 🐶 — Built with Code Puppy*
