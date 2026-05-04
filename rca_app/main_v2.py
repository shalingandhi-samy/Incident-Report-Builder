"""RCA Form Filler - FastAPI Application (Two-Phase Workflow)."""
import json
import uuid
import shutil
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, Form, File, UploadFile, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from models import (
    RCAIncident, Countermeasure, ObjectDetails, FiveWhys, FiveWhysCategory,
    INCIDENT_TYPES, INCIDENT_CATEGORIES, PYRAMID_LABELS, PYRAMID_DAYS
)
from pptx_generator import generate_rca_pptx
from xlsx_generator import generate_wmw738

app = FastAPI(title="RCA Form Filler", description="Safety Incident RCA Generator (Two-Phase)")

# Setup paths
BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR.parent / "Blank RCA Deck.pptx"
UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUTS_DIR = BASE_DIR / "outputs"
INCIDENTS_DIR = BASE_DIR / "incidents"  # Store saved Phase 1 data
STATIC_DIR = BASE_DIR / "static"
UPLOADS_DIR.mkdir(exist_ok=True)
OUTPUTS_DIR.mkdir(exist_ok=True)
INCIDENTS_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)

# Mount static files (images, css, etc.)
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# Templates
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


@app.get("/pyramid", response_class=HTMLResponse)
async def pyramid_viewer(request: Request):
    """Display the Countermeasure Pyramid reference image in its own page."""
    return templates.TemplateResponse(
        request=request,
        name="pyramid_viewer.html",
        context={},
    )

def get_saved_incidents():
    """Get list of saved incidents for dropdown."""
    incidents = []
    for f in INCIDENTS_DIR.glob("*.json"):
        try:
            data = json.loads(f.read_text())
            label = f"{data.get('site_location', 'Unknown')} - {data.get('assoc_last_name', '')}, {data.get('assoc_first_name', '')} ({data.get('incident_datetime', 'No date')[:10]})"
            incidents.append({"id": f.stem, "label": label})
        except:
            pass
    # Sort by most recent first
    incidents.sort(key=lambda x: x['id'], reverse=True)
    return incidents


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Render the main RCA form (two-phase)."""
    return templates.TemplateResponse(
        request=request,
        name="index_v2.html",
        context={
            "incident_types": INCIDENT_TYPES,
            "incident_categories": INCIDENT_CATEGORIES,
            "pyramid_labels": list(PYRAMID_LABELS.items()),
            "saved_incidents": get_saved_incidents(),
        }
    )


@app.get("/incident/{incident_id}")
async def get_incident(incident_id: str):
    """Load saved incident data."""
    incident_file = INCIDENTS_DIR / f"{incident_id}.json"
    if not incident_file.exists():
        return JSONResponse({"error": "Incident not found"}, status_code=404)
    
    data = json.loads(incident_file.read_text())
    return JSONResponse(data)


@app.delete("/incident/{incident_id}")
async def delete_incident(incident_id: str):
    """Delete a saved incident."""
    incident_file = INCIDENTS_DIR / f"{incident_id}.json"
    if not incident_file.exists():
        return JSONResponse({"error": "Incident not found"}, status_code=404)
    
    incident_file.unlink()
    return JSONResponse({"success": True, "message": "Incident deleted"})


@app.get("/reprint-738/{incident_id}")
async def reprint_738(incident_id: str):
    """Regenerate and download WMW-738 for a saved incident."""
    incident_file = INCIDENTS_DIR / f"{incident_id}.json"
    if not incident_file.exists():
        return JSONResponse({"error": "Incident not found"}, status_code=404)
    
    data = json.loads(incident_file.read_text())
    
    # Generate the Excel file
    site_location = data.get("site_location", "Unknown").replace(" ", "_")
    xlsx_filename = f"WMW-738_{site_location}_{incident_id}.xlsx"
    xlsx_path = OUTPUTS_DIR / xlsx_filename
    
    xlsx_bytes = generate_wmw738(data)
    with open(xlsx_path, "wb") as f:
        f.write(xlsx_bytes)
    
    return FileResponse(
        path=xlsx_path,
        filename=xlsx_filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.post("/send-email")
async def send_email(request: Request):
    """Send RCA PowerPoint via email by opening Outlook."""
    import subprocess
    
    data = await request.json()
    to_emails = data.get("to", "")
    custom_message = data.get("message", "")
    filename = data.get("filename", "")
    site_location = data.get("site_location", "Unknown")
    
    if not to_emails or not filename:
        return JSONResponse({"error": "Missing email or filename"}, status_code=400)
    
    filepath = OUTPUTS_DIR / filename
    if not filepath.exists():
        return JSONResponse({"error": "File not found"}, status_code=404)
    
    abs_path = str(filepath.resolve())
    
    # Build email body
    email_body = ""
    if custom_message:
        email_body += f"{custom_message}\n\n"
    email_body += "This RCA was developed with the assistance of Code Puppy and has been reviewed and validated before distribution."
    
    # Escape for PowerShell
    email_body_escaped = email_body.replace('"', '`"').replace("\n", "`n")
    
    try:
        # Use PowerShell to create and display email with Outlook
        ps_script = f'''
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.To = "{to_emails}"
$mail.Subject = "RCA PowerPoint - {site_location}"
$mail.Body = "{email_body_escaped}"
$mail.Attachments.Add("{abs_path}")
$mail.Display()
'''
        result = subprocess.run(
            ["powershell", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if result.returncode == 0:
            return JSONResponse({"success": True, "message": "Outlook opened with draft email"})
        else:
            return JSONResponse({"error": f"PowerShell error: {result.stderr}", "success": False})
    except subprocess.TimeoutExpired:
        return JSONResponse({"success": True, "message": "Outlook opening..."})
    except Exception as e:
        return JSONResponse({"error": f"Could not open Outlook: {str(e)}", "success": False})


@app.get("/incidents")
async def list_incidents():
    """List all saved incidents."""
    return JSONResponse(get_saved_incidents())


@app.post("/save-phase1")
async def save_phase1(
    request: Request,
    # Site & Incident Info
    site_location: str = Form(""),
    incident_datetime: str = Form(""),
    reported_datetime: str = Form(""),
    reporting_manager: str = Form(""),
    associates_manager: str = Form(""),
    incident_location: str = Form(""),
    # Associate Info
    assoc_last_name: str = Form(""),
    assoc_first_name: str = Form(""),
    assoc_mi: str = Form(""),
    assoc_hire_date: str = Form(""),
    assoc_department: str = Form(""),
    tpr_code: str = Form(""),
    assoc_schedule: str = Form(""),
    doing_normal_duties: str = Form(""),
    on_normal_shift: str = Form(""),
    on_overtime: str = Form(""),
    if_not_normal_duties: str = Form(""),
    # Incident Classification
    incident_types: list[str] = Form([]),
    # Associate Statement
    how_incident_occurred: str = Form(""),
    injury_description: str = Form(""),
    objects_involved: str = Form(""),
    # Power Equipment
    equipment_make: str = Form(""),
    equipment_model: str = Form(""),
    equipment_asset_id: str = Form(""),
    # Witnesses
    witness_names: str = Form(""),
    # Safety & Prevention (for WMW-738)
    safety_accountability: str = Form(""),
    behavior_conditions: str = Form(""),
    prevention: str = Form(""),
):
    """Save Phase 1 data and generate WMW-738 Excel."""
    
    # Generate unique ID for this incident
    incident_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + str(uuid.uuid4())[:4]
    
    # Build data dict
    phase1_data = {
        "id": incident_id,
        "created_at": datetime.now().isoformat(),
        "phase": 1,
        # Site & Incident
        "site_location": site_location,
        "incident_datetime": incident_datetime,
        "reported_datetime": reported_datetime,
        "reporting_manager": reporting_manager,
        "associates_manager": associates_manager,
        "incident_location": incident_location,
        # Associate
        "assoc_last_name": assoc_last_name,
        "assoc_first_name": assoc_first_name,
        "assoc_mi": assoc_mi,
        "assoc_hire_date": assoc_hire_date,
        "assoc_department": assoc_department,
        "tpr_code": tpr_code,
        "assoc_schedule": assoc_schedule,
        "doing_normal_duties": doing_normal_duties,
        "on_normal_shift": on_normal_shift,
        "on_overtime": on_overtime,
        "if_not_normal_duties": if_not_normal_duties,
        # Incident
        "incident_types": incident_types if isinstance(incident_types, list) else [incident_types],
        "how_incident_occurred": how_incident_occurred,
        "injury_description": injury_description,
        "objects_involved": objects_involved,
        # Equipment
        "equipment_make": equipment_make,
        "equipment_model": equipment_model,
        "equipment_asset_id": equipment_asset_id,
        # Witnesses
        "witness_names": witness_names,
        # Safety & Prevention
        "safety_accountability": safety_accountability,
        "behavior_conditions": behavior_conditions,
        "prevention": prevention,
    }
    
    # Save to JSON file
    incident_file = INCIDENTS_DIR / f"{incident_id}.json"
    incident_file.write_text(json.dumps(phase1_data, indent=2))
    
    # Generate WMW-738 Excel - pass the dict directly
    xlsx_filename = f"WMW-738_{site_location.replace(' ', '_')}_{incident_id}.xlsx"
    xlsx_path = OUTPUTS_DIR / xlsx_filename
    xlsx_bytes = generate_wmw738(phase1_data)
    with open(xlsx_path, "wb") as f:
        f.write(xlsx_bytes)
    
    # Return success page with Excel download
    return templates.TemplateResponse(
        request=request,
        name="phase1_success.html",
        context={
            "incident_id": incident_id,
            "xlsx_filename": xlsx_filename,
            "site_location": site_location,
            "assoc_name": f"{assoc_first_name} {assoc_last_name}",
        }
    )


@app.post("/generate-rca")
async def generate_rca(
    request: Request,
    incident_id: str = Form(""),
    # Phase 1 fields (editable)
    site_location: str = Form(""),
    incident_datetime: str = Form(""),
    reported_datetime: str = Form(""),
    reporting_manager: str = Form(""),
    associates_manager: str = Form(""),
    incident_location: str = Form(""),
    assoc_last_name: str = Form(""),
    assoc_first_name: str = Form(""),
    assoc_mi: str = Form(""),
    assoc_hire_date: str = Form(""),
    assoc_department: str = Form(""),
    tpr_code: str = Form(""),
    assoc_schedule: str = Form(""),
    doing_normal_duties: str = Form(""),
    on_normal_shift: str = Form(""),
    on_overtime: str = Form(""),
    if_not_normal_duties: str = Form(""),
    incident_types: list[str] = Form([]),
    how_incident_occurred: str = Form(""),
    injury_description: str = Form(""),
    objects_involved: str = Form(""),
    equipment_make: str = Form(""),
    equipment_model: str = Form(""),
    equipment_asset_id: str = Form(""),
    witness_names: str = Form(""),
    safety_accountability: str = Form(""),
    behavior_conditions: str = Form(""),
    prevention: str = Form(""),
    # Phase 2 RCA fields
    manager_incident_account: str = Form(""),
    # Object details (Phase 2 research)
    object_item_name: str = Form(""),
    object_size: str = Form(""),
    object_weight: str = Form(""),
    object_shape: str = Form(""),
    object_distance: str = Form(""),
    object_specifics: str = Form(""),
    # Associate work history
    job_description: str = Form(""),
    kind_of_injury: str = Form(""),
    hours_in_path: str = Form(""),
    hours_this_week: str = Form(""),
    glide_path: str = Form(""),
    no_training_paperwork: str = Form(""),
    labor_share: str = Form(""),
    yes_training_paperwork: str = Form(""),
    previous_stand_down: str = Form(""),
    roster_reviewed: str = Form(""),
    countermeasure_updated: str = Form(""),
    last_safety_observation: str = Form(""),
    observation_findings: str = Form(""),
    disciplinary_actions: str = Form(""),
    incident_category: list[str] = Form([]),
    # 5 Whys
    problem_statement: str = Form(""),
    material_why1: str = Form(""), material_why2: str = Form(""), material_why3: str = Form(""),
    material_why4: str = Form(""), material_why5: str = Form(""), material_root_cause: str = Form(""),
    machine_why1: str = Form(""), machine_why2: str = Form(""), machine_why3: str = Form(""),
    machine_why4: str = Form(""), machine_why5: str = Form(""), machine_root_cause: str = Form(""),
    method_why1: str = Form(""), method_why2: str = Form(""), method_why3: str = Form(""),
    method_why4: str = Form(""), method_why5: str = Form(""), method_root_cause: str = Form(""),
    human_why1: str = Form(""), human_why2: str = Form(""), human_why3: str = Form(""),
    human_why4: str = Form(""), human_why5: str = Form(""), human_root_cause: str = Form(""),
    # Countermeasures
    cm1_root_cause: str = Form(""), cm1_countermeasure: str = Form(""), cm1_owner: str = Form(""), cm1_pyramid: int = Form(6),
    cm1_due_date: str = Form(""), cm1_cost: str = Form(""),
    cm2_root_cause: str = Form(""), cm2_countermeasure: str = Form(""), cm2_owner: str = Form(""), cm2_pyramid: int = Form(6),
    cm2_due_date: str = Form(""), cm2_cost: str = Form(""),
    cm3_root_cause: str = Form(""), cm3_countermeasure: str = Form(""), cm3_owner: str = Form(""), cm3_pyramid: int = Form(6),
    cm3_due_date: str = Form(""), cm3_cost: str = Form(""),
    cm4_root_cause: str = Form(""), cm4_countermeasure: str = Form(""), cm4_owner: str = Form(""), cm4_pyramid: int = Form(6),
    cm4_due_date: str = Form(""), cm4_cost: str = Form(""),
    # Photos
    reenactment_photos: list[UploadFile] = File(None),
    cctv_screenshots: list[UploadFile] = File(None),
):
    """Generate RCA PowerPoint using form data (Phase 1 + Phase 2 combined)."""
    
    # Parse dates
    def parse_date(s):
        if not s:
            return None
        try:
            return datetime.strptime(s, "%Y-%m-%d").date()
        except:
            return None
    
    def parse_datetime(s):
        if not s:
            return None
        try:
            return datetime.fromisoformat(s)
        except:
            return None
    
    # Build countermeasures
    countermeasures = []
    for i, (rc, cm, owner, pyr, due, cost) in enumerate([
        (cm1_root_cause, cm1_countermeasure, cm1_owner, cm1_pyramid, cm1_due_date, cm1_cost),
        (cm2_root_cause, cm2_countermeasure, cm2_owner, cm2_pyramid, cm2_due_date, cm2_cost),
        (cm3_root_cause, cm3_countermeasure, cm3_owner, cm3_pyramid, cm3_due_date, cm3_cost),
        (cm4_root_cause, cm4_countermeasure, cm4_owner, cm4_pyramid, cm4_due_date, cm4_cost),
    ], 1):
        countermeasures.append(Countermeasure(
            root_cause=rc,
            countermeasure=cm,
            owner=owner,
            pyramid_level=pyr,
            due_date=parse_date(due),
            cost=cost,
        ))
    
    # Build full RCA data from form fields
    data = RCAIncident(
        # Phase 1 fields (from form, possibly edited)
        site_location=site_location,
        incident_datetime=parse_datetime(incident_datetime),
        reported_datetime=parse_datetime(reported_datetime),
        reporting_manager=reporting_manager,
        associates_manager=associates_manager,
        incident_location=incident_location,
        incident_types=incident_types if isinstance(incident_types, list) else [incident_types],
        doing_normal_duties=doing_normal_duties,
        on_normal_shift=on_normal_shift,
        on_overtime=on_overtime,
        how_incident_occurred=how_incident_occurred,
        manager_incident_account=manager_incident_account,
        injury_description=injury_description,
        object_details=ObjectDetails(
            specifics=object_specifics,
            item_name=object_item_name,
            size=object_size,
            shape=object_shape,
            weight=object_weight,
            distance_reach=object_distance,
        ),
        # Phase 2 fields
        job_description=job_description,
        kind_of_injury=kind_of_injury,
        hours_in_path=hours_in_path,
        hours_this_week=hours_this_week,
        glide_path=glide_path,
        no_training_paperwork=no_training_paperwork,
        labor_share=labor_share,
        yes_training_paperwork=yes_training_paperwork,
        previous_stand_down=previous_stand_down,
        roster_reviewed=roster_reviewed,
        countermeasure_updated=countermeasure_updated,
        last_safety_observation=last_safety_observation,
        observation_findings=observation_findings,
        disciplinary_actions=disciplinary_actions,
        incident_category=incident_category if isinstance(incident_category, list) else [incident_category],
        five_whys=FiveWhys(
            problem_statement=problem_statement,
            material=FiveWhysCategory(
                why1=material_why1, why2=material_why2, why3=material_why3,
                why4=material_why4, why5=material_why5, root_cause=material_root_cause,
            ),
            machine=FiveWhysCategory(
                why1=machine_why1, why2=machine_why2, why3=machine_why3,
                why4=machine_why4, why5=machine_why5, root_cause=machine_root_cause,
            ),
            method=FiveWhysCategory(
                why1=method_why1, why2=method_why2, why3=method_why3,
                why4=method_why4, why5=method_why5, root_cause=method_root_cause,
            ),
            human=FiveWhysCategory(
                why1=human_why1, why2=human_why2, why3=human_why3,
                why4=human_why4, why5=human_why5, root_cause=human_root_cause,
            ),
        ),
        countermeasures=countermeasures,
    )
    
    # Handle uploaded images
    session_id = str(uuid.uuid4())[:8]
    upload_dir = UPLOADS_DIR / session_id
    upload_dir.mkdir(exist_ok=True)
    
    reenactment_paths = []
    cctv_paths = []
    
    if reenactment_photos:
        for i, photo in enumerate(reenactment_photos):
            if photo.filename and photo.size > 0:
                ext = Path(photo.filename).suffix or ".jpg"
                path = upload_dir / f"reenactment_{i}{ext}"
                with open(path, "wb") as f:
                    content = await photo.read()
                    f.write(content)
                reenactment_paths.append(path)
    
    if cctv_screenshots:
        for i, screenshot in enumerate(cctv_screenshots):
            if screenshot.filename and screenshot.size > 0:
                ext = Path(screenshot.filename).suffix or ".jpg"
                path = upload_dir / f"cctv_{i}{ext}"
                with open(path, "wb") as f:
                    content = await screenshot.read()
                    f.write(content)
                cctv_paths.append(path)
    
    # Generate PowerPoint
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    site_clean = site_location.replace(" ", "_") if site_location else "Unknown"
    output_filename = f"RCA_{site_clean}_{timestamp}.pptx"
    output_path = OUTPUTS_DIR / output_filename
    
    generate_rca_pptx(
        data=data,
        template_path=TEMPLATE_PATH,
        output_path=output_path,
        reenactment_photos=reenactment_paths,
        cctv_screenshots=cctv_paths,
    )
    
    # Cleanup uploads
    shutil.rmtree(upload_dir, ignore_errors=True)
    
    # Update saved incident with ALL Phase 2 data so it can be reloaded
    if incident_id:
        incident_file = INCIDENTS_DIR / f"{incident_id}.json"
        if incident_file.exists():
            saved_data = json.loads(incident_file.read_text())
            # Update with Phase 2 fields
            saved_data.update({
                "phase": 2,
                "rca_completed_at": datetime.now().isoformat(),
                # Manager's account
                "manager_incident_account": manager_incident_account,
                # Object details
                "object_item_name": object_item_name,
                "object_size": object_size,
                "object_weight": object_weight,
                "object_shape": object_shape,
                "object_distance": object_distance,
                "object_specifics": object_specifics,
                # Work history
                "job_description": job_description,
                "kind_of_injury": kind_of_injury,
                "hours_in_path": hours_in_path,
                "hours_this_week": hours_this_week,
                "glide_path": glide_path,
                "no_training_paperwork": no_training_paperwork,
                "labor_share": labor_share,
                "yes_training_paperwork": yes_training_paperwork,
                "previous_stand_down": previous_stand_down,
                "roster_reviewed": roster_reviewed,
                "countermeasure_updated": countermeasure_updated,
                "last_safety_observation": last_safety_observation,
              "observation_findings": observation_findings,
                "disciplinary_actions": disciplinary_actions,
                "incident_category": incident_category if isinstance(incident_category, list) else [incident_category],
                # 5 Whys
                "problem_statement": problem_statement,
                "material_why1": material_why1, "material_why2": material_why2, "material_why3": material_why3,
                "material_why4": material_why4, "material_why5": material_why5, "material_root_cause": material_root_cause,
                "machine_why1": machine_why1, "machine_why2": machine_why2, "machine_why3": machine_why3,
                "machine_why4": machine_why4, "machine_why5": machine_why5, "machine_root_cause": machine_root_cause,
                "method_why1": method_why1, "method_why2": method_why2, "method_why3": method_why3,
                "method_why4": method_why4, "method_why5": method_why5, "method_root_cause": method_root_cause,
                "human_why1": human_why1, "human_why2": human_why2, "human_why3": human_why3,
                "human_why4": human_why4, "human_why5": human_why5, "human_root_cause": human_root_cause,
                # Countermeasures
                "cm1_root_cause": cm1_root_cause, "cm1_countermeasure": cm1_countermeasure, "cm1_owner": cm1_owner,
                "cm1_pyramid": cm1_pyramid, "cm1_due_date": cm1_due_date, "cm1_cost": cm1_cost,
                "cm2_root_cause": cm2_root_cause, "cm2_countermeasure": cm2_countermeasure, "cm2_owner": cm2_owner,
                "cm2_pyramid": cm2_pyramid, "cm2_due_date": cm2_due_date, "cm2_cost": cm2_cost,
                "cm3_root_cause": cm3_root_cause, "cm3_countermeasure": cm3_countermeasure, "cm3_owner": cm3_owner,
                "cm3_pyramid": cm3_pyramid, "cm3_due_date": cm3_due_date, "cm3_cost": cm3_cost,
                "cm4_root_cause": cm4_root_cause, "cm4_countermeasure": cm4_countermeasure, "cm4_owner": cm4_owner,
                "cm4_pyramid": cm4_pyramid, "cm4_due_date": cm4_due_date, "cm4_cost": cm4_cost,
            })
            incident_file.write_text(json.dumps(saved_data, indent=2))
    
    return templates.TemplateResponse(
        request=request,
        name="success.html",
        context={
            "filename": output_filename,
            "xlsx_filename": None,
            "site_location": site_location,
        }
    )


@app.get("/download/{filename}")
async def download_file(filename: str):
    """Download generated file."""
    file_path = OUTPUTS_DIR / filename
    if file_path.exists():
        if filename.endswith(".xlsx"):
            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else:
            media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        return FileResponse(path=file_path, filename=filename, media_type=media_type)
    return JSONResponse({"error": "File not found"}, status_code=404)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
