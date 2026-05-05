"""Excel generator for WMW-738 Manager's Investigation form."""
from pathlib import Path
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font

# Template path
TEMPLATE_PATH = Path(__file__).parent.parent / "WMW-738   new template.xlsx"


def format_datetime(dt: datetime | None, fmt: str = "%m/%d/%Y") -> str:
    """Format datetime to string."""
    if dt is None:
        return ""
    return dt.strftime(fmt)


def format_time(dt: datetime | None) -> str:
    """Format datetime to time string."""
    if dt is None:
        return ""
    return dt.strftime("%I:%M %p")


# Default font size for filled-in values (adjust this to make text bigger/smaller)
DEFAULT_FONT_SIZE = 11


def safe_set_cell(ws, cell_ref: str, value: str, font_size: int = DEFAULT_FONT_SIZE):
    """
    Safely set a cell value, handling merged cells.
    For merged cells, we need to write to the top-left cell of the merge range.
    Also applies a readable font size to the cell.
    """
    cell_font = Font(size=font_size, name='Calibri')
    
    try:
        cell = ws[cell_ref]
        if hasattr(cell, 'value'):
            cell.value = value
            cell.font = cell_font
    except AttributeError:
        for merge_range in ws.merged_cells.ranges:
            if cell_ref in merge_range:
                top_left = merge_range.start_cell.coordinate
                ws[top_left].value = value
                ws[top_left].font = cell_font
                return


def generate_wmw738(data: dict) -> bytes:
    """
    Fill in the WMW-738 template with incident data.
    
    Args:
        data: Dictionary containing all Phase 1 form fields
    
    Returns the filled workbook as bytes.
    
    Cell mapping (based on merged regions):
    Row 2:  E2 = DC/FC#
    Row 5:  E5 = Date of Hire, L5 = Department
    Row 6:  L6 = Schedule
    Row 11: B11 = Incident Date, F11 = Incident Time
            I11 = Reported Date, K11 = Reported Time, N11 = Reported To
    Row 13: C13 = Last Name, K13 = First Name, P13 = MI
    Row 14: E14 = Classification, J14 = Normal Duties, M14 = If not normal
    Row 15: E15 = Location
    Row 16: E16 = Equipment Make, J16 = Model, N16 = Asset ID
    Row 20: A20 = How incident occurred
    Row 23: A23 = Injury description
    Row 26: A26 = Objects/equipment
    Row 29: A29 = Safety accountability history
    Row 34: A34 = Behavior/conditions
    Row 37: A37 = Prevention
    Row 40: E40 = Witness names
    """
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    
    # Helper to get value from dict
    def get(key, default=""):
        return data.get(key, default) or default
    
    # === ROW 2: Site ===
    safe_set_cell(ws, "E2", get("site_location"))
    
    # === ROW 5: Associate Info ===
    safe_set_cell(ws, "E5", get("assoc_hire_date"))
    safe_set_cell(ws, "L5", get("assoc_department"))
    
    # === ROW 6: Schedule ===
    safe_set_cell(ws, "L6", get("assoc_schedule"))
    
    # === ROW 11: Dates and Times ===
    # Incident date/time
    incident_dt = get("incident_datetime")
    if incident_dt:
        try:
            dt = datetime.fromisoformat(incident_dt)
            safe_set_cell(ws, "B11", dt.strftime("%m/%d/%Y"))
            safe_set_cell(ws, "F11", dt.strftime("%I:%M %p"))
        except:
            pass
    
    # Reported date/time
    reported_dt = get("reported_datetime")
    if reported_dt:
        try:
            dt = datetime.fromisoformat(reported_dt)
            safe_set_cell(ws, "I11", dt.strftime("%m/%d/%Y"))
            safe_set_cell(ws, "K11", dt.strftime("%I:%M %p"))
        except:
            pass
    
    # Reported to
    safe_set_cell(ws, "N11", get("reporting_manager"))
    
    # === ROW 13: Name ===
    safe_set_cell(ws, "C13", get("assoc_last_name"))
    safe_set_cell(ws, "K13", get("assoc_first_name"))
    safe_set_cell(ws, "P13", get("assoc_mi"))
    
    # === ROW 14: Classification ===
    incident_types = get("incident_types", [])
    if isinstance(incident_types, list):
        safe_set_cell(ws, "E14", ", ".join(incident_types))
    else:
        safe_set_cell(ws, "E14", str(incident_types))
    
    safe_set_cell(ws, "J14", get("doing_normal_duties"))
    safe_set_cell(ws, "M14", get("if_not_normal_duties"))
    
    # === ROW 15: Location ===
    safe_set_cell(ws, "E15", get("incident_location"))
    
    # === ROW 16: Power Equipment ===
    safe_set_cell(ws, "E16", get("equipment_make"))
    safe_set_cell(ws, "J16", get("equipment_model"))
    safe_set_cell(ws, "N16", get("equipment_asset_id"))
    
    # === ROW 20: How incident occurred ===
    safe_set_cell(ws, "A20", get("how_incident_occurred"))
    
    # === ROW 23: Injury description ===
    safe_set_cell(ws, "A23", get("injury_description"))
    
    # === ROW 26: Objects/equipment involved ===
    safe_set_cell(ws, "A26", get("objects_involved"))
    
    # === ROW 29: Safety accountability history ===
    safe_set_cell(ws, "A29", get("safety_accountability"))
    
    # === ROW 34: Behavior/conditions ===
    safe_set_cell(ws, "A34", get("behavior_conditions"))
    
    # === ROW 37: Prevention ===
    safe_set_cell(ws, "A37", get("prevention"))
    
    # === ROW 40: Witnesses ===
    safe_set_cell(ws, "E40", get("witness_names"))
    
    # Save to bytes
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()
