"""Generate RCA PowerPoint from scratch — no template file required."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

from models import RCAIncident, PYRAMID_LABELS

# ── Walmart brand colours ─────────────────────────────────────────────────────
WMT_BLUE  = RGBColor(0x00, 0x53, 0xE2)
WMT_SPARK = RGBColor(0xFF, 0xC2, 0x20)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK      = RGBColor(0x1A, 0x1A, 0x1A)
LGREY     = RGBColor(0xF2, 0xF2, 0xF2)
MGREY     = RGBColor(0xD9, 0xD9, 0xD9)

# ── Slide canvas (16:9 widescreen) ────────────────────────────────────────────
SW = Inches(13.33)
SH = Inches(7.50)

# ── Low-level helpers ─────────────────────────────────────────────────────────

def _blank_slide(prs: Presentation):
    """Add a completely blank slide."""
    blank_layout = prs.slide_layouts[6]
    return prs.slides.add_slide(blank_layout)


def _rgb_str(rgb: RGBColor) -> str:
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def _solid_fill(shape, rgb: RGBColor):
    """Apply a solid fill colour to any shape."""
    fill = shape.fill
    fill.solid()
    fill.fore_color.rgb = rgb


def _add_rect(slide, left, top, width, height, rgb: RGBColor):
    """Add a filled rectangle (no outline)."""
    shape = slide.shapes.add_shape(1, left, top, width, height)  # 1 = MSO_SHAPE_TYPE.RECTANGLE
    _solid_fill(shape, rgb)
    shape.line.fill.background()
    return shape


def _tf(shape,
        text: str,
        bold: bool = False,
        size: int = 12,
        color: RGBColor = DARK,
        align=PP_ALIGN.LEFT,
        wrap: bool = True):
    """Set text-frame content with basic formatting."""
    tf = shape.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text) if text else ""
    run.font.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = color
    return tf


def _textbox(slide, left, top, width, height,
             text: str = "",
             bold: bool = False,
             size: int = 11,
             color: RGBColor = DARK,
             align=PP_ALIGN.LEFT,
             bg: Optional[RGBColor] = None):
    """Add a textbox with optional background fill."""
    box = slide.shapes.add_textbox(left, top, width, height)
    if bg:
        _solid_fill(box, bg)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text) if text else ""
    run.font.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = color
    return box


def _header(slide, title: str, subtitle: str = ""):
    """Blue banner across top with title + optional subtitle."""
    _add_rect(slide, 0, 0, SW, Inches(0.72), WMT_BLUE)
    _textbox(slide, Inches(0.15), Inches(0.06), Inches(9.5), Inches(0.36),
             text=title, bold=True, size=18, color=WHITE)
    if subtitle:
        _textbox(slide, Inches(0.15), Inches(0.42), Inches(11), Inches(0.28),
                 text=subtitle, bold=False, size=10, color=WMT_SPARK)


def _label(slide, left, top, width, height, text: str):
    """Small dark-background label pill."""
    _add_rect(slide, left, top, width, height, MGREY)
    _textbox(slide, left + Inches(0.04), top, width, height,
             text=text, bold=True, size=8, color=DARK)


def _value(slide, left, top, width, height, text: str, size: int = 10):
    """Light-background value field."""
    _add_rect(slide, left, top, width, height, LGREY)
    _textbox(slide, left + Inches(0.05), top + Inches(0.01), width - Inches(0.1), height,
             text=text, size=size, color=DARK)


def _section_bar(slide, top, text: str):
    """Full-width dark-blue section divider."""
    _add_rect(slide, 0, top, SW, Inches(0.26), WMT_BLUE)
    _textbox(slide, Inches(0.1), top + Inches(0.03), SW - Inches(0.2), Inches(0.22),
             text=text, bold=True, size=10, color=WHITE)


def _kv(slide, left, top, w_label, w_val, height, label: str, value: str):
    """Inline key → value pair."""
    _label(slide, left, top, w_label, height, label)
    _value(slide, left + w_label, top, w_val, height, value)


def _fmt_dt(dt) -> str:
    if dt is None:
        return ""
    if hasattr(dt, "strftime"):
        return dt.strftime("%m/%d/%Y  %I:%M %p")
    return str(dt)


def _fmt_date(dt) -> str:
    if dt is None:
        return ""
    if hasattr(dt, "strftime"):
        return dt.strftime("%m/%d/%Y")
    return str(dt)


# ── Slide builders ────────────────────────────────────────────────────────────

def _slide_1(prs: Presentation, d: RCAIncident):
    """Slide 1 — Incident Overview."""
    s = _blank_slide(prs)
    _header(s, "RCA — INCIDENT OVERVIEW", d.site_location)

    ROW_H = Inches(0.30)
    LBL_W = Inches(2.00)
    VAL_W = Inches(4.10)
    COL2  = Inches(6.60)
    LBL2  = Inches(1.80)
    VAL2  = Inches(4.73)
    top   = Inches(0.80)
    gap   = Inches(0.32)

    rows_left = [
        ("Incident Date/Time",  _fmt_dt(d.incident_datetime)),
        ("Reported Date/Time",  _fmt_dt(d.reported_datetime)),
        ("Associate's Manager", d.associates_manager),
        ("Reporting Manager",   d.reporting_manager),
        ("Incident Location",   d.incident_location),
        ("Job Description",     d.job_description),
        ("Kind of Injury",      d.kind_of_injury),
    ]
    for label, val in rows_left:
        _kv(s, Inches(0.15), top, LBL_W, VAL_W, ROW_H, label, val)
        top += gap

    # Normal duties / shift / overtime block
    top += Inches(0.05)
    _section_bar(s, top, "Work Status at Time of Incident")
    top += Inches(0.28)
    for label, val in [
        ("Normal Duties?", d.doing_normal_duties),
        ("Normal Shift?",  d.on_normal_shift),
        ("On Overtime?",   d.on_overtime),
    ]:
        _kv(s, Inches(0.15), top, LBL_W, VAL_W, ROW_H, label, val)
        top += gap

    # Right column — Incident Types
    _section_bar(s, Inches(0.80), "Incident Type(s)")
    t2 = Inches(1.10)
    for itype in [
        "Non-Med",
        "Med Only / Nurse Triage",
        "OSHA Recordable / Telemed",
        "Lost Time Injury",
        "Trailer Pullout",
        "PIT on PIT /PIT on Structure",
        "PIT on Pedestrian / Near Miss",
    ]:
        checked = "☑" if (d.incident_types and itype in d.incident_types) else "☐"
        _textbox(s, COL2, t2, LBL2 + VAL2, Inches(0.27),
                 text=f"  {checked}  {itype}", size=10, color=DARK, bg=LGREY)
        t2 += Inches(0.29)


def _slide_2(prs: Presentation, d: RCAIncident):
    """Slide 2 — Associate & Incident Details."""
    s = _blank_slide(prs)
    _header(s, "ASSOCIATE & INCIDENT DETAILS", d.site_location)

    ROW_H = Inches(0.28)
    LBL_W = Inches(3.20)
    VAL_W = Inches(2.90)
    top   = Inches(0.80)
    gap   = Inches(0.30)

    work_rows = [
        ("Hours in Same Path",          d.hours_in_path),
        ("Hours Worked This Week",       d.hours_this_week),
        ("Glide Path",                   d.glide_path),
        ("No Training Paperwork",        d.no_training_paperwork),
        ("Labor Share",                  d.labor_share),
        ("Yes Training Paperwork",       d.yes_training_paperwork),
        ("Previous Stand Down",          d.previous_stand_down),
        ("Roster Reviewed",              d.roster_reviewed),
        ("Countermeasure Updated",       d.countermeasure_updated),
        ("Last Safety Observation",      d.last_safety_observation),
        ("Observation Findings",         d.observation_findings),
        ("Disciplinary Actions",         d.disciplinary_actions),
    ]
    _section_bar(s, top - Inches(0.02), "Associate Work History")
    top += Inches(0.27)
    for label, val in work_rows:
        _kv(s, Inches(0.15), top, LBL_W, VAL_W, ROW_H, label, val)
        top += gap

    # Right column — Incident Category + Object Details
    C2    = Inches(6.60)
    C2_LW = Inches(2.00)
    C2_VW = Inches(4.50)
    t2 = Inches(0.80)

    _section_bar(s, t2 - Inches(0.02), "Incident Category")
    t2 += Inches(0.27)
    for cat in ["Material Handling", "Struck By/Against", "Slip/Trip/Fall"]:
        checked = "☑" if (d.incident_category and cat in d.incident_category) else "☐"
        _textbox(s, C2, t2, C2_LW + C2_VW, Inches(0.26),
                 text=f"  {checked}  {cat}", size=10, color=DARK, bg=LGREY)
        t2 += Inches(0.28)

    t2 += Inches(0.10)
    _section_bar(s, t2, "Object / Equipment Involved")
    t2 += Inches(0.28)
    obj = d.object_details
    for label, val in [
        ("Specifics",        obj.specifics),
        ("Item Name",        obj.item_name),
        ("Size",             obj.size),
        ("Shape",            obj.shape),
        ("Weight",           obj.weight),
        ("Distance / Reach", obj.distance_reach),
    ]:
        _kv(s, C2, t2, C2_LW, C2_VW, ROW_H, label, val)
        t2 += Inches(0.30)


def _slide_3(prs: Presentation, d: RCAIncident):
    """Slide 3 — Incident Description."""
    s = _blank_slide(prs)
    _header(s, "INCIDENT DESCRIPTION", d.site_location)

    _section_bar(s, Inches(0.80), "Injury / Illness Description")
    _add_rect(s, Inches(0.15), Inches(1.10), SW - Inches(0.30), Inches(2.40), LGREY)
    _textbox(s, Inches(0.20), Inches(1.13), SW - Inches(0.40), Inches(2.34),
             text=d.injury_description, size=11, color=DARK)

    _section_bar(s, Inches(3.60), "Manager's Account of Incident")
    _add_rect(s, Inches(0.15), Inches(3.90), SW - Inches(0.30), Inches(3.40), LGREY)
    account = d.manager_incident_account or d.how_incident_occurred
    _textbox(s, Inches(0.20), Inches(3.93), SW - Inches(0.40), Inches(3.34),
             text=account, size=11, color=DARK)


def _slide_photos(prs: Presentation, title: str, image_paths: list[Path]):
    """Generic photo grid slide (2×2)."""
    s = _blank_slide(prs)
    _header(s, title)

    positions = [
        (Inches(0.20),  Inches(0.80), Inches(6.40), Inches(3.10)),
        (Inches(6.73),  Inches(0.80), Inches(6.40), Inches(3.10)),
        (Inches(0.20),  Inches(4.05), Inches(6.40), Inches(3.10)),
        (Inches(6.73),  Inches(4.05), Inches(6.40), Inches(3.10)),
    ]

    for i, img_path in enumerate(image_paths[:4]):
        if not Path(img_path).exists():
            continue
        left, top, max_w, max_h = positions[i]
        try:
            from PIL import Image as PILImage
            with PILImage.open(img_path) as img:
                iw, ih = img.size
                ratio = iw / ih
                w = max_w.inches
                h = w / ratio
                if h > max_h.inches:
                    h = max_h.inches
                    w = h * ratio
                cx = left + Inches((max_w.inches - w) / 2)
                cy = top  + Inches((max_h.inches - h) / 2)
                s.shapes.add_picture(str(img_path), cx, cy, Inches(w), Inches(h))
        except Exception:
            s.shapes.add_picture(str(img_path), left, top, width=max_w)


def _slide_6(prs: Presentation, d: RCAIncident):
    """Slide 6 — 5 Whys Root Cause Analysis."""
    s = _blank_slide(prs)
    _header(s, "ROOT CAUSE ANALYSIS — 5 WHYS (4Ms)", d.site_location)

    fw = d.five_whys
    _section_bar(s, Inches(0.80), "Problem Statement")
    _add_rect(s, Inches(0.15), Inches(1.10), SW - Inches(0.30), Inches(0.55), LGREY)
    _textbox(s, Inches(0.20), Inches(1.12), SW - Inches(0.40), Inches(0.51),
             text=fw.problem_statement, size=11, color=DARK)

    # 5-Whys table
    tbl_top  = Inches(1.75)
    tbl_left = Inches(0.15)
    tbl_w    = SW - Inches(0.30)
    tbl_h    = SH - tbl_top - Inches(0.20)

    cols = 7
    rows = 5
    tbl = s.shapes.add_table(rows, cols, tbl_left, tbl_top, tbl_w, tbl_h).table

    # Header row
    headers = ["4M Category", "Why 1", "Why 2", "Why 3", "Why 4", "Why 5", "Root Cause"]
    for ci, hdr in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = hdr
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0]
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = WHITE
        cell.fill.solid()
        cell.fill.fore_color.rgb = WMT_BLUE

    # Data rows: (label, FiveWhysCategory)
    cat_rows = [
        ("Material",  fw.material),
        ("Machine",   fw.machine),
        ("Method",    fw.method),
        ("huMan",     fw.human),
    ]
    for ri, (cat_name, cat) in enumerate(cat_rows):
        row_idx = ri + 1
        vals = [cat_name, cat.why1, cat.why2, cat.why3, cat.why4, cat.why5, cat.root_cause]
        for ci, val in enumerate(vals):
            cell = tbl.cell(row_idx, ci)
            cell.text = val or ""
            p = cell.text_frame.paragraphs[0]
            run = p.runs[0] if p.runs else p.add_run()
            run.font.size = Pt(9)
            run.font.bold = (ci == 0)
            run.font.color.rgb = DARK
            bg = LGREY if ri % 2 == 0 else WHITE
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg if ci > 0 else MGREY


def _slide_7(prs: Presentation, d: RCAIncident):
    """Slide 7 — Countermeasures."""
    s = _blank_slide(prs)
    _header(s, "COUNTERMEASURES", d.site_location)

    tbl_top  = Inches(0.80)
    tbl_left = Inches(0.15)
    tbl_w    = SW - Inches(0.30)
    tbl_h    = SH - tbl_top - Inches(0.20)

    rows = 5  # header + 4 countermeasures
    cols = 6
    tbl = s.shapes.add_table(rows, cols, tbl_left, tbl_top, tbl_w, tbl_h).table

    headers = ["Root Cause", "Countermeasure", "Owner", "Due Date", "Cost", "Pyramid Level"]
    for ci, hdr in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = hdr
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.runs[0]
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = WHITE
        cell.fill.solid()
        cell.fill.fore_color.rgb = WMT_BLUE

    for ri, cm in enumerate(d.countermeasures[:4]):
        row_idx = ri + 1
        pyramid_label = PYRAMID_LABELS.get(cm.pyramid_level, str(cm.pyramid_level))
        vals = [
            cm.root_cause,
            cm.countermeasure,
            cm.owner,
            _fmt_date(cm.due_date),
            cm.cost,
            pyramid_label if (cm.root_cause or cm.countermeasure) else "",
        ]
        bg = LGREY if ri % 2 == 0 else WHITE
        for ci, val in enumerate(vals):
            cell = tbl.cell(row_idx, ci)
            cell.text = val or ""
            p = cell.text_frame.paragraphs[0]
            run = p.runs[0] if p.runs else p.add_run()
            run.font.size = Pt(9)
            run.font.color.rgb = DARK
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg

    # Spark-yellow footer note
    _textbox(s, Inches(0.15), SH - Inches(0.25), SW - Inches(0.30), Inches(0.22),
             text=f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  •  Incident Report Builder",
             size=8, color=WMT_SPARK)


# ── Public API ────────────────────────────────────────────────────────────────

def generate_rca_pptx(
    data: RCAIncident,
    output_path: Path,
    reenactment_photos: list[Path] | None = None,
    cctv_screenshots: list[Path] | None = None,
    template_path: Path | None = None,   # kept for backward-compat; ignored
) -> Path:
    """Build a full RCA PowerPoint from *data* and save to *output_path*.

    Args:
        data:               Populated RCAIncident model.
        output_path:        Destination .pptx file path.
        reenactment_photos: Optional list of image paths for slide 4.
        cctv_screenshots:   Optional list of image paths for slide 5.
        template_path:      Ignored — kept so existing callers don't break.

    Returns:
        output_path after saving.
    """
    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH

    _slide_1(prs, data)
    _slide_2(prs, data)
    _slide_3(prs, data)

    if reenactment_photos:
        _slide_photos(prs, "REENACTMENT PHOTOS", reenactment_photos)

    if cctv_screenshots:
        _slide_photos(prs, "CCTV SCREENSHOTS", cctv_screenshots)

    _slide_6(prs, data)
    _slide_7(prs, data)

    prs.save(str(output_path))
    return output_path
