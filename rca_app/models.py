"""Pydantic models for RCA Incident Form."""
from datetime import date, datetime
from typing import Optional
from pydantic import BaseModel, Field


class FiveWhysCategory(BaseModel):
    """A single 4M category row for 5 Whys analysis."""
    why1: str = ""
    why2: str = ""
    why3: str = ""
    why4: str = ""
    why5: str = ""
    root_cause: str = ""


class FiveWhys(BaseModel):
    """Complete 5 Whys analysis using 4Ms framework."""
    problem_statement: str = ""
    material: FiveWhysCategory = Field(default_factory=FiveWhysCategory)
    machine: FiveWhysCategory = Field(default_factory=FiveWhysCategory)
    method: FiveWhysCategory = Field(default_factory=FiveWhysCategory)
    human: FiveWhysCategory = Field(default_factory=FiveWhysCategory)


class Countermeasure(BaseModel):
    """A single countermeasure action item."""
    root_cause: str = ""
    countermeasure: str = ""
    owner: str = ""
    due_date: Optional[date] = None
    cost: str = ""
    pyramid_level: int = Field(default=1, ge=1, le=6)


class ObjectDetails(BaseModel):
    """Details about the object involved in the incident."""
    specifics: str = ""  # Specifics of Object (detailed description)
    item_name: str = ""
    size: str = ""
    shape: str = ""
    weight: str = ""
    distance_reach: str = ""


class RCAIncident(BaseModel):
    """Complete RCA Incident form data."""
    
    # === SLIDE 1: Incident Overview ===
    site_location: str = ""
    incident_types: list[str] = Field(default_factory=list)  # Multi-select checkboxes
    incident_datetime: Optional[datetime] = None
    reported_datetime: Optional[datetime] = None
    medical_datetime: Optional[datetime] = None
    associates_manager: str = ""
    reporting_manager: str = ""
    
    # Preliminary Info
    job_description: str = ""
    doing_normal_duties: str = ""  # Y/N
    on_normal_shift: str = ""  # Y/N
    on_overtime: str = ""  # Y/N
    incident_location: str = ""
    kind_of_injury: str = ""
    
    # === SLIDE 2: Associate Details ===
    hours_in_path: str = ""
    hours_this_week: str = ""
    glide_path: str = ""
    no_training_paperwork: str = ""  # If no proper training
    labor_share: str = ""
    yes_training_paperwork: str = ""  # If yes proper training
    previous_stand_down: str = ""
    roster_reviewed: str = ""  # If yes to previous
    countermeasure_updated: str = ""  # If no
    last_safety_observation: str = ""
    observation_findings: str = ""
    disciplinary_actions: str = "N"
    
    # Incident Type Details (checkboxes)
    incident_category: list[str] = Field(default_factory=list)  # Material Handling, Struck By, Slip/Trip/Fall
    
    # Object Details
    object_details: ObjectDetails = Field(default_factory=ObjectDetails)
    
    # === SLIDE 3: Incident Description ===
    injury_description: str = ""
    how_incident_occurred: str = ""  # Associate's account (WMW-738)
    manager_incident_account: str = ""  # Manager's account (for RCA PowerPoint)
    
    # === SLIDE 6: Root Cause Analysis (5 Whys) ===
    five_whys: FiveWhys = Field(default_factory=FiveWhys)
    
    # === SLIDE 7: Countermeasures ===
    countermeasures: list[Countermeasure] = Field(
        default_factory=lambda: [Countermeasure() for _ in range(4)]
    )


# Pyramid level to days mapping for auto-calculating due dates
PYRAMID_DAYS = {
    1: 120,  # Eliminate
    2: 90,   # Automate
    3: 25,   # Error Proof
    4: 5,    # At a Glance
    5: 5,    # Verify
    6: 3,    # Remind
}

PYRAMID_LABELS = {
    1: "Eliminate (120 days)",
    2: "Automate (90 days)",
    3: "Error Proof (25 days)",
    4: "At a Glance (5 days)",
    5: "Verify (5 days)",
    6: "Remind (3 days)",
}

INCIDENT_TYPES = [
    "Non-Med",
    "Med Only / Nurse Triage",
    "OSHA Recordable / Telemed",
    "Lost Time Injury",
    "Trailer Pullout",
    "PIT on PIT /PIT on Structure",
    "PIT on Pedestrian / Near Miss",
]

INCIDENT_CATEGORIES = [
    "Material Handling",
    "Struck By/Against",
    "Slip/Trip/Fall",
]
