"""
PMI Predictive Project Management Toolkit (PMBOK(R) Guide - 8th Edition aligned)
================================================================================
SINGLE SOURCE OF TRUTH for project metadata, document register, versioning and
visual identity. Edit the PROJECT block below and re-run build_toolkit.py to
regenerate all 25 artefacts with identical metadata.
"""

# ---------------------------------------------------------------------------
# 1. PROJECT METADATA  (edit these values, then re-run)
# ---------------------------------------------------------------------------
PROJECT = {
    "org": "[Organization Name]",
    "name": "[Project Name]",
    "id": "PRJ-2026-001",
    "sponsor": "[Project Sponsor]",
    "pm": "[Project Manager]",
    "pmo": "[PMO / Portfolio]",
    "business_unit": "[Business Unit]",
    "start": "2026-11-02",          # ISO yyyy-mm-dd
    "finish": "2027-10-29",
    "currency": "CAD",
    "approach": "Predictive (plan-driven)",
    "classification": "Internal",
}

# Default control values stamped on every new artefact
DOC_DEFAULTS = {
    "version": "0.1",
    "status": "Draft",
    "date": "2026-10-27",
    "author": PROJECT["pm"],
}

# ---------------------------------------------------------------------------
# 2. DOCUMENT REGISTER  (number, title, code, format, owner, PMBOK 8 mapping)
#    Domains: GOV SCO SCH FIN STK RES RSK   Focus areas: I P E MC C
# ---------------------------------------------------------------------------
DOMAINS = ["Governance", "Scope", "Schedule", "Finance",
           "Stakeholders", "Resources", "Risk"]
FOCUS_AREAS = ["Initiating", "Planning", "Executing",
               "Monitoring & Controlling", "Closing"]

REGISTER = [
    # no, title, code, ext, primary domain(s), focus areas, approver
    ("01", "Project Charter", "CHR", "docx", ["Governance"], ["Initiating"], "Sponsor"),
    ("02", "Integrated Project Management Plan", "PMP", "docx", ["Governance"], ["Planning", "Monitoring & Controlling"], "Sponsor"),
    ("03", "Scope Management Plan", "SCP", "docx", ["Scope"], ["Planning"], "Sponsor"),
    ("04", "Requirements & RTM", "RTM", "xlsx", ["Scope"], ["Planning", "Monitoring & Controlling"], "Business Owner"),
    ("05", "WBS & WBS Dictionary", "WBS", "xlsx", ["Scope"], ["Planning"], "Project Manager"),
    ("06", "Schedule Management Plan", "SMP", "docx", ["Schedule"], ["Planning"], "Sponsor"),
    ("07", "Master Schedule", "MSC", "xlsx", ["Schedule"], ["Planning", "Monitoring & Controlling"], "Project Manager"),
    ("08", "Financial Management Plan", "FMP", "docx", ["Finance"], ["Planning"], "Sponsor"),
    ("09", "Project Budget", "BUD", "xlsx", ["Finance"], ["Planning", "Monitoring & Controlling"], "Sponsor"),
    ("10", "Stakeholder Register", "SHR", "xlsx", ["Stakeholders"], ["Initiating", "Planning"], "Project Manager"),
    ("11", "Stakeholder Engagement Plan", "SEP", "docx", ["Stakeholders"], ["Planning", "Executing"], "Sponsor"),
    ("12", "Communication Plan", "COM", "xlsx", ["Stakeholders"], ["Planning", "Executing"], "Project Manager"),
    ("13", "Resource Management Plan", "RMP", "docx", ["Resources"], ["Planning", "Executing"], "Sponsor"),
    ("14", "RACI Matrix", "RACI", "xlsx", ["Resources", "Governance"], ["Planning"], "Project Manager"),
    ("15", "Risk Management Plan", "RKP", "docx", ["Risk"], ["Planning"], "Sponsor"),
    ("16", "Risk Register + 5\u00d75 Matrix", "RKR", "xlsx", ["Risk"], ["Planning", "Executing", "Monitoring & Controlling"], "Project Manager"),
    ("17", "Quality Management Plan", "QMP", "docx", ["Scope", "Governance"], ["Planning", "Executing"], "Sponsor"),
    ("18", "Procurement Plan", "PRC", "docx", ["Resources", "Finance"], ["Planning", "Executing"], "Sponsor"),
    ("19", "RAID Log", "RAID", "xlsx", ["Risk", "Governance"], ["Executing", "Monitoring & Controlling"], "Project Manager"),
    ("20", "Change Request", "CR", "docx", ["Governance"], ["Monitoring & Controlling"], "Change Control Board"),
    ("21", "Change Log", "CHL", "xlsx", ["Governance"], ["Monitoring & Controlling"], "Project Manager"),
    ("22", "Project Status Report", "PSR", "pptx", ["Governance", "Stakeholders"], ["Monitoring & Controlling"], "Project Manager"),
    ("23", "Decision Log", "DCL", "xlsx", ["Governance"], ["Executing", "Monitoring & Controlling"], "Project Manager"),
    ("24", "Lessons Learned Register", "LLR", "xlsx", ["Governance"], ["Executing", "Closing"], "Project Manager"),
    ("25", "Project Closure Report", "PCR", "docx", ["Governance", "Finance"], ["Closing"], "Sponsor"),
]


def doc_id(code):
    return f"{PROJECT['id']}-{code}"


def entry(no):
    for r in REGISTER:
        if r[0] == no:
            return r
    raise KeyError(no)


# ---------------------------------------------------------------------------
# 3. VERSIONING SCHEME & CONTROLLED VOCABULARY
# ---------------------------------------------------------------------------
VERSIONING = [
    ("0.1 - 0.9", "Draft", "Work in progress. Increment the minor digit for each review cycle (0.1, 0.2 ...)."),
    ("1.0", "Approved / Baselined", "First version formally approved by the designated approver. Becomes the baseline."),
    ("1.1, 1.2 ...", "Minor revision", "Editorial or clarifying change with NO impact on scope, schedule or cost baselines. PM approval."),
    ("2.0, 3.0 ...", "Major revision", "Change that alters a baseline. Requires an approved Change Request (doc 20) logged in the Change Log (doc 21)."),
]
STATUSES = ["Draft", "In Review", "Approved", "Baselined", "Superseded", "Archived"]
CLASSIFICATIONS = ["Public", "Internal", "Confidential", "Restricted"]
RAG = ["Green", "Amber", "Red"]

GLOSSARY = [
    ("Baseline", "Approved version of scope, schedule or cost, changeable only through integrated change control."),
    ("BAC", "Budget at Completion - total planned value of the work (cost baseline)."),
    ("CCB", "Change Control Board - body authorized to approve or reject change requests."),
    ("Contingency Reserve", "Budget/time within the baseline for identified risks (known-unknowns); controlled by the PM."),
    ("Control Account", "Management control point where scope, budget, actual cost and schedule are integrated and compared to earned value."),
    ("CPI / SPI", "Cost / Schedule Performance Index (EV/AC and EV/PV). >1.00 is favourable."),
    ("Deliverable", "Any unique, verifiable product, result or capability required to complete the project."),
    ("EAC / ETC", "Estimate at Completion / Estimate to Complete."),
    ("EV / PV / AC", "Earned Value / Planned Value / Actual Cost."),
    ("Focus Area", "PMBOK 8 grouping of processes by timing: Initiating, Planning, Executing, Monitoring & Controlling, Closing."),
    ("Issue", "A current condition or situation that may affect project objectives (a risk that has materialized)."),
    ("Management Reserve", "Budget outside the cost baseline for unknown-unknowns; released only by the Sponsor."),
    ("Performance Domain", "PMBOK 8 area of practice: Governance, Scope, Schedule, Finance, Stakeholders, Resources, Risk."),
    ("Planning Package", "WBS component below a control account with known content but no detailed activities yet."),
    ("RACI", "Responsible, Accountable, Consulted, Informed. Exactly one 'A' per activity."),
    ("RAG", "Red / Amber / Green status indicator, thresholds defined in the Integrated PMP."),
    ("Risk", "An uncertain event or condition that, if it occurs, has a positive (opportunity) or negative (threat) effect on objectives."),
    ("RTM", "Requirements Traceability Matrix - links each requirement to its source, WBS, deliverable and test."),
    ("Tailoring", "Deliberate adaptation of approach, governance and processes to suit the project context."),
    ("WBS", "Work Breakdown Structure - hierarchical decomposition of the total scope of work."),
    ("Work Package", "Lowest level of the WBS for which cost and duration are estimated and managed."),
]

PRINCIPLES = ["Adopt a holistic view", "Focus on value", "Embed quality",
              "Lead accountably", "Integrate sustainability", "Build empowered teams"]

# ---------------------------------------------------------------------------
# 4. VISUAL IDENTITY
# ---------------------------------------------------------------------------
BRAND = {
    "navy": "1F3A5F",      # primary
    "teal": "13807E",      # accent
    "sky": "DCE8F2",       # light fill
    "mist": "F3F6F9",      # zebra / input fill
    "ink": "1E2A36",       # body text
    "grey": "6B7785",      # secondary text
    "guide": "2F6E9E",     # guidance text
    "line": "B8C4D0",      # borders
    "green": "2E8B57", "amber": "E0A100", "red": "C0392B",
    "green_l": "D5EFDF", "amber_l": "FCEFC7", "red_l": "F6D5D1",
    "font": "Calibri",
    "font_head": "Calibri",
}
TOOLKIT_NAME = "PMI Predictive Project Management Toolkit"
TOOLKIT_SUB = "Aligned to the PMBOK® Guide – Eighth Edition"
