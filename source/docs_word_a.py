"""Word artefacts 01, 02, 03, 06, 08, 11 (and 00 Toolkit Guide)."""
from config import (PROJECT, REGISTER, DOMAINS, FOCUS_AREAS, VERSIONING, STATUSES,
                    CLASSIFICATIONS, GLOSSARY, PRINCIPLES, doc_id)
from docx_kit import Doc

CUR = PROJECT["currency"]


# ============================================================ 00 TOOLKIT GUIDE
def build_00(path):
    from docx_kit import Doc as _D
    import config
    # register a pseudo-entry for the guide
    if not any(r[0] == "00" for r in config.REGISTER):
        config.REGISTER.insert(0, ("00", "Toolkit Guide & Document Standards", "GDE", "docx",
                                   DOMAINS, FOCUS_AREAS, "PMO"))
    d = _D("00")
    d.cover("How to use this toolkit: the shared metadata, document-control standard, versioning scheme, "
            "terminology and PMBOK® 8 alignment that every artefact follows.")
    d.h("Purpose of the Toolkit")
    d.para("This toolkit provides a complete, consistent set of 25 predictive (plan-driven) project management "
           "artefacts. Each artefact shares the same project metadata, document-control table, versioning scheme, "
           "terminology and visual identity so that the set reads as one integrated project management plan.")
    d.para("The PMBOK® Guide – Eighth Edition (PMI, November 2025) is principle-based and non-prescriptive. "
           "It organizes practice into 6 principles, 7 performance domains and 5 focus areas with 40 processes. "
           "These templates are a tailored, predictive implementation of that guidance — they are not PMI "
           "publications and should be adapted to your organization and project size.")
    d.h("PMBOK 8 Structure Used in This Toolkit")
    d.table(["Element", "Items"], [
        ["6 Principles", "; ".join(PRINCIPLES)],
        ["7 Performance Domains", "; ".join(DOMAINS)],
        ["5 Focus Areas", "; ".join(FOCUS_AREAS)],
    ], widths=[1.8, 5.0], first_col_bold=True)
    d.h("Document Register & PMBOK 8 Mapping")
    d.guide("Use this register as the master index. Document IDs are generated as <Project ID>-<Code>. "
            "Mark the 'Used?' column during tailoring — small projects may merge subsidiary plans into the Integrated PMP.")
    rows = []
    for no, title, code, ext, doms, focus, appr in REGISTER:
        if no == "00":
            continue
        rows.append([no, title, doc_id(code), ext, ", ".join(doms), ", ".join(focus), appr, "Y / N"])
    d.table(["#", "Artefact", "Document ID", "Fmt", "Performance domain(s)", "Focus area(s)", "Approver", "Used?"],
            rows, widths=[0.3, 1.6, 1.25, 0.4, 1.2, 1.3, 0.9, 0.5], size=8)
    d.h("Domain Coverage Matrix")
    d.para("GOV = Governance, SCO = Scope, SCH = Schedule, FIN = Finance, STK = Stakeholders, RES = Resources, RSK = Risk.")
    hdr = ["#", "Artefact", "GOV", "SCO", "SCH", "FIN", "STK", "RES", "RSK"]
    rows = []
    for no, title, code, ext, doms, focus, appr in REGISTER:
        if no == "00":
            continue
        rows.append([no, title] + ["●" if dm in doms else "" for dm in DOMAINS])
    d.table(hdr, rows, widths=[0.3, 2.3] + [0.6] * 7, size=8)
    d.h("Shared Project Metadata")
    d.para("Every artefact carries the same metadata block on its cover (Word/PowerPoint) or Cover sheet (Excel):")
    d.kv_table([("Organization", PROJECT["org"]), ("Project name", PROJECT["name"]), ("Project ID", PROJECT["id"]),
                ("Project Sponsor", PROJECT["sponsor"]), ("Project Manager", PROJECT["pm"]),
                ("PMO / Portfolio", PROJECT["pmo"]), ("Business unit", PROJECT["business_unit"]),
                ("Planned start / finish", f"{PROJECT['start']} / {PROJECT['finish']}"),
                ("Currency", PROJECT["currency"]), ("Delivery approach", PROJECT["approach"]),
                ("Default classification", PROJECT["classification"])])
    d.guide("To change metadata across all 25 files at once, edit the PROJECT block in _source/config.py and run "
            "'python build_toolkit.py'. Alternatively use Find & Replace on the bracketed placeholders.")
    d.h("Document-Control Standard")
    d.bullets([
        ("Document control table", "Document ID, title, project, version, status, owner, approver, classification, location, next review."),
        ("Revision history", "One row per version: version, date, author, description, Change Request reference."),
        ("Review & approval", "Named approver per artefact (see register). Signature or e-approval record required for 1.0 and each major version."),
        ("Distribution", "Who receives the document and whether for action or information."),
        ("Header / footer", "Header shows project, title, Document ID and version; footer shows organization, classification and page x of y."),
    ])
    d.h("Versioning Scheme")
    d.table(["Version", "Meaning", "Rule"], [list(v) for v in VERSIONING], widths=[1.0, 1.5, 4.3])
    d.h("Controlled Vocabulary", level=2)
    d.kv_table([("Status values", " → ".join(STATUSES)),
                ("Classification", " / ".join(CLASSIFICATIONS)),
                ("RAG", "Green = on track within tolerance; Amber = at risk, tolerance threatened, PM action; "
                        "Red = tolerance breached, escalation to Sponsor/Steering Committee"),
                ("Priority", "Critical / High / Medium / Low"),
                ("File naming", "NN — Artefact Name_vX.Y.ext  (e.g. 01 — Project Charter_v1.0.docx)")])
    d.h("Visual Identity")
    d.table(["Element", "Specification"], [
        ["Primary colour", "Navy #1F3A5F — titles, table headers"],
        ["Accent colour", "Teal #13807E — sub-headings, rules"],
        ["Light fills", "Sky #DCE8F2 (labels), Mist #F3F6F9 (zebra rows / input cells)"],
        ["Status colours", "Green #2E8B57, Amber #E0A100, Red #C0392B"],
        ["Typeface", "Calibri — body 10.5 pt, H1 15 pt, H2 12 pt"],
        ["Guidance text", "Blue italic with teal left rule — delete before baselining"],
        ["Placeholders", "[Square brackets] in grey — replace with project content"],
        ["Excel conventions", "Navy headers = input columns; teal headers marked \u0192 = calculated (do not overwrite); amber cells = key assumptions; grey italic rows = examples"],
    ], widths=[1.8, 5.0], first_col_bold=True)
    d.h("How to Use the Toolkit (Recommended Sequence)")
    d.bullets([
        "Initiating: complete 01 Charter and 10 Stakeholder Register; obtain Sponsor approval.",
        "Planning: draft subsidiary plans 03, 06, 08, 11, 13, 15, 17, 18; build 04 RTM, 05 WBS, 07 Schedule, 09 Budget, "
        "12 Communication Plan, 14 RACI, 16 Risk Register; integrate everything in 02 Integrated PMP and baseline (v1.0).",
        "Executing / Monitoring & Controlling: maintain 19 RAID Log, 20/21 change control, 22 Status Report, 23 Decision Log, "
        "24 Lessons Learned; update registers at the cadence defined in the Integrated PMP.",
        "Closing: complete 25 Closure Report, finalize 24 Lessons Learned, archive all artefacts as 'Archived'.",
    ], style="List Number")
    d.h("Tailoring Guidance")
    d.table(["Project size", "Recommended set"], [
        ["Small / low complexity", "01, 02 (with scope, schedule, cost, risk, comms sections inside), 05, 07, 09, 10, 16/19, 20/21, 22, 25"],
        ["Medium", "All registers plus 03, 08, 15 and 11; other plans as sections of 02"],
        ["Large / high complexity or regulated", "Full set of 25, plus organization-specific artefacts (e.g. implementation, M&O transition, security plans)"],
    ], widths=[1.8, 5.0])
    d.h("Glossary of Terms")
    d.table(["Term", "Definition"], [list(g) for g in GLOSSARY], widths=[1.6, 5.2], first_col_bold=True)
    d.h("Recommended Free Reference Template Libraries")
    d.table(["Source", "What it offers", "Notes"], [
        ["California Project Management Framework (CA-PMF), California Dept. of Technology — capmf.cdt.ca.gov/Templates.html",
         "Complete Word/Excel set with instructions: charter, PMP, scope, schedule, cost, risk, issue, change, comms, governance, procurement, RTM, RACI, WBS, lessons learned, closeout.",
         "Free government source; downloadable as one zip. Strongest predictive coverage."],
        ["Smartsheet — smartsheet.com/content/project-management-plan-templates",
         "Excel/Word/Google formats: PM plan outline, integrated project plan, scope statement, WBS, stakeholder plan, cost plan, risk plan, resource plan, governance, Gantt.",
         "Free; good visual Excel layouts and dashboards."],
        ["PMI — pmi.org (Insights, PMBOK Guide, ProjectManagement.com)",
         "Authoritative guidance and articles; PMI members get PMBOK 8 PDF and member-only templates on ProjectManagement.com.",
         "Use for terminology and alignment; some templates require membership."],
        ["CDT SIMM 19 / SIMM 45 and CA-SDLC tools", "Project approval lifecycle and oversight reporting templates.", "Useful for public-sector governance."],
    ], widths=[2.3, 3.0, 1.5], size=8.5)
    d.save(path)


# ============================================================ 01 CHARTER
def build_01(path):
    d = Doc("01")
    d.cover("Formally authorizes the project, names the project manager and gives authority to apply "
            "organizational resources to project activities.")
    d.h("Project Purpose & Business Justification")
    d.guide("Describe the business need, problem or opportunity and why the project is being undertaken now. "
            "Reference the business case and benefits management plan if they exist.")
    d.ph("[Describe the problem/opportunity, current state and desired future state.]")
    d.h("Strategic Alignment & Value")
    d.guide("PMBOK 8 principle 'Focus on value': link the project to organizational strategy and expected benefits.")
    d.table(["Strategic objective", "Expected benefit / value", "Benefit owner", "Measure", "Realization timing"],
            [["[Objective]", "[Benefit]", "[Name]", "[KPI]", "[Quarter/Year]"], ["", "", "", "", ""]],
            widths=[1.6, 1.8, 1.1, 1.1, 1.1])
    d.h("Measurable Objectives & Success Criteria")
    d.table(["#", "Objective (SMART)", "Success criterion / metric", "Target", "How measured"],
            [["O1", "[Deliver ... by ...]", "[Metric]", "[Value]", "[Source]"],
             ["O2", "[Stay within approved budget]", "Cost variance at completion", "≤ ±5%", "Budget (doc 09)"],
             ["O3", "[Meet quality acceptance]", "Defects at acceptance", "[0 critical]", "QMP (doc 17)"]],
            widths=[0.4, 2.0, 1.6, 0.9, 1.4])
    d.h("High-Level Scope")
    d.table(["In scope", "Out of scope"], [["[Item]", "[Item]"], ["[Item]", "[Item]"], ["[Item]", "[Item]"]],
            widths=[3.4, 3.4])
    d.h("Key Deliverables", level=2, numbered=False)
    d.table(["#", "Deliverable", "Description", "Acceptance by"],
            [["D1", "[Deliverable]", "[Description]", "[Role]"], ["D2", "", "", ""]], widths=[0.4, 1.8, 3.2, 1.4])
    d.h("High-Level Requirements", level=2, numbered=False)
    d.bullets(["[Business requirement 1]", "[Regulatory / compliance requirement]", "[Technical requirement]"])
    d.h("Summary Milestone Schedule")
    d.table(["Milestone", "Target date", "Gate / approval"],
            [["Charter approved", "[yyyy-mm-dd]", "Gate 1 – Initiation"],
             ["Project management plan baselined", "[yyyy-mm-dd]", "Gate 2 – Planning"],
             ["[Key delivery milestone]", "[yyyy-mm-dd]", "[Gate]"],
             ["Final acceptance", "[yyyy-mm-dd]", "Gate 4 – Acceptance"],
             ["Project closed", PROJECT["finish"], "Gate 5 – Closure"]], widths=[3.0, 1.5, 2.3])
    d.h("Summary Budget & Funding")
    d.table(["Item", f"Amount ({CUR})", "Funding source / notes"],
            [["Estimated project cost (incl. contingency)", "[0.00]", "[Capital / Operating]"],
             ["Management reserve", "[0.00]", "Held by Sponsor"],
             ["Total authorized budget", "[0.00]", "Accuracy: ROM (-25% to +75%)"]], widths=[3.0, 1.4, 2.4])
    d.h("Key Stakeholders")
    d.table(["Name", "Role", "Organization", "Interest / expectation"],
            [[PROJECT["sponsor"], "Project Sponsor", "", ""], [PROJECT["pm"], "Project Manager", "", ""],
             ["[Name]", "[Business owner]", "", ""]], widths=[1.6, 1.5, 1.5, 2.2])
    d.guide("Full analysis is maintained in the Stakeholder Register (doc 10).")
    d.h("High-Level Risks, Assumptions & Constraints")
    d.table(["Type", "Description", "Owner"],
            [["Risk", "[If ... then ...]", ""], ["Assumption", "[Assumed true for planning]", ""],
             ["Constraint", "[Fixed date / budget cap / regulation]", ""], ["Dependency", "[External dependency]", ""]],
            widths=[1.1, 4.4, 1.3])
    d.h("Project Governance & PM Authority")
    d.table(["Area", "Project Manager authority", "Escalation to"],
            [["Staffing", "[Select team / request resources]", "Sponsor"],
             ["Budget", f"[Approve spend up to {CUR} X per item]", "Sponsor / Steering Committee"],
             ["Technical decisions", "[Within approved architecture]", "Design authority"],
             ["Conflict resolution", "[Within project team]", "Sponsor"],
             ["Schedule tolerance", "[± X working days]", "Steering Committee"]], widths=[1.5, 3.2, 2.1])
    d.h("Delivery Approach & Tailoring")
    d.para(f"Approach: {PROJECT['approach']}. Life cycle phases: Initiation → Planning → Execution "
           "→ Acceptance & Transition → Closure, each separated by a governance gate.")
    d.h("Sustainability Considerations")
    d.guide("PMBOK 8 principle 'Integrate sustainability': note environmental, social and economic impacts, "
            "and any ESG commitments the project must honour.")
    d.ph("[Environmental / social / economic considerations]")
    d.h("Project Exit Criteria")
    d.bullets(["All deliverables accepted by the designated approvers", "Operational handover signed",
               "Financial close-out completed", "Lessons learned captured and Closure Report approved"])
    d.approvals()
    d.save(path)


# ============================================================ 02 INTEGRATED PMP
def build_02(path):
    d = Doc("02")
    d.cover("Integrates all subsidiary plans and baselines and defines how the project will be executed, "
            "monitored and controlled, and closed.")
    d.h("Introduction")
    d.para("This Integrated Project Management Plan (PMP) is the authoritative source for how the project is managed. "
           "It consolidates the subsidiary plans and baselines listed in Section 6 and is maintained under change control.")
    d.h("Project Summary")
    d.kv_table([("Project name / ID", f"{PROJECT['name']} / {PROJECT['id']}"),
                ("Charter reference", doc_id("CHR")), ("Sponsor", PROJECT["sponsor"]),
                ("Project Manager", PROJECT["pm"]), ("Planned start / finish", f"{PROJECT['start']} / {PROJECT['finish']}"),
                ("Approved budget (BAC)", f"[{CUR} 0.00]"), ("Delivery approach", PROJECT["approach"])])
    d.h("Life Cycle, Phases & Governance Gates")
    d.table(["Phase", "Key outputs", "Gate", "Gate approver", "Exit criteria"],
            [["Initiation", "Charter, Stakeholder Register", "G1", "Sponsor", "Charter signed"],
             ["Planning", "Baselined PMP, WBS, schedule, budget, risk register", "G2", "Steering Committee", "PMP v1.0 approved"],
             ["Execution", "Deliverables, status reports", "G3", "Steering Committee", "Deliverables built & tested"],
             ["Acceptance & Transition", "Acceptance certificates, handover", "G4", "Sponsor / Business owner", "Formal acceptance"],
             ["Closure", "Closure Report, lessons learned, archive", "G5", "Sponsor", "Closure Report approved"]],
            widths=[1.2, 2.2, 0.5, 1.4, 1.5], size=8.5)
    d.h("Tailoring Decisions")
    d.guide("PMBOK 8 is non-prescriptive. Record what was adopted, adapted or omitted for each performance domain and why.")
    d.table(["Performance domain", "Tailoring decision", "Rationale"],
            [[dm, "[Adopted / adapted / combined into this plan]", "[Why]"] for dm in DOMAINS],
            widths=[1.6, 2.8, 2.4])
    d.h("Governance Structure")
    d.table(["Body / role", "Responsibilities", "Meets", "Decision rights"],
            [["Steering Committee", "Strategic direction, gate approvals, tolerance breaches", "Monthly", "Baseline changes > PM tolerance"],
             ["Project Sponsor", "Champions project, owns business case, releases management reserve", "Bi-weekly", "Scope/priority trade-offs"],
             ["Change Control Board", "Reviews and decides change requests", "As needed / bi-weekly", "Per thresholds in §8"],
             ["Project Manager", "Plans, leads and controls day-to-day delivery", "Daily", "Within delegated tolerances"],
             ["PMO", "Standards, assurance, portfolio reporting", "Monthly", "Advisory"]], widths=[1.4, 2.6, 1.1, 1.7], size=8.5)
    d.h("Subsidiary Plans, Baselines & Registers")
    rows = [[no, t, doc_id(c), ", ".join(dm)] for no, t, c, ext, dm, f, a in REGISTER if no not in ("00", "01", "02")]
    d.table(["#", "Component", "Document ID", "Performance domain"], rows, widths=[0.4, 2.8, 1.8, 1.8], size=8.5)
    d.h("Baselines", level=2, numbered=False)
    d.table(["Baseline", "Composition", "Source document", "Version"],
            [["Scope baseline", "Scope statement + WBS + WBS Dictionary", "03, 05", "1.0"],
             ["Schedule baseline", "Approved start/finish dates of activities & milestones", "07", "1.0"],
             ["Cost baseline", "Time-phased budget excluding management reserve", "09", "1.0"],
             ["Performance measurement baseline", "Integrated scope-schedule-cost baseline for EVM", "05, 07, 09", "1.0"]],
            widths=[1.8, 2.8, 1.2, 1.0])
    d.h("Integrated Change Control")
    d.table(["Step", "Activity", "Responsible", "Artefact"],
            [["1", "Raise change request", "Any stakeholder", "20 Change Request"],
             ["2", "Log and assign ID", "Project Manager", "21 Change Log"],
             ["3", "Impact analysis (scope, schedule, cost, quality, risk, resources)", "PM + SMEs", "20"],
             ["4", "Decide per authority thresholds", "PM / CCB / Steering Committee", "23 Decision Log"],
             ["5", "Update baselines & plans; re-version documents (major version)", "Project Manager", "Affected docs"],
             ["6", "Communicate decision and implement", "Project Manager", "12 Communication Plan"]],
            widths=[0.5, 3.2, 1.6, 1.5])
    d.table(["Change impact", "Approval authority"],
            [[f"Schedule ≤ [5] days and cost ≤ {CUR} [10,000], no scope change", "Project Manager"],
             [f"Schedule ≤ [20] days or cost ≤ {CUR} [50,000] or minor scope change", "Change Control Board"],
             ["Above CCB limits, milestone/gate date change, or change to objectives", "Steering Committee / Sponsor"]],
            widths=[4.6, 2.2])
    d.h("Configuration & Document Management")
    d.bullets(["Repository: [SharePoint / Confluence path]; single controlled copy per artefact.",
               "Versioning per toolkit scheme (0.x draft, 1.0 baseline, 1.x minor, 2.0+ major via approved CR).",
               "Configuration items: [list deliverables / components under configuration control]."])
    d.h("Performance Measurement & Reporting")
    d.table(["Indicator", "Green", "Amber", "Red"],
            [["Schedule (SPI / milestone slip)", "SPI ≥ 0.95 / no slip", "0.90–0.94 / ≤ 10 days", "< 0.90 / > 10 days"],
             ["Cost (CPI / variance)", "CPI ≥ 0.95", "0.90–0.94", "< 0.90"],
             ["Scope", "No unapproved change", "Pending CRs > 2 weeks", "Unapproved scope change"],
             ["Risk", "No critical risks open", "Critical risk with plan", "Critical risk without response"],
             ["Quality", "All criteria met", "Minor non-conformances", "Major non-conformance"]],
            widths=[2.0, 1.6, 1.6, 1.6])
    d.guide("Reporting cadence, audiences and channels are defined in the Communication Plan (doc 12); the Project "
            "Status Report (doc 22) applies these thresholds.")
    d.h("Issue Management & Escalation")
    d.table(["Level", "Escalate to", "Trigger", "Response time"],
            [["1", "Project Manager", "Issue within team tolerance", "2 business days"],
             ["2", "Project Sponsor", "Threatens tolerance or needs cross-functional resolution", "5 business days"],
             ["3", "Steering Committee", "Tolerance breached / strategic impact", "Next meeting or ad hoc"]],
            widths=[0.6, 1.8, 3.0, 1.4])
    d.h("Knowledge Management & Lessons Learned")
    d.para("Lessons are captured continuously in the Lessons Learned Register (doc 24) and reviewed at each gate.")
    d.h("Project Closure Approach")
    d.para("Closure follows the exit criteria in the Charter and is documented in the Project Closure Report (doc 25).")
    d.approvals([("Steering Committee Chair", "[Name]")])
    d.save(path)


# ============================================================ 03 SCOPE MGMT PLAN
def build_03(path):
    d = Doc("03")
    d.cover("Defines how project and product scope will be defined, decomposed, validated and controlled.")
    d.h("Purpose & Roles")
    d.table(["Role", "Scope management responsibility"],
            [["Project Sponsor", "Approves scope baseline and scope changes above PM tolerance"],
             ["Project Manager", "Owns this plan, WBS and scope baseline; facilitates change control"],
             ["Business Owner / Product Owner", "Defines and prioritizes requirements; accepts deliverables"],
             ["Business Analyst", "Elicits, documents and traces requirements (doc 04)"],
             ["Team leads", "Decompose work packages; estimate; confirm completion criteria"]], widths=[2.0, 4.8])
    d.h("Requirements Management Approach")
    d.bullets([("Elicitation", "Interviews, workshops, document analysis, observation, prototypes, surveys."),
               ("Analysis & prioritization", "MoSCoW (Must / Should / Could / Won't) agreed with the Business Owner."),
               ("Documentation", "Requirements & RTM workbook (doc 04): every requirement uniquely identified and traced."),
               ("Traceability", "Source → objective → WBS → deliverable → test case → acceptance."),
               ("Change", "New/changed requirements after baseline require a Change Request (doc 20).")])
    d.h("Project Scope Statement")
    d.kv_table([("Product scope description", "[Characteristics of the product, service or result]"),
                ("Project deliverables", "[List; see WBS level 2-3]"),
                ("Acceptance criteria", "[Conditions that must be met before deliverables are accepted]"),
                ("Exclusions (out of scope)", "[Explicitly excluded items]"),
                ("Constraints", "[Budget cap, fixed dates, regulations, technology]"),
                ("Assumptions", "[Factors assumed true]")])
    d.h("WBS Development Approach")
    d.bullets(["Deliverable-oriented decomposition; 100% rule — the WBS includes all and only the project scope.",
               "Numbering: 1, 1.1, 1.1.1 (max 4 levels). Level 1 = project; level 2 = phases/major deliverables.",
               "Work packages sized 8–80 hours of effort or ≤ one reporting period, with a single accountable owner.",
               "Each work package is mapped to a control account for EVM and costing (doc 09).",
               "Project management is a level-2 element so that management effort is planned and tracked."])
    d.h("WBS Dictionary", level=2, numbered=False)
    d.para("Maintained in the WBS & WBS Dictionary workbook (doc 05): description, deliverables, acceptance criteria, "
           "owner, milestones, effort, cost, assumptions and technical references per work package.")
    d.h("Scope Baseline")
    d.para("Scope baseline = approved Scope Statement + WBS + WBS Dictionary, version 1.0, approved at Gate 2.")
    d.h("Scope Validation (Deliverable Acceptance)")
    d.table(["Step", "Activity", "Responsible", "Record"],
            [["1", "Quality control confirms deliverable meets specs (doc 17)", "QA lead", "Test/inspection report"],
             ["2", "Deliverable submitted with acceptance criteria checklist", "Project Manager", "Acceptance form"],
             ["3", "Review/inspection by acceptor (≤ [10] business days)", "Business Owner", "Review notes"],
             ["4", "Accept, conditionally accept or reject", "Business Owner", "Signed acceptance"],
             ["5", "Update RTM status and Decision Log", "BA / PM", "Docs 04, 23"]], widths=[0.5, 3.2, 1.4, 1.7])
    d.h("Scope Control")
    d.bullets(["Scope variance is reviewed weekly against the baseline; any unapproved work is stopped.",
               "Requests follow integrated change control (doc 02 §8); thresholds as defined there.",
               "Scope creep indicators: work without WBS ID, requirements without RTM source, gold-plating."])
    d.approvals()
    d.save(path)


# ============================================================ 06 SCHEDULE MGMT PLAN
def build_06(path):
    d = Doc("06")
    d.cover("Establishes the criteria and activities for developing, monitoring and controlling the project schedule.")
    d.h("Scheduling Methodology & Tool")
    d.kv_table([("Methodology", "Critical Path Method (CPM) with resource levelling as needed"),
                ("Tool", "[MS Project / Primavera P6 / Smartsheet]; summary maintained in Master Schedule (doc 07)"),
                ("Calendar", "[5-day week, 8 h/day, statutory holidays per province/state]"),
                ("Units of measure", "Duration in working days; effort in hours"),
                ("Level of accuracy", "Activity durations rounded to 0.5 day; ROM ±50% at initiation, ±10% at baseline")])
    d.h("Schedule Development")
    d.table(["Step", "Description", "Output"],
            [["Define activities", "Decompose work packages (doc 05) into activities", "Activity list"],
             ["Sequence", "Logical dependencies (FS, SS, FF, SF), leads/lags; no open ends", "Network diagram"],
             ["Estimate durations", "Analogous, parametric, three-point (PERT: (O+4M+P)/6), expert judgement", "Duration estimates"],
             ["Develop schedule", "CPM, resource levelling, schedule compression (crashing/fast-tracking)", "Schedule model"],
             ["Baseline", "Approved at Gate 2 as part of PMP v1.0", "Schedule baseline"]], widths=[1.5, 3.6, 1.7])
    d.h("Schedule Quality Rules")
    d.bullets(["Every activity has at least one predecessor and successor (except start/finish milestones).",
               "No activity longer than [20] working days without decomposition.",
               "Hard constraints are minimized and justified; negative float is reported immediately.",
               "Milestones have zero duration and are linked to gates and deliverables."])
    d.h("Progress Measurement Rules")
    d.table(["Rule", "When used", "Credit"],
            [["0 / 100", "Activities ≤ 1 reporting period", "100% at finish"],
             ["50 / 50", "Activities spanning 2 periods", "50% at start, 50% at finish"],
             ["Weighted milestones", "Long work packages", "Per agreed milestone weights"],
             ["Physical % complete", "Measurable units (e.g. units installed)", "Units done / total"]],
            widths=[1.6, 3.0, 2.2])
    d.h("Monitoring & Control")
    d.table(["Item", "Approach"],
            [["Status date / update cycle", "[Weekly every Friday]; owners report actual start/finish and remaining duration"],
             ["Critical path & float", "Reviewed weekly; activities with float ≤ [5] days flagged as near-critical"],
             ["Variance thresholds", "Milestone slip > [5] days = Amber; > [10] days or SPI < 0.90 = Red"],
             ["Corrective actions", "Re-sequencing, resource reallocation, fast-tracking, crashing, scope trade-offs"],
             ["Re-baselining", "Only via approved Change Request; new major version of docs 07 and 02"]],
            widths=[1.9, 4.9])
    d.h("Roles & Responsibilities")
    d.table(["Role", "Responsibility"],
            [["Project Manager", "Owns schedule model, baseline, critical path analysis, reporting"],
             ["Scheduler / PMO", "Maintains tool, quality checks, schedule health metrics"],
             ["Work package owners", "Provide estimates and weekly progress"],
             ["Sponsor / Steering Committee", "Approve baseline and changes beyond PM tolerance"]], widths=[2.0, 4.8])
    d.h("Schedule Reporting")
    d.para("Milestone tracker, Gantt view and SPI are reported in the Project Status Report (doc 22) per the "
           "Communication Plan (doc 12).")
    d.approvals()
    d.save(path)


# ============================================================ 08 FINANCIAL MGMT PLAN
def build_08(path):
    d = Doc("08")
    d.cover("Defines how project costs and funding are estimated, budgeted, managed, controlled and reported "
            "(PMBOK 8 Finance performance domain).")
    d.h("Purpose & Financial Parameters")
    d.kv_table([("Currency / units", f"{CUR}; estimates rounded to nearest [100]"),
                ("Fiscal year", "[e.g. April – March]"),
                ("Capital / operating split", "[Capitalization policy reference]"),
                ("Chart of accounts / cost centre", "[Codes]"),
                ("Budget workbook", f"{doc_id('BUD')} (doc 09)")])
    d.h("Cost Estimating")
    d.table(["Estimate class", "Phase", "Accuracy range", "Methods"],
            [["Rough order of magnitude (ROM)", "Initiation", "-25% to +75%", "Analogous, expert judgement"],
             ["Budget estimate", "Early planning", "-10% to +25%", "Parametric, vendor quotes"],
             ["Definitive estimate", "Baseline", "-5% to +10%", "Bottom-up by work package"]], widths=[2.0, 1.3, 1.4, 2.1])
    d.h("Cost Breakdown Structure")
    d.para("Costs are estimated bottom-up per WBS work package and rolled up to control accounts. Categories: Labour, "
           "Contractors/Services, Hardware/Equipment, Software/Licences, Materials, Travel, Training, Facilities, Other.")
    d.h("Reserves")
    d.table(["Reserve", "Basis", "Held by", "Release"],
            [["Contingency reserve", "Quantified risks from Risk Register (doc 16) — [X]% of base estimate or EMV", "Project Manager", "Per risk response; logged"],
             ["Management reserve", "Unknown-unknowns — [Y]% of cost baseline", "Sponsor", "Via approved Change Request"]],
            widths=[1.5, 2.9, 1.2, 1.2])
    d.h("Cost Baseline & Funding")
    d.bullets(["Cost baseline = time-phased budget (S-curve) incl. contingency, excl. management reserve.",
               "Total project budget = cost baseline + management reserve.",
               "Funding requirements are reconciled with funding limits per fiscal period; gaps escalated to Sponsor."])
    d.h("Earned Value Management")
    d.table(["Metric", "Formula", "Interpretation"],
            [["Cost Variance (CV)", "EV − AC", "> 0 under budget"], ["Schedule Variance (SV)", "EV − PV", "> 0 ahead of schedule"],
             ["CPI", "EV / AC", "≥ 1.00 favourable"], ["SPI", "EV / PV", "≥ 1.00 favourable"],
             ["EAC (typical)", "BAC / CPI", "Forecast final cost"], ["ETC", "EAC − AC", "Cost to finish"],
             ["VAC", "BAC − EAC", "Forecast over/under run"], ["TCPI", "(BAC − EV) / (BAC − AC)", "Efficiency needed to meet BAC"]],
            widths=[1.8, 2.2, 2.8])
    d.h("Control Thresholds & Spending Authority")
    d.table(["Authority", f"Single commitment limit ({CUR})", "Cumulative variance limit"],
            [["Project Manager", "[10,000]", "±5% of control account"],
             ["Project Sponsor", "[100,000]", "±10% of BAC"],
             ["Steering Committee", "> [100,000]", "> 10% of BAC or use of management reserve"]], widths=[2.0, 2.4, 2.4])
    d.h("Financial Reporting & Forecasting")
    d.bullets(["Monthly: actuals from finance system reconciled to budget workbook by [day 5].",
               "Monthly: EVM metrics, EAC and cash-flow forecast reported in Status Report (doc 22).",
               "Quarterly: benefits and business-case review with Sponsor."])
    d.h("Value & Benefits Tracking")
    d.guide("PMBOK 8 emphasizes value delivery: record the financial measures used to confirm the investment remains justified.")
    d.table(["Measure", "Business case value", "Current forecast", "Comment"],
            [["NPV", "[ ]", "[ ]", ""], ["ROI / payback", "[ ]", "[ ]", ""], ["Annual benefit", "[ ]", "[ ]", ""]],
            widths=[1.6, 1.7, 1.7, 1.8])
    d.h("Financial Closure")
    d.para("At closure: confirm all invoices paid, accruals cleared, purchase orders closed, assets capitalized and "
           "final cost reported in the Closure Report (doc 25).")
    d.approvals([("Finance Business Partner", "[Name]")])
    d.save(path)


# ============================================================ 11 STAKEHOLDER ENGAGEMENT PLAN
def build_11(path):
    d = Doc("11")
    d.cover("Defines strategies and actions to promote productive stakeholder involvement in decisions and execution.")
    d.h("Purpose & Scope")
    d.para("This plan uses the analysis in the Stakeholder Register (doc 10) to define how each stakeholder group "
           "will be engaged. Communication logistics are detailed in the Communication Plan (doc 12).")
    d.h("Stakeholder Identification & Analysis")
    d.bullets(["Identification sources: charter, org charts, contracts, workshops, lessons learned.",
               "Analysis: power/influence (1-5), interest (1-5), attitude, current vs desired engagement.",
               "Register reviewed at each gate and whenever organizational change occurs."])
    d.h("Power / Interest Strategies")
    d.table(["Quadrant", "Profile", "Strategy", "Typical activities"],
            [["Manage closely", "High power, high interest", "Partner & co-decide", "Steering Committee, 1:1s, workshops"],
             ["Keep satisfied", "High power, low interest", "Consult on key decisions", "Executive briefings, gate reviews"],
             ["Keep informed", "Low power, high interest", "Inform & involve", "Newsletters, demos, user groups"],
             ["Monitor", "Low power, low interest", "Monitor with minimal effort", "Intranet updates"]],
            widths=[1.4, 1.6, 1.7, 2.1])
    d.h("Engagement Assessment Matrix")
    d.guide("C = current engagement, D = desired engagement. Close the gaps with targeted actions below.")
    d.table(["Stakeholder", "Unaware", "Resistant", "Neutral", "Supportive", "Leading"],
            [[PROJECT["sponsor"], "", "", "", "C", "D"], ["[Business owner]", "", "", "C", "D", ""],
             ["[End-user group]", "C", "", "", "D", ""], ["[Regulator]", "", "", "C/D", "", ""]],
            widths=[2.0, 0.95, 0.95, 0.95, 0.95, 0.95])
    d.h("Engagement Action Plan")
    d.table(["Stakeholder / group", "Gap", "Action", "Owner", "Timing", "Success indicator"],
            [["[Group]", "[Neutral → Supportive]", "[Action]", "[Name]", "[Date]", "[Indicator]"], ["", "", "", "", "", ""]],
            widths=[1.3, 1.1, 1.6, 0.9, 0.8, 1.1], size=8.5)
    d.h("Change Readiness & Resistance Management")
    d.bullets(["Identify impacted groups and degree of change (people, process, technology).",
               "Use change champions; address concerns early through two-way channels.",
               "Track resistance indicators: attendance, feedback sentiment, adoption metrics."])
    d.h("Monitoring Engagement")
    d.table(["Metric", "Target", "Frequency"],
            [["Key stakeholders at desired level", "≥ [80]%", "Monthly"],
             ["Meeting attendance (key forums)", "≥ [90]%", "Per meeting"],
             ["Stakeholder satisfaction pulse", "≥ [4/5]", "Quarterly"]], widths=[3.2, 1.6, 2.0])
    d.approvals()
    d.save(path)
