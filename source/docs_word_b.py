"""Word artefacts 13, 15, 17, 18, 20, 22 (docx), 25."""
from config import PROJECT, DOC_DEFAULTS, doc_id
from docx_kit import Doc, set_cell, shade, B
from docx.enum.text import WD_ALIGN_PARAGRAPH

CUR = PROJECT["currency"]


# ============================================================ 13 RESOURCE MGMT PLAN
def build_13(path):
    d = Doc("13")
    d.cover("Defines how team and physical resources are estimated, acquired, developed, managed and released.")
    d.h("Purpose & Approach")
    d.para("This plan covers human (team) and physical resources. Responsibility assignments are detailed in the "
           "RACI Matrix (doc 14); resource costs in the Project Budget (doc 09).")
    d.h("Project Organization")
    d.guide("Insert or link the project organization chart (Sponsor → Steering Committee → Project Manager "
            "→ workstream leads → team members).")
    d.ph("[Organization chart]")
    d.h("Roles, Responsibilities & Competencies")
    d.table(["Role", "Responsibilities", "Authority", "Required competencies", "Allocation"],
            [["Project Sponsor", "Business case, funding, escalations", "Approves baselines & reserve", "Executive leadership", "10%"],
             ["Project Manager", "Integrated planning, control, reporting", "Within delegated tolerances", "PMP, stakeholder mgmt", "100%"],
             ["Business Analyst", "Requirements, traceability", "Requirement sign-off prep", "Elicitation, modelling", "[ ]%"],
             ["Technical Lead", "Solution design, technical quality", "Technical decisions", "[Domain expertise]", "[ ]%"],
             ["QA Lead", "Quality assurance & testing", "Accept/reject test results", "Test management", "[ ]%"],
             ["[Role]", "", "", "", ""]], widths=[1.3, 1.9, 1.4, 1.4, 0.8], size=8.5)
    d.h("Resource Estimation & Acquisition")
    d.table(["Role / resource", "Qty", "Source (internal / contract)", "Start", "End", "Acquisition method"],
            [["[Role]", "1", "[Internal – dept]", "[date]", "[date]", "[Functional manager agreement]"],
             ["[Role]", "", "[Contractor]", "", "", "[Procurement – doc 18]"]],
            widths=[1.4, 0.4, 1.5, 0.8, 0.8, 1.9], size=8.5)
    d.h("Physical Resources")
    d.table(["Resource", "Specification", "Qty", "Needed from / to", "Source", "Owner"],
            [["[Equipment / facility / licence]", "", "", "", "", ""], ["", "", "", "", "", ""]],
            widths=[1.6, 1.5, 0.5, 1.2, 1.0, 1.0], size=8.5)
    d.h("Resource Calendars & Release")
    d.bullets(["Availability, holidays and part-time allocations are recorded in the scheduling tool.",
               "Onboarding checklist: access, induction, project orientation, safety briefing, tools.",
               "Release: knowledge transfer, performance feedback to functional manager, access revoked, assets returned."])
    d.h("Team Development")
    d.table(["Need", "Training / development action", "Who", "When"],
            [["[Skill gap]", "[Course / coaching / pairing]", "[Name]", "[Date]"]], widths=[1.8, 2.8, 1.2, 1.0])
    d.h("Team Charter & Ground Rules", level=2, numbered=False)
    d.guide("PMBOK 8 principle 'Build empowered teams': agree values, working agreements, decision-making and "
            "communication norms with the team.")
    d.bullets(["Core hours and response times: [ ]", "Decision-making: [consensus / consult-then-decide]",
               "Meeting etiquette: agenda, timebox, actions recorded", "Psychological safety and respectful challenge"])
    d.h("Recognition, Performance & Conflict")
    d.bullets([("Recognition", "[Milestone celebrations, peer recognition, input to performance reviews]"),
               ("Performance", "PM provides input to functional managers each [quarter]"),
               ("Conflict resolution", "Collaborate/problem-solve first; escalate to Sponsor if unresolved in [5] days")])
    d.h("Health, Safety & Compliance")
    d.ph("[Applicable HSE requirements, site access rules, labour agreements]")
    d.approvals()
    d.save(path)


# ============================================================ 15 RISK MGMT PLAN
def build_15(path):
    d = Doc("15")
    d.cover("Defines how risk management activities will be structured and performed, including scales, "
            "thresholds and the 5×5 probability-impact matrix used in the Risk Register (doc 16).")
    d.h("Methodology")
    d.table(["Process", "Approach", "Frequency"],
            [["Identify risks", "Workshops, checklists (RBS), assumption analysis, SWOT, expert interviews", "Planning + ongoing"],
             ["Qualitative analysis", "Probability × impact on 5×5 matrix; proximity; urgency", "On identification, monthly review"],
             ["Quantitative analysis", "EMV, three-point estimates, Monte Carlo (if project size warrants)", "Baseline & major changes"],
             ["Plan responses", "Strategy + owner + actions + due date + residual score", "Within [5] days"],
             ["Implement & monitor", "Track triggers, actions, new/closed risks; risk audits", "Weekly / monthly"]],
            widths=[1.6, 3.6, 1.6])
    d.h("Roles & Responsibilities")
    d.table(["Role", "Responsibility"],
            [["Project Sponsor", "Sets risk appetite; owns escalated risks; releases management reserve"],
             ["Project Manager", "Owns this plan and register; facilitates reviews; reports"],
             ["Risk owner", "Monitors trigger, implements responses, reports status"],
             ["Team members", "Identify and report risks continuously"]], widths=[1.8, 5.0])
    d.h("Risk Breakdown Structure (Categories)")
    d.table(["Category", "Examples"],
            [["Technical", "Requirements, technology, complexity, interfaces, performance, quality"],
             ["External", "Vendors, regulatory, market, weather, client"],
             ["Organizational", "Resources, funding, priorities, dependencies on other projects"],
             ["Project management", "Estimating, planning, control, communication"],
             ["Health, safety & environment", "Site safety, environmental compliance, sustainability"]], widths=[2.0, 4.8])
    d.h("Probability Scale")
    d.table(["Score", "Level", "Likelihood"],
            [["5", "Very high", "> 80%"], ["4", "High", "61% – 80%"], ["3", "Medium", "41% – 60%"],
             ["2", "Low", "21% – 40%"], ["1", "Very low", "≤ 20%"]], widths=[0.8, 1.6, 4.4])
    d.h("Impact Scale")
    d.table(["Score", "Cost", "Schedule", "Scope / quality", "Safety / reputation"],
            [["5 Very high", "> 20% BAC", "> 20% / misses gate", "Product unusable", "Serious injury / national media"],
             ["4 High", "10–20%", "10–20%", "Unacceptable to sponsor", "Lost-time injury / regulator action"],
             ["3 Medium", "5–10%", "5–10%", "Major areas affected", "Medical aid / local media"],
             ["2 Low", "1–5%", "1–5%", "Minor areas affected", "First aid / internal complaint"],
             ["1 Very low", "< 1%", "< 1%", "Barely noticeable", "Negligible"]], widths=[1.1, 1.1, 1.4, 1.6, 1.6], size=8.5)
    d.h("Probability × Impact Matrix (5×5)")
    t = d.d.add_table(rows=6, cols=6)
    from docx_kit import borders, cell_margins
    borders(t, "FFFFFF", 8); cell_margins(t, 80, 80, 80, 80)

    def band(s):
        if s >= 16: return B["red"], "FFFFFF"
        if s >= 10: return "E67E22", "FFFFFF"
        if s >= 5: return B["amber"], B["ink"]
        return B["green"], "FFFFFF"
    set_cell(t.rows[0].cells[0], "P \\ I", bold=True, color="FFFFFF"); shade(t.rows[0].cells[0], B["navy"])
    for i in range(1, 6):
        set_cell(t.rows[0].cells[i], str(i), bold=True, color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER); shade(t.rows[0].cells[i], B["navy"])
    for r in range(1, 6):
        p = 6 - r
        set_cell(t.rows[r].cells[0], str(p), bold=True, color="FFFFFF", align=WD_ALIGN_PARAGRAPH.CENTER); shade(t.rows[r].cells[0], B["navy"])
        for i in range(1, 6):
            s = p * i
            fill, fc = band(s)
            set_cell(t.rows[r].cells[i], str(s), bold=True, color=fc, align=WD_ALIGN_PARAGRAPH.CENTER); shade(t.rows[r].cells[i], fill)
    d._widths(t, [0.8] + [1.0] * 5)
    d.para("")
    d.table(["Score", "Rating", "Required action"],
            [["16 – 25", "Critical", "Escalate to Sponsor/Steering Committee within 24 h; response plan mandatory; weekly review"],
             ["10 – 15", "High", "Response plan within 5 days; PM reviews weekly; report in status report"],
             ["5 – 9", "Medium", "Response plan; monitor monthly"],
             ["1 – 4", "Low", "Accept and monitor via watch list"]], widths=[1.0, 1.0, 4.8])
    d.h("Risk Appetite & Thresholds")
    d.kv_table([("Overall appetite", "[Averse / Cautious / Balanced / Open]"),
                ("Cost threshold", f"Risks with EMV > {CUR} [ ] require Sponsor visibility"),
                ("Schedule threshold", "Any risk threatening a gate milestone is escalated"),
                ("Safety", "Zero tolerance — all safety risks treated as at least High")])
    d.h("Response Strategies")
    d.table(["Threats", "Opportunities", "Description"],
            [["Escalate", "Escalate", "Outside PM authority / project scope — hand to appropriate owner"],
             ["Avoid", "Exploit", "Eliminate the threat / ensure the opportunity occurs"],
             ["Transfer", "Share", "Shift ownership to a third party (insurance, contract) / partner"],
             ["Mitigate", "Enhance", "Reduce / increase probability or impact"],
             ["Accept (active/passive)", "Accept", "Acknowledge; set contingency reserve or no action"]], widths=[1.6, 1.4, 3.8])
    d.h("Contingency & Reserves")
    d.para("Contingency reserve is derived from the EMV of quantified risks and held in the cost baseline (doc 09). "
           "Management reserve is held by the Sponsor and released via Change Request.")
    d.h("Reporting & Tracking")
    d.bullets(["Top risks (score ≥ 10) reported in each Project Status Report (doc 22).",
               "Risk register reviewed weekly by PM, monthly in full with team; risk audit at each gate.",
               "Materialized risks move to the Issues tab of the RAID Log (doc 19)."])
    d.approvals()
    d.save(path)


# ============================================================ 17 QUALITY MGMT PLAN
def build_17(path):
    d = Doc("17")
    d.cover("Describes how quality policies, standards and metrics will be applied so that deliverables meet "
            "requirements (PMBOK 8 principle 'Embed quality').")
    d.h("Quality Policy & Standards")
    d.kv_table([("Organizational quality policy", "[Reference]"),
                ("Applicable standards", "[ISO 9001, ISO/IEC 25010, industry codes, internal standards]"),
                ("Regulatory requirements", "[ ]")])
    d.h("Quality Objectives & Metrics")
    d.table(["Deliverable / process", "Metric", "Target", "Measurement method", "Frequency", "Owner"],
            [["[Deliverable]", "Defect density", "[< X per unit]", "Test results", "Per release", "QA Lead"],
             ["Requirements", "Requirements with acceptance criteria", "100%", "RTM (doc 04)", "Gate 2", "BA"],
             ["Schedule process", "Schedule health checks passed", "≥ 95%", "Schedule audit", "Monthly", "PMO"],
             ["[Deliverable]", "Customer acceptance first pass", "≥ 90%", "Acceptance records", "Per deliverable", "PM"]],
            widths=[1.3, 1.5, 0.9, 1.2, 0.9, 0.9], size=8.5)
    d.h("Roles & Responsibilities")
    d.table(["Role", "Quality responsibility"],
            [["Project Manager", "Owns plan; ensures QA/QC activities are planned, resourced and reported"],
             ["QA Lead", "Audits, process compliance, test strategy, non-conformance management"],
             ["Deliverable owners", "Build to standard; peer reviews; fix defects"],
             ["Business Owner", "Defines acceptance criteria; accepts deliverables"]], widths=[1.8, 5.0])
    d.h("Quality Assurance (Process)")
    d.table(["QA activity", "Scope", "Timing", "Output"],
            [["Process audit", "Compliance with this PMP and standards", "Each gate", "Audit report + actions"],
             ["Peer / design reviews", "Designs, specifications", "Before build", "Review record"],
             ["Lessons learned review", "Process effectiveness", "Monthly", "Doc 24 updates"]], widths=[1.6, 2.2, 1.2, 1.8])
    d.h("Quality Control (Product)")
    d.table(["QC activity", "Deliverable", "Method / tool", "Acceptance criteria"],
            [["Inspection", "[ ]", "Checklist", "[ ]"], ["Testing", "[ ]", "Unit / integration / UAT", "[ ]"],
             ["Measurement", "[ ]", "Control charts, statistical sampling", "[ ]"]], widths=[1.4, 1.6, 2.0, 1.8])
    d.h("Non-Conformance & Corrective Action")
    d.bullets(["Record non-conformances with severity (Critical / Major / Minor) in the RAID Log Issues tab.",
               "Root cause analysis (5 Whys, fishbone); corrective and preventive action with owner and due date.",
               "Repeated non-conformance triggers process review and QA audit."])
    d.h("Continuous Improvement")
    d.para("Plan-Do-Check-Act cycles are applied to project processes; improvements are captured in the "
           "Lessons Learned Register (doc 24).")
    d.approvals([("QA Lead", "[Name]")])
    d.save(path)


# ============================================================ 18 PROCUREMENT PLAN
def build_18(path):
    d = Doc("18")
    d.cover("Documents procurement decisions, approach and vendors to be engaged, from make-or-buy through contract closure.")
    d.h("Procurement Strategy")
    d.kv_table([("Delivery method", "[Single vendor / multiple / design-build / in-house + contractors]"),
                ("Organizational policy", "[Procurement policy reference, trade agreements (e.g. CFTA)]"),
                ("Procurement lead", "[Name / department]")])
    d.h("Make-or-Buy Analysis")
    d.table(["Item", "Make (in-house)", "Buy (external)", "Decision", "Rationale"],
            [["[Item]", "[Cost / capability]", "[Cost / capability]", "Buy", "[Why]"], ["", "", "", "", ""]],
            widths=[1.3, 1.4, 1.4, 0.8, 1.9])
    d.h("Procurement Items")
    d.table(["#", "Goods / services", "Contract type", f"Est. value ({CUR})", "Method", "Start", "Award by", "Owner"],
            [["P1", "[ ]", "FFP", "[ ]", "RFP", "[ ]", "[ ]", "[ ]"], ["P2", "[ ]", "T&M", "[ ]", "RFQ", "", "", ""]],
            widths=[0.4, 1.4, 0.8, 0.9, 0.7, 0.7, 0.8, 0.8], size=8.5)
    d.h("Contract Type Guidance")
    d.table(["Type", "Use when", "Risk to buyer"],
            [["Firm Fixed Price (FFP)", "Scope well defined", "Low"],
             ["Fixed Price Incentive Fee (FPIF)", "Defined scope, performance incentives desired", "Low–medium"],
             ["Cost Plus Fixed/Incentive/Award Fee", "Scope uncertain, R&D", "High"],
             ["Time & Materials (T&M)", "Staff augmentation, small or uncertain scope", "Medium–high (use NTE cap)"]],
            widths=[2.3, 3.0, 1.5])
    d.h("Source Selection Criteria")
    d.table(["Criterion", "Weight", "Description"],
            [["Technical solution & approach", "35%", "Fit to SOW and requirements"],
             ["Price / total cost of ownership", "25%", "Lifecycle cost"],
             ["Experience & references", "15%", "Comparable projects"],
             ["Delivery capacity & schedule", "10%", "Team, availability"],
             ["Risk, security & compliance", "10%", "Insurance, certifications, data protection"],
             ["Sustainability / social value", "5%", "ESG commitments"]], widths=[2.6, 0.8, 3.4])
    d.h("Procurement Documents & Roles")
    d.table(["Document / step", "Owner", "Approver"],
            [["Statement of Work (SOW)", "Project Manager", "Sponsor"], ["RFI / RFP / RFQ", "Procurement", "Procurement lead"],
             ["Evaluation & recommendation", "Evaluation panel", "Sponsor"], ["Contract award", "Procurement / Legal", "Authorized signatory"]],
            widths=[2.6, 2.1, 2.1])
    d.h("Vendor Management & Performance")
    d.bullets(["Kick-off with each vendor; contract obligations register maintained.",
               "Performance reviewed [monthly] against KPIs/SLAs; scorecard shared with vendor.",
               "Invoices verified against accepted deliverables before approval for payment.",
               "Contract changes via change order linked to project Change Request (doc 20)."])
    d.h("Procurement Risks")
    d.table(["Risk", "Response"], [["[Vendor insolvency]", "[Financial checks, performance bond]"],
                                   ["[Late delivery]", "[Liquidated damages, milestone payments]"]], widths=[3.0, 3.8])
    d.h("Contract Closure")
    d.bullets(["Confirm all deliverables accepted and obligations met.", "Final payment and release of retention/bonds.",
               "Vendor performance evaluation recorded; records archived."])
    d.approvals([("Procurement Lead", "[Name]")])
    d.save(path)


# ============================================================ 20 CHANGE REQUEST
def build_20(path):
    d = Doc("20")
    d.cover("Form used to request, analyse and decide on any change to an approved baseline or project document.")
    d.h("Change Request Information")
    d.kv_table([("CR ID", f"{PROJECT['id']}-CR-[###]  (from Change Log, doc 21)"), ("CR title", "[Short title]"),
                ("Requested by / role", "[Name / role]"), ("Date submitted", "[yyyy-mm-dd]"),
                ("Priority", "Critical / High / Medium / Low"),
                ("Change category", "Scope / Schedule / Cost / Quality / Resource / Contract / Document"),
                ("Baselines affected", "Scope ☐  Schedule ☐  Cost ☐  None ☐")])
    d.h("Description of Change")
    d.ph("[Describe the proposed change clearly and precisely.]")
    d.h("Justification / Reason")
    d.ph("[Why is the change needed? What happens if it is not implemented?]")
    d.h("Impact Analysis")
    d.table(["Area", "Impact description", "Magnitude", "Quantified impact"],
            [["Scope", "", "None / Low / Med / High", "[WBS IDs affected]"],
             ["Schedule", "", "", "[+/- working days; milestones]"],
             ["Cost", "", "", f"[+/- {CUR}]"],
             ["Quality", "", "", ""], ["Resources", "", "", "[FTE / skills]"],
             ["Risk", "", "", "[New/changed risk IDs]"], ["Stakeholders / benefits", "", "", ""],
             ["Contracts / procurement", "", "", ""]], widths=[1.4, 2.6, 1.3, 1.5], first_col_bold=True)
    d.h("Options Considered")
    d.table(["Option", "Description", "Pros", "Cons"],
            [["A – Do nothing", "", "", ""], ["B – [Option]", "", "", ""]], widths=[1.4, 2.4, 1.5, 1.5])
    d.h("Recommendation")
    d.ph("[Project Manager recommendation]")
    d.h("Decision")
    d.kv_table([("Decision authority", "Project Manager / Change Control Board / Steering Committee"),
                ("Decision", "Approved ☐  Approved with conditions ☐  Rejected ☐  Deferred ☐"),
                ("Conditions / comments", ""), ("Decision date", ""), ("Decision Log ref.", f"{PROJECT['id']}-DEC-[###]"),
                ("Documents to re-version", "[e.g. 05 WBS v2.0, 07 Schedule v2.0, 09 Budget v2.0]")])
    d.table(["Role", "Name", "Signature", "Date"],
            [["Requestor", "", "", ""], ["Project Manager", PROJECT["pm"], "", ""], ["CCB Chair / Sponsor", "", "", ""]],
            widths=[1.8, 2.0, 2.0, 1.0])
    d.save(path)


# ============================================================ 22 STATUS REPORT (docx)
def build_22_docx(path):
    d = Doc("22")
    d.cover("Periodic report of project performance against baselines, using the RAG thresholds of the Integrated PMP.")
    d.h("Report Information")
    d.kv_table([("Reporting period", "[yyyy-mm-dd] to [yyyy-mm-dd]"), ("Status date", "[yyyy-mm-dd]"),
                ("Prepared by", PROJECT["pm"]), ("Distribution", "Per Communication Plan (doc 12)")])
    d.h("Overall Status Summary")
    t = d.table(["Dimension", "Current", "Previous", "Trend", "Comment"],
                [["Overall", "Green", "Green", "→", ""], ["Scope", "Green", "", "", ""],
                 ["Schedule", "Amber", "", "↓", ""], ["Cost", "Green", "", "", ""],
                 ["Risk", "Amber", "", "", ""], ["Quality", "Green", "", "", ""], ["Resources", "Green", "", "", ""],
                 ["Stakeholders", "Green", "", "", ""]], widths=[1.3, 0.9, 0.9, 0.6, 3.1])
    colors = {"Green": (B["green_l"], B["green"]), "Amber": (B["amber_l"], "8A6200"), "Red": (B["red_l"], B["red"])}
    for row in t.rows[1:]:
        for ci in (1, 2):
            v = row.cells[ci].text
            if v in colors:
                set_cell(row.cells[ci], v, bold=True, color=colors[v][1], size=9); shade(row.cells[ci], colors[v][0])
    d.h("Executive Summary")
    d.ph("[3–5 sentences: progress, key achievements, concerns, decisions required.]")
    d.h("Milestones")
    d.table(["Milestone", "Baseline", "Forecast / actual", "Variance (d)", "Status"],
            [["[Milestone]", "[date]", "[date]", "0", "Complete"], ["[Milestone]", "", "", "", "On track"]],
            widths=[2.4, 1.0, 1.3, 0.9, 1.2])
    d.h("Financial & Earned Value Performance")
    d.table(["BAC", "PV", "EV", "AC", "CPI", "SPI", "EAC", "VAC"],
            [["[ ]", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]", "[ ]"]], widths=[0.85] * 8)
    d.h("Accomplishments This Period / Planned Next Period")
    d.table(["Accomplished this period", "Planned for next period"], [["• ", "• "], ["• ", "• "]],
            widths=[3.4, 3.4])
    d.h("Top Risks & Issues")
    d.table(["ID", "Type", "Description", "Score / severity", "Response / action", "Owner"],
            [["R-001", "Risk", "", "", "", ""], ["I-001", "Issue", "", "", "", ""]],
            widths=[0.6, 0.6, 2.2, 1.0, 1.6, 0.8])
    d.h("Change Requests & Decisions Required")
    d.table(["Ref.", "Description", "Decision needed by", "From whom"], [["CR-001", "", "", ""]],
            widths=[0.9, 3.3, 1.3, 1.3])
    d.save(path)


# ============================================================ 25 CLOSURE REPORT
def build_25(path):
    d = Doc("25")
    d.cover("Documents final project performance, acceptance, handover, outstanding items and lessons, "
            "and formally closes the project.")
    d.h("Project Summary")
    d.kv_table([("Project", f"{PROJECT['name']} / {PROJECT['id']}"), ("Sponsor / PM", f"{PROJECT['sponsor']} / {PROJECT['pm']}"),
                ("Baseline start / finish", f"{PROJECT['start']} / {PROJECT['finish']}"),
                ("Actual start / finish", "[ ] / [ ]"), ("Closure type", "Completed / Terminated early")])
    d.h("Achievement of Objectives & Success Criteria")
    d.table(["Objective (from Charter)", "Target", "Actual result", "Met?", "Comment"],
            [["O1", "", "", "Yes / Partly / No", ""], ["O2", "", "", "", ""], ["O3", "", "", "", ""]],
            widths=[1.8, 1.0, 1.4, 1.0, 1.6])
    d.h("Deliverable Acceptance")
    d.table(["Deliverable", "Accepted by", "Date", "Acceptance ref.", "Conditions"],
            [["D1", "", "", "", ""], ["D2", "", "", "", ""]], widths=[1.8, 1.4, 0.9, 1.3, 1.4])
    d.h("Schedule Performance")
    d.table(["Milestone", "Baseline", "Actual", "Variance (d)", "Reason"],
            [["[ ]", "", "", "", ""]], widths=[2.0, 1.0, 1.0, 0.9, 1.9])
    d.h("Financial Performance")
    d.table(["Item", f"Baseline ({CUR})", f"Final ({CUR})", "Variance", "%"],
            [["Labour", "", "", "", ""], ["Contracts / services", "", "", "", ""], ["Materials / equipment", "", "", "", ""],
             ["Contingency used", "", "", "", ""], ["Management reserve used", "", "", "", ""], ["Total", "", "", "", ""]],
            widths=[2.2, 1.2, 1.2, 1.2, 1.0])
    d.h("Quality Summary")
    d.ph("[Quality metrics achieved, outstanding defects, audit results.]")
    d.h("Benefits Realization & Operational Handover")
    d.table(["Benefit / product", "Receiving owner", "Handover date", "Measurement plan after closure"],
            [["", "", "", ""]], widths=[2.0, 1.5, 1.2, 2.1])
    d.h("Outstanding Items & Transfer")
    d.table(["Item (issue/risk/action)", "Transferred to", "Due", "Notes"], [["", "", "", ""]], widths=[2.6, 1.6, 1.0, 1.6])
    d.h("Contract & Administrative Closure")
    d.bullets(["☐ All contracts closed and final payments made", "☐ Purchase orders closed, accruals cleared",
               "☐ Resources released and feedback provided", "☐ Access rights revoked; assets returned",
               "☐ Project records archived (status 'Archived')", "☐ Financial codes closed"])
    d.h("Lessons Learned Summary")
    d.guide("Summarize the top lessons from the Lessons Learned Register (doc 24) by performance domain.")
    d.table(["Performance domain", "Key lesson", "Recommendation"],
            [["Governance", "", ""], ["Scope", "", ""], ["Schedule", "", ""], ["Finance", "", ""],
             ["Stakeholders", "", ""], ["Resources", "", ""], ["Risk", "", ""]], widths=[1.6, 2.6, 2.6])
    d.h("Sustainability Outcomes")
    d.ph("[Environmental, social and economic outcomes against charter commitments.]")
    d.approvals([("Business Owner", "[Name]")])
    d.save(path)
