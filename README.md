# PMO Documentation Framework

A free predictive project management toolkit aligned to the **PMBOK® Guide – Eighth Edition** (PMI, 2025), with a Master PMO Documentation Register that governs it.

**Website:** open `index.html`, or visit the GitHub Pages address for this repository once Pages is enabled.

## What it contains

- **Master Document Register** for 25 artefacts. For each one it states whether the artefact is Mandatory, Conditional or Optional by project complexity, when it is created, who owns and approves it, when it is baselined, and which artefacts it feeds.
- **Lifecycle and gates.** The 7 performance domains (Governance, Scope, Schedule, Finance, Stakeholders, Resources, Risk) are mapped against the 5 focus areas (Initiating, Planning, Executing, Monitoring & Controlling, Closing), with checklists for gates G1–G5.
- **Governance controls:** delegated authority, escalation levels, RAG thresholds, and baseline and change-control rules.
- **25 templates** in Word, Excel and PowerPoint. They share the same metadata, document-control table, versioning scheme and visual identity.

## Repository layout

| Path | Contents |
|---|---|
| `index.html` | The website (static, no server needed) |
| `files/` | Downloadable templates and the complete toolkit zip |
| `source/` | Python scripts that generate the templates, the framework workbook and the website |

## Regenerating with your own project details

```bash
pip install python-docx openpyxl python-pptx
cd source
# edit config.py (PROJECT block) and framework.py (classifications, owners, thresholds)
python build_toolkit.py "../build"
python portal.py "../build" ".." --static --zip "<path to toolkit zip>"
```

## Versioning

Releases follow the toolkit's own document-control scheme:

- 0.x = draft
- 1.0 = first approved baseline
- 1.x = minor revision
- 2.0+ = major revision

Each release lists what changed.

## Disclaimer

This is an independent, tailored implementation of guidance in the PMBOK® Guide – Eighth Edition. It is not a PMI publication and is not endorsed by PMI. PMBOK and PMI are registered marks of the Project Management Institute, Inc. Adapt thresholds, authority limits and roles to your organization.
