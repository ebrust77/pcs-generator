# PCS Generator (Streamlit)

Auto-generate a **phase-appropriate Process Control Strategy (PCS)** from your **QTPP, CQAs, and Unit Operations**.

**Features**
- Interactive editors for QTPP, CQAs, Unit Operations, Parameters, and Mappings
- Auto-populates:
  - Control Strategy Matrix (Unit Operation × CQA)
  - CPP/IPC mapping
  - Acceptance criteria (phase-appropriate)
  - Justifications (phase-appropriate, ICH-aligned)
- Export to **DOCX** (based on your template) and **XLSX** (multi-tab workbook)
- Clean UI with tabs, download buttons, and persistent session state
- Works for biologics and cell/gene therapy (you can rename any fields)

## Quick Start
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud
1. Push this folder to GitHub.
2. Create a new Streamlit app (point at `app.py`).
3. Make sure secrets are not required.

## Inputs & Data Model
- **QTPP**: product goals/targets
- **CQAs**: attributes with acceptance criteria
- **Unit Operations**: name, area, description
- **Parameters**: per-unit-op parameters with class (CPP/nCPP/PM), targets, ranges, IPC status
- **Mappings**: link which CQAs are controlled/impacted by which unit operations & parameters

## Template
We included a placeholder template at `templates/base_template.docx` (copied from your upload). The app merges your data into this template and appends auto-generated tables.

## Exports
- **DOCX** with dynamic sections: Control Strategy Matrix, CPP/IPC mapping, Acceptance Criteria, and Phase-appropriate justifications
- **XLSX** with tabs mirroring the app's tables

## Notes
- Built with `python-docx` for DOCX generation, `pandas` + `xlsxwriter` for Excel export.
- You can import/export XLSX to edit data offline.
