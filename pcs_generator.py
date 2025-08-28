
import pandas as pd
import numpy as np

PHASE_ORDER = ["Preclinical", "Phase 1", "Phase 2", "Phase 3/PPQ"]

def make_control_matrix(unitops_df, cqas_df, mapping_df):
    # build matrix with default "—"
    uops = list(unitops_df["UnitOp"].dropna().unique())
    cqas = list(cqas_df["CQA"].dropna().unique())
    mat = pd.DataFrame("—", index=uops, columns=cqas)
    for _, row in mapping_df.iterrows():
        u = row.get("UnitOp")
        cqa = row.get("CQA")
        ctype = row.get("ControlType", "Monitors")
        if u in mat.index and cqa in mat.columns:
            mat.loc[u, cqa] = ctype
    mat.index.name = "Unit Operation"
    return mat.reset_index()

def make_cpp_ipc_mapping(params_df):
    # Show UnitOp, Parameter, Class, IPC, Target, Range
    cols = ["UnitOp","Parameter","Class","IPC","Target","Range"]
    keep = [c for c in cols if c in params_df.columns]
    out = params_df[keep].copy()
    # normalize class naming
    out["Class"] = out["Class"].replace({"CPP":"CPP","pCPP":"pCPP","nCPP":"nCPP","PM":"PM"})
    return out

def make_acceptance_criteria(cqas_df, phase):
    # Construct a tidy table of CQAs with phase-appropriate acceptance
    phase_col = None
    phase_map = {
        "Preclinical": ["Phase1_Accept"],
        "Phase 1": ["Phase1_Accept"],
        "Phase 2": ["Phase2_Accept"],
        "Phase 3/PPQ": ["Phase3_Accept"]
    }
    use_cols = phase_map.get(phase, ["Phase1_Accept"])
    rows = []
    for _, r in cqas_df.iterrows():
        cqa = r.get("CQA")
        method = r.get("Method","")
        note = r.get("Notes","")
        val = None
        for c in use_cols:
            if c in cqas_df.columns:
                val = r.get(c, None)
        rows.append({"CQA": cqa, "Method": method, "Acceptance": val, "Notes": note})
    return pd.DataFrame(rows)

def make_justifications(phase, modality, qtpp_df, cqas_df, params_df, mapping_df):
    # Very simple templated justification per phase
    base = {
        "Preclinical": "Preclinical strategy emphasizes rapid learning and trending. Parameters are initially classified as pCPP or nCPP; IPCs are informational with broad ranges while we establish process capability.",
        "Phase 1": "Phase 1 strategy applies risk-based controls aligned to ICH Q8–Q11 with preliminary action/alert limits. Ranges remain intentionally broad while we confirm process robustness.",
        "Phase 2": "Phase 2 strategy tightens ranges and elevates key parameters to CPPs where warranted. IPCs gain clear action/alert limits and justification is supported by accumulated manufacturing data.",
        "Phase 3/PPQ": "Phase 3/PPQ strategy finalizes CPP designations with NORs/ PARs, full acceptance criteria, and confirmatory PPQ evidence of control."
    }
    lines = [base.get(phase, base["Phase 1"])]
    if modality:
        lines.append(f"This PCS is tailored for a {modality} process, with unit operations and CQAs reflecting modality-specific risks.")
    if len(qtpp_df) > 0:
        lines.append("QTPP elements were used to prioritize CQAs and set acceptance criteria aligned to clinical performance objectives.")
    if len(params_df) > 0:
        n_cpp = (params_df["Class"].astype(str).str.upper().str.contains("CPP")).sum()
        lines.append(f"{n_cpp} parameter(s) are currently flagged as pCPP/CPP based on impact assessments and mapping to CQAs.")
    if len(mapping_df) > 0:
        ctrls = (mapping_df.get("ControlType","").astype(str)=="Controls").sum()
        lines.append(f"Control Strategy Matrix indicates {ctrls} control link(s) (UnitOp→CQA) where parameters directly control attribute outcomes.")
    return "\n\n".join(lines)
