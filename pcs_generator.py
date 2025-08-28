
import pandas as pd

PHASE_ORDER = ["Preclinical", "Phase 1", "Phase 2", "Phase 3/PPQ"]

def make_control_matrix(unitops_df, cqas_df, mapping_df):
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
    cols = ["UnitOp","Parameter","Class","IPC","Target","Range"]
    keep = [c for c in cols if c in params_df.columns]
    out = params_df[keep].copy()
    out["Class"] = out["Class"].astype(str)
    return out

def make_acceptance_criteria(cqas_df, phase):
    phase_map = {
        "Preclinical": "Phase1_Accept",
        "Phase 1": "Phase1_Accept",
        "Phase 2": "Phase2_Accept",
        "Phase 3/PPQ": "Phase3_Accept"
    }
    sel = phase_map.get(phase, "Phase1_Accept")
    rows = []
    for _, r in cqas_df.iterrows():
        rows.append({
            "CQA": r.get("CQA"),
            "Method": r.get("Method",""),
            "Acceptance": r.get(sel, ""),
            "Notes": r.get("Notes","")
        })
    return pd.DataFrame(rows)

def make_per_unitop_param_tables(unitops_df, params_df):
    """Return list of (subtitle, df) with parameters grouped by UnitOp."""
    tables = []
    if "UnitOp" not in params_df.columns or len(params_df)==0:
        return tables
    for u in unitops_df["UnitOp"].dropna().unique():
        df = params_df[params_df["UnitOp"]==u].copy()
        if df.empty: 
            continue
        # normalize columns
        cols = ["Parameter","Class","IPC","Target","Range"]
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        tables.append((f"{u} – Parameters & Controls", df[["Parameter","Class","IPC","Target","Range"]]))
    return tables

def make_unitop_narrative(uop, uops_df, params_df, mapping_df):
    desc = uops_df[uops_df["UnitOp"]==uop]["Description"].astype(str).head(1).tolist()
    descr = desc[0] if desc else ""
    # CQAs impacted at this unit op
    cqa_list = mapping_df[mapping_df["UnitOp"]==uop]["CQA"].dropna().unique().tolist()
    cqa_str = ", ".join(cqa_list) if cqa_list else "none identified"
    # Key CPP/pCPP
    sub = params_df[(params_df["UnitOp"]==uop) & (params_df["Class"].astype(str).str.contains("CPP", case=False, na=False))]
    key_params = sub["Parameter"].dropna().unique().tolist()
    kp_str = ", ".join(key_params) if key_params else "none designated"
    # IPCs
    ipc_params = params_df[(params_df["UnitOp"]==uop) & (params_df["IPC"].astype(str).str.lower()=="yes")]
    ipc_str = ", ".join(ipc_params["Parameter"].dropna().unique().tolist()) if not ipc_params.empty else "no IPCs"
    text = (
        f"**Unit Operation:** {uop}\n\n"
        f"Purpose/Description: {descr}\n\n"
        f"CQAs controlled/monitored here: {cqa_str}.\n\n"
        f"Key designated pCPP/CPP parameters: {kp_str}.\n\n"
        f"In-Process Controls (IPCs): {ipc_str}.\n\n"
        "Risk-based rationale: Parameters are classified based on impact to CQA(s) and process understanding. "
        "Ranges and IPCs are tightened with phase progression and supported by accumulated data."
    )
    return text

def make_justifications(phase, modality, qtpp_df, cqas_df, params_df, mapping_df):
    base = {
        "Preclinical": "Preclinical strategy emphasizes rapid learning and trending. Parameters are initially classified as pCPP or nCPP; IPCs are informational with broad ranges while process capability is established.",
        "Phase 1": "Phase 1 strategy applies risk-based controls aligned to ICH Q8–Q11 with preliminary action/alert limits. Ranges remain intentionally broad while robustness is confirmed.",
        "Phase 2": "Phase 2 strategy tightens ranges and elevates key parameters to CPPs where warranted. IPCs gain clear action/alert limits supported by manufacturing data.",
        "Phase 3/PPQ": "Phase 3/PPQ strategy finalizes CPP designations with NORs/PARs, full acceptance criteria, and confirmatory PPQ evidence of control."
    }
    lines = [base.get(phase, base["Phase 1"])]
    if modality:
        lines.append(f"This PCS is tailored for a {modality} process, with unit operations and CQAs reflecting modality-specific risks.")
    if len(qtpp_df) > 0:
        lines.append("QTPP elements prioritize CQAs and acceptance criteria aligned to clinical performance objectives.")
    if len(params_df) > 0:
        n_cpp = (params_df["Class"].astype(str).str.upper().str.contains("CPP")).sum()
        lines.append(f"{n_cpp} parameter(s) are currently flagged as pCPP/CPP based on impact assessments and mapping to CQAs.")
    if len(mapping_df) > 0:
        ctrls = (mapping_df.get("ControlType","").astype(str)=="Controls").sum()
        lines.append(f"Control Strategy Matrix indicates {ctrls} control link(s) (UnitOp→CQA) where parameters directly control attribute outcomes.")
    return "\n\n".join(lines)
