
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from pathlib import Path

from pcs_generator import make_control_matrix, make_cpp_ipc_mapping, make_acceptance_criteria, make_justifications
from docx_utils import save_docx_with_sections

st.set_page_config(page_title="PCS Generator", layout="wide")

st.title("Process Control Strategy (PCS) Generator")
st.caption("Generate a phase-appropriate PCS from your QTPP, CQAs, and Unit Operations.")

with st.sidebar:
    st.header("Settings")
    product_name = st.text_input("Product Name", value="PRODUCT X")
    modality = st.selectbox("Modality", ["General/Biologics","Cell Therapy","Viral Vector","mAb","Small Molecule"], index=1)
    phase = st.selectbox("Development Phase", ["Preclinical","Phase 1","Phase 2","Phase 3/PPQ"], index=1)
    st.markdown("---")
    st.write("**Template DOCX**")
    template_file = st.file_uploader("Upload a DOCX template (optional)", type=["docx"])
    use_uploaded_template = False
    template_path = Path("templates/pcs_built_in_template.docx")
    if template_file:
        # save to temp file
        uploaded_template_path = Path("templates/_uploaded_template.docx")
        uploaded_template_path.parent.mkdir(parents=True, exist_ok=True)
        with open(uploaded_template_path, "wb") as f:
            f.write(template_file.read())
        template_path = uploaded_template_path
        use_uploaded_template = True
    st.markdown("---")
    st.write("**Data Import (XLSX)**")
    data_xlsx = st.file_uploader("Upload XLSX with sheets (QTPP, CQAs, UnitOps, Parameters, Mapping)", type=["xlsx"])
    st.caption("Or start with tables below and export to XLSX for reuse.")

st.markdown("### 1) Define QTPP & CQAs")
c1, c2 = st.columns(2)

default_qtpp = pd.DataFrame([
    {"Attribute":"Dose accuracy","Target":"±10% of label","Rationale":"Clinical dosing consistency","Phase":"Phase 1+"},
    {"Attribute":"Sterility","Target":"No growth","Rationale":"Patient safety","Phase":"All"},
])
default_cqas = pd.DataFrame([
    {"CQA":"Potency","Method":"Bioassay","Phase1_Accept":"Trend only","Phase2_Accept":"Action/Alert","Phase3_Accept":"Meets Spec","Notes":"Aligned to MoA"},
    {"CQA":"Purity","Method":"HPLC","Phase1_Accept":"Trend only","Phase2_Accept":"Action/Alert","Phase3_Accept":"≥95%","Notes":"Phase-appropriate tightening"},
    {"CQA":"Identity","Method":"PCR/Seq","Phase1_Accept":"Qualitative","Phase2_Accept":"Qual/Quant","Phase3_Accept":"Pass/Fail","Notes":"Compendial where possible"},
])

if data_xlsx is not None:
    try:
        xls = pd.ExcelFile(data_xlsx)
        qtpp_df = xls.parse("QTPP")
        cqas_df = xls.parse("CQAs")
    except Exception as e:
        st.error(f"Failed to parse uploaded XLSX: {e}")
        qtpp_df = default_qtpp.copy()
        cqas_df = default_cqas.copy()
else:
    qtpp_df = default_qtpp.copy()
    cqas_df = default_cqas.copy()

qtpp_df = st.data_editor(qtpp_df, num_rows="dynamic", use_container_width=True, key="qtpp")
cqas_df = st.data_editor(cqas_df, num_rows="dynamic", use_container_width=True, key="cqas")

st.markdown("### 2) Unit Operations & Parameters")
c3, c4 = st.columns(2)
default_uops = pd.DataFrame([
    {"UnitOp":"Thaw","Area":"Upstream","Description":"Thaw of starting material"},
    {"UnitOp":"Activation","Area":"Upstream","Description":"Activate cells with beads/cytokines"},
    {"UnitOp":"Transduction","Area":"Upstream","Description":"LVV transduction"},
    {"UnitOp":"Expansion","Area":"Upstream","Description":"Cell expansion in bags or bioreactor"},
    {"UnitOp":"Harvest","Area":"Downstream","Description":"Harvest and wash"},
    {"UnitOp":"Formulation/Fill","Area":"Drug Product","Description":"Formulate, fill and cryopreserve"},
])
default_params = pd.DataFrame([
    {"UnitOp":"Activation","Parameter":"Bead:Cell ratio","Class":"pCPP","Target":"1:1","Range":"0.8–1.2:1","IPC":"Yes"},
    {"UnitOp":"Transduction","Parameter":"MOI","Class":"pCPP","Target":"5","Range":"3–7","IPC":"Yes"},
    {"UnitOp":"Expansion","Parameter":"Temp","Class":"nCPP","Target":"37°C","Range":"36–38°C","IPC":"No"},
    {"UnitOp":"Harvest","Parameter":"Spin g-force","Class":"pCPP","Target":"500 g","Range":"400–600 g","IPC":"Yes"},
    {"UnitOp":"Formulation/Fill","Parameter":"Fill volume","Class":"nCPP","Target":"2 mL","Range":"1.9–2.1 mL","IPC":"Yes"},
])

if data_xlsx is not None:
    try:
        uops_df = xls.parse("UnitOps")
        params_df = xls.parse("Parameters")
    except Exception as e:
        st.error(f"Failed to parse UnitOps/Parameters: {e}")
        uops_df = default_uops.copy()
        params_df = default_params.copy()
else:
    uops_df = default_uops.copy()
    params_df = default_params.copy()

uops_df = st.data_editor(uops_df, num_rows="dynamic", use_container_width=True, key="uops")
params_df = st.data_editor(params_df, num_rows="dynamic", use_container_width=True, key="params")

st.markdown("### 3) Mapping (Unit Operation × CQA)")
default_mapping = pd.DataFrame([
    {"UnitOp":"Activation","CQA":"Potency","ControlType":"Controls"},
    {"UnitOp":"Transduction","CQA":"Potency","ControlType":"Controls"},
    {"UnitOp":"Expansion","CQA":"Potency","ControlType":"Monitors"},
    {"UnitOp":"Harvest","CQA":"Purity","ControlType":"Controls"},
    {"UnitOp":"Formulation/Fill","CQA":"Purity","ControlType":"Monitors"},
    {"UnitOp":"Formulation/Fill","CQA":"Identity","ControlType":"Release Test"},
])
if data_xlsx is not None:
    try:
        mapping_df = xls.parse("Mapping")
    except Exception as e:
        st.error(f"Failed to parse Mapping: {e}")
        mapping_df = default_mapping.copy()
else:
    mapping_df = default_mapping.copy()

mapping_df = st.data_editor(mapping_df, num_rows="dynamic", use_container_width=True, key="mapping")

st.markdown("### 4) Generate PCS")
gen = st.button("Generate")

if gen:
    # Derived tables
    control_matrix = make_control_matrix(uops_df, cqas_df, mapping_df)
    cpp_ipc = make_cpp_ipc_mapping(params_df)
    accept = make_acceptance_criteria(cqas_df, phase)
    just = make_justifications(phase, modality, qtpp_df, cqas_df, params_df, mapping_df)

    st.success("PCS tables generated.")
    st.dataframe(control_matrix, use_container_width=True, height=240)
    st.dataframe(cpp_ipc, use_container_width=True, height=240)
    st.dataframe(accept, use_container_width=True, height=240)
    st.text_area("Phase-Appropriate Justifications", just, height=160)

    # Export to XLSX
    xlsx_buf = BytesIO()
    with pd.ExcelWriter(xlsx_buf, engine="xlsxwriter") as writer:
        qtpp_df.to_excel(writer, sheet_name="QTPP", index=False)
        cqas_df.to_excel(writer, sheet_name="CQAs", index=False)
        uops_df.to_excel(writer, sheet_name="UnitOps", index=False)
        params_df.to_excel(writer, sheet_name="Parameters", index=False)
        mapping_df.to_excel(writer, sheet_name="Mapping", index=False)
        control_matrix.to_excel(writer, sheet_name="ControlMatrix", index=False)
        cpp_ipc.to_excel(writer, sheet_name="CPP_IPC", index=False)
        accept.to_excel(writer, sheet_name=f"Acceptance_{phase.replace('/','_')}", index=False)
    xlsx_buf.seek(0)
    st.download_button("Download XLSX", data=xlsx_buf, file_name=f"PCS_Tables_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Export to DOCX (anchored sections)
    replacements = {
        "[PRODUCT NAME]": product_name,
        "PRODUCT NAME": product_name,
        "[DATE]": datetime.now().strftime("%Y-%m-%d")
    }

    from pcs_generator import make_per_unitop_param_tables, make_unitop_narrative
    sections = []

    # QTPP
    sections.append({"kind":"table","anchor":"1. Quality Target Product Profile (QTPP)",
                     "title":"QTPP", "content": qtpp_df})

    # CQAs
    sections.append({"kind":"table","anchor":"2. Critical Quality Attributes (CQAs)",
                     "title":"CQAs & Methods", "content": cqas_df})

    # Unit Ops overview
    sections.append({"kind":"table","anchor":"3. Process Overview & Unit Operations",
                     "title":"Unit Operations", "content": uops_df})

    # Parameters grouped by UnitOp
    for subtitle, df in make_per_unitop_param_tables(uops_df, params_df):
        sections.append({"kind":"table","anchor":"4. Parameters & Controls",
                         "title": subtitle, "content": df})

    # Control Strategy Matrix
    sections.append({"kind":"table","anchor":"5. Control Strategy Matrix (Unit Operation × CQA)",
                     "title":"Control Strategy Matrix", "content": control_matrix})

    # CPP / IPC mapping
    sections.append({"kind":"table","anchor":"6. CPP / IPC Mapping",
                     "title":"CPP / IPC Mapping", "content": cpp_ipc})

    # Acceptance criteria
    sections.append({"kind":"table","anchor":"7. Acceptance Criteria (Phase-Appropriate)",
                     "title": f"Acceptance – {phase}", "content": accept})

    # Per-Unit Operation Narrative
    for u in uops_df["UnitOp"].dropna().unique().tolist():
        narrative = make_unitop_narrative(u, uops_df, params_df, mapping_df)
        sections.append({"kind":"text","anchor":"8. Per-Unit Operation Narrative",
                         "title": u, "content": narrative})

    # Phase-Appropriate Justifications
    sections.append({"kind":"text","anchor":"9. Phase-Appropriate Justifications",
                     "title": f"Justification – {phase}", "content": f"Phase: {phase}\n\n{just}"})

    tmp_out = Path("PCS_Output.docx")
    save_docx_with_sections(str(template_path), str(tmp_out), replacements, sections)
    with open(tmp_out, "rb") as f:
        docx_bytes = f.read()
    st.download_button("Download DOCX", data=docx_bytes, file_name=f"PCS_{product_name.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

st.markdown("---")
st.markdown("**Tips**")
st.markdown("- Use XLSX import/export to collaborate offline.")
st.markdown("- Classify parameters as pCPP/CPP/nCPP/PM and set IPCs to enrich the mapping.")
st.markdown("- Tighten acceptance criteria as you move from Phase 1 → Phase 3/PPQ.")
