
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def replace_text_in_paragraphs(doc, replacements: dict):
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                inline = p.runs
                # Replace across runs
                full_text = p.text.replace(key, str(val))
                for i in range(len(inline)-1, -1, -1):
                    p.runs[i].clear()
                p.add_run(full_text)

def replace_text_in_tables(doc, replacements: dict):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in replacements.items():
                    if key in cell.text:
                        # Simplistic replacement; rebuild cell text
                        new_text = cell.text.replace(key, str(val))
                        # Clear cell and add new paragraph with replaced text
                        for paragraph in list(cell.paragraphs):
                            p = paragraph._element
                            p.getparent().remove(p)
                        new_para = cell.add_paragraph(new_text)

def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    return h

def add_simple_table(doc, dataframe, title=None):
    if title:
        add_heading(doc, title, level=2)
    rows, cols = dataframe.shape
    table = doc.add_table(rows=rows+1, cols=cols)
    table.style = "Light List Accent 1" if "Light List Accent 1" in [s.name for s in doc.styles] else None
    # header
    for j, col in enumerate(dataframe.columns):
        table.cell(0, j).text = str(col)
    # body
    for i in range(rows):
        for j in range(cols):
            table.cell(i+1, j).text = "" if dataframe.iat[i, j] is None else str(dataframe.iat[i, j])
    return table

def save_docx_with_sections(template_path, output_path, replacements, sections: list):
    """
    sections: list of tuples (title, pandas.DataFrame or list[ (subtitle, DataFrame) ] or str)
    """
    doc = Document(template_path)
    # Replace simple placeholders
    replace_text_in_paragraphs(doc, replacements)
    replace_text_in_tables(doc, replacements)

    for sec in sections:
        title, content = sec
        add_heading(doc, title, level=1)
        if isinstance(content, str):
            doc.add_paragraph(content)
        elif hasattr(content, "to_dict"):
            # a single DataFrame
            add_simple_table(doc, content, None)
        elif isinstance(content, list):
            # list of (subtitle, df)
            for (subtitle, df) in content:
                add_simple_table(doc, df, title=subtitle)
        else:
            doc.add_paragraph(str(content))
    doc.save(output_path)
