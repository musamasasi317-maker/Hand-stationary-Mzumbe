import streamlit as st
import requests
import io
import re
import os
import base64
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.oxml.shape import CT_Shape
from docx.oxml.drawing import CT_Drawing
from docx.oxml.text.run import CT_R
import tempfile
import subprocess

# -------------------- Page Config --------------------
st.set_page_config(page_title="FormatFix Premium", page_icon="📄", layout="centered")

# -------------------- CSS --------------------
st.markdown("""
<style>
    .stApp { background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%); }
    h1, h2, h3 { color: #1a2a6c; font-weight: 700; }
    .stButton button { background-color: #1a2a6c; color: white; font-weight: bold; border-radius: 10px; padding: 0.5rem 2rem; border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.1); transition: all 0.3s ease; }
    .stButton button:hover { transform: scale(1.02); background-color: #0f1a4a; box-shadow: 0 6px 10px rgba(0,0,0,0.2); }
    .stRadio label { font-weight: 600; color: #1a2a6c; }
    .stTextInput input, .stTextArea textarea { border-radius: 8px; border: 1px solid #ced4da; padding: 10px; }
    .stSuccess { background-color: #d4edda; border-radius: 10px; padding: 10px; }
    .footer { color: #6c757d; text-align: center; margin-top: 30px; }
    .preview-frame { width: 100%; height: 600px; border: 2px solid #1a2a6c; border-radius: 10px; box-shadow: 0 8px 16px rgba(0,0,0,0.2); }
</style>
""", unsafe_allow_html=True)

# -------------------- Helpers --------------------
def add_border_to_first_page(doc):
    if not doc.sections: return
    sect_pr = doc.sections[0]._sectPr
    pg = OxmlElement('w:pgBorders')
    pg.set(qn('w:display'), 'firstPage')
    pg.set(qn('w:offsetFrom'), 'text')
    for b in ['top','left','bottom','right']:
        e = OxmlElement(f'w:{b}')
        e.set(qn('w:val'), 'double')
        e.set(qn('w:sz'), '24')
        e.set(qn('w:space'), '4')
        e.set(qn('w:color'), 'auto')
        pg.append(e)
    sect_pr.append(pg)

def remove_borders_other_sections(doc):
    for idx, sec in enumerate(doc.sections):
        if idx == 0: continue
        for e in sec._sectPr.findall(qn('w:pgBorders')):
            sec._sectPr.remove(e)

def add_page_numbers(doc):
    for idx, sec in enumerate(doc.sections):
        if idx == 0:
            sec.different_first_page_header_footer = True
            sec.first_page_footer.paragraphs.clear()
            sec.footer.paragraphs.clear()
        else:
            sec.footer.paragraphs.clear()
            p = sec.footer.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            fld = OxmlElement('w:fldSimple')
            fld.set(qn('w:instr'), 'PAGE')
            p._element.append(fld)

def get_logo(uploaded=None):
    if uploaded: return io.BytesIO(uploaded.getvalue())
    if os.path.exists("mzumbe_logo.png"):
        with open("mzumbe_logo.png", "rb") as f: return io.BytesIO(f.read())
    for url in ["https://upload.wikimedia.org/wikipedia/commons/e/e1/Mzumbe_University_logo.png",
                "https://www.mzumbe.ac.tz/images/logo.png"]:
        try:
            r = requests.get(url, timeout=5)
            if r.status_code == 200: return io.BytesIO(r.content)
        except: pass
    return None

def apply_case(text, mode):
    if not text: return text
    if mode == "capital": return text.upper()
    if mode == "lower": return text.lower()
    if mode == "sentence":
        # simple sentence case
        sentences = re.split(r'(\.\s+)', text.lower())
        result = []
        for i, part in enumerate(sentences):
            if i % 2 == 0 and part:
                part = part[0].upper() + part[1:]
            result.append(part)
        return ''.join(result)
    return text

def format_ref(paragraph):
    # Simple italicize book titles (looks for "(Year). Title. Publisher")
    text = paragraph.text
    if not text: return
    m = re.search(r'\((\d{4})\)\.\s*', text)
    if not m: return
    start = m.end()
    rest = text[start:]
    m2 = re.search(r'\.\s+([A-Z][a-z])', rest)
    if m2:
        title_end = start + m2.start() + 1
        before = text[:start].strip()
        title = text[start:title_end].strip()
        publisher = text[title_end:].strip()
        paragraph.clear()
        if before:
            r = paragraph.add_run(before + ' ')
            r.font.name = 'Times New Roman'; r.font.size = Pt(12)
        r = paragraph.add_run(title)
        r.font.name = 'Times New Roman'; r.font.size = Pt(12); r.italic = True
        if publisher:
            r = paragraph.add_run(' ' + publisher)
            r.font.name = 'Times New Roman'; r.font.size = Pt(12)
    else:
        before = text[:start].strip()
        title = rest.strip()
        paragraph.clear()
        if before:
            r = paragraph.add_run(before + ' ')
            r.font.name = 'Times New Roman'; r.font.size = Pt(12)
        r = paragraph.add_run(title)
        r.font.name = 'Times New Roman'; r.font.size = Pt(12); r.italic = True

# -------------------- Cover Page --------------------
def create_cover(doc, atype, meta, students, uploaded_logo, case_mode):
    def center(text, bold=False):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(apply_case(text, case_mode))
        r.bold = bold
        r.font.name = 'Times New Roman'; r.font.size = Pt(12)
        return p

    center("MZUMBE UNIVERSITY", bold=True)
    logo = get_logo(uploaded_logo)
    if logo:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(); r.add_picture(logo, width=Inches(1.8))
    else:
        center("[Nembo ya Mzumbe - tafadhali weka picha]", bold=False)

    center("SCHOOL OF BUSINESS (SOB)", bold=True)
    for _ in range(2): doc.add_paragraph()

    if atype == "Individual Assignment":
        rows = [
            ("PROGRAMME :", meta.get("programme","")),
            ("SUBJECT NAME :", meta.get("subject_name","")),
            ("SUBJECT CODE :", meta.get("subject_code","")),
            ("LECTURER'S NAME :", meta.get("lecturer","")),
            ("NAME :", meta.get("student_name","")),
            ("REG.NUMBER :", meta.get("reg_number","")),
            ("NATURE OF WORK :", "INDIVIDUAL ASSIGNMENT"),
            ("DATE OF SUBMISSION :", meta.get("date_submission","")),
        ]
    else:
        rows = [
            ("PROGRAMME :", meta.get("programme","")),
            ("GROUP NUMBER :", meta.get("group_number","")),
            ("SUBJECT NAME :", meta.get("subject_name","")),
            ("SUBJECT CODE :", meta.get("subject_code","")),
            ("LECTURER'S NAME :", meta.get("lecturer","")),
            ("NATURE OF TASK :", "GROUP ASSIGNMENT"),
            ("DATE OF SUBMISSION :", meta.get("date_submission","")),
        ]

    # Borderless table
    tbl = doc.add_table(rows=len(rows), cols=2)
    tbl.autofit = False
    tbl.columns[0].width = Inches(2.5)
    tbl.columns[1].width = Inches(4.0)

    for i, (label, value) in enumerate(rows):
        # Label
        cell = tbl.rows[i].cells[0]
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 2.5
        r = p.add_run(apply_case(label.upper(), case_mode))
        r.bold = True
        r.font.name = 'Times New Roman'; r.font.size = Pt(12)

        # Value
        cell = tbl.rows[i].cells[1]
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 2.5
        r = p.add_run(apply_case(str(value), case_mode))
        r.font.name = 'Times New Roman'; r.font.size = Pt(12)

    # Make table borders completely invisible
    tblPr = tbl._tbl.tblPr
    tblBorders = parse_xml(f'<w:tblBorders {nsdecls("w")}>'
                           f'<w:top w:val="none"/>'
                           f'<w:left w:val="none"/>'
                           f'<w:bottom w:val="none"/>'
                           f'<w:right w:val="none"/>'
                           f'<w:insideH w:val="none"/>'
                           f'<w:insideV w:val="none"/>'
                           f'</w:tblBorders>')
    tblPr.append(tblBorders)

    # Also ensure each cell has no borders (some Word versions require this)
    for row in tbl.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            # Remove any existing borders
            for border in tcPr.findall(qn('w:tcBorders')):
                tcPr.remove(border)
            # Add borders with none
            borders = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                                f'<w:top w:val="none"/>'
                                f'<w:left w:val="none"/>'
                                f'<w:bottom w:val="none"/>'
                                f'<w:right w:val="none"/>'
                                f'</w:tcBorders>')
            tcPr.append(borders)

    # Group student table (visible)
    if atype == "Group Assignment" and students:
        doc.add_paragraph()
        stbl = doc.add_table(rows=1, cols=3)
        stbl.style = 'Table Grid'
        hdr = stbl.rows[0].cells
        hdr[0].text = "S/N"
        hdr[1].text = "NAME OF STUDENTS"
        hdr[2].text = "REGISTRATION NUMBER"
        for idx, (name, reg) in enumerate(students, 1):
            row = stbl.add_row().cells
            row[0].text = str(idx)
            row[1].text = apply_case(name.strip(), case_mode)
            row[2].text = apply_case(reg.strip(), case_mode)
        for row in stbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.font.name = 'Times New Roman'
                        r.font.size = Pt(12)

    for _ in range(3): doc.add_paragraph()

# -------------------- Copy content with images and shapes (experimental) --------------------
def copy_element_to_new_doc(new_doc, element, original_doc, in_reference, case_mode, force_copy=False):
    """
    Copy a paragraph or table element to new document.
    For paragraphs, attempt to copy both text and inline images (w:drawing).
    For shapes (w:drawing that are not inline), we copy as drawing if possible.
    """
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.drawing import CT_Drawing
    from docx.oxml.shape import CT_Shape

    if element.tag == qn('w:p'):
        p = Document._element_to_paragraph(element)
        text = p.text.strip()
        if text.upper() in ["REFERENCES", "BIBLIOGRAPHY"]:
            new_doc.add_page_break()
            in_reference = True
            heading = new_doc.add_paragraph()
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = heading.add_run(apply_case(text, case_mode))
            r.bold = True
            r.font.name = 'Times New Roman'; r.font.size = Pt(12)
            return in_reference

        new_p = new_doc.add_paragraph()
        # Copy runs
        for run in p.runs:
            # Check for drawing (images/shapes)
            drawing = run._element.find('.//' + qn('w:drawing'))
            if drawing is not None:
                # Try to copy as image
                blip = drawing.find('.//' + qn('a:blip'))
                if blip is not None:
                    rId = blip.get(qn('r:embed'))
                    if rId and rId in original_doc.part.related_parts:
                        img_part = original_doc.part.related_parts[rId]
                        new_run = new_p.add_run()
                        new_run.add_picture(io.BytesIO(img_part.blob), width=Inches(5.0))
                else:
                    # Possibly a shape (SmartArt, chart, etc.) - we'll try to copy the drawing element
                    # This is tricky; we'll clone the drawing element
                    try:
                        # Copy the drawing element as a shape (if it has a blip)
                        # Otherwise we add as a placeholder
                        # For simplicity, we'll just add a text placeholder
                        new_run = new_p.add_run("[Mchoro/Shape - hauwezi kunakiliwa kikamilifu]")
                        new_run.font.name = 'Times New Roman'; new_run.font.size = Pt(12)
                    except:
                        pass
            else:
                # Text run
                if run.text.strip():
                    new_run = new_p.add_run(apply_case(run.text, case_mode))
                    new_run.font.name = 'Times New Roman'
                    new_run.font.size = Pt(12)

        # Apply formatting
        for r in new_p.runs:
            r.font.name = 'Times New Roman'
            r.font.size = Pt(12)
        new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        new_p.paragraph_format.line_spacing = 1.5
        if in_reference and text.upper() not in ["REFERENCES", "BIBLIOGRAPHY"]:
            new_p.paragraph_format.first_line_indent = Inches(-0.5)
            new_p.paragraph_format.left_indent = Inches(0.5)
            format_ref(new_p)

        return in_reference

    elif element.tag == qn('w:tbl'):
        # Copy table (basic)
        try:
            table = Document._element_to_table(element)
            if table.rows:
                new_table = new_doc.add_table(rows=len(table.rows), cols=len(table.rows[0].cells))
                new_table.style = 'Table Grid'
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        new_table.cell(i, j).text = apply_case(cell.text, case_mode)
        except:
            pass
        return in_reference

    return in_reference

# -------------------- Main Processing --------------------
def process_document(input_mode, uploaded_file, raw_text, metadata, atype, students, uploaded_logo, case_mode):
    new_doc = Document()
    # Cover
    create_cover(new_doc, atype, metadata, students, uploaded_logo, case_mode)
    add_border_to_first_page(new_doc)
    new_doc.add_section()
    remove_borders_other_sections(new_doc)

    in_reference = False

    if input_mode == "Pakia Faili la DOCX" and uploaded_file is not None:
        original = Document(uploaded_file)
        # Iterate over body elements in order
        body = original._element.body
        for child in body.iterchildren():
            in_reference = copy_element_to_new_doc(new_doc, child, original, in_reference, case_mode)
    elif input_mode == "Bandika Maandishi" and raw_text:
        lines = [line.strip() for line in raw_text.splitlines() if line.strip()]
        for text in lines:
            if text.upper() in ["REFERENCES", "BIBLIOGRAPHY"]:
                new_doc.add_page_break()
                in_reference = True
                heading = new_doc.add_paragraph()
                heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = heading.add_run(apply_case(text, case_mode))
                r.bold = True
                r.font.name = 'Times New Roman'; r.font.size = Pt(12)
                continue
            p = new_doc.add_paragraph()
            r = p.add_run(apply_case(text, case_mode))
            r.font.name = 'Times New Roman'; r.font.size = Pt(12)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            if in_reference:
                p.paragraph_format.first_line_indent = Inches(-0.5)
                p.paragraph_format.left_indent = Inches(0.5)
                format_ref(p)
    else:
        raise ValueError("Hakuna maandishi yaliyopatikana.")

    add_page_numbers(new_doc)
    buffer = io.BytesIO()
    new_doc.save(buffer)
    buffer.seek(0)
    return buffer

# -------------------- PDF Preview (using win32com if available) --------------------
def convert_to_pdf_win32(docx_buffer):
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp.write(docx_buffer.getvalue())
            doc_path = tmp.name
        pdf_path = doc_path.replace(".docx", ".pdf")
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
        doc.Close()
        word.Quit()
        with open(pdf_path, "rb") as f:
            pdf_data = f.read()
        os.unlink(doc_path); os.unlink(pdf_path)
        return io.BytesIO(pdf_data)
    except ImportError:
        raise ImportError("Win32COM haipatikani. Endesha: pip install pywin32")
    except Exception as e:
        raise Exception(f"Hitilafu: {e}")

# -------------------- UI --------------------
st.markdown("<h1 style='text-align:center;'>📄 FormatFix Premium</h1>"
            "<p style='text-align:center;font-size:1.2rem;'>Panga kazi za wanafunzi kwa mtindo wa kitaaluma <strong>katika sekunde moja</strong>.</p><hr>",
            unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Mipangilio")
    uploaded_logo = st.file_uploader("Pakia Nembo Mwenyewe (hiari)", type=["png","jpg","jpeg"])
    case_mode = st.selectbox("Uandishi wa Herufi",
        ["Uandishi halisi (Original)", "Herufi kubwa (CAPITAL)", "Herufi ndogo (small)", "Herufi za kwanza kubwa (Sentence)"],
        index=0)
    case_map = {"Uandishi halisi (Original)":"original","Herufi kubwa (CAPITAL)":"capital","Herufi ndogo (small)":"lower","Herufi za kwanza kubwa (Sentence)":"sentence"}
    case_mode_internal = case_map[case_mode]

input_mode = st.radio("Chagua njia", ["Pakia Faili la DOCX", "Bandika Maandishi"], index=0, horizontal=True)

uploaded_file = None
raw_text = ""
if input_mode == "Pakia Faili la DOCX":
    uploaded_file = st.file_uploader("Pakua faili (.docx)", type=["docx"])
else:
    raw_text = st.text_area("Bandika maandishi", height=300, placeholder="Weka kila aya kwenye mstari mpya.")

atype = st.radio("Aina ya Kazi", ["Individual Assignment", "Group Assignment"], index=0, horizontal=True)

col1, col2 = st.columns(2)
with col1:
    programme = st.text_input("Programme", placeholder="BBA-EIM 1A")
    subject_name = st.text_input("Jina la Somo", placeholder="PRINCIPLES OF MANAGEMENT")
    subject_code = st.text_input("Msimbo wa Somo", placeholder="PUB 111")
    lecturer = st.text_input("Jina la Mhadhiri", placeholder="PAULO MARO")
with col2:
    if atype == "Individual Assignment":
        student_name = st.text_input("Jina la Mwanafunzi", placeholder="Jina kamili")
        reg_number = st.text_input("Namba ya Usajili", placeholder="1739110/T.25")
        group_number = ""
    else:
        group_number = st.text_input("Namba ya Kikundi", placeholder="Group 5")
        student_name = ""
        reg_number = ""
    date_submission = st.text_input("Tarehe ya Uwasilishaji", placeholder="15 JUNE 2026")

students = []
if atype == "Group Assignment":
    st.subheader("Taarifa za Wanafunzi")
    st.caption("Weka kila mwanafunzi kwenye mstari mpya: `Jina, Namba` (comma au slash)")
    student_text = st.text_area("Wanafunzi", height=150, placeholder="John Mushi, 2024-01-0002\nJane Kilima, 2024-01-0003")
    if student_text.strip():
        for line in student_text.strip().splitlines():
            line = line.strip()
            if not line: continue
            parts = re.split(r'[,/]\s*', line, maxsplit=1)
            if len(parts)==2: students.append((parts[0].strip(), parts[1].strip()))
            else: st.warning(f"Imepitishwa: {line}")

col_btn1, col_btn2 = st.columns(2)
process_btn = col_btn1.button("🚀 Pangwa Katika Sekunde", use_container_width=True)
preview_btn = col_btn2.button("👁️ Onyesha Hati (Preview)", use_container_width=True, disabled=not st.session_state.get('processed', False))

if 'doc_buffer' not in st.session_state: st.session_state.doc_buffer = None
if 'processed' not in st.session_state: st.session_state.processed = False
if 'pdf_buffer' not in st.session_state: st.session_state.pdf_buffer = None

if process_btn:
    errors = []
    if input_mode == "Pakia Faili la DOCX" and uploaded_file is None:
        errors.append("Pakua faili la DOCX.")
    elif input_mode == "Bandika Maandishi" and not raw_text.strip():
        errors.append("Bandika maandishi.")
    if not programme: errors.append("Programme inahitajika.")
    if not subject_name: errors.append("Jina la Somo linahitajika.")
    if not subject_code: errors.append("Msimbo wa Somo unahitajika.")
    if not lecturer: errors.append("Jina la Mhadhiri linahitajika.")
    if not date_submission: errors.append("Tarehe inahitajika.")
    if atype == "Individual Assignment":
        if not student_name: errors.append("Jina la Mwanafunzi linahitajika.")
        if not reg_number: errors.append("Namba ya Usajili inahitajika.")
    else:
        if not group_number: errors.append("Namba ya Kikundi inahitajika.")
        if not students: st.warning("Hakuna wanafunzi. Jedwali litakuwa tupu.")

    if errors:
        for err in errors: st.error(err)
    else:
        try:
            with st.spinner("Inapanga..."):
                meta = {
                    "programme": programme,
                    "subject_name": subject_name,
                    "subject_code": subject_code,
                    "lecturer": lecturer,
                    "date_submission": date_submission,
                    "student_name": student_name,
                    "reg_number": reg_number,
                    "group_number": group_number,
                }
                buffer = process_document(
                    input_mode, uploaded_file, raw_text, meta, atype, students,
                    uploaded_logo, case_mode_internal
                )
                st.session_state.doc_buffer = buffer
                st.session_state.processed = True
                st.success("🎉 Kazi imepangwa kikamilifu!")
                st.balloons()

                st.download_button(
                    label="⬇ Pakua Hati (DOCX)",
                    data=buffer,
                    file_name="FormatFix_Perfect_Document.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

                # Preview using win32com if possible
                try:
                    with st.spinner("Inaandaa preview..."):
                        pdf = convert_to_pdf_win32(buffer)
                        st.session_state.pdf_buffer = pdf
                        base64_pdf = base64.b64encode(pdf.getvalue()).decode('utf-8')
                        st.markdown(f'<iframe class="preview-frame" src="data:application/pdf;base64,{base64_pdf}"></iframe>', unsafe_allow_html=True)
                except ImportError:
                    st.info("💡 Ili kuona preview, hakikisha una Word na pywin32. Endesha: pip install pywin32")
                except Exception as e:
                    st.warning(f"Preview haipatikani: {e}")

        except Exception as e:
            st.error(f"Hitilafu: {e}")
            st.exception(e)

if preview_btn and st.session_state.processed and st.session_state.doc_buffer:
    try:
        with st.spinner("Inaandaa preview..."):
            pdf = convert_to_pdf_win32(st.session_state.doc_buffer)
            st.session_state.pdf_buffer = pdf
            base64_pdf = base64.b64encode(pdf.getvalue()).decode('utf-8')
            st.markdown(f'<iframe class="preview-frame" src="data:application/pdf;base64,{base64_pdf}"></iframe>', unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Preview haipatikani: {e}")

st.markdown("<hr><p class='footer'>Imetengenezwa kwa ❤️ na Musa Lucas Masasi · Chuo Kikuu cha Mzumbe</p>", unsafe_allow_html=True)
