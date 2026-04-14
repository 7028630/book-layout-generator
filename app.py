import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# ============================================================
#  XML / FORMATTING HELPERS
# ============================================================

def remove_all_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'none')
        tcBorders.append(border)
    tcPr.append(tcBorders)

def set_cell_width(cell, width_cm):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(int(width_cm * 567)))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def set_row_height(row, height_cm, exact=True):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(height_cm * 567)))
    trHeight.set(qn('w:hRule'), 'exact' if exact else 'atLeast')
    trPr.append(trHeight)

def add_styled_para(cell, text, font='Arial', size=11, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=0):
    """Generic paragraph adder with NO indents."""
    para = cell.add_paragraph(text)
    para.alignment = align
    run = para.runs[0]
    run.font.name = font
    # Fix for fonts not appearing in Word's font list correctly
    r = run._element
    r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:ascii'), font)
    r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:hAnsi'), font)
    
    run.font.size = Pt(size)
    run.font.bold = bold
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.first_line_indent = 0
    para.paragraph_format.left_indent = 0
    return para

# ============================================================
#  DOCUMENT SETUP
# ============================================================

def create_document(page_size='A4', margin_cm=1.5):
    doc = Document()
    for para in doc.paragraphs:
        p = para._element
        p.getparent().remove(p)
    
    section = doc.sections[0]
    sizes = {'A4': (29.7, 21.0), 'Letter': (27.94, 21.59), 'A5': (21.0, 14.85)}
    w, h = sizes.get(page_size, (29.7, 21.0))
    section.page_width, section.page_height = Cm(w), Cm(h)
    section.left_margin = section.right_margin = Cm(margin_cm)
    section.top_margin = section.bottom_margin = Cm(margin_cm)
    return doc, w, h

# ============================================================
#  CORE LAYOUT BUILDER
# ============================================================

def build_book_layout(doc, pages_data, page_width_cm, page_height_cm, margin_cm, settings):
    usable_width = page_width_cm - (margin_cm * 2)
    col_width = (usable_width - 0.5) / 2

    for sheet_idx in range(0, len(pages_data), 2):
        pair = pages_data[sheet_idx: sheet_idx + 2]
        while len(pair) < 2:
            pair.append({'content': []})

        if sheet_idx > 0:
            doc.add_page_break()

        outer = doc.add_table(rows=1, cols=2)
        remove_all_borders(outer.cell(0,0))
        remove_all_borders(outer.cell(0,1))

        for col_idx, page in enumerate(pair):
            cell = outer.rows[0].cells[col_idx]
            set_cell_width(cell, col_width)
            
            for item in page['content']:
                ctype = item['type']
                text = item['text']
                
                if ctype == 'Title':
                    add_styled_para(cell, text, settings['title_font'], settings['title_size'], True, WD_ALIGN_PARAGRAPH.CENTER, 10)
                elif ctype == 'Subtitle':
                    add_styled_para(cell, text, settings['title_font'], settings['subtitle_size'], False, WD_ALIGN_PARAGRAPH.CENTER, 8)
                elif ctype == 'Main Text':
                    # Split into actual paragraphs but ensure no indentation
                    for p_text in text.split('\n'):
                        if p_text.strip():
                            add_styled_para(cell, p_text, settings['main_font'], settings['main_size'], False)
                elif ctype == 'Special Note':
                    symbol = item.get('symbol', '***')
                    add_styled_para(cell, symbol, settings['main_font'], 10, False, WD_ALIGN_PARAGRAPH.CENTER)
                    add_styled_para(cell, text, 'Friedolin', settings['main_size']+2, True, WD_ALIGN_PARAGRAPH.CENTER)
                    add_styled_para(cell, symbol, settings['main_font'], 10, False, WD_ALIGN_PARAGRAPH.CENTER)

    return doc

# ============================================================
#  STREAMLIT UI
# ============================================================

def main():
    st.set_page_config(page_title='Optimized Book Gen', layout='wide')
    st.title('📖 Minimalist Book Layout')

    with st.sidebar:
        st.header('⚙️ Settings')
        page_size = st.selectbox('Paper Size', ['A4', 'Letter', 'A5'])
        margin_cm = st.slider('Margin (cm)', 0.5, 3.0, 1.0)
        
        fonts = ['Friedolin', 'Courier New', 'Times New Roman', 'Arial', 'Consolas']
        
        st.subheader('Title Typography')
        t_font = st.selectbox('Title Font', fonts, index=0)
        t_size = st.number_input('Title Size', 10, 40, 24)
        s_size = st.number_input('Subtitle Size', 10, 30, 16)
        
        st.subheader('Body Typography')
        m_font = st.selectbox('Main Font', fonts, index=2)
        m_size = st.number_input('Main Size', 8, 16, 11)

    if 'pages' not in st.session_state:
        st.session_state.pages = [{'content': [{'type': 'Title', 'text': 'Title Name'}]}]

    # --- UI Logic to add/remove content blocks ---
    for i, page in enumerate(st.session_state.pages):
        with st.expander(f'Page {i+1} Contents', expanded=True):
            for j, item in enumerate(page['content']):
                cols = st.columns([1, 4, 1])
                item['type'] = cols[0].selectbox("Type", ["Title", "Subtitle", "Main Text", "Special Note"], key=f"type_{i}_{j}")
                item['text'] = cols[1].text_area("Content", value=item['text'], key=f"text_{i}_{j}", height=100)
                if cols[2].button("🗑️", key=f"del_{i}_{j}"):
                    page['content'].pop(j)
                    st.rerun()
            
            if st.button("➕ Add Block", key=f"add_b_{i}"):
                page['content'].append({'type': 'Main Text', 'text': ''})
                st.rerun()

    if st.button("➕ Add New Page"):
        st.session_state.pages.append({'content': []})
        st.rerun()

    if st.button('📥 Generate Document', type='primary'):
        settings = {
            'title_font': t_font, 'title_size': t_size, 'subtitle_size': s_size,
            'main_font': m_font, 'main_size': m_size
        }
        doc_obj, pw, ph = create_document(page_size, margin_cm)
        doc_obj = build_book_layout(doc_obj, st.session_state.pages, pw, ph, margin_cm, settings)
        
        buf = io.BytesIO()
        doc_obj.save(buf)
        st.download_button('⬇️ Download .docx', buf.getvalue(), 'book.docx')

if __name__ == '__main__':
    main()
