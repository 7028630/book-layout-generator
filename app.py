import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import os

# ============================================================
#  FORMATTING HELPERS (No Indentations)
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

def add_styled_para(cell, text, font='Arial', size=11, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, space_after=0):
    """Generates a paragraph with strictly NO indentation."""
    para = cell.add_paragraph(text)
    para.alignment = align
    run = para.runs[0]
    run.font.name = font
    # XML force-set font for compatibility
    r = run._element
    r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:ascii'), font)
    r.get_or_add_rPr().get_or_add_rFonts().set(qn('w:hAnsi'), font)
    
    run.font.size = Pt(size)
    run.font.bold = bold
    # Force zero indents
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.first_line_indent = 0
    para.paragraph_format.left_indent = 0
    para.paragraph_format.line_spacing = 1.15
    return para

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
        outer.autofit = False
        
        for col_idx, page in enumerate(pair):
            cell = outer.rows[0].cells[col_idx]
            remove_all_borders(cell)
            set_cell_width(cell, col_width)
            
            # If page was blank/newly added
            if 'content' not in page: continue

            for item in page['content']:
                ctype = item['type']
                text = item['text']
                if not text.strip(): continue
                
                if ctype == 'Title':
                    add_styled_para(cell, text, settings['title_font'], settings['title_size'], True, WD_ALIGN_PARAGRAPH.CENTER, 12)
                elif ctype == 'Subtitle':
                    add_styled_para(cell, text, settings['title_font'], settings['subtitle_size'], False, WD_ALIGN_PARAGRAPH.CENTER, 10)
                elif ctype == 'Main Text':
                    for p_text in text.split('\n'):
                        if p_text.strip():
                            add_styled_para(cell, p_text, settings['main_font'], settings['main_size'], False, WD_ALIGN_PARAGRAPH.LEFT, 6)
                elif ctype == 'Note Block':
                    # Decorative block with symbols
                    symbol = "— ❧ —"
                    add_styled_para(cell, symbol, settings['main_font'], 10, False, WD_ALIGN_PARAGRAPH.CENTER, 2)
                    # Uses the uploaded font name
                    add_styled_para(cell, text, 'Friedolin', settings['main_size'] + 3, True, WD_ALIGN_PARAGRAPH.CENTER, 2)
                    add_styled_para(cell, symbol, settings['main_font'], 10, False, WD_ALIGN_PARAGRAPH.CENTER, 6)

    return doc

# ============================================================
#  STREAMLIT UI
# ============================================================

def main():
    st.set_page_config(page_title='Book Layout Pro', layout='wide')

    # --- SESSION RECOVERY / MIGRATION ---
    # This prevents the KeyError by resetting if old data format is found
    if 'pages' in st.session_state:
        if len(st.session_state.pages) > 0 and 'content' not in st.session_state.pages[0]:
            del st.session_state.pages

    if 'pages' not in st.session_state:
        st.session_state.pages = [{'content': [{'type': 'Title', 'text': 'My New Book'}]}]

    with st.sidebar:
        st.header('Settings')
        page_size = st.selectbox('Paper Size', ['A4', 'Letter'])
        margin_cm = st.slider('Margin (cm)', 0.5, 2.5, 1.2)
        
        # Restricted font list
        allowed_fonts = ['Courier New', 'Times New Roman', 'Arial', 'Consolas', 'Friedolin']
        
        st.subheader('Titles')
        t_font = st.selectbox('Title Font', allowed_fonts, index=1)
        t_size = st.slider('Title Pt', 14, 48, 28)
        s_size = st.slider('Subtitle Pt', 10, 24, 16)
        
        st.subheader('Body')
        m_font = st.selectbox('Main Font', allowed_fonts, index=2)
        m_size = st.slider('Main Pt', 8, 14, 11)

    st.title('📖 Book Layout Generator')

    # --- CONTENT EDITOR ---
    for i, page in enumerate(st.session_state.pages):
        with st.expander(f'Page {i+1}', expanded=True):
            # Sort buttons
            col_a, col_b = st.columns([6, 1])
            if col_b.button('🗑️ Page', key=f'del_pg_{i}'):
                st.session_state.pages.pop(i)
                st.rerun()

            for j, item in enumerate(page['content']):
                c1, c2, c3 = st.columns([1, 4, 0.5])
                item['type'] = c1.selectbox('Style', 
                    ['Title', 'Subtitle', 'Main Text', 'Note Block'], 
                    key=f't_{i}_{j}')
                item['text'] = c2.text_area('Text', value=item['text'], key=f'v_{i}_{j}', label_visibility='collapsed')
                if c3.button('✕', key=f'x_{i}_{j}'):
                    page['content'].pop(j)
                    st.rerun()
            
            if st.button('➕ Add Section', key=f'add_sec_{i}'):
                page['content'].append({'type': 'Main Text', 'text': ''})
                st.rerun()

    if st.button('➕ Add New Page'):
        st.session_state.pages.append({'content': [{'type': 'Main Text', 'text': ''}]})
        st.rerun()

    # --- GENERATION ---
    if st.button('📥 Build Word Document', type='primary'):
        doc = Document()
        # Setup section
        section = doc.sections[0]
        w, h = (29.7, 21.0) if page_size == 'A4' else (27.94, 21.59)
        section.page_width, section.page_height = Cm(w), Cm(h)
        section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Cm(margin_cm)
        
        # Clean doc
        for p in doc.paragraphs:
            p._element.getparent().remove(p._element)

        settings = {
            'title_font': t_font, 'title_size': t_size, 'subtitle_size': s_size,
            'main_font': m_font, 'main_size': m_size
        }
        
        doc = build_book_layout(doc, st.session_state.pages, w, h, margin_cm, settings)
        
        buf = io.BytesIO()
        doc.save(buf)
        st.download_button('⬇️ Download Book', buf.getvalue(), 'my_book.docx', use_container_width=True)

if __name__ == '__main__':
    main()
