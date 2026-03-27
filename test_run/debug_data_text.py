"""Check what data['text'] actually looks like after parsing the DOCX."""
import sys, os
sys.path.insert(0, os.getcwd())
import copy
from docx import Document
from docx.oxml import parse_xml as oxml_parse
from docx.document import Document as _Document
from docx.oxml.ns import qn
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl

docx_path = 'test_run/sub_bullet_test.docx'
doc = Document(docx_path)
w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# Manually replicate the parsing loop
sections = []
current_section = None

def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)

import re
for item in iter_block_items(doc):
    ilvl = 0
    is_list = False
    numPr = item._element.find(f'.//{{{w_ns}}}numPr')
    if numPr is not None:
        is_list = True
        ilvl_elem = numPr.find(f'{{{w_ns}}}ilvl')
        if ilvl_elem is not None:
            ilvl = int(ilvl_elem.get(f'{{{w_ns}}}val', '0'))
    style_name = item.style.name if item.style else ''
    if style_name and ('List' in style_name or 'Bullet' in style_name):
        is_list = True
        m = re.search(r'\d+', style_name)
        if m:
            style_level = max(0, int(m.group(0)) - 1)
            if ilvl == 0 and style_level > 0:
                ilvl = style_level

    parts = []
    for child in item._element.iterchildren():
        if child.tag.endswith('}r'):
            t_elem = child.find(f'.//{{{w_ns}}}t')
            if t_elem is not None and t_elem.text:
                rPr = child.find(f'.//{{{w_ns}}}rPr') 
                is_bold = rPr is not None and rPr.find(f'.//{{{w_ns}}}b') is not None
                parts.append({'type': 'text', 'value': t_elem.text, 'bold': is_bold})

    text = item.text.strip()
    if not text:
        continue
    
    entry = (parts if parts else text, ilvl, is_list)
    search_text = text
    match = re.search(r'^\[\s*([^\]]+?)\s*\]', search_text)
    if match:
        current_section = {'name': match.group(1), 'content': []}
        sections.append(current_section)
    elif current_section is not None and text:
        current_section['content'].append(entry)

for sec in sections:
    print(f"\n=== Section: {sec['name']} ===")
    for entry in sec['content'][:5]:
        print(f"  tuple len={len(entry)}, ilvl={entry[1]}, is_list={entry[2]}, text_type={type(entry[0])}")
