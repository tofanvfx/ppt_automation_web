"""Check if is_list is properly stored in the data_text_list tuples."""
import sys, os
sys.path.insert(0, os.getcwd())
from docx import Document
from docx.oxml.ns import qn

docx_path = 'test_run/sub_bullet_test.docx'
doc = Document(docx_path)

w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
for para in doc.paragraphs:
    ilvl = 0
    is_list = False
    numPr = para._element.find(f'.//{{{w_ns}}}numPr')
    if numPr is not None:
        is_list = True
        ilvl_elem = numPr.find(f'{{{w_ns}}}ilvl')
        if ilvl_elem is not None:
            ilvl = int(ilvl_elem.get(f'{{{w_ns}}}val', '0'))
    else:
        if para.style and para.style.name and ('List' in para.style.name or 'Bullet' in para.style.name):
            is_list = True
    
    if para.text.strip():
        print(f"  text={para.text[:30]!r} is_list={is_list} ilvl={ilvl}")
