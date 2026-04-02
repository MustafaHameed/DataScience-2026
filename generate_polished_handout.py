import zipfile
import xml.etree.ElementTree as ET
import os
import shutil
import tempfile

SRC = r"e:\Documents\Github\RnD\2026-DataScience\Foundations of Data Science.docx"
OUT = r"e:\Documents\Github\RnD\2026-DataScience\Lecture101-DataScience_polished.docx"

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', ns['w'])

def get_paragraph_text(p):
    texts = []
    for node in p.findall('.//w:t', ns):
        texts.append(node.text or '')
    return ''.join(texts).strip()

# extract full plain text
with zipfile.ZipFile(SRC) as zin:
    doc_xml = zin.read('word/document.xml')
root = ET.fromstring(doc_xml)
body = root.find('w:body', ns)
paras = []
for p in body.findall('w:p', ns):
    txt = get_paragraph_text(p)
    if txt:
        paras.append(txt)
full_text = '\n\n'.join(paras)

# create simple summaries
sentences = [s.strip() for s in full_text.replace('\n',' ').split('.') if s.strip()]
overview = '. '.join(sentences[:2]).strip()
if overview and not overview.endswith('.'):
    overview += '.'

# build new document.xml content
w = '{%s}' % ns['w']
new_root = ET.Element('{' + ns['w'] + '}document')
new_root.set('xmlns:w', ns['w'])
new_body = ET.SubElement(new_root, w + 'body')

# helper to add paragraph with style and text
def add_paragraph(body, text, style=None):
    p = ET.SubElement(body, w + 'p')
    if style:
        pPr = ET.SubElement(p, w + 'pPr')
        pStyle = ET.SubElement(pPr, w + 'pStyle')
        pStyle.set(w + 'val', style)
    r = ET.SubElement(p, w + 'r')
    t = ET.SubElement(r, w + 't')
    t.text = text
    return p

# Title
add_paragraph(new_body, 'Lecture101 — Foundations of Data Science', style='Title')
add_paragraph(new_body, '', None)

# Overview
add_paragraph(new_body, 'Overview', style='Heading2')
add_paragraph(new_body, overview or 'Overview: (see detailed content below).', None)
add_paragraph(new_body, '', None)

# Learning Objectives (basic inferred)
add_paragraph(new_body, 'Learning Objectives', style='Heading2')
obj1 = 'Understand core concepts and terminology in Data Science.'
obj2 = 'Familiarize with common data analysis workflows and tools.'
obj3 = 'Apply basic statistical and machine learning techniques to datasets.'
add_paragraph(new_body, '\u2022 ' + obj1)
add_paragraph(new_body, '\u2022 ' + obj2)
add_paragraph(new_body, '\u2022 ' + obj3)
add_paragraph(new_body, '', None)

# Prerequisites
add_paragraph(new_body, 'Prerequisites', style='Heading2')
add_paragraph(new_body, 'Basic programming (Python) and introductory statistics recommended.')
add_paragraph(new_body, '', None)

# Weekly Topics (placeholder headline)
add_paragraph(new_body, 'Weekly Topics / Detailed Content', style='Heading2')

# Insert original content paragraphs (as plain text paragraphs)
for p in paras:
    add_paragraph(new_body, p)

# Readings
add_paragraph(new_body, '', None)
add_paragraph(new_body, 'Readings', style='Heading2')
add_paragraph(new_body, 'Primary textbook and selected papers as indicated in class.')

# Assessment
add_paragraph(new_body, '', None)
add_paragraph(new_body, 'Assessment', style='Heading2')
add_paragraph(new_body, 'Assignments: 40%, Midterm: 30%, Final: 30%')

# Contact
add_paragraph(new_body, '', None)
add_paragraph(new_body, 'Contact', style='Heading2')
add_paragraph(new_body, 'Instructor: Dr. Mustafa Hameed. Office hours: TBD. Email: example@university.edu')

# final sectPr from original if present
sectPr = body.find('w:sectPr', ns)
if sectPr is not None:
    new_body.append(sectPr)

new_doc_xml = ET.tostring(new_root, encoding='utf-8', method='xml')

# write out new docx by copying all parts from original and replacing document.xml
tmpf = tempfile.NamedTemporaryFile(delete=False)
tmpname = tmpf.name
tmpf.close()
try:
    with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(tmpname, 'w') as zout:
        for item in zin.infolist():
            if item.filename == 'word/document.xml':
                zout.writestr(item, new_doc_xml)
            else:
                zout.writestr(item, zin.read(item.filename))
    if os.path.exists(OUT):
        os.remove(OUT)
    shutil.move(tmpname, OUT)
    print('Wrote polished handout to', OUT)
finally:
    if os.path.exists(tmpname):
        try:
            os.remove(tmpname)
        except Exception:
            pass
