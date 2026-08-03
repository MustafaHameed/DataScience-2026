# Helper script to reorder sections in a .docx by parsing document.xml
# Not intended for user execution directly in chat; used by assistant to process the docx.
import zipfile
import xml.etree.ElementTree as ET
import shutil
import os

SRC = r"e:\Documents\Github\RnD\2026-DataScience\Foundations of Data Science.docx"
OUT = r"e:\Documents\Github\RnD\2026-DataScience\Lecture101-DataScience.docx"

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def get_text_from_p(p):
    texts = []
    for node in p.findall('.//w:t', ns):
        texts.append(node.text or '')
    return ''.join(texts).strip()

def get_p_style(p):
    pPr = p.find('w:pPr', ns)
    if pPr is None:
        return None
    pStyle = pPr.find('w:pStyle', ns)
    if pStyle is None:
        return None
    return pStyle.get('{%s}val' % ns['w'])

# Extract document.xml
with zipfile.ZipFile(SRC) as zin:
    doc_xml = zin.read('word/document.xml')
root = ET.fromstring(doc_xml)
body = root.find('w:body', ns)

sections = []
current_heading = None
current_nodes = []

for child in list(body):
    if child.tag == '{%s}p' % ns['w']:
        style = get_p_style(child)
        text = get_text_from_p(child)
        if style and style.lower().startswith('heading') and text:
            # start new section
            if current_heading or current_nodes:
                sections.append((current_heading, current_nodes))
            current_heading = text
            current_nodes = [child]
        else:
            current_nodes.append(child)
    else:
        # tables or other content; attach to current_nodes
        current_nodes.append(child)

# append last
if current_heading or current_nodes:
    sections.append((current_heading, current_nodes))

# Build a summary file to inspect (print headings)
print('Found sections:')
for i,(h,nodes) in enumerate(sections):
    print(i, 'Heading:', repr(h), 'Nodes:', len(nodes))

# Decide new order heuristically
order_keywords = [
    'introduction', 'overview', 'learning objectives', 'objectives', 'prerequisite', 'prerequisites',
    'foundations', 'topics', 'syllabus', 'schedule', 'week', 'assessment', 'evaluation', 'reading', 'references', 'contact'
]

# Map headings to indices
mapping = {}
for idx, (h, nodes) in enumerate(sections):
    key = (h or '').lower()
    mapping[idx] = 99
    for j, kw in enumerate(order_keywords):
        if kw in key:
            mapping[idx] = j
            break

# sort indices by mapping then original order
sorted_indices = sorted(range(len(sections)), key=lambda i: (mapping[i], i))

# build new body
new_body = ET.Element('{%s}body' % ns['w'])
for i in sorted_indices:
    _, nodes = sections[i]
    for n in nodes:
        new_body.append(n)
# ensure final sectPr if exists in original
sectPr = body.find('w:sectPr', ns)
if sectPr is not None:
    new_body.append(sectPr)

# create new document XML
new_root = ET.Element(root.tag)
for k,v in root.attrib.items():
    new_root.set(k,v)
new_root.append(new_body)
new_doc_xml = ET.tostring(new_root, encoding='utf-8', method='xml')

# write out new docx by copying all files from src and replacing document.xml
shutil.copyfile(SRC, OUT)
with zipfile.ZipFile(OUT, 'a') as zout:
    # remove existing doc if present
    try:
        zout.getinfo('word/document.xml')
        # recreate zip without that entry is complex; instead write to a temp file
    except KeyError:
        pass

# Rebuild zip: create a temp zip
import tempfile
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
    # replace output
    if os.path.exists(OUT):
        os.remove(OUT)
    # shutil.move handles cross-device moves
    shutil.move(tmpname, OUT)
    print('Wrote new docx to', OUT)
finally:
    if os.path.exists(tmpname):
        try:
            os.remove(tmpname)
        except Exception:
            pass
