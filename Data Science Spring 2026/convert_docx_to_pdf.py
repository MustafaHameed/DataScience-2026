from docx2pdf import convert
import os

src = r"e:\Documents\Github\RnD\2026-DataScience\Lecture101-DataScience_polished.docx"
out = r"e:\Documents\Github\RnD\2026-DataScience\Lecture101-DataScience_polished.pdf"

print('Converting', src, '->', out)
convert(src, out)
print('Done')
