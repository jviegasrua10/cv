"""
Creates a compact reference.docx for use with pandoc --reference-doc.
Adjusts margins and default font size so the generated DOCX stays within 3 pages.
Run after pandoc has written the default reference.docx:
    pandoc --print-default-data-file reference.docx > reference.docx
    python create_reference_doc.py
"""
from docx import Document
from docx.shared import Inches, Pt

doc = Document('reference.docx')

# Set narrow margins on all sections
for section in doc.sections:
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

# Reduce default paragraph font size to 10pt
normal_style = doc.styles['Normal']
normal_style.font.size = Pt(10)

doc.save('reference.docx')
print("reference.docx updated with compact settings.")
