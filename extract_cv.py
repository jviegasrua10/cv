from docx import Document

doc = Document('cv-jmpvr.docx')

for para in doc.paragraphs:
    print(para.text)