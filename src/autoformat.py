import re
from docx import Document

def iterate_lines(doc_path):
    doc = Document(doc_path)
    for paragraph in doc.paragraphs:
        for line in paragraph.text.splitlines():
            yield line

def remove_digits_regex(text):
    return re.sub(r"\d", "", text)

def remove_broken_lines(original_doc_path, new_doc_path):
    new_doc = Document()
    paragraph_text = ""
    for line in iterate_lines(original_doc_path):
        print(paragraph_text)
        if not line.rstrip().endswith('.') :
            paragraph_text += line.strip("\t\n")
        elif line.rstrip().endswith('.'):
            paragraph_text += line.strip("\t\n")
            paragraph_text += "\n"
            new_doc.add_paragraph(remove_digits_regex(paragraph_text))
            paragraph_text = "\t"
    new_doc.save(new_doc_path)


remove_broken_lines("./documents/original.docx", "./documents/new.docx")
