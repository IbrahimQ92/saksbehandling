from pypdf import PdfReader
from docx import Document

def extract_text(file):
    if file.name.endswith(".pdf"):
        reader = PdfReader(file)
        return "\n".join(page.extract_text() or "" for page in reader.pages)

    if file.name.endswith(".docx"):
        doc = Document(file)
        return "\n".join(p.text for p in doc.paragraphs)

    if file.name.endswith(".txt"):
        return file.read().decode("utf-8")

    raise ValueError("Ugyldig filtype")
