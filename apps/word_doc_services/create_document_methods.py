from docx import Document
from docx.document import Document as doc #use 'as' here or it gets confused between imports from docx and docx.document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from uuid import uuid4

# this is a place holder for an actual location
save_location: str = '/Users/WordDocuments/{}.docx'

def create_word_document(doc_title: str):
    if not doc_title or len(doc_title) <= 0:
        raise ValueError("No Document title was provided")
    
    document = Document()
    document.add_heading(f"This is your business justification case template for {doc_title}", 0).alignment = WD_ALIGN_PARAGRAPH.CENTER

def save_word_doc(doc: doc, doc_file_name: str):
    doc_id = str(uuid4())
    
    doc_title = doc_file_name + "_" + doc_id

    doc.save(save_location.format(doc_title))


