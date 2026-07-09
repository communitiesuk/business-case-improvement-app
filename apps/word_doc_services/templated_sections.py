from docx.document import Document
from docx.shared import Pt

from .formatting_methods import add_text, add_subheader, insert_bullet_point_list

overview = (
    "1: Proposal participants and organisational details\n",
    "2: Proposal Summary"
)

sections = (
    "3: Strategic Case (4 questions)\n",
    "4: Economic Case (7 questions)\n",
    "5: Commercial Case (3 questions)\n",
    "6: Funding and Financial Case (4 questions)\n",
    "7: Management Case (5 questions)\n",
    "8: Other business cases"
)

approvals = (
    "9: Subject Matter Expert (SME) assurance, recommendations and declaration\n",
    "10: SRO/SCS sign off"
)

def create_what_youll_be_asked_section(doc: Document):
    p = doc.add_paragraph()

    what_youll_be_asked_run = p.add_run("What you'll be asked in this document:")
    what_youll_be_asked_run.font.size = Pt(16)
    what_youll_be_asked_run.font.bold = True
    what_youll_be_asked_run.font.name = "Aptos Bold"

    overview_paragraph = doc.add_paragraph()
    add_subheader(overview_paragraph, "Overview:\n")
    for item in overview:
        add_text(overview_paragraph, item)

    sections_paragraph = doc.add_paragraph()
    add_subheader(sections_paragraph, "Sections:\n")
    for item in sections:
        add_text(sections_paragraph, item)
    
    approvals_paragraph = doc.add_paragraph()
    add_subheader(approvals_paragraph, "Approvals:\n")
    for item in approvals:
        add_text(approvals_paragraph, item)

def create_before_you_start_section(doc: Document):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(45)
    header_text = p.add_run("Before you start:")
    header_text.font.size = Pt(16)
    header_text.font.bold = True
    header_text.font.name = "Aptos Bold"

    before_start_paragraph = doc.add_paragraph()
    add_text(before_start_paragraph, "This template follows the HM Treasury Green Book Five Case Model. The questions are designed to guide you through developing a clear, evidence-based business case.\n")
    
    insert_bullet_point_list(doc, [
        "Some questions may appear related, but each question collects different information. Take some time to read all questions in this template before you start drafting - it will help you avoid repetition.", 
        "Word counts are provided as a guide only. You do not need to use the full word count suggestions for each question."
        ]
    )
