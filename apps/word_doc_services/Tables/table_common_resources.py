
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.text.paragraph import Paragraph

from docx.oxml.parser import OxmlElement
import docx.oxml.ns
import docx

from docx.shared import Cm
from word_doc_services.common_resources import(
    translate_hex_to_rgb,
    _regular_font_name
)

_blue_help_box_background_hex: str = "#E4F2FF"
_help_box_footer_font_colour_hex: str = "#E8E8E8"
max_table_width: Cm = Cm(15.9)

table_footer_word_count: str = "Word count guideline: {} words"
table_width: Cm = Cm(16.30)

def get_blue_help_box_background_rgb() -> RGBColor:
    return translate_hex_to_rgb(_blue_help_box_background_hex)

def get_help_box_footer_font_colour_rgb() -> RGBColor:
    return translate_hex_to_rgb(_help_box_footer_font_colour_hex)

'''
Summary:
    Set some default formatting on the cell.
    Because we are using the default paragraph that exists when creating a table,
    we need to set this here via Oxml. If we add a paragraph through add_paragraph()
    this formatting won't be applied and the user will use the default styling in Word.
'''
def set_default_paragraph_formatting(p: Paragraph):
    pPr = p._p.get_or_add_pPr()

    # Get or create paragraph-level run properties (<w:rPr>)
    rPr = pPr.find(qn('w:rPr'))
    if rPr is None:
        rPr = OxmlElement('w:rPr')
        pPr.append(rPr)
        
    # Set the font name
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), _regular_font_name)
    rFonts.set(qn('w:hAnsi'), _regular_font_name)
    rPr.append(rFonts)

    # Set the deafult font colour
    color = OxmlElement('w:color')
    color.set(docx.oxml.ns.qn('w:val'), _regular_font_name)
    rPr.append(color)

    # Set the font size (Word measures this in half-points, so 12pt = 24)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), "24")
    rPr.append(sz)
