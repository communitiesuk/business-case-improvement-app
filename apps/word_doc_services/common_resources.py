from docx.shared import RGBColor

_regular_font_name: str = "Arial Regular"
_bold_font_name: str = "Arial Bold"
_italic_font_name: str = "Arial Italic"

_hyperlink_font_colour_hex: str = "#1D70B8"
_regular_font_colour_hex: str = "#000000"
_mhclg_green_hex: str = "#00625E"

def get_general_font_colour_rgb() -> RGBColor:
    return translate_hex_to_rgb(_regular_font_colour_hex)

def get_mhclg_green_rgb() -> RGBColor:
    return translate_hex_to_rgb(_mhclg_green_hex)


def translate_hex_to_rgb(hex_colour: str) -> RGBColor:
    return RGBColor(
        int(hex_colour[1:3],16),
        int(hex_colour[3:5],16),
        int(hex_colour[5:7],16)
    )
