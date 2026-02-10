"""
Style and color definitions for document generation.

Provides RGBColor helpers, predefined palettes, and the StyleConfig
dataclass that controls every visual aspect of generated documents.
"""

from dataclasses import dataclass, field
from typing import Optional

from docx.shared import Pt, Inches, RGBColor


# ---------------------------------------------------------------------------
# Predefined color palettes
# ---------------------------------------------------------------------------

COLORS = {
    # 327th Star Corps gold/amber theme
    "327th_gold": RGBColor(0xC8, 0x96, 0x16),
    "327th_dark_gold": RGBColor(0xA0, 0x78, 0x10),
    "327th_light_gold": RGBColor(0xE8, 0xC8, 0x50),
    "heading_bar_gold": RGBColor(0xBF, 0x8F, 0x00),
    "heading_bar_text": RGBColor(0x1A, 0x1A, 0x1A),

    # K Company colors
    "kc_orange": RGBColor(0xFF, 0x8C, 0x00),
    "kc_dark_orange": RGBColor(0xCC, 0x70, 0x00),

    # Republic / Clone Wars theme
    "republic_red": RGBColor(0xCC, 0x00, 0x00),
    "republic_blue": RGBColor(0x1A, 0x47, 0x8A),
    "republic_white": RGBColor(0xFF, 0xFF, 0xFF),

    # Document functional colors
    "black": RGBColor(0x00, 0x00, 0x00),
    "white": RGBColor(0xFF, 0xFF, 0xFF),
    "off_white": RGBColor(0xE8, 0xE8, 0xE8),
    "dark_gray": RGBColor(0x2D, 0x2D, 0x2D),
    "page_dark": RGBColor(0x3B, 0x3B, 0x3B),
    "medium_gray": RGBColor(0x88, 0x88, 0x88),
    "light_gray": RGBColor(0xCC, 0xCC, 0xCC),
    "callout_bg": RGBColor(0x4A, 0x4A, 0x4A),

    # Instructional color codes (brighter for dark backgrounds)
    "read_aloud_green": RGBColor(0x55, 0xCC, 0x55),
    "host_info_red": RGBColor(0xFF, 0x55, 0x55),
    "important_blue": RGBColor(0x66, 0xBB, 0xFF),

    # JDU temple codes (brighter for dark backgrounds)
    "temple_green": RGBColor(0x55, 0xCC, 0x55),
    "temple_yellow": RGBColor(0xFF, 0xDD, 0x33),
    "temple_orange": RGBColor(0xFF, 0xAA, 0x33),
    "temple_red": RGBColor(0xFF, 0x55, 0x55),
    "temple_black": RGBColor(0xCC, 0xCC, 0xCC),
    "temple_purple": RGBColor(0xCC, 0x77, 0xFF),
}


def hex_to_rgb(hex_str: str) -> RGBColor:
    """Convert a hex color string like '#D4A017' to an RGBColor."""
    hex_str = hex_str.lstrip("#")
    r, g, b = int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16)
    return RGBColor(r, g, b)


# ---------------------------------------------------------------------------
# StyleConfig — single object that controls all document formatting
# ---------------------------------------------------------------------------

@dataclass
class StyleConfig:
    """Holds every configurable style parameter for document generation."""

    # Page layout
    page_width: float = 8.5       # inches
    page_height: float = 11.0     # inches
    margin_top: float = 0.8       # inches
    margin_bottom: float = 0.8
    margin_left: float = 0.8
    margin_right: float = 0.8

    # Dark theme
    dark_mode: bool = True
    page_background: str = "page_dark"       # page fill color
    heading_bar_color: str = "heading_bar_gold"  # gold bar behind headings
    heading_bar_text_color: str = "heading_bar_text"  # text inside bars

    # Fonts
    title_font: str = "Arial"
    title_size: int = 30          # pt
    title_color: str = "327th_gold"
    title_bold: bool = True

    subtitle_font: str = "Arial"
    subtitle_size: int = 14
    subtitle_color: str = "light_gray"

    heading_font: str = "Arial"
    heading_size: int = 16
    heading_color: str = "heading_bar_text"
    heading_bold: bool = True

    subheading_font: str = "Arial"
    subheading_size: int = 13
    subheading_color: str = "327th_gold"
    subheading_bold: bool = True

    body_font: str = "Arial"
    body_size: int = 11
    body_color: str = "off_white"
    body_bold: bool = False

    # Accent / decorative
    accent_color: str = "327th_gold"
    divider_color: str = "327th_gold"
    bullet_color: str = "327th_gold"

    # Color-coded instruction text
    read_aloud_color: str = "read_aloud_green"
    host_info_color: str = "host_info_red"
    important_info_color: str = "important_blue"

    # Header / footer
    header_text: str = ""
    footer_text: str = ""
    header_font_size: int = 8
    footer_font_size: int = 8

    # Spacing
    paragraph_spacing_before: int = 2   # pt
    paragraph_spacing_after: int = 4    # pt
    line_spacing: float = 1.15

    # Table of contents
    toc_enabled: bool = True
    toc_title: str = "Table of Contents"

    # Decorative symbols
    section_symbol: str = ""          # e.g. "☬" for JDU docs
    use_section_symbols: bool = False

    # Emblem (Republic cog) on title page
    show_emblem: bool = True

    def resolve_color(self, color_key: str) -> RGBColor:
        """Resolve a color key to an RGBColor. Accepts palette names or hex."""
        if color_key in COLORS:
            return COLORS[color_key]
        if color_key.startswith("#"):
            return hex_to_rgb(color_key)
        return COLORS.get("black", RGBColor(0, 0, 0))

    def _bg_hex(self) -> str:
        """Return the page background as a hex string (no #)."""
        c = self.resolve_color(self.page_background)
        return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"

    def _heading_bar_hex(self) -> str:
        c = self.resolve_color(self.heading_bar_color)
        return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"


# ---------------------------------------------------------------------------
# Preset themes
# ---------------------------------------------------------------------------

THEME_327TH = StyleConfig()

THEME_K_COMPANY = StyleConfig(
    title_color="kc_orange",
    heading_bar_color="kc_dark_orange",
    subheading_color="kc_orange",
    accent_color="kc_orange",
    divider_color="kc_orange",
    section_symbol="\u2620",  # skull and crossbones
)

THEME_JDU = StyleConfig(
    section_symbol="\u262C",  # ☬
    use_section_symbols=True,
)

THEME_REPUBLIC = StyleConfig(
    title_color="republic_red",
    heading_bar_color="republic_blue",
    subheading_color="republic_red",
    accent_color="republic_red",
    divider_color="republic_blue",
)

THEMES = {
    "327th": THEME_327TH,
    "k_company": THEME_K_COMPANY,
    "jdu": THEME_JDU,
    "republic": THEME_REPUBLIC,
}
